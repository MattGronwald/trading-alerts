/**
 * Configuration parameters
 */
const CONFIG = {
  DEBUG_MODE: false,
  SHEET_NAME: "Aktienliste",
  EMAIL_RECIPIENT: "matthias.gronwald@gmail.com"
};

/**
 * Main function to check for Golden and Death cross signals
 * Processes data from spreadsheet and sends email notifications
 */
function checkForSignals() {
  const startTime = new Date().getTime();

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    SpreadsheetApp.flush();

    // Get all data at once
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const stockData = data.slice(1);

    const goldenCrossTickers = [];
    const deathCrossTickers = [];
    const now = new Date();

    // Prepare arrays for batch updates
    const newSignals = [];
    const prevSignals = [];
    const timestamps = [];

    // Conditional log
    if (CONFIG.DEBUG_MODE) Logger.log(`Processing ${stockData.length} stocks`);

    // Process all data in memory
    for (let i = 0; i < stockData.length; i++) {
      const row = stockData[i];
      const symbol = row[0];
      const name = row[1];
      const currentPrice = row[2];
      const sma50 = row[3];
      const sma200 = row[4];
      const prevSignal = row[6]; // Previous Signal (column 7)
      const currency = row[8]; // Currency (column I)
      
      // Skip processing if we have invalid data
      if (!symbol || !currentPrice || !sma50 || !sma200) {
        continue;
      }

      // Calculate new signal based on current SMA values
      const newSignal = sma50 > sma200 ? "GOLDEN" : "DEATH";

      // Debug logging
      if (CONFIG.DEBUG_MODE) {
        Logger.log(
          `Row ${i + 2}: ${symbol}, New Signal=${newSignal}, Prev Signal=${prevSignal}, 50 SMA=${sma50}, 200 SMA=${sma200}, Currency=${currency}`
        );
      }

      // Check for a Golden Cross (was DEATH, now GOLDEN)
      if (newSignal === "GOLDEN" && prevSignal === "DEATH") {
        goldenCrossTickers.push(createStockInfo(symbol, name, currentPrice, sma50, sma200, currency));
        row[7] = now; // Update timestamp for display in email
        if (CONFIG.DEBUG_MODE) Logger.log(`→ GOLDEN CROSS detected for ${symbol}`);
      }
      // Check for a Death Cross (was GOLDEN, now DEATH)
      else if (newSignal === "DEATH" && prevSignal === "GOLDEN") {
        deathCrossTickers.push(createStockInfo(symbol, name, currentPrice, sma50, sma200, currency));
        row[7] = now; // Update timestamp for display in email
        if (CONFIG.DEBUG_MODE) Logger.log(`→ DEATH CROSS detected for ${symbol}`);
      }

      // Collect data for batch updates
      newSignals.push([newSignal]);
      prevSignals.push([row[5] || newSignal]);
      timestamps.push([row[7] instanceof Date ? row[7] : null]);
    }

    // Perform batch updates
    if (stockData.length > 0) {
      sheet.getRange(2, 6, stockData.length, 1).setValues(newSignals);
      sheet.getRange(2, 7, stockData.length, 1).setValues(prevSignals);

      // Only update timestamps that are dates
      for (var i = 0; i < timestamps.length; i++) {
        if (timestamps[i][0] !== null) {
          sheet.getRange(i + 2, 8).setValue(timestamps[i][0]);
        }
      }
    }

    // Always log summary information (not too verbose)
    Logger.log(
      `Detected: ${goldenCrossTickers.length} Golden Cross and ${deathCrossTickers.length} Death Cross signals`
    );

    // Send email if there are any changes
    if (goldenCrossTickers.length > 0 || deathCrossTickers.length > 0) {
      sendCrossSignalEmail(goldenCrossTickers, deathCrossTickers, {
        date: now,
        debug: CONFIG.DEBUG_MODE,
        executionTime: new Date().getTime() - startTime,
      });
      Logger.log("Email sent successfully");
    } else {
      Logger.log("No crosses detected, no email sent");
    }

    // Always log total execution time (useful performance metric)
    Logger.log(
      `Execution completed in ${new Date().getTime() - startTime}ms`
    );
  } catch (error) {
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENT,
      subject: "Error in SMA Cross Signals Script",
      body: `An error occurred: ${error.toString()}\n\nStack trace: ${error.stack}`
    });
    Logger.log(`ERROR: ${error}`);
  }
}

/**
 * Creates a stock info object with calculated percentages
 * 
 * @param {string} symbol - Stock ticker symbol
 * @param {string} name - Company name
 * @param {number} price - Current price
 * @param {number} sma50 - 50-day SMA
 * @param {number} sma200 - 200-day SMA
 * @param {string} currency - Currency symbol
 * @return {Object} Formatted stock information
 */
function createStockInfo(symbol, name, price, sma50, sma200, currency) {
  return {
    symbol: symbol,
    name: name,
    price: price,
    sma50: sma50,
    sma200: sma200,
    pctAbove50: ((price / sma50 - 1) * 100).toFixed(2),
    pctAbove200: ((price / sma200 - 1) * 100).toFixed(2),
    currency: currency || "" // Include currency, default to empty string if not available
  };
}

/**
 * Generates and sends an email with cross signal information
 *
 * @param {Array} goldenCrossTickers - Array of stocks with golden cross signals
 * @param {Array} deathCrossTickers - Array of stocks with death cross signals
 * @param {Object} options - Additional options like date and debug info
 */
function sendCrossSignalEmail(goldenCrossTickers, deathCrossTickers, options) {
  const htmlBody = generateCrossSignalHtml(
    goldenCrossTickers,
    deathCrossTickers,
    options
  );
  
  try {
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENT,
      subject: `Golden Cross & Death Cross Signals - ${options.date.toLocaleDateString("en-US")}`,
      htmlBody: htmlBody,
    });
  } catch (error) {
    Logger.log(`Failed to send email: ${error.toString()}`);
    // Try a simpler email as fallback
    try {
      MailApp.sendEmail({
        to: CONFIG.EMAIL_RECIPIENT,
        subject: "Trading Signals Detected (Simplified Email)",
        body: `Golden Cross signals: ${goldenCrossTickers.length}\nDeath Cross signals: ${deathCrossTickers.length}\n\nError sending formatted email: ${error.toString()}`
      });
    } catch (fallbackError) {
      Logger.log(`Failed to send fallback email: ${fallbackError.toString()}`);
    }
  }
}

/**
 * Generates HTML for the cross signal email
 *
 * @param {Array} goldenCrossTickers - Array of stocks with golden cross signals
 * @param {Array} deathCrossTickers - Array of stocks with death cross signals
 * @param {Object} options - Additional options like date and debug info
 * @return {string} The formatted HTML content
 */
function generateCrossSignalHtml(
  goldenCrossTickers,
  deathCrossTickers,
  options
) {
  const now = options.date || new Date();
  const htmlParts = [];

  // Start HTML document with responsive design
  htmlParts.push(
    '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>' +
    '<body style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px;">',
  );
  htmlParts.push(
    `<h2>SMA Cross Signals for ${now.toLocaleDateString("en-US")}</h2>`,
  );
  
  // Add summary section
  htmlParts.push(
    `<p><b>Summary:</b> Found ${goldenCrossTickers.length} buy signals and ${deathCrossTickers.length} sell signals.</p>`
  );

  // Build Golden Cross table
  if (goldenCrossTickers.length > 0) {
    htmlParts.push(
      generateCrossTable(goldenCrossTickers, {
        title:
          '⬆️ <font color="green"><b>GOLDEN CROSS (BUY SIGNALS)</b></font>',
        subtitle: "(50-day SMA crossed above 200-day SMA):",
      }),
    );
  }

  // Build Death Cross table
  if (deathCrossTickers.length > 0) {
    htmlParts.push(
      generateCrossTable(deathCrossTickers, {
        title: '⬇️ <font color="red"><b>DEATH CROSS (SELL SIGNALS)</b></font>',
        subtitle: "(50-day SMA crossed below 200-day SMA):",
      }),
    );
  }

  htmlParts.push("<p><i>This is an automated message.</i></p>");

  // Only include execution time in debug mode
  if (options.debug && options.executionTime) {
    htmlParts.push(
      `<p><i>Execution time: ${options.executionTime}ms</i></p>`,
    );
  }

  htmlParts.push("</body></html>");

  // Join all HTML parts into one string
  return htmlParts.join("");
}

/**
 * Generates HTML table for a specific cross type
 *
 * @param {Array} tickers - Array of stocks with signals
 * @param {Object} options - Table options including title and subtitle
 * @return {string} The formatted HTML table
 */
function generateCrossTable(tickers, options) {
  const htmlParts = [];

  // Table header
  htmlParts.push(
    `<h3>${options.title} ${options.subtitle}</h3>`,
    '<table border="0" cellpadding="5" style="border-collapse: collapse; width: 100%; max-width: 800px;">',
    '<tr style="background-color: #f2f2f2;">',
    '<th align="left">Symbol</th>',
    '<th align="left">Name</th>',
    '<th align="right">Price</th>',
    '<th align="center">Currency</th>', // Added currency column
    '<th align="right">50 SMA</th>',
    '<th align="right">200 SMA</th>',
    '<th align="right">% Above 50</th>',
    '<th align="right">% Above 200</th>',
    "</tr>",
  );

  // Table rows with improved formatting and color indicators
  for (let i = 0; i < tickers.length; i++) {
    const stock = tickers[i];
    const pct50Color = parseFloat(stock.pctAbove50) >= 0 ? "green" : "red";
    const pct200Color = parseFloat(stock.pctAbove200) >= 0 ? "green" : "red";
    
    htmlParts.push(
      `<tr${i % 2 === 1 ? ' style="background-color: #f9f9f9;"' : ""}>`,
      `<td><b>${stock.symbol}</b></td>`,
      `<td>${stock.name}</td>`,
      `<td align="right"><b>${stock.price.toFixed(2)}</b></td>`,
      `<td align="center">${stock.currency}</td>`, // Display currency
      `<td align="right">${stock.sma50.toFixed(2)}</td>`,
      `<td align="right">${stock.sma200.toFixed(2)}</td>`,
      `<td align="right" style="color: ${pct50Color}">${stock.pctAbove50}%</td>`,
      `<td align="right" style="color: ${pct200Color}">${stock.pctAbove200}%</td>`,
      "</tr>",
    );
  }

  htmlParts.push("</table><br>");
  return htmlParts.join("");
}

/**
 * Trigger function that can be scheduled to run hourly
 */
function hourlySignalCheck() {
  checkForSignals();
}
