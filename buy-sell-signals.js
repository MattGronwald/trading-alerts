// Global logging control
var DEBUG_MODE = false;

function checkForSignals() {
  var startTime = new Date().getTime();

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Aktienliste");
    SpreadsheetApp.flush();

    // Get all data at once
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var stockData = data.slice(1);

    var goldenCrossTickers = [];
    var deathCrossTickers = [];
    var now = new Date();

    // Prepare arrays for batch updates
    var newSignals = [];
    var prevSignals = [];
    var timestamps = [];

    // Conditional log
    if (DEBUG_MODE) Logger.log("Processing " + stockData.length + " stocks");

    // Process all data in memory
    for (var i = 0; i < stockData.length; i++) {
      var row = stockData[i];
      var symbol = row[0];
      var name = row[1];
      var currentPrice = row[2];
      var sma50 = row[3];
      var sma200 = row[4];
      var prevSignal = row[6]; // Previous Signal (column 7)
      var currency = row[8]; // Currency (column I)

      // Calculate new signal based on current SMA values
      var newSignal = sma50 > sma200 ? "GOLDEN" : "DEATH";

      // Debug logging
      if (DEBUG_MODE) {
        Logger.log(
          "Row " +
            (i + 2) +
            ": " +
            symbol +
            ", New Signal=" +
            newSignal +
            ", Prev Signal=" +
            prevSignal +
            ", 50 SMA=" +
            sma50 +
            ", 200 SMA=" +
            sma200 +
            ", Currency=" +
            currency,
        );
      }

      // Check for a Golden Cross (was DEATH, now GOLDEN)
      if (newSignal == "GOLDEN" && prevSignal == "DEATH") {
        goldenCrossTickers.push({
          symbol: symbol,
          name: name,
          price: currentPrice,
          sma50: sma50,
          sma200: sma200,
          pctAbove50: ((currentPrice / sma50 - 1) * 100).toFixed(2),
          pctAbove200: ((currentPrice / sma200 - 1) * 100).toFixed(2),
          currency: currency || "", // Include currency, default to empty string if not available
        });
        row[7] = now; // Update timestamp for display in email
        if (DEBUG_MODE) Logger.log("→ GOLDEN CROSS detected for " + symbol);
      }
      // Check for a Death Cross (was GOLDEN, now DEATH)
      else if (newSignal == "DEATH" && prevSignal == "GOLDEN") {
        deathCrossTickers.push({
          symbol: symbol,
          name: name,
          price: currentPrice,
          sma50: sma50,
          sma200: sma200,
          pctAbove50: ((currentPrice / sma50 - 1) * 100).toFixed(2),
          pctAbove200: ((currentPrice / sma200 - 1) * 100).toFixed(2),
          currency: currency || "", // Include currency, default to empty string if not available
        });
        row[7] = now; // Update timestamp for display in email
        if (DEBUG_MODE) Logger.log("→ DEATH CROSS detected for " + symbol);
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
      "Detected: " +
        goldenCrossTickers.length +
        " Golden Cross and " +
        deathCrossTickers.length +
        " Death Cross signals",
    );

    // Send email if there are any changes
    if (goldenCrossTickers.length > 0 || deathCrossTickers.length > 0) {
      sendCrossSignalEmail(goldenCrossTickers, deathCrossTickers, {
        date: now,
        debug: DEBUG_MODE,
        executionTime: new Date().getTime() - startTime,
      });
      Logger.log("Email sent successfully");
    } else {
      Logger.log("No crosses detected, no email sent");
    }

    // Always log total execution time (useful performance metric)
    Logger.log(
      "Execution completed in " + (new Date().getTime() - startTime) + "ms",
    );
  } catch (error) {
    MailApp.sendEmail({
      to: "matthias.gronwald@gmail.com",
      subject: "Error in SMA Cross Signals Script",
      body:
        "An error occurred: " +
        error.toString() +
        "\n\nStack trace: " +
        error.stack,
    });
    Logger.log("ERROR: " + error);
  }
}

/**
 * Generates and sends an email with cross signal information
 *
 * @param {Array} goldenCrossTickers - Array of stocks with golden cross signals
 * @param {Array} deathCrossTickers - Array of stocks with death cross signals
 * @param {Object} options - Additional options like date and debug info
 */
function sendCrossSignalEmail(goldenCrossTickers, deathCrossTickers, options) {
  var htmlBody = generateCrossSignalHtml(
    goldenCrossTickers,
    deathCrossTickers,
    options,
  );

  MailApp.sendEmail({
    to: "matthias.gronwald@gmail.com",
    subject: "Golden Cross & Death Cross Signals",
    htmlBody: htmlBody,
  });
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
  options,
) {
  var now = options.date || new Date();
  var htmlParts = [];

  // Start HTML document
  htmlParts.push(
    '<!DOCTYPE html><html><body style="font-family: Arial, sans-serif;">',
  );
  htmlParts.push(
    "<h2>SMA Cross Signals for " + now.toLocaleDateString("en-US") + "</h2>",
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
      "<p><i>Execution time: " + options.executionTime + "ms</i></p>",
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
  var htmlParts = [];

  // Table header
  htmlParts.push(
    "<h3>" + options.title + " " + options.subtitle + "</h3>",
    '<table border="0" cellpadding="5" style="border-collapse: collapse;">',
    '<tr style="background-color: #f2f2f2;">',
    '<th align="left">Symbol</th>',
    '<th align="left">Name</th>',
    '<th align="right">Current Price</th>',
    '<th align="center">Currency</th>', // Added currency column
    '<th align="right">50 SMA</th>',
    '<th align="right">200 SMA</th>',
    '<th align="right">% Above 50</th>',
    '<th align="right">% Above 200</th>',
    "</tr>",
  );

  // Table rows
  for (var i = 0; i < tickers.length; i++) {
    var stock = tickers[i];
    htmlParts.push(
      "<tr" + (i % 2 == 1 ? ' style="background-color: #f9f9f9;"' : "") + ">",
      "<td><b>" + stock.symbol + "</b></td>",
      "<td>" + stock.name + "</td>",
      '<td align="right"><b>' + stock.price.toFixed(2) + "</b></td>",
      '<td align="center">' + stock.currency + "</td>", // Display currency
      '<td align="right">' + stock.sma50.toFixed(2) + "</td>",
      '<td align="right">' + stock.sma200.toFixed(2) + "</td>",
      '<td align="right">' + stock.pctAbove50 + "%</td>",
      '<td align="right">' + stock.pctAbove200 + "%</td>",
      "</tr>",
    );
  }

  htmlParts.push("</table><br>");
  return htmlParts.join("");
}
