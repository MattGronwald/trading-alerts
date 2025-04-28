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
            sma200,
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
        });
        row[7] = now; // Update timestamp for display in email
        if (DEBUG_MODE) Logger.log("→ DEATH CROSS detected for " + symbol);
      }

      // Set the current signal column (for next run)
      sheet.getRange(i + 2, 6).setValue(newSignal);
      // Set the previous signal column (for this run)
      sheet.getRange(i + 2, 7).setValue(row[5] || newSignal);
      // Update timestamp if needed
      if (row[7] instanceof Date) {
        sheet.getRange(i + 2, 8).setValue(row[7]);
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
      var htmlParts = [];
      htmlParts.push(
        '<!DOCTYPE html><html><body style="font-family: Arial, sans-serif;">',
      );
      htmlParts.push(
        "<h2>SMA Cross Signals for " +
          now.toLocaleDateString("en-US") +
          "</h2>",
      );

      // Build Golden Cross table
      if (goldenCrossTickers.length > 0) {
        htmlParts.push(
          '<h3>⬆️ <font color="green"><b>GOLDEN CROSS (BUY SIGNALS)</b></font> ',
          "(50-day SMA crossed above 200-day SMA):</h3>",
        );

        htmlParts.push(
          '<table border="0" cellpadding="5" style="border-collapse: collapse;">',
          '<tr style="background-color: #f2f2f2;">',
          '<th align="left">Symbol</th>',
          '<th align="left">Name</th>',
          '<th align="right">Current Price</th>',
          '<th align="right">50 SMA</th>',
          '<th align="right">200 SMA</th>',
          '<th align="right">% Above 50</th>',
          '<th align="right">% Above 200</th>',
          "</tr>",
        );

        for (var i = 0; i < goldenCrossTickers.length; i++) {
          var stock = goldenCrossTickers[i];
          htmlParts.push(
            "<tr" +
              (i % 2 == 1 ? ' style="background-color: #f9f9f9;"' : "") +
              ">",
            "<td><b>" + stock.symbol + "</b></td>",
            "<td>" + stock.name + "</td>",
            '<td align="right"><b>' + stock.price.toFixed(2) + "</b></td>",
            '<td align="right">' + stock.sma50.toFixed(2) + "</td>",
            '<td align="right">' + stock.sma200.toFixed(2) + "</td>",
            '<td align="right">' + stock.pctAbove50 + "%</td>",
            '<td align="right">' + stock.pctAbove200 + "%</td>",
            "</tr>",
          );
        }

        htmlParts.push("</table><br>");
      }

      // Build Death Cross table
      if (deathCrossTickers.length > 0) {
        htmlParts.push(
          '<h3>⬇️ <font color="red"><b>DEATH CROSS (SELL SIGNALS)</b></font> ',
          "(50-day SMA crossed below 200-day SMA):</h3>",
        );

        htmlParts.push(
          '<table border="0" cellpadding="5" style="border-collapse: collapse;">',
          '<tr style="background-color: #f2f2f2;">',
          '<th align="left">Symbol</th>',
          '<th align="left">Name</th>',
          '<th align="right">Current Price</th>',
          '<th align="right">50 SMA</th>',
          '<th align="right">200 SMA</th>',
          '<th align="right">% Above 50</th>',
          '<th align="right">% Above 200</th>',
          "</tr>",
        );

        for (var i = 0; i < deathCrossTickers.length; i++) {
          var stock = deathCrossTickers[i];
          htmlParts.push(
            "<tr" +
              (i % 2 == 1 ? ' style="background-color: #f9f9f9;"' : "") +
              ">",
            "<td><b>" + stock.symbol + "</b></td>",
            "<td>" + stock.name + "</td>",
            '<td align="right"><b>' + stock.price.toFixed(2) + "</b></td>",
            '<td align="right">' + stock.sma50.toFixed(2) + "</td>",
            '<td align="right">' + stock.sma200.toFixed(2) + "</td>",
            '<td align="right">' + stock.pctAbove50 + "%</td>",
            '<td align="right">' + stock.pctAbove200 + "%</td>",
            "</tr>",
          );
        }

        htmlParts.push("</table>");
      }

      htmlParts.push("<p><i>This is an automated message.</i></p>");

      // Only include execution time in debug mode
      if (DEBUG_MODE) {
        htmlParts.push(
          "<p><i>Execution time: " +
            (new Date().getTime() - startTime) +
            "ms</i></p>",
        );
      }

      htmlParts.push("</body></html>");

      // Join all HTML parts into one string
      var htmlBody = htmlParts.join("");

      MailApp.sendEmail({
        to: "matthias.gronwald@gmail.com",
        subject: "Golden Cross & Death Cross Signals",
        htmlBody: htmlBody,
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
