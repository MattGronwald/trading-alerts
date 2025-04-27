function checkForSignals() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Aktienliste";
    var sheet = spreadsheet.getSheetByName(sheetName);

    SpreadsheetApp.flush();

    var data = sheet.getDataRange().getValues();
    var goldenCrossTickers = [];
    var deathCrossTickers = [];
    var now = new Date();

    // First update previous states
    for (var i = 1; i < data.length; i++) {
      var currentSignal = data[i][5]; // Cross Signal column (index 5)
      sheet.getRange(i + 1, 7).setValue(currentSignal); // Previous Signal column (index 6)
    }

    // Then check for crosses and send notifications
    for (var i = 1; i < data.length; i++) {
      var symbol = data[i][0];
      var name = data[i][1];
      var currentPrice = data[i][2]; // Current price column (index 2)
      var sma50 = data[i][3]; // 50 SMA column (index 3)
      var sma200 = data[i][4]; // 200 SMA column (index 4)
      var currentSignal = "";

      // Determine current cross signal
      if (sma50 > sma200) {
        currentSignal = "GOLDEN";
      } else if (sma50 < sma200) {
        currentSignal = "DEATH";
      }

      var prevSignal = data[i][6]; // Previous Signal column (index 6)

      // Update the current signal in the sheet
      sheet.getRange(i + 1, 6).setValue(currentSignal);

      // Check for Golden Cross (50 SMA crosses above 200 SMA)
      if (currentSignal == "GOLDEN" && prevSignal == "DEATH") {
        goldenCrossTickers.push({
          symbol: symbol,
          name: name,
          price: currentPrice,
          sma50: sma50,
          sma200: sma200,
          pctAbove50: ((currentPrice / sma50 - 1) * 100).toFixed(2),
          pctAbove200: ((currentPrice / sma200 - 1) * 100).toFixed(2),
        });
        sheet.getRange(i + 1, 8).setValue(now); // Update timestamp
      }

      // Check for Death Cross (50 SMA crosses below 200 SMA)
      if (currentSignal == "DEATH" && prevSignal == "GOLDEN") {
        deathCrossTickers.push({
          symbol: symbol,
          name: name,
          price: currentPrice,
          sma50: sma50,
          sma200: sma200,
          pctAbove50: ((currentPrice / sma50 - 1) * 100).toFixed(2),
          pctAbove200: ((currentPrice / sma200 - 1) * 100).toFixed(2),
        });
        sheet.getRange(i + 1, 8).setValue(now); // Update timestamp
      }
    }

    // Send email if there are any changes
    if (goldenCrossTickers.length > 0 || deathCrossTickers.length > 0) {
      var htmlBody =
        '<!DOCTYPE html><html><body style="font-family: Arial, sans-serif;">';
      htmlBody +=
        "<h2>SMA Cross Signals for " +
        now.toLocaleDateString("en-US") +
        "</h2>";

      if (goldenCrossTickers.length > 0) {
        htmlBody +=
          '<h3>⬆️ <font color="green"><b>GOLDEN CROSS (BUY SIGNALS)</b></font> ';
        htmlBody += "(50-day SMA crossed above 200-day SMA):</h3>";

        htmlBody +=
          '<table border="0" cellpadding="5" style="border-collapse: collapse;">';
        htmlBody += '<tr style="background-color: #f2f2f2;">';
        htmlBody += '<th align="left">Symbol</th>';
        htmlBody += '<th align="left">Name</th>';
        htmlBody += '<th align="right">Current Price</th>';
        htmlBody += '<th align="right">50 SMA</th>';
        htmlBody += '<th align="right">200 SMA</th>';
        htmlBody += '<th align="right">% Above 50</th>';
        htmlBody += '<th align="right">% Above 200</th>';
        htmlBody += "</tr>";

        for (var i = 0; i < goldenCrossTickers.length; i++) {
          htmlBody +=
            "<tr" +
            (i % 2 == 1 ? ' style="background-color: #f9f9f9;"' : "") +
            ">";
          htmlBody += "<td><b>" + goldenCrossTickers[i].symbol + "</b></td>";
          htmlBody += "<td>" + goldenCrossTickers[i].name + "</td>";
          htmlBody +=
            '<td align="right"><b>' +
            goldenCrossTickers[i].price.toFixed(2) +
            "</b></td>";
          htmlBody +=
            '<td align="right">' +
            goldenCrossTickers[i].sma50.toFixed(2) +
            "</td>";
          htmlBody +=
            '<td align="right">' +
            goldenCrossTickers[i].sma200.toFixed(2) +
            "</td>";
          htmlBody +=
            '<td align="right">' + goldenCrossTickers[i].pctAbove50 + "%</td>";
          htmlBody +=
            '<td align="right">' + goldenCrossTickers[i].pctAbove200 + "%</td>";
          htmlBody += "</tr>";
        }

        htmlBody += "</table><br>";
      }

      if (deathCrossTickers.length > 0) {
        htmlBody +=
          '<h3>⬇️ <font color="red"><b>DEATH CROSS (SELL SIGNALS)</b></font> ';
        htmlBody += "(50-day SMA crossed below 200-day SMA):</h3>";

        htmlBody +=
          '<table border="0" cellpadding="5" style="border-collapse: collapse;">';
        htmlBody += '<tr style="background-color: #f2f2f2;">';
        htmlBody += '<th align="left">Symbol</th>';
        htmlBody += '<th align="left">Name</th>';
        htmlBody += '<th align="right">Current Price</th>';
        htmlBody += '<th align="right">50 SMA</th>';
        htmlBody += '<th align="right">200 SMA</th>';
        htmlBody += '<th align="right">% Above 50</th>';
        htmlBody += '<th align="right">% Above 200</th>';
        htmlBody += "</tr>";

        for (var i = 0; i < deathCrossTickers.length; i++) {
          htmlBody +=
            "<tr" +
            (i % 2 == 1 ? ' style="background-color: #f9f9f9;"' : "") +
            ">";
          htmlBody += "<td><b>" + deathCrossTickers[i].symbol + "</b></td>";
          htmlBody += "<td>" + deathCrossTickers[i].name + "</td>";
          htmlBody +=
            '<td align="right"><b>' +
            deathCrossTickers[i].price.toFixed(2) +
            "</b></td>";
          htmlBody +=
            '<td align="right">' +
            deathCrossTickers[i].sma50.toFixed(2) +
            "</td>";
          htmlBody +=
            '<td align="right">' +
            deathCrossTickers[i].sma200.toFixed(2) +
            "</td>";
          htmlBody +=
            '<td align="right">' + deathCrossTickers[i].pctAbove50 + "%</td>";
          htmlBody +=
            '<td align="right">' + deathCrossTickers[i].pctAbove200 + "%</td>";
          htmlBody += "</tr>";
        }

        htmlBody += "</table>";
      }

      htmlBody += "<p><i>This is an automated message.</i></p>";
      htmlBody += "</body></html>";

      MailApp.sendEmail({
        to: "matthias.gronwald@gmail.com",
        subject: "Golden Cross & Death Cross Signals",
        htmlBody: htmlBody,
      });
    }
  } catch (error) {
    MailApp.sendEmail({
      to: "matthias.gronwald@gmail.com",
      subject: "Error in SMA Cross Signals Script",
      body: "An error occurred: " + error.toString(),
    });
    Logger.log(error);
  }
}
