function checkAlarm() {
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
          sma50: sma50,
          sma200: sma200,
        });
        sheet.getRange(i + 1, 8).setValue(now); // Update timestamp
      }

      // Check for Death Cross (50 SMA crosses below 200 SMA)
      if (currentSignal == "DEATH" && prevSignal == "GOLDEN") {
        deathCrossTickers.push({
          symbol: symbol,
          name: name,
          sma50: sma50,
          sma200: sma200,
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

        htmlBody += '<table border="0" cellpadding="5">';
        htmlBody +=
          '<tr><th align="left">Symbol</th><th align="left">Name</th><th align="right">50 SMA</th><th align="right">200 SMA</th></tr>';

        for (var i = 0; i < goldenCrossTickers.length; i++) {
          htmlBody += "<tr>";
          htmlBody += "<td><b>" + goldenCrossTickers[i].symbol + "</b></td>";
          htmlBody += "<td>" + goldenCrossTickers[i].name + "</td>";
          htmlBody +=
            '<td align="right">' +
            goldenCrossTickers[i].sma50.toFixed(2) +
            "</td>";
          htmlBody +=
            '<td align="right">' +
            goldenCrossTickers[i].sma200.toFixed(2) +
            "</td>";
          htmlBody += "</tr>";
        }

        htmlBody += "</table><br>";
      }

      if (deathCrossTickers.length > 0) {
        htmlBody +=
          '<h3>⬇️ <font color="red"><b>DEATH CROSS (SELL SIGNALS)</b></font> ';
        htmlBody += "(50-day SMA crossed below 200-day SMA):</h3>";

        htmlBody += '<table border="0" cellpadding="5">';
        htmlBody +=
          '<tr><th align="left">Symbol</th><th align="left">Name</th><th align="right">50 SMA</th><th align="right">200 SMA</th></tr>';

        for (var i = 0; i < deathCrossTickers.length; i++) {
          htmlBody += "<tr>";
          htmlBody += "<td><b>" + deathCrossTickers[i].symbol + "</b></td>";
          htmlBody += "<td>" + deathCrossTickers[i].name + "</td>";
          htmlBody +=
            '<td align="right">' +
            deathCrossTickers[i].sma50.toFixed(2) +
            "</td>";
          htmlBody +=
            '<td align="right">' +
            deathCrossTickers[i].sma200.toFixed(2) +
            "</td>";
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

function updateRegime() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "SPY history"; // Name des richtigen Sheets
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Das Sheet "${sheetName}" wurde nicht gefunden.`);
  }

  // Sicherstellen, dass alle Berechnungen abgeschlossen sind
  SpreadsheetApp.flush();

  // Zellen definieren
  const maxValueCell = sheet.getRange("D2").getValue(); // Höchstwert der letzten 3 Jahre
  const currentCloseCell = sheet.getRange("B" + sheet.getLastRow()).getValue(); // Aktueller Close-Wert
  const regimeCells = {
    A: sheet.getRange("H2"),
    B: sheet.getRange("H3"),
    C: sheet.getRange("H4"),
  };

  const triggerDateCell = sheet.getRange("K1"); // Datum des Regimewechsels
  const triggerCloseCell = sheet.getRange("K2"); // Auslösender Schlusskurs

  // Schwellenwerte berechnen
  const threshold20 = maxValueCell * 0.8; // 20% unter Höchstwert
  const threshold40 = maxValueCell * 0.6; // 40% unter Höchstwert

  // Aktuelle Regime-Zustände
  const currentRegimeA = regimeCells.A.getValue() === "Ja";
  const currentRegimeB = regimeCells.B.getValue() === "Ja";
  const currentRegimeC = regimeCells.C.getValue() === "Ja";

  // Variable für das neue Regime
  let newRegime = "";
  let oldRegime = currentRegimeA
    ? "A"
    : currentRegimeB
      ? "B"
      : currentRegimeC
        ? "C"
        : "Unbekannt";

  // Logik für Regimewechsel
  if (currentCloseCell <= threshold40) {
    // Wechsel zu Regime C
    if (!currentRegimeC) {
      regimeCells.A.setValue("");
      regimeCells.B.setValue("");
      regimeCells.C.setValue("Ja");
      newRegime = "C - Eskalation der Eigenkapitalknappheit (Krise)";
    }
  } else if (
    currentCloseCell > threshold40 &&
    currentCloseCell <= threshold20
  ) {
    // Wechsel zu Regime B
    if (!currentRegimeB) {
      regimeCells.A.setValue("");
      regimeCells.B.setValue("Ja");
      regimeCells.C.setValue("");
      newRegime = "B - Eigenkapitalknappheit (Krise)";
    }
  } else if (currentCloseCell > threshold20) {
    // Wechsel zu Regime A
    if (!currentRegimeA) {
      regimeCells.A.setValue("Ja");
      regimeCells.B.setValue("");
      regimeCells.C.setValue("");
      newRegime = "A - Normal";
    }
  }

  // E-Mail-Benachrichtigung senden, wenn es einen Regimewechsel gibt
  if (newRegime) {
    const currentDate = new Date();
    triggerDateCell.setValue(currentDate); // Datum speichern
    triggerCloseCell.setValue(currentCloseCell); // Schlusskurs speichern
    sendEmailNotification(newRegime, oldRegime, currentCloseCell, maxValueCell);
  }
}

// Funktion zum Senden der E-Mail-Benachrichtigung
function sendEmailNotification(newRegime, oldRegime, currentClose, maxValue) {
  const recipient = "matthias.gronwald@gmail.com"; // E-Mail-Adresse ersetzen
  const subject = `S&P500 Regimewechsel: Von Regime ${oldRegime} zu ${newRegime}`;
  const body = `
    Es gab einen Regimewechsel im System.

    Alter Regimezustand: ${oldRegime}
    Neuer Regimezustand: ${newRegime}
    Aktueller Kurs: ${currentClose.toFixed(2)}
    3-Jahres-Hoch: ${maxValue.toFixed(2)}

    Beste Grüße,
    Dein Google Sheets System
  `;

  // E-Mail senden
  MailApp.sendEmail(recipient, subject, body);
}
