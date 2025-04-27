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
