// ─────────────────────────────────────────────────────────────────────────────
// Birdie Vrienden – Google Apps Script webhook
//
// Deployment:
//   1. Open het Google Sheet waarin je data wilt opslaan.
//   2. Kies Extensions > Apps Script.
//   3. Vervang de inhoud van Code.gs door deze code.
//   4. Sla op (Ctrl+S).
//   5. Klik op Deploy > New deployment.
//   6. Kies Type: Web App.
//   7. Stel in:
//        Execute as: Me
//        Who has access: Anyone
//   8. Klik Deploy en kopieer de Web App URL.
//   9. Plak die URL in index.html bij APPS_SCRIPT_URL.
// ─────────────────────────────────────────────────────────────────────────────

var SHEET_ID   = "1OFt91SSoiKGjSYIfT_mmaTUGJ2yKZSAH55oXfFvBwR8";
var SHEET_NAME = "Aanmeldingen";

var COLUMNS = ["timestamp", "naam", "email", "telefoon", "per_birdie", "max_seizoen", "whatsapp", "actief"];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SHEET_ID);

    if (data.type === "eenmalig") {
      var eenSheet = ss.getSheetByName("Eenmalige Sponsors");
      if (!eenSheet) {
        eenSheet = ss.insertSheet("Eenmalige Sponsors");
        eenSheet.appendRow(["NAAM", "EMAIL", "TELEFOON", "BEDRAG (€)", "BETAALD"]);
        eenSheet.setFrozenRows(1);
      }
      eenSheet.appendRow([
        data.naam     || "",
        data.email    || "",
        data.telefoon || "",
        data.bedrag   != null ? parseFloat(data.bedrag) : "",
        "Nee",
      ]);
      return buildResponse({ status: "ok", message: "Eenmalige sponsoring opgeslagen." });
    }

    // Bestaand gedrag: per-birdie naar Aanmeldingen
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(COLUMNS.map(function(c) { return c.toUpperCase(); }));
      sheet.setFrozenRows(1);
    }

    var row = [
      new Date(),
      data.naam        || "",
      data.email       || "",
      data.telefoon    || "",
      data.per_birdie  != null ? data.per_birdie : "",
      data.max_seizoen != null ? data.max_seizoen : "",
      data.whatsapp    ? "Ja" : "Nee",
      "Ja",
    ];

    sheet.appendRow(row);
    return buildResponse({ status: "ok", message: "Aanmelding opgeslagen." });

  } catch (err) {
    return buildResponse({ status: "error", message: err.toString() }, true);
  }
}

function doGet(e) {
  return buildResponse({ status: "ok", message: "Birdie Vrienden webhook actief." });
}

function buildResponse(obj, isError) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────────────────
// Menu
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Birdie Vrienden')
    .addItem('Setup Overzicht + Birdies',        'setupOverzicht')
    .addItem('Setup Betalingen',                 'setupBetalingen')
    .addItem('Setup Eenmalige Sponsors',          'setupEenmaligSponsors')
    .addItem('Refresh Overzicht formulas',       'refreshOverzichtFormulas')
    .addItem('Setup Aanmeldingen (ACTIEF kolom)','setupAanmeldingenKolommen')
    .addItem('Voeg datumpicker toe aan Birdies', 'setupBirdiesDatumPicker')
    .addSeparator()
    .addItem('Sync toernooien naar Betalingen',  'syncToernooien')
    .addSeparator()
    .addItem('Genereer betaalverzoek',           'genereerBetaalverzoek')
    .addToUi();
}

// Voegt datumvalidatie toe aan Birdies!A zodat de datumpicker verschijnt.
// Raakt geen bestaande data aan.
function setupBirdiesDatumPicker() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Birdies");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Birdies-tabblad niet gevonden.");
    return;
  }
  var range = sheet.getRange("A2:A200");
  range.setNumberFormat("dd-mm-yyyy");
  range.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDateAfter(new Date(2000, 0, 1))
      .setAllowInvalid(false)
      .setHelpText("Kies een datum via de kalender.")
      .build()
  );
  SpreadsheetApp.getUi().alert("Datumpicker toegevoegd aan de Birdies-tab.");
}

// ─────────────────────────────────────────────────────────────────────────────
// Aanmeldingen – ACTIEF kolom toevoegen aan bestaand sheet
// ─────────────────────────────────────────────────────────────────────────────
function setupAanmeldingenKolommen() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Aanmeldingen");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Aanmeldingen-tabblad niet gevonden.");
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var actiefCol = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h].toString().toUpperCase() === "ACTIEF") { actiefCol = h + 1; break; }
  }

  if (actiefCol === -1) {
    actiefCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, actiefCol)
      .setValue("ACTIEF")
      .setBackground("#1e1b2e").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setColumnWidth(actiefCol, 80);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var range  = sheet.getRange(2, actiefCol, lastRow - 1, 1);
    var values = range.getValues();
    for (var r = 0; r < values.length; r++) {
      if (values[r][0] === "") values[r][0] = "Ja";
    }
    range.setValues(values);
    range.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Ja", "Nee"], true).build()
    );
  }

  SpreadsheetApp.getUi().alert(
    "ACTIEF kolom toegevoegd aan Aanmeldingen.\n" +
    "Zet een sponsor op 'Nee' om hem te deactiveren.\n\n" +
    "Voer daarna 'Sync toernooien' uit om Betalingen bij te werken."
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Betalingen – aanmaken
// ─────────────────────────────────────────────────────────────────────────────
function setupBetalingen() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var sheet = ss.getSheetByName("Betalingen");
  if (sheet) {
    var ui = SpreadsheetApp.getUi();
    var antwoord = ui.alert(
      "Betalingen-tabblad bestaat al",
      "Wil je het volledig opnieuw aanmaken? Alle betalingsdata gaat verloren.",
      ui.ButtonSet.YES_NO
    );
    if (antwoord !== ui.Button.YES) return;
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    sheet.setConditionalFormatRules([]);
  } else {
    sheet = ss.insertSheet("Betalingen");
  }

  sheet.appendRow(["NAAM"]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 180);
  sheet.getRange("A1")
    .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");

  syncToernooien_(ss, sheet);
  SpreadsheetApp.getUi().alert("Betalingen-tabblad aangemaakt! Gebruik 'Sync toernooien' om het up-to-date te houden.");
}

// ─────────────────────────────────────────────────────────────────────────────
// Eenmalige Sponsors – aanmaken
// ─────────────────────────────────────────────────────────────────────────────
function setupEenmaligSponsors() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var sheet = ss.getSheetByName("Eenmalige Sponsors");
  if (sheet) {
    var ui = SpreadsheetApp.getUi();
    var antwoord = ui.alert(
      "Eenmalige Sponsors-tabblad bestaat al",
      "Wil je het volledig opnieuw aanmaken? Alle data gaat verloren.",
      ui.ButtonSet.YES_NO
    );
    if (antwoord !== ui.Button.YES) return;
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    sheet.setConditionalFormatRules([]);
  } else {
    sheet = ss.insertSheet("Eenmalige Sponsors");
  }

  // ── Headers ────────────────────────────────────────────────────────────────
  var headers = ["NAAM", "EMAIL", "TELEFOON", "BEDRAG (€)", "BETAALD"];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange("A1:E1")
    .setBackground("#9D174D")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  // ── Kolombreedtes ──────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 180); // NAAM
  sheet.setColumnWidth(2, 200); // EMAIL
  sheet.setColumnWidth(3, 130); // TELEFOON
  sheet.setColumnWidth(4, 130); // BEDRAG
  sheet.setColumnWidth(5, 100); // BETAALD

  // ── BEDRAG opmaak ──────────────────────────────────────────────────────────
  sheet.getRange("D2:D200").setNumberFormat('€#,##0.00');

  // ── BETAALD dropdown + voorwaardelijke opmaak ──────────────────────────────
  var betaaldRange = sheet.getRange("E2:E200");
  betaaldRange.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Ja", "Nee"], true)
      .build()
  );
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Ja")
      .setBackground("#d1fae5").setFontColor("#065f46")
      .setRanges([betaaldRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Nee")
      .setBackground("#fee2e2").setFontColor("#991b1b")
      .setRanges([betaaldRange])
      .build()
  ]);

  // ── Totaalrij (rij 202, vast onder datagebied E2:E200) ────────────────────
  sheet.getRange("A202").setValue("TOTAAL OPEN");
  sheet.getRange("D202").setFormula('=SUMIF(E2:E200,"Nee",D2:D200)');
  sheet.getRange("A202:E202")
    .setBackground("#fdf2f8")
    .setFontWeight("bold");
  sheet.getRange("D202").setNumberFormat('€#,##0.00');

  SpreadsheetApp.getUi().alert("Eenmalige Sponsors-tabblad aangemaakt.");
}

// ─────────────────────────────────────────────────────────────────────────────
// Sync – publieke wrapper
// ─────────────────────────────────────────────────────────────────────────────
function syncToernooien() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Betalingen");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Maak eerst het Betalingen-tabblad aan via Setup Betalingen.");
    return;
  }
  var result = syncToernooien_(ss, sheet);
  var msg = [];
  if (result.newSponsors > 0) msg.push(result.newSponsors + " nieuwe sponsor(s) toegevoegd.");
  if (result.added > 0)       msg.push(result.added + " groep(en) toegevoegd.");
  if (result.removed > 0)     msg.push(result.removed + " groep(en) verwijderd.");
  SpreadsheetApp.getUi().alert(msg.length > 0 ? msg.join(" ") : "Alles is al up-to-date.");
}

// ─────────────────────────────────────────────────────────────────────────────
// Sync – kern (idempotent)
//
// Wat het doet:
//  1. Groepen van 3 toernooien bouwen uit Birdies
//  2. Sponsormap bouwen uit Aanmeldingen (incl. actief-flag en joindatum)
//  3. Totaalrij verwijderen vóór verwerking
//  4. Nieuwe actieve sponsors toevoegen aan Betalingen
//  5. Ongeldige groepkolommen verwijderen
//  6. Ontbrekende groepkolommen aanmaken
//  7. Alle Open bedragen herberekenen (joindatum, cap, per_birdie-wijzigingen)
//     Inactieve sponsors krijgen nieuwe groepen als "Inactief" met €0
//  8. Dropdown + voorwaardelijke opmaak bijwerken
//  9. Totaalrij toevoegen
// ─────────────────────────────────────────────────────────────────────────────
function syncToernooien_(ss, betalingen) {
  var birdiesSheet = ss.getSheetByName("Birdies");
  var aanmSheet    = ss.getSheetByName("Aanmeldingen");
  if (!birdiesSheet || !aanmSheet) return { added: 0, removed: 0, newSponsors: 0 };

  // ── 1. Toernooien + groepen ───────────────────────────────────────────────
  var birdiesData    = birdiesSheet.getDataRange().getValues();
  var allTournaments = [];
  for (var i = 1; i < birdiesData.length; i++) {
    allTournaments.push({
      date:    birdiesData[i][0] ? new Date(birdiesData[i][0]) : null,
      birdies: Number(birdiesData[i][2]) || 0
    });
  }

  var groups = [];
  for (var g = 0; g * 3 + 2 < allTournaments.length; g++) {
    var from = g * 3 + 1;
    var to   = from + 2;
    groups.push({
      label:       "T" + from + "-T" + to,
      tournaments: [allTournaments[from - 1], allTournaments[from], allTournaments[from + 1]]
    });
  }
  var validLabels = {};
  for (var vl = 0; vl < groups.length; vl++) validLabels[groups[vl].label] = true;

  // ── 2. Sponsormap (dynamische kolomdetectie) ──────────────────────────────
  var aanmData    = aanmSheet.getDataRange().getValues();
  var aanmHeaders = aanmData[0];
  var ci          = {};
  for (var hh = 0; hh < aanmHeaders.length; hh++) {
    ci[aanmHeaders[hh].toString().toUpperCase()] = hh;
  }

  var sponsorMap      = {};
  var orderedSponsors = [];
  for (var s = 1; s < aanmData.length; s++) {
    var naam = aanmData[s][ci['NAAM'] !== undefined ? ci['NAAM'] : 1];
    if (!naam) continue;
    var joinDt   = aanmData[s][0] ? new Date(aanmData[s][0]) : null;
    var joinDate = joinDt
      ? new Date(joinDt.getFullYear(), joinDt.getMonth(), joinDt.getDate())
      : null;
    var actiefVal = ci['ACTIEF'] !== undefined ? aanmData[s][ci['ACTIEF']] : "";
    sponsorMap[naam] = {
      perBirdie:  parseFloat(aanmData[s][ci['PER_BIRDIE']  !== undefined ? ci['PER_BIRDIE']  : 4]) || 0,
      maxSeizoen: (function(v) { return v !== "" && v != null ? parseFloat(v) : null; })(
                    aanmData[s][ci['MAX_SEIZOEN'] !== undefined ? ci['MAX_SEIZOEN'] : 5]),
      joinDate:   joinDate,
      actief:     actiefVal !== "Nee"
    };
    orderedSponsors.push(naam);
  }

  // ── 3. Totaalrij verwijderen als die bestaat ──────────────────────────────
  var lastRowNow = betalingen.getLastRow();
  if (lastRowNow > 1) {
    var lastCellVal = betalingen.getRange(lastRowNow, 1).getValue().toString();
    if (lastCellVal === "TOTAAL OPEN") {
      betalingen.deleteRow(lastRowNow);
    }
  }

  // ── 4. Nieuwe actieve sponsors toevoegen ─────────────────────────────────
  var betalingenLastRow  = betalingen.getLastRow();
  var existingSponsorRow = {};
  if (betalingenLastRow > 1) {
    var sponsorCol = betalingen.getRange(2, 1, betalingenLastRow - 1, 1).getValues();
    for (var sp = 0; sp < sponsorCol.length; sp++) {
      if (sponsorCol[sp][0]) existingSponsorRow[sponsorCol[sp][0]] = sp + 2;
    }
  }
  var newSponsors = 0;
  for (var os = 0; os < orderedSponsors.length; os++) {
    var osNaam = orderedSponsors[os];
    if (!existingSponsorRow[osNaam] && sponsorMap[osNaam].actief) {
      betalingen.appendRow([osNaam]);
      existingSponsorRow[osNaam] = betalingen.getLastRow();
      newSponsors++;
    }
  }

  // ── 5. Ongeldige groepkolommen verwijderen (rechts → links) ──────────────
  var currentHeaders = betalingen.getRange(1, 1, 1, betalingen.getLastColumn()).getValues()[0];
  var removed = 0;
  for (var c = currentHeaders.length - 1; c >= 1; c--) {
    var chdr = currentHeaders[c].toString();
    if (chdr.indexOf(" BETAALD") !== -1) {
      var glabel = chdr.replace(" BETAALD", "").trim();
      if (!validLabels[glabel]) {
        betalingen.deleteColumn(c + 1);
        betalingen.deleteColumn(c);
        removed++;
      }
    }
  }

  // ── 6. Ontbrekende groepkolommen aanmaken ────────────────────────────────
  var updatedHeaders = betalingen.getRange(1, 1, 1, betalingen.getLastColumn()).getValues()[0];
  var groupColMap    = {};
  for (var uh = 0; uh < updatedHeaders.length; uh++) {
    var uhdr = updatedHeaders[uh].toString();
    if (uhdr.indexOf(" BEDRAG") !== -1) {
      var lbl = uhdr.replace(" BEDRAG", "").trim();
      groupColMap[lbl] = { bedragCol: uh + 1, betaaldCol: uh + 2 };
    }
  }
  var added = 0;
  for (var t = 0; t < groups.length; t++) {
    if (!groupColMap[groups[t].label]) {
      var bedragCol  = betalingen.getLastColumn() + 1;
      var betaaldCol = bedragCol + 1;
      betalingen.getRange(1, bedragCol).setValue(groups[t].label + " BEDRAG");
      betalingen.getRange(1, betaaldCol).setValue(groups[t].label + " BETAALD");
      betalingen.getRange(1, bedragCol, 1, 2)
        .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");
      betalingen.setColumnWidth(bedragCol, 130);
      betalingen.setColumnWidth(betaaldCol, 120);
      groupColMap[groups[t].label] = { bedragCol: bedragCol, betaaldCol: betaaldCol };
      added++;
    }
  }

  // ── 7. Herbereken alle Open/lege bedragen per sponsor per groep ───────────
  var totalRows = betalingen.getLastRow() - 1;
  if (totalRows < 1) return { added: added, removed: removed, newSponsors: newSponsors };

  var allData = betalingen.getRange(2, 1, totalRows, betalingen.getLastColumn()).getValues();

  for (var row = 0; row < allData.length; row++) {
    var naam    = allData[row][0];
    if (!naam) continue;
    var sponsor = sponsorMap[naam];
    if (!sponsor) continue;

    var capRunning = 0;

    for (var tg = 0; tg < groups.length; tg++) {
      var groep      = groups[tg];
      var cols       = groupColMap[groep.label];
      if (!cols) continue;

      var bedragIdx  = cols.bedragCol - 1;
      var betaaldIdx = cols.betaaldCol - 1;
      var status     = allData[row][betaaldIdx] || "";
      var huidigBedrag = parseFloat(allData[row][bedragIdx]) || 0;

      if (status === "Betaald") {
        capRunning += huidigBedrag;
        continue;
      }

      // Bereken eligible birdies (joindatum gefloord naar middernacht)
      var eligibleBirdies = 0;
      for (var ti = 0; ti < groep.tournaments.length; ti++) {
        var toernooi = groep.tournaments[ti];
        if (!sponsor.joinDate || !toernooi.date || toernooi.date >= sponsor.joinDate) {
          eligibleBirdies += toernooi.birdies;
        }
      }

      var rawBedrag = sponsor.perBirdie * eligibleBirdies;
      var newBedrag = rawBedrag;
      if (sponsor.maxSeizoen !== null) {
        newBedrag = Math.min(rawBedrag, Math.max(0, sponsor.maxSeizoen - capRunning));
      }

      var newStatus;
      if (!sponsor.actief) {
        newStatus  = "Inactief";
        newBedrag  = 0;
      } else {
        newStatus  = (status === "Open" || status === "") ? "Open" : status;
      }

      betalingen.getRange(row + 2, cols.bedragCol).setValue(newBedrag).setNumberFormat('€#,##0.00');
      betalingen.getRange(row + 2, cols.betaaldCol).setValue(newStatus);

      capRunning += newBedrag;
    }
  }

  // ── 8. Dropdown + voorwaardelijke opmaak ─────────────────────────────────
  var totalSponsorRows = betalingen.getLastRow() - 1;
  if (totalSponsorRows > 0) {
    var allRules = [];
    for (var dg = 0; dg < groups.length; dg++) {
      var dgCols = groupColMap[groups[dg].label];
      if (!dgCols) continue;
      var dgRange = betalingen.getRange(2, dgCols.betaaldCol, totalSponsorRows, 1);
      dgRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(["Open", "Betaald", "Inactief"], true).build()
      );
      allRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Betaald").setBackground("#d1fae5").setFontColor("#065f46")
        .setRanges([dgRange]).build());
      allRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Open").setBackground("#fee2e2").setFontColor("#991b1b")
        .setRanges([dgRange]).build());
      allRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Inactief").setBackground("#f3f4f6").setFontColor("#9ca3af")
        .setRanges([dgRange]).build());
    }
    betalingen.setConditionalFormatRules(allRules);
  }

  // ── 9. Totaalrij ─────────────────────────────────────────────────────────
  var finalHeaders = betalingen.getRange(1, 1, 1, betalingen.getLastColumn()).getValues()[0];
  var finalData    = betalingen.getLastRow() > 1
    ? betalingen.getRange(2, 1, betalingen.getLastRow() - 1, betalingen.getLastColumn()).getValues()
    : [];

  var totaalRow = ["TOTAAL OPEN"];
  for (var col = 1; col < finalHeaders.length; col += 2) {
    var som = 0;
    for (var frow = 0; frow < finalData.length; frow++) {
      if (finalData[frow][col + 1] === "Open") {
        som += parseFloat(finalData[frow][col]) || 0;
      }
    }
    totaalRow.push(som);
    totaalRow.push("");
  }
  betalingen.appendRow(totaalRow);
  var totaalRowIdx = betalingen.getLastRow();
  betalingen.getRange(totaalRowIdx, 1, 1, betalingen.getLastColumn())
    .setBackground("#fdf2f8")
    .setFontWeight("bold");
  for (var bc = 2; bc <= betalingen.getLastColumn(); bc += 2) {
    betalingen.getRange(totaalRowIdx, bc).setNumberFormat('€#,##0.00');
  }

  return { added: added, removed: removed, newSponsors: newSponsors };
}

// ─────────────────────────────────────────────────────────────────────────────
// Betaalverzoek popup
// ─────────────────────────────────────────────────────────────────────────────
function genereerBetaalverzoek() {
  var ss         = SpreadsheetApp.openById(SHEET_ID);
  var betalingen = ss.getSheetByName("Betalingen");
  if (!betalingen) {
    SpreadsheetApp.getUi().alert("Betalingen-tabblad niet gevonden.");
    return;
  }

  var data    = betalingen.getDataRange().getValues();
  var headers = data[0];
  var lines   = [];
  var grandTotal = 0;

  for (var i = 1; i < data.length; i++) {
    var naam = data[i][0];
    if (!naam || naam === "TOTAAL OPEN") continue;

    var totaal      = 0;
    var openGroepen = [];
    for (var j = 1; j + 1 < headers.length; j += 2) {
      if (data[i][j + 1] === "Open" && data[i][j]) {
        totaal += parseFloat(data[i][j]) || 0;
        openGroepen.push(headers[j].replace(" BEDRAG", "").trim());
      }
    }
    if (totaal > 0) {
      grandTotal += totaal;
      lines.push(naam + "  →  €" + totaal.toFixed(2) + "  (" + openGroepen.join(", ") + ")");
    }
  }

  if (lines.length === 0) {
    SpreadsheetApp.getUi().alert("✓ Alle sponsors hebben betaald!");
    return;
  }

  lines.push("");
  lines.push("TOTAAL  →  €" + grandTotal.toFixed(2));

  SpreadsheetApp.getUi().alert(
    "BETAALVERZOEK OVERZICHT\n" +
    "────────────────────────────────────\n\n" +
    lines.join("\n")
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Overzicht – refresh
// ─────────────────────────────────────────────────────────────────────────────
function refreshOverzichtFormulas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var overzichtSheet = ss.getSheetByName("Overzicht");
  if (!overzichtSheet) {
    SpreadsheetApp.getUi().alert("Overzicht-tabblad niet gevonden. Voer Setup Overzicht uit.");
    return;
  }

  overzichtSheet.clear();
  overzichtSheet.getRange(1, 1, overzichtSheet.getMaxRows(), overzichtSheet.getMaxColumns()).clearDataValidations();
  overzichtSheet.setConditionalFormatRules([]);

  var headers = ["NAAM", "EMAIL", "PER BIRDIE (€)", "MAX SEIZOEN (€)", "TOTAAL BIRDIES", "BEREKEND BEDRAG (€)", "CAP BEREIKT?"];
  overzichtSheet.appendRow(headers);
  overzichtSheet.setFrozenRows(1);

  [140, 200, 130, 130, 120, 160, 110].forEach(function(w, i) {
    overzichtSheet.setColumnWidth(i + 1, w);
  });
  overzichtSheet.getRange("A1:G1")
    .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");

  overzichtSheet.getRange(2, 1, 199, 7).setFormulas(buildOverzichtFormulas_());
  overzichtSheet.getRange("C2:D200").setNumberFormat('€#,##0.00');
  overzichtSheet.getRange("F2:F200").setNumberFormat('€#,##0.00');

  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Ja")
    .setBackground("#d1fae5").setFontColor("#065f46")
    .setRanges([overzichtSheet.getRange("G2:G200")])
    .build();
  overzichtSheet.setConditionalFormatRules([rule]);

  overzichtSheet.getRange("I1").setValue("TOTAAL PER BIRDIE (€)")
    .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");
  overzichtSheet.getRange("I2").setFormula("=SUM(C2:C200)")
    .setNumberFormat('€#,##0.00').setFontWeight("bold");
  overzichtSheet.setColumnWidth(9, 180);

  SpreadsheetApp.getUi().alert("Overzicht bijgewerkt.");
}

// Aanmeldingen kolommen: A=timestamp B=naam C=email D=telefoon E=per_birdie F=max_seizoen G=whatsapp H=actief
// TOTAAL BIRDIES en BEREKEND BEDRAG houden rekening met de joindatum via SUMPRODUCT.
function buildOverzichtFormulas_() {
  var sp = 'SUMPRODUCT((Birdies!A$2:A$200<>"")*(Birdies!A$2:A$200>=INT(Aanmeldingen!A';
  var formulas = [];
  for (var r = 2; r <= 200; r++) {
    var eligible = sp + r + '))*Birdies!C$2:C$200)';
    formulas.push([
      '=IF(Aanmeldingen!B' + r + '="";"";Aanmeldingen!B' + r + ')',
      '=IF(Aanmeldingen!C' + r + '="";"";Aanmeldingen!C' + r + ')',
      '=IF(Aanmeldingen!E' + r + '="";"";Aanmeldingen!E' + r + ')',
      '=IF(Aanmeldingen!B' + r + '="";"";IF(Aanmeldingen!F' + r + '="";"–";Aanmeldingen!F' + r + '))',
      '=IF(Aanmeldingen!B' + r + '="";"";' + eligible + ')',
      '=IF(Aanmeldingen!B' + r + '="";"";IF(Aanmeldingen!F' + r + '="";' +
        'Aanmeldingen!E' + r + '*' + eligible + ';' +
        'MIN(Aanmeldingen!E' + r + '*' + eligible + ';Aanmeldingen!F' + r + ')))',
      '=IF(Aanmeldingen!B' + r + '="";"";IF(Aanmeldingen!F' + r + '="";"–";' +
        'IF(Aanmeldingen!E' + r + '*' + eligible + '>=Aanmeldingen!F' + r + ';"✓ Ja";"Nee")))',
    ]);
  }
  return formulas;
}

// ─────────────────────────────────────────────────────────────────────────────
// Setup Overzicht + Birdies (eenmalig)
// ─────────────────────────────────────────────────────────────────────────────
function setupOverzicht() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var heeftData = ss.getSheetByName("Birdies") || ss.getSheetByName("Overzicht");
  if (heeftData) {
    var ui = SpreadsheetApp.getUi();
    var antwoord = ui.alert(
      "Tabbladen bestaan al",
      "Wil je Birdies en Overzicht volledig opnieuw aanmaken? Alle toernooigegevens gaan verloren.",
      ui.ButtonSet.YES_NO
    );
    if (antwoord !== ui.Button.YES) return;
  }

  var birdiesSheet = ss.getSheetByName("Birdies");
  if (!birdiesSheet) {
    birdiesSheet = ss.insertSheet("Birdies");
  } else {
    birdiesSheet.clear();
  }

  birdiesSheet.appendRow(["DATUM", "RONDE / OMSCHRIJVING", "AANTAL BIRDIES"]);
  birdiesSheet.setFrozenRows(1);
  birdiesSheet.setColumnWidth(1, 120);
  birdiesSheet.setColumnWidth(2, 220);
  birdiesSheet.setColumnWidth(3, 160);
  birdiesSheet.getRange("A1:C1")
    .setBackground("#1e1b2e").setFontColor("#ffffff").setFontWeight("bold");
  birdiesSheet.getRange("A2:A200").setNumberFormat("dd-mm-yyyy");
  birdiesSheet.getRange("A2:A200").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireDateAfter(new Date(2000, 0, 1))
      .setAllowInvalid(false)
      .setHelpText("Kies een datum via de kalender.")
      .build()
  );

  var overzichtSheet = ss.getSheetByName("Overzicht");
  if (!overzichtSheet) {
    overzichtSheet = ss.insertSheet("Overzicht");
  } else {
    overzichtSheet.clear();
  }

  var headers = ["NAAM", "EMAIL", "PER BIRDIE (€)", "MAX SEIZOEN (€)", "TOTAAL BIRDIES", "BEREKEND BEDRAG (€)", "CAP BEREIKT?"];
  overzichtSheet.appendRow(headers);
  overzichtSheet.setFrozenRows(1);

  [140, 200, 130, 130, 120, 160, 110].forEach(function(w, i) {
    overzichtSheet.setColumnWidth(i + 1, w);
  });
  overzichtSheet.getRange("A1:G1")
    .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");

  overzichtSheet.getRange(2, 1, 199, 7).setFormulas(buildOverzichtFormulas_());
  overzichtSheet.getRange("C2:D200").setNumberFormat('€#,##0.00');
  overzichtSheet.getRange("F2:F200").setNumberFormat('€#,##0.00');

  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Ja")
    .setBackground("#d1fae5").setFontColor("#065f46")
    .setRanges([overzichtSheet.getRange("G2:G200")])
    .build();
  overzichtSheet.setConditionalFormatRules([rule]);

  overzichtSheet.getRange("I1").setValue("TOTAAL PER BIRDIE (€)")
    .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");
  overzichtSheet.getRange("I2").setFormula("=SUM(C2:C200)")
    .setNumberFormat('€#,##0.00').setFontWeight("bold");
  overzichtSheet.setColumnWidth(9, 180);

  SpreadsheetApp.getUi().alert(
    "Klaar! Tabbladen 'Birdies' en 'Overzicht' zijn aangemaakt.\n\n" +
    "Voer daarna 'Setup Betalingen' uit via het Birdie Vrienden-menu."
  );
}
