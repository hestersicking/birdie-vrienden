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

// Kolomvolgorde in de sheet (pas aan als je kolommen wilt herordenen)
var COLUMNS = ["timestamp", "naam", "email", "telefoon", "bedrijf", "per_birdie", "max_seizoen", "whatsapp"];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    // Maak het tabblad aan als het nog niet bestaat
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // Schrijf kolomkoppen
      sheet.appendRow(COLUMNS.map(function(c) { return c.toUpperCase(); }));
      sheet.setFrozenRows(1);
    }

    // Schrijf een nieuwe rij
    var row = [
      new Date(),                                          // timestamp
      data.naam        || "",
      data.email       || "",
      data.telefoon    || "",
      data.bedrijf     || "",
      data.per_birdie  != null ? data.per_birdie : "",
      data.max_seizoen != null ? data.max_seizoen : "",
      data.whatsapp    ? "Ja" : "Nee",
    ];

    sheet.appendRow(row);

    return buildResponse({ status: "ok", message: "Aanmelding opgeslagen." });

  } catch (err) {
    return buildResponse({ status: "error", message: err.toString() }, true);
  }
}

// Antwoord op OPTIONS preflight (browsers sturen dit soms vooraf)
function doGet(e) {
  return buildResponse({ status: "ok", message: "Birdie Vrienden webhook actief." });
}

function buildResponse(obj, isError) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ─────────────────────────────────────────────────────────────────────────────
// Eenmalig uitvoeren via de Apps Script editor: selecteer setupOverzicht en
// klik op "Uitvoeren". Maakt de tabbladen "Birdies" en "Overzicht" aan.
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Birdie Vrienden')
    .addItem('Setup Overzicht + Birdies', 'setupOverzicht')
    .addItem('Setup Betalingen', 'setupBetalingen')
    .addItem('Refresh Overzicht formulas', 'refreshOverzichtFormulas')
    .addSeparator()
    .addItem('Sync toernooien naar Betalingen', 'syncToernooien')
    .addSeparator()
    .addItem('Genereer betaalverzoek', 'genereerBetaalverzoek')
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────────────────
// Betalingen – matrix sponsors × toernooien met betaalstatus
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
    sheet.clearDataValidations();
    sheet.setConditionalFormatRules([]);
  } else {
    sheet = ss.insertSheet("Betalingen");
  }

  sheet.appendRow(["NAAM"]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 180);
  sheet.getRange("A1")
    .setBackground("#9D174D")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  var aanmSheet = ss.getSheetByName("Aanmeldingen");
  if (aanmSheet) {
    var aanmData = aanmSheet.getDataRange().getValues();
    for (var i = 1; i < aanmData.length; i++) {
      if (aanmData[i][1]) sheet.appendRow([aanmData[i][1]]);
    }
  }

  syncToernooien_(ss, sheet);
  SpreadsheetApp.getUi().alert("Betalingen-tabblad aangemaakt!");
}

// Voegt toernooien toe die nog niet in Betalingen staan.
// Veilig om meerdere keren te draaien (idempotent).
function syncToernooien() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Betalingen");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Maak eerst het Betalingen-tabblad aan via Setup Betalingen.");
    return;
  }
  var added = syncToernooien_(ss, sheet);
  SpreadsheetApp.getUi().alert(
    added === 0 ? "Alles is al up-to-date." : added + " nieuw(e) toernooi(en) toegevoegd."
  );
}

function syncToernooien_(ss, betalingen) {
  var birdiesSheet = ss.getSheetByName("Birdies");
  var aanmSheet    = ss.getSheetByName("Aanmeldingen");
  if (!birdiesSheet || !aanmSheet) return 0;

  // Alle toernooien uit Birdies (rij 2 t/m einde)
  var birdiesData = birdiesSheet.getDataRange().getValues();
  var allBirdies = [];
  for (var i = 1; i < birdiesData.length; i++) {
    allBirdies.push(Number(birdiesData[i][2]) || 0);
  }

  // Groepen van 3 toernooien: T1-T3, T4-T6, …
  // Een groep wordt pas toegevoegd als alle 3 toernooien in Birdies staan.
  var groups = [];
  for (var g = 0; g * 3 + 2 < allBirdies.length; g++) {
    var from  = g * 3 + 1;
    var to    = from + 2;
    var label = "T" + from + "-T" + to;
    var totalBirdies = allBirdies[from - 1] + allBirdies[from] + allBirdies[from + 1];
    groups.push({ label: label, birdies: totalBirdies });
  }

  // Welke groepen staan al in Betalingen?
  var lastCol = betalingen.getLastColumn();
  var headers = betalingen.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
  var existing = {};
  for (var h = 0; h < headers.length; h++) {
    var hdr = headers[h].toString();
    if (hdr.indexOf(" BETAALD") !== -1) {
      existing[hdr.replace(" BETAALD", "").trim()] = true;
    }
  }

  // per_birdie per sponsor
  var aanmData = aanmSheet.getDataRange().getValues();
  var perBirdieMap = {};
  for (var s = 1; s < aanmData.length; s++) {
    if (aanmData[s][1]) perBirdieMap[aanmData[s][1]] = parseFloat(aanmData[s][5]) || 0;
  }

  // Sponsornamen uit Betalingen
  var lastRow = betalingen.getLastRow();
  var sponsorNames = lastRow > 1
    ? betalingen.getRange(2, 1, lastRow - 1, 1).getValues()
    : [];

  var added = 0;
  for (var t = 0; t < groups.length; t++) {
    var groep = groups[t];
    if (existing[groep.label]) continue;

    var bedragCol  = betalingen.getLastColumn() + 1;
    var betaaldCol = bedragCol + 1;

    betalingen.getRange(1, bedragCol).setValue(groep.label + " BEDRAG");
    betalingen.getRange(1, betaaldCol).setValue(groep.label + " BETAALD");
    betalingen.getRange(1, bedragCol, 1, 2)
      .setBackground("#9D174D").setFontColor("#ffffff").setFontWeight("bold");
    betalingen.setColumnWidth(bedragCol, 130);
    betalingen.setColumnWidth(betaaldCol, 120);

    for (var r = 0; r < sponsorNames.length; r++) {
      var naam = sponsorNames[r][0];
      if (!naam) continue;
      var bedrag = (perBirdieMap[naam] || 0) * groep.birdies;
      betalingen.getRange(r + 2, bedragCol).setValue(bedrag).setNumberFormat('€#,##0.00');
      betalingen.getRange(r + 2, betaaldCol).setValue("Open");
    }

    if (sponsorNames.length > 0) {
      var dataRange = betalingen.getRange(2, betaaldCol, sponsorNames.length, 1);
      dataRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(["Open", "Betaald"], true).build()
      );
      var rules = betalingen.getConditionalFormatRules();
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Betaald").setBackground("#d1fae5").setFontColor("#065f46")
        .setRanges([dataRange]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Open").setBackground("#fee2e2").setFontColor("#991b1b")
        .setRanges([dataRange]).build());
      betalingen.setConditionalFormatRules(rules);
    }

    added++;
  }
  return added;
}

// Toont per sponsor het openstaande bedrag per groep van 3 toernooien.
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

  for (var i = 1; i < data.length; i++) {
    var naam = data[i][0];
    if (!naam) continue;

    var totaal = 0;
    var openGroepen = [];
    for (var j = 1; j + 1 < headers.length; j += 2) {
      if (data[i][j + 1] === "Open" && data[i][j]) {
        totaal += parseFloat(data[i][j]) || 0;
        openGroepen.push(headers[j].replace(" BEDRAG", "").trim());
      }
    }
    if (totaal > 0) {
      lines.push(naam + "  →  €" + totaal.toFixed(2) + "  (" + openGroepen.join(", ") + ")");
    }
  }

  if (lines.length === 0) {
    SpreadsheetApp.getUi().alert("✓ Alle sponsors hebben betaald!");
    return;
  }

  SpreadsheetApp.getUi().alert(
    "BETAALVERZOEK OVERZICHT\n" +
    "────────────────────────────────────\n\n" +
    lines.join("\n")
  );
}

// Herplaatst alleen de formules in Overzicht – raakt Birdies niet aan.
function refreshOverzichtFormulas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var overzichtSheet = ss.getSheetByName("Overzicht");
  if (!overzichtSheet) {
    SpreadsheetApp.getUi().alert("Overzicht-tabblad niet gevonden. Voer Setup Overzicht uit.");
    return;
  }
  overzichtSheet.getRange(2, 1, 199, 8).setFormulas(buildOverzichtFormulas_());
  SpreadsheetApp.getUi().alert("Overzicht formulas bijgewerkt.");
}

function buildOverzichtFormulas_() {
  var formulas = [];
  for (var r = 2; r <= 200; r++) {
    formulas.push([
      '=IF(Aanmeldingen!B' + r + '="";"";Aanmeldingen!B' + r + ')',
      '=IF(Aanmeldingen!C' + r + '="";"";Aanmeldingen!C' + r + ')',
      '=IF(Aanmeldingen!E' + r + '="";"";Aanmeldingen!E' + r + ')',
      '=IF(Aanmeldingen!F' + r + '="";"";Aanmeldingen!F' + r + ')',
      '=IF(Aanmeldingen!B' + r + '="";"";IF(Aanmeldingen!G' + r + '="";"–";Aanmeldingen!G' + r + '))',
      '=IF(Aanmeldingen!B' + r + '="";"";SUM(Birdies!C:C))',
      '=IF(Aanmeldingen!B' + r + '="";"";' +
        'IF(Aanmeldingen!G' + r + '="";Aanmeldingen!F' + r + '*SUM(Birdies!C:C);' +
        'MIN(Aanmeldingen!F' + r + '*SUM(Birdies!C:C);Aanmeldingen!G' + r + ')))',
      '=IF(Aanmeldingen!B' + r + '="";"";' +
        'IF(Aanmeldingen!G' + r + '="";"–";' +
        'IF(Aanmeldingen!F' + r + '*SUM(Birdies!C:C)>=Aanmeldingen!G' + r + ';"✓ Ja";"Nee")))',
    ]);
  }
  return formulas;
}

// ─────────────────────────────────────────────────────────────────────────────
// Eenmalig uitvoeren via de Apps Script editor: selecteer setupOverzicht en
// klik op "Uitvoeren". Maakt de tabbladen "Birdies" en "Overzicht" aan.
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

  // ── Tabblad "Birdies" ──────────────────────────────────────────────────────
  var birdiesSheet = ss.getSheetByName("Birdies");
  if (!birdiesSheet) {
    birdiesSheet = ss.insertSheet("Birdies");
  } else {
    birdiesSheet.clear();
  }

  birdiesSheet.appendRow(["DATUM", "RONDE / OMSCHRIJVING", "AANTAL BIRDIES"]);
  birdiesSheet.setFrozenRows(1);

  // Kolombreedte
  birdiesSheet.setColumnWidth(1, 120);
  birdiesSheet.setColumnWidth(2, 220);
  birdiesSheet.setColumnWidth(3, 160);

  // Opmaak kolomkop
  birdiesSheet.getRange("A1:C1")
    .setBackground("#1e1b2e")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  // Datumnotatie kolom A (rij 2 t/m 200)
  birdiesSheet.getRange("A2:A200")
    .setNumberFormat("dd-mm-yyyy");

  // ── Tabblad "Overzicht" ───────────────────────────────────────────────────
  var overzichtSheet = ss.getSheetByName("Overzicht");
  if (!overzichtSheet) {
    overzichtSheet = ss.insertSheet("Overzicht");
  } else {
    overzichtSheet.clear();
  }

  // Kolomkoppen
  var headers = [
    "NAAM", "EMAIL", "BEDRIJF", "PER BIRDIE (€)",
    "MAX SEIZOEN (€)", "TOTAAL BIRDIES", "BEREKEND BEDRAG (€)", "CAP BEREIKT?"
  ];
  overzichtSheet.appendRow(headers);
  overzichtSheet.setFrozenRows(1);

  // Kolombreedte
  [140, 200, 160, 130, 130, 120, 160, 110].forEach(function(w, i) {
    overzichtSheet.setColumnWidth(i + 1, w);
  });

  // Opmaak kolomkop
  overzichtSheet.getRange("A1:H1")
    .setBackground("#9D174D")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  overzichtSheet.getRange(2, 1, 199, 8).setFormulas(buildOverzichtFormulas_());

  // Getalnotatie voor bedragen
  overzichtSheet.getRange("D2:E200").setNumberFormat('€#,##0.00');
  overzichtSheet.getRange("G2:G200").setNumberFormat('€#,##0.00');

  // Voorwaardelijke opmaak: cap bereikt → groene achtergrond
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Ja")
    .setBackground("#d1fae5")
    .setFontColor("#065f46")
    .setRanges([overzichtSheet.getRange("H2:H200")])
    .build();
  overzichtSheet.setConditionalFormatRules([rule]);

  SpreadsheetApp.getUi().alert("Klaar! Tabbladen 'Birdies' en 'Overzicht' zijn aangemaakt.\n\nVoer daarna 'Setup Betalingen' uit via het Birdie Vrienden-menu.");
}
