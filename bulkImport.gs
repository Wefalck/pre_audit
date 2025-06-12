function parseCSVSemrush(file) {
  const fileName = file.fileName;

  Logger.log(`    [parseCSVSemrush] Traitement du fichier : ${fileName}`);

  // Étape 1 - Extraction de la date depuis le nom du fichier (format YYYYMMDD)
  const dateMatch = fileName.match(/\d{8}/);
  if (!dateMatch) throw new Error(`⚠️ ${fileName} : date introuvable`);

  // Étape 2 - Conversion de la date au format MM-YY
  const sheetName = formatDateToMMYY(dateMatch[0]);
  Logger.log(`    [parseCSVSemrush] Date extraite = ${sheetName}`);

  // Étape 3 - Parsing brut du contenu CSV
  const parsedRows = parseCSV(file.csvString);
  if (parsedRows.length <= 1) {
    throw new Error(`⚠️ ${fileName} : données vides`);
  }
  Logger.log(`    [parseCSVSemrush] ${parsedRows.length - 1} lignes de données brutes extraites`);

  // Étape 4 - Transformation des données utiles
  const rows = parsedRows.slice(1);
  const mappedData = rows.map(row => [
    row[0], row[3], row[1], "", "", "", "", row[6], row[7], ""
  ]);
  Logger.log(`    [parseCSVSemrush] ${mappedData.length} lignes mappées pour injection`);

  // Étape 5 - PURGE : suppression des onglets date > 15 plus récents (1 seule fois par run, donc ici c’est redondant si plusieurs fichiers)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  // Filtrer uniquement les feuilles avec nom de type MM-YY
  const dateSheets = allSheets.filter(sheet => /^\d{2}-\d{2}$/.test(sheet.getName()));

  // Trier par date décroissante (de la plus récente à la plus ancienne)
  dateSheets.sort((a, b) => {
    const [am, ay] = a.getName().split('-').map(Number);
    const [bm, by] = b.getName().split('-').map(Number);
    const aDate = new Date(2000 + ay, am - 1);
    const bDate = new Date(2000 + by, bm - 1);
    return bDate - aDate;
  });

  // Supprimer les feuilles au-delà des 15 plus récentes
  const excessSheets = dateSheets.slice(15);
  excessSheets.forEach(sheet => {
    Logger.log("    [parseCSVSemrush] 🗑️ Suppression de l'onglet ancien : " + sheet.getName());
    ss.deleteSheet(sheet);
  });

  // Étape 6 - Retour des données pour traitement en aval
  return { sheetName, mappedData };
}

function createTargetSheet(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Supprimer la feuille si elle existe déjà
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  // Créer la nouvelle feuille
  const sheet = ss.insertSheet(sheetName);
  SpreadsheetApp.setActiveSheet(sheet); // rester dessus

  // Injecter les données ligne 5
  sheet.getRange(5, 1, data.length, 10).setValues(data);

  Logger.log(`📥 Données injectées dans la feuille : ${sheetName}`);
}

function applyDynamicFormulas(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  // 1. Déterminer les noms de feuilles M-1 et N-1
  const prevMonth = getPreviousMonthOrYearSheetName(sheetName, 'month');
  const prevYear = getPreviousMonthOrYearSheetName(sheetName, 'year');

  Logger.log("🧩 Début de l'injection des formules ordonnées dans " + sheetName);

  // 2. Formule colonne E (Position M-1)
  const colE = sheet.getRange(5, 5, lastRow - 4);
  colE.setFormula(`=IFERROR(VLOOKUP(A5; '${prevMonth}'!$A$4:I; 3; FALSE); "NO POS")`);
  SpreadsheetApp.flush();
  colE.copyTo(colE, { contentsOnly: true });
  Logger.log("✅ Colonne E (M-1) injectée et figée.");

  // 3. Formule colonne G (Position N-1)
  const colG = sheet.getRange(5, 7, lastRow - 4);
  colG.setFormula(`=IFERROR(VLOOKUP(A5; '${prevYear}'!$A$4:I; 3; FALSE); "NO POS")`);
  SpreadsheetApp.flush();
  colG.copyTo(colG, { contentsOnly: true });
  Logger.log("✅ Colonne G (N-1) injectée et figée.");

  // 4. Formule colonne J (Trafic M-1)
  const colJ = sheet.getRange(5, 10, lastRow - 4);
  colJ.setFormula(`=IFERROR(VLOOKUP(A5; '${prevMonth}'!$A$5:J; 9; FALSE); "0")`);
  SpreadsheetApp.flush();
  colJ.copyTo(colJ, { contentsOnly: true });
  Logger.log("✅ Colonne J (Trafic M-1) injectée et figée.");

  // 5. Formule colonne D (Var M-1)
  const colD = sheet.getRange(5, 4, lastRow - 4);
  colD.setFormula(`=IFERROR(E5-C5; "Nouveau")`);
  SpreadsheetApp.flush();
  colD.copyTo(colD, { contentsOnly: true });
  Logger.log("✅ Colonne D (Variation M-1) injectée et figée.");

  // 6. Formule colonne F (Var N-1)
  const colF = sheet.getRange(5, 6, lastRow - 4);
  colF.setFormula(`=IFERROR(G5-C5; "Nouveau")`);
  SpreadsheetApp.flush();
  colF.copyTo(colF, { contentsOnly: true });
  Logger.log("✅ Colonne F (Variation N-1) injectée et figée.");

  // 7. Ligne 3 : I3, J3, E3, D3, G3, F3
  sheet.getRange("I3").setFormula(`=SUM(I5:I)`);
  sheet.getRange("J3").setFormula(`=IFERROR('${prevMonth}'!I3; "0")`);
  sheet.getRange("E3").setFormula(`=J3`);
  sheet.getRange("D3").setFormula(`=IFERROR(I3 - E3; "")`);
  sheet.getRange("G3").setFormula(`=IFERROR('${prevYear}'!I3; "")`);
  sheet.getRange("F3").setFormula(`=IFERROR(I3 - G3; "")`);
  SpreadsheetApp.flush();
  sheet.getRange("D3:G3").copyTo(sheet.getRange("D3:G3"), { contentsOnly: true });
  sheet.getRange("I3:J3").copyTo(sheet.getRange("I3:J3"), { contentsOnly: true });
  Logger.log("📌 Ligne 3 figée en valeurs.");

  Logger.log("🎯 Ligne 3 mise à jour avec les formules de synthèse.");
  Logger.log("🏁 Formules ordonnées appliquées avec succès à " + sheetName);
}

function appliquerFormatageFinal(sheetName, nbLignes) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error("Feuille introuvable : " + sheetName);

  Logger.log("🎨 [Formatage] Début du formatage pour : " + sheetName);
  Logger.log("[DEBUG] Paramètres initiaux : sheetName=" + sheetName + " | nbLignes=" + nbLignes);

  // 1. Suppression des colonnes à droite de J
  const maxCols = sheet.getMaxColumns();
  Logger.log("[DEBUG] maxCols = " + maxCols);
  if (maxCols > 10) {
    Logger.log("[DEBUG] Suppression des colonnes à droite de J (col 11 à " + maxCols + ")");
    sheet.deleteColumns(11, maxCols - 10);
  }

  // 2. Ligne 1 : titre principal
  const configSheet = spreadsheet.getSheetByName("Configuration");
  const siteName = configSheet.getRange("C2").getValue();

  const [month, yearShort] = sheetName.split('-');
  const year = "20" + yearShort;
  const monthNames = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
  const fullDate = monthNames[parseInt(month, 10) - 1] + " " + year;

  const row1 = sheet.getRange("A1:J1");
  row1.merge().setValue(
    `Positionnement SEO ${siteName} FR ${fullDate}\nRelevé fait à un instant T (Le 15 du mois), le positionnement peut avoir évolué depuis.`
  ).setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF").setFontFamily("Arial")
   .setHorizontalAlignment("center").setBackground("#073763");
  sheet.setRowHeight(1, 80);

  // 3. Ligne 2 et 3
  sheet.getRange("A2:J2").setFontSize(10).setFontColor("#000000").setFontFamily("Arial");
  sheet.getRange("A3:J3").setFontSize(10).setFontColor("#000000").setFontFamily("Arial");
  sheet.getRange("I3:J3").setHorizontalAlignment("right").setNumberFormat("0");
  sheet.getRange("D3:G3").setHorizontalAlignment("center");
  sheet.getRangeList(["D3", "F3"]).setNumberFormat("+0;-0");

  // 4. Ligne 4 : en-tête
  sheet.getRange("A4:J4")
    .setFontWeight("bold").setFontColor("#FFFFFF")
    .setFontFamily("Arial").setHorizontalAlignment("center")
    .setValues([["Mots clés", "Volume", "Position", "Variation M-1", "Position M-1", "Variation N-1", "Position N-1", "URL", "Trafic", "Trafic M-1"]]);

  // 5. Suppression des lignes inutiles après les données (version strictement sécurisée)
  const maxRows = sheet.getMaxRows();
  const lastRowToKeep = 4 + nbLignes;
  Logger.log(`[DEBUG] Suppression lignes : maxRows=${maxRows} | lastRowToKeep=${lastRowToKeep}`);

  if (maxRows > lastRowToKeep) {
    const nbRowsToDelete = maxRows - lastRowToKeep;
    Logger.log(`[DEBUG] Suppression de ${nbRowsToDelete} lignes à partir de la ligne ${lastRowToKeep + 1}`);
    sheet.deleteRows(lastRowToKeep + 1, nbRowsToDelete);
  } else {
    Logger.log("[DEBUG] Rien à supprimer, la feuille est déjà à la bonne taille ou plus petite");
  }

  // 6. Mise en forme des colonnes
  const columns = [
    { col: 1, align: "left",   width: 350 },
    { col: 2, align: "center", width: 120, format: "#,##0" },
    { col: 3, align: "center", width: 120, format: "0" },
    { col: 4, align: "center", width: 120, format: "+#,##0;-#,##0;#,##0" },
    { col: 5, align: "center", width: 120, format: "#,##0" },
    { col: 6, align: "center", width: 120, format: "+#,##0;-#,##0;#,##0" },
    { col: 7, align: "center", width: 120, format: "#,##0" },
    { col: 8, align: "left",   width: 550 },
    { col: 9, align: "right",  width: 100, format: "#,##0" },
    { col: 10, align: "right", width: 100, format: "#,##0" }
  ];

  columns.forEach(({ col, align, width, format }) => {
    const range = sheet.getRange(5, col, nbLignes);
    range.setHorizontalAlignment(align);
    if (format) range.setNumberFormat(format);
    sheet.setColumnWidth(col, width);
  });

  // 7. Figer les lignes d’en-tête
  sheet.setFrozenRows(4);

  // 8. Quadrillage off, alignement vertical middle
  sheet.setHiddenGridlines(true);
  sheet.getRange("A1:J" + (4 + nbLignes)).setVerticalAlignment("middle");

  // 9. Mise en forme conditionnelle
  const rules = [
    // "NO POS" → texte rouge gras
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("NO POS").setFontColor("#FF0000").setBold(true)
      .setRanges([
        sheet.getRange("E5:E" + (4 + nbLignes)),
        sheet.getRange("G5:G" + (4 + nbLignes))
      ]).build(),

    // "Nouveau" → fond vert clair
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Nouveau").setBackground("#B7E1CD")
      .setRanges([
        sheet.getRange("D5:D" + (4 + nbLignes)),
        sheet.getRange("F5:F" + (4 + nbLignes))
      ]).build(),

    // Variation positive ≥ 0 → fond vert clair
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0).setBackground("#B7E1CD")
      .setRanges([
        sheet.getRange("D5:D" + (4 + nbLignes)),
        sheet.getRange("F5:F" + (4 + nbLignes)),
        sheet.getRange("D3"),
        sheet.getRange("F3")
      ]).build(),

    // Variation négative < 0 → fond rouge clair
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setBackground("#F4CCCC")
      .setRanges([
        sheet.getRange("D5:D" + (4 + nbLignes)),
        sheet.getRange("F5:F" + (4 + nbLignes)),
        sheet.getRange("D3"),
        sheet.getRange("F3")
      ]).build(),

    // Dégradé sur la colonne B (Volume)
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint("#CFE2F3")
      .setGradientMaxpoint("#3C78D8")
      .setRanges([sheet.getRange("B5:B" + (4 + nbLignes))])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
  Logger.log("🎨 Mise en forme conditionnelle appliquée, incluant D3 et F3.");

  // 10. Banding (couleurs alternées)
  const bandingRange = sheet.getRange("A4:J" + (4 + nbLignes));
  bandingRange.getBandings().forEach(b => b.remove());
  bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor("#073763")
    .setFirstRowColor("#FFFFFF")
    .setSecondRowColor("#F3F3F3");

  // 11. Nommer la plage
  const monthNamesForRange = ["janvier", "fevrier", "mars", "avril", "mai", "juin", "juillet", "aout", "septembre", "octobre", "novembre", "decembre"];
  const rangeName = monthNamesForRange[parseInt(month, 10) - 1] + "_" + yearShort;
  spreadsheet.setNamedRange(rangeName, sheet.getRange("A1:J" + (4 + nbLignes)));

  // 12. Appliquer filtre sur A4:J
  sheet.getRange("A4:J").createFilter();

  //13. Trier les onglets
  trierOnglets();

  Logger.log("✅ [Formatage] Terminé pour : " + sheetName);
}

function getPreviousMonthOrYearSheetName(currentSheetName, mode) {
    var parts = currentSheetName.split('-');
    var month = parseInt(parts[0], 10);
    var year = parseInt(parts[1], 10);

    if (mode === "month") {
        // Si le mois est janvier, le mois précédent serait décembre de l'année précédente
        if (month === 6) {
            month = 12;
            year -= 1;
        } else {
            month -= 1;
        }
    } else if (mode === "year") {
        year -= 1;  // simplement décrémenter l'année
    }

    // Convertir le mois en une chaîne de deux caractères (par exemple, '01' pour janvier)
    var monthString = month < 10 ? '0' + month : '' + month;

    return monthString + '-' + year;
}

function trierOnglets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var pattern = /^(\d{2})-(\d{2})$/; // Pattern pour matcher le format "MM-YY"

  // Sépare les onglets en trois listes : ceux qui suivent le format "MM-YY", "Import données" et les autres
  var dateSheets = [];
  var otherSheets = [];
  var importSheet = null;
  
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (name === "Import données") {
      importSheet = sheet;
    } else if (pattern.test(name)) {
      dateSheets.push(sheet);
    } else {
      otherSheets.push(sheet);
    }
  });

  // Trie les onglets avec le format "MM-YY" dans l'ordre décroissant
  dateSheets.sort(function(a, b) {
    var aName = a.getName().match(pattern);
    var bName = b.getName().match(pattern);
    var aDate = new Date("20" + aName[2], parseInt(aName[1]) - 1); // Transforme "MM-YY" en date
    var bDate = new Date("20" + bName[2], parseInt(bName[1]) - 1); // Transforme "MM-YY" en date
    return bDate - aDate; // Trie dans l'ordre décroissant
  });

  // Place d'abord les onglets non-datés, puis les onglets datés, et enfin l'onglet "Import données"
  var orderedSheets = otherSheets.concat(dateSheets);
  if (importSheet) {
    orderedSheets.push(importSheet);
  }

  for (var i = 0; i < orderedSheets.length; i++) {
    spreadsheet.setActiveSheet(orderedSheets[i]);
    spreadsheet.moveActiveSheet(i + 1);
  }
}

function getBulkImportInstructionsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Configuration");
  if (!configSheet) throw new Error("Feuille 'Configuration' introuvable.");

  // [1] Lecture des valeurs attendues
  var months = [
    configSheet.getRange("C11").getValue(),
    configSheet.getRange("C10").getValue(),
    configSheet.getRange("C9").getValue()
  ];
  var paramC3 = configSheet.getRange("C3").getValue();

  // [2] Log pour debug
  Logger.log("[getBulkImportInstructionsData] months=" + JSON.stringify(months) + ", C3=" + paramC3);

  // [3] Retourne les valeurs au front
  return {
    months: months,
    c3: paramC3
  };
}