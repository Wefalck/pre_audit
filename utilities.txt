function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Admin SF')
        .addItem('📊 Ajouter les données', 'showDialogBulk')
      .addToUi();
  }

function showDialogBulk() { 
  var htmlOutput = HtmlService.createHtmlOutputFromFile('uploadFormBulk') 
      .setWidth(600)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Importer les données en masse'); 
}

function parseCSV(content, delimiter = ",") {
  const rows = [];
  let current = '';
  let inQuotes = false;
  let row = [];

  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const nextChar = content[i + 1];

    if (char === '"' && inQuotes && nextChar === '"') {
      current += '"';
      i++; // guillemet échappé
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === delimiter && !inQuotes) {
      row.push(current);
      current = '';
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      if (current || row.length > 0) {
        row.push(current);
        rows.push(row.map(c => c.trim()));
        row = [];
        current = '';
      }
    } else {
      current += char;
    }
  }

  if (current || row.length > 0) {
    row.push(current);
    rows.push(row.map(c => c.trim()));
  }

  return rows;
}

function listerFeuilles() {
  var classeur = SpreadsheetApp.getActiveSpreadsheet();
  var feuilles = classeur.getSheets();
  var feuillesFormatMMYY = [];
  
  // Boucle pour parcourir chaque feuille et vérifier le format de son nom
  for (var i = 0; i < feuilles.length; i++) {
    var nomFeuille = feuilles[i].getName();
    if (/^\d{2}-\d{2}$/.test(nomFeuille)) { // Si le nom correspond au format MM-YY
      feuillesFormatMMYY.push(nomFeuille);
    }
  }
  
  // Convertir MM-YY en Mois Année
  var mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
  var feuillesFormatMoisAnnee = feuillesFormatMMYY.map(function(mmYY) {
    var elements = mmYY.split("-");
    var nomMois = mois[parseInt(elements[0]) - 1];
    var annee = "20" + elements[1];
    return nomMois + " " + annee;
  });
  
  return feuillesFormatMoisAnnee;
}

function rangeNameFromDate(date) {
  var monthNames = ["janvier", "fevrier", "mars", "avril", "mai", "juin", 
                    "juillet", "aout", "septembre", "octobre", "novembre", "decembre"];
  var month = date.getMonth();
  var year = date.getFullYear();
  return monthNames[month] + "_" + year.toString().slice(-2);
}

function isDateSheet(sheetName) {
  return /^(\d{2})-(\d{2})$/.test(sheetName);
}

function formatDateToMMYY(importDate) {
    var year = importDate.substring(2, 4); // YY
    var month = importDate.substring(4, 6); // MM
    return month + '-' + year;
}

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}

function extractDomain(url) {
  var match = url.match(/^https?:\/\/([^\/]+)/i); // Cette regex extrait le domaine avec le protocole
  if (match) {
    var hostname = match[1]; // Récupère le domaine complet (avec sous-domaines)
    var domain = hostname;

    if (hostname != null) {
      var parts = hostname.split('.').reverse(); // Découpe le domaine en parties
      if (parts != null && parts.length > 1) {
        domain = parts[1] + '.' + parts[0];  // Concatène les deux dernières parties pour le domaine
        // Gère les cas spéciaux comme .co.uk
        if (hostname.toLowerCase().indexOf('.co.uk') != -1 && parts.length > 2) {
          domain = parts[2] + '.' + domain;
        }
      }
    }
    return domain;
  }
  return null; // Retourne null si l'URL n'est pas valide
}

function extractFullDomain(url) {
  var match = url.match(/^https?:\/\/([^\/]+)/i); // Cette regex extrait le domaine avec le protocole
  if (match) {
    return match[1]; // Retourne le domaine complet avec les sous-domaines
  }
  return null; // Retourne null si l'URL n'est pas valide
}
