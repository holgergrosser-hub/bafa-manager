// ===== KONFIGURATION =====
const CONFIG = {
  MASTER_FOLDER_NAME: 'Template BAFA',
  TEMPLATE_FOLDER_ID: '1S_KMGYI500ACGZR3y37O3LPpVw--V0J5',
  PARENT_FOLDER_ID: '1CxrypRxuDROK_QWJje397LmrzvvIheGJ',
  LOGO_FOLDER_ID: '1X0yHW8IwoacCj9YT1wYVLzLF0ZcmnI0B',
  PDF_API_KEY: 'holger.grosser@iso9001.info_txrVNeBl6yauboNS7fLvlcoS23nOFptOLTl0CyROqXC0BwTZ1DaNtG5aeTDnI6Le',
  PDF_CO_URL: 'https://api.pdf.co/v1',
  // CRM Super Master
  CRM_SUPER_MASTER_ID: '1FWbeX3YeK9Uidyn9obKJ7z-J-zXX1h5PsXcfk_YHAyU',
  CRM_SHEET_NAME: 'Super Master'
};

// Mapping: Super Master Spalten → BAFA Kundentabelle Spalten
// Anpassen falls sich die CRM-Spalten ändern!
const CRM_IMPORT_MAP = {
  'A':  'companyName',       // Firmenname
  'BS': 'Strasse',           // Straße
  'BW': 'PLZ_Ort',           // PLZ + Ort
  'CA': 'Ansprechpartner',   // Ansprechpartner
  'I':  'email',             // E-Mail
  'C':  'Webpage',           // Webpage (laut Super Master: Spalte C)
  'L':  'AUDITOR'            // Auditor (laut Super Master: Spalte L)
  // Weitere CRM-Spalten hier ergänzen:
  // 'XX': 'Webpage',
  // 'YY': 'Geltungsbereich',
};

// ===== TABELLEN-DEFINITION =====
// Spaltenname = Platzhalter im Dokument: {{Spaltenname}}

const TABLE_COLUMNS = {
  // CRM-Stammdaten (Import)
  crm: [
    'companyName',
    'Strasse',
    'PLZ_Ort',
    'Ansprechpartner'
  ],

  // Weitere Platzhalter für Dokumente (erweiterbar)
  placeholders: [
    'email',
    'Webpage',
    'Gruenderdatum',
    'AnzahlderMitarbeiter',
    'Geltungsbereich',
    'Zielgruppe_Zielgebiet',
    'Ausgelagerte_Prozesse',
    'Norm',
    'AUDITOR',
    'Aktuelles_Jahr',
    'Datum_Heute'
  ],

  // System-Spalten (automatisch befüllt)
  system: [
    'folderID',
    'folderLink',
    'createdDate',
    'logoUrl'
  ],

  // Dokument-ID Spalten 01–14 (automatisch befüllt beim Kopieren)
  documents: [
    'doc_01_Beraterbewertung',
    'doc_02_Kundenrueckmeldung',
    'doc_03_Normen_Gesetze',
    'doc_04_Managementbewertung',
    'doc_05_Massnahmenplan',
    'doc_06_Prozessbeschreibungen',
    'doc_07_Schulungsplan',
    'doc_08_Ziele_Prozesskennzahlen',
    'doc_09_Unternehmenshandbuch',
    'doc_10_Auditbericht',
    'doc_11_Vollmacht',
    'doc_12_Firmeninfo_Foerdergeld',
    'doc_13_Projektbericht',
    'doc_14_Ausfuellanleitung'
  ]
};

// Mapping: Dokumentnamen-Prefix → Spaltenname
const DOC_COLUMN_MAP = {
  '01': 'doc_01_Beraterbewertung',
  '02': 'doc_02_Kundenrueckmeldung',
  '03': 'doc_03_Normen_Gesetze',
  '04': 'doc_04_Managementbewertung',
  '05': 'doc_05_Massnahmenplan',
  '06': 'doc_06_Prozessbeschreibungen',
  '07': 'doc_07_Schulungsplan',
  '08': 'doc_08_Ziele_Prozesskennzahlen',
  '09': 'doc_09_Unternehmenshandbuch',
  '10': 'doc_10_Auditbericht',
  '11': 'doc_11_Vollmacht',
  '12': 'doc_12_Firmeninfo_Foerdergeld',
  '13': 'doc_13_Projektbericht',
  '14': 'doc_14_Ausfuellanleitung'
};

// Dokumente die KEIN Google Doc sind (kein Logo, keine Text-Platzhalter)
const SKIP_TEXT_REPLACE = ['doc_05_Massnahmenplan'];

// Mapping: Spaltenname → tatsächlicher Platzhalter im Dokument
// (falls der Platzhalter anders geschrieben ist als die Spalte)
const PLACEHOLDER_ALIAS = {
  'companyName':             '{{FIRMENNAME}}',
  'Strasse':                 '{{Straße}}',
  'PLZ_Ort':                 '{{PLZ_Ort}}',
  'Ansprechpartner':         '{{Ansprechpartner}}',
  'email':                   '{{email}}',
  'Webpage':                 '{{Webpage}}',
  'Gruenderdatum':           '{{Gruenderdatum}}',
  'AnzahlderMitarbeiter':    '{{AnzahderMitarbtier}}',
  'Geltungsbereich':         '{{Geltungsbereich}}',
  'Zielgruppe_Zielgebiet':   '{{Zielgruppe/Zielgebiet}}',
  'Ausgelagerte_Prozesse':   '{{Ausgelagerte Prozesse}}',
  'Norm':                    '{{Norm}}',
  'AUDITOR':                 '{{AUDITOR}}',
  'Aktuelles_Jahr':          '{{Aktuelles_Jahr}}',
  'Datum_Heute':             '{{Datum_Heute}}'
};

// ===== WEB APP =====

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('BAFA Dokumente erstellen')
      .setWidth(1200)
      .setHeight(800);
}

// ===== KUNDENTABELLE ERSTELLEN =====

/**
 * Erstellt die Kundentabelle mit allen Spalten.
 * Wird einmalig manuell aufgerufen – danach CRM-Import.
 */
function createCustomerSheet() {
  var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var spreadsheet = SpreadsheetApp.create('BAFA Kundentabelle');
  var sheet = spreadsheet.getActiveSheet();
  sheet.setName('Kunden');

  // In richtigen Ordner verschieben
  var file = DriveApp.getFileById(spreadsheet.getId());
  parentFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Alle Headers zusammenbauen
  var allHeaders = []
    .concat(TABLE_COLUMNS.crm)
    .concat(TABLE_COLUMNS.placeholders)
    .concat(TABLE_COLUMNS.system)
    .concat(TABLE_COLUMNS.documents);

  // Headers schreiben
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);

  // === FORMATIERUNG ===
  var headerRange = sheet.getRange(1, 1, 1, allHeaders.length);
  headerRange.setFontWeight('bold');
  headerRange.setWrap(true);
  headerRange.setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  var crmLen = TABLE_COLUMNS.crm.length;
  var phLen  = TABLE_COLUMNS.placeholders.length;
  var sysLen = TABLE_COLUMNS.system.length;
  var docLen = TABLE_COLUMNS.documents.length;

  // CRM-Spalten (blau)
  sheet.getRange(1, 1, 1, crmLen).setBackground('#1565c0').setFontColor('#ffffff');
  // Platzhalter (grün)
  sheet.getRange(1, crmLen + 1, 1, phLen).setBackground('#2e7d32').setFontColor('#ffffff');
  // System (grau)
  sheet.getRange(1, crmLen + phLen + 1, 1, sysLen).setBackground('#616161').setFontColor('#ffffff');
  // Dokumente (lila)
  sheet.getRange(1, crmLen + phLen + sysLen + 1, 1, docLen).setBackground('#6a1b9a').setFontColor('#ffffff');

  // Spaltenbreiten
  for (var i = 1; i <= crmLen; i++) sheet.setColumnWidth(i, 180);
  for (var i = crmLen + 1; i <= crmLen + phLen; i++) sheet.setColumnWidth(i, 160);
  for (var i = crmLen + phLen + 1; i <= crmLen + phLen + sysLen; i++) sheet.setColumnWidth(i, 200);
  for (var i = crmLen + phLen + sysLen + 1; i <= allHeaders.length; i++) sheet.setColumnWidth(i, 180);

  // Sheet-ID speichern
  var sheetId = spreadsheet.getId();
  PropertiesService.getScriptProperties().setProperty('CUSTOMER_SHEET_ID', sheetId);

  Logger.log('=== KUNDENTABELLE ERSTELLT ===');
  Logger.log('Sheet-ID: ' + sheetId);
  Logger.log('URL: https://docs.google.com/spreadsheets/d/' + sheetId);
  Logger.log('Spalten: ' + allHeaders.length);

  return {
    sheetId: sheetId,
    url: 'https://docs.google.com/spreadsheets/d/' + sheetId,
    columns: allHeaders.length
  };
}

function getCustomerSheetId() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('CUSTOMER_SHEET_ID');
  if (!sheetId) {
    throw new Error('Kundentabelle noch nicht erstellt. Bitte zuerst über Menü erstellen.');
  }
  return sheetId;
}

function getCustomerSheet() {
  var sheetId = getCustomerSheetId();
  return SpreadsheetApp.openById(sheetId).getSheetByName('Kunden');
}

function setCustomerSheetId(sheetId) {
  PropertiesService.getScriptProperties().setProperty('CUSTOMER_SHEET_ID', sheetId);
  Logger.log('Sheet-ID gesetzt: ' + sheetId);
  return true;
}

// ===== ORDNER KOPIEREN =====

/**
 * BAFA Template kopieren für einen Kunden aus der Tabelle.
 * Schreibt Ordner-ID + alle Dokument-IDs zurück in die Tabelle.
 */
function copyMasterFolderWithRename(customerName, date, createArchive) {
  try {
    var sourceFolder = DriveApp.getFolderById(CONFIG.TEMPLATE_FOLDER_ID);
    var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);

    // Kundenordner erstellen
    var targetFolder = parentFolder.createFolder(customerName);

    // Kopieren mit Tracking
    var copiedFiles = [];
    copyFolderWithTracking(sourceFolder, targetFolder, copiedFiles);

    // Dateien umbenennen
    var replaceMap = {
      '_Kunde_Stand: Datum_': '_' + customerName + '_Stand: ' + date + '_',
      '_Kunde_Stand: Datum':  '_' + customerName + '_Stand: ' + date,
      '_Kunde_':              '_' + customerName + '_',
      'Template_Firma':       customerName
    };

    var renamedFiles = 0;

    copiedFiles.forEach(function(fileInfo) {
      try {
        var file = DriveApp.getFileById(fileInfo.id);
        var oldName = file.getName();
        var newName = oldName;
        var wasChanged = false;

        for (var search in replaceMap) {
          if (newName.indexOf(search) !== -1) {
            newName = newName.split(search).join(replaceMap[search]);
            wasChanged = true;
          }
        }

        if (wasChanged && newName !== oldName) {
          file.setName(newName);
          fileInfo.name = newName;
          renamedFiles++;
        }
      } catch (e) {
        Logger.log('Umbenennungsfehler: ' + e.toString());
      }
    });

    if (createArchive) {
      targetFolder.createFolder('Archiv');
    }

    // In Kundentabelle eintragen
    try {
      writeDocumentIdsToSheet(customerName, targetFolder.getId(), copiedFiles);
    } catch (sheetError) {
      Logger.log('Tabellen-Warnung: ' + sheetError.toString());
    }

    return {
      success: true,
      folderId: targetFolder.getId(),
      renamedFiles: renamedFiles,
      totalFiles: copiedFiles.length
    };

  } catch (error) {
    throw new Error('Kopierfehler: ' + error.toString());
  }
}

function copyFolderWithTracking(source, target, copiedFiles) {
  var files = source.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var copy = file.makeCopy(file.getName(), target);
    copiedFiles.push({
      id: copy.getId(),
      name: copy.getName(),
      mimeType: copy.getMimeType()
    });
  }
  var folders = source.getFolders();
  while (folders.hasNext()) {
    var subFolder = folders.next();
    var newSub = target.createFolder(subFolder.getName());
    copyFolderWithTracking(subFolder, newSub, copiedFiles);
  }
}

/**
 * Schreibt Ordner-ID und Dokument-IDs in die Kundenzeile der Tabelle
 */
function writeDocumentIdsToSheet(customerName, folderId, copiedFiles) {
  var sheet = getCustomerSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var companyCol = headers.indexOf('companyName');

  // Kundenzeile finden
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][companyCol] === customerName) {
      rowIndex = i + 1; // 1-basiert
      break;
    }
  }

  if (rowIndex === -1) {
    // Kunde nicht in Tabelle → neue Zeile
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, companyCol + 1).setValue(customerName);
    Logger.log('Neuer Kunde angelegt: ' + customerName);
  }

  // System-Spalten
  var folderIdCol   = headers.indexOf('folderID');
  var folderLinkCol = headers.indexOf('folderLink');
  var createdCol    = headers.indexOf('createdDate');

  if (folderIdCol !== -1) sheet.getRange(rowIndex, folderIdCol + 1).setValue(folderId);
  if (folderLinkCol !== -1) {
    sheet.getRange(rowIndex, folderLinkCol + 1).setFormula(
      '=HYPERLINK("https://drive.google.com/drive/folders/' + folderId + '","📁 Ordner öffnen")'
    );
  }
  if (createdCol !== -1) {
    sheet.getRange(rowIndex, createdCol + 1).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
    );
  }

  // Dokument-IDs zuordnen
  copiedFiles.forEach(function(fileInfo) {
    var match = fileInfo.name.match(/^(\d{2})[_\s]/);
    if (match) {
      var prefix = match[1];
      var columnName = DOC_COLUMN_MAP[prefix];
      if (columnName) {
        var colIndex = headers.indexOf(columnName);
        if (colIndex !== -1) {
          var docUrl = 'https://docs.google.com/document/d/' + fileInfo.id + '/edit';
          sheet.getRange(rowIndex, colIndex + 1).setFormula(
            '=HYPERLINK("' + docUrl + '","📄 ' + prefix + ' öffnen")'
          );
        }
      }
    }
  });

  Logger.log('Tabelle aktualisiert: ' + customerName + ' (' + copiedFiles.length + ' Dateien)');
  return true;
}

// ===== PLATZHALTER IN DOKUMENTE SCHREIBEN =====

/**
 * Liest Werte aus der Tabelle und ersetzt alle Platzhalter in den Dokumenten.
 */
function fillPlaceholdersForCustomer(customerName) {
  try {
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');

    // Kundenzeile finden
    var rowData = null;
    var rowIdx  = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][companyCol] === customerName) {
        rowData = data[i];
        rowIdx  = i;
        break;
      }
    }
    if (!rowData) throw new Error('Kunde "' + customerName + '" nicht gefunden.');

    // Platzhalter-Werte sammeln
    var replacements = {};
    var allPlaceholderCols = TABLE_COLUMNS.crm.concat(TABLE_COLUMNS.placeholders);

    allPlaceholderCols.forEach(function(colName) {
      var colIndex = headers.indexOf(colName);
      if (colIndex === -1) return;

      var value = rowData[colIndex];
      if (value === '' || value === null || value === undefined) return;

      // Alias-Platzhalter verwenden (z.B. companyName → {{FIRMENNAME}})
      var placeholder = PLACEHOLDER_ALIAS[colName] || ('{{' + colName + '}}');
      replacements[placeholder] = String(value);
    });

    // Automatische Werte falls leer
    if (!replacements['{{Aktuelles_Jahr}}']) {
      replacements['{{Aktuelles_Jahr}}'] = new Date().getFullYear().toString();
    }
    if (!replacements['{{Datum_Heute}}']) {
      replacements['{{Datum_Heute}}'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    }

    Logger.log('=== PLATZHALTER BEFÜLLUNG ===');
    Logger.log('Kunde: ' + customerName);
    Logger.log('Platzhalter: ' + Object.keys(replacements).length);
    for (var key in replacements) {
      Logger.log('  ' + key + ' → ' + replacements[key]);
    }

    // Logo-URL
    var logoUrlCol = headers.indexOf('logoUrl');
    var logoUrl = (logoUrlCol !== -1) ? rowData[logoUrlCol] : '';

    // Dokumente verarbeiten
    var processedDocs = 0;
    var errors = [];

    TABLE_COLUMNS.documents.forEach(function(docColName) {
      var colIndex = headers.indexOf(docColName);
      if (colIndex === -1) return;

      // Formel aus Zelle lesen um Doc-ID zu extrahieren
      var cell = sheet.getRange(rowIdx + 1, colIndex + 1);
      var formula = cell.getFormula();
      var cellValue = cell.getValue();

      var docId = null;
      if (formula) {
        var m = formula.match(/document\/d\/([a-zA-Z0-9_-]+)/);
        if (m) docId = m[1];
      }
      if (!docId && cellValue) {
        docId = extractDocIdFromCell(String(cellValue));
      }
      if (!docId) return;

      try {
        var driveFile = DriveApp.getFileById(docId);
        if (driveFile.getMimeType() !== MimeType.GOOGLE_DOCS) return;

        var doc = DocumentApp.openById(docId);
        var body = doc.getBody();
        var header = doc.getHeader();

        // Logo einfügen
        if (logoUrl) {
          insertLogoInDocument(doc, logoUrl);
        }

        // Platzhalter ersetzen
        for (var placeholder in replacements) {
          var escaped = escapeRegExp(placeholder);
          body.replaceText(escaped, replacements[placeholder]);
          if (header) header.replaceText(escaped, replacements[placeholder]);
        }

        doc.saveAndClose();
        processedDocs++;
        Logger.log('✓ ' + docColName);

      } catch (docError) {
        errors.push(docColName + ': ' + docError.toString());
        Logger.log('✗ ' + docColName + ': ' + docError.toString());
      }
    });

    Logger.log('=== ERGEBNIS: ' + processedDocs + ' Dokumente, ' + errors.length + ' Fehler ===');

    return {
      success: true,
      processedDocs: processedDocs,
      totalPlaceholders: Object.keys(replacements).length,
      errors: errors
    };

  } catch (error) {
    throw new Error('Platzhalter-Befüllung fehlgeschlagen: ' + error.toString());
  }
}

/**
 * Vorschau: Zeigt welche Platzhalter gefüllt/leer sind
 */
function previewPlaceholdersForCustomer(customerName) {
  try {
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');

    var rowData = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][companyCol] === customerName) { rowData = data[i]; break; }
    }
    if (!rowData) throw new Error('Kunde nicht gefunden');

    var filled = [];
    var empty  = [];
    var allCols = TABLE_COLUMNS.crm.concat(TABLE_COLUMNS.placeholders);

    allCols.forEach(function(colName) {
      var colIndex = headers.indexOf(colName);
      if (colIndex === -1) return;

      var placeholder = PLACEHOLDER_ALIAS[colName] || ('{{' + colName + '}}');
      var value = rowData[colIndex];

      if (value !== '' && value !== null && value !== undefined) {
        filled.push({ placeholder: placeholder, column: colName, value: String(value) });
      } else {
        empty.push({ placeholder: placeholder, column: colName });
      }
    });

    // Dokumente zählen
    var docCount = 0;
    TABLE_COLUMNS.documents.forEach(function(docCol) {
      var idx = headers.indexOf(docCol);
      if (idx !== -1 && rowData[idx]) docCount++;
    });

    return {
      customerName: customerName,
      filled: filled,
      empty: empty,
      documentCount: docCount
    };

  } catch (error) {
    throw new Error('Vorschau fehlgeschlagen: ' + error.toString());
  }
}

function extractDocIdFromCell(cellValue) {
  var str = String(cellValue);
  var m1 = str.match(/document\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  var m2 = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{25,}$/.test(str)) return str;
  return null;
}

// ===== BELIEBIGEN ORDNER KOPIEREN =====

function copyFolderByUrl(sourceUrl, newName) {
  try {
    var folderId = extractIdFromUrl(sourceUrl);
    var sourceFolder = DriveApp.getFolderById(folderId);
    var parents = sourceFolder.getParents();
    var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var targetFolder = parentFolder.createFolder(newName);
    var copiedFiles = [];
    copyFolderWithTracking(sourceFolder, targetFolder, copiedFiles);
    return { success: true, folderId: targetFolder.getId() };
  } catch (error) {
    throw new Error('Kopierfehler: ' + error.toString());
  }
}

// ===== DATEIEN UMBENENNEN =====

function renameFilesSimple(folderId, searchStrings, replaceString, recursive) {
  try {
    var mainFolder = DriveApp.getFolderById(folderId);
    var searchArray = Array.isArray(searchStrings) ? searchStrings : [searchStrings];
    var result = { totalFiles: 0, renamedFiles: 0, processedFolders: 0 };

    function processFolder(folder) {
      result.processedFolders++;
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var oldName = file.getName();
        result.totalFiles++;
        var newName = oldName;
        var wasChanged = false;
        for (var s = 0; s < searchArray.length; s++) {
          if (oldName.indexOf(searchArray[s]) !== -1) {
            newName = newName.replace(new RegExp(escapeRegExp(searchArray[s]), 'g'), replaceString);
            wasChanged = true;
          }
        }
        if (wasChanged && newName !== oldName) {
          file.setName(newName);
          result.renamedFiles++;
        }
      }
      if (recursive) {
        var subs = folder.getFolders();
        while (subs.hasNext()) processFolder(subs.next());
      }
    }

    processFolder(mainFolder);
    return result;
  } catch (error) {
    throw new Error('Umbenennungsfehler: ' + error.toString());
  }
}

function renameFilesFromUI(folderId, searchPattern, replacePattern, options) {
  try {
    if (!folderId || !searchPattern) throw new Error('Ordner-ID und Suchmuster erforderlich');
    var result = renameFilesSimple(folderId, searchPattern, replacePattern, options.recursive !== false);
    return {
      success: true,
      message: result.renamedFiles > 0
        ? result.renamedFiles + ' von ' + result.totalFiles + ' Dateien umbenannt'
        : 'Keine Dateien zum Umbenennen gefunden',
      details: { totalFiles: result.totalFiles, renamedFiles: result.renamedFiles }
    };
  } catch (error) {
    return { success: false, message: error.toString(), details: null };
  }
}

// ===== LOGO FUNKTIONEN =====

function getCompanyNames() {
  try {
    var sheet = getCustomerSheet();
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var companyCol = data[0].indexOf('companyName');
    return data.slice(1)
      .map(function(row) { return row[companyCol]; })
      .filter(function(name) { return name !== ''; });
  } catch (e) {
    return [];
  }
}

function getCustomerInfo(companyName) {
  try {
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');

    for (var i = 1; i < data.length; i++) {
      if (data[i][companyCol] === companyName) {
        var docCount = 0;
        TABLE_COLUMNS.documents.forEach(function(docCol) {
          var idx = headers.indexOf(docCol);
          if (idx !== -1 && data[i][idx]) docCount++;
        });
        return {
          companyName: companyName,
          folderId: data[i][headers.indexOf('folderID')] || null,
          hasLogo: !!data[i][headers.indexOf('logoUrl')],
          documentCount: docCount
        };
      }
    }
    return { companyName: companyName, folderId: null, hasLogo: false, documentCount: 0 };
  } catch (error) {
    throw new Error('Fehler: ' + error.toString());
  }
}

function getRecentFolders() {
  try {
    var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
    var folders = parentFolder.getFolders();
    var recent = [];
    while (folders.hasNext() && recent.length < 10) {
      var f = folders.next();
      recent.push({ id: f.getId(), name: f.getName(), created: f.getDateCreated() });
    }
    recent.sort(function(a, b) { return b.created - a.created; });
    return recent;
  } catch (error) {
    throw new Error('Fehler: ' + error.toString());
  }
}

/**
 * Logo hochladen UND direkt in alle Dokumente des Kunden einfügen.
 * 1. Logo in Logo-Ordner speichern
 * 2. Logo-URL in Kundentabelle speichern
 * 3. Ordner-ID + Doc-IDs aus Tabelle lesen
 * 4. {{LOGO_URL}} in allen Dokumenten durch das Logo-Bild ersetzen
 */
function uploadLogoAndProcessDocuments(logoBlob, companyName) {
  try {
    // 1. Logo hochladen + in Tabelle speichern
    var logoUrl = uploadLogo(logoBlob, companyName);
    Logger.log('Logo hochgeladen: ' + logoUrl);

    // 2. Dokument-IDs aus Tabelle lesen
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');

    var rowData = null;
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][companyCol] === companyName) {
        rowData = data[i];
        rowIdx = i;
        break;
      }
    }

    if (!rowData) {
      return { success: true, logoUrl: logoUrl, processedDocs: 0, message: 'Logo gespeichert, aber Kunde nicht in Tabelle gefunden.' };
    }

    // 3. Alle Dokumente durchgehen und Logo einfügen
    var processedDocs = 0;
    var errors = [];

    TABLE_COLUMNS.documents.forEach(function(docColName) {
      var colIndex = headers.indexOf(docColName);
      if (colIndex === -1) return;

      // Doc-ID aus Formel oder Wert extrahieren
      var cell = sheet.getRange(rowIdx + 1, colIndex + 1);
      var formula = cell.getFormula();
      var cellValue = cell.getValue();

      var docId = null;
      if (formula) {
        var m = formula.match(/document\/d\/([a-zA-Z0-9_-]+)/);
        if (m) docId = m[1];
      }
      if (!docId && cellValue) {
        docId = extractDocIdFromCell(String(cellValue));
      }
      if (!docId) return;

      try {
        var driveFile = DriveApp.getFileById(docId);
        if (driveFile.getMimeType() !== MimeType.GOOGLE_DOCS) return;

        var doc = DocumentApp.openById(docId);
        insertLogoInDocument(doc, logoUrl);
        doc.saveAndClose();
        processedDocs++;
        Logger.log('✓ Logo eingefügt: ' + docColName);

      } catch (docError) {
        errors.push(docColName + ': ' + docError.toString());
        Logger.log('✗ Logo-Fehler: ' + docColName + ': ' + docError.toString());
      }
    });

    Logger.log('=== LOGO ERGEBNIS: ' + processedDocs + ' Dokumente, ' + errors.length + ' Fehler ===');

    return {
      success: true,
      logoUrl: logoUrl,
      processedDocs: processedDocs,
      errors: errors
    };

  } catch (error) {
    throw new Error('Logo-Prozess fehlgeschlagen: ' + error.message);
  }
}

function uploadLogo(logoBlob, companyName) {
  try {
    var logoFolder = DriveApp.getFolderById(CONFIG.LOGO_FOLDER_ID);
    var mimeType = logoBlob.match(/data:([^;]+);/)[1];
    var ext = { 'image/jpeg':'jpg','image/jpg':'jpg','image/png':'png','image/gif':'gif','image/webp':'webp' };
    var fileExt = ext[mimeType.toLowerCase()];
    if (!fileExt) throw new Error('Nicht unterstütztes Format');

    var safeName = companyName.replace(/[^a-zA-Z0-9]/g, '_');
    var fileName = 'logo_' + safeName + '_' + new Date().getTime() + '.' + fileExt;
    var blob = Utilities.newBlob(Utilities.base64Decode(logoBlob.split(',')[1]), mimeType, fileName);

    var file = logoFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var logoUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();

    // In Tabelle speichern
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');
    var logoCol = headers.indexOf('logoUrl');

    for (var i = 1; i < data.length; i++) {
      if (data[i][companyCol] === companyName) {
        sheet.getRange(i + 1, logoCol + 1).setValue(logoUrl);
        break;
      }
    }
    return logoUrl;
  } catch (error) {
    throw new Error('Upload-Fehler: ' + error.message);
  }
}

function insertLogoInDocument(doc, logoUrl) {
  var body = doc.getBody();
  var header = doc.getHeader();
  var MAX_W = 200, MAX_H = 100;

  var doInsert = function(container, searchResult) {
    if (!searchResult) return;
    try {
      var el = searchResult.getElement();
      var offset = searchResult.getStartOffset();
      var logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();
      var para = el.getParent().asParagraph();
      var img = para.insertInlineImage(offset, logoBlob);
      var w = img.getWidth(), h = img.getHeight(), r = w / h;
      if (w > MAX_W) { w = MAX_W; h = w / r; }
      if (h > MAX_H) { h = MAX_H; w = h * r; }
      img.setWidth(w).setHeight(h);
      el.asText().deleteText(offset, offset + '{{LOGO_URL}}'.length - 1);
    } catch (e) { Logger.log('Logo-Fehler: ' + e); }
  };

  if (header) doInsert(header, header.findText('\\{\\{LOGO_URL\\}\\}'));
  doInsert(body, body.findText('\\{\\{LOGO_URL\\}\\}'));
}

// ===== CRM IMPORT AUS SUPER MASTER =====

/**
 * Holt Firmenliste aus dem CRM Super Master Sheet
 */
function getCrmCompanies() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.CRM_SUPER_MASTER_ID);
    var sheet = ss.getSheetByName(CONFIG.CRM_SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return data.flat().filter(function(name) {
      return name && name.toString().trim() !== '';
    });
  } catch (error) {
    Logger.log('CRM-Fehler: ' + error);
    throw new Error('CRM-Zugriff fehlgeschlagen: ' + error.toString());
  }
}

/**
 * Importiert einen Kunden aus dem CRM Super Master in die BAFA Kundentabelle.
 * Prüft ob Kunde bereits existiert und aktualisiert ggf.
 */
function importFromCrm(companyName) {
  try {
    if (!companyName || companyName.trim() === '') {
      throw new Error('Kein Firmenname ausgewählt');
    }

    Logger.log('=== CRM IMPORT: ' + companyName + ' ===');

    // 1. Daten aus Super Master holen
    var crmSS = SpreadsheetApp.openById(CONFIG.CRM_SUPER_MASTER_ID);
    var crmSheet = crmSS.getSheetByName(CONFIG.CRM_SHEET_NAME);
    var crmData = crmSheet.getRange(1, 1, crmSheet.getLastRow(), 1).getValues();

    var crmRowIndex = -1;
    for (var i = 0; i < crmData.length; i++) {
      if (crmData[i][0] && crmData[i][0].toString().trim() === companyName.trim()) {
        crmRowIndex = i + 1;
        break;
      }
    }

    if (crmRowIndex === -1) {
      throw new Error('Firma "' + companyName + '" nicht im CRM gefunden');
    }

    // 2. Werte aus den gemappten Spalten lesen
    var importedValues = {};
    for (var colLetter in CRM_IMPORT_MAP) {
      var targetCol = CRM_IMPORT_MAP[colLetter];
      var cellValue = crmSheet.getRange(colLetter + crmRowIndex).getValue();
      importedValues[targetCol] = cellValue ? cellValue.toString().trim() : '';
    }

    
    var plzCandidate = importedValues['Ansprechpartner'] || '';
    var ortCandidate = importedValues['PLZ_Ort'] || '';
    if (/^\d{4,5}$/.test(plzCandidate) && ortCandidate && !/\d/.test(ortCandidate)) {
      importedValues['PLZ_Ort'] = (plzCandidate + ' ' + ortCandidate).trim();
      importedValues['Ansprechpartner'] = '';
    }

    Logger.log('CRM-Daten: ' + JSON.stringify(importedValues));

    // 3. In BAFA Kundentabelle schreiben
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');

    // Prüfen ob Kunde schon existiert
    var existingRow = -1;
    for (var j = 1; j < data.length; j++) {
      if (data[j][companyCol] === companyName) {
        existingRow = j + 1;
        break;
      }
    }

    // Neue Zeile oder bestehende aktualisieren
    var targetRow = existingRow > 0 ? existingRow : sheet.getLastRow() + 1;
    var isNew = existingRow <= 0;

    for (var col in importedValues) {
      var colIndex = headers.indexOf(col);
      if (colIndex !== -1 && importedValues[col]) {
        sheet.getRange(targetRow, colIndex + 1).setValue(importedValues[col]);
      }
    }

    // Standardwerte setzen
    var normCol = headers.indexOf('Norm');
    var auditorCol = headers.indexOf('AUDITOR');
    if (normCol !== -1 && !sheet.getRange(targetRow, normCol + 1).getValue()) {
      sheet.getRange(targetRow, normCol + 1).setValue('ISO 9001:2015');
    }
    if (auditorCol !== -1 && !sheet.getRange(targetRow, auditorCol + 1).getValue()) {
      sheet.getRange(targetRow, auditorCol + 1).setValue('Holger Grosser');
    }

    Logger.log('=== CRM IMPORT ABGESCHLOSSEN: Zeile ' + targetRow + (isNew ? ' (NEU)' : ' (AKTUALISIERT)') + ' ===');

    return {
      success: true,
      isNew: isNew,
      row: targetRow,
      importedFields: Object.keys(importedValues).filter(function(k) { return importedValues[k]; }).length,
      companyName: companyName
    };

  } catch (error) {
    Logger.log('CRM Import Fehler: ' + error);
    throw new Error('Import fehlgeschlagen: ' + error.toString());
  }
}

// ===== KI-ZUORDNUNG: FREITEXT → PLATZHALTER =====

/**
 * Claude analysiert Freitext und ordnet Inhalte den Platzhaltern zu.
 * Gibt Zuordnungs-Vorschlag zurück zur manuellen Prüfung.
 */
function analyzeWithAI(freeText, companyName) {
  try {
    var claudeKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!claudeKey) {
      throw new Error('Claude API-Key fehlt! Bitte über Menü "🔑 Claude Key einrichten" setzen.');
    }

    // Alle verfügbaren Platzhalter-Spalten
    var allColumns = TABLE_COLUMNS.crm.concat(TABLE_COLUMNS.placeholders);

    // Bestehende Werte laden falls vorhanden
    var existingValues = {};
    try {
      var sheet = getCustomerSheet();
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var companyCol = headers.indexOf('companyName');
      for (var i = 1; i < data.length; i++) {
        if (data[i][companyCol] === companyName) {
          allColumns.forEach(function(col) {
            var idx = headers.indexOf(col);
            if (idx !== -1 && data[i][idx]) {
              existingValues[col] = String(data[i][idx]);
            }
          });
          break;
        }
      }
    } catch (e) {
      Logger.log('Bestehende Werte nicht geladen: ' + e);
    }

    // Platzhalter-Beschreibungen für Claude
    var fieldDescriptions = {
      'companyName': 'Firmenname (z.B. Müller GmbH)',
      'Strasse': 'Straße und Hausnummer',
      'PLZ_Ort': 'Postleitzahl und Ort (z.B. 90402 Nürnberg)',
      'Ansprechpartner': 'Kontaktperson / Geschäftsführer',
      'email': 'E-Mail-Adresse',
      'Webpage': 'Website-URL',
      'Gruenderdatum': 'Gründungsjahr oder -datum',
      'AnzahlderMitarbeiter': 'Anzahl der Mitarbeiter',
      'Geltungsbereich': 'Geltungsbereich der Zertifizierung (Branche/Tätigkeit)',
      'Zielgruppe_Zielgebiet': 'Zielgruppe und Zielgebiet des Unternehmens',
      'Ausgelagerte_Prozesse': 'Ausgelagerte Prozesse (z.B. Buchhaltung, IT)',
      'Norm': 'ISO-Norm (z.B. ISO 9001:2015)',
      'AUDITOR': 'Name des Auditors',
      'Aktuelles_Jahr': 'Aktuelles Jahr (automatisch)',
      'Datum_Heute': 'Heutiges Datum (automatisch)'
    };

    var fieldsPrompt = allColumns.map(function(col) {
      var desc = fieldDescriptions[col] || col;
      var existing = existingValues[col] ? ' [AKTUELL: ' + existingValues[col] + ']' : '';
      return '- ' + col + ': ' + desc + existing;
    }).join('\n');

    var prompt = 'Analysiere den folgenden Text für die Firma "' + companyName + '".\n\n' +
      'Ordne die gefundenen Informationen den folgenden Platzhalter-Feldern zu:\n\n' +
      fieldsPrompt + '\n\n' +
      'TEXT ZUM ANALYSIEREN:\n' + freeText + '\n\n' +
      'REGELN:\n' +
      '- Nur Felder zurückgeben, für die du Informationen im Text findest\n' +
      '- Bei mehreren möglichen Werten den wahrscheinlichsten wählen\n' +
      '- Felder die bereits einen Wert haben [AKTUELL:...] nur überschreiben wenn der neue Wert besser/aktueller ist\n' +
      '- Aktuelles_Jahr und Datum_Heute NICHT zuordnen (werden automatisch gesetzt)\n' +
      '- Werte sauber formatieren (z.B. PLZ mit Leerzeichen vor Ort)\n\n' +
      'Antworte NUR mit JSON-Objekt, keine Erklärungen:\n' +
      '{"feldname": "wert", ...}';

    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': claudeKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1500,
        temperature: 0.2,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      var errText = response.getContentText();
      Logger.log('Claude API Error: ' + errText);
      throw new Error('Claude API Error: HTTP ' + response.getResponseCode());
    }

    var respData = JSON.parse(response.getContentText());
    var text = respData.content[0].text;
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('Keine JSON-Antwort von Claude');
    }

    var assignments = JSON.parse(jsonMatch[0]);
    Logger.log('KI-Zuordnung: ' + JSON.stringify(assignments));

    // Ergebnis-Array für UI erstellen
    var result = [];
    allColumns.forEach(function(col) {
      if (col === 'Aktuelles_Jahr' || col === 'Datum_Heute') return;

      var entry = {
        column: col,
        description: fieldDescriptions[col] || col,
        existingValue: existingValues[col] || '',
        aiValue: assignments[col] || '',
        useAI: !!assignments[col]
      };
      result.push(entry);
    });

    return {
      success: true,
      assignments: result,
      rawText: freeText.substring(0, 200) + '...'
    };

  } catch (error) {
    Logger.log('KI-Zuordnung Fehler: ' + error);
    throw new Error('KI-Analyse fehlgeschlagen: ' + error.toString());
  }
}

/**
 * Speichert die geprüften Zuordnungen in die Kundentabelle.
 * confirmedData = [{column: 'email', value: 'info@firma.de'}, ...]
 */
function saveAIAssignments(companyName, confirmedData) {
  try {
    var sheet = getCustomerSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var companyCol = headers.indexOf('companyName');

    // Kundenzeile finden oder erstellen
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][companyCol] === companyName) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
      sheet.getRange(targetRow, companyCol + 1).setValue(companyName);
    }

    // Werte schreiben
    var savedCount = 0;
    confirmedData.forEach(function(item) {
      var colIndex = headers.indexOf(item.column);
      if (colIndex !== -1 && item.value && item.value.trim() !== '') {
        sheet.getRange(targetRow, colIndex + 1).setValue(item.value.trim());
        savedCount++;
      }
    });

    Logger.log('=== KI-ZUORDNUNG GESPEICHERT: ' + savedCount + ' Felder für ' + companyName + ' ===');

    return {
      success: true,
      savedFields: savedCount,
      companyName: companyName,
      row: targetRow
    };

  } catch (error) {
    Logger.log('Speichern fehlgeschlagen: ' + error);
    throw new Error('Speichern fehlgeschlagen: ' + error.toString());
  }
}

/**
 * Claude API-Key einrichten
 */
function setupClaudeApiKey() {
  var ui = SpreadsheetApp.getUi();
  var existing = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  var msg = existing ? 'Key ist gesetzt (sk-ant-...' + existing.slice(-8) + ')\nNeuen Key eingeben oder Cancel:' : 'Claude API-Key eingeben (sk-ant-...):';

  var response = ui.prompt('🔑 Claude API-Key', msg, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var key = response.getResponseText().trim();
    if (key && key.startsWith('sk-ant-')) {
      PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY', key);
      ui.alert('✅ Claude Key gespeichert!');
    } else {
      ui.alert('❌ Ungültiger Key (muss mit sk-ant- beginnen)');
    }
  }
}

// ===== PDF FUNKTIONEN =====

function convertToPDF(parentFolderId, createDateFolder) {
  try {
    var folder = DriveApp.getFolderById(parentFolderId);
    var pdfFolder = folder;
    if (createDateFolder) {
      pdfFolder = folder.createFolder('PDF ' + Utilities.formatDate(new Date(), 'GMT+1', 'yyyy-MM-dd'));
    }
    var files = folder.getFiles();
    var fileCount = 0, pdfCount = 0;
    while (files.hasNext()) {
      var f = files.next(); fileCount++;
      try {
        var pdfBlob = null;
        if (f.getMimeType() === MimeType.GOOGLE_DOCS) {
          pdfBlob = f.getAs('application/pdf');
        } else if (f.getMimeType() === MimeType.MICROSOFT_EXCEL || f.getMimeType() === MimeType.MICROSOFT_EXCEL_LEGACY) {
          var tmp = Drive.Files.copy({ title: f.getName().replace(/\.[^/.]+$/, ''), mimeType: MimeType.GOOGLE_SHEETS }, f.getId());
          Utilities.sleep(2000);
          var gs = DriveApp.getFileById(tmp.id);
          pdfBlob = gs.getAs('application/pdf');
          gs.setTrashed(true);
        }
        if (pdfBlob) {
          pdfBlob.setName(f.getName().replace(/\.[^/.]+$/, '') + '.pdf');
          pdfFolder.createFile(pdfBlob);
          pdfCount++;
        }
      } catch (e) { Logger.log('PDF-Fehler: ' + f.getName() + ': ' + e); }
    }
    return { fileCount: fileCount, pdfCount: pdfCount };
  } catch (error) {
    throw new Error('PDF-Fehler: ' + error.toString());
  }
}

function splitMultiplePDFs(pdfUrls) {
  var ok = 0, err = 0, details = [];
  for (var i = 0; i < pdfUrls.length; i++) {
    var url = pdfUrls[i].trim();
    if (!url) continue;
    try {
      var r = splitPDFWithPdfCo(url);
      if (r.success) { ok++; details.push('✓ ' + r.message); }
      else { err++; details.push('✗ ' + r.error); }
    } catch (e) { err++; details.push('✗ ' + e.toString()); }
  }
  return { success: ok, errors: err, details: details };
}

function splitPDFWithPdfCo(fileUrl) {
  try {
    var fileId = extractIdFromUrl(fileUrl);
    if (!fileId) return { success: false, error: 'Ungültige URL' };
    var pdfFile = DriveApp.getFileById(fileId);
    var fileName = pdfFile.getName().replace('.pdf', '');

    var upResp = UrlFetchApp.fetch(CONFIG.PDF_CO_URL + '/file/upload', {
      method: 'POST', headers: { 'x-api-key': CONFIG.PDF_API_KEY },
      payload: { file: pdfFile.getBlob() }, muteHttpExceptions: true
    });
    var upJson = JSON.parse(upResp.getContentText());
    if (upJson.error || !upJson.url) return { success: false, error: 'Upload fehlgeschlagen' };

    var splitResp = UrlFetchApp.fetch(CONFIG.PDF_CO_URL + '/pdf/split', {
      method: 'POST',
      headers: { 'x-api-key': CONFIG.PDF_API_KEY, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ url: upJson.url, async: false, split_mode: 'extract', pages: '*' }),
      muteHttpExceptions: true
    });
    var splitJson = JSON.parse(splitResp.getContentText());
    if (splitJson.error || !splitJson.urls || !splitJson.urls.length) return { success: false, error: 'Split fehlgeschlagen' };

    var folder = pdfFile.getParents().next().createFolder(fileName + '_Seiten');
    for (var p = 0; p < splitJson.urls.length; p++) {
      var blob = UrlFetchApp.fetch(splitJson.urls[p]).getBlob();
      folder.createFile(blob.setName(fileName + '-' + (p + 1) + '.pdf'));
    }
    return { success: true, message: fileName + ': ' + splitJson.urls.length + ' Seiten' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ===== STAPELVERARBEITUNG =====

function runBatchProcess(actions, customerName, date) {
  try {
    var results = [];
    var folderId = null;

    if (actions.copyMaster) {
      var cr = copyMasterFolderWithRename(customerName, date, true);
      folderId = cr.folderId;
      results.push('✓ Template kopiert (' + cr.renamedFiles + ' umbenannt)');
    }
    if (actions.fillPlaceholders) {
      var fr = fillPlaceholdersForCustomer(customerName);
      results.push('✓ ' + fr.processedDocs + ' Dokumente mit Platzhaltern befüllt');
    }
    if (actions.uploadLogo) {
      results.push('⚠ Logo-Upload separat ausführen');
    }
    if (actions.createPdf && folderId) {
      var pr = convertToPDF(folderId, true);
      results.push('✓ ' + pr.pdfCount + ' PDFs erstellt');
    }

    return {
      success: true,
      summary: '<ul>' + results.map(function(r) { return '<li>' + r + '</li>'; }).join('') + '</ul>'
    };
  } catch (error) {
    throw new Error('Stapelfehler: ' + error.toString());
  }
}

// ===== HILFSFUNKTIONEN =====

function getDatePlus3BusinessDays() {
  var d = new Date();
  var added = 0;
  while (added < 3) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) added++;
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getDefaultDate() {
  return getDatePlus3BusinessDays();
}

function extractIdFromUrl(url) {
  var patterns = [
    /\/folders\/([a-zA-Z0-9-_]+)/,
    /\/file\/d\/([a-zA-Z0-9-_]+)/,
    /\/d\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = url.match(patterns[i]);
    if (m) return m[1];
  }
  throw new Error('Keine gültige ID gefunden');
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ===== MENÜ =====

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 BAFA Manager')
    .addItem('📊 Manager öffnen', 'showWebApp')
    .addSeparator()
    .addItem('🆕 Kundentabelle erstellen', 'createCustomerSheetMenu')
    .addItem('📋 Kundentabelle öffnen', 'openCustomerSheet')
    .addSeparator()
    .addItem('📥 CRM-Import', 'showCrmImportDialog')
    .addItem('🤖 KI-Zuordnung (Freitext)', 'showAIAssignDialog')
    .addSeparator()
    .addItem('📁 Template kopieren', 'showMasterCopyDialog')
    .addItem('📝 Platzhalter befüllen', 'showFillPlaceholdersDialog')
    .addItem('🖼️ Logo hochladen', 'showLogoDialog')
    .addItem('📑 PDFs erstellen', 'showPdfDialog')
    .addItem('✂️ PDFs aufteilen', 'showSplitDialog')
    .addSeparator()
    .addItem('🔑 Claude Key einrichten', 'setupClaudeApiKey')
    .addToUi();
}

function createCustomerSheetMenu() {
  var result = createCustomerSheet();
  SpreadsheetApp.getUi().alert(
    'Kundentabelle erstellt!\n\nURL: ' + result.url + '\nSpalten: ' + result.columns
  );
}

function openCustomerSheet() {
  var sheetId = getCustomerSheetId();
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("https://docs.google.com/spreadsheets/d/' + sheetId + '","_blank");google.script.host.close();</script>'
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Öffne...');
}

function showWebApp() {
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'BAFA Dokumente erstellen');
}

function showMasterCopyDialog() {
  var defaultDate = getDatePlus3BusinessDays();
  var companies = getCompanyNames();
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;padding:20px}.form-group{margin-bottom:15px}label{display:block;margin-bottom:5px;font-weight:bold}input,select{width:100%;padding:8px;box-sizing:border-box}button{background:#4CAF50;color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer}#status{margin-top:15px;padding:10px;border-radius:4px;display:none}.success{background:#d4edda;color:#155724}.error{background:#f8d7da;color:#721c24}.info{background:#d1ecf1;color:#0c5460;margin-bottom:15px;padding:10px;border-radius:4px}</style>' +
    '<h3>BAFA Template kopieren</h3>' +
    '<div class="info">ℹ️ Kunde muss in der Kundentabelle existieren (CRM-Import)</div>' +
    '<div class="form-group"><label>Kunde:</label><select id="c"><option value="">-- Bitte wählen --</option>' +
    companies.map(function(c) { return '<option value="' + c + '">' + c + '</option>'; }).join('') +
    '</select></div>' +
    '<div class="form-group"><label>Datum (Stand):</label><input type="date" id="d" value="' + defaultDate + '"></div>' +
    '<button onclick="go()">Template kopieren</button><div id="status"></div>' +
    '<script>function go(){var c=document.getElementById("c").value,d=document.getElementById("d").value;if(!c||!d){s("Bitte alle Felder ausfüllen","error");return}s("Wird kopiert...","info");google.script.run.withSuccessHandler(function(r){s("✓ Kopiert! "+r.renamedFiles+" umbenannt, IDs in Tabelle.","success")}).withFailureHandler(function(e){s("Fehler: "+e,"error")}).copyMasterFolderWithRename(c,d,true)}function s(m,t){var el=document.getElementById("status");el.textContent=m;el.className=t;el.style.display="block"}</script>'
  ).setWidth(450).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'BAFA Template kopieren');
}

function showFillPlaceholdersDialog() {
  var companies = getCompanyNames();
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;padding:20px}.form-group{margin-bottom:15px}label{display:block;margin-bottom:5px;font-weight:bold}select{width:100%;padding:8px;box-sizing:border-box}button{color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;margin-right:10px}.btn-preview{background:#2196F3}.btn-fill{background:#4CAF50}#status{margin-top:15px;padding:10px;border-radius:4px;display:none}.success{background:#d4edda;color:#155724}.error{background:#f8d7da;color:#721c24}.info{background:#d1ecf1;color:#0c5460}#preview{margin-top:15px;display:none;max-height:300px;overflow-y:auto}table{width:100%;border-collapse:collapse;margin-top:10px}td,th{padding:6px 10px;border:1px solid #ddd;text-align:left;font-size:13px}th{background:#f0f0f0}.empty{color:#dc3545}.filled{color:#28a745}</style>' +
    '<h3>📝 Platzhalter befüllen</h3>' +
    '<div class="form-group"><label>Kunde:</label><select id="c"><option value="">-- Bitte wählen --</option>' +
    companies.map(function(c) { return '<option value="' + c + '">' + c + '</option>'; }).join('') +
    '</select></div>' +
    '<button class="btn-preview" onclick="preview()">👁️ Vorschau</button>' +
    '<button class="btn-fill" onclick="fill()">📝 Befüllen</button>' +
    '<div id="status"></div><div id="preview"></div>' +
    '<script>function preview(){var c=document.getElementById("c").value;if(!c){s("Bitte Kunde wählen","error");return}s("Lade...","info");google.script.run.withSuccessHandler(function(r){var h="<h4>"+r.customerName+" ("+r.documentCount+" Dokumente)</h4><table><tr><th>Platzhalter</th><th>Wert</th></tr>";r.filled.forEach(function(f){h+="<tr><td>"+f.placeholder+"</td><td class=filled>✓ "+f.value+"</td></tr>"});r.empty.forEach(function(e){h+="<tr><td>"+e.placeholder+"</td><td class=empty>— leer</td></tr>"});h+="</table>";document.getElementById("preview").innerHTML=h;document.getElementById("preview").style.display="block";document.getElementById("status").style.display="none"}).withFailureHandler(function(e){s("Fehler: "+e,"error")}).previewPlaceholdersForCustomer(c)}function fill(){var c=document.getElementById("c").value;if(!c){s("Bitte Kunde wählen","error");return}if(!confirm("Platzhalter für \\""+c+"\\" in alle Dokumente schreiben?"))return;s("Wird befüllt...","info");google.script.run.withSuccessHandler(function(r){var m="✓ "+r.processedDocs+" Dokumente, "+r.totalPlaceholders+" Platzhalter";if(r.errors.length>0)m+=" ("+r.errors.length+" Fehler)";s(m,"success")}).withFailureHandler(function(e){s("Fehler: "+e,"error")}).fillPlaceholdersForCustomer(c)}function s(m,t){var el=document.getElementById("status");el.textContent=m;el.className=t;el.style.display="block"}</script>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Platzhalter befüllen');
}

function showLogoDialog() {
  var companies = getCompanyNames();
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:20px}.form-group{margin-bottom:15px}label{display:block;margin-bottom:5px;font-weight:bold}select,input{width:100%;padding:8px;box-sizing:border-box}button{background:#4CAF50;color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer}#status{margin-top:15px;padding:10px;border-radius:4px;display:none}.success{background:#d4edda;color:#155724}.error{background:#f8d7da;color:#721c24}.info{background:#d1ecf1;color:#0c5460}#preview{max-width:200px;max-height:100px;margin-top:10px;display:none}#custInfo{margin:10px 0;padding:10px;background:#f8f9fa;border-radius:4px;display:none;font-size:13px}</style>' +
    '<h3>🖼️ Logo hochladen & einfügen</h3>' +
    '<div class="form-group"><label>Kunde:</label><select id="co" onchange="loadInfo()">' +
    '<option value="">-- Bitte wählen --</option>' +
    companies.map(function(c) { return '<option value="' + c + '">' + c + '</option>'; }).join('') +
    '</select></div>' +
    '<div id="custInfo"></div>' +
    '<div class="form-group"><label>Logo:</label>' +
    '<input type="file" id="logo" accept="image/*" onchange="var r=new FileReader();r.onload=function(e){document.getElementById(\'preview\').src=e.target.result;document.getElementById(\'preview\').style.display=\'block\'};r.readAsDataURL(this.files[0])">' +
    '<img id="preview"></div>' +
    '<button onclick="up()">🖼️ Hochladen & in Dokumente einfügen</button><div id="status"></div>' +
    '<script>' +
    'function loadInfo(){var c=document.getElementById("co").value;var d=document.getElementById("custInfo");if(!c){d.style.display="none";return}' +
    'google.script.run.withSuccessHandler(function(i){var h="";h+=i.folderId?"📁 Ordner vorhanden":"⚠️ Kein Ordner";h+=" | "+i.documentCount+" Dokumente";h+=i.hasLogo?" | ✓ Logo vorhanden":" | Kein Logo";d.innerHTML=h;d.style.display="block"}).getCustomerInfo(c)}' +
    'function up(){var c=document.getElementById("co").value,f=document.getElementById("logo");if(!c||!f.files[0]){s("Bitte alles ausfüllen","error");return}s("Logo wird hochgeladen und eingefügt...","info");var r=new FileReader();r.onloadend=function(){google.script.run.withSuccessHandler(function(res){var m="✓ Logo hochgeladen! "+res.processedDocs+" Dokumente aktualisiert.";if(res.errors&&res.errors.length>0)m+=" ("+res.errors.length+" Fehler)";s(m,"success")}).withFailureHandler(function(e){s("Fehler: "+e,"error")}).uploadLogoAndProcessDocuments(r.result,c)};r.readAsDataURL(f.files[0])}' +
    'function s(m,t){var el=document.getElementById("status");el.textContent=m;el.className=t;el.style.display="block"}' +
    '</script>'
  ).setWidth(450).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Logo hochladen & einfügen');
}

function showPdfDialog() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:20px}label{display:block;margin-bottom:5px;font-weight:bold}input{width:100%;padding:8px;box-sizing:border-box;margin-bottom:15px}button{background:#4CAF50;color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer}#status{margin-top:15px;padding:10px;border-radius:4px;display:none}.success{background:#d4edda;color:#155724}.error{background:#f8d7da;color:#721c24}</style>' +
    '<h3>PDFs erstellen</h3><label>Ordner ID:</label><input type="text" id="fid" placeholder="Google Drive Ordner ID">' +
    '<label><input type="checkbox" id="cf" checked> PDF-Ordner mit Datum</label><br><br>' +
    '<button onclick="go()">PDFs erstellen</button><div id="status"></div>' +
    '<script>function go(){var id=document.getElementById("fid").value;if(!id){s("Bitte ID eingeben","error");return}s("Erstelle...","info");google.script.run.withSuccessHandler(function(r){s("✓ "+r.pdfCount+" PDFs erstellt!","success")}).withFailureHandler(function(e){s("Fehler: "+e,"error")}).convertToPDF(id,document.getElementById("cf").checked)}function s(m,t){var el=document.getElementById("status");el.textContent=m;el.className=t;el.style.display="block"}</script>'
  ).setWidth(400).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'PDFs erstellen');
}

function showSplitDialog() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:20px}label{display:block;margin-bottom:5px;font-weight:bold}textarea{width:100%;padding:8px;box-sizing:border-box;min-height:100px}button{background:#4CAF50;color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer}#status{margin-top:15px;padding:10px;border-radius:4px;display:none}.success{background:#d4edda;color:#155724}.error{background:#f8d7da;color:#721c24}</style>' +
    '<h3>PDFs aufteilen</h3><label>PDF URLs (eine pro Zeile):</label>' +
    '<textarea id="urls" placeholder="https://drive.google.com/file/d/..."></textarea><br><br>' +
    '<button onclick="go()">Aufteilen</button><div id="status"></div>' +
    '<script>function go(){var urls=document.getElementById("urls").value.trim().split("\\n").filter(function(u){return u});if(!urls.length){s("Bitte URLs eingeben","error");return}s("Teile auf...","info");google.script.run.withSuccessHandler(function(r){s("✓ "+r.success+" PDFs aufgeteilt!","success")}).withFailureHandler(function(e){s("Fehler: "+e,"error")}).splitMultiplePDFs(urls)}function s(m,t){var el=document.getElementById("status");el.textContent=m;el.className=t;el.style.display="block"}</script>'
  ).setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'PDFs aufteilen');
}

// ===== CRM IMPORT DIALOG =====

function showCrmImportDialog() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:20px}' +
    '.form-group{margin-bottom:15px}' +
    'label{display:block;margin-bottom:5px;font-weight:bold}' +
    'select{width:100%;padding:8px;box-sizing:border-box}' +
    'input[type=text]{width:100%;padding:6px;box-sizing:border-box;margin-bottom:8px}' +
    'button{background:#1a73e8;color:#fff;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;font-size:14px}' +
    'button:hover{background:#1557b0}' +
    'button:disabled{background:#ccc;cursor:not-allowed}' +
    '#status{margin-top:15px;padding:12px;border-radius:4px;display:none;font-size:13px}' +
    '.success{background:#d4edda;color:#155724}' +
    '.error{background:#f8d7da;color:#721c24}' +
    '.info{background:#d1ecf1;color:#0c5460}' +
    '.hint{background:#e8f0fe;border-left:4px solid #1a73e8;padding:12px;margin-bottom:15px;border-radius:4px;font-size:12px;line-height:1.6}' +
    '</style>' +
    '<h3>📥 CRM-Import aus Super Master</h3>' +
    '<div class="hint">Importiert Stammdaten (Firma, Adresse, Kontakt, E-Mail) aus dem CRM in die BAFA Kundentabelle.<br>Bestehende Kunden werden aktualisiert, neue werden angelegt.</div>' +
    '<div class="form-group">' +
    '<label>🔍 Firma suchen:</label>' +
    '<input type="text" id="search" placeholder="Tippen zum Filtern..." onkeyup="filterList()">' +
    '<select id="company" size="8"><option disabled>Lade CRM-Firmenliste...</option></select>' +
    '</div>' +
    '<button id="btn" onclick="doImport()" disabled>📥 Importieren</button>' +
    '<div id="status"></div>' +
    '<script>' +
    'var allCo=[];' +
    'google.script.run.withSuccessHandler(function(list){' +
    '  allCo=list;renderList(list);document.getElementById("btn").disabled=false;' +
    '}).withFailureHandler(function(e){' +
    '  showSt("Fehler: "+e,"error");' +
    '}).getCrmCompanies();' +
    'function renderList(arr){var s=document.getElementById("company");s.innerHTML="";arr.forEach(function(c){var o=document.createElement("option");o.value=c;o.text=c;s.appendChild(o)});if(!arr.length){var o=document.createElement("option");o.text="Keine Treffer";o.disabled=true;s.appendChild(o)}}' +
    'function filterList(){var q=document.getElementById("search").value.toLowerCase();renderList(allCo.filter(function(c){return c.toLowerCase().indexOf(q)>-1}))}' +
    'function doImport(){var c=document.getElementById("company").value;if(!c){showSt("Bitte Firma wählen","error");return}document.getElementById("btn").disabled=true;showSt("⏳ Importiere "+c+"...","info");' +
    'google.script.run.withSuccessHandler(function(r){' +
    '  showSt("✅ "+r.companyName+(r.isNew?" NEU angelegt":" aktualisiert")+" ("+r.importedFields+" Felder) in Zeile "+r.row,"success");' +
    '  document.getElementById("btn").disabled=false;' +
    '}).withFailureHandler(function(e){showSt("❌ "+e,"error");document.getElementById("btn").disabled=false}).importFromCrm(c)}' +
    'function showSt(m,t){var el=document.getElementById("status");el.innerHTML=m;el.className=t;el.style.display="block"}' +
    '</script>'
  ).setWidth(450).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'CRM-Import');
}

// ===== KI-ZUORDNUNGS-DIALOG =====

function showAIAssignDialog() {
  var companies = getCompanyNames();
  var companyOptions = companies.map(function(c) {
    return '<option value="' + c.replace(/"/g, '&quot;') + '">' + c + '</option>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{font-family:Arial,sans-serif;padding:15px;margin:0;background:#f5f5f5}' +
    'h3{color:#1a73e8;margin:0 0 10px}' +
    '.hint{background:#e8f0fe;border-left:4px solid #1a73e8;padding:10px;margin-bottom:12px;border-radius:4px;font-size:12px;line-height:1.5}' +
    '.form-group{margin-bottom:12px}' +
    'label{display:block;margin-bottom:4px;font-weight:bold;font-size:13px}' +
    'select,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;font-family:inherit;font-size:13px}' +
    'textarea{min-height:120px;resize:vertical}' +
    'button{padding:10px 20px;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:600;margin-right:8px}' +
    '.btn-blue{background:#1a73e8;color:#fff}.btn-blue:hover{background:#1557b0}' +
    '.btn-green{background:#22c55e;color:#fff}.btn-green:hover{background:#16a34a}' +
    '.btn-gray{background:#e8eaed;color:#5f6368}.btn-gray:hover{background:#dadce0}' +
    'button:disabled{opacity:0.5;cursor:not-allowed}' +
    '#status{margin-top:10px;padding:10px;border-radius:4px;display:none;font-size:13px}' +
    '.success{background:#d4edda;color:#155724}.error{background:#f8d7da;color:#721c24}.info{background:#d1ecf1;color:#0c5460}' +
    '#resultTable{width:100%;border-collapse:collapse;margin-top:10px;font-size:12px}' +
    '#resultTable th{background:#1a73e8;color:#fff;padding:8px;text-align:left}' +
    '#resultTable td{padding:6px 8px;border-bottom:1px solid #e0e0e0}' +
    '#resultTable tr:hover{background:#f5f5f5}' +
    '#resultTable input[type=text]{width:100%;padding:4px;border:1px solid #ddd;border-radius:3px;font-size:12px}' +
    '#resultTable input[type=checkbox]{transform:scale(1.2)}' +
    '.existing{color:#666;font-size:11px;font-style:italic}' +
    '.ai-new{background:#f0fdf4}' +
    '.ai-update{background:#fef3c7}' +
    '#resultArea{display:none;margin-top:15px;background:#fff;padding:15px;border-radius:8px;border:1px solid #e0e0e0}' +
    '</style>' +

    '<h3>🤖 KI-Zuordnung: Freitext → Platzhalter</h3>' +
    '<div class="hint">' +
    '<strong>So gehts:</strong> 1) Kunde wählen → 2) Kundendaten reinkopieren (E-Mail, PDF, Notizen...) → 3) Claude ordnet zu → 4) Prüfen & Speichern' +
    '</div>' +

    '<div class="form-group">' +
    '<label>Kunde:</label>' +
    '<select id="company"><option value="">-- Bitte wählen --</option>' + companyOptions + '</select>' +
    '</div>' +

    '<div class="form-group">' +
    '<label>Kundendaten (Freitext):</label>' +
    '<textarea id="freetext" placeholder="Hier alles reinkopieren was der Kunde geliefert hat...\n\nZ.B.:\nFirma Müller GmbH\nHauptstr. 5, 90402 Nürnberg\nGF: Thomas Müller\ninfo@mueller.de\nwww.mueller.de\nGegründet 2005\n12 Mitarbeiter\nDruckerei / Digitaldruck\nISO 9001 Zertifizierung gewünscht"></textarea>' +
    '</div>' +

    '<button class="btn-blue" id="analyzeBtn" onclick="analyze()">🤖 Claude analysieren lassen</button>' +
    '<button class="btn-gray" onclick="clearAll()">🗑️ Zurücksetzen</button>' +
    '<div id="status"></div>' +

    '<div id="resultArea">' +
    '<h3 style="margin-top:0">📋 Zuordnung prüfen</h3>' +
    '<p style="font-size:12px;color:#666;margin-bottom:10px">✅ = wird gespeichert | Grün = neu von KI | Gelb = überschreibt bestehend</p>' +
    '<table id="resultTable"><thead><tr><th>✅</th><th>Feld</th><th>Aktuell</th><th>Neuer Wert</th></tr></thead><tbody id="resultBody"></tbody></table>' +
    '<div style="margin-top:15px;text-align:right">' +
    '<button class="btn-gray" onclick="toggleAll(true)">Alle an</button>' +
    '<button class="btn-gray" onclick="toggleAll(false)">Alle aus</button>' +
    '<button class="btn-green" id="saveBtn" onclick="saveAssignments()">💾 In Kundentabelle speichern</button>' +
    '</div>' +
    '</div>' +

    '<script>' +
    'var currentAssignments=[];' +

    'function analyze(){' +
    '  var co=document.getElementById("company").value;' +
    '  var txt=document.getElementById("freetext").value.trim();' +
    '  if(!co){showSt("Bitte Kunde wählen","error");return}' +
    '  if(!txt||txt.length<10){showSt("Bitte Kundendaten eingeben (min. 10 Zeichen)","error");return}' +
    '  document.getElementById("analyzeBtn").disabled=true;' +
    '  showSt("⏳ Claude analysiert die Daten...","info");' +
    '  google.script.run.withSuccessHandler(function(r){' +
    '    currentAssignments=r.assignments;' +
    '    renderResults(r.assignments);' +
    '    document.getElementById("resultArea").style.display="block";' +
    '    showSt("✅ Analyse fertig! Bitte Zuordnung prüfen.","success");' +
    '    document.getElementById("analyzeBtn").disabled=false;' +
    '  }).withFailureHandler(function(e){' +
    '    showSt("❌ "+e,"error");' +
    '    document.getElementById("analyzeBtn").disabled=false;' +
    '  }).analyzeWithAI(txt,co);' +
    '}' +

    'function renderResults(items){' +
    '  var tb=document.getElementById("resultBody");tb.innerHTML="";' +
    '  items.forEach(function(item,idx){' +
    '    if(!item.aiValue&&!item.existingValue)return;' +
    '    var isNew=item.aiValue&&!item.existingValue;' +
    '    var isUpdate=item.aiValue&&item.existingValue&&item.aiValue!==item.existingValue;' +
    '    var cls=isNew?"ai-new":isUpdate?"ai-update":"";' +
    '    var tr=document.createElement("tr");tr.className=cls;' +
    '    tr.innerHTML=' +
    '      \'<td><input type="checkbox" id="chk_\'+idx+\'" \'+(item.useAI?"checked":"")+\' data-idx="\'+idx+\'"></td>\'+' +
    '      \'<td><strong>\'+item.column+\'</strong><br><span style="font-size:11px;color:#666">\'+item.description+\'</span></td>\'+' +
    '      \'<td><span class="existing">\'+(item.existingValue||"—")+\'</span></td>\'+' +
    '      \'<td><input type="text" id="val_\'+idx+\'" value="\'+escHtml(item.aiValue||item.existingValue||"")+\'"></td>\';' +
    '    tb.appendChild(tr);' +
    '  });' +
    '}' +

    'function saveAssignments(){' +
    '  var co=document.getElementById("company").value;' +
    '  if(!co){showSt("Kein Kunde gewählt","error");return}' +
    '  var toSave=[];' +
    '  currentAssignments.forEach(function(item,idx){' +
    '    var chk=document.getElementById("chk_"+idx);' +
    '    var val=document.getElementById("val_"+idx);' +
    '    if(chk&&chk.checked&&val&&val.value.trim()){' +
    '      toSave.push({column:item.column,value:val.value.trim()});' +
    '    }' +
    '  });' +
    '  if(!toSave.length){showSt("Keine Felder zum Speichern ausgewählt","error");return}' +
    '  document.getElementById("saveBtn").disabled=true;' +
    '  showSt("⏳ Speichere "+toSave.length+" Felder...","info");' +
    '  google.script.run.withSuccessHandler(function(r){' +
    '    showSt("✅ "+r.savedFields+" Felder für "+r.companyName+" gespeichert! (Zeile "+r.row+")","success");' +
    '    document.getElementById("saveBtn").disabled=false;' +
    '  }).withFailureHandler(function(e){' +
    '    showSt("❌ "+e,"error");' +
    '    document.getElementById("saveBtn").disabled=false;' +
    '  }).saveAIAssignments(co,toSave);' +
    '}' +

    'function toggleAll(state){' +
    '  currentAssignments.forEach(function(item,idx){' +
    '    var chk=document.getElementById("chk_"+idx);' +
    '    if(chk)chk.checked=state;' +
    '  });' +
    '}' +

    'function clearAll(){' +
    '  document.getElementById("freetext").value="";' +
    '  document.getElementById("resultArea").style.display="none";' +
    '  document.getElementById("status").style.display="none";' +
    '  currentAssignments=[];' +
    '}' +

    'function escHtml(s){return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}' +
    'function showSt(m,t){var el=document.getElementById("status");el.innerHTML=m;el.className=t;el.style.display="block"}' +
    '</script>'
  ).setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🤖 KI-Zuordnung');
}
