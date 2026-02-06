# BAFA Manager – Entwicklungskonzept

> **Stand:** 06.02.2026 | **Code:** 1.635 Zeilen (Code.gs) + 859 Zeilen (index.html) | **63 Funktionen**

---

## 1. System-Übersicht

### Was ist der BAFA Manager?

Ein Google Apps Script Tool zur automatisierten Erstellung von ISO-Beratungsdokumenten für BAFA-geförderte Projekte. Von der Kundendaten-Erfassung bis zum fertigen PDF-Ordner – alles in einem System.

### Architektur

```
┌──────────────────────────────────────────────────────────────────────┐
│                        BAFA MANAGER SYSTEM                           │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  DATENQUELLEN                          AUSGABE                       │
│  ┌──────────────┐                      ┌──────────────────────────┐  │
│  │ CRM Super    │──┐                   │ 14 BAFA-Dokumente        │  │
│  │ Master Sheet │  │                   │ (Google Docs + 1 Sheet)  │  │
│  └──────────────┘  │  ┌────────────┐   │ + Logo eingefügt         │  │
│  ┌──────────────┐  ├─▶│ Kunden-    │──▶│ + Platzhalter befüllt    │  │
│  │ Kunden-E-Mail│  │  │ tabelle    │   │ + PDF-Export             │  │
│  │ (Freitext)   │──┘  │ (32 Sp.)  │   └──────────────────────────┘  │
│  └──────────────┘     └────────────┘                                 │
│  ┌──────────────┐          │                                         │
│  │ Claude API   │──────────┘  KI-Zuordnung                          │
│  └──────────────┘                                                    │
│                                                                      │
│  SPEICHER                                                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐               │
│  │ Template-    │  │ Kunden-      │  │ Logo-        │               │
│  │ Ordner       │  │ Ordner       │  │ Ordner       │               │
│  │ (14 Vorlagen)│  │ (pro Kunde)  │  │ (alle Logos) │               │
│  └──────────────┘  └──────────────┘  └──────────────┘               │
└──────────────────────────────────────────────────────────────────────┘
```

### Google Drive IDs

| Ressource | ID |
|---|---|
| Template-Ordner | `1S_KMGYI500ACGZR3y37O3LPpVw--V0J5` |
| Parent-Ordner (Kundenordner) | `1CxrypRxuDROK_QWJje397LmrzvvIheGJ` |
| Logo-Ordner | `1X0yHW8IwoacCj9YT1wYVLzLF0ZcmnI0B` |
| CRM Super Master Sheet | `1FWbeX3YeK9Uidyn9obKJ7z-J-zXX1h5PsXcfk_YHAyU` |

---

## 2. Workflow (5 Schritte)

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         BAFA WORKFLOW (5 SCHRITTE)                       │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  ① CRM-IMPORT                                            ✅ FERTIG     │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────────────────┐  │
│  │ Super Master │ ── │ Firma wählen │ ── │ Stammdaten in Kunden-    │  │
│  │ CRM Sheet    │    │ (Suchfilter) │    │ tabelle (neu/aktualisiert│  │
│  └──────────────┘    └──────────────┘    └──────────────────────────┘  │
│                                                     │                   │
│  ② KI-ZUORDNUNG                                     │    ✅ FERTIG     │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────▼───────────────┐  │
│  │ Kundendaten  │ ── │ Claude API   │ ── │ Prüfen (Tabelle mit     │  │
│  │ reinkopieren │    │ analysiert + │    │ ✅/❌ pro Feld) →        │  │
│  │ (Freitext)   │    │ ordnet zu    │    │ Speichern in Sheet       │  │
│  └──────────────┘    └──────────────┘    └──────────────────────────┘  │
│                                                     │                   │
│  ③ TEMPLATE KOPIEREN                                │    ✅ FERTIG     │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────▼───────────────┐  │
│  │ Kunde wählen │ ── │ 14 Templates │ ── │ Doc-IDs automatisch      │  │
│  │ + Datum      │    │ kopieren     │    │ in Tabelle geschrieben   │  │
│  └──────────────┘    └──────────────┘    └──────────────────────────┘  │
│                                                     │                   │
│  ④ PLATZHALTER BEFÜLLEN                             │    ✅ FERTIG     │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────▼───────────────┐  │
│  │ Vorschau:    │ ── │ Bestätigen   │ ── │ {{Platzhalter}} in allen │  │
│  │ gefüllt/leer │    │              │    │ 13 Docs ersetzt          │  │
│  └──────────────┘    └──────────────┘    └──────────────────────────┘  │
│                                                     │                   │
│  ⑤ LOGO + PDF                                       │    ✅ FERTIG     │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────▼───────────────┐  │
│  │ Logo upload  │ ── │ Auto-Insert  │ ── │ PDF-Export +             │  │
│  │ pro Kunde    │    │ in alle Docs │    │ PDF-Split                │  │
│  └──────────────┘    └──────────────┘    └──────────────────────────┘  │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 3. Kundentabelle (32 Spalten)

### Spalten-Gruppen

| Gruppe | Spalten | Beschreibung |
|---|---|---|
| **CRM** (4) | companyName, Strasse, PLZ_Ort, Ansprechpartner | Import aus Super Master |
| **Platzhalter** (11) | email, Webpage, Gruenderdatum, AnzahlderMitarbeiter, Geltungsbereich, Zielgruppe_Zielgebiet, Ausgelagerte_Prozesse, Norm, AUDITOR, Aktuelles_Jahr, Datum_Heute | Werden in Dokumente geschrieben |
| **System** (4) | folderID, folderLink, createdDate, logoUrl | Automatisch befüllt |
| **Dokumente** (14) | doc_01 bis doc_14 | HYPERLINK-Formeln zu Google Docs |

### CRM-Import Mapping

```
Super Master Spalte  →  BAFA Kundentabelle
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
A  (Firmenname)      →  companyName
BS (Straße)          →  Strasse
BW (PLZ + Ort)       →  PLZ_Ort
CA (Ansprechpartner) →  Ansprechpartner
I  (E-Mail)          →  email

Konfigurierbar in: CRM_IMPORT_MAP (Code.gs Zeile ~15)
```

### 14 Dokumente

| Nr. | Spalte | Typ | Hinweis |
|---|---|---|---|
| 01 | doc_01_Beraterbewertung | Google Doc | |
| 02 | doc_02_Kundenrueckmeldung | Google Doc | |
| 03 | doc_03_Normen_Gesetze | Google Doc | |
| 04 | doc_04_Managementbewertung | Google Doc | |
| 05 | doc_05_Massnahmenplan | **Google Sheet** | ⚠️ Kein Logo, keine Platzhalter |
| 06 | doc_06_Prozessbeschreibungen | Google Doc | Nicht im Template, wird vorab erstellt |
| 07 | doc_07_Schulungsplan | Google Doc | |
| 08 | doc_08_Ziele_Prozesskennzahlen | Google Doc | |
| 09 | doc_09_Unternehmenshandbuch | Google Doc | |
| 10 | doc_10_Auditbericht | Google Doc | |
| 11 | doc_11_Vollmacht | Google Doc | |
| 12 | doc_12_Firmeninfo_Foerdergeld | Google Doc | |
| 13 | doc_13_Projektbericht | Google Doc | |
| 14 | doc_14_Ausfuellanleitung | Google Doc | |

### Platzhalter-Mapping (PLACEHOLDER_ALIAS)

```
Spalte                    →  Platzhalter im Dokument
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
companyName               →  {{FIRMENNAME}}
Strasse                   →  {{Straße}}
PLZ_Ort                   →  {{PLZ_Ort}}
Ansprechpartner           →  {{Ansprechpartner}}
email                     →  {{email}}
Webpage                   →  {{Webpage}}
Gruenderdatum             →  {{Gruenderdatum}}
AnzahlderMitarbeiter      →  {{AnzahderMitarbtier}}    ⚠️ Tippfehler im Template!
Geltungsbereich           →  {{Geltungsbereich}}
Zielgruppe_Zielgebiet     →  {{Zielgruppe/Zielgebiet}}
Ausgelagerte_Prozesse     →  {{Ausgelagerte Prozesse}}
Norm                      →  {{Norm}}
AUDITOR                   →  {{AUDITOR}}
Aktuelles_Jahr            →  {{Aktuelles_Jahr}}
Datum_Heute               →  {{Datum_Heute}}
```

---

## 4. Menü-Struktur

```
📋 BAFA Manager
├── 📊 Manager öffnen           → Web App (10 Tabs)
├── ─────────────
├── 🆕 Kundentabelle erstellen  → Neues Sheet mit 32 Spalten
├── 📋 Kundentabelle öffnen     → Direktlink
├── ─────────────
├── 📥 CRM-Import               → Dialog: Super Master → Kundentabelle
├── 🤖 KI-Zuordnung             → Dialog: Freitext → Claude → Platzhalter
├── ─────────────
├── 📁 Template kopieren        → Dialog: 14 Dokumente erstellen
├── 📝 Platzhalter befüllen     → Dialog: Vorschau → Bestätigen → Ersetzen
├── 🖼️ Logo hochladen          → Dialog: Kunde → Logo → In alle Docs
├── 📑 PDFs erstellen           → Dialog: Ordner → PDF-Konvertierung
├── ✂️ PDFs aufteilen           → Dialog: PDF-URLs → Einzelseiten
├── ─────────────
└── 🔑 Claude Key einrichten    → API-Key für KI-Zuordnung
```

### Web App (10 Tabs)

| Tab | Funktion |
|---|---|
| 🏠 Übersicht | Workflow-Grafik (klickbar), letzte Kundenordner |
| 📥 CRM-Import | Suchfilter, Firmenliste, Import-Button |
| 🤖 KI-Zuordnung | Freitext → Claude → Prüf-Tabelle → Speichern |
| 📁 Template kopieren | Kunde + Datum → 14 Dokumente erstellen |
| 📝 Platzhalter | Vorschau → Bestätigen → Alle Docs befüllen |
| 🖼️ Logo | Kunde → Info-Anzeige → Upload → Auto-Insert alle Docs |
| ⚡ Stapel | Kopieren + Befüllen in einem Schritt |
| ✏️ Umbenennen | Dateien im Ordner umbenennen |
| 📑 PDF | Google Docs/Sheets → PDF konvertieren |
| ✂️ PDF teilen | Mehrseitige PDFs → Einzelseiten |

---

## 5. KI-Zuordnung im Detail

### Flow

```
┌─────────────┐    ┌──────────────────┐    ┌───────────────────────┐
│ 1. Kunde    │    │ 2. Freitext      │    │ 3. Claude API         │
│    wählen   │───▶│    reinkopieren   │───▶│    analysiert         │
│ (Dropdown)  │    │ (E-Mail, PDF...) │    │    (sonnet-4)         │
└─────────────┘    └──────────────────┘    └───────────┬───────────┘
                                                       │
┌─────────────┐    ┌──────────────────┐    ┌───────────▼───────────┐
│ 6. In       │    │ 5. Werte        │    │ 4. Ergebnis-Tabelle   │
│ Kundentab.  │◀───│    editieren    │◀───│    ✅ Feld │ Wert     │
│ speichern   │    │    per Feld     │    │    Grün=neu Gelb=upd. │
└─────────────┘    └──────────────────┘    └───────────────────────┘
```

### Was Claude bekommt

- Alle 15 Platzhalter-Felder mit Beschreibung
- Bereits vorhandene Werte aus der Kundentabelle `[AKTUELL: ...]`
- Den Freitext vom Benutzer
- Regeln: Nur zuordnen was im Text steht, bestehende Werte nur bei besserem Match überschreiben

### Ergebnis-Tabelle (UI)

```
┌────┬────────────────────┬──────────────────┬──────────────────────┐
│ ✅ │ Feld               │ Aktuell          │ Neuer Wert (editbar) │
├────┼────────────────────┼──────────────────┼──────────────────────┤
│ ☑  │ companyName        │ —                │ Müller GmbH      🟢 │
│ ☑  │ Strasse            │ —                │ Hauptstr. 5      🟢 │
│ ☑  │ PLZ_Ort            │ —                │ 90402 Nürnberg   🟢 │
│ ☑  │ email              │ alt@firma.de     │ neu@firma.de     🟡 │
│ ☐  │ Norm               │ ISO 9001:2015    │ ISO 9001:2015       │
└────┴────────────────────┴──────────────────┴──────────────────────┘
🟢 = Neu von KI    🟡 = Überschreibt bestehend    ☐ = Nicht ausgewählt
```

### Backend-Funktionen

| Funktion | Beschreibung |
|---|---|
| `analyzeWithAI(freeText, companyName)` | Sendet Freitext + Feld-Definitionen an Claude, gibt Zuordnungs-Array zurück |
| `saveAIAssignments(companyName, confirmedData)` | Schreibt geprüfte Zuordnungen in Kundentabelle |
| `setupClaudeApiKey()` | Speichert Claude API-Key in Script Properties |

### Voraussetzung

Claude API-Key muss einmalig über Menü **🔑 Claude Key einrichten** gesetzt werden (Format: `sk-ant-...`).

---

## 6. Logo-Workflow

```
Kunde wählen → Kundeninfo anzeigen → Logo hochladen → In ALLE Dokumente einfügen
     │                  │                    │                     │
     ▼                  ▼                    ▼                     ▼
 Dropdown aus      • Ordner vorhanden?   Logo in Logo-        Jedes doc_01-14:
 Kundentabelle     • Anzahl Dokumente    Ordner speichern     • Doc öffnen
                   • Logo vorhanden?     logoUrl in Tabelle   • {{LOGO_URL}} suchen
                   (wird überschrieben)                       • Durch Bild ersetzen
                                                              • Max 200x100px
                                                              • Sheet überspringen
```

---

## 7. Entwicklungsumgebung: VS Code + GitHub + clasp

### Warum?

| Problem heute | Lösung |
|---|---|
| Code direkt im Apps Script Editor bearbeiten | VS Code mit Syntax-Highlighting, Autocomplete, Linting |
| Keine Versionierung | Git/GitHub – jede Änderung nachvollziehbar |
| Kein Review vor Deployment | Pull Requests, Code Review |
| Copy/Paste zwischen Claude und Apps Script | `clasp push` direkt aus VS Code |
| Fehler erst beim Testen im Browser | ESLint + lokale Tests vor dem Push |

### Entwicklungsflow

```
┌─────────────────────────────────────────────────────────────────┐
│                      ENTWICKLUNGSFLOW                            │
│                                                                  │
│  VS Code (lokal)                                                 │
│  ├── src/Code.gs            ← 1.635 Zeilen Backend              │
│  ├── src/index.html         ← 859 Zeilen Frontend (10 Tabs)     │
│  ├── src/CompanySelection.html  ← CRM-Auswahl Dialog            │
│  ├── src/appsscript.json    ← Projekt-Konfiguration             │
│  ├── .clasp.json            ← Script-ID Verknüpfung             │
│  ├── .eslintrc.json         ← Linting-Regeln                    │
│  └── tests/                 ← Lokale Tests                      │
│         │                                                        │
│         ▼                                                        │
│  ┌─────────────┐    ┌──────────────┐                             │
│  │ git commit  │ ── │ GitHub Push  │                             │
│  │ git push    │    │ (Repository) │                             │
│  └─────────────┘    └──────┬───────┘                             │
│                            │                                     │
│                     GitHub Actions (CI/CD)                        │
│                     ┌──────▼───────┐                             │
│                     │ 1. ESLint    │                             │
│                     │ 2. Tests     │                             │
│                     │ 3. clasp push│ ── Google Apps Script       │
│                     │    (deploy)  │    Projekt                  │
│                     └──────────────┘                             │
└─────────────────────────────────────────────────────────────────┘
```

### Setup-Schritte (Schritt für Schritt)

```bash
# 1. Node.js + clasp installieren
npm install -g @google/clasp

# 2. Bei Google anmelden (öffnet Browser)
clasp login

# 3. Script-ID aus dem Apps Script Editor holen:
#    Apps Script öffnen → Einstellungen → Script-ID kopieren

# 4. Projektordner erstellen
mkdir bafa-manager
cd bafa-manager

# 5. Bestehendes Projekt klonen
clasp clone <DEINE_SCRIPT_ID>
#    → erstellt .clasp.json + alle Dateien

# 6. In src/ Ordner verschieben (Struktur wie unten)
mkdir src
mv Code.gs index.html appsscript.json src/

# 7. .clasp.json anpassen (rootDir auf src):
# { "scriptId": "DEINE_ID", "rootDir": "src" }

# 8. Git Repository initialisieren
git init
git remote add origin https://github.com/holger-grosser/bafa-manager.git

# 9. .gitignore erstellen
echo ".clasp.json" > .gitignore
echo "node_modules/" >> .gitignore

# 10. Initialer Commit
git add -A
git commit -m "v1.0 - BAFA Manager mit CRM-Import + KI-Zuordnung"
git push -u origin main

# 11. Code nach Google Apps Script pushen
clasp push

# 12. Version erstellen
clasp deploy --description "v1.0"
```

### Projektstruktur

```
bafa-manager/
├── src/                           # → wird nach Google Apps Script gepusht
│   ├── Code.gs                    #   Backend (1.635 Zeilen, 63 Funktionen)
│   ├── index.html                 #   Frontend (859 Zeilen, 10 Tabs)
│   ├── CompanySelection.html      #   CRM-Auswahl (separater Dialog)
│   └── appsscript.json            #   Apps Script Manifest
├── tests/
│   ├── test-placeholders.js       #   Platzhalter-Tests
│   ├── test-crm-import.js         #   CRM-Import Tests
│   ├── test-doc-ids.js            #   Dokument-ID Extraktion Tests
│   └── run-all.js                 #   Test-Runner
├── .clasp.json                    # ⚠️ NICHT committen (enthält Script-ID)
├── .claspignore                   #   Dateien die nicht gepusht werden
├── .eslintrc.json                 #   Linting-Regeln
├── .gitignore                     #   .clasp.json, node_modules
├── .github/
│   └── workflows/
│       └── deploy.yml             #   CI/CD Pipeline
├── package.json                   #   Node Dependencies
└── README.md                      #   Dokumentation
```

### .claspignore

```
tests/**
.github/**
.eslintrc.json
.gitignore
package.json
package-lock.json
node_modules/**
README.md
BAFA_Entwicklungskonzept.md
```

---

## 8. Qualitätssicherung

### 3-Stufen-QA-Modell

```
┌──────────────────────────────────────────────────────────────────┐
│                    QUALITÄTSSICHERUNG                              │
│                                                                    │
│  STUFE 1: Prävention (vor dem Code)                               │
│  ┌────────────────────────────────────────────────────────────┐   │
│  │ • ESLint: Syntaxfehler, Style, Best Practices              │   │
│  │ • TypeScript JSDoc: Typen dokumentieren                    │   │
│  │ • Claude: Code-Review vor Übergabe                         │   │
│  └────────────────────────────────────────────────────────────┘   │
│                            │                                       │
│  STUFE 2: Verifikation (automatische Tests)                       │
│  ┌────────────────────────────────────────────────────────────┐   │
│  │ • Unit Tests: Einzelne Funktionen isoliert                 │   │
│  │ • Integration Tests: Zusammenspiel prüfen                  │   │
│  │ • Mock Tests: Google Services simulieren                   │   │
│  │ • Regression Tests: Alte Fehler nicht wiederholen           │   │
│  └────────────────────────────────────────────────────────────┘   │
│                            │                                       │
│  STUFE 3: Absicherung (Deployment)                                │
│  ┌────────────────────────────────────────────────────────────┐   │
│  │ • GitHub Actions: Tests vor jedem Push                     │   │
│  │ • Staging-Deployment: Erst testen, dann live               │   │
│  │ • Rollback: Jederzeit zur letzten Version zurück           │   │
│  │ • Monitoring: Logger + Fehler-Benachrichtigung             │   │
│  └────────────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────┘
```

### Stufe 1: ESLint Konfiguration

```json
// .eslintrc.json
{
  "env": { "es6": true },
  "globals": {
    "SpreadsheetApp": "readonly",
    "DriveApp": "readonly",
    "DocumentApp": "readonly",
    "HtmlService": "readonly",
    "UrlFetchApp": "readonly",
    "PropertiesService": "readonly",
    "Utilities": "readonly",
    "Logger": "readonly",
    "Session": "readonly",
    "MimeType": "readonly",
    "Drive": "readonly",
    "CONFIG": "readonly",
    "TABLE_COLUMNS": "readonly",
    "DOC_COLUMN_MAP": "readonly",
    "PLACEHOLDER_ALIAS": "readonly",
    "CRM_IMPORT_MAP": "readonly",
    "SKIP_TEXT_REPLACE": "readonly"
  },
  "rules": {
    "no-unused-vars": "warn",
    "no-undef": "error",
    "eqeqeq": "error",
    "no-var": "warn",
    "prefer-const": "warn"
  }
}
```

### Stufe 2: Tests

```javascript
// tests/test-placeholders.js

// Mock für Google Services
const mockSheet = {
  getDataRange: () => ({
    getValues: () => [
      ['companyName','Strasse','PLZ_Ort','Ansprechpartner','email'],
      ['Müller GmbH','Hauptstr. 5','90402 Nürnberg','Hr. Müller','info@mueller.de']
    ]
  }),
  getRange: () => ({ getFormula: () => '', getValue: () => '' })
};

// Test: Platzhalter-Mapping
function testPlaceholderAlias() {
  const expected = {
    'companyName': '{{FIRMENNAME}}',
    'Strasse': '{{Straße}}',
    'AnzahlderMitarbeiter': '{{AnzahderMitarbtier}}'
  };
  for (const [col, placeholder] of Object.entries(expected)) {
    console.assert(
      PLACEHOLDER_ALIAS[col] === placeholder,
      `FAIL: ${col} → erwartet "${placeholder}", bekommen "${PLACEHOLDER_ALIAS[col]}"`
    );
  }
  console.log('✓ Platzhalter-Alias Test bestanden');
}

// Test: Dokument-ID Extraktion
function testExtractDocId() {
  const testCases = [
    { input: '=HYPERLINK("https://docs.google.com/document/d/abc123/edit","📄")', expected: 'abc123' },
    { input: 'https://docs.google.com/document/d/xyz789/edit', expected: 'xyz789' },
    { input: 'abcdefghijklmnopqrstuvwxyz', expected: 'abcdefghijklmnopqrstuvwxyz' },
    { input: 'kurz', expected: null }
  ];
  testCases.forEach(function(tc) {
    const result = extractDocIdFromCell(tc.input);
    console.assert(
      result === tc.expected,
      `FAIL: "${tc.input}" → erwartet "${tc.expected}", bekommen "${result}"`
    );
  });
  console.log('✓ DocId-Extraktion Test bestanden');
}

// Test: CRM Import Map Vollständigkeit
function testCrmImportMap() {
  const requiredColumns = ['companyName', 'Strasse', 'PLZ_Ort', 'Ansprechpartner', 'email'];
  const mappedColumns = Object.values(CRM_IMPORT_MAP);
  requiredColumns.forEach(function(col) {
    console.assert(
      mappedColumns.includes(col),
      `FAIL: CRM_IMPORT_MAP fehlt Mapping für "${col}"`
    );
  });
  console.log('✓ CRM Import Map Test bestanden');
}
```

### Stufe 3: GitHub Actions CI/CD

```yaml
# .github/workflows/deploy.yml
name: Lint, Test & Deploy

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  quality:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: '20'
      - run: npm install
      - name: ESLint
        run: npx eslint src/Code.gs
      - name: Tests
        run: node tests/run-all.js

  deploy:
    needs: quality
    if: github.ref == 'refs/heads/main' && github.event_name == 'push'
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: '20'
      - run: npm install -g @google/clasp
      - name: clasp push
        env:
          CLASP_TOKEN: ${{ secrets.CLASP_TOKEN }}
        run: |
          echo $CLASP_TOKEN > ~/.clasprc.json
          clasp push --force
```

---

## 9. Alle Funktionen (Referenz)

### Backend (Code.gs – 63 Funktionen)

| Bereich | Funktion | Beschreibung |
|---|---|---|
| **CRM** | `getCrmCompanies()` | Firmenliste aus Super Master |
| | `importFromCrm(companyName)` | Stammdaten in Kundentabelle importieren |
| **KI** | `analyzeWithAI(freeText, companyName)` | Freitext → Claude → Zuordnungs-Array |
| | `saveAIAssignments(companyName, confirmedData)` | Geprüfte Zuordnung speichern |
| | `setupClaudeApiKey()` | Claude API-Key setzen |
| **Tabelle** | `createCustomerSheet()` | Kundentabelle mit 32 Spalten erstellen |
| | `getCustomerSheet()` | Kundentabelle-Referenz holen |
| | `getCompanyNames()` | Firmenliste aus Kundentabelle |
| **Template** | `copyMasterFolderWithRename(customer, date, archive)` | 14 Templates kopieren + IDs in Tabelle |
| **Platzhalter** | `fillPlaceholders(companyName)` | Alle {{...}} in allen Docs ersetzen |
| | `getPlaceholderPreview(companyName)` | Vorschau: gefüllt/leer |
| **Logo** | `uploadLogoAndProcessDocuments(blob, company)` | Logo speichern + in alle Docs einfügen |
| | `getCustomerInfo(companyName)` | Ordner, Docs, Logo-Status |
| | `insertLogoInDocument(docId, imageUrl)` | {{LOGO_URL}} → Bild ersetzen |
| **PDF** | `convertToPDF(folderId, createDateFolder)` | Google Docs → PDF |
| | `splitPDFWithPdfCo(fileUrl)` | PDF → Einzelseiten (pdf.co) |
| **Stapel** | `batchProcess(customer, date, archive)` | Kopieren + Befüllen in einem Schritt |
| **Menü** | `onOpen()` | Menü mit 12 Einträgen erstellen |

---

## 10. Nächste Schritte

### Umsetzungsreihenfolge

| Phase | Was | Aufwand | Status |
|---|---|---|---|
| ~~Phase 1~~ | ~~CRM-Import einbauen~~ | ~~1 Std~~ | ✅ FERTIG |
| ~~Phase 2~~ | ~~KI-Zuordnungs-Dialog einbauen~~ | ~~2 Std~~ | ✅ FERTIG |
| **Phase 3** | **VS Code + clasp Setup** | **30 Min** | ⭐⭐⭐ JETZT |
| **Phase 4** | **GitHub Repo + initialer Commit** | **30 Min** | ⭐⭐⭐ JETZT |
| **Phase 5** | **ESLint Konfiguration** | **15 Min** | ⭐⭐⭐ JETZT |
| Phase 6 | .claspignore + .gitignore | 10 Min | ⭐⭐ Diese Woche |
| Phase 7 | Unit Tests schreiben | 2 Std | ⭐⭐ Diese Woche |
| Phase 8 | GitHub Actions CI/CD | 1 Std | ⭐⭐ Nächste Woche |
| Phase 9 | Staging/Prod Trennung | 1 Std | ⭐ Später |

### Ergebnis nach Phase 3-5 (heute)

- ✅ Kein Copy/Paste mehr zwischen Claude und Apps Script Editor
- ✅ `clasp push` direkt aus VS Code Terminal
- ✅ Jede Änderung versioniert in Git
- ✅ ESLint fängt Syntaxfehler ab bevor sie deployt werden
- ✅ Rollback jederzeit: `git checkout <commit>` + `clasp push`
