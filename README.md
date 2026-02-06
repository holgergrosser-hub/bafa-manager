# BAFA Manager (Google Apps Script)

Dieses Repo enthält den Quellcode für das gemeinsame Google Apps Script Projekt (clasp).

## Voraussetzungen

- Node.js (inkl. npm)
- `git`
- `clasp` (Google Apps Script CLI)

Installation:

```bash
npm install -g @google/clasp
```

## Setup (erstes Mal)

1. Repo klonen
2. Einmalig bei Google anmelden:

```bash
clasp login
```

3. Status prüfen:

```bash
clasp status
```

## Standard-Workflow

- Änderungen aus der Cloud holen:

```bash
clasp pull
```

- Lokale Änderungen in die Cloud pushen:

```bash
clasp push
```

Hinweis: In Apps Script können Dateinamen-Endungen unterschiedlich sein (`.gs`, `.js`, `.html`). In der Cloud gibt es pro Datei nur **einen** Namen ohne Endung – vermeide daher z.B. gleichzeitig `Code.js` und `Code.gs`.

## Projekt öffnen

Die Projekt-ID steht in `.clasp.json` (`scriptId`). Öffnen im Browser:

`https://script.google.com/d/<scriptId>/edit`

Oder per Skript (Windows PowerShell):

```powershell
.\scripts\open.ps1
```

## Was wird gepusht?

- `clasp` pusht nur getrackte Dateien (i.d.R. `appsscript.json`, `*.gs/*.js`, `*.html`).
- Regeln, was *nicht* gepusht wird, stehen in `.claspignore`.

## Team-Hinweis (gleiches Script für alle)

`.clasp.json` ist **Teil des Repos**, damit alle Entwickler gegen dasselbe Apps-Script-Projekt arbeiten.

## Lint & Tests

```bash
npm ci
npm run lint
npm test
```
