const fs = require('fs');
const path = require('path');
const assert = require('assert');

function extractConstObject(source, constName) {
  const re = new RegExp(`const\\s+${constName}\\s*=\\s*({[\\s\\S]*?^});`, 'm');
  const match = source.match(re);
  if (!match) throw new Error(`Could not find const ${constName}`);
  // eslint-disable-next-line no-new-func
  return Function(`"use strict"; return (${match[1]});`)();
}

function run() {
  const codePath = path.join(__dirname, '..', 'src', 'Code.gs');
  const src = fs.readFileSync(codePath, 'utf8');

  const crmMap = extractConstObject(src, 'CRM_IMPORT_MAP');
  const mustMapTo = ['companyName', 'Strasse', 'PLZ_Ort', 'Ansprechpartner', 'email'];
  const values = new Set(Object.values(crmMap));

  for (const col of mustMapTo) {
    assert.ok(values.has(col), `CRM_IMPORT_MAP must include mapping to ${col}`);
  }

  console.log('✓ CRM_IMPORT_MAP tests passed');
}

module.exports = { run };
