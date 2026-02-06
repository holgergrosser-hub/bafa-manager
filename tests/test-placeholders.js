const fs = require('fs');
const path = require('path');
const assert = require('assert');

function extractConstObject(source, constName) {
  const re = new RegExp(`const\\s+${constName}\\s*=\\s*({[\\s\\S]*?^});`, 'm');
  const match = source.match(re);
  if (!match) throw new Error(`Could not find const ${constName}`);
  // Evaluate object literal safely in a function scope.
  // eslint-disable-next-line no-new-func
  return Function(`"use strict"; return (${match[1]});`)();
}

function run() {
  const codePath = path.join(__dirname, '..', 'src', 'Code.gs');
  const src = fs.readFileSync(codePath, 'utf8');

  const placeholderAlias = extractConstObject(src, 'PLACEHOLDER_ALIAS');
  const required = {
    companyName: '{{FIRMENNAME}}',
    Strasse: '{{Straße}}',
    PLZ_Ort: '{{PLZ_Ort}}',
    AnzahlderMitarbeiter: '{{AnzahderMitarbtier}}'
  };

  for (const [key, expected] of Object.entries(required)) {
    assert.strictEqual(
      placeholderAlias[key],
      expected,
      `PLACEHOLDER_ALIAS[${key}] expected ${expected} but got ${placeholderAlias[key]}`
    );
  }

  // Basic sanity: all placeholders are wrapped in {{ }}
  for (const [key, value] of Object.entries(placeholderAlias)) {
    assert.ok(typeof value === 'string', `PLACEHOLDER_ALIAS[${key}] must be string`);
    assert.ok(value.startsWith('{{') && value.endsWith('}}'), `PLACEHOLDER_ALIAS[${key}] must be wrapped in {{}}`);
  }

  console.log('✓ PLACEHOLDER_ALIAS tests passed');
}

module.exports = { run };
