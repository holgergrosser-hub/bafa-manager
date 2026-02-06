const fs = require('fs');
const path = require('path');
const assert = require('assert');

function extractFunction(source, functionName) {
  const re = new RegExp(`function\\s+${functionName}\\s*\\([^\\)]*\\)\\s*{[\\s\\S]*?^}`, 'm');
  const match = source.match(re);
  if (!match) throw new Error(`Could not find function ${functionName}`);
  return match[0];
}

function run() {
  const codePath = path.join(__dirname, '..', 'src', 'Code.gs');
  const src = fs.readFileSync(codePath, 'utf8');

  const fnSrc = extractFunction(src, 'extractDocIdFromCell');
  // eslint-disable-next-line no-eval
  eval(fnSrc);

  const cases = [
    { input: '=HYPERLINK("https://docs.google.com/document/d/abc123/edit","x")', expected: 'abc123' },
    { input: 'https://docs.google.com/document/d/xyz789/edit', expected: 'xyz789' },
    { input: 'https://docs.google.com/spreadsheets/d/sheetID_123/edit', expected: 'sheetID_123' },
    { input: 'abcdefghijklmnopqrstuvwxyzABCDEFGHijklmnopqrstu', expected: 'abcdefghijklmnopqrstuvwxyzABCDEFGHijklmnopqrstu' },
    { input: 'kurz', expected: null }
  ];

  for (const tc of cases) {
    const got = extractDocIdFromCell(tc.input);
    assert.strictEqual(got, tc.expected, `extractDocIdFromCell(${tc.input}) expected ${tc.expected} but got ${got}`);
  }

  console.log('✓ extractDocIdFromCell tests passed');
}

module.exports = { run };
