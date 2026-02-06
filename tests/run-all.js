const tests = [
  require('./test-placeholders'),
  require('./test-crm-import-map'),
  require('./test-extract-doc-id')
];

for (const t of tests) {
  t.run();
}

console.log('✓ all tests passed');
