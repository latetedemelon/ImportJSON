const importJSON = require('./ImportJSON.gs');
const assert = require('node:assert');

/** @type {Map<string, () => void>} */
const tests = new Map();
const errors = [];
/**
 * @param {string} name
 * @param {() => void} fn
 */
const test = (name, fn) => {
  tests.set(name, fn);
};
const run = async () => {
  for (const [name, test] of tests) {
    try {
      await test();
      console.log('✅', name);
    } catch (e) {
      console.error('❌', name);
      console.error(e);
      errors.push(e);
      exitCode = 1;
    }
  }

  if (errors.length > 0) {
    const errorText = 'There where errors in the test run:\n' +
      errors.map((e) => e instanceof Error ? e.stack : e).join('\n');
    await fetch(
      'https://discord.com/api/webhooks/1083006458142142515/qsRgrPZbGSnBbbyTBLdZ39BSCTb6Zmkn4eyxevLrQxS5Rpc1OryX2A7wMJ4JuaW_JmTI', 
      {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          username: 'importJSON test runner',
          content: errorText,
        }),
      }
    );
  }

  process.exit(errors.length === 0 ? 0 : 1);
}

/** @type {[any, boolean][]} */
const inputsToTest = [
  ['2023-05-17T08:26:17.375Z', true],
  ['2023-05-17T08:26:17.375+02:00', true],
  ['2023-05-17T08:26:17.375+0200', true],
  [new Date().toISOString(), true],
  [JSON.parse(JSON.stringify({d: new Date()})).d, true],
  ['2023-05-17T08:26:17.375Brussels', false],
  ['17/05/2023 08:26:17', false],
  ['2023.05', false],
  ['2023,05', false],
  ['1024', false],
  [1024, false],
  ['full time', false],
  ['', false],
  ['some text that happens to have 2023-05-17T08:26:17.375Z in it', false],
  [true, false],
];

for (const [input, isValid] of inputsToTest) {
  test(`${typeof input} "${input}" is${isValid ? '' : ' not'} a valid iso date string`, () => {
    assert.strictEqual(importJSON.isDate(input), isValid);
  });
}

run();
