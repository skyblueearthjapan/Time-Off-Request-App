// すべての src/*.gs が構文として正しいこと、固有の文字列が他アプリのものでないこと。node --test gas-tests/
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { SRC } = require('./harness');

const files = fs.readdirSync(SRC).filter(f => f.endsWith('.gs')).sort();

test('src/*.gs がすべて構文として正しい', () => {
  assert.ok(files.length >= 15);
  for (const f of files) {
    assert.doesNotThrow(() => new vm.Script(fs.readFileSync(path.join(SRC, f), 'utf8'), { filename: f }), f);
  }
});

test('新しい 2 ファイルに残業・貯蔵品の固有の文字列が無い', () => {
  for (const f of ['Internal.gs', 'Cutover.gs']) {
    const t = fs.readFileSync(path.join(SRC, f), 'utf8');
    for (const bad of ['overtime', 'supplies', 'Sebastian', 'SEBASTIAN', '残業', '休日出勤', '貯蔵品', '1Knx_kaQ', 'ACCUMULATION', 'FormApp', 'morningBatch']) {
      assert.ok(!t.includes(bad), f + ' contains ' + bad);
    }
  }
});

test('appsscript.json の webapp は DOMAIN・USER_DEPLOYING のまま', () => {
  const m = JSON.parse(fs.readFileSync(path.join(SRC, 'appsscript.json'), 'utf8'));
  assert.deepEqual(m.webapp, { executeAs: 'USER_DEPLOYING', access: 'DOMAIN' });
});
