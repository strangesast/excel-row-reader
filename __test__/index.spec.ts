import * as fs from 'fs';

import test from 'ava';

import { parse, dump, Fmt } from '../index';

const EXAMPLE_HEADERS = [
  ['one', 1],
  ['two', 2],
  ['forty-four', '44'],
] as [string, string][];

test('parse', (t) => {
  t.assert(typeof parse === 'function');
  for (const fmt of ['xlsb', 'xls', 'xlsm', 'xlsx'] as Fmt[]) {
    const filename = `example.${fmt}`;
    const buf = fs.readFileSync(filename);
    t.assert(parse(fmt, buf, EXAMPLE_HEADERS, 0).length > 0);
  }
});

test('dump', (t) => {
  t.assert(typeof dump === 'function');
});
