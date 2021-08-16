import * as fs from 'fs';

import b from 'benny';

import { parse } from '../index';

const BUF = fs.readFileSync('./example.xlsb');
const HEADERS = [
  ['one', 1],
  ['two', 2],
  ['forty-four', '44'],
] as [string, string][];

async function run() {
  await b.suite(
    'parse',

    b.add('xlsb', () => {
      parse('xlsb', BUF, HEADERS, 0);
    }),

    b.cycle(),
    b.complete(),
  );
}

run().catch((e) => {
  console.error(e);
});
