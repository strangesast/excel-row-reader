import test from 'ava';

import { parse } from '../index';

test('parse', (t) => {
  t.assert(typeof parse === 'function');
});
