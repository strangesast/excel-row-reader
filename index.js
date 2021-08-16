const { loadBinding } = require('@node-rs/helper');

const { dump, ...mod } = loadBinding(__dirname, 'rusty-excel-reader', 'rusty-excel-reader');
const parse = (s, b, headers, sheet_name) => {
  const buf = dump(s, b, headers, sheet_name);
  return JSON.parse(buf.toString());
};

module.exports = { ...mod, dump, parse };
