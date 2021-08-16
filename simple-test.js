const fs = require('fs/promises');

const mod = require('./index');

(async () => {
  const HEADER = [
    ['one', 1],
    ['two', 2],
    ['forty-four', '44'],
  ];

  const buf = await fs.readFile('example.xlsb');

  console.info('xlsb');
  const result = mod.parse('xlsb', buf, HEADER, 0);
  console.info(result.length);
  console.info(result[result.length - 1]);
})();
