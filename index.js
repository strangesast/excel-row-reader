const { loadBinding } = require('@node-rs/helper')

module.exports = loadBinding(__dirname, 'excel-row-reader', 'excel-row-reader')
