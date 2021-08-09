const { loadBinding } = require('@node-rs/helper')

module.exports = loadBinding(__dirname, 'rusty-excel-reader', 'rusty-excel-reader')
