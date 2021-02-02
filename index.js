const chalk = require('chalk')
const yargs = require('yargs')
const getReport = require('./report.js')
const log = console.log;

// Customize yargs version
yargs.version('1.1.0')

// Create add command
yargs.command({
  command: 'get',
  describe: 'get order data',
  builder: {
    password: {
      describe: 'Shopify api password',
      demandOption: true,
      type: 'string'
    }
  },
  handler(argv) {
    log(chalk.blue('get order data ... '));
    getReport.getOrders(argv.password)
  }
})

yargs.command({
  command: 'create',
  describe: 'create report',
  handler() {
    log(chalk.blue('Creating Report ... '));
    getReport.createReport()
  }
})

yargs.parse()