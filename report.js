const fs = require('fs')
const chalk = require('chalk')
const log = console.log
const axios = require('axios')
const ExcelJS = require('exceljs')

const host = 'https://wall-vs-me.myshopify.com/admin/api'
const getOrders = (password) => {
  axios({
    method: 'get',
    baseURL: host,
    url: '/2021-01/orders.json',
    params: {
      status: 'any'
    },
    responseType: 'json',
    headers: { 'X-Shopify-Access-Token': password }
  }).then((response) => {
    fs.writeFileSync('./orders.json', JSON.stringify(response.data))
    log(chalk.blue('Orders Data is Ready ... '));
  });

  log(chalk.blue('Create excel ... '));;
  log(chalk.blue('Excel is Ready ... '));
}

const createReport = async () => {
  log(chalk.blue('Create excel ... '));
  const { orders } = JSON.parse(fs.readFileSync('./orders.json'));
  let workbook =  createExcel(), worksheet = workbook.getWorksheet('Orders');
  let rowValues = [], rowNumber = 2, insertedRow;
  for (let order of orders) {
    rowValues = [order['created_at'], order['name'], order['total_price'] ,order['total_price'] - order['subtotal_price'],
    order['subtotal_price'], order['total_shipping_price_set']['shop_money']['amount'],
    order['total_tip_received'], order['total_tax_set']['shop_money']['amount'], parseInt(order['subtotal_price']) * 0.2,
      '', '', '']
    console.log(rowValues)
    insertedRow = worksheet.insertRow(rowNumber++, rowValues);
    insertedRow.font = { family: 4, size: 16 }
    insertedRow.alignment = { vertical: 'middle', horizontal: 'center' };
    rowValues = []
  }
  
  log(chalk.blue('Excel is Ready ... '));
  workbook.xlsx.writeFile('./report.xlsx');
}

const createExcel = () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Orders', { properties: { defaultRowHeight: 25, defaultColWidth: 35 } })
  const rowValues = ['Date', 'Order #', 'Total Amount', 'Transaction Fee',
    'Received Amount', 'Shipping', 'Tip for Artist', 'Customer Tax', 'Artist Commission', 'Manufacturing Cost', 'Profit', 'Profit %']
  const insertedRow = worksheet.insertRow(1, rowValues);
  insertedRow.font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true }
  insertedRow.alignment = { vertical: 'middle', horizontal: 'center' };
  
  return workbook
}

module.exports = {
  getOrders,
  createReport
}