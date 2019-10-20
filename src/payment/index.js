const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
const {
  parseAlipayData,
  parseAgentOrderData,
  getAlipaySummaryData,
} = require('./alipay');

const rootDir = '/Users/johvin/Documents/财务/支付平台/6月';
const yearMonth = '2018-06';

const filenames = {
  input: [
    '支付宝_原表.xlsx',
  ],
  output: [
    '支付宝_结算2.xlsx',
  ]
};

const hmAlipay = {
  account: 0,
  companyName: 1,
  money: 2,
  remaining: 3,
  detail: 4,
  logDate: 5,
  optUser: 6
};

const [{ data: alipayList }] = xlsx.parse(path.resolve(rootDir, filenames.input[0]));

alipayList.shift();

const alipayData = parseAlipayData(alipayList, hmAlipay, yearMonth);

function genAlipaySummaryReport() {
  const reportData = getAlipaySummaryData(alipayData);
  const buffer = xlsx.build([{ name: 'sheet', data: reportData }]);

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, filenames.output[0])).end(buffer, resolve);
  }).then(() => {
    console.log('支付宝结算报告搞定 ✌️️️️️✌️️️️️✌️️️️️');
  })
  .catch ((err) => {
    console.log('我也不知道哪里出错了，😒😒😒 你来看看吧');
    console.log(err);
  });;
}

genAlipaySummaryReport();
