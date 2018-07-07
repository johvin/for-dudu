const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
const {
  parseAgentLogData,
  parseAgentOrderData,
  matchAgentOrderAndLog,
  getAgentConsumptionSummaryData,
} = require('./consumptionSummary');
const {
  parseAdvancePaymentData,
  matchAdvancePaymentAndAgent,
  getAdvancePaymentCheckData,
} = require('./advancePayment');

const rootDir = '/Users/johvin/Documents/财务/代理商报表/6月';
const yearMonth = '2018-06';

const filenames = {
  input: [
    'max.xls',
    '代理商订单.xlsx',
    'U8.xls',
  ],
  output: [
    '代理商消耗汇总.xlsx',
    '预收账款核对.xlsx'
  ]
};

const hmLog = {
  account: 0,
  companyName: 1,
  money: 2,
  remaining: 3,
  detail: 4,
  logDate: 5,
  optUser: 6
};

const hmOrder = {
  account: 0,
  companyName: 3,
  orderId: 4,
  orderType: 5,
  orderDate: 10,
  realMoney: 12
};

const hmPayment = {
  id: 4,
  companyName: 5,
  jieMoney: 9,
  daiMoney: 10,
  direction: 11,
  remaining: 12,
  note: 8
};

const [{ data: agentLogList }] = xlsx.parse(path.resolve(rootDir, filenames.input[0]));
const [{ data: agentOrderList }] = xlsx.parse(path.resolve(rootDir, filenames.input[1]));

agentLogList.shift();
agentOrderList.shift();

const agentLogData = parseAgentLogData(agentLogList, hmLog, yearMonth);
const agentOrderData = parseAgentOrderData(agentOrderList, hmOrder, yearMonth);

function genAgentConsumptionSummaryReport() {
  matchAgentOrderAndLog(agentOrderData, agentLogData);

  const reportData = getAgentConsumptionSummaryData(agentOrderData);
  const buffer = xlsx.build([{ name: 'sheet', data: reportData }]);

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, filenames.output[0])).end(buffer, resolve);
  }).then(() => {
    console.log('代理商消耗汇总搞定 ✌️️️️️✌️️️️️✌️️️️️');
  });
}

function genConsumptionAndPaymentReport() {
  return genAgentConsumptionSummaryReport()
  .then(() => {
    const [{ data: originList }] = xlsx.parse(path.resolve(rootDir, filenames.input[2]));
    originList.shift();

    const advancePaymentList = parseAdvancePaymentData(originList, hmPayment, true);
    matchAdvancePaymentAndAgent(advancePaymentList, agentOrderData, agentLogData);
    return getAdvancePaymentCheckData(advancePaymentList);
  })
  .then((reportData) => {
    const buffer = xlsx.build([{ name: 'sheet', data: reportData }]);

    return new Promise((resolve) => {
      fs.createWriteStream(path.resolve(rootDir, filenames.output[1])).end(buffer, resolve);
    })
    .then(() => {
      console.log('预收账款核对搞定 ✌️️️️️✌️️️️️✌️️️️️');
    });
  })
  .catch((err) => {
    console.log('我也不知道哪里出错了，😒😒😒 你来看看吧');
    console.log(err);
  });;
}


// genAgentConsumptionSummaryReport();

genConsumptionAndPaymentReport();
