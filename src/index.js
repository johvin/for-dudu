const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');

const rootDir = '/Users/johvin/Documents/财务/代理商报表/6月';
const yearMonth = '2018-06'

const agentLogData = getAgentLogData(yearMonth);
const agentOrderData = getAgentOrderData(yearMonth);

calcAgentConsumeSummary(agentOrderData, agentLogData);
generateAgentConsumeReport(agentOrderData, '代理商消耗汇总.xlsx')
.then(() => {
  console.log('代理商消耗汇总搞定 ✌️️️️️✌️️️️️✌️️️️️');
  // process.exit(0);
})
.then(() => {
  const u8DepositData = getU8DepositData();
  matchU8AndAgent(u8DepositData, agentOrderData, agentLogData);
    return generateU8CheckReport(u8DepositData, '预收账款核对.xlsx');
  })
  .then(() => {
    console.log('预收账款核对搞定 ✌️️️️️✌️️️️️✌️️️️️');
  })
  .catch((err) => {
    console.log('我也不知道哪里出错了，😒😒😒 你来看看吧');
    console.log(err);
  });

// 获取代理商指定月份的操作日志
function getAgentLogData (yearMonth) {
  const [{ data: agentLogList }] = getData('max.xls');
  // remove header
  agentLogList.shift();

  return agentLogList.reduce((a, b) => {
    const record = {
      money: parseFloat(b[2]),
      remaining: parseFloat(b[3]) || 0, // 余额，暂时都是整数
      detail: b[4],
      date: b[5],
      optUser: b[6]
    };

    if (!(b[0] in a)) {
      a[b[0]] = {
        account: b[0],
        name: ('' + b[1]).trim(),
        remaining: record.remaining,
        recharge: 0, // 充值金额，暂时没啥用
        consume: 0,
        records: []
      };
    }

    if (record.date.startsWith(yearMonth)) {
      a[b[0]].records.push(record);
      // update recharge or consume
      a[b[0]][record.money > 0 ? 'recharge' : 'consume'] += record.money;
    }

    return a;
  }, Object.create(null));
}

// 获取代理商（消耗）订单列表数据
function getAgentOrderData(yearMonth) {
  const [{ data: agentOrderList }] = getData('代理商订单.xlsx');
  // remove header
  agentOrderList.shift();

  // // 距离 1900 年的天数
  // const days = agentOrderList[0][10];
  // // 转化为 js 中时间的毫秒数
  // const time = (days - 2) * 24 * 3600 * 1000 + new Date('1900/1/1').getTime();

  // if (!dateMatch(time, yearMonth)) {
  //   throw new Error('代理商消耗明细表月份不对');
  // }

  return agentOrderList.reduce((a, b) => {
    const order = {
      id: b[4],
      type: b[5],
      realMoney: parseFloat(b[12]) // 实际金额
    };

    if (!(b[0] in a)) {
      a[b[0]] = {
        account: b[0],
        name: ('' + b[3]).trim(),
        realMoney: 0,
        orders: []
      };
    }

    const { orders } = a[b[0]];
    const orderSummary = orders.find(it => it.type === order.type);

    if (!orderSummary) {
      orders.push({
        type: order.type,
        realMoney: order.realMoney,
        records: [order]
      })
    } else {
      orderSummary.realMoney += order.realMoney;
      orderSummary.records.push(order);
    }

    a[b[0]].realMoney += order.realMoney;

    return a;
  }, Object.create(null));
}

// 获取 U8 预收账款数据
function getU8DepositData() {
  const [{ data: depositList }] = getData('U8.xls');
  // remove header
  depositList.shift();

  return depositList.reduce((a, b) => {
    if (b[8] === '小计') {
      a.push({
        id: parseInt(b[4], 10),
        name: ('' + b[5]).trim(),
        jieMoney: parseFloat(b[9]),
        daiMoney: parseFloat(b[10]),
        direction: b[11],
        remaining: parseFloat(b[12])
      });
    }

    return a;
  }, []).sort((a, b) => a.id - b.id);
}

// 匹配代理商消耗和操作日志
function calcAgentConsumeSummary(agentOrderData, agentLogData) {
  Object.values(agentOrderData).forEach((agentOrder) => {
    if (agentOrder.found = (agentOrder.account in agentLogData)) {
      const agentLog = agentLogData[agentOrder.account];
      agentOrder.match = agentOrder.realMoney + agentLog.consume === 0

      // 如果不 match，记录下操作日志中的数值
      if (!agentOrder.match) {
        agentOrder.consume = agentLog.consume;
      }
    }
  });
}

// 匹配 u8 数据和代理商订单详情中的税额、操作日志中的余额
function matchU8AndAgent(u8DepositData, agentOrderData, agentLogData) {
  agentOrderList = Object.values(agentOrderData);
  agentLogList = Object.values(agentLogData);

  u8DepositData.forEach((deposit) => {
    const { name } = deposit;
    // 客户名称匹配
    let findLog = agentLogList.find(it => it.name === name);

    if (!findLog) {
      // try 中文分词
    }

    if (findLog) {
      const { account, remaining: maxRemaining } = findLog;
      const findOrder = agentOrderList.find(it => it.account === account);

      deposit.account = account;
      deposit.adjust = findOrder ? findOrder.tax : 0; // 没有消耗税额为 0
      deposit.maxRemaining = maxRemaining; // max 系统余额
      deposit.adjustRemaining = toFixed(deposit.remaining - deposit.adjust, 1, 2); // 调整后余额

      // 由于小数精度问题，导致计算结果有时候会有非常小的差 e.g. 1e-12
      if (Math.abs(deposit.adjustRemaining) < 1e-6) {
        deposit.adjustRemaining = 0;
      }
      deposit.remainingDiff = toFixed(deposit.adjustRemaining - deposit.maxRemaining, 1, 2); // 差额
    } else {
      deposit.account = null;
      deposit.adjust = null; // excel里面，会计专用格式里面0就都是显示成-
      deposit.maxRemaining = null;
      deposit.adjustRemaining = null;
      deposit.remainingDiff = null;
    }

    deposit.match = !!findLog;
  });
}

// 生成代理商消耗汇总表
function generateAgentConsumeReport(agentOrderData, filename) {
  const data = [
    [
      '代理商名称',
      '扣款金额（元）',
      '销售额（元）',
      '税额（元）',
      '备注'
    ]
  ];

  const total = {
    realMoney: 0,
    sales: 0,
    tax: 0
  }

  Object.values(agentOrderData).forEach((agentOrder) => {
    // 销售额
    agentOrder.sales = toFixed(agentOrder.realMoney, 1.06, 2);
    agentOrder.tax = agentOrder.realMoney - agentOrder.sales;

    data.push([
      agentOrder.name,
      agentOrder.realMoney,
      agentOrder.sales,
      agentOrder.tax,
      !agentOrder.found ? '操作日志中未找到匹配记录' : !agentOrder.match ? `日志中总消耗：${agentOrder.consume}` : null
    ]);

    total.realMoney += agentOrder.realMoney;
    total.sales += agentOrder.sales;
    total.tax += agentOrder.tax;

    // 订单类目
    agentOrder.orders.forEach((order) => {
      // 销售额
      order.sales = toFixed(order.realMoney, 1.06, 2);
      order.tax = order.realMoney - order.sales;

      data.push([
        order.type,
        order.realMoney,
        order.sales,
        order.tax,
        null
      ]);
    });
  });

  data.push([
    '总计',
    total.realMoney,
    total.sales,
    total.tax,
    null
  ]);

  const buffer = xlsx.build([{ name: 'sheet', data: data }]);
  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, filename)).end(buffer, resolve);
  });
}

// 生成 U8 预收账款核对明细表
function generateU8CheckReport(u8DepositData, filename) {
  const data = [
    [
      '客户编号',
      '代理商账号',
      '客户名称',
      '借方金额',
      '贷方金额',
      '方向',
      '余额金额',
      '本期调整',
      '调后余额',
      'MAX系统余额',
      '差额',
      '备注'
    ]
  ];

  u8DepositData.forEach((deposit) => {
    data.push([
      deposit.id,
      deposit.account,
      deposit.name,
      deposit.jieMoney,
      deposit.daiMoney,
      deposit.direction,
      deposit.remaining,
      deposit.adjust,
      deposit.adjustRemaining,
      deposit.maxRemaining,
      deposit.remainingDiff,
      deposit.match ? null : 'Max 系统无匹配账户名称'
    ]);
  });

  const buffer = xlsx.build([{ name: 'sheet', data: data }]);
  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, filename)).end(buffer, resolve);
  });
}

function getData(filename) {
  return xlsx.parse(path.resolve(rootDir, filename));
}

function dateMatch(dateStr, yearMonth) {
  console.log(dateStr, yearMonth);
  const d = new Date(dateStr);
  const m = d.getMonth() + 1 + '';
  return d.getFullYear + m === yearMonth;
}

function toFixed(dividend, divisor, n = 2) {
  const weight = 10 ** n;
  return Math.round(dividend * weight / divisor) / weight;
}
