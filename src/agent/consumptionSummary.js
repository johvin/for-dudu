const {
  toFixed,
  getYYYYMMDDDateStr,
} = require('../utils');

// 获取指定月份的代理商操作日志（max)
function parseAgentLogData(agentLogList, headerMap, yearMonth) {
  return agentLogList.reduce((a, b) => {
    const logDate = getYYYYMMDDDateStr(b[headerMap.logDate]);

    if (logDate.slice(0, 7) > yearMonth) {
      console.warn(colors.warn(`忽略时间大于 ${yearMonth} 的 max 日志：\n`), b);
      return a;
    }

    const record = {
      money: parseFloat(b[headerMap.money]),
      remaining: parseFloat(b[headerMap.remaining]) || 0, // 余额，暂时都是整数
      detail: b[headerMap.detail],
      date: logDate,
      optUser: b[headerMap.optUser]
    };

    if (!(b[headerMap.account] in a)) {
      a[b[headerMap.account]] = {
        account: b[headerMap.account],
        name: ('' + b[headerMap.companyName]).trim(),
        remaining: record.remaining,
        recharge: 0, // 充值金额，暂时没啥用
        consume: 0,
        records: []
      };
    }

    // 不是指定月的数据不做计算
    if (record.date.startsWith(yearMonth)) {
      a[b[headerMap.account]].records.push(record);
      // update recharge or consume
      a[b[headerMap.account]][record.money > 0 ? 'recharge' : 'consume'] += record.money;
    }

    return a;
  }, Object.create(null));
}

// 获取指定月份的代理商（消耗）订单数据
function parseAgentOrderData(agentOrderList, headerMap, yearMonth) {
  const agentOrderData = agentOrderList.reduce((a, b) => {
    const account = b[headerMap.account];
    const dateStr = getYYYYMMDDDateStr(b[headerMap.orderDate]);

    // 订单不属于该月份
    if (!dateStr.startsWith(yearMonth)) {
      console.warn(colors.warn(`不属于 ${yearMonth} 的代理商订单：\n`), b);
      return a;
    }

    const order = {
      id: b[headerMap.orderId],
      type: b[headerMap.orderType],
      realMoney: parseFloat(b[headerMap.realMoney]) // 实际金额
    };

    if (!(account in a)) {
      a[account] = {
        account,
        name: ('' + b[headerMap.companyName]).trim(),
        realMoney: 0,
        orders: []
      };
    }

    const { orders } = a[account];
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

    a[account].realMoney += order.realMoney;

    return a;
  }, Object.create(null));

  Object.values(agentOrderData).forEach((agentOrder) => {
    // 销售额 = 扣款金额 / 1.06
    agentOrder.sales = toFixed(agentOrder.realMoney, 1.06, 2);
    agentOrder.tax = agentOrder.realMoney - agentOrder.sales;

    // 订单类目（秀点、企业基础班等）
    agentOrder.orders.forEach((order) => {
      // 销售额 = 扣款金额 / 1.06
      order.sales = toFixed(order.realMoney, 1.06, 2);
      order.tax = order.realMoney - order.sales;
    });
  });

  return agentOrderData;
}


// 匹配代理商消耗(订单）和操作日志
function matchAgentOrderAndLog(agentOrderData, agentLogData) {
  Object.values(agentOrderData).forEach((agentOrder) => {
    if (agentOrder.found = (agentOrder.account in agentLogData)) {
      const agentLog = agentLogData[agentOrder.account];
      agentOrder.consume = agentLog.consume;
      agentOrder.match = agentOrder.realMoney + agentLog.consume === 0
    }
  });
}


// 获取代理商消耗（订单）汇总数据
function getAgentConsumptionSummaryData(agentOrderData) {
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

    // 订单类目（秀点、企业基础班等）
    agentOrder.orders.forEach((order) => {
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

  return data;
}

exports.parseAgentLogData = parseAgentLogData;
exports.parseAgentOrderData = parseAgentOrderData;
exports.matchAgentOrderAndLog = matchAgentOrderAndLog;
exports.getAgentConsumptionSummaryData = getAgentConsumptionSummaryData;
