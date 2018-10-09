const { toFixed } = require('../utils');

// 获取预付款数据
// containNote: 是否包含“小记”列
function parseAdvancePaymentData(advancePaymentList, headerMap, containNote) {
  return advancePaymentList.reduce((a, b) => {
    // 空行
    if (b.length === 0) return a;

    // 不包含“小记”列 or 只提取“小记”所在行数据
    if (!containNote || b[headerMap.note] === '小计') {
      const id = parseInt(b[headerMap.id], 10);

      if (isNaN(id)) {
        console.warn('record with invalid id: ', b);
      } else {
        a.push({
          id,
          name: ('' + b[headerMap.companyName]).trim(),
          jieMoney: parseFloat(b[headerMap.jieMoney]),
          daiMoney: parseFloat(b[headerMap.daiMoney]),
          direction: b[headerMap.direction],
          remaining: parseFloat(b[headerMap.remaining])
        });
      }
    }

    return a;
  }, []).sort((a, b) => a.id - b.id);
}

// 匹配 u8 预付款数据和代理商订单详情中的税额、操作日志中的余额
function matchAdvancePaymentAndAgent(advancePaymentList, agentOrderData, agentLogData) {
  agentOrderList = Object.values(agentOrderData);
  agentLogList = Object.values(agentLogData);

  advancePaymentList.forEach((ap) => {
    const { name } = ap;
    // 客户名称匹配
    let findLog = agentLogList.find(it => it.name === name);

    if (!findLog) {
      // try 中文分词
    }

    if (findLog) {
      const { account, remaining: maxRemaining } = findLog;
      const findOrder = agentOrderList.find(it => it.account === account);

      ap.account = account;
      ap.adjust = findOrder ? findOrder.tax : 0; // 没有消耗税额为 0
      ap.maxRemaining = maxRemaining; // max 系统余额
      ap.adjustRemaining = toFixed(ap.remaining - ap.adjust, 1, 2); // 调整后余额

      // 由于小数精度问题，导致计算结果有时候会有非常小的差 e.g. 1e-12
      if (Math.abs(ap.adjustRemaining) < 1e-6) {
        ap.adjustRemaining = 0;
      }
      ap.remainingDiff = toFixed(ap.adjustRemaining - ap.maxRemaining, 1, 2); // 差额
    } else {
      ap.account = null;
      ap.adjust = null; // excel里面，会计专用格式里面0就都是显示成-
      ap.maxRemaining = null;
      ap.adjustRemaining = null;
      ap.remainingDiff = null;
    }

    ap.match = !!findLog;
  });
}

// 获取 U8 预收账款核对明细数据
function getAdvancePaymentCheckData(advancePaymentList) {
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

  advancePaymentList.forEach((ap) => {
    data.push([
      ap.id,
      ap.account,
      ap.name,
      ap.jieMoney,
      ap.daiMoney,
      ap.direction,
      ap.remaining,
      ap.adjust,
      ap.adjustRemaining,
      ap.maxRemaining,
      ap.remainingDiff,
      ap.match ? null : 'Max 系统无匹配账户名称'
    ]);
  });

  return data;
}

exports.parseAdvancePaymentData = parseAdvancePaymentData;
exports.matchAdvancePaymentAndAgent = matchAdvancePaymentAndAgent;
exports.getAdvancePaymentCheckData = getAdvancePaymentCheckData;
