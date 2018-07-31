const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
const {
  toFixed,
  getYYYYMMDDDateStr,
} = require('../utils');

const rootDir = '/Users/nilianzhu/Documents/财务/例子-nlz/7月';
// 当月 key
const thisMonth = '2018-07';
let d = new Date(thisMonth);
d.setMonth(d.getMonth() - 1);
// 上个月 key
const lastMonth = d.toISOString().slice(0, 7);
// 上个月之前的月份 key
const monthBeforeLast = 'monthBeforeLast';
// 不存在创建日期
const noDate = 'noDate';
d = null;

const filenames = {
  input: [
    '2018.07开票统计-lrx（北京）.xlsx'
  ],
  output: [
    '开票明细汇总.xlsx'
  ]
};

// 发票 header map
const hmInvoice = {
  orderId: 3,
  orderDate: 5,
  paymentType: 6,
  orderType: 8,
  invoiceStatus: 10,
  invoiceValue: 17,
  noTaxInvoiceValue: 18,
  invoiceTax: 19
};

// 订单日期在数据表中的映射类型
const invoiceSummaryDateTypeMap = {
  [thisMonth]: '当月',
  [lastMonth]: '上个月',
  [monthBeforeLast]: '前期',
  [noDate]: '空日期'
};

// 支付类型识别 regexp
const paymentTypeRe = /微信|支付宝|对公打款|POS/i;

// 订单类型排序权值 map
const invoiceSummaryOrderTypeOrderMap = {
  '秀点': 0,
  '秀点（营销推广）': 3,
  '易企秀推广服务': 7,
  '企业基础版': 10,
  '企业标准版': 20,
  '企业高级版': 30,
  '定制服务': 40,
  '代理商服务费': 50,
  '代理商消耗': 60
};

// 发票明细汇总表中统计指标对应的名称
const invoiceIndicatorMap = {
  invoiceValue: '已开票金额',
  noTaxInvoiceValue: '不含税金额',
  invoiceTax: '税额'
};

const [{ data: invoiceList }] = xlsx.parse(path.resolve(rootDir, filenames.input[0]));
invoiceList.shift();

const invoiceData = parseInvoiceData(invoiceList, hmInvoice);
const negativeListObj = dealWithNegativeInvoice(invoiceData);
const invoiceSummary = getInvoiceSummaryByDimensions(invoiceData);
// console.log(JSON.stringify(invoiceSummary, null, 2));
genInvoiceDetailSummaryReport();

// 生成发票明细汇总表
function genInvoiceDetailSummaryReport() {
  const reportData = getInvoiceSummaryReportData(invoiceSummary);
  const { negativeMatchOrderIdList, negativeUnmatchOrderIdList } = negativeListObj;
  const matchData = negativeMatchOrderIdList.map(id => [id]);
  const unmatchData = negativeUnmatchOrderIdList.map(id => [id]);
  const buffer = xlsx.build([{
    name: '汇总',
    data: reportData
  }, {
    name: '红冲',
    data: unmatchData
  }, {
    name: '匹配负票',
    data: matchData
  }]);

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, thisMonth + filenames.output[0])).end(buffer, resolve);
  }).then(() => {
    console.log('发票明细汇总搞定 ✌️️️️️✌️️️️️✌️️️️️');
  });
}

// 获取汇总报告的 excel 数据
function getInvoiceSummaryReportData(invoiceSummary) {
  const data = [];
  for(let platform in invoiceSummary) {
    const d1 = invoiceSummary[platform];
    for (let date of Object.keys(invoiceSummaryDateTypeMap)) {
      if (date in d1) {
        const d2 = d1[date];
        // 表前加空行
        data.push([]);

        // 要显示的订单类型 columns，显示所有已定义类型和未定义但存在的类型
        const orderTypeArr = Object.keys(Object.assign({}, invoiceSummaryOrderTypeOrderMap, d2));
        
        orderTypeArr.sort((a, b) => {
          if (a in invoiceSummaryOrderTypeOrderMap && b in invoiceSummaryOrderTypeOrderMap) {
            return invoiceSummaryOrderTypeOrderMap[a] - invoiceSummaryOrderTypeOrderMap[b];
          }
          if (a in invoiceSummaryOrderTypeOrderMap) {
            return -1;
          }

          if (b in invoiceSummaryOrderTypeOrderMap) {
            return 1;
          }

          return a < b ? -1 : 1;
        });

        data.push([null, platform + invoiceSummaryDateTypeMap[date]].concat(orderTypeArr));

        const total = {
          invoiceValue: 0,
          noTaxInvoiceValue: 0,
          invoiceTax: 0
        };

        for(let key of ['invoiceValue', 'noTaxInvoiceValue', 'invoiceTax']) {
          const t = [invoiceIndicatorMap[key], 0];
          for(let orderType of orderTypeArr) {
            if (orderType in d2) {
              total[key] += d2[orderType][key];
              t.push(toFixed(d2[orderType][key], 1, 2));
            } else {
              t.push(null);
            }
          }
          t[1] = toFixed(total[key], 1, 2);
          data.push(t);
        }
      }
    }
  }

  return data;
}

// 计算不同维度的开票明细总计
function getInvoiceSummaryByDimensions(invoiceList) {
  const dims = {};

  invoiceList.forEach((invoice) => {
    if (!invoice) return;

    if (!(invoice.paymentType in dims)) {
      dims[invoice.paymentType] = {};
    }

    const dim1 = dims[invoice.paymentType];
    const dateType = getDateType(invoice.orderDate);

    if (!(dateType in dim1)) {
      dim1[dateType] = {};
    }

    const dim2 = dim1[dateType];

    if (!(invoice.orderType in dim2)) {
      dim2[invoice.orderType] = {
        invoiceValue: 0,
        noTaxInvoiceValue: 0,
        invoiceTax: 0
      };
    }

    const dim3 = dim2[invoice.orderType];

    dim3.invoiceValue += invoice.invoiceValue;
    dim3.noTaxInvoiceValue += invoice.noTaxInvoiceValue;
    dim3.invoiceTax += invoice.invoiceTax;
  });

  return dims;
}

// 处理负票&红冲，将负票和对应的正票设置为 null
// return 负票和红冲 orderId list
function dealWithNegativeInvoice(invoiceList) {
  const negativeList = invoiceList.filter(x => x.invoiceValue < 0);
  const negativeMatchOrderIdList = [];
  for(let i = 0; i < negativeList.length; i++) {
    const it = negativeList[i];
    const findIndex = invoiceList.findIndex(x => x && x.orderId === it.orderId && Math.abs(x.invoiceValue + it.invoiceValue) < 1e-6);
    if (findIndex > -1) {
      // 记录能匹配的
      negativeMatchOrderIdList.push(it.orderId);
      // 匹配负票的正票不计算到结果中，去掉
      invoiceList[findIndex] = null;
      negativeList[i] = null;
    }
  }

  invoiceList.forEach((it, index) => {
    // 去掉负票
    if (it && it.invoiceValue < 0) {
      invoiceList[index] = null;
    }
  });

  return {
    negativeMatchOrderIdList,
    negativeUnmatchOrderIdList: negativeList.filter(Boolean).map(it => it.orderId)
  };
}

// 获取发票信息列表，去除无效数据
function parseInvoiceData(invoiceList, headerMap, thisMonth) {
  // 日期 regexp
  const dateRe = /^\d{4}-\d{2}-\d{2}$/;
  const unpredefinedOrderTypeSet = new Set();

  const filterList = [];
  invoiceList.forEach((b, index) => {
    const invoiceStatus = b[headerMap.invoiceStatus];
    // 废票对明细没有影响，不统计
    if (typeof invoiceStatus !== 'string' || invoiceStatus.includes('废')) {
      return;
    }

    const paymentTypeRes = paymentTypeRe.exec(b[headerMap.paymentType]);
    // 支付类型不在匹配范围内，不处理
    if (!paymentTypeRes || paymentTypeRes.length === 0) {
      return;
    }

    const orderType = (b[headerMap.orderType] || '').replace(/\s/g, '');

    // 非已定义的订单类型给出 warning
    if (!(orderType in invoiceSummaryOrderTypeOrderMap)) {
      unpredefinedOrderTypeSet.add(orderType);
    }
  
    const record = {
      orderId: b[headerMap.orderId],
      orderDate: b[headerMap.orderDate],
      paymentType: paymentTypeRes[0],
      orderType,
      invoiceValue: parseFloat(b[headerMap.invoiceValue]),
      noTaxInvoiceValue: parseFloat(b[headerMap.noTaxInvoiceValue]),
      invoiceTax: parseFloat(b[headerMap.invoiceTax])
    };

    if (typeof record.orderDate === 'number') {
      record.orderDate = getYYYYMMDDDateStr(record.orderDate);
    } else if (typeof record.orderDate === 'string' && dateRe.test(record.orderDate.trim())) {
      record.orderDate = record.orderDate.trim();
    } else if (record.orderDate === null || record.orderDate === undefined) {
      // 空日期一般是预开票或代理商消耗（对上月消费的统一开票），两者都不存在实际订单
      record.orderDate = null;
    } else {
      console.error(`日期格式错误, row: ${index + 1}`, typeof record.orderDate, record.orderDate);
      throw new Error(`日期格式错误, row: ${index + 1}`);
    }

    filterList.push(record);
  });

  if (unpredefinedOrderTypeSet.size > 0) {
    console.warn(`订单类型中包含 ${unpredefinedOrderTypeSet.size} 个不在已定义范围内的值：`);
    for(let it of unpredefinedOrderTypeSet) {
      console.log(it === '' || it === null || it === undefined ? '空值' : it);
    }
  }

  return filterList;
}

// 获取日期所属类型
// return thisMonth/lastMonth/monthBeforeLast/noDate
function getDateType(date) {
  switch (true) {
    case date === null:
      return noDate;
    case date.startsWith(thisMonth):
      return thisMonth;
    case date.startsWith(lastMonth):
      return lastMonth;
    default:
      return monthBeforeLast;
  }
}