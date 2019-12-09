const fs = require('fs');
const path = require('path'); const xlsx = require('node-xlsx');
require('../colors');
const {
  toFixed,
  getYYYYMMDDDateStr,
} = require('../utils');

const rootDir = '/Users/nilianzhu/Documents/财务/发票/11月';
// 当月 key
const thisMonth = '2019-11';
// 上个月 key
const lastMonth = ((d) => (d.setMonth(d.getMonth() - 1), d.toISOString().slice(0, 7)))(new Date(thisMonth));
// 上个月之前的月份 key
const monthBeforeLast = 'monthBeforeLast';
// 不存在创建日期
const noDate = 'noDate';

const inputFilenames = [
  '2019.11开票统计（北京）_exception.xlsx',
];

// 发票 header map
// index 为负数表示不存在该列
const hmInvoice = {
  orderId: 0,
  orderDate: 7,
  paymentType: 8,
  orderType: 2,
  invoiceStatus: -1,
  invoiceValue: 4,
  noTaxInvoiceValue: 5,
  invoiceTax: 6
};

// 订单日期在数据表中的映射类型
const invoiceSummaryDateTypeMap = {
  [thisMonth]: '当月',
  [lastMonth]: '上个月',
  [monthBeforeLast]: '前期',
  [noDate]: '空日期' // 空日期代表预开票
};

// 支付类型识别 regexp
const paymentTypeRe = /微信|支付宝|苹果支付|对公打款|POS/i;

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

printMeta();
inputFilenames.forEach(process);


function printMeta() {
  console.log(colors.verbose(`正在处理 ${colors.em(colors.green(thisMonth))} 数据 ...\n文件夹路径: ${colors.em(rootDir)}，一共 ${inputFilenames.length} 个文件\n`));
}

// 处理
function process(inputFilename) {
  console.log(colors.verbose(`\n正在处理 "${colors.green(inputFilename)}" ...\n`));

  const filePath = path.resolve(rootDir, inputFilename);
  
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  const [{ data: invoiceList }] = xlsx.parse(filePath);
  invoiceList.shift();

  const invoiceData = parseInvoiceData(invoiceList, hmInvoice);
  const negativeListObj = dealWithNegativeInvoice(invoiceData);
  const invoiceSummary = getInvoiceSummaryByDimensions(invoiceData);
  genInvoiceDetailSummaryReport(invoiceSummary, negativeListObj, inputFilename);
}

// 生成发票明细汇总表
function genInvoiceDetailSummaryReport(invoiceSummary, negativeListObj, inputFilename) {
  const reportData = getInvoiceSummaryReportData(invoiceSummary);
  const { negativeMatchOrderList, negativeUnmatchOrderList } = negativeListObj;

  const getNegativeData = orderList => [['订单号', '发票金额', '不含税金额', '税额', '支付方式']].concat(orderList.map(it => [it.orderId, it.invoiceValue, it.noTaxInvoiceValue, it.invoiceTax, it.paymentType]));
  const matchData = getNegativeData(negativeMatchOrderList);
  const unmatchData = getNegativeData(negativeUnmatchOrderList);
  const buffer = xlsx.build([{
    name: '汇总',
    data: reportData
  }, {
    name: '不匹配红冲',
    data: unmatchData
  }, {
    name: '匹配红冲',
    data: matchData
  }]);
  
  const ext = path.extname(inputFilename);
  const outputFilename = `${path.basename(inputFilename, ext)}-汇总${ext}`;

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, outputFilename)).end(buffer, resolve);
  }).then(() => {
    console.log(colors.ok(`"${inputFilename}" 发票明细汇总搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    console.log(colors.verbose(`输出文件路径: ${outputFilename}`));
  });
}

// 获取汇总报告的 excel 数据
function getInvoiceSummaryReportData(invoiceSummary) {
  const data = [];
  for(let platform in invoiceSummary) { // 平台是支付方式
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
// 数据格式：
// {
//   '微信': {
//     [thisMonth]: {
//       '秀点': {
//         invoiceValue: 0,
//         noTaxInvoiceValue: 0,
//         invoiceTax: 0
//       }
//     }
//   }
// }
function getInvoiceSummaryByDimensions(invoiceList) {
  const dims = {};

  invoiceList.forEach((invoice) => {
    if (!invoice) return;

    // 支付类型作为一级 key
    if (!(invoice.paymentType in dims)) {
      dims[invoice.paymentType] = {};
    }

    const dim1 = dims[invoice.paymentType];
    const dateType = getDateType(invoice.orderDate);

    // 日期类型作为二级 key
    if (!(dateType in dim1)) {
      dim1[dateType] = {};
    }

    const dim2 = dim1[dateType];

    // 订单类型作为三级 key
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

// 处理红冲，将负票和对应的正票设置为 null
// @return 匹配上的红冲和不匹配的红冲 orderList
// 能匹配的红冲通常是重新变更单位开票的情况，不匹配的红冲通常是退款、预开票未交钱的情况
function dealWithNegativeInvoice(invoiceList) {
  const negativeList = invoiceList.filter(x => x.invoiceValue < 0);
  const negativeMatchOrderList = [];
  for(let i = 0; i < negativeList.length; i++) {
    const it = negativeList[i];
    // 没有订单号是预开票未交钱的情况，属于不匹配的情况
    const findIndex = it.orderId ? invoiceList.findIndex(x => x && x.orderId === it.orderId && Math.abs(x.invoiceValue + it.invoiceValue) < 1e-6) : -1;
    if (findIndex > -1) {
      // 记录能匹配的
      negativeMatchOrderList.push(it);
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
    negativeMatchOrderList,
    negativeUnmatchOrderList: negativeList.filter(Boolean)
  };
}

// 获取发票信息列表，去除无效数据
function parseInvoiceData(invoiceList, headerMap) {
  // 日期 regexp
  const dateRe = /^\d{4}-\d{2}-\d{2}$/;
  // 记录订单类型不在预定义集合中的
  const unpredefinedOrderTypeSet = new Set();

  const filterList = [];
  invoiceList.forEach((b, index) => {
    const invoiceStatus = headerMap.invoiceStatus >= 0 && b[headerMap.invoiceStatus];
    // 存在该列时，状态为废票的数据对明细没有影响，不统计
    // 不存在该列时，认为所有数据均为有效状态
    if (headerMap.invoiceStatus >= 0 && (typeof invoiceStatus !== 'string' || invoiceStatus.includes('废'))) {
      return;
    }

    const paymentTypeRes = paymentTypeRe.exec(b[headerMap.paymentType]);
    // 支付类型不在匹配范围内，不处理
    if (!paymentTypeRes || paymentTypeRes.length === 0) {
      return;
    }

    let orderType = (b[headerMap.orderType] || '').replace(/\s/g, '');
    if (/^\d+秀点$/.test(orderType)) {
      orderType = '秀点';
    }

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
    } else if (typeof record.orderDate === 'string') {
      const t = record.orderDate.trim();
      if (dateRe.test(t)) {
        record.orderDate = t;
      } else {
        const dateArr = t.split(' ').filter(x => dateRe.test(x.trim()));
        if (dateArr.length === 0) {
          console.error(colors.error(`日期格式错误, row: ${index + 1}`), typeof record.orderDate, record.orderDate);
          throw new Error(`日期格式错误, row: ${index + 1}`);
        }
        record.orderDate = dateArr[0];
        if (dateArr.length > 1) {
          console.warn(colors.warn(`多个日期，row: ${index + 1}`), record.orderDate);
        }

      }
    } else if (record.orderDate === null || record.orderDate === undefined) {
      // 空日期一般是预开票或代理商消耗（对上月消费的统一开票），两者都不存在实际订单
      record.orderDate = null;
    } else {
      console.error(colors.error(`日期格式错误, row: ${index + 1}`), typeof record.orderDate, record.orderDate);
      throw new Error(`日期格式错误, row: ${index + 1}`);
    }

    filterList.push(record);
  });

  if (unpredefinedOrderTypeSet.size > 0) {
    console.warn(colors.warn(`订单类型中包含 ${unpredefinedOrderTypeSet.size} 个不在已定义范围内的值：`));
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
