const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
require('../colors');
const {
  getYYYYMMDDDateStr,
} = require('../utils');

const rootDir = '/Users/johvin/Documents/财务/营销会员';

const thisMonth = '2019/10';

// 要处理的数据的起始时间
const startMonthOfHandleData = '';

// 尚未开通
const notStart = 'notStart';

const inputFilenames = [
  '营销云会员明细表2016.01-2019.09.xlsx',
];

// header map
const hmOrder = {
  orderDate: 0,
  orderId: 1,
  region: 2,
  orderType: 3,
  platform: 4,
  paymentType: 5,
  source: 6,
  startDate: 7,
  endDate: 8,
  price: 9,
  noTaxPrice: 10,
  tax: 11,
  pricePerMonth: 12,
};

// 订单类型排序
const orderTypeOrder = {
  '基础版服务': -10,
  '高级版服务': -4,
  '畅享版服务': 0,
  '企业体验版': 4,

  '企业基础版': 10,
  '企业标准版': 20,
  '企业高级版': 30,
  '秀推基础版': 40,
  '秀推-企业基础版': 43,
  '秀推-企业标准版': 45,
  '秀推-企业高级版': 48,
  '秀推服务费': 50,
  '秀推-客脉追踪': 60,
  '秀推-人脉红包': 70,
  '秀推-商城': 80,
  '秀推-微信优惠券': 85,
  '秀推-客容': 90,
  '秀推-内容子账号': 100,
  '秀推-基础版管理子帐号': 110,
  '秀推-客脉追踪管理子帐号': 120,
  '秀推-人脉红包管理子帐号': 130,
  '秀推-商城管理子帐号': 140,
  '秀推-微信优惠券管理子帐号': 150,
  '全版本管理子账号': 160,
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

  const sheetList = xlsx.parse(filePath);

  sheetList.filter(it => it.name.includes('非白名单')).forEach((sheet) => {
    const { data: list } = sheet;
    list.shift();

    const dataMap = parseData(list, hmOrder);
    const summaryData1 = getSummaryByOrderTypeAndOrderDate(dataMap);
    const summaryData2 = getSummaryByOrderDateAndAverageMonth(dataMap);
    gen营销云SummaryReport(inputFilename, [{
      name: '营销云-月收款金额',
      data: summaryData1,
    }, {
      name: '营销云-月收入确认额',
      data: summaryData2,
    }]);
  });

}

// 收款金额
function getSummaryByOrderTypeAndOrderDate(dataMap) {
  const dateBounds = [];
  const table = [];
  const map = new Map();

  for (let [key, orderList] of dataMap.entries()) {
    const [orderMonth, orderType] = key.split('#');

    if (dateBounds.length === 0) {
      dateBounds[0] = dateBounds[1] = orderMonth;
    } else {
      if (orderMonth < dateBounds[0]) {
        dateBounds[0] = orderMonth;
      } else if (orderMonth > dateBounds[1]) {
        dateBounds[1] = orderMonth;
      }
    }

    if (!map.has(orderType)) {
      map.set(orderType, new Map());
    }

    const tmap = map.get(orderType);
    if (!tmap.has(orderMonth)) {
      tmap.set(orderMonth, 0);
    }
    let s = tmap.get(orderMonth);
    for (let record of orderList) {
      s += record.price;
    }
    tmap.set(orderMonth, s);
  }

  const tableHeader = ['收款金额'].concat(getDateRange(dateBounds[0], dateBounds[1], false));

  for (let orderType of map.keys()) {
    const tmap = map.get(orderType);
    table.push(
      tableHeader.map(
        (name, index) => index === 0 ? orderType : (tmap.get(name) || 0)
      )
    );
  }
  table.sort((a, b) => orderTypeOrder[a[0]] < orderTypeOrder[b[0]] ? -1 : 1);

  const total = ['合计'];
  for(let col = 1; col < tableHeader.length; col++) {
    for(let row of table) {
      total[col] = (total[col] || 0) + row[col];
    }
  }

  table.unshift(tableHeader);
  table.push(total);

  return table;
}

// 月收入确认额
function getSummaryByOrderDateAndAverageMonth(dataMap) {
  const dateBounds = [];
  const table = [];
  const map = new Map();

  for (let [key, orderList] of dataMap.entries()) {
    const [orderMonth, orderType] = key.split('#');

    if (!map.has(orderMonth)) {
      map.set(orderMonth, new Map());
    }
    const mmap = map.get(orderMonth);

    if (!mmap.has(orderType)) {
      mmap.set(orderType, new Map());
    }
    const tmap = mmap.get(orderType);

    for(let record of orderList) {
      const startMonth = record.startDate ? getYearMonth(record.startDate) : notStart;
      const endMonth = record.startDate ? getYearMonth(record.endDate) : notStart;

      if (startMonth === notStart) {
        tmap.set(startMonth, (tmap.get(startMonth) || 0) + record.pricePerMonth);
      } else {
        if (dateBounds.length === 0) {
          dateBounds[0] = startMonth;
          dateBounds[1] = endMonth;
        } else {
          if (startMonth < dateBounds[0]) {
            dateBounds[0] = startMonth;
          } else if (endMonth > dateBounds[1]) {
            dateBounds[1] = endMonth;
          }
        }

        if (orderType.includes('服务费')) {
          tmap.set(startMonth, (tmap.get(startMonth) || 0) + record.pricePerMonth);
        } else {
          for (let s of getDateRange(startMonth, endMonth)) {
            tmap.set(s, (tmap.get(s) || 0) + record.pricePerMonth);
          }
        }
      }
    }
  }

  const tableHeader = ['月收入确认额'].concat(getDateRange(dateBounds[0], dateBounds[1]), '尚未开通');

  const total = ['合计'];
  for (let [orderMonth, mmap] of map.entries()) {
    const mTotal = [orderMonth];
    const mArr = [];
    for(let [orderType, tmap] of mmap.entries()) {

      mArr.push(
        tableHeader.map((name, index) => {
          const val = index === 0 ? orderType : (tmap.get(index === tableHeader.length - 1 ? notStart : name) || 0);
          if (index > 0) {
            mTotal[index] = (mTotal[index] || 0) + val;
            total[index] = (total[index] || 0) + val;
          }
          return val;
        })
      );
    }
    mArr.sort((a, b) => orderTypeOrder[a[0]] < orderTypeOrder[b[0]] ? -1 : 1);
    mArr.unshift(mTotal);
    table.push(...mArr);
  }

  table.unshift(tableHeader);
  table.push(total);

  return table;
}

// 生成营销云汇总
function gen营销云SummaryReport(inputFilename, sheetData) {
  const buffer = xlsx.build(sheetData);

  const ext = path.extname(inputFilename);
  const outputFilename = `${path.basename(inputFilename, ext)}-汇总${ext}`;

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, outputFilename)).end(buffer, resolve);
  }).then(() => {
    console.log(colors.ok(`"${inputFilename}" 营销云汇总搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    console.log(colors.verbose(`输出文件路径: ${outputFilename}`));
  });
}

// 获取列表，去除无效数据
function parseData(list, headerMap) {
  const unpredefinedOrderTypeSet = new Set();
  const startMonth = startMonthOfHandleData ? getYearMonth(startMonthOfHandleData) : '';

  const map = list.reduce((a, b, index) => {
    const record = {
      orderDate: b[headerMap.orderDate],
      orderId: b[headerMap.orderId],
      orderType: b[headerMap.orderType],
      startDate: b[headerMap.startDate],
      endDate: b[headerMap.endDate],
      price: b[headerMap.price],
      pricePerMonth: b[headerMap.pricePerMonth],
    };

    try {
      record.orderDate = parseDate(record.orderDate);
    } catch (e) {
      // index + 1 + 1, 第一行是 header
      throw new Error(`row: ${index + 2}, ${e.message}`);
    }

    const orderMonth = getYearMonth(record.orderDate);

    if (startMonth && orderMonth < startMonth) {
      return a;
    }

    // 有些类型中会有多余空格
    record.orderType = record.orderType.split(' ').join('');

    if (!(record.orderType in orderTypeOrder)) {
      unpredefinedOrderTypeSet.add(record.orderType);
    }

    // 服务费开通结束时间认为是订单时间
    if (record.orderType.includes('服务费')) {
      record.startDate = record.endDate = record.orderDate;
      record.pricePerMonth = b[headerMap.noTaxPrice];
    }

    try {
      record.startDate = parseDate(record.startDate);
      record.endDate = parseDate(record.endDate);
    } catch (e) {
      // console.error(`row: ${index + 2}，尚未开通`);
      // console.info(b);
      record.startDate = record.endDate = '';
      record.pricePerMonth = b[headerMap.noTaxPrice];
    }

    const key = `${orderMonth}#${record.orderType}`;

    if (!a.has(key)) {
      a.set(key, []);
    }
    a.get(key).push(record);

    return a;
  }, new Map());

  if (unpredefinedOrderTypeSet.size > 0) {
    console.warn(colors.warn(`订单类型中包含 ${unpredefinedOrderTypeSet.size} 个不在已定义范围内的值：`));
    for (let it of unpredefinedOrderTypeSet) {
      console.log(it === '' || it === null || it === undefined ? '空值' : it);
    }
  }

  return map;
}

// 日期区间，
// isLessMaxDate 右开闭区间
function getDateRange(minDateStr, maxDateStr, isLessMaxDate = true) {
  const arr = [];
  for (let s = minDateStr; isLessMaxDate ? s < maxDateStr : s <= maxDateStr;) {
    arr.push(s);
    s = new Date(`${s}/01`);
    s.setMonth(s.getMonth() + 1);
    s = getYearMonth(s);
  }

  return arr;
}

function parseDate(value) {
  const dateRe = /^\d{4}(\/|-)\d+\1\d+$/;
  if (typeof value === 'number') {
    return getYYYYMMDDDateStr(value);
  } else if (typeof value === 'string' && dateRe.test(value.trim())) {
    return value.trim();
  }

  throw new Error(`日期格式错误, value: ${value}(${typeof value})`);
}

function getYearMonth(orderDate) {
  const d = new Date(orderDate);
  return d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0');
}
