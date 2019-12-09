const readdirp = require('readdirp');
const xlsx = require('node-xlsx');
const utils = require('./utils');
require('../colors');

const rootDir = '/Users/nilianzhu/Documents/财务/收入流水/';
const thisMonth = '2019.06';
const initMemory = process.memoryUsage();

function printMemory() {
  const {
    rss,
    heapTotal,
    heapUsed,
    external,
  } = process.memoryUsage();
  const getUnit = (init = 'B') => {
    const unitArr = ['B', 'KB', 'MB', 'GB'];
    function unit() {
      return unitArr[unit.index];
    }
    unit.index = unitArr.indexOf(init);
    unit.increase = () => (unit.index += 1, unit);

    return unit;
  };
  const f = (n, u = getUnit()) => n < 1024 ? `${n}${u()}` : (n = ~~(n/1024), f(n, u.increase()));
  console.log(`memory delta => rss: ${f(rss - initMemory.rss)}, heapTotal: ${f(heapTotal - initMemory.heapTotal)}, heapUsed: ${f(heapUsed - initMemory.heapUsed)}, external: ${f(external - initMemory.external)}`);
  console.log(`memory value => rss: ${f(rss)}, heapTotal: ${f(heapTotal)}, heapUsed: ${f(heapUsed)}, external: ${f(external)}`);
}

(async () => {
  const ret = await getFilesTree(rootDir);
  processFile('成都', ret.get('成都'));
  // console.log(ret);
})();

async function getFilesTree(rootDir) {
  const files = await readdirp.promise(rootDir, { fileFilter: ['*.xlsx', '!~*.xlsx'] });
  const area = new Map();
  
  for(let file of files) {
    // rootDir 下的文件忽略
    if (!file.path.includes('/')) continue;

    const slices = file.path.split('/');
    if (slices.length < 3) throw new Error(`unexpected file: ${file.path}`);

    if (!area.has(slices[0])) {
      area.set(slices[0], new Map());
    }
    if (!area.get(slices[0]).has(slices[1])) {
      area.get(slices[0]).set(slices[1], []);
    }
    area.get(slices[0]).get(slices[1]).push(file);
  }

  return area;
}

// printMeta();
// inputFilenames.forEach(process);

function printMeta(parent, child) {
  console.log(colors.verbose(`正在处理 ${colors.em(colors.green(`${parent} => ${child}`))} 数据 ...\n`));
}

// 处理
function processFile(area, fileMap) {
  for(let type of fileMap.keys()) {

    // TODO
    if (type === '订单列表') continue;

    printMeta(area, type);
    const files = fileMap.get(type);

    for(let file of files) {
      if (!/微信|支付宝|银行|pos/i.test(file.basename)) continue;
      printMeta(type, file.basename);
      getFileData(file);
      printMemory();
    }
  }

  getAreaSheet(fileMap);
  printMemory();
}

function getAreaSheet(fileMap) {
  const table = [];
  table.push([], ['Per 收款渠道']);
  const files = fileMap.get('支付平台').filter(f => f.content);
  // 要统计的月份
  const monthArr = files[0].content.map(it => it.name).reduce((a, b) => {
    if (b <= thisMonth) a.push(b);
    return a;
  }, []);
  
  table.push([''].concat(monthArr));

  const thirdPartyFiles = files.filter(it => !/银行/.test(it.basename));
  if (thirdPartyFiles.length > 0) {
    table.push(...get3rdPartyTable(thirdPartyFiles, monthArr));
  }
  const bankFiles = files.filter(it => /银行/.test(it.basename));
  if (bankFiles.length > 0) {
    table.push([], [''].concat(monthArr), ...getBankTableData(bankFiles[0], monthArr));
  }

  console.table(table);
}

// 获取第三方渠道数据
function get3rdPartyTable(fileArr, monthArr) {
  const data = [];
  for(let file of fileArr) {
    if (/微信/.test(file.basename)) {
      data.push(...getWechatTableData(file, monthArr));
    } else if (/支付宝/.test(file.basename)) {
      data.push(...getAlipayTableData(file, monthArr));
    } else if (/pos/i.test(file.basename)) {
      data.push(...getPOSTableData(file, monthArr));
    } else {
      console.warn(colors.warn(`unknown income file: ${file.basename}`));
    }
  }

  const total = Array(monthArr.length + 1).fill(0);
  total[0] = ['第三方渠道收获'];
  const totalRe = /微信|支付宝|pos/i;

  for(let row of data) {
    if (totalRe.test(row[0])) {
      for(let i = 1; i < total.length; i++) {
        total[i] += row[i];
      }
    }
  }
  data.unshift(total);

  return data;
}

function getWechatTableData(file, monthArr) {
  const ignoreType = /测试|白名单|红包充值/i;
  const data = [];
  // 北京分3个账户统计，成都一个账户
  for(let summaryType of file.path.includes('北京') ? ['微信1', '微信2', '微信3'] : ['微信']) {
    data.push(...getSummaryTableData({
      file,
      summaryType,
      monthArr,
      handleRow: (row) => {
        return !ignoreType.test(row.type) && (row.account ? row.account === summaryType : true);
      },
    }));
  }
  return data;
}

function getAlipayTableData(file, monthArr) {
  const ignoreType = /测试|白名单/i;
  return getSummaryTableData({
    file,
    summaryType: '支付宝',
    monthArr,
    handleRow: (row) => {
      return !ignoreType.test(row.type);
    },
  });
}

function getPOSTableData(file, monthArr) {
  const ignoreType = /测试|白名单/i;
  return getSummaryTableData({
    file,
    summaryType: 'POS',
    monthArr,
    handleRow: (row) => {
      return !ignoreType.test(row.type);
    },
  });
}

function getBankTableData(file, monthArr) {
  const ignoreType = /微信|支付宝|pos|测试|白名单/i;
  return getSummaryTableData({
    file,
    summaryType: '转账收款',
    monthArr,
    handleRow: (row) => {
      return !ignoreType.test(row.type);
    },
  });
}

// 根据 type 做分类汇总
function getSummaryTableData({ file, monthArr, summaryType, handleRow }) {
  const map = new Map();

  for(let sheet of file.content) {
    for(let row of sheet.data) {
      if (!handleRow(row, sheet.name)) continue;
      if (row.type === '无人认领' || row.type === '其它') row.type = '其他';
      if (!map.has(row.type)) map.set(row.type, new Map());
      
      const m = map.get(row.type);
      m.set(sheet.name, (m.get(sheet.name) || 0) + row.money);
    }
  }

  const totalMap = new Map();
  const data = [];
  for(let type of map.keys()) {
    const m = map.get(type);
    const arr = [type];
    for(let month of monthArr) {
      arr.push(m.get(month) || 0);
      totalMap.set(month, (totalMap.get(month) || 0) + (m.get(month) || 0));
    }
    data.push(arr);
  }
  data.push([summaryType].concat(monthArr.map(month => totalMap.get(month))));

  data.sort((a, b) => {
    return typeOrder[a[0]] < typeOrder[b[0]] ? -1 : 1;
  });

  return data; 
}

function getAreaTotalTable() {
}

// 用于最终 excel 的行排列
const typeOrder = {
  '流水收入': -30,
  '营销中心': -26,
  '内容中心': -20,
  '第三方渠道收款': -16,
  '微信1': -12,
  '微信2': -10,
  '微信3': -8,
  '支付宝': -6,
  '转账收款': -4,
  '流量变现': 0,
  '增值服务': 4,
  '营销云会员': 8,
  '创意云会员': 12,
  '代理商预付款': 16,
  '代理商保证金': 20,
  '其他': 24,
};

// 获取 sheet 数据的 utils map
const getDataMap = new Map();
const thirdParty = new Map();
getDataMap.set('支付平台', thirdParty);
thirdParty.set('支付宝', utils.getAlipayData);
thirdParty.set('微信', utils.getWechatData);
// TODO 成都微信需要和北京微信统计格式
thirdParty.set('微信cd', utils.getChengDuWechatData);
thirdParty.set('银行', utils.getBankData);
thirdParty.set('pos', utils.getPOSData);
thirdParty.set('POS', utils.getPOSData);
const orderList = new Map();
getDataMap.set('订单列表', orderList);
thirdParty.set('支付宝', utils.getAlipayData);
thirdParty.set('微信', utils.getWechatData);
thirdParty.set('银行', utils.getBankData);
thirdParty.set('pos', utils.getPOSData);
thirdParty.set('POS', utils.getPOSData);

// 读取 excel 文件数据
function getFileData(file) {
  const match = file.path.match(/(支付平台|订单列表)\/(微信|支付宝|银行|pos)/i);
  if (!match || match.length < 3) {
    throw new Error(`unexpected file: ${file.path}`);
  }
  if (match[1] === '支付平台' && match[2] === '微信' && file.path.includes('成都')) {
    match[2] = '微信cd';
  }

  const extractMethod = getDataMap.get(match[1]).get(match[2])
  const ret = xlsx.parse(file.fullPath);
  file.content = ret.map(it => ({
    name: it.name,
    data: extractMethod(it.data),
  }));
}
