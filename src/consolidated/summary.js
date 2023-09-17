const fs = require('fs');
const path = require('path');
const { rootDir, fileDir, outputFile } = require('./constants');
const { readData, readText, genExcel } = require('../excel');
require('../colors');

// 统计项目名称
const resultItem = {
  property: readText(__dirname, './resultItems/property.txt').split(/\n/).map(it => it.trimEnd()),
  debtEquity: readText(__dirname, './resultItems/debtEquity.txt').split(/\n/).map(it => it.trimEnd()),
  profit: readText(__dirname, './resultItems/profit.txt').split(/\n/).map(it => it.trimEnd()),
  cashFlow: readText(__dirname, './resultItems/cashFlow.txt').split(/\n/).map(it => it.trimEnd()),
};

// console.log(resultItem);

process();

// 处理
function process() {
  const sourceDir = path.resolve(rootDir, fileDir.summary);
  console.log(colors.verbose(`正在处理数据 ...\n源文件夹路径: ${colors.em(colors.green(sourceDir))}`));

  const filenames = fs.readdirSync(sourceDir);

  const inputFiles = filenames.filter(n => /^\d+_/.test(n)).sort();

  if (inputFiles.length === 0) {
    console.log(colors.error('未找到要处理文件，请检查目录或文件名'));
    global.process.exit(1);
  }

  const propertyList = [];
  const debtEquityList = [];
  const profitList = [];
  const cashFlowList = [];

  // 读取数据
  inputFiles.forEach(filename => {
    const company = filename.split(/_|\s/).slice(0, 2).map(t => t.trim()).join('_');
    const sheets = readData(sourceDir, filename);
    const [balanceSheet, profitSheet, cashFlowSheet] = sheets;

    // 前 4 行无用
    balanceSheet.data.splice(0, 4);
    propertyList.push({
      name: company,
      // 资产负债表数据一拆二
      list: balanceSheet.data.map(it => it.slice(0, 3)),
    });
    debtEquityList.push({
      name: company,
      // 资产负债表数据一拆二
      list: balanceSheet.data.map(it => it.slice(3, 6)),
    });
    profitSheet.data.splice(0, 4);
    profitList.push({
      name: company,
      list: profitSheet.data,
    });
    cashFlowSheet.data.splice(0, 4);
    cashFlowList.push({
      name: company,
      list: cashFlowSheet.data,
    });
  });

  cleanData(propertyList);
  cleanData(debtEquityList);
  cleanData(profitList);
  cleanData(cashFlowList);

  // for (const c of debtEquityList) {
  //   console.log(c.name);
  //   for (const d of c.list) {
  //     console.log(JSON.stringify(d));
  //   }

  // }

  const propertySummary = summaryList(resultItem.property, propertyList);
  const debtEquitySummary = summaryList(resultItem.debtEquity, debtEquityList);
  const profitSummary = summaryList(resultItem.profit, profitList);
  const cashFlowSummary = summaryList(resultItem.cashFlow, cashFlowList);

  genSummaryReport([
    propertySummary[0],
    propertySummary[1].concat(debtEquitySummary[1]),
    propertySummary[2].concat(debtEquitySummary[2]),
  ], profitSummary, cashFlowSummary);
}

function cleanData(sheet) {
  for (const c of sheet) {
    c.list.forEach(it => {
      it[0] = (it[0] ?? '').trim();
    });
    // 项目不存在的过滤掉
    c.list = c.list.filter(it => it[0].length > 0);
  }
}

function summaryList(items, companyList) {
  const sheetHeaders = ['项目名称'];
  const data1 = [];
  const data2 = [];

  companyList.forEach(c => {
    sheetHeaders.push(c.name);
  });

  // 项目名称匹配不到的公司列表
  const misMatchList = new Map();

  for (const item of items) {
    const row1 = [item];
    const row2 = [item];

    const matchItemName = item.trimStart();
    if (matchItemName.length === 0) {
      continue;
    }

    for (const c of companyList) {
      const info = c.list.find(it => it[0] === matchItemName);

      if (!info) {
        if (misMatchList.has(matchItemName)) {
          misMatchList.get(matchItemName).push(c.name);
        } else {
          misMatchList.set(matchItemName, [ c.name ]);
        }
      }

      row1.push(info?.[1]);
      row2.push(info?.[2]);
    }
    data1.push(row1);
    data2.push(row2);
  }

  for(const [mName, list] of misMatchList) {
    console.warn(`匹配不到 [${mName}] 的公司列表：`);
    for(const l of list) {
      console.warn(`  ${l}`);
    }
  }

  return [sheetHeaders, data1, data2];
}

// 生成合并报表
function genSummaryReport(...summarys) {
  const outputFilename = `${outputFile.summary}.xlsx`;
  const sheetNames = [
    ['资产负债表-期末', '资产负债表-年初'],
    ['利润表-当期', '利润表-累计'],
    ['现金流量表-当期', '现金流量表-累计'],
  ];

  const sheetData = [];

  summarys.forEach((s, idx) => {
    const [header, list1, list2] = s;
    const data1 = [header].concat(list1);
    const data2 = [header].concat(list2);

    sheetData.push({
      name: sheetNames[idx][0],
      data: data1,
    });
    sheetData.push({
      name: sheetNames[idx][1],
      data: data2,
    });
  });
  
  return genExcel(rootDir, outputFilename, sheetData).then(() => {
    console.log(colors.ok(`${outputFile.summary}搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    console.log(`结果文件：${colors.em(path.resolve(rootDir, outputFilename))}`)
  });
}
