const fs = require('fs');
const path = require('path');
const { rootDir, thisMonth } = require('./constants');
const { readData, genExcel } = require('../excel');
require('../colors');

const resultItem = {
  property: `流动资产： 
  货币资金 
  应收票据 
  应收账款 
  预付款项 
  应收利息 
  其他应收款 
  存货 
  一年内到期的非流动资产 
  其他流动资产 
流动资产合计 
非流动资产： 
  长期应收款 
  长期股权投资 
  固定资产 
  在建工程 
  固定资产清理 
  无形资产 
  开发支出 
  商誉 
  长期待摊费用 
  其他非流动资产 
非流动资产合计 
资产总计`.split(/\n/).map(it => it.trimEnd()),
  debtEquity: `
  负债及所有者权益（或股东权益） 
  流动负债： 
    短期借款 
    应付票据 
    应付账款 
    预收款项 
    应付职工薪酬 
    应交税费 
    应付利息 
    其他应付款 
    一年内到期的非流动负债 
   其他流动负债 
  流动负债合计 
    非流动负债： 
    长期借款 
    应付债券 
    长期应付款 
    其他非流动负债 
  非流动负债合计 
  负债合计 
  所有者权益（或股东权益）： 
    实收资本（或股本） 
    资本公积 
    盈余公积 
    未分配利润 
    所有者权益（或股东权益）合计 
  负债和所有者权益（或股东权益）总计`.split(/\n/).map(it => it.trimEnd()),
  profit: `
  一、营业收入
  　　减：营业成本
  　    　营业税金及附加
      　　销售费用
  　　    管理费用
  　　    财务费用
  　    　资产减值损失
  　　加：公允价值变动收益)损失以"-"号填列)
  　　    投资收益(损失以"-"号填列)
  　　　  其中：对联营企业和合营企业的投资收益
  二、营业利润(亏损以"-"号填列)
  　　加：营业外收入
  　　减：营业外支出
  　　  　其中：非流动资产处置损失
         加以前年度损益调整
  三、利润总额(亏损总额以"-"号填列)
  　　减：所得税费用
  四、净利润(净亏损以"-"号填列)`.split(/\n/).map(it => it.trimEnd()),
  cashFlow: `一、经营活动产生的现金流量：
  　　销售商品、提供劳务收到的现金
  　　收到的税费返还
  　　收到的其他与经营活动有关的资金
  　   经营活动现金流入小计
  　　购买商品、接受劳务支付的现金
  　　支付给职工以及为职工支付的现金
  　　支付的各项税费
  　　支付其他与经营活动有关的现金
  　   经营活动现金流出小计
  　　经营活动产生的现金流量净额
  二、投资活动产生的现金流量：
  　　收回投资收到的现金
  　　取得投资收益收到的现金
  　　处置固定资产、无形资产和其他长期资产收回的现金净额
  　　处置子公司及其他营业单位收到的现金净额
  　　收到其他与投资活动有关的现金
  　   投资活动现金流入小计
  　　购建固定资产、无形资产和其他长期资产所支付的现金
  　　投资支付的现金
  　　取得子公司及其他营业单位支付的现金净额
  　　支付其他与投资活动有关的现金
  　   投资活动现金流出小计
  　　投资活动产生的现金流量净额
  三、筹资活动所产生的现金流量：
  　　吸收投资收到的现金
  　　取得借款收到的现金
  　　收到其他与筹资活动有关的现金
  　   筹资活动现金流入小计
  　　偿还债务支付的现金
  　　分配股利、利润或偿付利息所支付的现金
  　　支付其他与筹资活动有关的现金
  　   筹资活动现金流出小计
  　　筹资活动产生的现金流量净额
  四、汇率变动对现金及现金等价物的影响
  五、现金及现金等价物净增加额
  　　加：期初现金及现金等价物余额
  六、期末现金及现金等价物余额`.split(/\n/).map(it => it.trimEnd()),
};

// console.log(JSON.stringify(resultItem.cashFlow, null, 2));

process();

// 处理
function process() {
  const sourceDir = path.resolve(rootDir, 'summary');
  console.log(colors.verbose(`正在处理 ${colors.em(colors.green(thisMonth))} 数据 ...\n源文件夹路径: ${colors.em(sourceDir)}`));

  const filenames = fs.readdirSync(sourceDir);

  const inputFiles = filenames.filter(n => /^\d+_/.test(n)).sort();

  if (inputFiles.length === 0) {
    console.log(colors.error('无其他应收、付文件，请检查文件'));
    process.exit(1);
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

  for (const item of items) {
    const row1 = [item];
    const row2 = [item];

    const matchItemName = item.trimStart();
    for (const c of companyList) {
      const info = c.list.find(it => it[0] === matchItemName);

      row1.push(info?.[1]);
      row2.push(info?.[2]);
    }
    data1.push(row1);
    data2.push(row2);
  }

  return [sheetHeaders, data1, data2];
}

// 生成合并报表
function genSummaryReport(...summarys) {
  const outputFilename = `汇总表.xlsx`;
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
    console.log(colors.ok(`汇总报表搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    console.log(`结果文件：${colors.em(path.resolve(rootDir, outputFilename))}`)
  });
}
