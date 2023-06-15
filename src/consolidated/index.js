const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
const { readData } = require('../excel');
require('../colors');

const rootDir = '/Users/bytedance/Documents/财务/合并/4月';

const thisMonth = '2023-04';

const getColumnIndex = col => col.codePointAt(0) - 'A'.codePointAt(0);

// 应收账款 header map
const receivableHM = {
  base: getColumnIndex('A'),
  other: getColumnIndex('E'),
  curMoneyJie: getColumnIndex('F'),
  curMoneyDai: getColumnIndex('G'),
  endMoneyJie: getColumnIndex('L'),
  endMoneyDai: getColumnIndex('M'),
};

// 应付账款 header map
const payableHM = {
  base: getColumnIndex('A'),
  other: getColumnIndex('E'),
  curMoneyJie: getColumnIndex('F'),
  curMoneyDai: getColumnIndex('G'),
  endMoneyJie: getColumnIndex('L'),
  endMoneyDai: getColumnIndex('M'),
};

process();

// 处理
function process() {
  console.log(colors.verbose(`正在处理 ${colors.em(colors.green(thisMonth))} 数据 ...\n文件夹路径: ${colors.em(rootDir)}`));

  const filenames = fs.readdirSync(rootDir);

  const receivables = filenames.filter(n => /其他应收/.test(n));
  const payables = filenames.filter(n => /其他应付/.test(n));

  if (receivables.length === 0 || payables.length === 0) {
    console.log(colors.error('无其他应收、付文件，请检查文件'));
    process.exit(1);
  }

  let receivableList = [];
  let payableList = [];

  // 读取数据
  receivables.forEach(filename => {
    const [{ data }] = readData(rootDir, filename);
    // 前 3 行无用
    data.splice(0, 3);
    data.forEach(it => {
      receivableList.push({
        base: (it[receivableHM.base] ?? '').trim(),
        other: (it[receivableHM.other] ?? '').trim(),
        curMoneyJie: it[receivableHM.curMoneyJie] ?? 0,
        curMoneyDai: it[receivableHM.curMoneyDai] ?? 0,
        endMoneyJie: it[receivableHM.endMoneyJie] ?? 0,
        endMoneyDai: it[receivableHM.endMoneyDai] ?? 0,
      });
    });
  });

  payables.forEach(filename => {
    const [{ data }] = readData(rootDir, filename);
    // 前 3 行无用
    data.splice(0, 3);
    data.forEach(it => {
      payableList.push({
        base: (it[payableHM.base] ?? '').trim(),
        other: (it[payableHM.other] ?? '').trim(),
        curMoneyJie: it[payableHM.curMoneyJie] ?? 0,
        curMoneyDai: it[payableHM.curMoneyDai] ?? 0,
        endMoneyJie: it[payableHM.endMoneyJie] ?? 0,
        endMoneyDai: it[payableHM.endMoneyDai] ?? 0,
      });
    });
  });

  [receivableList, payableList] = cleanAndMergeData(receivableList, payableList);

  // receivableList.forEach((it, idx) => {
  //   console.log(`${idx}: ${it.base} <= ${it.other}, ${it.endMoneyJie}, ${it.endMoneyDai}`);
  // });

  // payableList.forEach((it, idx) => {
  //   console.log(`${idx}: ${it.base} => ${it.other}, ${it.endMoneyJie}, ${it.endMoneyDai}`);
  // });

  const result = [];

  for (const rit of receivableList) {
    const pit = payableList.filter(it => it.base === rit.other && it.other === rit.base)[0];
    const curMoneyJie = rit.curMoneyJie - rit.curMoneyDai;
    const curMoneyDai = pit ? pit.curMoneyDai - pit.curMoneyJie : 0;
    const endMoneyJie = rit.endMoneyJie - rit.endMoneyDai;
    const endMoneyDai = pit ? pit.endMoneyDai - pit.endMoneyJie: 0;
    result.push({
      base: rit.base,
      other: rit.other,
      curMoneyJie,
      curMoneyDai,
      endMoneyJie,
      endMoneyDai,
      diff: endMoneyJie - endMoneyDai,
    });
  }

  // result.forEach((it, idx) => {
  //   console.log(`${idx}: ${it.base} <= ${it.other}, ${it.endMoneyJie}, ${it.endMoneyDai}, ${it.diff}`);
  // });

  genConsolidatedReport(result);
}

function cleanAndMergeData(receivableList, payableList) {
  // 名字清洗
  for (let it of receivableList.concat(payableList)) {
    if (/\(\d+\)$/.test(it.base)) {
      const idx = it.base.lastIndexOf('(');
      it.base = it.base.slice(0, idx);
    }

    if (/\(\d+\)$/.test(it.other)) {
      const idx = it.other.lastIndexOf('(');
      it.other = it.other.slice(0, idx);
    }
  }

  // 无效数据剔除
  const receivableMap = new Map();
  const payableMap = new Map();

  for (let it of payableList) {
    payableMap.set(it.base);
  }
  
  for (let it of receivableList) {
    receivableMap.set(it.base);
  }

  receivableList = receivableList.filter(it => it.base && it.other && payableMap.has(it.other));
  payableList = payableList.filter(it => it.base && it.other && receivableMap.has(it.other));

  // 合并相同公司数据
  const merge = list => list.reduce((a, b) => {
    const it = a.filter(t => t.base === b.base && t.other === b.other)[0];

    if (it) {
      it.curMoneyJie += b.curMoneyJie;
      it.curMoneyDai += b.curMoneyDai;
      it.endMoneyJie += b.endMoneyJie;
      it.endMoneyDai += b.endMoneyDai;
    } else {
      a.push(b);
    }
    return a;
  }, []);

  return [merge(receivableList), merge(payableList)];
}

// 生成合并报表
function genConsolidatedReport(reportData) {
  const data  = [];
  const tHeader = ['base', 'other', '本期发生(借)', '本期发生(贷)', '期末余额(借)', '期末余额(贷)', '期末差额'];

  data.push(tHeader);
  reportData.forEach(it => {
    data.push(
      [it.base, it.other, it.curMoneyJie, it.curMoneyDai, it.endMoneyJie, it.endMoneyDai, it.diff]
    );
  });

  const buffer = xlsx.build([{
    name: '合并',
    data,
  }]);

  const outputFilename = `合并报表.xlsx`;

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, outputFilename)).end(buffer, resolve);
  }).then(() => {
    console.log(colors.ok(`合并报表搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    // console.log(colors.verbose(`输出文件路径: ${outputFilename}`));
  });
}
