const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');
require('../colors');

const rootDir = '/Users/bytedance/Documents/财务/合并/5月';

const thisMonth = '2023-05';

const getColumnIndex = col => col.codePointAt(0) - 'A'.codePointAt(0);

// 应收账款 header map
const receivableHM = {
  base: getColumnIndex('A'),
  other: getColumnIndex('E'),
  moneyJie: getColumnIndex('L'),
  moneyDai: getColumnIndex('M'),
};

// 应付账款 header map
const payableHM = {
  base: getColumnIndex('A'),
  other: getColumnIndex('E'),
  moneyJie: getColumnIndex('L'),
  moneyDai: getColumnIndex('M'),
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

  receivables.forEach(filename => {
    const filePath = path.resolve(rootDir, filename);
    const [{ data }] = xlsx.parse(filePath);
    // 前 3 行无用
    data.splice(0, 3);
    data.forEach(it => {
      receivableList.push({
        base: (it[receivableHM.base] ?? '').trim(),
        other: (it[receivableHM.other] ?? '').trim(),
        moneyJie: it[receivableHM.moneyJie] ?? 0,
        moneyDai: it[receivableHM.moneyDai] ?? 0,
      });
    });
  });

  payables.forEach(filename => {
    const filePath = path.resolve(rootDir, filename);
    const [{ data }] = xlsx.parse(filePath);
    // 前 3 行无用
    data.splice(0, 3);
    data.forEach(it => {
      payableList.push({
        base: (it[receivableHM.base] ?? '').trim(),
        other: (it[receivableHM.other] ?? '').trim(),
        moneyJie: it[payableHM.moneyJie] ?? 0,
        moneyDai: it[payableHM.moneyDai] ?? 0,
      });
    });
  });

  [receivableList, payableList] = mergeData(receivableList, payableList);

  // receivableList.forEach((it, idx) => {
  //   console.log(`${idx}: ${it.base} <= ${it.other}, ${it.moneyJie}, ${it.moneyDai}`);
  // });

  // payableList.forEach((it, idx) => {
  //   console.log(`${idx}: ${it.base} => ${it.other}, ${it.moneyJie}, ${it.moneyDai}`);
  // });

  const result = [];

  for (const rit of receivableList) {
    const pit = payableList.filter(it => it.base === rit.other && it.other === rit.base)[0];
    const jie = rit.moneyJie - rit.moneyDai;
    const dai = pit ? pit.moneyDai - pit.moneyJie: 0;
    result.push({
      base: rit.base,
      other: rit.other,
      moneyJie: jie,
      moneyDai: dai,
      diff: jie - dai,
    });
  }

  // result.forEach((it, idx) => {
  //   console.log(`${idx}: ${it.base} <= ${it.other}, ${it.moneyJie}, ${it.moneyDai}, ${it.diff}`);
  // });

  genConsolidatedReport(result);
}

function mergeData(receivableList, payableList) {
  // clean first
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

  // merge
  const nRec = receivableList.reduce((a, b) => {
    const it = a.filter(t => t.base === b.base && t.other === b.other)[0];

    if (it) {
      it.moneyJie += b.moneyJie;
      it.moneyDai += b.moneyDai;
    } else {
      a.push(b);
    }
    return a;
  }, []);

  const nPay = payableList.reduce((a, b) => {
    const it = a.filter(t => t.base === b.base && t.other === b.other)[0];

    if (it) {
      it.moneyJie += b.moneyJie;
      it.moneyDai += b.moneyDai;
    } else {
      a.push(b);
    }
    return a;
  }, []);

  return [nRec, nPay];
}

// 生成合并报表
function genConsolidatedReport(reportData) {
  const data  = [];
  const tHeader = ['base', 'other', '期末余额(借)', '期末余额(贷)', '差额'];

  data.push(tHeader);
  reportData.forEach(it => {
    data.push(
      [it.base, it.other, it.moneyJie, it.moneyDai, it.diff]
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
