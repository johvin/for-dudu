const fs = require('fs');
const path = require('path');
const xlsxSync = require('node-xlsx');
const { XLSX } = require('xlsx-extract');
require('../colors');

const rootDir = '/Users/johvin/Documents/财务/订单/';

const filename1 = '附件：2015-2017秀点订单.xlsx';
const filename2 = '2017年及之前其他年份的秀点收入核算表.xlsx';

const startMemory = process.memoryUsage();
const diffMemory = (m1, m2) => Object.keys(m2).reduce((a, b) => {
  let n = m2[b] - m1[b];
  const unit = ['B', 'KB', 'MB'];
  let cnt = 0;
  while (n > 1024) {
    n /= 1024;
    if (++cnt == unit.length) {
      break;
    }
  }
  a[b] = Math.round(n * 100) / 100 + unit[cnt];
  return a;
}, {});

const printDiffMemory = () => {
  console.log(colors.verbose(`memory diff: ${JSON.stringify(diffMemory(startMemory, process.memoryUsage()))}`));
};

// process.on('exit', () => printDiffMemory());
process.on('exit', () => {
  console.log('exit');
});
process.on('uncaughtException', () => printDiffMemory());

diff(filename1, filename2);

async function diff(filename1, filename2) {
  console.log(colors.verbose(`\n正在比较 "${colors.green(filename1)} & ${colors.green(filename2)}" ...\n`));

  const outputFilename = `${path.basename(filename1, path.extname(filename1))}_based_diff.xlsx`;

  const file1OrderIdList = await getOrderIdList(filename1, 0);
  const file2OrderIdList = await getOrderIdList(filename2, 0);

  const diff = file2OrderIdList.filter(id => !file1OrderIdList.includes(id));

  output(diff, outputFilename);
  printDiffMemory();
}

function getOrderIdList(filename, orderIdIndex) {
  const filePath = path.resolve(rootDir, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  console.log(colors.verbose(`\n正在读取 ${filename} 数据 ...\n`));
  printDiffMemory();

  return new Promise((res, rej) => {
    const data = [];
    new XLSX().extract(filePath, { sheet_nr: 1 })
    .on('sheet', function (sheet) {
      console.log('sheet', sheet[0]);  //sheet is array [sheetname, sheetid, sheetnr]
    })
    .on('row', function (row) {
      data.push(row[orderIdIndex]);
      if (data.length === 1) {
        console.log(row);
        throw 'a';
      }
      if (data.length % 10000 === 0) {
        console.log('row ', data.length);
        console.log('row', row)
      }
    })
    .on('error', rej)
    .on('end', function (err) {
      res(data);
      console.log('eof total', data.length);
      printDiffMemory();
    });
  });
}

function getOrderIdList2(filename, orderIdIndex) {
  const filePath = path.resolve(rootDir, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  console.log(colors.verbose(`\n正在读取 ${filename} 数据 ...\n`));
  printDiffMemory();

  const [{ data: orderList }] = xlsxSync.parse(filePath);
  orderList.shift();

  return orderList.map(it => it[orderIdIndex]);
}

function output(data, outputFilename) {
  const buffer = xlsxSync.build([{
    name: 'diff',
    data,
  }]);

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, outputFilename)).end(buffer, resolve);
  }).then(() => {
    console.log(colors.ok(`diff 搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    console.log(colors.verbose(`输出文件路径: ${outputFilename}`));
  });
}
