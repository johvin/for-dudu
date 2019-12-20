const fs = require('fs');
const path = require('path');
const xlsxSync = require('node-xlsx');
const { XLSX } = require('xlsx-extract');
require('../colors');
const {
  toFixed,
  getYYYYMMDDDateStr,
  getYYYYMMDateStr,
  updateProgress,
} = require('../utils');

const rootDir = '/Users/nilianzhu/Documents/财务/订单/';

const filename1 = '创意云月明细（2018.01-2019.05）.xlsx';
const filename2 = 'CYY-1-2 会员核算订单表内逻辑测试.xlsx';
const file1Header = {
  orderId: 3,
  orderDate: 0,
};
const file2Header = {
  orderId: 2,
  orderDate: 3,
};

const getRunTime = ((start) => () => {
  const diff = process.hrtime(start);
  let min = 0;
  let s = diff[0];
  if (s >= 60) {
    min += Math.floor(s / 60);
    s = s % 60;
  }

  return {
    min,
    s,
    ms: Math.round(diff[1] / 1e6 ),
  };
})(process.hrtime());

const printRunTime = (label) => {
  const t = getRunTime();
  console.log(`run time(${label || ''}): ${t.min > 0 ? `${t.min}min` : ''}${t.s}s${t.ms}ms`);
};

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
  a[b] = toFixed(n, 1, 2) + unit[cnt];
  return a;
}, {});

const printDiffMemory = () => {
  console.log(colors.verbose(`memory diff: ${JSON.stringify(diffMemory(startMemory, process.memoryUsage()))}`));
};

process.on('exit', (exitCode) => {
  console.log('exit', exitCode);
  printRunTime('exit');
  printDiffMemory();
});
process.on('uncaughtException', () => {
  console.log('uncaughtException');
  printRunTime('uncaughtException');
  printDiffMemory();
});

diff(filename1, filename2);

async function diff(filename1, filename2) {
  console.log(colors.verbose(`\n正在比较 "${colors.green(filename1)} & ${colors.green(filename2)}" ...\n`));

  const outputFilename = `${path.basename(filename1, path.extname(filename1))}_based_diff.xlsx`;

  const file1OrderIdMap = await getOrderIdListInFile1(filename1, file1Header);

  for(let [month, arr] of file1OrderIdMap) {
    console.log(month, arr.length);
  }

  // sort for binary search
  for(let [, arr] of file1OrderIdMap) {
    arr.sort((a, b) => {
      if (a.length < b.length) {
        return -1;
      }
      if (a.length > b.length) {
        return 1;
      }
      return a < b ? -1 : 1;
    });
  }
  printRunTime('sort file1OrderList');

  const orderDateRange = Array.from(file1OrderIdMap.keys()).sort();

  const diffs = new Map();

  await diffOrderIdInFile2(filename2, file2Header, (orderId, monthStr) => {
    // file1 中没有的日期不 check
    if (!orderDateRange.includes(monthStr)) return;

    const addDiff = () => {
      if (!diffs.has(monthStr)) {
        diffs.set(monthStr, []);
      }
      diffs.get(monthStr).push([orderId, monthStr]);
    };

    if (!file1OrderIdMap.has(monthStr)) {
      addDiff();
    } else {
      // derived logic
      // const derivedMonthStrArr = ((monthStr, nextCnt) => {
      //   const arr = [ monthStr ];
      //   const date = new Date(monthStr);
      //   for(let i = 0; i < nextCnt; i++) {
      //     date.setMonth(date.getMonth() + 1);
      //     arr.push(`${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`);
      //   }
      //   return arr;
      // })(monthStr, 0);
      // const findedIndex = derivedMonthStrArr.findIndex(monthStr => file1OrderIdMap.has(monthStr) ? file1OrderIdMap.get(monthStr).includes(orderId) : false);
      //
      // if (findedIndex === -1) {
      //   addDiff();
      // }

      (function binarySearch() {
        const arr = file1OrderIdMap.get(monthStr);
        
        let finded = false;
        let left = 0, right = arr.length - 1, mid;

        do {
          mid = (left + right) >> 1;
          if (arr[mid] === orderId) {
            finded = true;
            break;
          }

          if (arr[mid].length === orderId.length) {
            if (arr[mid] < orderId) {
              left = mid + 1;
            } else {
              right = mid - 1;
            }
          } else if (arr[mid].length < orderId.length) {
            left = mid + 1;
          } else {
            right = mid - 1;
          }
        } while (left <= right);

        if (!finded) {
          addDiff();
        }
      })();
    }
  }, (err) => {
    console.log('err', err);
    console.log('current diffs: ', diffs);
  });

  const dateRange = orderDateRange.reduce((a, b) => {
    if (a.length === 0) {
      a.push([b, b]);
    } else {
      const cur = a[a.length - 1];
      const prevDate = new Date(cur[1]);
      prevDate.setMonth(prevDate.getMonth() + 1);
      const curDate = new Date(b);

      if (prevDate.getMonth() === curDate.getMonth()) {
        cur[1] = b;
      } else {
        a.push([b, b]);
      }
    }
    return a;
  }, []);

  console.log('check date range: ', dateRange.map(r => r[0] === r[1] ? r[0] : `${r[0]} - ${r[1]}`).join(', '));

  await outputDiff(diffs, outputFilename);

  // for (let it of diffs) {
  //   await outputDiff([it], outputFilename.replace('.xlsx', `_${it[0]}.xlsx`));
  // }
}

function getOrderIdListInFile1(filename, header) {
  const filePath = path.resolve(rootDir, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  console.log(colors.verbose(`\n正在读取 ${filename} 数据 ...\n`));
  printDiffMemory();

  return new Promise((res, rej) => {
    const data = new Map();

    let lines = 0;
    new XLSX().extract(filePath, { sheet_nr: 1, ignore_header: 1, convert_values: { dates: false } })
    .on('sheet', function (sheet) {
      console.log('sheet', sheet[0]);  //sheet is array [sheetname, sheetid, sheetnr]
      printRunTime('file1 sheet')
      console.log();
    })
    .on('row', function (row) {
      try {
        if (++lines % 10000 === 0) {
          const t = getRunTime();
          updateProgress(`progress: row ${lines} (${t.min > 0 ? `${t.min}min` : ''}${t.s}s${t.ms}ms)`);
        }

        // if (lines < 10) {
        //   console.log('row', row);
        // }

        let monthStr = '';

        try {
          monthStr = getYYYYMMDDDateStr(row[ header.orderDate ]).slice(0, 7);
        } catch(e) {
          monthStr = getYYYYMMDateStr(row[ header.orderDate ]);
        }

        if (!data.has(monthStr)) {
          data.set(monthStr, []);
        }

        data.get(monthStr).push('' + row[ header.orderId ]);
      } catch(e) {
        console.log(`lines ${lines} error`);
        console.log(e);
        process.exit(1);
      }
    })
    .on('error', (err) => {
      console.log(data.keys());
      console.log(err);
      rej(err);
      process.exit(1);
    })
    .on('end', function (err) {
      console.log('eof total', lines);
      
      printRunTime('read end');
      printDiffMemory();
      res(data);
    });
  });
}

function diffOrderIdInFile2(filename, header, diffFn, errorHandler) {
  const filePath = path.resolve(rootDir, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  console.log(colors.verbose(`\n正在 diff ${filename} 数据 ...\n`));
  printDiffMemory();

  return new Promise((res, rej) => {
    let lines = 0;
    new XLSX().extract(filePath, { sheet_nr: 1, ignore_header: 1, convert_values: { dates: false } })
    .on('sheet', function (sheet) {
      console.log('sheet', sheet[0]);  //sheet is array [sheetname, sheetid, sheetnr]
      printRunTime('file2 sheet');
      console.log();
    })
    .on('row', function (row) {
      try {
        lines++;

        if (lines % 10000 === 0) {
          const t = getRunTime();
          updateProgress(`progress: row ${lines} (${t.min > 0 ? `${t.min}min` : ''}${t.s}s${t.ms}ms)`);
        }

        let monthStr = '';

        try {
          monthStr = getYYYYMMDDDateStr(row[ header.orderDate ]).slice(0, 7);
        } catch(e) {
          monthStr = getYYYYMMDateStr(row[ header.orderDate ]);
        }

        if (diffFn) {
          diffFn('' + row[ header.orderId ], monthStr);
        }
      } catch(e) {
        console.log(`lines ${lines} error`);
        console.log(e);
        if (errorHandler) {
          errorHandler(e);
        }
      }
    })
    .on('error', (err) => {
      console.log(err);
      rej(err);
      process.exit(1);
    })
    .on('end', function (err) {
      console.log('eof total', lines);
      printRunTime('read end');
      printDiffMemory();
      res();
    });
  });
}

function outputDiff(diffData, outputFilename) {
  const sheets = [];
  console.log('diff data:');
  let total = 0;
  for(let [monthStr, arr] of diffData) {
    console.log(`${monthStr}: ${arr.length}`);
    total += arr.length;
    sheets.push({
      name: monthStr,
      data: arr,
    });
    if (arr.length < 10) {
      console.log(JSON.stringify(arr));
    }
  }
  console.log(`total diff count: ${total}`);

  const buffer = xlsxSync.build(sheets);

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, outputFilename)).end(buffer, resolve);
  }).then(() => {
    console.log(colors.ok(`diff 搞定 ✌️️️️️✌️️️️️✌️️️️️`));
    console.log(colors.verbose(`输出文件路径: ${outputFilename}`));
  });
}
