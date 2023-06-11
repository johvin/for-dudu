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
  strBinarySearch,
} = require('../utils');

const rootDir = '/Users/nilianzhu/Documents/财务/订单/';

const filename1 = '创意云月明细（2018.01-2019.05）.xlsx';
const filename2 = 'CYY-1-2 会员核算订单表内逻辑测试.xlsx';

// file2 是否有订单日期
const file2HasOrderDate = true;
const file1Header = {
  orderId: 3,
  orderDate: 0,
};
const file2Header = {
  orderId: 2,
  orderDate: 3,
  orderStartDate: 0,
};

const getRunTime = ((init) => (start) => {
  const diff = process.hrtime(start || init);
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
process.on('uncaughtException', (err) => {
  console.log('uncaughtException', err);
  printRunTime('uncaughtException');
  printDiffMemory();
});

diff(filename1, filename2);

async function diff(filename1, filename2) {
  console.log(colors.verbose(`\n正在比较 "${colors.green(filename1)} & ${colors.green(filename2)}" ...\n`));

  const outputFilename = `${path.basename(filename1, path.extname(filename1))}_based_diff.xlsx`;

  const file1OrderIdMap = await getOrderIdListInFile1(filename1, file1Header);

  const orderMonthRange = Array.from(file1OrderIdMap.keys()).sort();

  for(let month of orderMonthRange) {
    console.log(month, file1OrderIdMap.get(month).length);
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

  const continousDateRange = orderMonthRange.reduce((a, b) => {
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

  const diffRange = continousDateRange.map(r => r[0] === r[1] ? r[0] : `${r[0]} - ${r[1]}`).join(', ');

  const diffs = new Map();

  // 用来打印平均搜索次数
  const derivedSearchTimes = [];

  await diffOrderIdInFile2(filename2, (row, lineNo, header) => {
    let orderDateStr = '';
    let orderMonthStr = '';
    let orderStartDateStr = '';
    let orderStartMonthStr = '';

    try {
      orderDateStr = getYYYYMMDDDateStr(row[ file2Header.orderDate ]);
    } catch(e) {
      orderDateStr = getYYYYMMDateStr(row[ file2Header.orderDate ]);
      if (orderDateStr.length === 6) {
        orderDateStr += 'xx';
      } else {
        orderDateStr += orderDateStr[4] + 'xx';
      }
    }

    orderMonthStr = /\d/.test(orderDateStr[4]) ? orderDateStr.slice(0, 6) : orderDateStr.slice(0, 7);

    if (!file2HasOrderDate) {
      try {
        orderStartDateStr = getYYYYMMDDDateStr(row[ file2Header.orderStartDate ]);
      } catch(e) {
        orderStartDateStr = getYYYYMMDateStr(row[ file2Header.orderStartDate ]);
        if (orderStartDateStr.length === 6) {
          orderStartDateStr += 'xx';
        } else {
          orderStartDateStr += orderStartDateStr[4] + 'xx';
        }
      }

      orderStartMonthStr = /\d/.test(orderStartDateStr[4]) ? orderStartDateStr.slice(0, 6) : orderStartDateStr.slice(0, 7);
    }

    // file1 中没有的日期不 check
    if (file2HasOrderDate && !orderMonthRange.includes(orderMonthStr)) return;

    const orderId = '' + row[ file2Header.orderId ];

    const addDiff = () => {
      if (!diffs.has(orderMonthStr)) {
        diffs.set(orderMonthStr, [ header ]);
      }
      diffs.get(orderMonthStr).push(row);
    };

    // 只按照当前月份对比
    if (file2HasOrderDate) {
      if (!file1OrderIdMap.has(orderMonthStr) || strBinarySearch(file1OrderIdMap.get(orderMonthStr), orderId) === -1) {
        addDiff();
      }
      return;
    } else {
      // derived month from earliest month,
      // 月末几天的订单在另一个表中可能会被统计到下个月，因此需要 nextCnt = 1
      const derivedMonthStrArr = ((curMonthStr, nextCnt) => {
        const arr = [];
        const earliestMonth = orderMonthRange[0];

        const date = new Date(curMonthStr);
        if (nextCnt >= 0) {
          // 先查询当前月，提高效率
          arr.push( curMonthStr );
        }
        if (nextCnt > 0) {
          date.setMonth(date.getMonth() + nextCnt);
        }

        do {
          const cur = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
          if (cur < earliestMonth) {
            break;
          }

          if (nextCnt > 0) {
            if (!arr.includes(cur)) {
              arr.push(cur);
            }
          } else {
            arr.push(cur);
          }
          date.setMonth(date.getMonth() - 1);
        } while(true);

        return arr;
      })(orderStartMonthStr, 1);

      const findedIndex = derivedMonthStrArr.findIndex(
        monthStr => file1OrderIdMap.has(monthStr) ? strBinarySearch(file1OrderIdMap.get(monthStr), orderId) !== -1 : false
      );

      if (findedIndex === -1) {
        addDiff();
      } else {
        derivedSearchTimes.push(findedIndex + 1);
      }
    }

  }, (err) => {
    console.log('err', err);
    console.log('current diffs: ', diffs);
  });

  if (derivedSearchTimes.length > 0) {
    console.log('average derived search times:', toFixed(derivedSearchTimes.reduce((a, b) => a + b, 0), derivedSearchTimes.length, 2));
    // console.log(derivedSearchTimes.slice(0, 20))
  }

  // console.log('check date range: ', diffRange);

  if (diffs.size > 0) {
    const diffArr = [
      // sheet name, sheet data
      ['说明', [ ['检查时间范围'], [diffRange] ]],
    ];
    for(let m of orderMonthRange) {
      if (diffs.has(m)) {
        diffArr.push([m, diffs.get(m)]);
      }
    }
    await outputDiff(diffArr, outputFilename);
  } else {
    console.log(colors.ok('数据无差异，👍'));
  }

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
    new XLSX().extract(filePath, { parser: 'expat', sheet_nr: 1, ignore_header: 1, convert_values: { dates: false } })
    .on('sheet', function (sheet) {
      console.log('sheet', sheet[0]);  //sheet is array [sheetname, sheetid, sheetnr]
      printRunTime('file1 sheet')
      console.log();
    })
    .on('row', function (row) {
      try {
        if (++lines % 1000 === 0) {
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

function diffOrderIdInFile2(filename, rowHandler, errorHandler) {
  const filePath = path.resolve(rootDir, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  console.log(colors.verbose(`\n正在 diff ${filename} 数据 ...\n`));
  printDiffMemory();

  return new Promise((res, rej) => {
    let tableHeader;
    let lines = 0;
    new XLSX().extract(filePath, { parser: 'expat', sheet_nr: 1, ignore_header: 0, convert_values: { dates: false } })
    .on('sheet', function (sheet) {
      console.log('sheet', sheet[0]);  //sheet is array [sheetname, sheetid, sheetnr]
      printRunTime('file2 sheet');
      console.log();
    })
    .on('row', function (row) {
      try {
        if (!tableHeader) {
          tableHeader = row;
          return;
        }

        if (++lines % 1000 === 0) {
          const t = getRunTime();
          updateProgress(`progress: row ${lines} (${t.min > 0 ? `${t.min}min` : ''}${t.s}s${t.ms}ms)`);
        }

        if (rowHandler) {
          rowHandler(row, lines, tableHeader);
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
      console.log(JSON.stringify(arr.slice(0, 4)));
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
