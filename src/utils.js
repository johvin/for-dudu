// 数字格式化小数位数
function toFixed(dividend, divisor, n = 2) {
  const weight = 10 ** n;
  return Math.round(dividend * weight / divisor) / weight;
}

// 将数字格式的 excel date 转换成 js 中的年月日
// n 表示距离 1900-01-00 的天数
function parseYYYYMMDDFromExcelDateNumber(n) {
  if (typeof n !== 'number' || isNaN(n) || n < 0) {
    throw new Error('invalid Excel Date Number "' + n + '"');
  }
  n = parseInt(n);

  if (n === 0) {
    return '1900-01-00';
  }
  const d = new Date('1900-01-01T00:00:00.000Z');
  // excel 1900/1/1 是第1天
  d.setDate(n - 1);

  return d.toISOString().slice(0, 10);
}

// 20180212 格式
const dateStrRe1 = /^\d{4}\d{2}\d{2}(?=\b|[^\d]|\s|$)/;
// 2018/8/12 格式
const dateStrRe2 = /^(\d{4})(-|\/)(\d{1,2})\2(\d{1,2})(?=\b|[^\d]|\s|$)/;

// 获取 YYYY-MM-DD 格式（标准格式）的日期字符串
function getYYYYMMDDDateStr(dStr) {
  const getYMDByDate = d => `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;

  if (typeof dStr === 'string') {
    if (dateStrRe1.test(dStr)) {
      return dStr.slice(0, 8);
    }
    if (dateStrRe2.test(dStr)) {
      return dStr.replace(dateStrRe2, function (s0, s1, s2, s3, s4) {
        return `${s1}${s2}${s3.padStart(2, '0')}${s2}${s4.padStart(2, '0')}`;
      }).slice(0, 10);
    }
    // 时间戳 or excel 天数
    if (/^\d+$/.test(dStr)) {
      const d = Number(dStr);
      return d > 1e8 ? getYMDByDate(new Date(d)) : parseYYYYMMDDFromExcelDateNumber(d);
    }
    throw new Error('invalid date string: ' + dStr);
  } else if (typeof dStr === 'number') {
    return parseYYYYMMDDFromExcelDateNumber(dStr);
  } else if (typeof dStr === 'object' && dStr && dStr.constructor === Date) {
    return getYMDByDate(dStr);
  } else {
    throw new Error('invalid date string or date number: ' + dStr);
  }
}

const monthStrRe1 = /^\d{4}\d{2}(?=\b|[^\d]|\s|$)/;
const monthStrRe2 = /^(\d{4})(-|\/)(\d{1,2})(?=\b|[^\d]|s|$)/;

function getYYYYMMDateStr(input) {
  if (typeof input === 'string') {
    if (monthStrRe1.test(input)) {
      return input.slice(0, 6);
    }
    if (monthStrRe2.test(input)) {
      return input.replace(monthStrRe2, function (s0, s1, s2, s3) {
        return `${s1}${s2}${s3.padStart(2, '0')}`;
      }).slice(0, 7);
    }
  }
  throw new Error('invalid month string: ' + input);
}

function updateProgress(txt) {
  const rl = require('readline');
  rl.moveCursor(process.stdout, 0, -1);
  rl.clearLine(process.stdout, 0);
  console.log(txt);
}

function strBinarySearch(arr, target) {
  let left = 0, right = arr.length - 1, mid;

  do {
    mid = (left + right) >> 1;
    if (arr[mid] === target) {
      break;
    }

    if (arr[mid].length === target.length) {
      if (arr[mid] < target) {
        left = mid + 1;
      } else {
        right = mid - 1;
      }
    } else if (arr[mid].length < target.length) {
      left = mid + 1;
    } else {
      right = mid - 1;
    }
  } while (left <= right);

  return left > right ? -1 : mid;
}

function readFile(filePath, rowHandler) {
  const fs = require('fs');
  const path = require('path');
  const { XLSX } = require('xlsx-extract');

  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在 => ${filePath}`);
  }

  console.log(colors.verbose(`\n正在读取 ${path.basename(filePath)} 数据 ...\n`));

  return new Promise((res, rej) => {
    let curSheet;
    let header;
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
            monthStr = getYYYYMMDDDateStr(row[header.orderDate]).slice(0, 7);
          } catch (e) {
            monthStr = getYYYYMMDateStr(row[header.orderDate]);
          }

          if (!data.has(monthStr)) {
            data.set(monthStr, []);
          }

          data.get(monthStr).push('' + row[header.orderId]);
        } catch (e) {
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

exports.toFixed = toFixed;
exports.parseYYYYMMDDFromExcelDateNumber = parseYYYYMMDDFromExcelDateNumber;
exports.getYYYYMMDDDateStr = getYYYYMMDDDateStr;
exports.getYYYYMMDateStr = getYYYYMMDateStr;
exports.updateProgress = updateProgress;
exports.strBinarySearch = strBinarySearch;
