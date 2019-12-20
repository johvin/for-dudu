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

exports.toFixed = toFixed;
exports.parseYYYYMMDDFromExcelDateNumber = parseYYYYMMDDFromExcelDateNumber;
exports.getYYYYMMDDDateStr = getYYYYMMDDDateStr;
exports.getYYYYMMDateStr = getYYYYMMDateStr;
exports.updateProgress = updateProgress;
