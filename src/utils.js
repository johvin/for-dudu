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

// 获取 YYYY-MM-DD 格式（标准格式）的日期字符串
function getYYYYMMDDDateStr(dStr) {
  const getYMDByDate = d => `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;

  if (typeof dStr === 'string') {
    if (/^\d{4}-\d{2}-\d{2}/.test(dStr)) {
      return dStr.slice(0, 10);
    }
    if (/^\d{4}.\d{1,2}.\d{1,2}/.test(dStr)) {
      return dStr.replace(/^(\d{4}).(\d{1,2}).(\d{1,2})/, function (s0, s1, s2, s3) {
        return `${s1}-${s2.padStart(2, '0')}-${s3.padStart(2, '0')}`;
      }).slice(0, 10);
    }
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
};

function updateProgress(txt) {
  const rl = require('readline');
  rl.moveCursor(process.stdout, 0, -1);
  rl.clearLine(process.stdout, 0);
  console.log(txt);
}

exports.toFixed = toFixed;
exports.parseYYYYMMDDFromExcelDateNumber = parseYYYYMMDDFromExcelDateNumber;
exports.getYYYYMMDDDateStr = getYYYYMMDDDateStr;
exports.updateProgress = updateProgress;
