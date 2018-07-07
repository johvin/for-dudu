function toFixed(dividend, divisor, n = 2) {
  const weight = 10 ** n;
  return Math.round(dividend * weight / divisor) / weight;
}

// 将数字格式的 excel date 转换成 js 中的年月日
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
  switch (typeof dStr) {
    case 'string':
      if (/^\d{4}-\d{2}-\d{2}$/.test(dStr)) {
        return dStr;
      }
      if (/^\d{4}.\d{1,2}.\d{1,2}$/.test(dStr)) {
        return dStr.replace(/^(\d{4}).(\d{1,2}).(\d{1,2})$/, function (s0, s1, s2, s3) {
          return `${s1}-${s2.length === 1 ? '0' + s2 : s2}-${s3.length === 1 ? '0' + s3 : s3}`;
        });
      }
      throw new Error('invalid date string: ' + dStr);
    case 'number':
      return parseYYYYMMDDFromExcelDateNumber(dStr);
    default:
      throw new Error('invalid date string or date number: ' + dStr);
  }
};

exports.toFixed = toFixed;
exports.parseYYYYMMDDFromExcelDateNumber = parseYYYYMMDDFromExcelDateNumber;
exports.getYYYYMMDDDateStr = getYYYYMMDDDateStr;
