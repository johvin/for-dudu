const path = require('path');
const xlsx = require('node-xlsx');

function readData(rootDir, filename) {
  const filePath = path.resolve(rootDir, filename);
  // [sheet1, sheet2, ...] => [{ data1 }, { data2 }, ...]
  return xlsx.parse(filePath);
}

exports.readData = readData;