const fs = require('fs');
const path = require('path');
const xlsx = require('node-xlsx');

/** 读取 excel data */
function readData(rootDir, filename) {
  const filePath = path.resolve(rootDir, filename);
  // [sheet1, sheet2, ...] => [{ data1 }, { data2 }, ...]
  return xlsx.parse(filePath);
}

/** 读取小型纯文本 */
function readText(rootDir, filename) {
  const filePath = path.resolve(rootDir, filename);
  return fs.readFileSync(filePath, { encoding: 'utf-8'});
}


/**
 * 生成文件表
 * @param {*} rootDir 
 * @param {*} outFilename 
 * @param {*} sheetList ItemType => { name: string, data: Array }
 * @returns 
 */
function genExcel(rootDir, outFilename, sheetList) {
  const buffer = xlsx.build(sheetList);

  return new Promise((resolve) => {
    fs.createWriteStream(path.resolve(rootDir, outFilename)).end(buffer, resolve);
  });
}

exports.readData = readData;
exports.readText = readText;
exports.genExcel = genExcel;
