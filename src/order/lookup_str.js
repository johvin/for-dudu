const { updateProgress } = require('../utils');

const printRunTime = ((start) => (label) => {
  const diff = process.hrtime(start);
  console.log(`${(label || '').padEnd(16, ' ')}${label ? ':' : ' '}run time: ${diff[0]}s${Math.round(diff[1]/1e6)}ms`)
})(process.hrtime());

const dataSize = 1e6;
const lookupDataSize = 1e4;
const maxData = 1e8;

function lookup() {
  const arr = [];
  for(let i = 0; i < dataSize; i++) {
    const t = ~~(Math.random() * maxData);
    arr.push('' + t);
  }

  console.log('data size:', dataSize);

  printRunTime('gen data');
  arr.sort((a, b) => {
    if (a.length < b.length) {
      return -1;
    }
    return a < b ? -1 : 1;
  });

  printRunTime('data sort');

  const lookup = [];

  for(let i = 0; i < lookupDataSize; i++) {
    const t = ~~(Math.random() * maxData);
    lookup.push('' + t);
  }

  console.log('lookup data size', lookupDataSize);
  printRunTime('gen lookup data');

  ////////////////////////

  console.log('bi search:');
  console.log();

  for(let i = 0; i < lookup.length; i++) {
    const target = lookup[i];

    let left = 0;
    let right = arr.length - 1;
    let mid;
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
    } while(left <= right);

    updateProgress(`search: ${Math.round(i * 10000 / lookup.length) / 100}%`);
  }

  printRunTime('bi search');

  ////////////////////

  console.log('lookup:');
  console.log();

  for(let i = 0; i < lookup.length; i++) {
    arr.includes(lookup[i]);
    updateProgress(`search: ${Math.round(i * 10000 / lookup.length) / 100}%`);
  }

  printRunTime('lookup');

}

lookup();