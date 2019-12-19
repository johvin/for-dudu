const { updateProgress } = require('../utils');

const printRunTime = ((start) => (label) => {
  const diff = process.hrtime(start);
  console.log(`${(label || '').padEnd(16, ' ')}${label ? ':' : ' '}run time: ${diff[0]}s${Math.round(diff[1]/1e6)}ms`)
})(process.hrtime());

const dataSize = 1e2;
const lookupDataSize = 1e1;
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
    if (a.length > b.length) {
      return 1;
    }
    return a < b ? -1 : 1;
  });

  printRunTime('data sort');

  console.log(JSON.stringify(arr, null, 2));

  const lookup = [];

  for(let i = 0; i < lookupDataSize; i++) {
    lookup.push( arr[ Math.floor( Math.random() * dataSize ) ] );
  }

  console.log('lookup data size', lookupDataSize);
  printRunTime('gen lookup data');

  console.log(JSON.stringify(lookup, null, 2));

  ////////////////////////

  console.log('bi search:');
  console.log();
  let find1 = 0;

  for(let i = 0; i < lookup.length; i++) {
    const target = lookup[i];

    let left = 0;
    let right = arr.length - 1;
    let mid;
    do {
      mid = (left + right) >> 1;
      if (arr[mid] === target) {
        find1++;
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
  console.log('find', find1);

  ////////////////////

  console.log('lookup:');
  console.log();
  let find2 = 0;

  for(let i = 0; i < lookup.length; i++) {
    if (arr.includes(lookup[i])) {
      find2++;
    }
    updateProgress(`search: ${Math.round(i * 10000 / lookup.length) / 100}%`);
  }

  printRunTime('lookup');
  console.log('find', find2);

}

lookup();