const printRunTime = ((start) => (label) => {
  const diff = process.hrtime(start);
  console.log(`${(label || '').padEnd(16, ' ')}${label ? ':' : ' '}run time: ${diff[0]}s${Math.round(diff[1]/1e6)}ms`)
})(process.hrtime());

function lookup() {
  const arr = [];
  for(let i = 0; i < 1e6; i++) {
    arr.push( Math.round(Math.random() * 1e6) );
  }

  console.log('data size:', 1e6);

  printRunTime('gen data');
  arr.sort();
  printRunTime('data sort');

  const lookup = [];

  for(let i = 0; i < 10000; i++) {
    lookup.push( Math.round(Math.random() * 1e6) );
  }

  console.log('lookup data size', 1e4);
  printRunTime('gen lookup data');

  ///////////////////

  for(let i = 0; i < lookup.length; i++) {
    const target = lookup[i];

    let left = 0;
    let right = arr.length - 1;
    let mid;
    do {
      mid = (left + right) >> 1;
      if (arr[mid] === target) break;
      if (arr[mid] < target) {
        left = mid + 1;
      } else {
        right = mid - 1;
      }
    } while(left < right)
  }

  printRunTime('bi search');

  /////////////////

  for(let i = 0; i < lookup.length; i++) {
    arr.includes(lookup[i]);
  }

  printRunTime('lookup');


}

lookup();