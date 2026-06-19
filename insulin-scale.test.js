const rules = `0-200=0
201-250=4
251-300=8
301-350=8
351-400=10
401-=送醫院`;

function advice(value) {
  const sugar = Number(value);
  for (const rule of rules.split(/\n|,/)) {
    const [rawRange, dose] = rule.split('=');
    const [rawLow, rawHigh] = rawRange.split('-');
    const low = rawLow ? Number(rawLow) : -Infinity;
    const high = rawHigh === '' ? Infinity : Number(rawHigh);
    if (sugar >= low && sugar <= high) return dose.trim();
  }
  return '';
}

assert(advice(199) === '0');
assert(advice(250) === '4');
assert(advice(300) === '8');
assert(advice(350) === '8');
assert(advice(400) === '10');
assert(advice(401) === '送醫院');

function assert(ok) {
  if (!ok) throw new Error('insulin scale check failed');
}

console.log('ok');
