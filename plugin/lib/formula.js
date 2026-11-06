// A small spreadsheet formula engine: enough of Excel's grammar and the
// everyday functions to give written formulas a cached value, so reading
// and previewing show results. Anything it does not know throws
// Unsupported, and the cell is left for Excel/WPS to calculate on open.

export class Unsupported extends Error {}

export const ERR = (code) => ({ error: code })
export const isErr = (v) => v && typeof v === 'object' && 'error' in v

// ------------------------------------------------------------ addresses

export function colNum (letters) {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64
  return n
}

export function colName (n) {
  let s = ''
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) }
  return s
}

export function parseCell (ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref).trim())
  if (!m) return undefined
  return { r: +m[2], c: colNum(m[1]) }
}

export const cellName = (r, c) => colName(c) + r

export function parseRange (ref) {
  const [a, b] = String(ref).trim().split(':')
  const p = parseCell(a)
  const q = b === undefined ? p : parseCell(b)
  if (!p || !q) return undefined
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) }
}

// ------------------------------------------------------------ tokenizer

const TOKEN = new RegExp([
  /(?<ws>\s+)/,
  /(?<str>"(?:[^"]|"")*")/,
  /(?<err>#(?:DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!))/,
  // Optional sheet, then a cell, a cell range, a column range or a row range.
  /(?<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*)!)?(?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(]))/,
  /(?<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)/,
  /(?<fn>[A-Za-z_][\w.]*(?=\())/,
  /(?<name>[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*)/,
  /(?<op><>|<=|>=|[-+*/^&=<>%(),:;{}])/
].map((r) => r.source).join('|'), 'y')

export function tokenize (src) {
  const out = []
  TOKEN.lastIndex = 0
  while (TOKEN.lastIndex < src.length) {
    const at = TOKEN.lastIndex
    const m = TOKEN.exec(src)
    if (!m || m.index !== at) throw new Unsupported(`cannot read formula at "${src.slice(at, at + 12)}"`)
    const g = m.groups
    const kind = Object.keys(g).find((k) => g[k] !== undefined)
    out.push({ kind, text: m[0] })
  }
  return out
}

// shift moves the relative parts of the references in a formula by dr
// rows and dc columns (shared formulas store one master for many cells).
export function shift (src, dr, dc) {
  return tokenize(src).map((t) => {
    if (t.kind !== 'ref') return t.text
    const bang = t.text.lastIndexOf('!')
    const sheet = bang >= 0 ? t.text.slice(0, bang + 1) : ''
    const body = bang >= 0 ? t.text.slice(bang + 1) : t.text
    return sheet + body.replace(/(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?/g, (all, ca, col, ra, row) => {
      if (!all) return all
      let s = ''
      if (col) s += ca + (ca ? col : colName(colNum(col) + dc))
      if (row) s += ra + (ra ? row : String(+row + dr))
      return s
    })
  }).join('')
}

// ------------------------------------------------------------ parser

export function parseFormula (src) {
  const toks = tokenize(src.replace(/^=/, '')).filter((t) => t.kind !== 'ws')
  let i = 0
  const peek = () => toks[i]
  const next = () => toks[i++]
  const isOp = (t, ...ops) => t && t.kind === 'op' && ops.includes(t.text)
  const binary = (sub, ops) => () => {
    let left = sub()
    while (isOp(peek(), ...ops)) {
      const op = next().text
      left = { op, args: [left, sub()] }
    }
    return left
  }
  const primary = () => {
    const t = next()
    if (!t) throw new Unsupported('formula ends early')
    switch (t.kind) {
      case 'num': return { lit: parseFloat(t.text) }
      case 'str': return { lit: t.text.slice(1, -1).replace(/""/g, '"') }
      case 'err': return { lit: ERR(t.text) }
      case 'ref': return { ref: t.text }
      case 'name': {
        const u = t.text.toUpperCase()
        if (u === 'TRUE' || u === 'FALSE') return { lit: u === 'TRUE' }
        return { name: t.text }
      }
      case 'fn': {
        next() // (
        const args = []
        if (!isOp(peek(), ')')) {
          for (;;) {
            if (isOp(peek(), ',', ';')) args.push({ lit: null })
            else args.push(expr())
            if (isOp(peek(), ',', ';')) { next(); if (isOp(peek(), ')')) args.push({ lit: null }); continue }
            break
          }
        }
        if (!isOp(next(), ')')) throw new Unsupported('missing )')
        return { fn: t.text.toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, ''), args }
      }
      case 'op':
        if (t.text === '(') { const e = expr(); if (!isOp(next(), ')')) throw new Unsupported('missing )'); return e }
        if (t.text === '-' || t.text === '+') return { op: 'u' + t.text, args: [unary()] }
        if (t.text === '{') throw new Unsupported('array constants')
    }
    throw new Unsupported(`unexpected "${t.text}"`)
  }
  const postfix = () => {
    let e = primary()
    while (isOp(peek(), '%')) { next(); e = { op: '%', args: [e] } }
    return e
  }
  const unary = () => {
    if (isOp(peek(), '-', '+')) { const op = next().text; return { op: 'u' + op, args: [unary()] } }
    return postfix()
  }
  const power = binary(unary, ['^'])
  const term = binary(power, ['*', '/'])
  const additive = binary(term, ['+', '-'])
  const concat = binary(additive, ['&'])
  const expr = binary(concat, ['=', '<>', '<', '>', '<=', '>='])
  const tree = expr()
  if (i < toks.length) throw new Unsupported(`unexpected "${toks[i].text}"`)
  return tree
}

// ------------------------------------------------------------ evaluation

// env: { sheet, cell(sheet, r, c) → value, bounds(sheet) → { r2, c2 },
// has(sheet), name(text) → formula source of a defined name, at: { r, c }
// the cell being calculated }
export function evaluate (tree, env) {
  const v = ev(tree, env)
  if (v && v.range) return v.values.length && v.values[0].length ? scalarOf(v.values[0][0]) : 0
  return v === null || v === undefined ? 0 : v
}

function resolveRef (text, env) {
  const bang = text.lastIndexOf('!')
  let sheet = env.sheet
  let body = text
  if (bang >= 0) {
    sheet = text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
    body = text.slice(bang + 1)
  }
  body = body.replace(/\$/g, '')
  let rg
  if (/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(body)) {
    const [a, b] = body.split(':')
    const bd = env.bounds(sheet)
    rg = { r1: 1, r2: Math.max(1, bd.r2), c1: Math.min(colNum(a), colNum(b)), c2: Math.max(colNum(a), colNum(b)) }
  } else if (/^\d+:\d+$/.test(body)) {
    const [a, b] = body.split(':').map(Number)
    const bd = env.bounds(sheet)
    rg = { r1: Math.min(a, b), r2: Math.max(a, b), c1: 1, c2: Math.max(1, bd.c2) }
  } else rg = parseRange(body)
  if (!rg) return ERR('#REF!')
  if (!env.has(sheet)) return ERR('#REF!')
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !body.includes(':')) return env.cell(sheet, rg.r1, rg.c1)
  if (env.range) return { range: true, values: env.range(sheet, rg) }
  const values = []
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = []
    for (let c = rg.c1; c <= rg.c2; c++) row.push(env.cell(sheet, r, c))
    values.push(row)
  }
  return { range: true, values }
}

const scalarOf = (v) => (v === null || v === undefined ? null : v)

function ev (n, env) {
  if ('lit' in n) return n.lit
  if (n.ref) return resolveRef(n.ref, env)
  if (n.name) {
    const own = env.lets?.get(n.name.toLowerCase())
    if (own !== undefined) return own
    // Defined names hold a formula of their own, usually a reference.
    const src = env.name?.(n.name)
    if (src === undefined) throw new Unsupported(`defined name ${n.name}`)
    if ((env.depth ?? 0) > 16) return ERR('#REF!')
    return ev(parseFormula(src), { ...env, depth: (env.depth ?? 0) + 1 })
  }
  if (n.fn) {
    const f = FUNCS[n.fn]
    if (!f) throw new Unsupported(`function ${n.fn}`)
    if (f.lazy) return f(n.args, env)
    return f(...n.args.map((a) => ev(a, env)))
  }
  const [a, b] = n.args
  const op = OPS[n.op]
  return n.op[0] === 'u' || n.op === '%' ? lift(ev(a, env), null, op) : lift(ev(a, env), ev(b, env), op)
}

const cmp = (k) => (x, y) => {
  if (isErr(x)) return x
  if (isErr(y)) return y
  const c = compare(x, y)
  return k === '=' ? c === 0 : k === '<>' ? c !== 0 : k === '<' ? c < 0 : k === '>' ? c > 0 : k === '<=' ? c <= 0 : c >= 0
}
const OPS = {
  'u-': (x) => arith(x, 0, (v) => -v),
  'u+': (x) => arith(x, 0, (v) => v),
  '%': (x) => arith(x, 0, (v) => v / 100),
  '+': (x, y) => arith2(x, y, (p, q) => p + q),
  '-': (x, y) => arith2(x, y, (p, q) => p - q),
  '*': (x, y) => arith2(x, y, (p, q) => p * q),
  '/': (x, y) => arith2(x, y, (p, q) => (q === 0 ? ERR('#DIV/0!') : p / q)),
  '^': (x, y) => arith2(x, y, Math.pow),
  '&': (x, y) => (isErr(x) ? x : isErr(y) ? y : str(x) + str(y)),
  '=': cmp('='), '<>': cmp('<>'), '<': cmp('<'), '>': cmp('>'), '<=': cmp('<='), '>=': cmp('>=')
}

// lift applies an operator cell by cell when a side is a range, as
// Excel does inside array-taking functions: (A1:A9="x")*B1:B9. A single
// row or column stretches across the other side.
function lift (x, y, f) {
  const gx = x && x.range ? x.values : null
  const gy = y && y.range ? y.values : null
  if (!gx && !gy) return f(x, y)
  const rows = Math.max(gx ? gx.length : 1, gy ? gy.length : 1)
  const cs = Math.max(gx ? gx[0]?.length ?? 0 : 1, gy ? gy[0]?.length ?? 0 : 1)
  const at = (g, v, r, c) => {
    if (!g) return v
    const row = g.length === 1 ? g[0] : g[r]
    if (!row) return ERR('#N/A')
    const cell = row.length === 1 ? row[0] : row[c]
    return cell === undefined ? ERR('#N/A') : cell
  }
  const values = []
  for (let r = 0; r < rows; r++) {
    const row = new Array(cs)
    for (let c = 0; c < cs; c++) row[c] = f(at(gx, x, r, c), at(gy, y, r, c))
    values.push(row)
  }
  return { range: true, values }
}

function single (v) {
  if (v && v.range) return v.values.length === 1 && v.values[0].length === 1 ? scalarOf(v.values[0][0]) : ERR('#VALUE!')
  return v
}

function num (v) {
  v = single(v)
  if (isErr(v)) return v
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  const n = Number(String(v).trim())
  return Number.isFinite(n) && String(v).trim() !== '' ? n : ERR('#VALUE!')
}

function str (v) {
  if (v === null || v === undefined) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number') return fmtGeneral(v)
  return String(v)
}

export function fmtGeneral (n) {
  if (Number.isInteger(n)) return String(n)
  const s = String(+n.toPrecision(15))
  return s
}

function bool (v) {
  v = single(v)
  if (isErr(v)) return v
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v === null || v === undefined || v === '') return false
  const u = String(v).toUpperCase()
  if (u === 'TRUE') return true
  if (u === 'FALSE') return false
  return ERR('#VALUE!')
}

function arith (x, _, f) {
  const a = num(x)
  return isErr(a) ? a : f(a)
}

function arith2 (x, y, f) {
  const a = num(x); if (isErr(a)) return a
  const b = num(y); if (isErr(b)) return b
  const r = f(a, b)
  return typeof r === 'number' && !Number.isFinite(r) ? ERR('#NUM!') : r
}

function compare (x, y) {
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 0)
  if (x === null || x === undefined) x = typeof y === 'string' ? '' : typeof y === 'boolean' ? false : 0
  if (y === null || y === undefined) y = typeof x === 'string' ? '' : typeof x === 'boolean' ? false : 0
  if (rank(x) !== rank(y)) return rank(x) - rank(y)
  if (typeof x === 'string') { const a = x.toLowerCase(); const b = y.toLowerCase(); return a < b ? -1 : a > b ? 1 : 0 }
  // Excel compares numbers to 15 significant digits: 0.1+0.2=0.3.
  if (typeof x === 'number' && x !== y && Math.abs(x - y) <= Math.max(Math.abs(x), Math.abs(y)) * 5e-15) return 0
  return x < y ? -1 : x > y ? 1 : 0
}

// values flattens arguments: ranges give their cells, scalars themselves.
function* values (args) {
  for (const a of args) {
    if (a && a.range) { for (const row of a.values) for (const v of row) yield { v, fromRange: true } } else yield { v: a, fromRange: false }
  }
}

// numbers takes what SUM-like functions count: numbers in ranges, and
// anything convertible given directly.
function numbers (args) {
  const out = []
  for (const { v, fromRange } of values(args)) {
    if (isErr(v)) throw new ErrorValue(v)
    if (fromRange) { if (typeof v === 'number') out.push(v) } else if (v !== null) {
      const n = num(v)
      if (isErr(n)) throw new ErrorValue(n)
      out.push(n)
    }
  }
  return out
}

class ErrorValue extends Error { constructor (v) { super(v.error); this.value = v } }

const guard = (f) => (...a) => { try { return f(...a) } catch (e) { if (e instanceof ErrorValue) return e.value; throw e } }

function criterion (c) {
  c = single(c)
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => compare(v, c) === 0 && typeof v === typeof c
  const m = /^(<>|<=|>=|=|<|>)?(.*)$/s.exec(String(c ?? ''))
  const op = m[1] ?? '='
  const raw = m[2]
  const n = raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : undefined
  let test
  if (n !== undefined) test = (v) => typeof v === 'number' && cmpOp(op, compare(v, n))
  else if (op === '=' || op === '<>') {
    const re = new RegExp('^' + raw.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'is')
    test = (v) => {
      const s = v === null || v === undefined ? '' : str(v)
      const hit = raw === '' ? s === '' : re.test(s)
      return op === '=' ? hit : !hit
    }
  } else test = (v) => typeof v === 'string' && cmpOp(op, compare(v, raw))
  return test
}

const cmpOp = (op, c) => ({ '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op])

// Ranges read during one recalculation are shared, so their flat form
// is kept with them; callers never change either.
const flats = new WeakMap()
const flat = (v) => {
  if (!(v && v.range)) return [v]
  let f = flats.get(v.values)
  if (!f) { f = v.values.length === 1 ? v.values[0] : v.values.flat(); flats.set(v.values, f) }
  return f
}

const round = (x, d, mode) => {
  const k = Math.pow(10, d)
  const y = Math.abs(x) * k
  const r = mode === 'up' ? Math.ceil(y - 1e-9) : mode === 'down' ? Math.floor(y + 1e-9) : Math.round(y + 1e-9)
  return Math.sign(x) * r / k
}

const FUNCS = {
  SUM: guard((...a) => numbers(a).reduce((s, x) => s + x, 0)),
  PRODUCT: guard((...a) => numbers(a).reduce((s, x) => s * x, 1)),
  AVERAGE: guard((...a) => { const n = numbers(a); return n.length ? n.reduce((s, x) => s + x, 0) / n.length : ERR('#DIV/0!') }),
  MIN: guard((...a) => { const n = numbers(a); return n.length ? Math.min(...n) : 0 }),
  MAX: guard((...a) => { const n = numbers(a); return n.length ? Math.max(...n) : 0 }),
  MEDIAN: guard((...a) => {
    const n = numbers(a).sort((x, y) => x - y)
    if (!n.length) return ERR('#NUM!')
    const h = n.length >> 1
    return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2
  }),
  COUNT: (...a) => [...values(a)].filter(({ v, fromRange }) => typeof v === 'number' || (!fromRange && !isErr(num(v)) && v !== null)).length,
  COUNTA: (...a) => [...values(a)].filter(({ v }) => v !== null && v !== undefined && v !== '').length,
  COUNTBLANK: (...a) => [...values(a)].filter(({ v }) => v === null || v === undefined || v === '').length,
  COUNTIF: (range, c) => { const t = criterion(c); return flat(range).filter((v) => t(v)).length },
  SUMIF: (range, c, sumRange) => {
    const t = criterion(c)
    const keys = flat(range)
    const vals = sumRange ? flat(sumRange) : keys
    let s = 0
    keys.forEach((k, i) => { if (t(k) && typeof vals[i] === 'number') s += vals[i] })
    return s
  },
  AVERAGEIF: (range, c, avgRange) => {
    const t = criterion(c)
    const keys = flat(range)
    const vals = avgRange ? flat(avgRange) : keys
    const hit = keys.map((k, i) => (t(k) && typeof vals[i] === 'number' ? vals[i] : undefined)).filter((x) => x !== undefined)
    return hit.length ? hit.reduce((s, x) => s + x, 0) / hit.length : ERR('#DIV/0!')
  },
  ROUND: (x, d) => arith2(x, d ?? 0, (a, b) => round(a, b)),
  ROUNDUP: (x, d) => arith2(x, d ?? 0, (a, b) => round(a, b, 'up')),
  ROUNDDOWN: (x, d) => arith2(x, d ?? 0, (a, b) => round(a, b, 'down')),
  INT: (x) => arith(x, 0, Math.floor),
  ABS: (x) => arith(x, 0, Math.abs),
  SQRT: (x) => arith(x, 0, (a) => (a < 0 ? ERR('#NUM!') : Math.sqrt(a))),
  POWER: (x, y) => arith2(x, y, Math.pow),
  MOD: (x, y) => arith2(x, y, (a, b) => (b === 0 ? ERR('#DIV/0!') : a - b * Math.floor(a / b))),
  IF: Object.assign((args, env) => {
    const c = bool(ev(args[0], env))
    if (isErr(c)) return c
    if (c) return args[1] ? scalar(ev(args[1], env)) : true
    return args[2] ? scalar(ev(args[2], env)) : false
  }, { lazy: true }),
  IFERROR: Object.assign((args, env) => {
    const v = scalar(ev(args[0], env))
    return isErr(v) ? scalar(ev(args[1], env)) : v
  }, { lazy: true }),
  AND: (...a) => { for (const { v } of values(a)) { if (v === null) continue; const b = bool(v); if (isErr(b)) return b; if (!b) return false } return true },
  OR: (...a) => { let any = false; for (const { v } of values(a)) { if (v === null) continue; const b = bool(v); if (isErr(b)) return b; if (b) any = true } return any },
  NOT: (x) => { const b = bool(x); return isErr(b) ? b : !b },
  CONCAT: (...a) => [...values(a)].map(({ v }) => str(v)).join(''),
  CONCATENATE: (...a) => a.map((v) => str(single(v))).join(''),
  LEN: (x) => str(single(x)).length,
  LEFT: (x, n) => str(single(x)).slice(0, n === undefined ? 1 : num(n)),
  RIGHT: (x, n) => { const s = str(single(x)); const k = n === undefined ? 1 : num(n); return k ? s.slice(-k) : '' },
  MID: (x, s, n) => str(single(x)).substr(num(s) - 1, num(n)),
  UPPER: (x) => str(single(x)).toUpperCase(),
  LOWER: (x) => str(single(x)).toLowerCase(),
  TRIM: (x) => str(single(x)).trim().replace(/ +/g, ' ')
}

function scalar (v) { return v && v.range ? single(v) : v }

// ------------------------------------------------------------ more functions

const grid = (v) => (v && v.range ? v.values : [[v]])
const cols = (v) => grid(v)[0]?.length ?? 0
const DAY = 86400000
// Excel's serial day numbers, with its 1900 leap-year bug.
const serialOf = (y, m, d) => { let s = Date.UTC(y, m, d) / DAY + 25569; if (s < 61) s -= 1; return s }
const dateOf = (n) => { let s = Math.floor(n); if (s < 61) s += 1; return new Date((s - 25569) * DAY) }
const dnum = (v) => { const n = num(v); if (isErr(n)) return n; return n < 0 ? ERR('#NUM!') : n }
const ifs = (pairs) => {
  const tests = []
  for (let i = 0; i + 1 < pairs.length; i += 2) tests.push([flat(pairs[i]), criterion(pairs[i + 1])])
  const n = tests[0]?.[0].length ?? 0
  if (tests.some(([r]) => r.length !== n)) return null
  const hit = []
  for (let k = 0; k < n; k++) if (tests.every(([r, t]) => t(r[k]))) hit.push(k)
  return hit
}
const pick = (range, hit) => { const v = flat(range); return hit.map((k) => v[k]).filter((x) => typeof x === 'number') }
const lookupIn = (key, list, mode) => {
  key = single(key)
  if (isErr(key)) return -1
  if (mode === 0) {
    const t = typeof key === 'string' ? criterion(key) : (v) => compare(v, key) === 0 && typeof v === typeof key
    return list.findIndex((v) => t(v))
  }
  // Sorted lookup: the last value not above (mode 1) or not below (-1) the key.
  let best = -1
  for (let i = 0; i < list.length; i++) {
    const v = list[i]
    if (v === null || v === undefined || typeof v !== typeof key) continue
    const c = compare(v, key)
    if (c === 0) return i
    if (mode === 1 && c < 0) best = i
    if (mode === -1 && c > 0) best = i
  }
  return best
}
const refPos = (arg, env) => {
  if (!arg) return env.at
  if (!arg.ref) return undefined
  const body = arg.ref.slice(arg.ref.lastIndexOf('!') + 1).replace(/\$/g, '')
  const rg = parseRange(body)
  return rg && { r: rg.r1, c: rg.c1, rows: rg.r2 - rg.r1 + 1, cols: rg.c2 - rg.c1 + 1 }
}
const stats = (a, sample) => {
  const n = numbers(a)
  if (n.length < (sample ? 2 : 1)) return ERR('#DIV/0!')
  const m = n.reduce((s, x) => s + x, 0) / n.length
  return n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - (sample ? 1 : 0))
}
const textFmt = (v, code) => {
  v = single(v)
  if (isErr(v)) return v
  if (typeof v !== 'number') return str(v)
  code = str(single(code))
  if (/[ymdhs]/i.test(code.replace(/"[^"]*"/g, ''))) {
    const d = dateOf(v)
    const t = (v % 1) * DAY
    const pad = (x, k = 2) => String(x).padStart(k, '0')
    const H = Math.floor(t / 3600000); const M = Math.floor(t / 60000) % 60; const S = Math.round(t / 1000) % 60
    return code.replace(/yyyy|yy|mm|m|dd|d|hh|h|ss|s/gi, (k) => {
      const u = k.toLowerCase()
      if (u === 'yyyy') return d.getUTCFullYear()
      if (u === 'yy') return pad(d.getUTCFullYear() % 100)
      if (u === 'mm') return pad(d.getUTCMonth() + 1)
      if (u === 'm') return d.getUTCMonth() + 1
      if (u === 'dd') return pad(d.getUTCDate())
      if (u === 'd') return d.getUTCDate()
      if (u === 'hh') return pad(H)
      if (u === 'h') return H
      if (u === 'ss') return pad(S)
      return S
    }).replace(/"/g, '')
  }
  if (/^general$/i.test(code.trim())) return fmtGeneral(v)
  const parts = code.split(';')
  const neg = v < 0 && parts.length > 1
  const sec = parts[neg ? 1 : v === 0 && parts.length > 2 ? 2 : 0]
  const lit = (t) => t.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1')
  const m = /[#0?][#0?,]*(\.[#0?]*)?/.exec(sec.replace(/"[^"]*"/g, (q) => ' '.repeat(q.length)))
  if (!m) return lit(sec)
  const x = Math.abs(v) * 100 ** (sec.replace(/"[^"]*"/g, '').split('%').length - 1)
  const intPat = m[0].split('.')[0]
  const decPat = m[1] ? m[1].slice(1) : ''
  const k = 10 ** decPat.length
  let [ip, dp = ''] = (Math.round(+(x * k).toPrecision(15)) / k).toFixed(decPat.length).split('.')
  const optional = /#*$/.exec(decPat)[0].length
  for (let i = 0; i < optional && dp.endsWith('0'); i++) dp = dp.slice(0, -1)
  const minInt = intPat.replace(/[^0]/g, '').length
  if (ip === '0' && !minInt) ip = ''
  ip = ip.padStart(minInt, '0')
  if (intPat.includes(',')) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  const body = ip + (dp || (decPat && !optional) ? '.' + dp : '')
  return (v < 0 && !neg ? '-' : '') + lit(sec.slice(0, m.index)) + body + lit(sec.slice(m.index + m[0].length))
}

Object.assign(FUNCS, {
  SUMIFS: (sum, ...p) => { const h = ifs(p); return h ? pick(sum, h).reduce((s, x) => s + x, 0) : ERR('#VALUE!') },
  COUNTIFS: (...p) => { const h = ifs(p); return h ? h.length : ERR('#VALUE!') },
  AVERAGEIFS: (avg, ...p) => { const h = ifs(p); if (!h) return ERR('#VALUE!'); const v = pick(avg, h); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : ERR('#DIV/0!') },
  MAXIFS: (rg, ...p) => { const h = ifs(p); if (!h) return ERR('#VALUE!'); const v = pick(rg, h); return v.length ? Math.max(...v) : 0 },
  MINIFS: (rg, ...p) => { const h = ifs(p); if (!h) return ERR('#VALUE!'); const v = pick(rg, h); return v.length ? Math.min(...v) : 0 },
  SUMPRODUCT: (...a) => {
    const lists = a.map(flat)
    if (lists.some((l) => l.length !== lists[0].length)) return ERR('#VALUE!')
    let s = 0
    for (let i = 0; i < lists[0].length; i++) s += lists.reduce((p, l) => p * (typeof l[i] === 'number' ? l[i] : 0), 1)
    return s
  },
  VLOOKUP: (key, table, col, approx) => {
    const g = grid(table); const k = num(col)
    if (isErr(k)) return k
    if (k < 1 || k > cols(table)) return ERR('#REF!')
    const i = lookupIn(key, g.map((r) => r[0]), approx === undefined || approx === null || bool(approx) === true ? 1 : 0)
    return i < 0 ? ERR('#N/A') : scalarOf(g[i][k - 1])
  },
  HLOOKUP: (key, table, row, approx) => {
    const g = grid(table); const k = num(row)
    if (isErr(k)) return k
    if (k < 1 || k > g.length) return ERR('#REF!')
    const i = lookupIn(key, g[0], approx === undefined || approx === null || bool(approx) === true ? 1 : 0)
    return i < 0 ? ERR('#N/A') : scalarOf(g[k - 1][i])
  },
  MATCH: (key, list, mode) => {
    const i = lookupIn(key, flat(list), mode === undefined || mode === null ? 1 : num(mode))
    return i < 0 ? ERR('#N/A') : i + 1
  },
  INDEX: (table, row, col) => {
    const g = grid(table)
    let r = num(row ?? 0); let c = num(col ?? 0)
    if (isErr(r)) return r
    if (isErr(c)) return c
    if (g.length === 1 && (col === undefined || col === null)) { c = r; r = 1 }
    // Row or column 0 picks the whole column or row.
    if (!r && g.length > 1 && c) return { range: true, values: g.map((x) => [x[c - 1]]) }
    if (!c && cols(table) > 1 && r) return g[r - 1] ? { range: true, values: [g[r - 1]] } : ERR('#REF!')
    if (!r) r = 1
    if (!c) c = 1
    return g[r - 1]?.[c - 1] === undefined ? ERR('#REF!') : scalarOf(g[r - 1][c - 1])
  },
  XLOOKUP: (key, list, ret, missing, mode) => {
    const m = mode === undefined || mode === null ? 0 : num(mode)
    const l = flat(list)
    const i = m === 0 ? lookupIn(key, l, 0) : lookupIn(key, l, m === -1 ? 1 : -1)
    if (i < 0) return missing === undefined || missing === null ? ERR('#N/A') : single(missing)
    const g = grid(ret)
    return scalarOf(g.length === 1 ? g[0][i] : g[i]?.[0])
  },
  CHOOSE: (k, ...a) => { const n = num(k); if (isErr(n)) return n; return n >= 1 && n <= a.length ? single(a[Math.floor(n) - 1]) : ERR('#VALUE!') },
  ROWS: (v) => grid(v).length,
  COLUMNS: (v) => cols(v),
  ROW: Object.assign((args, env) => { const p = refPos(args[0], env); if (!p) throw new Unsupported('ROW without a cell'); return p.r }, { lazy: true }),
  COLUMN: Object.assign((args, env) => { const p = refPos(args[0], env); if (!p) throw new Unsupported('COLUMN without a cell'); return p.c }, { lazy: true }),
  TODAY: () => { const d = new Date(); return serialOf(d.getFullYear(), d.getMonth(), d.getDate()) },
  NOW: () => { const d = new Date(); return serialOf(d.getFullYear(), d.getMonth(), d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400 },
  DATE: (y, m, d) => { const a = num(y); const b = num(m); const c = num(d); for (const x of [a, b, c]) if (isErr(x)) return x; return serialOf(a < 1900 ? a + 1900 : a, b - 1, c) },
  YEAR: (v) => { const n = dnum(v); return isErr(n) ? n : dateOf(n).getUTCFullYear() },
  MONTH: (v) => { const n = dnum(v); return isErr(n) ? n : dateOf(n).getUTCMonth() + 1 },
  DAY: (v) => { const n = dnum(v); return isErr(n) ? n : dateOf(n).getUTCDate() },
  WEEKDAY: (v, t) => {
    const n = dnum(v); if (isErr(n)) return n
    const w = dateOf(n).getUTCDay(); const k = t === undefined || t === null ? 1 : num(t)
    return k === 2 ? (w + 6) % 7 + 1 : k === 3 ? (w + 6) % 7 : w + 1
  },
  EDATE: (v, m) => { const n = dnum(v); const k = num(m); if (isErr(n)) return n; if (isErr(k)) return k; const d = dateOf(n); const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + k + 1, 0)).getUTCDate(); return serialOf(d.getUTCFullYear(), d.getUTCMonth() + k, Math.min(d.getUTCDate(), last)) },
  EOMONTH: (v, m) => { const n = dnum(v); const k = num(m); if (isErr(n)) return n; if (isErr(k)) return k; const d = dateOf(n); return serialOf(d.getUTCFullYear(), d.getUTCMonth() + k + 1, 0) },
  DATEDIF: (a, b, unit) => {
    const x = dnum(a); const y = dnum(b); if (isErr(x)) return x; if (isErr(y)) return y
    if (y < x) return ERR('#NUM!')
    const p = dateOf(x); const q = dateOf(y); const u = str(single(unit)).toUpperCase()
    let months = (q.getUTCFullYear() - p.getUTCFullYear()) * 12 + q.getUTCMonth() - p.getUTCMonth()
    if (q.getUTCDate() < p.getUTCDate()) months--
    if (u === 'D') return Math.floor(y) - Math.floor(x)
    if (u === 'M') return months
    if (u === 'Y') return Math.floor(months / 12)
    if (u === 'YM') return months % 12
    if (u === 'MD') return q.getUTCDate() >= p.getUTCDate() ? q.getUTCDate() - p.getUTCDate() : Math.round((y - serialOf(q.getUTCFullYear(), q.getUTCMonth() - 1, p.getUTCDate())))
    return ERR('#NUM!')
  },
  TEXT: (v, code) => textFmt(v, code),
  OFFSET: Object.assign((args, env) => {
    const a = args[0]
    if (!a?.ref) throw new Unsupported('OFFSET of something other than a reference')
    const p = refPos(a, env)
    if (!p) throw new Unsupported('OFFSET of a whole row or column')
    const at = (i, d) => { if (!args[i] || 'lit' in args[i] && args[i].lit === null) return d; const v = num(ev(args[i], env)); return isErr(v) ? v : Math.trunc(v) }
    const r = at(1, 0); const c = at(2, 0); const h = at(3, p.rows); const w = at(4, p.cols)
    for (const x of [r, c, h, w]) if (isErr(x)) return x
    const r1 = p.r + r; const c1 = p.c + c
    if (r1 < 1 || c1 < 1 || h < 1 || w < 1) return ERR('#REF!')
    const bang = a.ref.lastIndexOf('!')
    return resolveRef((bang >= 0 ? a.ref.slice(0, bang + 1) : '') + cellName(r1, c1) + (h > 1 || w > 1 ? ':' + cellName(r1 + h - 1, c1 + w - 1) : ''), env)
  }, { lazy: true }),
  INDIRECT: Object.assign((args, env) => {
    const t = str(single(ev(args[0], env))).trim()
    if (args[1] && bool(ev(args[1], env)) === false) throw new Unsupported('R1C1 references')
    TOKEN.lastIndex = 0
    const m = TOKEN.exec(t)
    if (m?.groups.ref === t) return resolveRef(t, env)
    const src = env.name?.(t)
    return src === undefined ? ERR('#REF!') : ev(parseFormula(src), env)
  }, { lazy: true }),
  ADDRESS: (row, col, abs, a1, sheet) => {
    const r = num(row); const c = num(col); if (isErr(r)) return r; if (isErr(c)) return c
    if (r < 1 || c < 1) return ERR('#VALUE!')
    if (a1 !== undefined && a1 !== null && bool(a1) === false) throw new Unsupported('R1C1 references')
    const k = abs === undefined || abs === null ? 1 : num(abs)
    const ref = (k === 1 || k === 3 ? '$' : '') + colName(Math.trunc(c)) + (k === 1 || k === 2 ? '$' : '') + Math.trunc(r)
    if (sheet === undefined || sheet === null) return ref
    const n = str(single(sheet))
    return (/^[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`) + '!' + ref
  },
  LET: Object.assign((args, env) => {
    if (args.length < 3 || args.length % 2 === 0) return ERR('#VALUE!')
    const lets = new Map(env.lets)
    const inner = { ...env, lets }
    for (let i = 0; i + 1 < args.length; i += 2) {
      if (!args[i].name) return ERR('#NAME?')
      lets.set(args[i].name.toLowerCase(), ev(args[i + 1], inner))
    }
    return ev(args[args.length - 1], inner)
  }, { lazy: true }),
  VALUE: (v) => { v = single(v); if (typeof v === 'number') return v; const n = Number(str(v).replace(/,/g, '').trim()); return str(v).trim() !== '' && Number.isFinite(n) ? n : ERR('#VALUE!') },
  SUBSTITUTE: (t, a, b, k) => {
    const s = str(single(t)); const from = str(single(a)); const to = str(single(b))
    if (!from) return s
    if (k === undefined || k === null) return s.split(from).join(to)
    const n = num(k); let at = -1
    for (let i = 0; i < n; i++) { at = s.indexOf(from, at + 1); if (at < 0) return s }
    return s.slice(0, at) + to + s.slice(at + from.length)
  },
  REPLACE: (t, start, n, by) => { const s = str(single(t)); const a = num(start) - 1; return s.slice(0, a) + str(single(by)) + s.slice(a + num(n)) },
  FIND: (a, t, start) => { const i = str(single(t)).indexOf(str(single(a)), (start === undefined ? 1 : num(start)) - 1); return i < 0 ? ERR('#VALUE!') : i + 1 },
  SEARCH: (a, t, start) => {
    const re = new RegExp(str(single(a)).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.'), 'i')
    const s = str(single(t)); const from = (start === undefined ? 1 : num(start)) - 1
    const m = re.exec(s.slice(from))
    return m ? m.index + from + 1 : ERR('#VALUE!')
  },
  REPT: (t, n) => { const k = num(n); return isErr(k) ? k : str(single(t)).repeat(Math.max(0, Math.floor(k))) },
  EXACT: (a, b) => str(single(a)) === str(single(b)),
  PROPER: (t) => str(single(t)).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p, c) => p + c.toUpperCase()),
  TEXTJOIN: (sep, skip, ...a) => { const e = bool(skip) === true; return [...values(a)].map(({ v }) => str(v)).filter((s) => !e || s !== '').join(str(single(sep))) },
  IFS: Object.assign((args, env) => {
    for (let i = 0; i + 1 < args.length; i += 2) { const c = bool(ev(args[i], env)); if (isErr(c)) return c; if (c) return scalar(ev(args[i + 1], env)) }
    return ERR('#N/A')
  }, { lazy: true }),
  SWITCH: Object.assign((args, env) => {
    const v = single(ev(args[0], env))
    let i = 1
    for (; i + 1 < args.length; i += 2) if (compare(v, single(ev(args[i], env))) === 0) return scalar(ev(args[i + 1], env))
    return i < args.length ? scalar(ev(args[i], env)) : ERR('#N/A')
  }, { lazy: true }),
  IFNA: Object.assign((args, env) => { const v = scalar(ev(args[0], env)); return isErr(v) && v.error === '#N/A' ? scalar(ev(args[1], env)) : v }, { lazy: true }),
  XOR: (...a) => { let n = 0; for (const { v } of values(a)) { if (v === null) continue; const b = bool(v); if (isErr(b)) return b; if (b) n++ } return n % 2 === 1 },
  ISBLANK: (v) => { v = single(v); return v === null || v === undefined },
  ISNUMBER: (v) => typeof single(v) === 'number',
  ISTEXT: (v) => typeof single(v) === 'string',
  ISERROR: (v) => !!isErr(single(v)),
  ISNA: (v) => { v = single(v); return !!isErr(v) && v.error === '#N/A' },
  ISLOGICAL: (v) => typeof single(v) === 'boolean',
  ISEVEN: (v) => { const n = num(v); return isErr(n) ? n : Math.floor(Math.abs(n)) % 2 === 0 },
  ISODD: (v) => { const n = num(v); return isErr(n) ? n : Math.floor(Math.abs(n)) % 2 === 1 },
  NA: () => ERR('#N/A'),
  RANK: (v, rg, order) => {
    const x = num(v); if (isErr(x)) return x
    const list = flat(rg).filter((n) => typeof n === 'number')
    if (!list.includes(x)) return ERR('#N/A')
    const asc = order !== undefined && order !== null && num(order) !== 0
    return list.filter((n) => (asc ? n < x : n > x)).length + 1
  },
  LARGE: guard((rg, k) => { const n = numbers([rg]).sort((a, b) => b - a); const i = num(k); return i >= 1 && i <= n.length ? n[i - 1] : ERR('#NUM!') }),
  SMALL: guard((rg, k) => { const n = numbers([rg]).sort((a, b) => a - b); const i = num(k); return i >= 1 && i <= n.length ? n[i - 1] : ERR('#NUM!') }),
  STDEV: guard((...a) => { const v = stats(a, true); return isErr(v) ? v : Math.sqrt(v) }),
  STDEVP: guard((...a) => { const v = stats(a, false); return isErr(v) ? v : Math.sqrt(v) }),
  VAR: guard((...a) => stats(a, true)),
  VARP: guard((...a) => stats(a, false)),
  CEILING: (x, s) => arith2(x, s ?? 1, (a, b) => (b === 0 ? 0 : Math.ceil(a / b - 1e-12) * b)),
  FLOOR: (x, s) => arith2(x, s ?? 1, (a, b) => (b === 0 ? ERR('#DIV/0!') : Math.floor(a / b + 1e-12) * b)),
  PI: () => Math.PI,
  EXP: (x) => arith(x, 0, Math.exp),
  LN: (x) => arith(x, 0, (a) => (a <= 0 ? ERR('#NUM!') : Math.log(a))),
  LOG: (x, b) => arith2(x, b ?? 10, (a, c) => (a <= 0 || c <= 0 || c === 1 ? ERR('#NUM!') : Math.log(a) / Math.log(c))),
  LOG10: (x) => arith(x, 0, (a) => (a <= 0 ? ERR('#NUM!') : Math.log10(a))),
  SIGN: (x) => arith(x, 0, Math.sign),
  TRUNC: (x, d) => arith2(x, d ?? 0, (a, b) => Math.trunc(a * 10 ** b) / 10 ** b),
  RAND: () => Math.random(),
  RANDBETWEEN: (a, b) => arith2(a, b, (x, y) => Math.floor(Math.random() * (Math.floor(y) - Math.ceil(x) + 1)) + Math.ceil(x))
})
FUNCS['STDEV.S'] = FUNCS.STDEV
FUNCS['STDEV.P'] = FUNCS.STDEVP
FUNCS['VAR.S'] = FUNCS.VAR
FUNCS['VAR.P'] = FUNCS.VARP
FUNCS['RANK.EQ'] = FUNCS.RANK

export const FUNCTION_NAMES = Object.keys(FUNCS)
