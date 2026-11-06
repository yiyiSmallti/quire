// A small XML tree that round-trips what it does not understand: element
// names keep their prefixes, attributes keep their order, comments and
// processing instructions stay where they were, and text keeps its
// whitespace. Office parts only use UTF-8, entities from the predefined
// set and numeric references, which is all this parser decodes.

export class XmlError extends Error {}

// Element: { name, attrs, kids }; text: { text }; anything else verbatim:
// { raw }.
export function el (name, attrs = {}, kids = []) {
  return { name, attrs, kids: kids.filter(Boolean) }
}

export const text = (v) => ({ text: String(v) })

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

function decode (s) {
  if (s.indexOf('&') < 0) return s
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const n = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
      return Number.isFinite(n) ? String.fromCodePoint(n) : m
    }
    return ENT[e] ?? m
  })
}

// parse returns { prolog, root } where prolog keeps the declaration and
// anything before the root element.
export function parse (src) {
  if (Buffer.isBuffer(src)) src = src.toString('utf8')
  if (src.charCodeAt(0) === 0xfeff) src = src.slice(1)
  let i = 0
  const n = src.length
  const prolog = []
  const stack = []
  let root = null
  const push = (node) => {
    if (stack.length) stack[stack.length - 1].kids.push(node)
    else if (node.name) {
      if (root) throw new XmlError('more than one root element')
      root = node
    } else if (!root) prolog.push(node)
  }
  while (i < n) {
    const lt = src.indexOf('<', i)
    if (lt < 0) {
      if (stack.length) push(text(decode(src.slice(i))))
      break
    }
    if (lt > i && stack.length) push(text(decode(src.slice(i, lt))))
    if (src.startsWith('<!--', lt)) {
      const end = src.indexOf('-->', lt + 4)
      if (end < 0) throw new XmlError('unterminated comment')
      push({ raw: src.slice(lt, end + 3) })
      i = end + 3
    } else if (src.startsWith('<![CDATA[', lt)) {
      const end = src.indexOf(']]>', lt + 9)
      if (end < 0) throw new XmlError('unterminated CDATA')
      push(text(src.slice(lt + 9, end)))
      i = end + 3
    } else if (src[lt + 1] === '?' || src[lt + 1] === '!') {
      const end = src.indexOf('>', lt)
      if (end < 0) throw new XmlError('unterminated declaration')
      push({ raw: src.slice(lt, end + 1) })
      i = end + 1
    } else if (src[lt + 1] === '/') {
      const end = src.indexOf('>', lt)
      const name = src.slice(lt + 2, end).trim()
      const top = stack.pop()
      if (!top || top.name !== name) throw new XmlError(`mismatched closing tag </${name}>`)
      if (!stack.length) {
        if (root !== top) throw new XmlError('stray element')
      }
      i = end + 1
    } else {
      // Start tag: name, then attributes until > or />.
      let j = lt + 1
      while (j < n && !/[\s/>]/.test(src[j])) j++
      const node = { name: src.slice(lt + 1, j), attrs: {}, kids: [] }
      for (;;) {
        while (j < n && /\s/.test(src[j])) j++
        if (src[j] === '>' || src[j] === '/') break
        const eq = src.indexOf('=', j)
        if (eq < 0) throw new XmlError(`malformed attribute in <${node.name}>`)
        const an = src.slice(j, eq).trim()
        let k = eq + 1
        while (/\s/.test(src[k])) k++
        const q = src[k]
        if (q !== '"' && q !== "'") throw new XmlError(`unquoted attribute ${an} in <${node.name}>`)
        const close = src.indexOf(q, k + 1)
        if (close < 0) throw new XmlError('unterminated attribute value')
        node.attrs[an] = decode(src.slice(k + 1, close))
        j = close + 1
      }
      const selfClose = src[j] === '/'
      const end = src.indexOf('>', j)
      push(node)
      if (!selfClose) stack.push(node)
      else if (!stack.length && root !== node) root = node
      i = end + 1
    }
  }
  if (stack.length) throw new XmlError(`unclosed <${stack[stack.length - 1].name}>`)
  if (!root) throw new XmlError('no root element')
  return { prolog, root }
}

const escText = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const escAttr = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;').replace(/\r/g, '&#13;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;')

export function serializeNode (node, out = []) {
  if (node.text !== undefined) { out.push(escText(node.text)); return out }
  if (node.raw !== undefined) { out.push(node.raw); return out }
  out.push('<', node.name)
  for (const k in node.attrs) {
    const v = node.attrs[k]
    if (v === undefined || v === null) continue
    out.push(' ', k, '="', escAttr(v), '"')
  }
  if (!node.kids.length) { out.push('/>'); return out }
  out.push('>')
  for (const k of node.kids) serializeNode(k, out)
  out.push('</', node.name, '>')
  return out
}

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

export function serialize (doc) {
  const out = []
  if (doc.prolog && doc.prolog.length) {
    for (const p of doc.prolog) if (p.raw !== undefined) out.push(p.raw)
  } else out.push(DECL)
  out.push('\r\n')
  serializeNode(doc.root, out)
  return out.join('')
}

// A new document around root.
export const doc = (root) => ({ prolog: [{ raw: DECL }], root })

// ------------------------------------------------------------ navigation

export const isEl = (n) => n && n.name !== undefined
export const child = (node, name) => node?.kids?.find((k) => k.name === name)
export const children = (node, name) => (node?.kids ?? []).filter((k) => name === undefined ? isEl(k) : k.name === name)

// Every descendant element named name (or matching a predicate), in
// document order.
export function descendants (node, name, out = []) {
  if (!node?.kids) return out
  const match = typeof name === 'function' ? name : (k) => k.name === name
  for (const k of node.kids) {
    if (!isEl(k)) continue
    if (match(k)) out.push(k)
    descendants(k, name, out)
  }
  return out
}

// path follows child names: path(root, 'p:cSld', 'p:spTree').
export function path (node, ...names) {
  let cur = node
  for (const nm of names) {
    cur = child(cur, nm)
    if (!cur) return undefined
  }
  return cur
}

// ensure returns the child named name, appending it (or inserting it in
// order) when missing.
export function ensure (node, name, order) {
  let c = child(node, name)
  if (!c) {
    c = el(name)
    insertOrdered(node, c, order)
  }
  return c
}

// insertOrdered puts kid among node's children before the first sibling
// whose name comes later in order (OOXML schemas fix element order).
export function insertOrdered (node, kid, order) {
  if (!order) { node.kids.push(kid); return kid }
  const rank = order.indexOf(kid.name)
  if (rank < 0) { node.kids.push(kid); return kid }
  const at = node.kids.findIndex((k) => isEl(k) && order.indexOf(k.name) > rank)
  if (at < 0) node.kids.push(kid)
  else node.kids.splice(at, 0, kid)
  return kid
}

export function remove (parent, kid) {
  const i = parent.kids.indexOf(kid)
  if (i >= 0) parent.kids.splice(i, 1)
  return i
}

// textContent concatenates every text node below node.
export function textContent (node) {
  if (!node) return ''
  if (node.text !== undefined) return node.text
  if (!node.kids) return ''
  let s = ''
  for (const k of node.kids) s += textContent(k)
  return s
}

// Deep copy.
export function clone (node) {
  if (node.text !== undefined) return { text: node.text }
  if (node.raw !== undefined) return { raw: node.raw }
  return { name: node.name, attrs: { ...node.attrs }, kids: node.kids.map(clone) }
}

export const esc = escText
