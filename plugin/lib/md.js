// The Markdown subset Agents write documents in (SPEC 3.2), parsed into
// blocks of inline runs that the Word writer turns into paragraphs.

// Block: { type: 'heading', level, runs } | { type: 'paragraph', runs }
//   | { type: 'item', ordered, level, list, runs } | { type: 'quote', runs }
//   | { type: 'code', text } | { type: 'table', rows: [[runs]], align }
//   | { type: 'image', src, alt } | { type: 'pagebreak' } | { type: 'rule' }
// Run: { text, b?, i?, s?, code?, link? } | { image, alt } | { br: true }
export function parseMarkdown (src) {
  const lines = String(src ?? '').replace(/\r\n?/g, '\n').split('\n')
  const blocks = []
  let list = 0 // id of the current list; a new list restarts numbering
  let inList = false
  let para = []
  const flush = () => {
    if (!para.length) return
    const textLines = para
    para = []
    const joined = textLines.join('\n')
    const img = /^!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)$/.exec(joined.trim())
    if (img) blocks.push({ type: 'image', alt: img[1], src: img[2] })
    else blocks.push({ type: 'paragraph', runs: inlineLines(textLines) })
  }
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    const trimmed = line.trim()
    if (!trimmed) { flush(); inList = false; continue }
    let m
    if (/^(```|~~~)/.test(trimmed)) {
      flush(); inList = false
      const fence = trimmed.slice(0, 3)
      const body = []
      i++
      while (i < lines.length && !lines[i].trim().startsWith(fence)) body.push(lines[i++])
      blocks.push({ type: 'code', text: body.join('\n') })
      continue
    }
    if (trimmed === '\\pagebreak' || trimmed === '\\newpage') { flush(); inList = false; blocks.push({ type: 'pagebreak' }); continue }
    if ((m = /^(#{1,6})\s+(.*?)\s*#*\s*$/.exec(trimmed))) {
      flush(); inList = false
      blocks.push({ type: 'heading', level: m[1].length, runs: inline(m[2]) })
      continue
    }
    if (/^([-*_])(\s*\1){2,}$/.test(trimmed)) { flush(); inList = false; blocks.push({ type: 'rule' }); continue }
    if (trimmed.startsWith('|') && i + 1 < lines.length && /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(lines[i + 1].trim())) {
      flush(); inList = false
      const cells = (l) => {
        let s = l.trim()
        if (s.startsWith('|')) s = s.slice(1)
        if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1)
        return s.split(/(?<!\\)\|/).map((c) => inline(c.trim().replace(/\\\|/g, '|')))
      }
      const align = lines[i + 1].trim().replace(/^\||\|$/g, '').split('|').map((c) => {
        c = c.trim()
        return c.startsWith(':') && c.endsWith(':') ? 'center' : c.endsWith(':') ? 'right' : 'left'
      })
      const rows = [cells(line)]
      i += 2
      while (i < lines.length && lines[i].trim().startsWith('|')) rows.push(cells(lines[i++]))
      i--
      blocks.push({ type: 'table', rows, align })
      continue
    }
    if ((m = /^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/.exec(line))) {
      flush()
      if (!inList) list++
      inList = true
      const level = Math.min(8, Math.floor(m[1].replace(/\t/g, '  ').length / 2))
      const ordered = /\d/.test(m[2])
      let body = m[3]
      // Task list markers read as text.
      body = body.replace(/^\[( |x|X)\]\s+/, (_, x) => (x === ' ' ? '☐ ' : '☑ '))
      blocks.push({ type: 'item', ordered, level, list, runs: inline(body) })
      continue
    }
    if (trimmed.startsWith('>')) {
      flush(); inList = false
      const body = [trimmed.replace(/^>\s?/, '')]
      while (i + 1 < lines.length && lines[i + 1].trim().startsWith('>')) body.push(lines[++i].trim().replace(/^>\s?/, ''))
      blocks.push({ type: 'quote', runs: inlineLines(body) })
      continue
    }
    if (inList && /^\s{2,}\S/.test(line) && blocks.length && blocks[blocks.length - 1].type === 'item') {
      // A continuation line of a list item.
      blocks[blocks.length - 1].runs.push({ text: ' ' }, ...inline(trimmed))
      continue
    }
    inList = false
    para.push(line.replace(/\s+$/, (s) => (s.length >= 2 ? '  ' : '')))
  }
  flush()
  return blocks
}

// inlineLines joins paragraph lines: a hard break where a line ends in two
// spaces or a backslash, a space otherwise.
function inlineLines (lines) {
  const runs = []
  lines.forEach((l, k) => {
    let hard = false
    if (/ {2}$/.test(l) || /\\$/.test(l)) { hard = true; l = l.replace(/( {2}|\\)$/, '') }
    runs.push(...inline(l.trim()))
    if (k < lines.length - 1) runs.push(hard ? { br: true } : { text: ' ' })
  })
  return merge(runs)
}

const PUNCT = /[\s!-/:-@[-`{-~　-〿＀-￯]/

// inline parses emphasis, code, links and images.
export function inline (s, style = {}) {
  const out = []
  let buf = ''
  const push = () => { if (buf) { out.push({ ...style, text: buf }); buf = '' } }
  let i = 0
  while (i < s.length) {
    const c = s[i]
    const rest = s.slice(i)
    let m
    if (c === '\\' && i + 1 < s.length && /[\\`*_{}[\]()#+\-.!|~<>]/.test(s[i + 1])) { buf += s[i + 1]; i += 2; continue }
    if (/^<br\s*\/?>/i.test(rest)) { push(); out.push({ br: true }); i += rest.match(/^<br\s*\/?>/i)[0].length; continue }
    if (c === '`') {
      const end = s.indexOf('`', i + 1)
      if (end > i) { push(); out.push({ ...style, text: s.slice(i + 1, end), code: true }); i = end + 1; continue }
    }
    if ((m = /^!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/.exec(rest))) { push(); out.push({ image: m[2], alt: m[1] }); i += m[0].length; continue }
    if ((m = /^\[((?:[^\]\\]|\\.)+)\]\(([^)\s]+)(?:\s+"[^"]*")?\)/.exec(rest))) { push(); out.push(...inline(m[1], { ...style, link: m[2] })); i += m[0].length; continue }
    if ((m = /^<(https?:\/\/[^>\s]+)>/.exec(rest))) { push(); out.push({ ...style, text: m[1], link: m[1] }); i += m[0].length; continue }
    const delim = rest.startsWith('**') ? '**' : rest.startsWith('__') ? '__' : rest.startsWith('~~') ? '~~' : (c === '*' || c === '_') ? c : null
    if (delim) {
      const before = i === 0 ? ' ' : s[i - 1]
      const after = s[i + delim.length]
      const opens = after && !/\s/.test(after) && (delim[0] !== '_' || PUNCT.test(before))
      if (opens) {
        let j = i + delim.length
        let end = -1
        while ((j = s.indexOf(delim, j)) >= 0) {
          const pre = s[j - 1]
          const post = s[j + delim.length]
          // A single * must not close on the start of **.
          if (delim.length === 1 && s[j + 1] === delim && s[j - 1] !== delim) { j += 2; continue }
          if (!/\s/.test(pre) && (delim[0] !== '_' || !post || PUNCT.test(post))) { end = j; break }
          j += delim.length
        }
        if (end > i + delim.length) {
          push()
          const key = delim === '~~' ? 's' : delim.length === 2 ? 'b' : 'i'
          out.push(...inline(s.slice(i + delim.length, end), { ...style, [key]: true }))
          i = end + delim.length
          continue
        }
      }
    }
    buf += c
    i++
  }
  push()
  return merge(out)
}

function merge (runs) {
  const out = []
  for (const r of runs) {
    const p = out[out.length - 1]
    if (p && r.text !== undefined && p.text !== undefined && !r.code && !p.code &&
        !!p.b === !!r.b && !!p.i === !!r.i && !!p.s === !!r.s && p.link === r.link) {
      p.text += r.text
    } else out.push({ ...r })
  }
  return out
}

// plain is the text of runs.
export const plain = (runs) => runs.map((r) => r.text ?? (r.br ? '\n' : r.alt ?? '')).join('')
