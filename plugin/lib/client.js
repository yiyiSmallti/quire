// Quire in dsh web: open .docx, .xlsx and .pptx from the right Sidebar's
// file preview and edit them in place. The Host half (ui-host.js) reads
// the file into an editing model and applies what the person changes,
// so untouched content stays as it was; saving is automatic, one batch
// at a time, and refused when the Agent changed the file meanwhile.
//
// Loaded by dsh's client module system (window.__ModuleLoader__); React
// comes from the platform module table, the editors themselves are plain
// DOM.
window.__ModuleLoader__.load({
  id: '@precinct/quire',
  factory: (require) => {
    const module = { exports: {} }
    const React = require('react')
    const { useEffect, useRef } = React
    const ID = '@precinct/quire/editor'

    // ------------------------------------------------------------ helpers

    function el (tag, attrs, ...kids) {
      const n = document.createElement(tag)
      for (const [k, v] of Object.entries(attrs ?? {})) {
        if (v === undefined || v === null || v === false) continue
        if (k === 'class') n.className = v
        else if (k === 'style' && typeof v === 'object') Object.assign(n.style, v)
        else if (k === 'dataset') Object.assign(n.dataset, v)
        else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v)
        else if (k === 'html') n.innerHTML = v
        else n.setAttribute(k, v === true ? '' : v)
      }
      for (const kid of kids.flat()) if (kid !== undefined && kid !== null && kid !== false) n.append(kid instanceof Node ? kid : String(kid))
      return n
    }
    const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    const mod = (e) => e.metaKey || e.ctrlKey
    const CSS_ID = (s) => String(s).replace(/["\\]/g, '\\$&')
    const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) } return s }
    const cellName = (r, c) => colName(c) + r

    // The file-address grammar of dsh's resource model.
    function parseFileAddress (address) {
      const PREFIX = 'dsh-resource://file/'
      if (typeof address !== 'string' || !address.startsWith(PREFIX)) return undefined
      const end = address.search(/[?#]/)
      const [scope, id, ...segments] = address.slice(PREFIX.length, end === -1 ? undefined : end).split('/')
      if (scope !== 'session' || !id || !segments.length) return undefined
      try { return { sessionId: decodeURIComponent(id), path: segments.map(decodeURIComponent).join('/') } } catch { return undefined }
    }

    class ApiError extends Error {
      constructor (status, body) { super(body?.error?.message ?? `HTTP ${status}`); this.status = status; this.code = body?.error?.code; this.body = body }
    }
    async function api (name, body) {
      let res
      try {
        res = await fetch(`api/quire/${name}`, { method: 'POST', headers: { 'content-type': 'application/json' }, credentials: 'same-origin', body: JSON.stringify(body) })
      } catch (err) { throw new ApiError(0, { error: { code: 'network', message: '连接不到 dsh：' + err.message } }) }
      const data = await res.json().catch(() => ({}))
      if (!res.ok) throw new ApiError(res.status, data)
      return data
    }

    const ICONS = {
      undo: '<path d="M4 7h7a4 4 0 0 1 0 8H7" /><path d="M7 4L4 7l3 3" />',
      redo: '<path d="M14 7H7a4 4 0 0 0 0 8h4" /><path d="M11 4l3 3-3 3" />',
      full: '<path d="M3 7V3h4M13 3h4v4M17 13v4h-4M7 17H3v-4" />',
      exit: '<path d="M7 3v4H3M13 3v4h4M13 17v-4h4M7 17v-4H3" />',
      left: '<path d="M3 5h14M3 9h9M3 13h14M3 17h9" />',
      center: '<path d="M3 5h14M5.5 9h9M3 13h14M5.5 17h9" />',
      right: '<path d="M3 5h14M8 9h9M3 13h14M8 17h9" />',
      justify: '<path d="M3 5h14M3 9h14M3 13h14M3 17h14" />',
      bullet: '<circle cx="4" cy="6" r="1.2" fill="currentColor"/><circle cx="4" cy="14" r="1.2" fill="currentColor"/><path d="M8 6h9M8 14h9" />',
      number: '<path d="M3 4h1.5v4M3 8h3M3 12.5c0-1 2.5-1 2.5 0S3 14.5 3 16h3" /><path d="M9 6h8M9 14h8" />',
      outdent: '<path d="M3 4h14M9 8h8M9 12h8M3 16h14M6 8l-3 2 3 2" />',
      indent: '<path d="M3 4h14M9 8h8M9 12h8M3 16h14M3 8l3 2-3 2" />',
      wrap: '<path d="M3 5h14M3 10h11a3 3 0 0 1 0 6h-3M3 15h5" /><path d="M12 14l-2 2 2 2" />',
      border: '<rect x="3" y="3" width="14" height="14" /><path d="M3 10h14M10 3v14" />',
      merge: '<rect x="3" y="5" width="14" height="10" /><path d="M7 10h6M11 8l2 2-2 2M9 8l-2 2 2 2" />',
      table: '<rect x="3" y="4" width="14" height="12" /><path d="M3 8h14M3 12h14M8 4v12M13 4v12" />',
      text: '<rect x="3" y="4" width="14" height="12" stroke-dasharray="2 2" /><path d="M7 8h6M10 8v5" />',
      trash: '<path d="M4 6h12M8 6V4h4v2M6 6l1 11h6l1-11" />',
      up: '<path d="M10 16V4M5 9l5-5 5 5" />',
      down: '<path d="M10 4v12M5 11l5 5 5-5" />',
      copy: '<rect x="7" y="7" width="10" height="10" /><path d="M13 7V3H3v10h4" />',
      plus: '<path d="M10 4v12M4 10h12" />',
      eye: '<path d="M2 10s3-5 8-5 8 5 8 5-3 5-8 5-8-5-8-5z" /><circle cx="10" cy="10" r="2" />',
      freeze: '<rect x="3" y="3" width="14" height="14" /><path d="M3 8h14M8 3v14" stroke-width="2.2" />',
      rowIns: '<rect x="3" y="3" width="14" height="5" /><rect x="3" y="12" width="14" height="5" /><path d="M10 8.5v3M8.5 10h3" />',
      rowDel: '<rect x="3" y="3" width="14" height="5" /><rect x="3" y="12" width="14" height="5" /><path d="M8.5 10h3" />',
      colIns: '<rect x="3" y="3" width="5" height="14" /><rect x="12" y="3" width="5" height="14" /><path d="M10 8.5v3M8.5 10h3" />',
      colDel: '<rect x="3" y="3" width="5" height="14" /><rect x="12" y="3" width="5" height="14" /><path d="M8.5 10h3" />',
      pagebreak: '<path d="M4 3v5h12V3M4 17v-5h12v5" /><path d="M2 10h2M6 10h2M10 10h2M14 10h2" />',
      print: '<path d="M6 8V3h8v5" /><rect x="3" y="8" width="14" height="6" rx="1" /><path d="M6 12h8v5H6z" />',
      search: '<circle cx="8.5" cy="8.5" r="5" /><path d="M12.5 12.5L17 17" />',
      image: '<rect x="3" y="4" width="14" height="12" rx="1" /><circle cx="7.5" cy="8" r="1.5" /><path d="M3 14l4-4 3 3 2-2 5 5" />',
      link: '<path d="M8.5 11.5a3 3 0 0 0 4.2 0l2.6-2.6a3 3 0 0 0-4.2-4.2L10 5.8" /><path d="M11.5 8.5a3 3 0 0 0-4.2 0l-2.6 2.6a3 3 0 0 0 4.2 4.2l1.1-1.1" />',
      sup: '<path d="M3 15l6-8M3 7l6 8" /><path d="M12 6.5c0-1.5 3-1.5 3 0 0 1.2-3 2.2-3 3.5h3" />',
      sub: '<path d="M3 13l6-8M3 5l6 8" /><path d="M12 14.5c0-1.5 3-1.5 3 0 0 1.2-3 2.2-3 3.5h3" />',
      sortUp: '<path d="M6 16V4M3 7l3-3 3 3" /><path d="M11 5h6M11 9h4.5M11 13h3M11 17h1.5" />',
      sortDown: '<path d="M6 4v12M3 13l3 3 3-3" /><path d="M11 5h1.5M11 9h3M11 13h4.5M11 17h6" />',
      filter: '<path d="M3 4h14l-5.5 6.5V16l-3 1.5v-7z" />',
      decInc: '<path d="M2 15h1" /><text x="5" y="15" font-size="8" fill="currentColor" stroke="none">.00</text><path d="M14 4l3 2-3 2" />',
      decDec: '<path d="M2 15h1" /><text x="5" y="15" font-size="8" fill="currentColor" stroke="none">.0</text><path d="M17 4l-3 2 3 2" />',
      colInsL: '<rect x="9" y="3" width="8" height="14" /><path d="M2 10h5M4.5 7.5v5" />',
      colInsR: '<rect x="3" y="3" width="8" height="14" /><path d="M13 10h5M15.5 7.5v5" />',
      firstLine: '<path d="M8 4h9M3 8h14M3 12h14M3 16h14" /><path d="M3 2.5v3l2.5-1.5z" fill="currentColor" />',
      play: '<path d="M6 4l10 6-10 6z" />',
      front: '<rect x="3" y="3" width="9" height="9" /><rect x="8" y="8" width="9" height="9" fill="currentColor" fill-opacity=".35" />',
      shape: '<rect x="2.5" y="9" width="7" height="7" /><circle cx="13.5" cy="6.5" r="4" />',
      noFill: '<rect x="3" y="3" width="14" height="14" /><path d="M3 17L17 3" />',
      comment: '<path d="M3 4h14v9.5H8.5L4.5 17v-3.5H3z" />',
      commentAdd: '<path d="M3 4h14v9.5H8.5L4.5 17v-3.5H3z" /><path d="M10 6.5v4.5M7.75 8.75h4.5" />',
      sum: '<path d="M15 4H5l6 6-6 6h10" />',
      fx: '<path d="M7 17c2 0 2.4-2.5 2.8-7S11 3 13.5 3" /><path d="M6 9h6" /><path d="M12 12l4.5 5M16.5 12L12 17" />',
      cf: '<rect x="3" y="3" width="14" height="14" rx="1" /><path d="M3 8h14M8 3v14" /><rect x="10" y="10" width="5" height="5" fill="currentColor" stroke="none" opacity=".5" />',
      dv: '<rect x="3" y="4" width="14" height="12" rx="1" /><path d="M6.5 10l2.5 2.5L14 7.5" />',
      names: '<path d="M3 6h14M3 10h14M3 14h8" /><circle cx="15" cy="14.5" r="2" />',
      brush: '<path d="M4 3h10v4H4z" /><path d="M14 5h2v4H9v3" /><rect x="7.5" y="12" width="3" height="5" rx="1" />',
      lock: '<rect x="4" y="9" width="12" height="8" rx="1.5" /><path d="M7 9V6.5a3 3 0 0 1 6 0V9" />',
      group: '<path d="M4 4v12M4 4h3M4 16h3" /><path d="M10 7h7M10 10h7M10 13h7" />',
      chart: '<path d="M3 17h14" /><rect x="4.5" y="10" width="2.5" height="5" /><rect x="8.75" y="5" width="2.5" height="10" /><rect x="13" y="8" width="2.5" height="7" />',
      pin: '<path d="M4 4h12v9h-6l-4 3.5V13H4z" /><path d="M10 6.5v4M8 8.5h4" />',
      nav: '<rect x="3" y="4" width="14" height="12" rx="1.5" /><path d="M7.5 4v12M10 7.5h5M10 10h4M10 12.5h5" />'
    }
    const icon = (name) => `<svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${ICONS[name]}</svg>`

    // numText writes a list counter in a Word number format.
    function numText (n, fmt) {
      const roman = (x) => { let out = ''; for (const [v, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) while (x >= v) { out += r; x -= v } return out }
      const letter = (x) => { let out = ''; while (x > 0) { x--; out = String.fromCharCode(97 + (x % 26)) + out; x = Math.floor(x / 26) } return out }
      const cn = (x) => {
        const d = '零一二三四五六七八九'
        if (x < 10) return d[x]
        if (x < 20) return '十' + (x % 10 ? d[x % 10] : '')
        if (x < 100) return d[Math.floor(x / 10)] + '十' + (x % 10 ? d[x % 10] : '')
        return String(x)
      }
      switch (fmt) {
        case 'lowerLetter': return letter(n)
        case 'upperLetter': return letter(n).toUpperCase()
        case 'lowerRoman': return roman(n)
        case 'upperRoman': return roman(n).toUpperCase()
        case 'chineseCounting': case 'chineseCountingThousand': case 'chineseLegalSimplified': case 'ideographDigital': return cn(n)
        case 'decimalEnclosedCircle': case 'decimalEnclosedCircleChinese': return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n)
        case 'ideographTraditional': return '甲乙丙丁戊己庚辛壬癸'[(n - 1) % 10]
        case 'decimalZero': return n < 10 ? '0' + n : String(n)
        default: return String(n)
      }
    }

    // htmlTable reads pasted HTML that is one table and nothing else (from
    // Word, Excel or a web page) as rows of cell text; spanned cells leave
    // blanks after them. The header is a first row of <th> or a <thead>.
    function htmlTable (html) {
      const docu = new DOMParser().parseFromString(html, 'text/html')
      const tables = docu.body.querySelectorAll('table')
      const t = tables[0]
      if (!t || [...tables].some((x) => x !== t && !t.contains(x))) return null
      const squash = (s) => s.replace(/\s+/g, '')
      if (squash(docu.body.textContent) !== squash(t.textContent)) return null
      const text = (n) => {
        let out = ''
        const walk = (x) => { for (const c of x.childNodes) { if (c.nodeType === 3) out += c.data.replace(/[\t\n\r ]+/g, ' '); else if (c.tagName === 'BR') out += '\n'; else if (c.nodeType === 1 && !/^(STYLE|SCRIPT)$/.test(c.tagName)) { if (/^(P|DIV|LI)$/.test(c.tagName) && out && !out.endsWith('\n')) out += '\n'; walk(c) } } }
        walk(n)
        return out.split('\n').map((l) => l.trim()).join('\n').replace(/^\n+|\n+$/g, '')
      }
      const rows = [...t.rows].map((tr) => [...tr.cells].flatMap((td) => [text(td), ...Array.from({ length: Math.min(20, +td.colSpan || 1) - 1 }, () => '')]))
      if (!rows.length || !rows.some((r) => r.length)) return null
      const header = !!t.tHead || (t.rows[0] && [...t.rows[0].cells].every((c) => c.tagName === 'TH'))
      return { rows, header }
    }

    // parseHtml turns pasted HTML into paragraphs of runs the editor knows:
    // bold, italic, underline, strike, colour, highlight, size, font,
    // super/subscript and links; headings, quotes, code and lists.
    function parseHtml (html) {
      const docu = new DOMParser().parseFromString(html, 'text/html')
      const out = []
      let cur = null
      const hex = (c) => {
        const m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/.exec(c ?? '')
        if (m) return m[4] !== undefined && +m[4] === 0 ? undefined : [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase()
        const h = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec((c ?? '').trim())
        if (!h) return undefined
        return (h[1].length === 3 ? [...h[1]].map((x) => x + x).join('') : h[1]).toUpperCase()
      }
      const inline = (n, props) => {
        const p = { ...props }
        const tag = n.tagName
        if (tag === 'B' || tag === 'STRONG') p.b = true
        if (tag === 'I' || tag === 'EM') p.i = true
        if (tag === 'U' || tag === 'INS') p.u = true
        if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') p.s = true
        if (tag === 'SUP') p.vert = 'sup'
        if (tag === 'SUB') p.vert = 'sub'
        if (tag === 'A' && /^(https?:|mailto:)/i.test(n.getAttribute('href') ?? '')) p.link = n.getAttribute('href')
        if (tag === 'FONT' && hex(n.getAttribute('color'))) p.color = hex(n.getAttribute('color'))
        const st = n.style
        if (st) {
          const fw = st.fontWeight
          if (fw) p.b = fw === 'bold' || fw === 'bolder' || +fw >= 600
          if (st.fontStyle) p.i = st.fontStyle === 'italic' || st.fontStyle === 'oblique'
          const td = st.textDecorationLine || st.textDecoration
          if (td) { if (/underline/.test(td)) p.u = true; if (/line-through/.test(td)) p.s = true; if (td === 'none') { delete p.u; delete p.s } }
          if (st.color && hex(st.color)) p.color = hex(st.color)
          const bg = hex(st.backgroundColor)
          if (bg && bg !== 'FFFFFF') p.highlight = bg
          const fs = /^([\d.]+)(pt|px)$/.exec(st.fontSize ?? '')
          if (fs) { const pt = fs[2] === 'pt' ? +fs[1] : +fs[1] * 0.75; if (pt >= 1 && pt <= 400) p.size = Math.round(pt * 2) }
          const ff = (st.fontFamily ?? '').split(',')[0].replace(/["']/g, '').trim()
          if (ff && !/^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial|-apple-system|BlinkMacSystemFont)$/i.test(ff)) p.font = ff
          if (st.verticalAlign === 'super') p.vert = 'sup'
          if (st.verticalAlign === 'sub') p.vert = 'sub'
        }
        return p
      }
      const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE', 'TR', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'UL', 'OL', 'TABLE', 'TBODY', 'THEAD', 'DT', 'DD', 'FIGURE', 'FIGCAPTION', 'MAIN', 'NAV', 'ASIDE', 'ADDRESS'])
      const add = (text, props, blk) => {
        if (!cur) { cur = { runs: [], ...blk }; out.push(cur) }
        const last = cur.runs[cur.runs.length - 1]
        const clean = Object.fromEntries(Object.entries(props).filter(([, v]) => v !== undefined))
        if (last && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...clean, text: '' })) last.text += text
        else cur.runs.push({ ...clean, text })
      }
      const walk = (node, props, blk, pre) => {
        for (const c of node.childNodes) {
          if (c.nodeType === 3) {
            let t = c.data
            if (!pre) t = t.replace(/[\t\n\r ]+/g, ' ')
            if (!t || (!cur && !t.trim())) continue
            if (!cur && !pre) t = t.replace(/^ +/, '')
            add(t, props, blk)
            continue
          }
          if (c.nodeType !== 1) continue
          const tag = c.tagName
          if (/^(SCRIPT|STYLE|META|LINK|TITLE|HEAD|TEMPLATE|NOSCRIPT|IMG|svg|SVG|BUTTON|INPUT|SELECT|TEXTAREA)$/.test(tag)) continue
          if (c.getAttribute?.('style')?.includes('mso-list:Ignore')) continue
          if (tag === 'BR') { add('\n', props, blk); continue }
          if (tag === 'TD' || tag === 'TH') { if (cur && cur.runs.length) add('\t', props, blk); walk(c, inline(c, props), blk, pre); continue }
          const p2 = inline(c, props)
          if (!BLOCK.has(tag)) { walk(c, p2, blk, pre); continue }
          const nb = { ...blk }
          const h = /^H([1-6])$/.exec(tag)
          if (h) nb.style = '@h' + Math.min(3, +h[1])
          if (tag === 'BLOCKQUOTE') nb.style = '@quote'
          if (tag === 'PRE') nb.style = '@code'
          if (tag === 'UL' || tag === 'OL') { nb.listKind = tag; nb.depth = (blk.depth ?? -1) + 1 }
          if (tag === 'LI') nb.list = { ordered: blk.listKind === 'OL', level: Math.min(8, Math.max(0, blk.depth ?? 0)) }
          const al = c.style?.textAlign || c.getAttribute('align')
          if (al && ['left', 'center', 'right', 'justify'].includes(al)) nb.align = al === 'justify' ? 'both' : al
          cur = null
          walk(c, p2, nb, pre || tag === 'PRE')
          cur = null
        }
      }
      walk(docu.body, {}, {}, false)
      for (const p of out) {
        delete p.listKind; delete p.depth
        while (p.runs.length && !p.runs[p.runs.length - 1].text.replace(/[ \n]+$/, '')) p.runs.pop()
        const last = p.runs[p.runs.length - 1]
        if (last) last.text = last.text.replace(/[ \n]+$/, '')
        for (const r of p.runs) r.text = r.text.replace(/ /g, ' ')
      }
      return out.filter((p) => p.runs.some((r) => r.text.trim()))
    }

    // A toolbar of small controls; each returns its element.
    // With a strip, tab(name) starts a ribbon tab: the controls after it
    // go on that tab's panel, and one panel shows at a time.
    class Toolbar {
      constructor (root, strip) { this.root = root; this.base = root; this.strip = strip; this.panels = [] }
      tab (name) {
        const panel = el('div', { class: 'q-panel', role: 'tabpanel', dataset: { tab: name }, hidden: true })
        const btn = el('button', { type: 'button', class: 'q-tab', role: 'tab', 'aria-selected': 'false', dataset: name === '文件' ? { file: '1' } : undefined }, name)
        btn.addEventListener('mousedown', (e) => e.preventDefault())
        btn.addEventListener('click', () => this.show(name))
        this.base.append(panel)
        this.strip.append(btn)
        this.panels.push({ name, panel, btn })
        this.root = panel
        return panel
      }
      show (name, remember = true) {
        for (const p of this.panels) {
          const on = p.name === name
          p.panel.hidden = !on
          p.btn.classList.toggle('on', on)
          p.btn.setAttribute('aria-selected', String(on))
        }
        // 文件 is a place to pass through, as in Office: never reopened on.
        if (remember && this.key && name !== '文件') try { localStorage.setItem('quire.tab.' + this.key, name) } catch {}
        this.onShow?.(name)
      }
      // restore opens the tab last used for this kind of file.
      restore (key) {
        this.key = key
        let saved
        try { saved = localStorage.getItem('quire.tab.' + key) } catch {}
        this.show(saved !== '文件' && this.panels.some((p) => p.name === saved) ? saved : (this.panels.find((p) => p.name !== '文件') ?? this.panels[0])?.name, false)
      }
      button (label, title, onClick, { ico, cls } = {}) {
        const b = el('button', { type: 'button', class: 'q-btn' + (cls ? ' ' + cls : ''), title, 'aria-label': title, html: ico ? icon(ico) + (label ? `<span class="q-lbl">${esc(label)}</span>` : '') : esc(label) })
        b.addEventListener('mousedown', (e) => e.preventDefault())
        b.addEventListener('click', (e) => { e.preventDefault(); if (!b.disabled) onClick(e) })
        this.root.append(b)
        return b
      }
      select (title, options, onChange, { width } = {}) {
        const s = el('select', { class: 'q-sel', title, 'aria-label': title, style: width ? { width } : undefined })
        for (const o of options) s.append(el('option', { value: o.value }, o.label))
        s.addEventListener('change', () => onChange(s.value))
        this.root.append(s)
        return s
      }
      color (title, glyph, onPick, initial) {
        const input = el('input', { type: 'color', value: '#' + initial })
        const swatch = el('span', { class: 'q-sw', style: { background: '#' + initial } })
        const wrap = el('label', { class: 'q-btn q-color', title, 'aria-label': title }, el('span', { class: 'q-glyph' }, glyph), swatch, input)
        wrap.addEventListener('mousedown', (e) => { if (e.target !== input) e.preventDefault() })
        input.addEventListener('input', () => { swatch.style.background = input.value })
        input.addEventListener('change', () => onPick(input.value.slice(1).toUpperCase()))
        this.root.append(wrap)
        return wrap
      }
      sep () { this.root.append(el('span', { class: 'q-sep' })) }
      space () { this.root.append(el('span', { class: 'q-space' })) }
    }

    // ------------------------------------------------------------ pins
    // A pin is a note about one spot in the file for the Agent. Pins wait
    // on the dsh server (api/quire/pins), per Session, with a picture of
    // the boxed area, and fold into a pill above the chat composer
    // (QuirePinDock), never in the draft. The next message the person
    // sends takes them and their pictures along to the Agent, and the
    // server forgets them. Spots use the numbering office_read gives the
    // Agent: 汇总!B4, [p12], 第 3 页 [s3.2].
    const oneLine = (x) => String(x ?? '').replace(/[\s\u200b]+/g, ' ').trim()
    const clipPin = (x, n) => { const t = oneLine(x); return t.length > n ? t.slice(0, n) + '…' : t }

    // pinStore mirrors the server's pins for the Sessions on screen; the
    // editors and the composer dock both read it.
    const NO_PINS = { rev: 0, items: [] }
    const pinStore = {
      data: new Map(),
      subs: new Set(),
      frames: new Set(),
      get (sid) { return this.data.get(sid) ?? NO_PINS },
      set (sid, r) {
        if (!r || !Array.isArray(r.items)) return
        this.data.set(sid, { rev: r.rev, items: r.items })
        for (const f of this.subs) f(sid)
      },
      async call (sid, op, extra) {
        const r = await api('pins', { sessionId: sid, op, ...extra })
        this.set(sid, r)
        return r
      },
      load (sid) { return sid ? this.call(sid, 'list').catch(() => {}) : Promise.resolve() },
      subscribe (f) { this.subs.add(f); return () => this.subs.delete(f) },
      // reveal shows a pin in an open editor on its file.
      reveal (sid, pin) {
        for (const fr of this.frames) if (fr.doc?.file?.sessionId === sid && fr.doc.file.path === pin.path) fr.editor?.revealPin?.(pin)
      }
    }

    // textRange finds text inside an element, across its text nodes.
    function textRange (root, needle) {
      if (!needle) return null
      const nodes = []
      let all = ''
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      for (let n = w.nextNode(); n; n = w.nextNode()) { nodes.push([n, all.length]); all += n.data }
      const flat = all.replace(/\u200b/g, ' ')
      let at = flat.indexOf(needle)
      if (at < 0) {
        // The pin line squeezed white space; match it loosely.
        const re = new RegExp(needle.split(/\s+/).map((x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s+'))
        const m = re.exec(flat)
        if (!m) return null
        at = m.index; needle = m[0]
      }
      const end = at + needle.length
      const pos = (x) => { for (let i = nodes.length - 1; i >= 0; i--) if (nodes[i][1] <= x) return [nodes[i][0], Math.min(x - nodes[i][1], nodes[i][0].data.length)]; return null }
      const a = pos(at); const b = pos(end)
      if (!a || !b) return null
      const r = document.createRange()
      r.setStart(a[0], a[1]); r.setEnd(b[0], b[1])
      return r
    }

    // Pins in every open document share one highlight.
    const pinRanges = new Map()
    function paintPinRanges () {
      const H = window.CSS?.highlights
      if (!H || typeof Highlight === 'undefined') return
      const all = [...pinRanges.values()].flat()
      if (all.length) H.set('q-pin', new Highlight(...all)); else H.delete('q-pin')
    }

    // menu shows a context menu at (x, y). Items are { label, fn, kbd,
    // disabled } or '-' for a separator.
    // ------------------------------------------------------- snapshots
    // A pin carries a picture of the spot. The sheet is a canvas and is
    // cropped; documents and slides are DOM, copied with every computed
    // style written inline into an SVG foreignObject and drawn onto a
    // canvas. Pseudo-element counters and web fonts may come out plainer
    // than on screen; the layout and text are exact.
    const SNAP_EDGE = 1600
    function cloneStyled (src) {
      if (src.nodeType === 3) return document.createTextNode(src.data)
      if (src.nodeType !== 1) return null
      const tag = src.localName
      if (tag === 'script' || tag === 'style' || tag === 'link') return null
      const cs = getComputedStyle(src)
      if (cs.display === 'none') return null
      if (tag === 'canvas') {
        const img = document.createElement('img')
        try { img.src = src.toDataURL() } catch { return null }
        img.setAttribute('style', `display:${cs.display};width:${cs.width};height:${cs.height}`)
        return img
      }
      const c = src.cloneNode(false)
      let css = ''
      for (let i = 0; i < cs.length; i++) { const k = cs[i]; css += `${k}:${cs.getPropertyValue(k)};` }
      c.setAttribute('style', css)
      for (const a of ['class', 'id', 'contenteditable', 'tabindex']) c.removeAttribute?.(a)
      if (tag === 'img' && src.currentSrc) c.setAttribute('src', src.currentSrc)
      if (tag === 'textarea' || tag === 'input') c.textContent = src.value
      for (const pseudo of ['::before', '::after']) {
        const ps = getComputedStyle(src, pseudo)
        const m = /^"(.*)"$/.exec(ps.content ?? '')
        if (!m || tag === 'img') continue
        const span = document.createElement('span')
        let pcss = ''
        for (let i = 0; i < ps.length; i++) { const k = ps[i]; if (k !== 'content') pcss += `${k}:${ps.getPropertyValue(k)};` }
        span.setAttribute('style', pcss)
        span.textContent = m[1].replace(/\\"/g, '"')
        if (pseudo === '::before') c.prepend(span); else c.append(span)
      }
      for (const k of src.childNodes) { const x = cloneStyled(k); if (x) c.append(x) }
      return c
    }

    const dataUrlOf = (url) => fetch(url, { credentials: 'same-origin' }).then((r) => r.ok ? r.blob() : Promise.reject(new Error(r.status))).then((b) => new Promise((resolve, reject) => {
      const fr = new FileReader(); fr.onload = () => resolve(String(fr.result)); fr.onerror = () => reject(fr.error); fr.readAsDataURL(b)
    }))

    // domSnap draws the part of the screen inside rect, showing only
    // nodes (each copied where it sits) over bg.
    async function domSnap (nodes, rect, bg) {
      const W = Math.ceil(rect.right - rect.left); const H = Math.ceil(rect.bottom - rect.top)
      if (W < 2 || H < 2) return null
      const XHTML = 'http://www.w3.org/1999/xhtml'
      const wrap = document.createElementNS(XHTML, 'div')
      wrap.setAttribute('style', `position:relative;width:${W}px;height:${H}px;overflow:hidden;background:${bg || '#fff'}`)
      for (const n of nodes) {
        const r = n.getBoundingClientRect()
        const c = cloneStyled(n)
        if (!c) continue
        const w = n.offsetWidth || r.width; const h = n.offsetHeight || r.height
        c.style.setProperty('position', 'absolute')
        c.style.setProperty('margin', '0')
        c.style.setProperty('left', r.left - rect.left + 'px')
        c.style.setProperty('top', r.top - rect.top + 'px')
        c.style.setProperty('width', w + 'px')
        c.style.setProperty('height', h + 'px')
        c.style.setProperty('transform-origin', '0 0')
        c.style.setProperty('transform', `scale(${r.width / w || 1}, ${r.height / h || 1})`)
        wrap.append(c)
      }
      let xml = new XMLSerializer().serializeToString(wrap)
      // An SVG image loads nothing from outside: inline the pictures.
      const urls = new Set()
      for (const m of xml.matchAll(/(?:src|href)="((?!data:)[^"]+)"|url\(&quot;((?!data:)[^&]+)&quot;\)|url\("((?!data:)[^"]+)"\)/g)) urls.add((m[1] ?? m[2] ?? m[3]).replace(/&amp;/g, '&'))
      for (const u of urls) {
        const d = await dataUrlOf(u).catch(() => '')
        xml = xml.split(u.replace(/&/g, '&amp;')).join(d)
      }
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><foreignObject x="0" y="0" width="100%" height="100%">${xml}</foreignObject></svg>`
      const img = new Image()
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg)
      await img.decode()
      const k = Math.min(Math.max(1, window.devicePixelRatio || 1), 2, SNAP_EDGE / Math.max(W, H))
      const cv = document.createElement('canvas')
      cv.width = Math.round(W * k); cv.height = Math.round(H * k)
      const g = cv.getContext('2d')
      g.scale(k, k)
      g.drawImage(img, 0, 0)
      return cv
    }

    // canvasCrop copies the part of a canvas under rect (screen pixels).
    function canvasCrop (src, rect) {
      const cr = src.getBoundingClientRect()
      const sx = src.width / cr.width; const sy = src.height / cr.height
      const x = Math.max(0, (rect.left - cr.left) * sx); const y = Math.max(0, (rect.top - cr.top) * sy)
      const w = Math.min(src.width - x, (rect.right - rect.left) * sx); const h = Math.min(src.height - y, (rect.bottom - rect.top) * sy)
      if (w < 2 || h < 2) return null
      const k = Math.min(1, SNAP_EDGE / Math.max(w, h))
      const cv = document.createElement('canvas')
      cv.width = Math.round(w * k); cv.height = Math.round(h * k)
      cv.getContext('2d').drawImage(src, x, y, w, h, 0, 0, cv.width, cv.height)
      return cv
    }

    // snapOut turns a snapshot into what a pin stores: the PNG for the
    // Agent and a small preview for the composer.
    function snapOut (cv) {
      if (!cv) return null
      let data = cv.toDataURL('image/png'); let mediaType = 'image/png'
      if (data.length > 6e6) { data = cv.toDataURL('image/jpeg', 0.9); mediaType = 'image/jpeg' }
      const k = Math.min(1, 240 / Math.max(cv.width, cv.height))
      const t = document.createElement('canvas')
      t.width = Math.max(1, Math.round(cv.width * k)); t.height = Math.max(1, Math.round(cv.height * k))
      const g = t.getContext('2d')
      g.fillStyle = '#fff'; g.fillRect(0, 0, t.width, t.height)
      g.drawImage(cv, 0, 0, t.width, t.height)
      return { image: { mediaType, data: data.slice(data.indexOf(',') + 1) }, thumb: t.toDataURL('image/jpeg', 0.85), width: cv.width, height: cv.height }
    }

    // pinWhere is a spot as people read it: a part of a slide is shown
    // by its picture, not by percentages.
    const pinWhere = (where) => String(where ?? '').replace(/ 区域 [\d%,.–-]+$/, ' 框选区域')

    function menu (x, y, items) {
      document.querySelector('.q-menu')?.remove()
      const m = el('div', { class: 'q-menu', role: 'menu' })
      let last = '-'
      for (const it of items) {
        if (!it) continue
        if (it === '-') { if (last !== '-') m.append(el('div', { class: 'q-menu-sep' })); last = '-'; continue }
        const b = el('button', { type: 'button', role: 'menuitem', disabled: it.disabled }, el('span', { class: 'q-mi' }, it.pre ?? null, it.label), it.kbd ? el('kbd', {}, it.kbd) : null)
        b.addEventListener('mousedown', (e) => e.preventDefault())
        b.addEventListener('click', () => { m.remove(); it.fn() })
        m.append(b)
        last = it
      }
      if (m.lastChild?.classList.contains('q-menu-sep')) m.lastChild.remove()
      document.body.append(m)
      const r = m.getBoundingClientRect()
      m.style.left = Math.max(4, Math.min(x, window.innerWidth - r.width - 4)) + 'px'
      m.style.top = Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) + 'px'
      const close = (ev) => { if (!m.contains(ev.target)) { m.remove(); document.removeEventListener('pointerdown', close, true); document.removeEventListener('keydown', esc, true) } }
      const esc = (ev) => { if (ev.key === 'Escape') { ev.stopPropagation(); m.remove(); document.removeEventListener('pointerdown', close, true); document.removeEventListener('keydown', esc, true) } }
      setTimeout(() => { document.addEventListener('pointerdown', close, true); document.addEventListener('keydown', esc, true) })
      return m
    }

    // dialog shows a small modal form and resolves with its values, or null
    // when cancelled. fields: { key, label, type, options, value,
    // placeholder, hint, show(values) }; body adds content of its own and
    // read() gives what to resolve with instead; check returns a message to
    // keep the dialog open.
    function dialog ({ title, fields = [], body, read, ok = '确定', check, wide, buttons = [] }) {
      return new Promise((resolve) => {
        document.querySelector('.q-dlg-back')?.remove()
        const back = el('div', { class: 'q-dlg-back' })
        const box = el('form', { class: 'q-dlg' + (wide ? ' q-dlg-wide' : ''), role: 'dialog', 'aria-label': title, novalidate: true })
        box.append(el('div', { class: 'q-dlg-h' }, title))
        const inputs = {}
        const rows = []
        const values = () => Object.fromEntries(Object.entries(inputs).map(([k, n]) => [k, n.type === 'checkbox' ? n.checked : n.value]))
        for (const f of fields) {
          let n
          if (f.type === 'select') n = el('select', { class: 'q-sel' }, ...f.options.map(([v, l]) => el('option', { value: v, selected: String(f.value ?? '') === String(v) }, l)))
          else if (f.type === 'textarea') n = el('textarea', { rows: 3, placeholder: f.placeholder ?? '', spellcheck: 'false' })
          else n = el('input', { type: f.type ?? 'text', placeholder: f.placeholder ?? '', spellcheck: 'false', autocomplete: 'off' })
          if (f.type === 'checkbox') n.checked = !!f.value
          else if (f.type !== 'select') n.value = f.value ?? ''
          n.name = f.key
          inputs[f.key] = n
          const row = el('label', { class: 'q-dlg-row' + (f.type === 'checkbox' ? ' q-dlg-check' : '') }, el('span', {}, f.label), n, f.hint ? el('small', {}, f.hint) : null)
          rows.push([f, row])
          box.append(row)
        }
        const sync = () => { const v = values(); for (const [f, row] of rows) if (f.show) row.hidden = !f.show(v) }
        box.addEventListener('input', sync)
        box.addEventListener('change', sync)
        sync()
        if (body) box.append(typeof body === 'function' ? body((v) => done(v)) : body)
        const err = el('div', { class: 'q-dlg-err', hidden: true })
        const done = (v) => { back.remove(); document.removeEventListener('keydown', onKey, true); resolve(v) }
        const okBtn = el('button', { type: 'submit', class: 'q-dlg-ok' }, ok)
        const cancel = el('button', { type: 'button' }, '取消')
        cancel.addEventListener('click', () => done(null))
        const extra = buttons.map(([label, fn]) => { const b = el('button', { type: 'button', class: 'q-dlg-x' }, label); b.addEventListener('click', () => fn(done)); return b })
        box.append(err, el('div', { class: 'q-dlg-f' }, ...extra, el('span', { class: 'q-space' }), cancel, okBtn))
        box.addEventListener('submit', (e) => {
          e.preventDefault()
          const v = read ? read() : values()
          const why = check?.(v)
          if (why) { err.textContent = why; err.hidden = false; return }
          done(v)
        })
        const onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); done(null) } }
        document.addEventListener('keydown', onKey, true)
        back.addEventListener('pointerdown', (e) => { if (e.target === back) done(null) })
        back.append(box)
        document.body.append(back)
        const first = box.querySelector('input:not([type=checkbox]),select,textarea')
        first?.focus()
        first?.select?.()
      })
    }

    // pickImage asks for a picture on the person's computer and resolves
    // with it as a data: URI.
    function pickImage () {
      return new Promise((resolve) => {
        const input = el('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif', style: { display: 'none' } })
        input.addEventListener('change', () => { const f = input.files?.[0]; input.remove(); if (f) readImage(f).then(resolve, () => resolve(undefined)); else resolve(undefined) })
        document.body.append(input)
        input.click()
      })
    }
    function readImage (file) {
      return new Promise((resolve, reject) => {
        if (!/^image\/(png|jpe?g|gif)$/i.test(file.type)) { reject(new Error('只能插入 PNG、JPEG、GIF 图片')); return }
        if (file.size > 20 * 1024 * 1024) { reject(new Error('图片超过 20 MB')); return }
        const r = new FileReader()
        r.onload = () => resolve(String(r.result))
        r.onerror = () => reject(r.error)
        r.readAsDataURL(file)
      })
    }
    const imageFile = (dt) => [...(dt?.files ?? [])].find((f) => /^image\/(png|jpe?g|gif)$/i.test(f.type))

    // FONTS are the faces offered in the font menus: common on Windows,
    // macOS and in WPS, Chinese first.
    const FONTS = ['宋体', '黑体', '微软雅黑', '楷体', '仿宋', '等线', 'Arial', 'Calibri', 'Times New Roman', 'Helvetica', 'Georgia', 'Verdana', 'Courier New']
    const fontOptions = () => [{ value: '', label: '字体' }, ...FONTS.map((f) => ({ value: f, label: f }))]
    // fontCss names a face with fallbacks the browser can find.
    const FONT_ALIAS = { 宋体: 'SimSun, "Songti SC", STSong', 黑体: 'SimHei, "Heiti SC", STHeiti', 微软雅黑: '"Microsoft YaHei", "PingFang SC"', 楷体: 'KaiTi, "Kaiti SC", STKaiti', 仿宋: 'FangSong, STFangsong', 等线: 'DengXian, "PingFang SC"', Calibri: 'Carlito, "Helvetica Neue", Arial, sans-serif', 'Calibri Light': 'Carlito, "Helvetica Neue", Arial, sans-serif', Cambria: 'Caladea, Georgia, serif', Arial: 'Helvetica, sans-serif' }
    const fontCss = (f) => `"${f.replace(/"/g, '')}"${FONT_ALIAS[f] ? ', ' + FONT_ALIAS[f] : ''}`
    const PRINT_CSS = '@media print{.bar{display:none!important}html,body{background:#fff!important}.slide{box-shadow:none!important}.slide-wrap{break-inside:avoid;page-break-inside:avoid}}'

    // ------------------------------------------------------------ document session

    // Doc owns one open file: its version, saving, undo and redo, and
    // changes made elsewhere. The editor supplies collect(), load() and
    // saved().
    // hostCtx is dsh's client context, for opening files in the sidebar.
    let hostCtx
    const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent || '')
    // keyName writes a shortcut the way this computer names its keys.
    const keyName = (s) => isMac ? s : s.replace(/⇧/g, 'Shift+').replace(/⌘/g, 'Ctrl+').replace(/⌥/g, 'Alt+')

    // zoomSlider is the status-bar zoom: − slider + and the percentage,
    // which fits the view when clicked.
    function zoomSlider (set, fit) {
      const steps = [10, 25, 50, 67, 75, 80, 90, 100, 110, 125, 150, 175, 200, 250, 300, 400]
      let cur = 100
      const go = (v) => set(Math.max(10, Math.min(400, Math.round(v))))
      const btn = (label, title, fn, cls) => { const b = el('button', { type: 'button', title, class: cls }, label); b.addEventListener('mousedown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b }
      const range = el('input', { type: 'range', min: '10', max: '400', step: '5', value: '100', class: 'q-zoom-r', 'aria-label': '缩放' })
      range.addEventListener('input', () => go(+range.value))
      const pct = btn('100%', fit ? '适应窗口' : '恢复 100%', () => (fit ? fit() : go(100)), 'q-zoom-p')
      const node = el('span', { class: 'q-zoom' }, btn('−', '缩小', () => go([...steps].reverse().find((x) => x < cur - 0.5) ?? 10)), range, btn('+', '放大', () => go(steps.find((x) => x > cur + 0.5) ?? 400)), pct)
      return { node, update (v) { cur = v; range.value = String(Math.round(v)); pct.textContent = Math.round(v) + '%' } }
    }

    // syncZoomSel shows a zoom in a toolbar zoom list, adding odd values.
    function syncZoomSel (sel, v) {
      if (!sel) return
      if (!v) { if ([...sel.options].some((o) => o.value === '')) sel.value = ''; return }
      const k = String(Math.round(v))
      if (![...sel.options].some((o) => o.value === k)) sel.append(el('option', { value: k }, k + '%'))
      sel.value = k
    }

    // fileTab is the 文件 tab every editor starts with.
    function fileTab (tb, frame) {
      tb.tab('文件')
      tb.button('新建', '新建空白的 Word 文档、Excel 表格或 PowerPoint 演示文稿', (e) => {
        const r = e.currentTarget.getBoundingClientRect()
        menu(r.left, r.bottom + 2, [
          { label: '空白 Word 文档', pre: el('b', { class: 'q-kind q-kind-docx' }, 'W'), fn: () => frame.newFile('docx') },
          { label: '空白 Excel 表格', pre: el('b', { class: 'q-kind q-kind-xlsx' }, 'X'), fn: () => frame.newFile('xlsx') },
          { label: '空白 PowerPoint 演示文稿', pre: el('b', { class: 'q-kind q-kind-pptx' }, 'P'), fn: () => frame.newFile('pptx') }
        ])
      }, { ico: 'plus' })
      tb.button('另存副本', '把当前文件另存一份副本并打开它', () => frame.saveCopy(), { ico: 'copy' })
      tb.sep()
      tb.button('导出 PDF', '打开打印对话框，选择“存储为 PDF”', () => frame.print(true), { ico: 'print' })
      tb.button('打印', '打印（⌘P）', () => frame.print(), { ico: 'print' })
      tb.sep()
      tb.button('键盘快捷键', '查看快捷键（⌘/）', () => frame.shortcuts())
    }

    class Doc {
      constructor (frame, file) {
        this.frame = frame
        this.file = file
        this.version = undefined
        this.saving = false
        this.timer = undefined
        this.idle = []
        this.disposed = false
      }

      async open () {
        const r = await api('open', this.file)
        this.version = r.version
        this.writable = r.writable
        this.why = r.why
        this.history(r)
        return r
      }

      history (r) { this.frame.history(r.undo ?? 0, r.redo ?? 0, r.undoLabels ?? [], r.redoLabels ?? []) }

      // dirty asks for a save soon; now saves at once.
      dirty (now) {
        if (!this.writable) return
        this.frame.status('pending')
        clearTimeout(this.timer)
        this.timer = setTimeout(() => this.kick(), now ? 0 : 700)
      }

      async kick () {
        if (this.saving || this.disposed) return
        const batch = this.editor.collect()
        if (!batch || !batch.ops.length) { this.frame.status('saved'); this.settle(); return }
        this.saving = true
        this.frame.status('saving')
        try {
          const r = await api('save', { ...this.file, version: this.version, ops: batch.ops, ...(batch.steps ? { steps: batch.steps } : {}) })
          this.version = r.version
          this.history(r)
          this.editor.saved(r.model, batch)
          this.frame.status('saved')
        } catch (err) {
          if (err.code === 'conflict') {
            this.version = err.body.version
            this.editor.load(err.body.model, { keep: true })
            this.frame.toast('文件已在别处被修改（可能是 Agent），已载入最新内容；你最后一次的改动没有保存。', 'warn')
            this.frame.status('saved')
          } else {
            this.frame.status('error', err.message)
            this.editor.failed?.(batch, err)
          }
        } finally {
          this.saving = false
          if (this.editor.hasChanges?.()) this.dirty(true)
          else this.settle()
        }
      }

      settle () { for (const f of this.idle.splice(0)) f() }

      // flush saves what is pending and resolves when nothing is in flight.
      flush () {
        clearTimeout(this.timer)
        return new Promise((resolve) => {
          if (!this.saving && !this.editor.hasChanges?.()) return resolve()
          this.idle.push(resolve)
          if (!this.saving) this.kick()
        })
      }

      // step undoes or redoes count steps.
      async step (dir, count = 1) {
        if (!this.writable) return
        await this.flush()
        try {
          const r = await api(dir, { ...this.file, count })
          this.version = r.version
          this.history(r)
          this.editor.load(r.model, { keep: true })
        } catch (err) {
          this.frame.toast(err.message, err.code === 'nothing' ? 'info' : 'warn')
          if (err.code === 'stale') this.frame.history(0, 0)
        }
      }

      // remote checks whether the file changed elsewhere and, when it did
      // and nothing local is pending, shows the new content.
      async remote () {
        if (!this.editor || this.saving || this.editor.hasChanges?.() || this.disposed) return
        let r
        try { r = await api('open', this.file) } catch { return }
        if (r.version === this.version || this.saving || this.editor.hasChanges?.()) return
        this.version = r.version
        this.writable = r.writable
        this.why = r.why
        this.history(r)
        this.editor.load(r.model, { keep: true })
        this.frame.toast('已载入文件的最新内容。', 'info')
      }
    }

    // Frame is the chrome around an editor: toolbar row, status, undo and
    // redo, full screen, banners and toasts.
    class Frame {
      constructor (host) {
        this.root = el('div', { class: 'quire', tabindex: '-1' })
        this.strip = el('div', { class: 'q-tabs', role: 'tablist' })
        this.side = el('div', { class: 'q-side' })
        this.head = el('div', { class: 'q-head' }, this.strip, this.side)
        this.tools = el('div', { class: 'q-tools' })
        this.bar = el('div', { class: 'q-bar' }, this.tools)
        this.banner = el('div', { class: 'q-banner', hidden: true })
        this.main = el('div', { class: 'q-main' })
        this.toasts = el('div', { class: 'q-toasts' })
        this.buildFind()
        this.root.append(this.head, this.bar, this.findBar, this.banner, this.main, this.toasts)
        host.append(this.root)
        const tb = new Toolbar(this.side)
        this.stat = el('span', { class: 'q-status' })
        this.side.append(this.stat)
        // Pins are how the person talks to the Agent about the file, so
        // the button stands out: labelled, in its own colour, first.
        this.pinBtn = tb.button('标注', '框选要标注的部分，写下问题，随下一条消息发给 Agent（⇧⌘M）', () => this.pin(), { ico: 'pin', cls: 'q-pin' })
        this.pinN = el('span', { class: 'q-pin-n', hidden: true })
        this.pinBtn.append(this.pinN)
        this.pinMore = tb.button('▾', '待发送的标注', (e) => this.pinMenu(e), { cls: 'q-more q-pin-more' })
        this.pinMore.hidden = true
        this.pins = []
        tb.sep()
        this.undoBtn = tb.button('', '撤销（⌘Z）', () => this.doc?.step('undo'), { ico: 'undo' })
        this.undoMore = tb.button('▾', '撤销多步', (e) => this.historyMenu(e, 'undo'), { cls: 'q-more' })
        this.redoBtn = tb.button('', '重做（⇧⌘Z）', () => this.doc?.step('redo'), { ico: 'redo' })
        this.redoMore = tb.button('▾', '重做多步', (e) => this.historyMenu(e, 'redo'), { cls: 'q-more' })
        tb.button('', '查找和替换（⌘F）', () => this.openFind(), { ico: 'search' })
        tb.button('', '打印 / 导出 PDF（⌘P）', () => this.print(), { ico: 'print' })
        this.fullBtn = tb.button('', '全屏', () => this.full(), { ico: 'full' })
        this.history(0, 0)
        this.onKey = (e) => { if (e.key === 'Escape' && this.root.classList.contains('q-full') && !document.querySelector('.q-menu') && !this.pinPopEl && !this.pinArm) { this.full(false) } }
        document.addEventListener('keydown', this.onKey)
        // Editor-wide shortcuts, ahead of the editors' own key handling.
        this.root.addEventListener('keydown', (e) => {
          if (!mod(e) || e.altKey) return
          const k = e.key.toLowerCase()
          if (k === 'm' && e.shiftKey) { e.preventDefault(); e.stopPropagation(); if (!this.pinPopEl) this.pin() } else if (k === 'f' || (k === 'h' && e.shiftKey)) { e.preventDefault(); e.stopPropagation(); this.openFind() } else if (k === 'p') { e.preventDefault(); e.stopPropagation(); this.print() } else if (k === 's') { e.preventDefault(); e.stopPropagation(); this.saveNow() } else if (k === '/') { e.preventDefault(); e.stopPropagation(); this.shortcuts() }
        }, true)
      }

      // ------------------------------------------------ find and replace

      buildFind () {
        this.findInput = el('input', { class: 'q-find-in', placeholder: '查找', 'aria-label': '查找', spellcheck: 'false' })
        this.replInput = el('input', { class: 'q-find-in', placeholder: '替换为', 'aria-label': '替换为', spellcheck: 'false' })
        this.findCount = el('span', { class: 'q-find-n' })
        const btn = (label, title, fn) => { const b = el('button', { type: 'button', class: 'q-btn q-find-b', title }, label); b.addEventListener('mousedown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b }
        this.replWrap = el('span', { class: 'q-find-r' }, this.replInput, btn('替换', '替换当前', () => this.replaceOne()), btn('全部替换', '全部替换', () => this.replaceAll()))
        this.findBar = el('div', { class: 'q-find', hidden: true }, this.findInput, btn('↑', '上一个（⇧⏎）', () => this.step(-1)), btn('↓', '下一个（⏎）', () => this.step(1)), this.findCount, this.replWrap, btn('✕', '关闭（Esc）', () => this.closeFind()))
        this.findInput.addEventListener('input', () => { clearTimeout(this.findTimer); this.findTimer = setTimeout(() => this.search(), 150) })
        const keys = (e) => {
          if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.closeFind() } else if (e.key === 'Enter' && e.target === this.findInput) { e.preventDefault(); this.step(e.shiftKey ? -1 : 1) } else if (e.key === 'Enter') { e.preventDefault(); this.replaceOne() }
        }
        this.findInput.addEventListener('keydown', keys)
        this.replInput.addEventListener('keydown', keys)
        this.matches = []
        this.at = -1
      }

      openFind () {
        if (!this.editor?.findAll) return
        this.findBar.hidden = false
        this.replWrap.hidden = !this.doc?.writable
        const sel = this.editor.selectedText?.()
        if (sel && !sel.includes('\n') && sel.length < 80) this.findInput.value = sel
        this.findInput.focus()
        this.findInput.select()
        this.search()
      }

      closeFind () {
        this.findBar.hidden = true
        this.matches = []
        this.editor?.focus?.()
      }

      search (keepAt) {
        const q = this.findInput.value
        this.matches = q ? this.editor.findAll(q) : []
        if (!this.matches.length) { this.at = -1; this.findCount.textContent = q ? '无结果' : ''; return }
        this.at = keepAt !== undefined ? Math.min(keepAt, this.matches.length - 1) : 0
        this.show()
      }

      show () {
        this.findCount.textContent = `${this.at + 1}/${this.matches.length}`
        this.editor.showMatch(this.matches[this.at])
        this.findInput.focus({ preventScroll: true })
      }

      step (d) {
        if (!this.matches.length) { this.search(); return }
        this.at = (this.at + d + this.matches.length) % this.matches.length
        this.show()
      }

      replaceOne () {
        if (!this.doc?.writable || !this.matches.length) return
        this.editor.replaceMatch(this.matches[this.at], this.findInput.value, this.replInput.value)
        setTimeout(() => this.search(this.at), 30)
      }

      replaceAll () {
        if (!this.doc?.writable || !this.findInput.value) return
        const n = this.editor.replaceAll(this.findInput.value, this.replInput.value)
        this.toast(n ? `已替换 ${n} 处。` : '没有找到。', 'info')
        setTimeout(() => this.search(), 30)
      }

      // ------------------------------------------------ print

      async print (pdf) {
        if (!this.doc) return
        this.toast(pdf ? '在打印对话框的“目标打印机”里选择“存储为 PDF”。' : '正在准备打印…', 'info')
        try {
          await this.doc.flush()
          const r = await api('print', this.doc.file)
          const f = el('iframe', { class: 'q-print', title: '打印' })
          document.body.append(f)
          f.addEventListener('load', () => {
            try { f.contentWindow.focus(); f.contentWindow.print() } catch (err) { this.toast('无法打开打印对话框：' + err.message, 'warn') }
            setTimeout(() => f.remove(), 60000)
          }, { once: true })
          f.srcdoc = r.html.replace('</style>', PRINT_CSS + '</style>')
        } catch (err) { this.toast('无法打印：' + err.message, 'warn') }
      }

      history (u, r, ul = [], rl = []) {
        this.labels = { undo: ul, redo: rl }
        this.undoBtn.disabled = this.undoMore.disabled = !u
        this.redoBtn.disabled = this.redoMore.disabled = !r
        this.undoBtn.title = ul[0] ? `撤销：${ul[0]}（${keyName('⌘Z')}）` : `撤销（${keyName('⌘Z')}）`
        this.redoBtn.title = rl[0] ? `重做：${rl[0]}（${keyName('⇧⌘Z')}）` : `重做（${keyName('⇧⌘Z')}）`
      }

      // historyMenu lists the steps to undo or redo, most recent first;
      // picking one goes back (or forward) through it.
      historyMenu (e, dir) {
        const list = this.labels?.[dir] ?? []
        if (!list.length || !this.doc) return
        const r = e.currentTarget.getBoundingClientRect()
        menu(r.left - 20, r.bottom + 2, list.map((l, i) => ({ label: `${dir === 'undo' ? '撤销' : '重做'} ${i + 1} 步：${l}`, fn: () => this.doc.step(dir, i + 1) })))
      }

      saveNow () {
        if (!this.doc?.writable) return
        this.doc.flush().then(() => this.toast('已保存。', 'info'))
      }

      shortcuts () {
        const common = [['⌘Z', '撤销'], ['⇧⌘Z / ⌘Y', '重做'], ['⌘S', '立即保存'], ['⌘F', '查找'], ['⇧⌘H', '查找和替换'], ['⌘P', '打印 / 导出 PDF'], ['⇧⌘M', '标注问题，加入聊天框'], ['⌘/', '键盘快捷键'], ['Esc', '退出全屏']]
        const own = this.editor?.shortcuts?.() ?? []
        const table = (rows) => el('table', { class: 'q-keys' }, ...rows.map(([k, d]) => el('tr', {}, el('td', {}, ...k.split(' / ').flatMap((x, i) => [i ? ' / ' : null, el('kbd', {}, keyName(x))])), el('td', {}, d))))
        dialog({ title: '键盘快捷键', wide: true, ok: '关闭', body: el('div', { class: 'q-keys-wrap' }, own.length ? el('h4', {}, own.title) : null, own.length ? table(own) : null, el('h4', {}, '通用'), table(common)) })
      }

      // ------------------------------------------------ files

      // fileName asks for a name until the server takes it.
      async fileName (title, value, hint, make) {
        let why = ''
        for (;;) {
          const v = await dialog({ title, fields: [{ key: 'name', label: '文件名', value, hint: why || hint }], check: (x) => x.name.trim() ? '' : '请输入文件名' })
          if (!v) return
          try { return await make(v.name.trim()) } catch (err) {
            if (err.code !== 'exists' && err.code !== 'bad-name') { this.toast(err.message, 'warn'); return }
            why = err.message
            value = v.name
          }
        }
      }

      async newFile (kind) {
        if (!this.doc) return
        const name = { docx: '新建文档', xlsx: '新建表格', pptx: '新建演示文稿' }[kind]
        const r = await this.fileName({ docx: '新建 Word 文档', xlsx: '新建 Excel 表格', pptx: '新建 PowerPoint 演示文稿' }[kind], `${name}.${kind}`, '保存在当前文件所在的文件夹', (n) => api('new', { ...this.doc.file, kind, name: n }))
        if (r) this.openFile(r.path)
      }

      async saveCopy () {
        if (!this.doc) return
        const base = this.doc.file.path.split('/').pop()
        const ext = base.split('.').pop()
        await this.doc.flush()
        const r = await this.fileName('另存副本', `${base.slice(0, -ext.length - 1)} - 副本.${ext}`, '副本保存在同一个文件夹，原文件保持不变', (n) => api('copy', { ...this.doc.file, name: n }))
        if (r) this.openFile(r.path)
      }

      // openFile shows another workspace file in a sidebar tab of its own.
      openFile (path) {
        const address = 'dsh-resource://file/session/' + encodeURIComponent(this.doc.file.sessionId) + '/' + path.split('/').map(encodeURIComponent).join('/')
        const side = hostCtx?.get?.('sidebarRight')
        try { if (side?.openResource) { side.openResource(address); this.toast(`已创建 ${path}`, 'info'); return } } catch {}
        this.toast(`已创建 ${path}，在“文件”里打开它。`, 'info')
      }

      status (state, msg) {
        const text = { pending: '有改动…', saving: '保存中…', saved: '已保存', error: '保存失败' }[state] ?? ''
        this.stat.textContent = text
        this.stat.dataset.state = state
        this.stat.title = msg ?? ''
        if (state === 'error') this.toast(`保存失败：${msg}`, 'warn')
      }

      readOnly (why) {
        this.banner.hidden = !why
        this.banner.textContent = why ? `只读：${why}` : ''
        this.undoBtn.hidden = this.redoBtn.hidden = this.undoMore.hidden = this.redoMore.hidden = !!why
        this.stat.hidden = !!why
      }

      full (on = !this.root.classList.contains('q-full')) {
        this.root.classList.toggle('q-full', on)
        this.fullBtn.innerHTML = icon(on ? 'exit' : 'full')
        this.fullBtn.title = on ? '退出全屏（Esc）' : '全屏'
        this.editor?.resize?.()
      }

      toast (msg, kind = 'info') {
        const t = el('div', { class: 'q-toast q-' + kind }, msg)
        this.toasts.append(t)
        setTimeout(() => t.remove(), kind === 'warn' ? 7000 : 3500)
      }

      fail (msg) {
        this.main.replaceChildren(el('div', { class: 'q-fail' }, msg))
      }

      // ------------------------------------------------ pins

      // pin arms the box tool: the next drag over the file marks the
      // part to talk about, a plain click marks the one spot under it.
      pin () {
        if (this.pinArm) { this.pinDisarm(); return }
        if (!this.editor?.pinAt) { this.pinHere(); return }
        this.pinClose()
        const main = this.main
        const inside = (e) => main.contains(e.target)
        const block = (e) => { if (inside(e)) { e.preventDefault(); e.stopPropagation() } }
        const tip = el('div', { class: 'q-pintip' }, '拖动框选要标注的部分，单击标注一处　Esc 退出')
        this.root.append(tip)
        const R = this.root.getBoundingClientRect(); let M = main.getBoundingClientRect()
        tip.style.top = M.top - R.top + 10 + 'px'
        const clamp = (x, y) => [Math.max(M.left, Math.min(x, M.right)), Math.max(M.top, Math.min(y, M.bottom))]
        const down = (e) => {
          if (!inside(e)) return
          block(e)
          if (e.button !== 0) return
          M = main.getBoundingClientRect()
          const [ox, oy] = clamp(e.clientX, e.clientY)
          const marq = el('div', { class: 'q-pinmarq', hidden: true })
          this.root.append(marq)
          let box = { left: ox, top: oy, right: ox, bottom: oy }
          const move = (ev) => {
            const [x, y] = clamp(ev.clientX, ev.clientY)
            box = { left: Math.min(ox, x), top: Math.min(oy, y), right: Math.max(ox, x), bottom: Math.max(oy, y) }
            const r0 = this.root.getBoundingClientRect()
            marq.hidden = box.right - box.left < 4 && box.bottom - box.top < 4
            Object.assign(marq.style, { left: box.left - r0.left + 'px', top: box.top - r0.top + 'px', width: box.right - box.left + 'px', height: box.bottom - box.top + 'px' })
          }
          const up = (ev) => {
            ev.preventDefault()
            window.removeEventListener('pointermove', move, true)
            window.removeEventListener('pointerup', up, true)
            // The click that follows this release is still swallowed.
            setTimeout(() => this.pinDisarm())
            const t = this.editor.pinAt(box)
            if (!t) { marq.remove(); this.toast('框选范围里没有可标注的内容。', 'info'); return }
            this.pinPop({ ...t, rect: marq.hidden ? t.rect : box, shot: marq.hidden ? t.shot : box, boxed: !marq.hidden })
            if (marq.hidden) marq.remove(); else this.pinMarq = marq
          }
          window.addEventListener('pointermove', move, true)
          window.addEventListener('pointerup', up, true)
        }
        const key = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.pinDisarm() } }
        const on = [['pointerdown', down], ['mousedown', block], ['click', block], ['dblclick', block], ['contextmenu', block], ['keydown', key]]
        for (const [k, f] of on) window.addEventListener(k, f, true)
        this.root.classList.add('q-pinning')
        this.pinBtn.classList.add('arm')
        this.pinArm = () => { for (const [k, f] of on) window.removeEventListener(k, f, true); tip.remove() }
      }

      pinDisarm () {
        if (!this.pinArm) return
        this.pinArm()
        this.pinArm = null
        this.root.classList.remove('q-pinning')
        this.pinBtn.classList.remove('arm')
      }

      // pinHere marks what is selected now, for the context menus.
      pinHere () {
        this.pinDisarm()
        const t = this.editor?.pinTarget?.()
        if (!t) { this.toast('先选中要标注的内容：单元格、文字、形状或幻灯片。', 'info'); return }
        this.pinPop(t)
      }

      // pinShot pictures the spot: the box drawn, else what is selected,
      // with a little margin, inside the editing area.
      pinShot (t) {
        const area = (r) => r && r.right - r.left > 2 && r.bottom - r.top > 2
        let r = t.shot ?? this.editor?.snapRect?.()
        if (!area(r)) r = area(t.rect) ? t.rect : null
        if (!r || !this.editor?.shoot) return Promise.resolve(null)
        const M = this.main.getBoundingClientRect()
        const pad = t.boxed ? 0 : 8
        const c = { left: Math.max(M.left, r.left - pad), top: Math.max(M.top, r.top - pad), right: Math.min(M.right, r.right + pad), bottom: Math.min(M.bottom, r.bottom + pad) }
        if (!area(c)) return Promise.resolve(null)
        return Promise.resolve().then(() => this.editor.shoot(c)).then(snapOut).catch((err) => { console.warn('quire: snapshot failed', err); return null })
      }

      pinPop (t) {
        this.pinClose()
        // Taken now, before the popover or the next edit covers the spot.
        t.picture = this.pinShot(t)
        const box = el('textarea', { class: 'q-pin-in', rows: '3', placeholder: '这里有什么问题，或想让 Agent 怎么改？', 'aria-label': '标注内容', spellcheck: 'false' })
        const ok = el('button', { type: 'button', class: 'q-btn q-pin-ok', disabled: true }, '添加标注')
        const cancel = el('button', { type: 'button', class: 'q-btn' }, '取消')
        const pop = el('div', { class: 'q-pinpop', role: 'dialog', 'aria-label': '标注' },
          el('div', { class: 'q-pin-h' }, el('b', {}, pinWhere(t.where)), t.quote !== undefined ? el('span', { class: 'q-pin-q', title: oneLine(t.quote) }, clipPin(t.quote, 80) || '空') : null),
          box,
          el('div', { class: 'q-pin-f' }, el('span', { class: 'q-pin-k' }, '⏎ 添加　⇧⏎ 换行'), el('span', { class: 'q-space' }), cancel, ok))
        this.root.append(pop)
        this.pinPopEl = pop
        t.picture.then((p) => {
          if (!p || this.pinPopEl !== pop) return
          pop.querySelector('.q-pin-h').prepend(el('img', { class: 'q-pin-shot', src: p.thumb, alt: '所选区域截图', title: '所选区域截图会随标注发给 Agent' }))
        })
        // The spot stays lit while the note is written.
        if (t.range && window.CSS?.highlights && typeof Highlight !== 'undefined') window.CSS.highlights.set('q-pin-on', new Highlight(t.range))
        const R = this.root.getBoundingClientRect()
        const W = pop.offsetWidth; const H = pop.offsetHeight
        const a = t.rect ?? { left: R.left + 40, right: R.left + 40, top: R.top + 80, bottom: R.top + 80 }
        let top = a.bottom - R.top + 6
        if (top + H > R.height - 8) top = a.top - R.top - H - 6
        pop.style.top = Math.max(8, Math.min(top, R.height - H - 8)) + 'px'
        pop.style.left = Math.max(8, Math.min(a.left - R.left, R.width - W - 8)) + 'px'
        box.addEventListener('input', () => { ok.disabled = !box.value.trim() })
        box.addEventListener('keydown', (e) => {
          if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.pinClose(true) } else if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); ok.click() }
        })
        ok.addEventListener('click', () => { if (!box.value.trim()) return; this.pinClose(true); this.pinSend(t, box.value) })
        cancel.addEventListener('click', () => this.pinClose(true))
        this.pinAway = (e) => { if (!pop.contains(e.target)) this.pinClose() }
        setTimeout(() => document.addEventListener('pointerdown', this.pinAway, true))
        box.focus()
      }

      pinClose (refocus) {
        this.pinMarq?.remove()
        this.pinMarq = null
        if (!this.pinPopEl) return
        this.pinPopEl.remove()
        this.pinPopEl = null
        document.removeEventListener('pointerdown', this.pinAway, true)
        window.CSS?.highlights?.delete?.('q-pin-on')
        if (refocus) this.editor?.focus?.()
      }

      async pinSend (t, note) {
        const file = this.doc.file
        const pin = { path: file.path, where: t.where, quote: t.quote !== undefined ? oneLine(t.quote) : '', extra: t.extra ? oneLine(t.extra) : '', note: note.trim() }
        const picture = await t.picture
        if (picture) { pin.image = picture.image; pin.thumb = picture.thumb }
        try {
          const r = await pinStore.call(file.sessionId, 'add', { pin })
          this.toast(`已添加到聊天框上方，共 ${r.items.length} 条标注，随下一条消息发给 Agent`)
        } catch (err) {
          this.toast(`标注没有保存：${err.message}`, 'warn')
        }
      }

      // syncPins marks the pins waiting for this file.
      syncPins () {
        if (!this.doc?.file) return
        const { sessionId, path } = this.doc.file
        this.pins = pinStore.get(sessionId).items.filter((p) => p.path === path)
        const n = this.pins.length
        this.pinN.hidden = !n
        this.pinN.textContent = String(n)
        this.pinMore.hidden = !n
        this.pinBtn.classList.toggle('on', n > 0)
        this.editor?.showPins?.(this.pins)
      }

      pinMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const { sessionId, path } = this.doc.file
        const drop = (op, extra) => pinStore.call(sessionId, op, extra).catch((err) => this.toast(`没能移除：${err.message}`, 'warn'))
        menu(r.left, r.bottom + 2, [
          ...this.pins.map((p) => ({ label: `${p.where}　${clipPin(p.note, 24)}`, fn: () => this.editor?.revealPin?.(p) })),
          '-',
          { label: '移除这个文件的标注', fn: () => drop('remove', { path }) },
          { label: `清空全部标注（${pinStore.get(sessionId).items.length} 条）`, fn: () => drop('clear') }
        ])
      }

      dispose () { this.pinDisarm(); this.pinClose(); document.removeEventListener('keydown', this.onKey); this.root.remove() }
    }

    // ------------------------------------------------------------ editors
    // ---------------------------------------------------------- rich text

    // Helpers shared by the document and slide editors: text nodes in a
    // range, each in a run element of its own, so formatting can be set
    // on exactly the selected text.
    const ZW = '​'
    const isRun = (n, root) => n && n !== root && n.nodeType === 1 && n.tagName === 'SPAN' && !n.classList.contains('bu') && !n.hasAttribute('data-pb')

    function currentRange (root) {
      const sel = window.getSelection()
      if (!sel || !sel.rangeCount) return undefined
      const r = sel.getRangeAt(0)
      return root.contains(r.commonAncestorContainer) ? r : undefined
    }

    function setCaret (node, offset) {
      const r = document.createRange()
      r.setStart(node, offset)
      r.collapse(true)
      const sel = window.getSelection()
      sel.removeAllRanges()
      sel.addRange(r)
    }

    // textsIn splits the range's boundary text nodes and returns the text
    // nodes inside it that accept() allows.
    function textsIn (range, root, accept) {
      let { startContainer: sc, startOffset: so, endContainer: ec, endOffset: eo } = range
      if (ec.nodeType === 3 && eo < ec.length && eo > 0) ec.splitText(eo)
      if (sc.nodeType === 3 && so > 0 && so < sc.length) { const tail = sc.splitText(so); if (ec === sc) ec = tail; sc = tail; so = 0 }
      const r = document.createRange()
      if (sc.nodeType === 3 && so >= sc.length) r.setStartAfter(sc); else r.setStart(sc, so)
      if (ec.nodeType === 3) r.setEnd(ec, ec === sc ? ec.length : Math.min(eo, ec.length)); else r.setEnd(ec, eo)
      const out = []
      const walk = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      for (let t = walk.nextNode(); t; t = walk.nextNode()) {
        if (!t.data.replace(/​/g, '') || !r.intersectsNode(t)) continue
        if (accept && !accept(t)) continue
        out.push(t)
      }
      return out
    }

    // isolate puts the text node in a run element of its own and returns it.
    function isolate (t, root, cls) {
      const p = t.parentNode
      if (!isRun(p, root)) {
        const s = el('span', { class: cls })
        p.insertBefore(s, t)
        s.append(t)
        return s
      }
      if (p.childNodes.length === 1) return p
      const before = p.cloneNode(false)
      while (p.firstChild !== t) before.append(p.firstChild)
      if (before.childNodes.length) p.before(before)
      const after = p.cloneNode(false)
      while (t.nextSibling) after.append(t.nextSibling)
      if (after.childNodes.length) p.after(after)
      return p
    }

    function selectTexts (texts) {
      if (!texts.length) return
      const r = document.createRange()
      r.setStart(texts[0], 0)
      const last = texts[texts.length - 1]
      r.setEnd(last, last.length)
      const sel = window.getSelection()
      sel.removeAllRanges()
      sel.addRange(r)
    }

    // caretRun inserts an empty run at the caret, carrying formatting for
    // what is typed next.
    function caretRun (range, root, make) {
      const span = make()
      span.textContent = ZW
      const c = range.startContainer
      const o = range.startOffset
      if (c.nodeType === 3) {
        const p = c.parentNode
        const tail = o < c.length ? c.splitText(o) : null
        if (isRun(p, root)) {
          for (const k of Object.keys(p.dataset)) if (!(k in span.dataset)) span.dataset[k] = p.dataset[k]
          if (tail) { const after = p.cloneNode(false); for (let n = tail; n;) { const nx = n.nextSibling; after.append(n); n = nx } p.after(after) }
          p.after(span)
        } else p.insertBefore(span, tail ?? c.nextSibling)
      } else {
        const ref = c.childNodes[o] ?? null
        if (ref?.nodeName === 'BR') ref.remove()
        c.insertBefore(span, ref?.isConnected ? ref : null)
      }
      setCaret(span.firstChild, 1)
      return span
    }

    const rgbHex = (s) => { const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(s || ''); return m ? m.slice(1, 4).map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase() : undefined }

    // ---------------------------------------------------------- spreadsheet

    const HW = 46 // row header width
    const HH = 22 // column header height
    const FONT = 'Calibri, "PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", sans-serif'
    const FORMATS = [
      { value: 'General', label: '常规' },
      { value: '0', label: '整数' },
      { value: '0.00', label: '两位小数' },
      { value: '#,##0', label: '千分位' },
      { value: '#,##0.00', label: '千分位·两位小数' },
      { value: '"¥"#,##0.00', label: '货币 ¥' },
      { value: '0%', label: '百分比' },
      { value: '0.00%', label: '百分比·两位' },
      { value: 'yyyy-mm-dd', label: '日期' },
      { value: 'yyyy-mm-dd hh:mm', label: '日期时间' },
      { value: '@', label: '文本' }
    ]
    const FN_HELP = {
      SUM: 'SUM(数值1, [数值2], …) 求和', AVERAGE: 'AVERAGE(数值1, …) 平均值', MIN: 'MIN(数值1, …) 最小值', MAX: 'MAX(数值1, …) 最大值',
      COUNT: 'COUNT(值1, …) 数字个数', COUNTA: 'COUNTA(值1, …) 非空个数', COUNTBLANK: 'COUNTBLANK(区域) 空白个数', COUNTIF: 'COUNTIF(区域, 条件) 满足条件的个数',
      SUMIF: 'SUMIF(区域, 条件, [求和区域]) 条件求和', AVERAGEIF: 'AVERAGEIF(区域, 条件, [平均区域]) 条件平均', PRODUCT: 'PRODUCT(数值1, …) 乘积', MEDIAN: 'MEDIAN(数值1, …) 中位数',
      ROUND: 'ROUND(数值, 位数) 四舍五入', ROUNDUP: 'ROUNDUP(数值, 位数) 向上舍入', ROUNDDOWN: 'ROUNDDOWN(数值, 位数) 向下舍入', INT: 'INT(数值) 向下取整',
      ABS: 'ABS(数值) 绝对值', SQRT: 'SQRT(数值) 平方根', POWER: 'POWER(底数, 指数) 乘幂', MOD: 'MOD(被除数, 除数) 余数',
      IF: 'IF(条件, 真值, [假值]) 条件判断', IFERROR: 'IFERROR(值, 出错时的值) 出错时替换', AND: 'AND(条件1, …) 全部为真', OR: 'OR(条件1, …) 任一为真', NOT: 'NOT(条件) 取反',
      CONCAT: 'CONCAT(文本1, …) 连接文本', CONCATENATE: 'CONCATENATE(文本1, …) 连接文本', LEN: 'LEN(文本) 字符数', LEFT: 'LEFT(文本, [个数]) 左边的字符',
      RIGHT: 'RIGHT(文本, [个数]) 右边的字符', MID: 'MID(文本, 起点, 个数) 中间的字符', UPPER: 'UPPER(文本) 转大写', LOWER: 'LOWER(文本) 转小写', TRIM: 'TRIM(文本) 去掉多余空格'
    }
    Object.assign(FN_HELP, {
      VLOOKUP: 'VLOOKUP(查找值, 表格区域, 列序号, [近似匹配]) 按首列纵向查找', HLOOKUP: 'HLOOKUP(查找值, 表格区域, 行序号, [近似匹配]) 按首行横向查找',
      XLOOKUP: 'XLOOKUP(查找值, 查找区域, 返回区域, [未找到时], [匹配方式]) 查找并返回', INDEX: 'INDEX(区域, 行号, [列号]) 取指定位置的值', MATCH: 'MATCH(查找值, 区域, [匹配方式]) 查找位置',
      CHOOSE: 'CHOOSE(序号, 值1, 值2, …) 按序号选值', ROW: 'ROW([引用]) 行号', COLUMN: 'COLUMN([引用]) 列号', ROWS: 'ROWS(区域) 行数', COLUMNS: 'COLUMNS(区域) 列数',
      SUMIFS: 'SUMIFS(求和区域, 条件区域1, 条件1, …) 多条件求和', COUNTIFS: 'COUNTIFS(条件区域1, 条件1, …) 多条件计数', AVERAGEIFS: 'AVERAGEIFS(平均区域, 条件区域1, 条件1, …) 多条件平均',
      MAXIFS: 'MAXIFS(最大值区域, 条件区域1, 条件1, …) 多条件最大值', MINIFS: 'MINIFS(最小值区域, 条件区域1, 条件1, …) 多条件最小值', SUMPRODUCT: 'SUMPRODUCT(数组1, [数组2], …) 对应相乘再求和',
      TODAY: 'TODAY() 今天的日期', NOW: 'NOW() 当前日期和时间', DATE: 'DATE(年, 月, 日) 组成日期', YEAR: 'YEAR(日期) 年份', MONTH: 'MONTH(日期) 月份', DAY: 'DAY(日期) 日',
      WEEKDAY: 'WEEKDAY(日期, [类型]) 星期几', EDATE: 'EDATE(开始日期, 月数) 几个月前后的日期', EOMONTH: 'EOMONTH(开始日期, 月数) 月末日期', DATEDIF: 'DATEDIF(开始日期, 结束日期, 单位) 日期间隔（"Y"/"M"/"D"）',
      TEXT: 'TEXT(值, 格式) 按格式转为文本', VALUE: 'VALUE(文本) 文本转数字', SUBSTITUTE: 'SUBSTITUTE(文本, 旧文本, 新文本, [第几个]) 替换文本', REPLACE: 'REPLACE(文本, 起点, 个数, 新文本) 按位置替换',
      FIND: 'FIND(要找的文本, 文本, [起点]) 查找位置（区分大小写）', SEARCH: 'SEARCH(要找的文本, 文本, [起点]) 查找位置（可用通配符）', REPT: 'REPT(文本, 次数) 重复文本', EXACT: 'EXACT(文本1, 文本2) 是否完全相同',
      PROPER: 'PROPER(文本) 单词首字母大写', TEXTJOIN: 'TEXTJOIN(分隔符, 忽略空白, 文本1, …) 用分隔符连接',
      IFS: 'IFS(条件1, 值1, [条件2, 值2], …) 多条件判断', SWITCH: 'SWITCH(表达式, 值1, 结果1, …, [默认]) 按值匹配', IFNA: 'IFNA(值, 为 #N/A 时的值) #N/A 时替换', XOR: 'XOR(条件1, …) 奇数个为真',
      ISBLANK: 'ISBLANK(值) 是否为空', ISNUMBER: 'ISNUMBER(值) 是否为数字', ISTEXT: 'ISTEXT(值) 是否为文本', ISERROR: 'ISERROR(值) 是否为错误值', ISNA: 'ISNA(值) 是否为 #N/A',
      ISLOGICAL: 'ISLOGICAL(值) 是否为逻辑值', ISEVEN: 'ISEVEN(数值) 是否为偶数', ISODD: 'ISODD(数值) 是否为奇数', NA: 'NA() 返回 #N/A',
      RANK: 'RANK(数值, 区域, [升序]) 排名', LARGE: 'LARGE(区域, k) 第 k 大的值', SMALL: 'SMALL(区域, k) 第 k 小的值', STDEV: 'STDEV(数值1, …) 样本标准差', STDEVP: 'STDEVP(数值1, …) 总体标准差',
      VAR: 'VAR(数值1, …) 样本方差', VARP: 'VARP(数值1, …) 总体方差', CEILING: 'CEILING(数值, 基数) 向上舍入到倍数', FLOOR: 'FLOOR(数值, 基数) 向下舍入到倍数', PI: 'PI() 圆周率',
      EXP: 'EXP(数值) e 的乘幂', LN: 'LN(数值) 自然对数', LOG: 'LOG(数值, [底数]) 对数', LOG10: 'LOG10(数值) 以 10 为底的对数', SIGN: 'SIGN(数值) 正负号', TRUNC: 'TRUNC(数值, [位数]) 截去小数',
      RAND: 'RAND() 0 到 1 的随机数', RANDBETWEEN: 'RANDBETWEEN(下限, 上限) 范围内的随机整数'
    })
    const FN_CATS = [
      ['常用', ['SUM', 'AVERAGE', 'IF', 'COUNT', 'MAX', 'MIN', 'VLOOKUP', 'XLOOKUP', 'SUMIF', 'COUNTIF', 'ROUND', 'IFERROR']],
      ['数学', ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'PRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'TRUNC', 'CEILING', 'FLOOR', 'ABS', 'SQRT', 'POWER', 'MOD', 'SIGN', 'PI', 'EXP', 'LN', 'LOG', 'LOG10', 'RAND', 'RANDBETWEEN']],
      ['统计', ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MAX', 'MAXIFS', 'MIN', 'MINIFS', 'MEDIAN', 'RANK', 'LARGE', 'SMALL', 'STDEV', 'STDEVP', 'VAR', 'VARP']],
      ['逻辑', ['IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'AND', 'OR', 'NOT', 'XOR']],
      ['文本', ['CONCAT', 'CONCATENATE', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'TRIM', 'UPPER', 'LOWER', 'PROPER', 'REPT', 'EXACT', 'TEXT', 'VALUE']],
      ['日期', ['TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'EDATE', 'EOMONTH', 'DATEDIF']],
      ['查找', ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'CHOOSE', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS']],
      ['信息', ['ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISNA', 'ISLOGICAL', 'ISEVEN', 'ISODD', 'NA']]
    ]
    const FN_NAMES = Object.keys(FN_HELP)
    const REF_COLORS = ['#1a73e8', '#d93025', '#8e24aa', '#188038', '#e37400', '#c2185b', '#00838f']
    const CF_STYLES = [
      ['red', '浅红填充色深红色文本', { fill: 'FFC7CE', color: '9C0006' }], ['yellow', '黄填充色深黄色文本', { fill: 'FFEB9C', color: '9C5700' }],
      ['green', '绿填充色深绿色文本', { fill: 'C6EFCE', color: '006100' }], ['fill', '浅红色填充', { fill: 'FFC7CE' }], ['text', '红色文本', { color: 'C00000' }], ['bold', '加粗', { bold: true }]
    ]
    const CF_OPS = { greaterThan: '大于', lessThan: '小于', between: '介于', equal: '等于', notEqual: '不等于', greaterThanOrEqual: '大于或等于', lessThanOrEqual: '小于或等于', notBetween: '未介于' }
    const BAR_COLORS = [['638EC6', '蓝色'], ['63C384', '绿色'], ['FF555A', '红色'], ['FFB628', '橙色'], ['008AEF', '浅蓝色'], ['D6007B', '紫色']]
    const SCALES = [[['F8696B', 'FFEB84', '63BE7B'], '红 - 黄 - 绿'], [['63BE7B', 'FFEB84', 'F8696B'], '绿 - 黄 - 红'], [['F8696B', 'FFFFFF', '5A8AC6'], '红 - 白 - 蓝'], [['5A8AC6', 'FFFFFF', 'F8696B'], '蓝 - 白 - 红'], [['FFFFFF', '63BE7B'], '白 - 绿'], [['FFFFFF', 'F8696B'], '白 - 红'], [['FCFCFF', '5A8AC6'], '白 - 蓝']]
    const CELL_STYLES = [
      ['常规', { bold: false, italic: false, color: '000000', fill: null, format: 'General' }],
      ['好', { fill: 'C6EFCE', color: '006100' }], ['差', { fill: 'FFC7CE', color: '9C0006' }], ['适中', { fill: 'FFEB9C', color: '9C5700' }],
      '-',
      ['标题', { size: 18, bold: true, color: '44546A' }], ['标题 1', { size: 15, bold: true, color: '44546A' }], ['标题 2', { size: 13, bold: true, color: '44546A' }], ['汇总', { bold: true, border: true }],
      '-',
      ['计算', { fill: 'F2F2F2', color: 'FA7D00', bold: true, border: true }], ['输入', { fill: 'FFCC99', color: '3F3F76', border: true }], ['注释', { fill: 'FFFFCC', border: true }], ['警告文本', { color: 'FF0000' }],
      '-',
      ['着色 1', { fill: '4472C4', color: 'FFFFFF' }], ['着色 2', { fill: 'ED7D31', color: 'FFFFFF' }], ['着色 3', { fill: 'A5A5A5', color: 'FFFFFF' }], ['着色 4', { fill: 'FFC000', color: '000000' }], ['着色 5', { fill: '5B9BD5', color: 'FFFFFF' }], ['着色 6', { fill: '70AD47', color: 'FFFFFF' }],
      '-',
      ['货币', { format: '¥#,##0.00' }], ['百分比', { format: '0%' }], ['千位分隔', { format: '#,##0.00' }]
    ]
    // rangeList reads a space-separated list of A1 ranges.
    const rangeList = (s) => String(s ?? '').split(/\s+/).map((x) => {
      const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(x)
      if (!m) return undefined
      const col = (t) => [...t.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
      const r1 = +m[2]; const c1 = col(m[1]); const r2 = m[4] ? +m[4] : r1; const c2 = m[3] ? col(m[3]) : c1
      return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }
    }).filter(Boolean)
    const dateSerial = (t) => {
      const m = /^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$/.exec(String(t).trim())
      if (!m) return NaN
      let n = Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569
      if (n < 61) n -= 1
      return n
    }
    const BORDERS = [
      { value: '', label: '边框…' }, { value: 'all', label: '所有框线' }, { value: 'outside', label: '外侧框线' }, { value: 'thick', label: '粗外侧框线' },
      { value: 'bottom', label: '下框线' }, { value: 'top', label: '上框线' }, { value: 'left', label: '左框线' }, { value: 'right', label: '右框线' }, { value: 'none', label: '无框线' }
    ]
    const reEsc = (q) => q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

    // ------------------------------------------------------------ charts

    const CHART_TYPES = [['col', '柱形图'], ['bar', '条形图'], ['line', '折线图'], ['area', '面积图'], ['pie', '饼图'], ['doughnut', '圆环图'], ['scatter', '散点图']]
    const CHART_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300']

    // niceTicks spreads about n round values over lo..hi.
    function niceTicks (lo, hi, n = 5) {
      if (!(hi > lo)) { hi = lo + 1 }
      const raw = (hi - lo) / n
      const mag = Math.pow(10, Math.floor(Math.log10(raw)))
      const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((x) => x >= raw * 0.999)
      const out = []
      for (let v = Math.floor(lo / step + 1e-9) * step; v <= hi + step * 0.999 && out.length < 40; v += step) out.push(+v.toPrecision(12))
      if (out.length < 2) out.push(+(out[0] + step).toPrecision(12))
      return out
    }
    const tickText = (v) => (Math.abs(v) >= 1e8 ? +(v / 1e8).toPrecision(4) + '亿' : Math.abs(v) >= 1e4 ? +(v / 1e4).toPrecision(4) + '万' : String(+v.toPrecision(7)))
    const textW = (s) => [...String(s)].reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 10 : 5.6), 0)
    const clipText = (s, w) => { s = String(s); if (textW(s) <= w) return s; let o = ''; for (const ch of s) { if (textW(o + ch + '…') > w) break; o += ch } return o + '…' }
    const hexOr = (c, i) => (/^[0-9A-F]{6}$/i.test(c ?? '') ? c : CHART_COLORS[i % CHART_COLORS.length])

    // chartSvg draws a chart from the model the host sends (series with
    // categories, values and colors) at w×h pixels.
    function chartSvg (ch, W, H) {
      const out = []
      const f1 = (v) => Math.round(v * 10) / 10
      const txt = (x, y, s, o = {}) => `<text x="${f1(x)}" y="${f1(y)}" font-size="${o.size ?? 10}" fill="${o.fill ?? '#595959'}" text-anchor="${o.anchor ?? 'middle'}"${o.bold ? ' font-weight="600"' : ''}${o.rot ? ` transform="rotate(${o.rot} ${f1(x)} ${f1(y)})"` : ''}>${esc(s)}</text>`
      const line = (a, b, c, d, col, w = 1) => `<line x1="${f1(a)}" y1="${f1(b)}" x2="${f1(c)}" y2="${f1(d)}" stroke="${col}" stroke-width="${w}"/>`
      const box = (x, y, w, h, col) => `<rect x="${f1(x)}" y="${f1(y)}" width="${f1(Math.max(0, w))}" height="${f1(Math.max(0, h))}" fill="#${col}"/>`
      const svg = () => `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family='${FONT}'>${out.join('')}</svg>`
      const num = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : null)
      const series = (ch.series ?? []).filter((x) => x.vals?.length).map((x, i) => ({ ...x, color: hexOr(x.color, i) }))
      const pie = ch.type === 'pie' || ch.type === 'doughnut'
      let x0 = 8; let y0 = 8; let x1 = W - 8; let y1 = H - 8
      if (ch.title) { out.push(txt(W / 2, 20, clipText(ch.title, W - 16), { size: 13, fill: '#404040', bold: 1 })); y0 = 30 }
      if (!series.length) { out.push(txt(W / 2, H / 2, '没有可显示的数据')); return svg() }
      const first = series[0]
      const legend = pie ? first.vals.map((_, i) => [String(first.cats[i] ?? i + 1), CHART_COLORS[i % CHART_COLORS.length]]) : series.map((x) => [x.name, x.color])
      // The legend sits where the file says: a column on the right or
      // left (r, l, tr), or rows above or below (t, b).
      const lp = ch.legend === null ? null : ch.legend ?? (pie ? 'r' : 'b')
      if (lp && ['r', 'l', 'tr'].includes(lp)) {
        const lw = Math.min(W * 0.4, 16 + Math.max(...legend.map(([n]) => textW(n))))
        const items = legend.slice(0, Math.max(1, Math.floor((y1 - y0) / 16)))
        const top = lp === 'tr' ? y0 : Math.max(y0, (y0 + y1) / 2 - items.length * 8)
        const lx = lp === 'l' ? x0 : x1 - lw
        items.forEach(([n, c], i) => { const y = top + 2 + i * 16; out.push(box(lx, y + 2, 9, 9, c), txt(lx + 13, y + 10, clipText(n, lw - 14), { anchor: 'start' })) })
        if (lp === 'l') x0 += lw + 8; else x1 -= lw + 8
      } else if (lp) {
        const rows = []; let row = []; let rw = 0
        for (const it of legend) { const iw = 24 + textW(it[0]); if (rw + iw > x1 - x0 && row.length) { rows.push([row, rw]); row = []; rw = 0 } row.push(it); rw += iw }
        if (row.length) rows.push([row, rw])
        const shown = rows.slice(0, 2)
        shown.forEach(([items, w], ri) => {
          let x = Math.max(x0, (W - w) / 2)
          const y = lp === 't' ? y0 + ri * 16 : y1 - (shown.length - ri) * 16 + 4
          for (const [n, c] of items) { out.push(box(x, y + 2, 9, 9, c), txt(x + 13, y + 10, n, { anchor: 'start' })); x += 24 + textW(n) }
        })
        if (lp === 't') y0 += shown.length * 16 + 4; else y1 -= shown.length * 16 + 2
      }
      // Axis titles: the one along the left turned on its side.
      if (!pie) {
        const horiz = ch.type === 'bar'
        const left = horiz ? ch.xTitle : ch.yTitle
        const bottom = horiz ? ch.yTitle : ch.xTitle
        if (bottom) { out.push(txt((x0 + x1) / 2, y1 - 2, clipText(bottom, x1 - x0), { size: 11, fill: '#404040' })); y1 -= 16 }
        if (left) { out.push(txt(x0 + 9, (y0 + y1) / 2, clipText(left, y1 - y0), { size: 11, fill: '#404040', rot: -90 })); x0 += 16 }
      }
      // Data labels: the value, the pie share or the category of a point.
      const fmtVal = (v) => (Math.abs(v) >= 1000 ? Math.round(v).toLocaleString('en-US') : String(+v.toPrecision(4)))
      const label = (v, cat, share) => (ch.labels === 'category' ? String(cat ?? '') : ch.labels === 'percent' && share !== undefined ? Math.round(share * 100) + '%' : fmtVal(v))
      // The value axis: its ticks, inside the bounds the file fixes.
      const valTicks = (lo, hi) => {
        const a = typeof ch.yMin === 'number' ? ch.yMin : lo; const b = typeof ch.yMax === 'number' ? ch.yMax : hi
        let t = niceTicks(a, b > a ? b : a + 1)
        if (typeof ch.yMin === 'number') { t = t.filter((v) => v > a); t.unshift(a) }
        if (typeof ch.yMax === 'number' && ch.yMax > t[0]) { t = t.filter((v) => v < ch.yMax); t.push(ch.yMax) }
        return t
      }
      const grid = ch.gridlines !== false
      if (pie) {
        const vals = first.vals.map((v) => Math.max(0, num(v) ?? 0))
        const total = vals.reduce((a, b) => a + b, 0)
        const cx = (x0 + x1) / 2; const cy = (y0 + y1) / 2; const r = Math.min(x1 - x0, y1 - y0) / 2 - 2
        if (r < 8 || !total) return svg()
        let a = -Math.PI / 2
        vals.forEach((v, i) => {
          if (!v) return
          const c = CHART_COLORS[i % CHART_COLORS.length]
          const b = a + v / total * Math.PI * 2
          if (v === total) out.push(`<circle cx="${f1(cx)}" cy="${f1(cy)}" r="${f1(r)}" fill="#${c}"/>`)
          else out.push(`<path d="M${f1(cx)} ${f1(cy)}L${f1(cx + r * Math.cos(a))} ${f1(cy + r * Math.sin(a))}A${f1(r)} ${f1(r)} 0 ${b - a > Math.PI ? 1 : 0} 1 ${f1(cx + r * Math.cos(b))} ${f1(cy + r * Math.sin(b))}Z" fill="#${c}" stroke="#fff" stroke-width="1"/>`)
          if (ch.labels && v / total > 0.06 && r > 40) { const m = (a + b) / 2; const k = ch.type === 'doughnut' ? 0.75 : 0.62; out.push(txt(cx + r * k * Math.cos(m), cy + r * k * Math.sin(m) + 3.5, clipText(label(v, first.cats[i], v / total), r * 0.8), { fill: '#fff', bold: 1 })) }
          a = b
        })
        if (ch.type === 'doughnut') out.push(`<circle cx="${f1(cx)}" cy="${f1(cy)}" r="${f1(r * 0.5)}" fill="#fff"/>`)
        return svg()
      }
      if (ch.type === 'scatter') {
        const pts = series.map((x) => x.vals.map((v, i) => { const xv = num(+x.cats[i]) ?? (x.cats[i] === undefined || x.cats[i] === '' ? i + 1 : null); const yv = num(v); return xv === null || yv === null ? null : [xv, yv] }).filter(Boolean))
        const all = pts.flat()
        if (!all.length) return svg()
        const xt = niceTicks(Math.min(0, ...all.map((p) => p[0])), Math.max(...all.map((p) => p[0])))
        const yt = valTicks(Math.min(0, ...all.map((p) => p[1])), Math.max(...all.map((p) => p[1])))
        const px0 = x0 + Math.max(...yt.map((v) => textW(tickText(v)))) + 6; const px1 = x1 - 6; const py0 = y0 + 4; const py1 = y1 - 16
        if (px1 - px0 < 20 || py1 - py0 < 20) return svg()
        const X = (v) => px0 + (v - xt[0]) / (xt[xt.length - 1] - xt[0]) * (px1 - px0)
        const Y = (v) => py1 - (v - yt[0]) / (yt[yt.length - 1] - yt[0]) * (py1 - py0)
        for (const v of yt) out.push(grid ? line(px0, Y(v), px1, Y(v), '#d9d9d9') : '', txt(px0 - 4, Y(v) + 3.5, tickText(v), { anchor: 'end' }))
        for (const v of xt) out.push(grid ? line(X(v), py0, X(v), py1, '#ececec') : '', txt(X(v), py1 + 12, tickText(v)))
        pts.forEach((ps, si) => { for (const [a, b] of ps) if (b >= yt[0] && b <= yt[yt.length - 1]) out.push(`<circle cx="${f1(X(a))}" cy="${f1(Y(b))}" r="3.5" fill="#${series[si].color}"/>`, ch.labels ? txt(X(a), Y(b) - 6, label(b, a)) : '') })
        return svg()
      }
      // Category charts: columns, bars, lines and areas.
      const n = Math.max(...series.map((x) => x.vals.length))
      const labels = Array.from({ length: n }, (_, i) => String(series.find((x) => x.cats?.length)?.cats[i] ?? i + 1))
      const pct = ch.stacked === 'percent'
      const sums = []
      let lo = 0; let hi = 0
      if (ch.stacked) {
        for (let i = 0; i < n; i++) {
          let p = 0; let q = 0
          for (const x of series) { const v = num(x.vals[i]); if (v === null) continue; if (v >= 0) p += v; else q += v }
          sums.push([p, q])
          if (!pct) { hi = Math.max(hi, p); lo = Math.min(lo, q) }
        }
        if (pct) { hi = 1; lo = sums.some((x) => x[1] < 0) ? -1 : 0 }
      } else for (const x of series) for (const v of x.vals) { const y = num(v); if (y !== null) { hi = Math.max(hi, y); lo = Math.min(lo, y) } }
      const ticks = valTicks(lo, hi)
      const tlo = ticks[0]; const thi = ticks[ticks.length - 1]
      const fmt = (v) => (pct ? Math.round(v * 100) + '%' : tickText(v))
      const horiz = ch.type === 'bar'
      let px0; let px1; let py0; let py1
      if (horiz) { px0 = x0 + Math.min((x1 - x0) * 0.35, Math.max(...labels.map(textW)) + 6); px1 = x1 - 8; py0 = y0 + 4; py1 = y1 - 14 } else { px0 = x0 + Math.max(...ticks.map((v) => textW(fmt(v)))) + 6; px1 = x1 - 4; py0 = y0 + 4; py1 = y1 - 16 }
      if (px1 - px0 < 20 || py1 - py0 < 20) return svg()
      const vpos = (v) => { v = Math.max(tlo, Math.min(thi, v)); return horiz ? px0 + (v - tlo) / (thi - tlo) * (px1 - px0) : py1 - (v - tlo) / (thi - tlo) * (py1 - py0) }
      for (const v of ticks) {
        const p = vpos(v)
        out.push(horiz ? (grid ? line(p, py0, p, py1, '#d9d9d9') : '') + txt(p, py1 + 11, fmt(v)) : (grid ? line(px0, p, px1, p, '#d9d9d9') : '') + txt(px0 - 4, p + 3.5, fmt(v), { anchor: 'end' }))
      }
      const zero = vpos(Math.max(tlo, Math.min(thi, 0)))
      out.push(horiz ? line(zero, py0, zero, py1, '#bfbfbf') : line(px0, zero, px1, zero, '#bfbfbf'))
      const len = horiz ? py1 - py0 : px1 - px0
      const band = len / n
      const centre = (i) => (horiz ? py0 : px0) + band * (i + 0.5)
      const every = Math.max(1, Math.ceil(n * (horiz ? 13 : Math.max(...labels.map(textW)) + 6) / len))
      labels.forEach((l, i) => {
        if (i % every) return
        out.push(horiz ? txt(px0 - 4, centre(i) + 3.5, clipText(l, px0 - x0 - 6), { anchor: 'end' }) : txt(centre(i), py1 + 12, clipText(l, band * every - 2)))
      })
      const share = (v, i) => { if (!pct) return v; const t = sums[i][0] - sums[i][1]; return t ? v / t : 0 }
      if (ch.type === 'col' || ch.type === 'bar') {
        const k = ch.stacked ? 1 : series.length
        const b = band / (k + 1.5)
        const acc = Array.from({ length: n }, () => [0, 0])
        series.forEach((x, si) => x.vals.forEach((v0, i) => {
          let v = num(v0)
          if (v === null) return
          v = share(v, i)
          let a = 0; let e = v
          if (ch.stacked) { const side = v >= 0 ? 0 : 1; a = acc[i][side]; e = a + v; acc[i][side] = e }
          const s0 = (horiz ? py0 : px0) + band * i + b * 0.75 + (ch.stacked ? 0 : si) * b
          const p1 = vpos(a); const p2 = vpos(e)
          out.push(horiz ? box(Math.min(p1, p2), s0, Math.abs(p2 - p1), b, x.color) : box(s0, Math.min(p1, p2), b, Math.abs(p2 - p1), x.color))
          if (ch.labels) {
            const s = label(num(v0), labels[i])
            // Stacked pieces carry their label inside, the rest past the end.
            if (ch.stacked) out.push(horiz ? txt((p1 + p2) / 2, s0 + b / 2 + 3.5, s, { fill: '#fff', size: 9 }) : txt(s0 + b / 2, (p1 + p2) / 2 + 3.5, s, { fill: '#fff', size: 9 }))
            else out.push(horiz ? txt(v >= 0 ? Math.max(p1, p2) + 3 : Math.min(p1, p2) - 3, s0 + b / 2 + 3.5, s, { anchor: v >= 0 ? 'start' : 'end', size: 9 }) : txt(s0 + b / 2, v >= 0 ? Math.min(p1, p2) - 3 : Math.max(p1, p2) + 10, s, { size: 9 }))
          }
        }))
        return svg()
      }
      const acc = new Array(n).fill(0)
      const tops = series.map((x) => Array.from({ length: n }, (_, i) => {
        let v = num(x.vals[i])
        if (v === null) return null
        v = share(v, i)
        if (ch.stacked) { acc[i] += v; return acc[i] }
        return v
      }))
      if (ch.type === 'area') {
        tops.forEach((t, si) => {
          const base = ch.stacked && si ? tops[si - 1] : null
          const ptsTop = t.map((v, i) => `${f1(centre(i))},${f1(vpos(v ?? 0))}`)
          const ptsBase = t.map((_, i) => `${f1(centre(i))},${f1(vpos(base ? base[i] ?? 0 : 0))}`).reverse()
          out.push(`<polygon points="${[...ptsTop, ...ptsBase].join(' ')}" fill="#${series[si].color}" fill-opacity="${ch.stacked ? 1 : 0.8}"/>`)
        })
        if (ch.labels) tops.forEach((t, si) => t.forEach((v, i) => { if (v !== null) out.push(txt(centre(i), vpos(v) - 5, label(num(series[si].vals[i]), labels[i]), { size: 9 })) }))
        return svg()
      }
      tops.forEach((t, si) => {
        let d = ''
        t.forEach((v, i) => { if (v === null) return; d += `${d && t[i - 1] !== null ? 'L' : 'M'}${f1(centre(i))} ${f1(vpos(v))}` })
        out.push(`<path d="${d}" fill="none" stroke="#${series[si].color}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`)
        if (ch.labels) t.forEach((v, i) => { if (v !== null) out.push(txt(centre(i), vpos(v) - 6, label(num(series[si].vals[i]), labels[i]), { size: 9 })) })
      })
      return svg()
    }

    // The sheet's selection wears Excel's green.
    const SEL = '#217346'; const SEL_FILL = 'rgba(33,115,70,0.10)'; const SEL_HEAD = '#dcebe2'; const SEL_TEXT = '#185c37'
    class XlsxEditor {
      constructor (frame, doc) {
        this.frame = frame
        this.doc = doc
        this.queue = []
        this.overlay = new Map()
        this.seq = 0
        this.sheetIx = 0
        this.sel = { r1: 1, c1: 1, r2: 1, c2: 1, ar: 1, ac: 1 }
        this.editing = null
        this.z = 1
        this.build()
      }

      build () {
        const f = this.frame
        const tb = new Toolbar(f.tools, f.strip)
        fileTab(tb, f)
        const ro = () => !this.doc.writable
        const style = (props) => { if (!ro()) this.styleSel(props) }
        this.ui = {}
        tb.tab('开始')
        this.ui.b = tb.button('B', '加粗（⌘B）', () => style({ bold: !this.cur().b }), { cls: 'q-b' })
        this.ui.i = tb.button('I', '倾斜（⌘I）', () => style({ italic: !this.cur().i }), { cls: 'q-i' })
        this.ui.u = tb.button('U', '下划线（⌘U）', () => style({ underline: !this.cur().u }), { cls: 'q-u' })
        this.ui.size = tb.select('字号', [{ value: '', label: '字号' }, ...[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((n) => ({ value: String(n), label: String(n) }))], (v) => { if (v) style({ size: +v }) }, { width: '58px' })
        tb.color('文字颜色', 'A', (c) => style({ color: c }), 'C00000')
        tb.color('填充颜色', '▧', (c) => style({ fill: c }), 'FFF2CC')
        tb.button('', '清除填充', () => style({ fill: null }), { ico: 'trash', cls: 'q-mini' })
        this.ui.painter = tb.button('', '格式刷：单击刷一次，双击连续刷，Esc 退出', () => { if (!ro()) this.startPainter(false) }, { ico: 'brush' })
        this.ui.painter.addEventListener('dblclick', () => { if (!ro()) this.startPainter(true) })
        tb.sep()
        this.ui.al = tb.button('', '左对齐', () => style({ align: this.cur().align === 'left' ? 'general' : 'left' }), { ico: 'left' })
        this.ui.ac = tb.button('', '居中', () => style({ align: this.cur().align === 'center' ? 'general' : 'center' }), { ico: 'center' })
        this.ui.ar = tb.button('', '右对齐', () => style({ align: this.cur().align === 'right' ? 'general' : 'right' }), { ico: 'right' })
        this.ui.wrap = tb.button('', '自动换行', () => style({ wrap: !this.cur().wrap }), { ico: 'wrap' })
        this.ui.merge = tb.button('', '合并/取消合并', () => this.mergeSel(), { ico: 'merge' })
        this.ui.border = tb.select('边框', BORDERS, (v) => { if (v) this.border(v); this.ui.border.value = '' }, { width: '70px' })
        tb.sep()
        this.ui.fmt = tb.select('数字格式', FORMATS, (v) => style({ format: v }), { width: '96px' })
        this.ui.decInc = tb.button('', '增加小数位数', () => this.decimals(1), { ico: 'decInc' })
        this.ui.decDec = tb.button('', '减少小数位数', () => this.decimals(-1), { ico: 'decDec' })
        tb.sep()
        this.ui.cf = tb.button('条件格式', '条件格式：突出显示、数据条、色阶', (e) => this.cfMenu(e), { ico: 'cf' })
        tb.button('单元格样式 ▾', '单元格样式', (e) => { const r = e.currentTarget.getBoundingClientRect(); menu(r.left, r.bottom + 2, CELL_STYLES.map((x) => x === '-' ? '-' : { label: x[0], disabled: ro(), fn: () => style(x[1]) })) })
        tb.sep()
        tb.button('', '在上方插入行', () => this.lines('insert_rows'), { ico: 'rowIns' })
        tb.button('', '删除选中行', () => this.lines('delete_rows'), { ico: 'rowDel' })
        tb.button('', '在左侧插入列', () => this.lines('insert_cols'), { ico: 'colIns' })
        tb.button('', '删除选中列', () => this.lines('delete_cols'), { ico: 'colDel' })
        tb.tab('插入')
        this.ui.chart = tb.button('图表', '插入图表：先选中数据区域', (e) => { const r = e.currentTarget.getBoundingClientRect(); this.chartTypeMenu(r.left, r.bottom + 2) }, { ico: 'chart' })
        tb.sep()
        tb.button('插入行', '在选区上方插入行', () => this.lines('insert_rows'))
        tb.button('插入列', '在选区左侧插入列', () => this.lines('insert_cols'))
        this.ui.addSheet = tb.button('新工作表', '新建工作表', () => this.tabs.querySelector('.qx-add')?.click())
        tb.sep()
        tb.button('图片', '插入图片', () => this.insertPic(), { ico: 'image' })
        tb.button('链接', '插入链接（⌘K）', () => this.editLink(), { ico: 'link' })
        tb.button('批注', '新建批注（⇧F2）', () => this.editNote(), { ico: 'comment' })
        tb.tab('页面布局')
        const pageMenu = (items) => (e) => { const r = e.currentTarget.getBoundingClientRect(); menu(r.left, r.bottom + 2, items().map((x) => x === '-' ? x : { ...x, disabled: ro() || x.disabled })) }
        const pg = () => this.sheet.page ?? {}
        const tick = (on, label) => (on ? '✓ ' : '\u3000') + label
        tb.button('纸张方向 ▾', '纸张方向', pageMenu(() => [['portrait', '纵向'], ['landscape', '横向']].map(([v, l]) => ({ label: tick(pg().orientation === v, l), fn: () => this.page({ orientation: v }) }))))
        tb.button('纸张大小 ▾', '纸张大小', pageMenu(() => ['A4', 'A3', 'A5', 'B5', 'Letter', 'Legal'].map((v) => ({ label: tick(pg().paper === v, v), fn: () => this.page({ paper: v }) }))))
        tb.button('页边距 ▾', '页边距', pageMenu(() => [['normal', '常规'], ['narrow', '窄'], ['wide', '宽']].map(([v, l]) => ({ label: tick(pg().margins === v, l), fn: () => this.page({ margins: v }) }))))
        tb.button('打印区域 ▾', '打印区域', pageMenu(() => [{ label: '设置打印区域', fn: () => this.page({ area: this.selRef(true) }) }, { label: '取消打印区域', disabled: !pg().area, fn: () => this.page({ area: null }) }]))
        tb.button('打印标题', '每页都打印的标题行', () => this.printTitles())
        tb.sep()
        tb.button('缩放 ▾', '打印缩放', pageMenu(() => [{ label: tick(!pg().fit, '无缩放'), fn: () => this.page({ fit: null }) }, { label: tick(pg().fit?.w === 1 && !pg().fit?.h, '所有列调整为一页'), fn: () => this.page({ fit: { w: 1, h: 0 } }) }, { label: tick(pg().fit?.w === 1 && pg().fit?.h === 1, '整张表调整为一页'), fn: () => this.page({ fit: { w: 1, h: 1 } }) }]))
        this.ui.gridPrint = tb.button('打印网格线', '打印时带网格线', () => this.page({ grid: !pg().grid }))
        tb.tab('公式')
        tb.button('插入函数', '插入函数（⇧F3）', () => this.insertFunction(), { ico: 'fx' })
        tb.button('自动求和', '自动求和（⌥=）', () => this.autoSum('SUM'), { ico: 'sum' })
        tb.button('▾', '平均值、计数、最大值、最小值', (e) => { const r = e.currentTarget.getBoundingClientRect(); menu(r.left, r.bottom + 2, [['SUM', '求和'], ['AVERAGE', '平均值'], ['COUNT', '计数'], ['MAX', '最大值'], ['MIN', '最小值']].map(([f, l]) => ({ label: l, disabled: ro(), fn: () => this.autoSum(f) }))) }, { cls: 'q-mini' })
        tb.sep()
        for (const [cat, list] of FN_CATS.slice(1)) {
          tb.button(cat + ' ▾', cat + '函数', (e) => {
            const r = e.currentTarget.getBoundingClientRect()
            menu(r.left, r.bottom + 2, list.map((n) => ({ label: n, kbd: FN_HELP[n].replace(/^\S+\(.*?\)\s*/, ''), disabled: ro(), fn: () => this.startFormula(n) })))
          })
        }
        tb.sep()
        tb.button('名称管理器', '名称管理器', () => this.nameManager(), { ico: 'names' })
        tb.button('定义名称', '把选区定义为名称', () => this.defineName())
        tb.tab('数据')
        this.ui.sortUp = tb.button('升序', '升序排序', () => this.sortSel(false), { ico: 'sortUp' })
        this.ui.sortDown = tb.button('降序', '降序排序', () => this.sortSel(true), { ico: 'sortDown' })
        this.ui.filter = tb.button('筛选', '筛选', () => this.toggleFilter(), { ico: 'filter' })
        tb.button('排序…', '按多列排序', () => this.sortDialog())
        tb.sep()
        tb.button('删除重复项', '删除重复项', () => this.dedupeDialog())
        tb.button('分列', '按分隔符把一列拆成多列', () => this.splitColumn())
        tb.sep()
        tb.button('组合', '组合所选行或列（⌥⇧→）', () => this.outline(1), { ico: 'group' })
        tb.button('取消组合', '取消组合（⌥⇧←）', () => this.outline(-1))
        tb.button('隐藏明细', '折叠所选的组', () => this.detail(true))
        tb.button('显示明细', '展开所选的组', () => this.detail(false))
        tb.sep()
        this.ui.dv = tb.button('数据验证', '数据验证：下拉列表、数值范围', (e) => { const r = e.currentTarget.getBoundingClientRect(); this.dvMenu(r.left, r.bottom + 2) }, { ico: 'dv' })
        tb.sep()
        this.ui.fillDown = tb.button('向下填充', '向下填充（⌘D）', () => this.fillDown())
        tb.tab('审阅')
        this.ui.noteBtn = tb.button('新建批注', '新建或编辑批注（⇧F2）', () => this.editNote(), { ico: 'comment' })
        tb.button('删除批注', '删除所选单元格的批注', () => this.deleteNotes())
        this.ui.showNotes = tb.button('显示所有批注', '显示所有批注', () => { this.showNotes = !this.showNotes; this.ui.showNotes.classList.toggle('on', this.showNotes); this.draw() })
        tb.sep()
        this.ui.protect = tb.button('保护工作表', '保护工作表：只能改未锁定的单元格', () => this.toggleProtect(), { ico: 'lock' })
        this.ui.lockCell = tb.button('锁定单元格', '保护工作表后，锁定的单元格不能修改', () => { if (!ro()) this.styleSel({ locked: !!this.cur().unlocked }) })
        tb.tab('视图')
        this.ui.freeze = tb.select('冻结', [{ value: '', label: '冻结窗格…' }, { value: 'row', label: '冻结首行' }, { value: 'col', label: '冻结首列' }, { value: 'here', label: '冻结至选中单元格' }, { value: 'none', label: '取消冻结' }], (v) => { this.freeze(v); this.ui.freeze.value = '' }, { width: '92px' })
        this.viewZoom = tb.select('显示比例', [50, 75, 90, 100, 125, 150, 200].map((n) => ({ value: String(n), label: n + '%' })), (v) => { this.zoomSel.value = v; this.zoomSel.dispatchEvent(new Event('change')) }, { width: '68px' })
        this.viewZoom.value = '100'
        tb.restore('xlsx')

        this.nameBox = el('input', { class: 'qx-name', 'aria-label': '单元格', spellcheck: 'false' })
        this.fbar = el('input', { class: 'qx-fx', 'aria-label': '编辑栏', spellcheck: 'false' })
        const fxl = el('button', { type: 'button', class: 'qx-fxl', title: '插入函数' }, 'fx')
        fxl.addEventListener('mousedown', (e) => e.preventDefault())
        fxl.addEventListener('click', () => this.insertFunction())
        const fx = el('div', { class: 'qx-fbar' }, this.nameBox, fxl, this.fbar)
        this.canvas = el('canvas', { class: 'qx-canvas' })
        this.spacer = el('div', { class: 'qx-spacer' })
        this.scroll = el('div', { class: 'qx-scroll' }, this.spacer)
        this.inp = el('textarea', { class: 'qx-input', spellcheck: 'false', autocomplete: 'off', 'aria-label': '单元格内容' })
        this.guide = el('div', { class: 'qx-guide', hidden: true })
        this.tip = el('div', { class: 'qx-tip', hidden: true })
        this.pop = el('div', { class: 'qx-pop', hidden: true })
        this.noteLayer = el('div', { class: 'qx-notes', style: { left: HW + 'px', top: HH + 'px' } })
        this.chartLayer = el('div', { class: 'qx-charts', style: { left: HW + 'px', top: HH + 'px' } })
        this.grid = el('div', { class: 'qx-grid' }, this.canvas, this.scroll, this.chartLayer, this.noteLayer, this.inp, this.guide, this.tip, this.pop)
        this.scroll.addEventListener('pointerleave', () => { this.pop.hidden = true })
        this.tabs = el('div', { class: 'qx-tabs' })
        this.stats = el('div', { class: 'qx-stats' })
        this.zoomSel = el('select', { class: 'q-sel qx-zoom', title: '缩放', 'aria-label': '缩放' }, ...[50, 75, 90, 100, 125, 150, 200].map((n) => el('option', { value: String(n), selected: n === 100 }, n + '%')))
        this.zoomSel.addEventListener('change', () => { this.z = +this.zoomSel.value / 100; this.viewZoom.value = this.zoomSel.value; this.zoomer.update(this.z * 100); this.layout(); this.draw(); this.placeInput(); this.focus() })
        this.zoomer = zoomSlider((v) => { this.z = v / 100; syncZoomSel(this.zoomSel, v); syncZoomSel(this.viewZoom, v); this.zoomer.update(v); this.layout(); this.draw(); this.placeInput() })
        this.ac = el('div', { class: 'qx-ac', hidden: true })
        this.root = el('div', { class: 'qx' }, fx, this.grid, el('div', { class: 'qx-foot' }, this.tabs, this.stats, this.zoomSel, this.zoomer.node), this.ac)
        f.main.append(this.root)
        this.scroll.addEventListener('contextmenu', (e) => this.context(e))
        this.scroll.addEventListener('wheel', (e) => {
          if (!mod(e)) return
          e.preventDefault()
          const steps = [50, 75, 90, 100, 125, 150, 200]
          const i = steps.indexOf(Math.round(this.z * 100))
          const n = steps[Math.max(0, Math.min(steps.length - 1, (i < 0 ? 3 : i) + (e.deltaY < 0 ? 1 : -1)))]
          this.zoomSel.value = this.viewZoom.value = String(n); this.z = n / 100; this.zoomer.update(n); this.layout(); this.draw(); this.placeInput()
        }, { passive: false })

        this.scroll.addEventListener('scroll', () => { this.grow(); this.draw(); this.placeInput() })
        this.scroll.addEventListener('pointerdown', (e) => this.down(e))
        this.scroll.addEventListener('pointermove', (e) => this.hover(e))
        this.scroll.addEventListener('dblclick', (e) => { const h = this.hit(e); if (h.kind === 'cell') this.beginEdit('edit') })
        this.inp.addEventListener('keydown', (e) => this.key(e))
        this.inp.addEventListener('compositionstart', () => { if (!this.editing) this.beginEdit('enter', true) })
        this.inp.addEventListener('input', () => {
          if (!this.editing) this.beginEdit('enter', true)
          this.fbar.value = this.inp.value
          this.pointing = null
          this.fitInput()
          this.suggest()
          if (this.inp.value.startsWith('=') || this.hadRefs) this.draw()
        })
        this.inp.addEventListener('paste', (e) => this.paste(e))
        this.inp.addEventListener('copy', (e) => this.copy(e, false))
        this.inp.addEventListener('cut', (e) => this.copy(e, true))
        this.fbar.addEventListener('focus', () => { if (!this.editing && this.doc.writable) { this.editing = { mode: 'edit', bar: true, orig: this.input(this.sel.ar, this.sel.ac) } } })
        this.fbar.addEventListener('keydown', (e) => {
          if (this.acKey(e, this.fbar)) return
          if (e.key === 'Enter') { e.preventDefault(); this.inp.value = this.fbar.value; this.commit(1, 0) } else if (e.key === 'Escape') { e.preventDefault(); this.cancel() }
        })
        this.fbar.addEventListener('input', () => { if (this.editing && !this.editing.bar) this.inp.value = this.fbar.value; this.pointing = null; this.suggest(); if (this.fbar.value.startsWith('=') || this.hadRefs) this.draw() })
        this.fbar.addEventListener('blur', () => setTimeout(() => this.hideAc(), 150))
        this.nameBox.addEventListener('keydown', (e) => {
          if (e.key !== 'Enter') return
          e.preventDefault()
          const m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(this.nameBox.value.trim())
          if (!m) return
          const col = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
          const r1 = +m[2]; const c1 = col(m[1]); const r2 = m[4] ? +m[4] : r1; const c2 = m[3] ? col(m[3]) : c1
          this.select(Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2), r1, c1)
          this.focus()
          this.refreshBar()
        })
        this.ro = new ResizeObserver(() => this.resize())
        this.ro.observe(this.grid)
      }

      dispose () { this.ro?.disconnect() }

      // ------------------------------------------------ model

      load (model, { keep } = {}) {
        const name = this.sheet?.name
        this.chartBusy = false
        this.model = model
        const visible = model.sheets.map((s, i) => [s, i]).filter(([s]) => !s.state || s.state === 'visible')
        let ix = keep && name ? model.sheets.findIndex((s) => s.name === name) : -1
        if (ix < 0) ix = visible[0]?.[1] ?? 0
        this.useSheet(ix, keep)
      }

      useSheet (ix, keep) {
        this.sheetIx = ix
        this.sheet = this.model.sheets[ix]
        const sh = this.sheet
        this.colMap = new Map()
        for (const c of sh.cols ?? []) for (let k = c.min; k <= Math.min(c.max, (sh.maxC ?? 0) + 200); k++) this.colMap.set(k, c.hidden ? 0 : c.px)
        this.merges = sh.merges ?? []
        this.mergeAt = new Map()
        for (const m of this.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) this.mergeAt.set(`${r},${c}`, m)
        this.nr = Math.max((sh.maxR ?? 0) + 100, 200, ...(sh.charts ?? []).map((c) => c.to.r + 30))
        this.nc = Math.max((sh.maxC ?? 0) + 26, 40, ...(sh.charts ?? []).map((c) => c.to.c + 6))
        this.layout()
        this.buildCharts()
        if (!keep) { this.sel = { r1: 1, c1: 1, r2: 1, c2: 1, ar: 1, ac: 1 }; this.scroll.scrollTop = 0; this.scroll.scrollLeft = 0 }
        this.renderTabs()
        this.resize()
        this.refreshBar()
      }

      layout () {
        const sh = this.sheet
        this.xs = new Float64Array(this.nc + 2)
        for (let c = 1; c <= this.nc + 1; c++) this.xs[c] = this.xs[c - 1] + (c > 1 ? this.w(c - 1) : 0)
        this.ys = new Float64Array(this.nr + 2)
        for (let r = 1; r <= this.nr + 1; r++) this.ys[r] = this.ys[r - 1] + (r > 1 ? this.h(r - 1) : 0)
        this.fr = sh.frozen?.r ?? 0
        this.fc = sh.frozen?.c ?? 0
        this.spacer.style.width = HW + this.xs[this.nc + 1] + 'px'
        this.spacer.style.height = HH + this.ys[this.nr + 1] + 'px'
      }

      w (c) { const v = this.colMap.get(c); return (v === undefined ? this.sheet.defW ?? 64 : v) * this.z }
      h (r) { const v = this.sheet.rows?.[r]; return (v === undefined ? this.sheet.defH ?? 20 : v) * this.z }
      X (c) { return this.xs[Math.min(c, this.nc + 1)] }
      Y (r) { return this.ys[Math.min(r, this.nr + 1)] }

      grow () {
        const s = this.scroll
        let changed = false
        if (s.scrollTop + s.clientHeight > HH + this.ys[this.nr + 1] - 300 && this.nr < 1048576) { this.nr = Math.min(1048576, this.nr + 300); changed = true }
        if (s.scrollLeft + s.clientWidth > HW + this.xs[this.nc + 1] - 300 && this.nc < 16384) { this.nc = Math.min(16384, this.nc + 20); changed = true }
        if (changed) this.layout()
      }

      cell (r, c) { return this.sheet.cells?.[`${r},${c}`] }
      shown (r, c) {
        const o = this.overlay.get(`${this.sheet.name}|${r},${c}`)
        if (o) return o.text.startsWith('=') ? '…' : o.text
        return this.cell(r, c)?.[0] ?? ''
      }
      input (r, c) {
        const o = this.overlay.get(`${this.sheet.name}|${r},${c}`)
        return o ? o.text : this.cell(r, c)?.[1] ?? ''
      }
      styleOf (r, c) { const x = this.cell(r, c); return (x && this.model.styles[x[2]]) || {} }
      cur () { return this.styleOf(this.sel.ar, this.sel.ac) }
      // paintOf is a cell's style with its conditional formats laid over.
      paintOf (r, c) {
        const st = this.styleOf(r, c)
        const x = this.sheet.cfx?.[`${r},${c}`]
        if (!x) return st
        const out = { ...st }
        if (x.fill || x.scale) out.fill = x.fill ?? x.scale
        if (x.color) out.color = x.color
        if (x.b !== undefined) out.b = x.b
        if (x.i !== undefined) out.i = x.i
        if (x.u) out.u = true
        if (x.bar) out.bar = x.bar
        return out
      }

      // ------------------------------------------------ geometry

      resize () {
        const r = this.grid.getBoundingClientRect()
        const dpr = window.devicePixelRatio || 1
        this.vw = Math.max(1, Math.floor(r.width))
        this.vh = Math.max(1, Math.floor(r.height))
        this.canvas.width = this.vw * dpr
        this.canvas.height = this.vh * dpr
        this.canvas.style.width = this.vw + 'px'
        this.canvas.style.height = this.vh + 'px'
        this.dpr = dpr
        this.grow()
        this.draw()
        this.placeInput()
      }

      sx (c) { return c <= this.fc ? HW + this.X(c) : HW + this.X(c) - this.scroll.scrollLeft }
      sy (r) { return r <= this.fr ? HH + this.Y(r) : HH + this.Y(r) - this.scroll.scrollTop }
      FW () { return this.X(this.fc + 1) }
      FH () { return this.Y(this.fr + 1) }

      find (arr, v, n) {
        let lo = 1; let hi = n
        while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (arr[mid] <= v) lo = mid; else hi = mid - 1 }
        return lo
      }
      colAt (x) { const ax = x - HW < this.FW() ? x - HW : x - HW + this.scroll.scrollLeft; return Math.max(1, this.find(this.xs, ax, this.nc)) }
      rowAt (y) { const ay = y - HH < this.FH() ? y - HH : y - HH + this.scroll.scrollTop; return Math.max(1, this.find(this.ys, ay, this.nr)) }

      visibleCols () {
        const out = []
        for (let c = 1; c <= this.fc; c++) out.push(c)
        const start = this.find(this.xs, this.scroll.scrollLeft + this.FW(), this.nc)
        for (let c = Math.max(start, this.fc + 1); c <= this.nc && this.sx(c) < this.vw; c++) out.push(c)
        return out
      }
      visibleRows () {
        const out = []
        for (let r = 1; r <= this.fr; r++) out.push(r)
        const start = this.find(this.ys, this.scroll.scrollTop + this.FH(), this.nr)
        for (let r = Math.max(start, this.fr + 1); r <= this.nr && this.sy(r) < this.vh; r++) out.push(r)
        return out
      }

      rectOf (r1, c1, r2, c2) {
        const x1 = this.sx(c1); const y1 = this.sy(r1)
        return { x: x1, y: y1, w: this.sx(c2 + 1 > this.fc && c2 <= this.fc ? c2 + 1 : c2 + 1) - x1, h: this.sy(r2 + 1) - y1 }
      }

      // ------------------------------------------------ drawing

      draw () {
        if (!this.sheet || !this.vw) return
        const g = this.canvas.getContext('2d')
        g.setTransform(this.dpr, 0, 0, this.dpr, 0, 0)
        g.fillStyle = '#fff'
        g.fillRect(0, 0, this.vw, this.vh)
        if (this.sheet.unsupported) { g.fillStyle = '#666'; g.font = `13px ${FONT}`; g.fillText(this.sheet.unsupported, 20, 40); return }
        const rows = this.visibleRows()
        const cols = this.visibleCols()
        const fw = HW + this.FW(); const fh = HH + this.FH()
        const regions = [
          [rows.filter((r) => r > this.fr), cols.filter((c) => c > this.fc), fw, fh, this.vw, this.vh],
          [rows.filter((r) => r <= this.fr), cols.filter((c) => c > this.fc), fw, HH, this.vw, fh],
          [rows.filter((r) => r > this.fr), cols.filter((c) => c <= this.fc), HW, fh, fw, this.vh],
          [rows.filter((r) => r <= this.fr), cols.filter((c) => c <= this.fc), HW, HH, fw, fh]
        ]
        for (const [rs, cs, x1, y1, x2, y2] of regions) {
          if (!rs.length || !cs.length || x2 <= x1 || y2 <= y1) continue
          g.save()
          g.beginPath(); g.rect(x1, y1, x2 - x1, y2 - y1); g.clip()
          this.drawCells(g, rs, cs, x2)
          this.drawSelection(g)
          this.drawPins(g)
          this.drawRefs(g)
          g.restore()
        }
        this.drawHeaders(g, rows, cols)
        if (this.fr) { g.strokeStyle = '#9aa0a6'; g.lineWidth = 1; g.beginPath(); g.moveTo(0, fh + 0.5); g.lineTo(this.vw, fh + 0.5); g.stroke() }
        if (this.fc) { g.strokeStyle = '#9aa0a6'; g.lineWidth = 1; g.beginPath(); g.moveTo(fw + 0.5, 0); g.lineTo(fw + 0.5, this.vh); g.stroke() }
        this.placeCharts()
        this.placeNotes()
        this.updateStats()
      }

      font (st) {
        const px = (st.size ? +st.size : 11) * 4 / 3 * this.z
        return `${st.i ? 'italic ' : ''}${st.b ? '700 ' : ''}${px.toFixed(1)}px ${FONT}`
      }

      drawCells (g, rows, cols, clipRight) {
        // Fills.
        for (const r of rows) {
          for (const c of cols) {
            if (this.mergeAt.has(`${r},${c}`)) continue
            const st = this.paintOf(r, c)
            if (st.fill) { g.fillStyle = '#' + st.fill; g.fillRect(this.sx(c), this.sy(r), this.w(c), this.h(r)) }
            if (st.bar) this.dataBar(g, st.bar, this.sx(c), this.sy(r), this.w(c), this.h(r))
          }
        }
        // Grid lines.
        g.strokeStyle = '#e3e3e3'
        g.lineWidth = 1
        g.beginPath()
        for (const c of cols) { const x = Math.round(this.sx(c) + this.w(c)) + 0.5; g.moveTo(x, this.sy(rows[0])); g.lineTo(x, this.sy(rows[rows.length - 1]) + this.h(rows[rows.length - 1])) }
        for (const r of rows) { const y = Math.round(this.sy(r) + this.h(r)) + 0.5; g.moveTo(this.sx(cols[0]), y); g.lineTo(this.sx(cols[cols.length - 1]) + this.w(cols[cols.length - 1]), y) }
        g.stroke()
        // Text and borders.
        g.textBaseline = 'middle'
        const heads = []
        for (const r of rows) {
          for (let k = 0; k < cols.length; k++) {
            const c = cols[k]
            if (this.mergeAt.has(`${r},${c}`)) continue
            const text = this.shown(r, c)
            const st = this.paintOf(r, c)
            const x = this.sx(c); const y = this.sy(r)
            if (st.bd) this.sides(g, st.bd, x, y, this.w(c), this.h(r))
            if (this.filterHead(r, c)) heads.push([r, c])
            if (this.sheet.notes?.[`${r},${c}`] && this.w(c) > 6) { g.fillStyle = '#d93025'; g.beginPath(); g.moveTo(x + this.w(c) - 6, y + 1); g.lineTo(x + this.w(c), y + 1); g.lineTo(x + this.w(c), y + 7); g.closePath(); g.fill() }
            if (text === '' || !this.w(c) || !this.h(r)) continue
            // Left-aligned text runs on into empty cells to its right.
            const kind = this.cell(r, c)?.[3] ?? 0
            let width = this.w(c)
            const align = st.align && st.align !== 'general' ? st.align : kind === 1 ? 'right' : kind === 2 ? 'center' : 'left'
            if (align === 'left' && !st.wrap) {
              for (let n = c + 1; n <= Math.min(c + 30, this.nc); n++) { if (this.shown(r, n) !== '' || this.mergeAt.has(`${r},${n}`)) break; width += this.w(n) }
              width = Math.min(width, clipRight - x)
            }
            this.text(g, text, st, x, y, width, this.h(r), align, kind)
          }
        }
        // Merged cells, drawn whole.
        const rset = new Set(rows); const cset = new Set(cols)
        for (const m of this.merges) {
          let hitR = false; let hitC = false
          for (let r = m.r1; r <= m.r2 && !hitR; r++) hitR = rset.has(r)
          for (let c = m.c1; c <= m.c2 && !hitC; c++) hitC = cset.has(c)
          if (!hitR || !hitC) continue
          const x = this.sx(m.c1); const y = this.sy(m.r1)
          const w = this.X(m.c2 + 1) - this.X(m.c1); const h = this.Y(m.r2 + 1) - this.Y(m.r1)
          const st = this.paintOf(m.r1, m.c1)
          g.fillStyle = st.fill ? '#' + st.fill : '#fff'
          g.fillRect(x + 1, y + 1, w - 1, h - 1)
          if (st.bar) this.dataBar(g, st.bar, x, y, w, h)
          g.strokeStyle = '#e3e3e3'
          g.strokeRect(Math.round(x) + 0.5, Math.round(y) + 0.5, w, h)
          if (st.bd) this.sides(g, st.bd, x, y, w, h)
          const text = this.shown(m.r1, m.c1)
          if (text) {
            const kind = this.cell(m.r1, m.c1)?.[3] ?? 0
            const align = st.align && st.align !== 'general' ? st.align : kind === 1 ? 'right' : 'left'
            this.text(g, text, st, x, y, w, h, align, kind)
          }
        }
        for (const [r, c] of heads) this.filterButton(g, r, c)
      }

      dataBar (g, [pct, color], x, y, w, h) {
        const bw = Math.max(0, (w - 4) * pct)
        if (bw < 1 || h < 6) return
        const grd = g.createLinearGradient(x + 2, 0, x + 2 + bw, 0)
        grd.addColorStop(0, '#' + color)
        grd.addColorStop(1, '#' + color + '33')
        g.fillStyle = grd
        g.fillRect(x + 2, y + 2, bw, h - 4)
        g.strokeStyle = '#' + color
        g.lineWidth = 1
        g.strokeRect(x + 2.5, y + 2.5, bw - 1, h - 5)
      }

      // sides draws a cell's own borders, side by side.
      sides (g, bd, x, y, w, h) {
        const W = { thin: 1, hair: 1, dotted: 1, dashed: 1, mediumDashed: 2, medium: 2, thick: 3, double: 3, dashDot: 1, mediumDashDot: 2, slantDashDot: 2, dashDotDot: 1, mediumDashDotDot: 2 }
        const line = (k, x1, y1, x2, y2) => {
          const s = bd[k]
          if (!s) return
          g.save()
          g.strokeStyle = '#' + (s[1] || '000000')
          g.lineWidth = W[s[0]] ?? 1
          if (/dash/i.test(s[0])) g.setLineDash([4, 2]); else if (s[0] === 'dotted' || s[0] === 'hair') g.setLineDash([1, 2])
          g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke()
          g.restore()
        }
        const L = Math.round(x) + 0.5; const T = Math.round(y) + 0.5; const R = Math.round(x + w) + 0.5; const B = Math.round(y + h) + 0.5
        line('t', L, T, R, T); line('b', L, B, R, B); line('l', L, T, L, B); line('r', R, T, R, B)
      }

      filterHead (r, c) {
        const af = this.sheet.autoFilter
        return af && r === af.r1 && c >= af.c1 && c <= af.c2
      }

      filterButton (g, r, c) {
        const sz = Math.min(16, this.h(r) - 3)
        if (sz < 8 || this.w(c) < sz + 4) return
        const x = this.sx(c) + this.w(c) - sz - 2; const y = this.sy(r) + (this.h(r) - sz) / 2
        const on = Array.isArray(this.sheet.autoFilter.cols?.[c])
        g.save()
        g.fillStyle = on ? '#dbe7f9' : '#f6f7f8'; g.strokeStyle = on ? '#1a73e8' : '#b6bbc2'
        g.fillRect(x, y, sz, sz); g.strokeRect(x + 0.5, y + 0.5, sz - 1, sz - 1)
        g.fillStyle = on ? '#1a73e8' : '#555'
        g.beginPath()
        if (on) { g.moveTo(x + 3, y + 4); g.lineTo(x + sz - 3, y + 4); g.lineTo(x + sz / 2 + 1, y + sz / 2 + 1); g.lineTo(x + sz / 2 + 1, y + sz - 3); g.lineTo(x + sz / 2 - 1, y + sz - 4); g.lineTo(x + sz / 2 - 1, y + sz / 2 + 1) } else { g.moveTo(x + 4, y + sz / 2 - 2); g.lineTo(x + sz - 4, y + sz / 2 - 2); g.lineTo(x + sz / 2, y + sz / 2 + 3) }
        g.closePath(); g.fill()
        g.restore()
      }

      text (g, text, st, x, y, w, h, align, kind) {
        g.save()
        g.beginPath(); g.rect(x + 1, y + 1, w - 2, h - 1); g.clip()
        g.font = this.font(st)
        g.fillStyle = st.color && st.color !== '000000' ? '#' + st.color : '#000'
        if (kind === 2 && /^#/.test(text)) g.fillStyle = '#c00000'
        const pad = 3
        // A number that does not fit shows as ####, as in Excel, rather
        // than losing digits off one side.
        if (kind === 1 && !st.wrap && g.measureText(text).width > w - pad * 2) {
          // A plain decimal (General format) gives up decimals first.
          const m = /^(-?\d+)\.(\d+)$/.exec(text)
          let fit = ''
          if (m) for (let d = m[2].length - 1; d >= 0 && !fit; d--) { const t = String(+(+text).toFixed(d)); if (g.measureText(t).width <= w - pad * 2) fit = t }
          text = fit || '#'.repeat(Math.max(1, Math.floor((w - pad * 2) / Math.max(1, g.measureText('#').width))))
        }
        const lines = st.wrap ? this.wrapLines(g, text, w - pad * 2) : [text.replace(/\n/g, ' ')]
        const lh = (st.size ? +st.size : 11) * 4 / 3 * 1.25 * this.z
        const total = lines.length * lh
        const va = st.valign === 'top' ? 'top' : st.valign === 'center' ? 'center' : 'bottom'
        let ty = va === 'top' ? y + 2 + lh / 2 : va === 'center' ? y + h / 2 - total / 2 + lh / 2 : y + h - total - 1 + lh / 2
        if (!st.wrap && va === 'bottom') ty = y + h - lh / 2 - 1
        for (const line of lines) {
          let tx = x + pad
          if (align === 'right') { g.textAlign = 'right'; tx = x + w - pad } else if (align === 'center' || align === 'centerContinuous') { g.textAlign = 'center'; tx = x + w / 2 } else g.textAlign = 'left'
          g.fillText(line, tx, ty)
          if (st.u) { const m = g.measureText(line); const ux = g.textAlign === 'right' ? tx - m.width : g.textAlign === 'center' ? tx - m.width / 2 : tx; g.fillRect(ux, ty + lh * 0.32, m.width, 1) }
          ty += lh
        }
        g.restore()
      }

      wrapLines (g, text, width) {
        const out = []
        for (const para of String(text).split('\n')) {
          let line = ''
          for (const ch of para) {
            if (g.measureText(line + ch).width > width && line) { out.push(line); line = ch } else line += ch
          }
          out.push(line)
        }
        return out
      }

      drawHeaders (g, rows, cols) {
        const s = this.sel
        g.font = `12px ${FONT}`
        g.textAlign = 'center'
        g.textBaseline = 'middle'
        g.fillStyle = '#f3f4f6'
        g.fillRect(0, 0, this.vw, HH)
        g.fillRect(0, 0, HW, this.vh)
        for (const c of cols) {
          const x = this.sx(c); const w = this.w(c)
          if (x + w < HW || x > this.vw || !w) continue
          if (c >= s.c1 && c <= s.c2) { g.fillStyle = SEL_HEAD; g.fillRect(x, 0, w, HH); g.fillStyle = SEL; g.fillRect(x, HH - 2, w, 2) }
          g.fillStyle = c >= s.c1 && c <= s.c2 ? SEL_TEXT : '#555'
          g.save(); g.beginPath(); g.rect(Math.max(HW, x), 0, w, HH); g.clip()
          g.fillText(colName(c), x + w / 2, HH / 2 + 1)
          g.restore()
          g.fillStyle = '#d4d4d4'; g.fillRect(Math.round(x + w), 0, 1, HH)
        }
        for (const r of rows) {
          const y = this.sy(r); const h = this.h(r)
          if (y + h < HH || y > this.vh || !h) continue
          if (r >= s.r1 && r <= s.r2) { g.fillStyle = SEL_HEAD; g.fillRect(0, y, HW, h); g.fillStyle = SEL; g.fillRect(HW - 2, y, 2, h) }
          g.fillStyle = r >= s.r1 && r <= s.r2 ? SEL_TEXT : '#555'
          g.save(); g.beginPath(); g.rect(0, Math.max(HH, y), HW, h); g.clip()
          g.fillText(String(r), HW / 2, y + h / 2 + 1)
          g.restore()
          g.fillStyle = '#d4d4d4'; g.fillRect(0, Math.round(y + h), HW, 1)
        }
        // Grouped rows and columns get a bar per level along the header edge.
        const lv = this.sheet.lv
        if (lv) {
          g.fillStyle = '#8a94a6'
          for (const r of rows) { const n = lv.r?.[r]; if (!n) continue; const y = this.sy(r); if (y + this.h(r) < HH) continue; for (let k = 0; k < n; k++) g.fillRect(2 + k * 3, Math.max(HH, y), 2, this.h(r)) }
          for (const c of cols) { const n = lv.c?.[c]; if (!n) continue; const x = this.sx(c); if (x + this.w(c) < HW) continue; for (let k = 0; k < n; k++) g.fillRect(Math.max(HW, x), 2 + k * 3, this.w(c), 2) }
        }
        g.fillStyle = '#e8e9eb'
        g.fillRect(0, 0, HW, HH)
        g.fillStyle = '#c8c8c8'
        g.fillRect(HW - 1, 0, 1, this.vh)
        g.fillRect(0, HH - 1, this.vw, 1)
      }

      drawSelection (g) {
        const s = this.sel
        let { r1, c1, r2, c2 } = s
        const m = this.mergeAt.get(`${s.ar},${s.ac}`)
        if (r1 === r2 && c1 === c2 && m) ({ r1, c1, r2, c2 } = m)
        const x = this.sx(c1); const y = this.sy(r1)
        const w = this.X(c2 + 1) - this.X(c1); const h = this.Y(r2 + 1) - this.Y(r1)
        if (r1 !== r2 || c1 !== c2) {
          g.fillStyle = SEL_FILL
          g.fillRect(x, y, w, h)
          const am = this.mergeAt.get(`${s.ar},${s.ac}`)
          const ax = this.sx(am ? am.c1 : s.ac); const ay = this.sy(am ? am.r1 : s.ar)
          const aw = am ? this.X(am.c2 + 1) - this.X(am.c1) : this.w(s.ac); const ah = am ? this.Y(am.r2 + 1) - this.Y(am.r1) : this.h(s.ar)
          g.fillStyle = '#fff'
          const st = this.paintOf(s.ar, s.ac)
          if (!st.fill) g.fillRect(ax + 1, ay + 1, aw - 1, ah - 1)
          if (st.bar) this.dataBar(g, st.bar, ax, ay, aw, ah)
          const t = this.shown(s.ar, s.ac)
          if (t) this.text(g, t, st, ax, ay, aw, ah, st.align && st.align !== 'general' ? st.align : (this.cell(s.ar, s.ac)?.[3] === 1 ? 'right' : 'left'), this.cell(s.ar, s.ac)?.[3] ?? 0)
        }
        g.strokeStyle = SEL
        g.lineWidth = 2
        g.strokeRect(x + 1, y + 1, w - 1, h - 1)
        g.fillStyle = SEL
        if (this.doc.writable) { g.fillRect(x + w - 3, y + h - 3, 6, 6); g.strokeStyle = '#fff'; g.lineWidth = 1; g.strokeRect(x + w - 3.5, y + h - 3.5, 7, 7) }
        g.lineWidth = 1
        this.dvBtn = null
        const dv = !this.editing && this.dvAt(s.ar, s.ac)
        if (dv && dv.type === 'list' && !dv.noArrow && this.doc.writable) {
          const am = this.mergeAt.get(`${s.ar},${s.ac}`)
          const bx = this.sx(am ? am.c2 + 1 : s.ac + 1) + 1; const by = this.sy(am ? am.r2 : s.ar)
          const bh = Math.min(18, this.h(am ? am.r2 : s.ar))
          const by2 = by + this.h(am ? am.r2 : s.ar) - bh
          g.fillStyle = '#f6f7f8'; g.strokeStyle = '#b6bbc2'
          g.fillRect(bx, by2, 16, bh); g.strokeRect(bx + 0.5, by2 + 0.5, 15, bh - 1)
          g.fillStyle = '#444'; g.beginPath(); g.moveTo(bx + 4, by2 + bh / 2 - 2); g.lineTo(bx + 12, by2 + bh / 2 - 2); g.lineTo(bx + 8, by2 + bh / 2 + 3); g.closePath(); g.fill()
          this.dvBtn = { x: bx, y: by2, w: 16, h: bh }
        }
        const f = this.fillTo
        if (f) {
          g.save()
          g.setLineDash([4, 3]); g.strokeStyle = SEL; g.lineWidth = 1.5
          g.strokeRect(this.sx(f.c1) + 1, this.sy(f.r1) + 1, this.X(f.c2 + 1) - this.X(f.c1) - 1, this.Y(f.r2 + 1) - this.Y(f.r1) - 1)
          g.restore()
        }
      }

      // formulaRefs finds the cell references in a formula being typed, each
      // with the colour Excel would give it.
      formulaRefs (text) {
        const out = []
        if (!text?.startsWith('=')) return out
        const re = /"(?:[^"]|"")*"|((?:'(?:[^']|'')+'|[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![\w(!])/g
        let m; const colors = new Map()
        while ((m = re.exec(text))) {
          if (!m[2]) continue
          const prev = text[m.index - 1]
          if (prev && /[\w.$]/.test(prev)) continue
          const sheet = m[1] ? m[1].slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : this.sheet.name
          const g = rangeList(m[2])[0]
          if (!g) continue
          const key = sheet + '!' + m[2].replace(/\$/g, '').toUpperCase()
          if (!colors.has(key)) colors.set(key, REF_COLORS[colors.size % REF_COLORS.length])
          out.push({ ...g, sheet, start: m.index, end: m.index + m[0].length, color: colors.get(key) })
        }
        return out
      }

      // ------------------------------------------------ pins

      pinTarget () {
        if (!this.sheet || this.sheet.unsupported) return null
        const s = this.sel
        const m = this.mergeAt.get(`${s.ar},${s.ac}`)
        const single = (s.r1 === s.r2 && s.c1 === s.c2) || (m && m.r1 === s.r1 && m.c1 === s.c1 && m.r2 === s.r2 && m.c2 === s.c2)
        const ref = single ? cellName(s.ar, s.ac) : this.selRef(false)
        const where = `${this.sheet.name}!${ref}`
        let quote; let extra
        if (single) {
          quote = String(this.shown(s.ar, s.ac) ?? '')
          const inp = String(this.input(s.ar, s.ac) ?? '')
          if (inp.startsWith('=')) extra = '公式 ' + inp
        } else {
          const vals = []; let more = false; let seen = 0
          const r2 = Math.min(s.r2, this.nr); const c2 = Math.min(s.c2, this.nc)
          for (let r = s.r1; r <= r2 && !more; r++) {
            for (let c = s.c1; c <= c2; c++) {
              if (++seen > 4000) { more = true; break }
              const v = this.shown(r, c)
              if (v === '' || v === undefined) continue
              if (vals.length === 8) { more = true; break }
              vals.push(String(v))
            }
          }
          quote = vals.join(' | ') + (more ? ' …' : '')
          extra = `${s.r2 - s.r1 + 1} 行 × ${s.c2 - s.c1 + 1} 列`
        }
        const cr = this.canvas.getBoundingClientRect()
        const b = this.rectOf(s.r1, s.c1, Math.min(s.r2, this.nr), Math.min(s.c2, this.nc))
        const left = cr.left + Math.max(HW, Math.min(b.x, this.vw)); const top = cr.top + Math.max(HH, Math.min(b.y, this.vh))
        const bottom = cr.top + Math.max(HH, Math.min(b.y + b.h, this.vh))
        return { where, quote, extra, rect: { left, right: left, top, bottom } }
      }

      // pinAt selects the cells a box on screen covers and marks them.
      pinAt (b) {
        if (!this.sheet || this.sheet.unsupported) return null
        const cr = this.canvas.getBoundingClientRect()
        const x1 = Math.max(HW + 1, b.left - cr.left); const y1 = Math.max(HH + 1, b.top - cr.top)
        const x2 = Math.min(this.vw - 1, b.right - cr.left); const y2 = Math.min(this.vh - 1, b.bottom - cr.top)
        if (x2 < x1 || y2 < y1) return null
        const r1 = this.rowAt(y1); const c1 = this.colAt(x1); const r2 = this.rowAt(y2); const c2 = this.colAt(x2)
        this.select(r1, c1, r2, c2, r1, c1)
        this.refreshBar()
        return this.pinTarget()
      }

      shoot (rect) { return canvasCrop(this.canvas, rect) }

      // snapRect is what a pin's picture shows: the selected cells.
      snapRect () {
        const s = this.sel
        const cr = this.canvas.getBoundingClientRect()
        const b = this.rectOf(s.r1, s.c1, Math.min(s.r2, this.nr), Math.min(s.c2, this.nc))
        const x1 = Math.max(HW, b.x); const y1 = Math.max(HH, b.y); const x2 = Math.min(this.vw, b.x + b.w); const y2 = Math.min(this.vh, b.y + b.h)
        return x2 > x1 && y2 > y1 ? { left: cr.left + x1, top: cr.top + y1, right: cr.left + x2, bottom: cr.top + y2 } : null
      }

      // pinBox reads a spot like 汇总!B4:C6 or 汇总!A:A back into rows and columns.
      pinBox (where) {
        const i = where.lastIndexOf('!')
        if (i < 0) return null
        const sheet = where.slice(0, i)
        const col = (x) => [...x.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
        const ref = where.slice(i + 1).replace(/\$/g, '')
        let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(ref)
        if (m) return { sheet, r1: +m[2], c1: col(m[1]), r2: m[4] ? +m[4] : +m[2], c2: m[3] ? col(m[3]) : col(m[1]) }
        m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(ref)
        if (m) return { sheet, r1: 1, c1: col(m[1]), r2: 1048576, c2: col(m[2]) }
        m = /^(\d+):(\d+)$/.exec(ref)
        if (m) return { sheet, r1: +m[1], c1: 1, r2: +m[2], c2: 16384 }
        return null
      }

      showPins (pins) {
        this.pinMarks = pins.map((p) => this.pinBox(p.where)).filter(Boolean)
        this.draw()
      }

      drawPins (g) {
        for (const b of this.pinMarks ?? []) {
          if (b.sheet !== this.sheet.name) continue
          const r2 = Math.min(b.r2, this.nr); const c2 = Math.min(b.c2, this.nc)
          const x = this.sx(b.c1); const y = this.sy(b.r1)
          const w = this.X(c2 + 1) - this.X(b.c1); const h = this.Y(r2 + 1) - this.Y(b.r1)
          if (x > this.vw || y > this.vh || x + w < 0 || y + h < 0) continue
          g.save()
          g.fillStyle = 'rgba(232,113,10,0.08)'
          g.fillRect(x, y, w, h)
          g.strokeStyle = '#e8710a'; g.lineWidth = 1.5; g.setLineDash([4, 3])
          g.strokeRect(x + 1, y + 1, w - 2, h - 2)
          g.setLineDash([])
          g.fillStyle = '#e8710a'
          g.beginPath(); g.arc(x + 4, y + 4, 3, 0, Math.PI * 2); g.fill()
          g.restore()
        }
      }

      revealPin (p) {
        const b = this.pinBox(p.where)
        if (!b) return
        const ix = this.model.sheets.findIndex((x) => x.name === b.sheet)
        if (ix < 0) return
        if (this.model.sheets[ix].name !== this.sheet.name) this.useSheet(ix)
        this.select(b.r1, b.c1, Math.min(b.r2, this.nr), Math.min(b.c2, this.nc), b.r1, b.c1)
        this.ensureVisible(b.r1, b.c1)
        this.focus()
        this.refreshBar()
      }

      drawRefs (g) {
        const src = this.editing ? (this.editing.bar ? this.fbar : this.inp).value : ''
        const refs = this.formulaRefs(src).filter((x) => x.sheet === this.sheet.name)
        this.hadRefs = refs.length > 0
        for (const x of refs) {
          const px = this.sx(x.c1); const py = this.sy(x.r1)
          const w = this.X(x.c2 + 1) - this.X(x.c1); const h = this.Y(x.r2 + 1) - this.Y(x.r1)
          g.save()
          g.globalAlpha = 0.08; g.fillStyle = x.color; g.fillRect(px, py, w, h)
          g.globalAlpha = 1; g.strokeStyle = x.color; g.lineWidth = 2
          g.strokeRect(px + 1, py + 1, w - 1, h - 1)
          g.fillStyle = x.color
          for (const [cx, cy] of [[px, py], [px + w, py], [px, py + h], [px + w, py + h]]) g.fillRect(cx - 2, cy - 2, 4, 4)
          g.restore()
        }
      }

      // selBox is the selection's rectangle on screen, merges included.
      selBox () {
        const s = this.sel
        let { r1, c1, r2, c2 } = s
        const m = this.mergeAt.get(`${s.ar},${s.ac}`)
        if (r1 === r2 && c1 === c2 && m) ({ r1, c1, r2, c2 } = m)
        const x = this.sx(c1); const y = this.sy(r1)
        return { x, y, w: this.X(c2 + 1) - this.X(c1), h: this.Y(r2 + 1) - this.Y(r1) }
      }

      updateStats () {
        const s = this.sel
        if (s.r1 === s.r2 && s.c1 === s.c2) { this.stats.textContent = ''; return }
        let n = 0; let count = 0; let sum = 0
        const cells = this.sheet.cells ?? {}
        const area = (s.r2 - s.r1 + 1) * (s.c2 - s.c1 + 1)
        if (area > 200000) {
          for (const [k, v] of Object.entries(cells)) { const [r, c] = k.split(',').map(Number); if (r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2) { if (v[0] !== '') count++; if (v[3] === 1) { const x = Number(String(v[1]).replace(/%$/, '')); if (!Number.isNaN(x)) { n++; sum += /%$/.test(v[1]) ? x / 100 : x } } } }
        } else {
          for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
            const v = cells[`${r},${c}`]
            if (!v) continue
            if (v[0] !== '') count++
            if (v[3] === 1) { const raw = String(v[1]); const x = Number(raw.replace(/%$/, '')); const val = v[1].startsWith('=') ? Number(String(v[0]).replace(/[,¥$%\s]/g, '')) : (/%$/.test(raw) ? x / 100 : x); if (!Number.isNaN(val)) { n++; sum += val } }
          }
        }
        const fmt = (x) => (Math.abs(x) >= 1e15 ? x.toExponential(4) : +x.toPrecision(12)).toLocaleString('zh-CN', { maximumFractionDigits: Math.abs(x) >= 100 ? 2 : 6 })
        this.stats.textContent = n ? `求和 ${fmt(sum)}　平均 ${fmt(sum / n)}　计数 ${count}` : count ? `计数 ${count}` : ''
      }

      // ------------------------------------------------ tabs

      renderTabs () {
        this.tabs.replaceChildren()
        this.model.sheets.forEach((s, i) => {
          if (s.state && s.state !== 'visible') return
          const t = el('button', { type: 'button', class: 'qx-tab' + (i === this.sheetIx ? ' on' : '') }, s.name)
          t.addEventListener('click', () => { if (this.editing) this.commit(0, 0); this.useSheet(i, false); this.focus() })
          t.addEventListener('dblclick', () => this.renameSheet(s.name))
          t.addEventListener('contextmenu', (e) => { e.preventDefault(); this.sheetMenu(e, s.name) })
          this.tabs.append(t)
        })
        if (this.doc.writable) {
          const add = el('button', { type: 'button', class: 'qx-tab qx-add', title: '新建工作表', html: icon('plus') })
          add.addEventListener('click', () => {
            const names = new Set(this.model.sheets.map((s) => s.name.toLowerCase()))
            let n = this.model.sheets.length + 1
            while (names.has(`sheet${n}`)) n++
            this.push([{ op: 'add_sheet', name: `Sheet${n}` }], { now: true, switchTo: `Sheet${n}` })
          })
          this.tabs.append(add)
        }
      }

      renameSheet (name) {
        if (!this.doc.writable) return
        const to = window.prompt('工作表名称', name)
        if (!to || to === name) return
        this.push([{ op: 'rename_sheet', sheet: name, name: to }], { now: true, switchTo: to })
      }

      sheetMenu (e, name) {
        if (!this.doc.writable) return
        const sheets = this.model.sheets
        const ix = sheets.findIndex((x) => x.name === name)
        const visible = sheets.filter((x) => !x.state || x.state === 'visible')
        const hidden = sheets.filter((x) => x.state && x.state !== 'visible')
        menu(e.clientX, e.clientY, [
          { label: '插入工作表', fn: () => this.tabs.querySelector('.qx-add')?.click() },
          { label: '重命名', fn: () => this.renameSheet(name) },
          { label: '删除', disabled: sheets.length < 2, fn: () => { if (window.confirm(`删除工作表“${name}”？`)) this.push([{ op: 'delete_sheet', sheet: name }], { now: true }) } },
          '-',
          { label: '左移', disabled: ix === 0, fn: () => this.push([{ op: 'move_sheet', sheet: name, to: ix - 1 }], { now: true, switchTo: name }) },
          { label: '右移', disabled: ix === sheets.length - 1, fn: () => this.push([{ op: 'move_sheet', sheet: name, to: ix + 1 }], { now: true, switchTo: name }) },
          '-',
          { label: '隐藏', disabled: visible.length < 2, fn: () => this.push([{ op: 'sheet_state', sheet: name, hidden: true }], { now: true, switchTo: visible.find((x) => x.name !== name)?.name }) },
          ...hidden.map((x) => ({ label: `取消隐藏“${x.name}”`, fn: () => this.push([{ op: 'sheet_state', sheet: x.name, hidden: false }], { now: true, switchTo: x.name }) }))
        ])
      }

      // ------------------------------------------------ selection

      select (r1, c1, r2, c2, ar = r1, ac = c1) {
        this.sel = { r1, c1, r2, c2, ar, ac }
        this.ensureVisible(ar, ac)
        this.draw()
        this.placeInput()
        this.refreshBar()
      }

      ensureVisible (r, c) {
        const s = this.scroll
        if (c > this.fc) {
          const left = this.X(c); const right = this.X(c + 1)
          const viewL = s.scrollLeft + this.FW(); const viewR = s.scrollLeft + this.vw - HW
          if (left < viewL) s.scrollLeft = left - this.FW()
          else if (right > viewR) s.scrollLeft = right - (this.vw - HW)
        }
        if (r > this.fr) {
          const top = this.Y(r); const bottom = this.Y(r + 1)
          const viewT = s.scrollTop + this.FH(); const viewB = s.scrollTop + this.vh - HH
          if (top < viewT) s.scrollTop = top - this.FH()
          else if (bottom > viewB) s.scrollTop = bottom - (this.vh - HH)
        }
      }

      refreshBar () {
        const s = this.sel
        // A save landing while the person types a reference must not undo it.
        if (document.activeElement !== this.nameBox) this.nameBox.value = s.r1 === s.r2 && s.c1 === s.c2 ? cellName(s.ar, s.ac) : `${cellName(s.r1, s.c1)}:${cellName(s.r2, s.c2)}`
        if (!this.editing) this.fbar.value = this.input(s.ar, s.ac)
        const st = this.cur()
        this.ui.b.classList.toggle('on', !!st.b)
        this.ui.i.classList.toggle('on', !!st.i)
        this.ui.u.classList.toggle('on', !!st.u)
        this.ui.al.classList.toggle('on', st.align === 'left')
        this.ui.ac.classList.toggle('on', st.align === 'center')
        this.ui.ar.classList.toggle('on', st.align === 'right')
        this.ui.wrap.classList.toggle('on', !!st.wrap)
        this.ui.lockCell?.classList.toggle('on', !st.unlocked)
        this.ui.protect.querySelector('span:last-child') && (this.ui.protect.lastChild.textContent = this.sheet.protect ? '撤销工作表保护' : '保护工作表')
        this.ui.protect.classList.toggle('on', !!this.sheet.protect)
        this.ui.gridPrint?.classList.toggle('on', !!this.sheet.page?.grid)
        this.ui.noteBtn.lastChild.textContent = this.sheet.notes?.[`${s.ar},${s.ac}`] ? '编辑批注' : '新建批注'
        this.showTip()
        this.ui.merge.classList.toggle('on', this.mergeAt.has(`${s.ar},${s.ac}`))
        this.ui.filter.classList.toggle('on', !!this.sheet.autoFilter)
        this.ui.size.value = st.size ? String(+st.size) : ''
        const fmt = st.fmt ?? 'General'
        if (![...this.ui.fmt.options].some((o) => o.value === fmt)) this.ui.fmt.append(el('option', { value: fmt }, fmt))
        this.ui.fmt.value = fmt
        const ro = !this.doc.writable
        for (const k of ['b', 'i', 'u', 'al', 'ac', 'ar', 'wrap', 'merge', 'decInc', 'decDec', 'sortUp', 'sortDown', 'filter']) this.ui[k].disabled = ro
        this.ui.size.disabled = this.ui.fmt.disabled = this.ui.freeze.disabled = this.ui.border.disabled = ro
        this.fbar.readOnly = ro
      }

      hit (e) {
        const gr = this.grid.getBoundingClientRect()
        const x = e.clientX - gr.left; const y = e.clientY - gr.top
        if (x < HW && y < HH) return { kind: 'corner', x, y }
        if (y < HH) {
          const c = this.colAt(x)
          const edge = this.sx(c) + this.w(c)
          if (Math.abs(x - edge) <= 4) return { kind: 'colEdge', c, x, y }
          if (Math.abs(x - this.sx(c)) <= 4 && c > 1) return { kind: 'colEdge', c: c - 1, x, y }
          return { kind: 'col', c, x, y }
        }
        if (x < HW) {
          const rr = this.rowAt(y)
          const edge = this.sy(rr) + this.h(rr)
          if (Math.abs(y - edge) <= 3) return { kind: 'rowEdge', r: rr, x, y }
          if (Math.abs(y - this.sy(rr)) <= 3 && rr > 1) return { kind: 'rowEdge', r: rr - 1, x, y }
          return { kind: 'row', r: rr, x, y }
        }
        if (this.doc.writable) {
          const b = this.selBox()
          if (Math.abs(x - (b.x + b.w)) <= 5 && Math.abs(y - (b.y + b.h)) <= 5) return { kind: 'fill', x, y }
        }
        const d = this.dvBtn
        if (d && x >= d.x && x <= d.x + d.w && y >= d.y && y <= d.y + d.h) return { kind: 'dvBtn', x, y }
        const r = this.rowAt(y); const c = this.colAt(x)
        if (this.filterHead(r, c) && x > this.sx(c) + this.w(c) - 19) return { kind: 'filterBtn', r, c, x, y }
        return { kind: 'cell', r, c, x, y }
      }

      hover (e) {
        if (this.dragging) return
        const h = this.hit(e)
        this.scroll.style.cursor = h.kind === 'colEdge' ? 'col-resize' : h.kind === 'rowEdge' ? 'row-resize' : h.kind === 'cell' ? (this.painter ? 'copy' : 'cell') : h.kind === 'fill' ? 'crosshair' : 'default'
        const key = h.kind === 'cell' ? `${h.r},${h.c}` : ''
        const note = key && !this.showNotes && this.sheet.notes?.[key]
        const link = key && this.sheet.links?.[key]
        if (!note && !link) { this.pop.hidden = true; return }
        if (this.popKey === key && !this.pop.hidden) return
        this.popKey = key
        this.pop.replaceChildren(...[
          note && el('div', { class: 'qx-pop-note' }, note.a ? el('b', {}, note.a + '：') : null, note.t),
          link && el('div', { class: 'qx-pop-link' }, link, el('small', {}, link.startsWith('#') ? '⌘+单击跳转' : '⌘+单击打开'))
        ].filter(Boolean))
        Object.assign(this.pop.style, { left: Math.min(this.vw - 220, this.sx(h.c) + this.w(h.c) + 4) + 'px', top: Math.max(HH, this.sy(h.r)) + 'px' })
        this.pop.hidden = false
      }

      placeNotes () {
        if (!this.noteLayer) return
        const notes = this.showNotes ? Object.entries(this.sheet.notes ?? {}) : []
        const kids = []
        for (const [key, n] of notes) {
          const [r, c] = key.split(',').map(Number)
          const x = this.sx(c) + this.w(c) + 6; const y = this.sy(r)
          if (x > this.vw || y > this.vh || y < HH - 40 || x < HW) continue
          kids.push(el('div', { class: 'qx-note', style: { left: x - HW + 'px', top: y - HH + 'px' } }, n.a ? el('b', {}, n.a + '：') : null, n.t))
        }
        this.noteLayer.replaceChildren(...kids)
      }

      down (e) {
        if (e.button !== 0) return
        if (this.selChart !== undefined) this.pickChart(undefined)
        const h = this.hit(e)
        if (this.editing && h.kind === 'cell' && this.pointable()) { e.preventDefault(); this.pointRef(h); return }
        if (this.editing) this.commit(0, 0)
        const s = this.sel
        if (h.kind === 'dvBtn') { e.preventDefault(); this.dvList(); return }
        const link = h.kind === 'cell' && mod(e) && this.sheet.links?.[`${h.r},${h.c}`]
        if (link) { e.preventDefault(); this.followLink(link); return }
        if (h.kind === 'colEdge' || h.kind === 'rowEdge') {
          if (!this.doc.writable) return
          e.preventDefault()
          const col = h.kind === 'colEdge'
          const start = col ? this.w(h.c) : this.h(h.r)
          const origin = col ? e.clientX : e.clientY
          const base = col ? this.sx(h.c) : this.sy(h.r)
          this.guide.hidden = false
          this.guide.className = 'qx-guide ' + (col ? 'v' : 'h')
          const move = (ev) => {
            const size = Math.max(col ? 12 : 8, start + (col ? ev.clientX : ev.clientY) - origin)
            if (col) this.guide.style.left = base + size + 'px'; else this.guide.style.top = base + size + 'px'
            this.dragging = size
          }
          move(e)
          const up = () => {
            window.removeEventListener('pointermove', move)
            this.guide.hidden = true
            const size = Math.round(this.dragging)
            this.dragging = undefined
            const px = Math.round(size / this.z)
            if (col) {
              const [a, b] = s.c1 <= h.c && h.c <= s.c2 && s.r1 === 1 && s.r2 >= this.nr ? [s.c1, s.c2] : [h.c, h.c]
              this.setWidths(a, b, px)
            } else {
              const [a, b] = s.r1 <= h.r && h.r <= s.r2 && s.c1 === 1 && s.c2 >= this.nc ? [s.r1, s.r2] : [h.r, h.r]
              this.setHeights(a, b, px)
            }
            this.layout(); this.draw()
          }
          window.addEventListener('pointermove', move)
          window.addEventListener('pointerup', up, { once: true })
          return
        }
        e.preventDefault()
        this.focus()
        if (h.kind === 'fill') { this.fillDrag(e); return }
        if (h.kind === 'filterBtn') { this.select(h.r, h.c, h.r, h.c); this.filterMenu(h.c, e.clientX, e.clientY); return }
        const extend = e.shiftKey
        if (h.kind === 'corner') { this.select(1, 1, this.nr, this.nc, 1, 1); return }
        let anchor
        if (h.kind === 'col') { anchor = extend ? { r: 1, c: s.ac } : { r: 1, c: h.c }; this.select(1, Math.min(anchor.c, h.c), this.nr, Math.max(anchor.c, h.c), 1, anchor.c) }
        if (h.kind === 'row') { anchor = extend ? { r: s.ar, c: 1 } : { r: h.r, c: 1 }; this.select(Math.min(anchor.r, h.r), 1, Math.max(anchor.r, h.r), this.nc, anchor.r, 1) }
        if (h.kind === 'cell') { anchor = extend ? { r: s.ar, c: s.ac } : { r: h.r, c: h.c }; this.selectRange(anchor, h) }
        const kind = h.kind
        const move = (ev) => {
          const g = this.hit(ev)
          if (kind === 'cell') this.selectRange(anchor, { r: this.rowAt(g.y), c: this.colAt(g.x) })
          else if (kind === 'col') { const c = this.colAt(g.x); this.select(1, Math.min(anchor.c, c), this.nr, Math.max(anchor.c, c), 1, anchor.c) } else if (kind === 'row') { const r = this.rowAt(g.y); this.select(Math.min(anchor.r, r), 1, Math.max(anchor.r, r), this.nc, anchor.r, 1) }
        }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', () => { window.removeEventListener('pointermove', move); if (this.painter && kind === 'cell') this.paintSel() }, { once: true })
      }

      // selectRange selects from anchor to the cell, widened to whole merges.
      selectRange (a, b) {
        let r1 = Math.min(a.r, b.r); let r2 = Math.max(a.r, b.r); let c1 = Math.min(a.c, b.c); let c2 = Math.max(a.c, b.c)
        for (let grew = true; grew;) {
          grew = false
          for (const m of this.merges) {
            if (m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2) continue
            if (m.r1 < r1) { r1 = m.r1; grew = true } if (m.r2 > r2) { r2 = m.r2; grew = true }
            if (m.c1 < c1) { c1 = m.c1; grew = true } if (m.c2 > c2) { c2 = m.c2; grew = true }
          }
        }
        const am = this.mergeAt.get(`${a.r},${a.c}`)
        this.select(r1, c1, r2, c2, am ? am.r1 : a.r, am ? am.c1 : a.c)
      }

      focus () { if (!this.editing) this.inp.focus({ preventScroll: true }) }

      move (dr, dc, extend) {
        const s = this.sel
        if (extend) {
          // Move the far corner; the active cell stays.
          let er = s.r1 === s.ar ? s.r2 : s.r1
          let ec = s.c1 === s.ac ? s.c2 : s.c1
          er = Math.max(1, Math.min(this.nr, er + dr)); ec = Math.max(1, Math.min(this.nc, ec + dc))
          this.selectRange({ r: s.ar, c: s.ac }, { r: er, c: ec })
          this.ensureVisible(er, ec)
          this.draw()
          return
        }
        let r = s.ar; let c = s.ac
        const m = this.mergeAt.get(`${r},${c}`)
        if (m && dr > 0) r = m.r2
        if (m && dc > 0) c = m.c2
        do { r += dr } while (dr && r > 1 && r < this.nr && this.h(r) === 0)
        do { c += dc } while (dc && c > 1 && c < this.nc && this.w(c) === 0)
        r = Math.max(1, Math.min(this.nr, r)); c = Math.max(1, Math.min(this.nc, c))
        const t = this.mergeAt.get(`${r},${c}`)
        if (t) { r = t.r1; c = t.c1 }
        this.select(r, c, t ? t.r2 : r, t ? t.c2 : c, r, c)
      }

      // jump moves to the edge of the data block, like ⌘ + arrow.
      jump (dr, dc, extend) {
        const s = this.sel
        let r = s.ar; let c = s.ac
        const filled = (rr, cc) => this.shown(rr, cc) !== ''
        const lim = (rr, cc) => rr >= 1 && cc >= 1 && rr <= Math.max(this.sheet.maxR + 1, 1) && cc <= Math.max(this.sheet.maxC + 1, 1)
        if (filled(r, c) && lim(r + dr, c + dc) && filled(r + dr, c + dc)) { while (lim(r + dr, c + dc) && filled(r + dr, c + dc)) { r += dr; c += dc } } else {
          r += dr; c += dc
          while (lim(r, c) && !filled(r, c)) { r += dr; c += dc }
          if (!lim(r, c)) { r = dr > 0 ? Math.max(this.sheet.maxR, s.ar) : dr < 0 ? 1 : s.ar; c = dc > 0 ? Math.max(this.sheet.maxC, s.ac) : dc < 0 ? 1 : s.ac }
        }
        r = Math.max(1, r); c = Math.max(1, c)
        if (extend) this.selectRange({ r: s.ar, c: s.ac }, { r, c }); else this.select(r, c, r, c, r, c)
      }

      // ------------------------------------------------ keyboard

      key (e) {
        if (e.isComposing || e.keyCode === 229) return
        const s = this.sel
        if (this.editing && this.acKey(e, this.inp)) return
        if (this.editing && e.altKey && e.code === 'Equal') { e.preventDefault(); this.autoSum('SUM'); return }
        if (this.editing && e.key === 'F3' && e.shiftKey) { e.preventDefault(); this.insertFunction(); return }
        if (this.editing) {
          const enterMode = this.editing.mode === 'enter'
          if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); this.commit(e.shiftKey ? -1 : 1, 0); return }
          if (e.key === 'Enter' && e.altKey) { e.preventDefault(); this.insertText('\n'); return }
          if (e.key === 'Tab') { e.preventDefault(); this.commit(0, e.shiftKey ? -1 : 1); return }
          if (e.key === 'Escape') { e.preventDefault(); this.cancel(); return }
          if (enterMode && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            e.preventDefault()
            this.commit(e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0, e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0)
          }
          return
        }
        const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }
        if (arrows[e.key]) { e.preventDefault(); const [dr, dc] = arrows[e.key]; if (mod(e)) this.jump(dr, dc, e.shiftKey); else this.move(dr, dc, e.shiftKey); return }
        if (e.key === 'Enter') { e.preventDefault(); this.move(e.shiftKey ? -1 : 1, 0); return }
        if (e.key === 'Tab') { e.preventDefault(); this.move(0, e.shiftKey ? -1 : 1); return }
        if (e.key === 'Home') { e.preventDefault(); this.select(mod(e) ? 1 : s.ar, 1, mod(e) ? 1 : s.ar, 1); return }
        if (e.key === 'PageDown' || e.key === 'PageUp') { e.preventDefault(); const n = Math.max(1, Math.floor((this.vh - HH) / (this.sheet.defH ?? 20)) - 1); this.move(e.key === 'PageDown' ? n : -n, 0, e.shiftKey); return }
        if (e.key === 'F2') { e.preventDefault(); this.beginEdit('edit'); return }
        if (e.key === 'F3' && e.shiftKey) { e.preventDefault(); this.insertFunction(); return }
        if (e.key === 'F2' && e.shiftKey) { e.preventDefault(); this.editNote(); return }
        if (e.key === 'Escape' && this.painter) { e.preventDefault(); this.stopPainter(); return }
        if (e.altKey && e.shiftKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) { e.preventDefault(); this.outline(e.key === 'ArrowRight' ? 1 : -1); return }
        if (mod(e) && e.key.toLowerCase() === 'k') { e.preventDefault(); this.editLink(); return }
        if (e.altKey && e.key === 'ArrowDown') { e.preventDefault(); this.dvList(); return }
        if (e.altKey && e.code === 'Equal') { e.preventDefault(); this.autoSum('SUM'); return }
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.clearSel(); return }
        if (mod(e)) {
          const k = e.key.toLowerCase()
          if (k === 'a') { e.preventDefault(); this.select(1, 1, this.nr, this.nc, s.ar, s.ac); return }
          if (k === 'b') { e.preventDefault(); if (this.doc.writable) this.styleSel({ bold: !this.cur().b }); return }
          if (k === 'i') { e.preventDefault(); if (this.doc.writable) this.styleSel({ italic: !this.cur().i }); return }
          if (k === 'u') { e.preventDefault(); if (this.doc.writable) this.styleSel({ underline: !this.cur().u }); return }
          if (k === 'z') { e.preventDefault(); this.doc.step(e.shiftKey ? 'redo' : 'undo'); return }
          if (k === 'y') { e.preventDefault(); this.doc.step('redo'); return }
          if (k === 'c' || k === 'x') { this.inp.value = this.tsv(); this.inp.select(); return }
          if (k === 'v') { this.inp.value = ''; return }
          if (k === 'd' && this.doc.writable) { e.preventDefault(); this.fillDown(); return }
        }
        if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) { e.preventDefault(); const b = this.selBox(); const g = this.grid.getBoundingClientRect(); this.cellMenu(g.left + b.x + b.w / 2, g.top + b.y + b.h / 2) }
      }

      // ------------------------------------------------ editing

      beginEdit (mode, typed) {
        if (!this.doc.writable || this.editing) return
        if (this.sheet.protect && !this.cur().unlocked) { this.frame.toast('工作表受保护，这个单元格是锁定的。到“审阅”里撤销保护后再改。', 'warn'); this.inp.value = ''; return }
        const s = this.sel
        this.editing = { mode, orig: this.input(s.ar, s.ac) }
        if (!typed) this.inp.value = this.editing.orig
        this.inp.classList.add('on')
        this.placeInput()
        this.inp.focus({ preventScroll: true })
        if (!typed) { const n = this.inp.value.length; this.inp.setSelectionRange(n, n) }
        this.fbar.value = this.inp.value
      }

      insertText (t) {
        const a = this.inp.selectionStart; const b = this.inp.selectionEnd
        this.inp.value = this.inp.value.slice(0, a) + t + this.inp.value.slice(b)
        this.inp.setSelectionRange(a + t.length, a + t.length)
        this.fbar.value = this.inp.value
        this.fitInput()
      }

      placeInput () {
        if (!this.sheet) return
        const s = this.sel
        const m = this.mergeAt.get(`${s.ar},${s.ac}`)
        const r1 = m ? m.r1 : s.ar; const c1 = m ? m.c1 : s.ac
        const x = this.sx(c1); const y = this.sy(r1)
        const w = m ? this.X(m.c2 + 1) - this.X(m.c1) : this.w(c1); const h = m ? this.Y(m.r2 + 1) - this.Y(m.r1) : this.h(r1)
        Object.assign(this.inp.style, { left: Math.max(HW, x) + 'px', top: Math.max(HH, y) + 'px' })
        if (this.editing) {
          const st = this.cur()
          Object.assign(this.inp.style, { minWidth: Math.max(w, 40) + 'px', minHeight: Math.max(h, 18) + 'px', font: this.font(st), color: st.color ? '#' + st.color : '#000', background: st.fill ? '#' + st.fill : '#fff', textAlign: st.align && st.align !== 'general' ? st.align : 'left' })
          this.fitInput()
        } else Object.assign(this.inp.style, { minWidth: '', minHeight: '', width: '1px', height: '1px' })
      }

      fitInput () {
        if (!this.editing) return
        this.inp.style.width = '1px'
        this.inp.style.height = '1px'
        const maxW = this.vw - parseFloat(this.inp.style.left) - 4
        // Measure unwrapped, so text grows to the right before it wraps.
        this.inp.style.whiteSpace = 'pre'
        const want = this.inp.scrollWidth + 8
        this.inp.style.whiteSpace = ''
        this.inp.style.width = Math.min(maxW, want) + 'px'
        this.inp.style.height = this.inp.scrollHeight + 'px'
      }

      commit (dr, dc) {
        const ed = this.editing
        if (!ed) return
        const text = ed.bar ? this.fbar.value : this.inp.value
        const bad = text !== ed.orig && this.invalid(this.sel.ar, this.sel.ac, text)
        if (bad?.stop) {
          this.frame.toast(bad.msg, 'warn')
          ;(ed.bar ? this.fbar : this.inp).focus({ preventScroll: true })
          return
        }
        if (bad) this.frame.toast(bad.msg, 'info')
        this.pointing = null
        this.editing = null
        this.inp.classList.remove('on')
        this.inp.value = ''
        this.hideAc()
        const s = this.sel
        if (text !== ed.orig) this.push([{ op: 'input', sheet: this.sheet.name, cell: cellName(s.ar, s.ac), text }], { overlay: [[s.ar, s.ac, text]] })
        if (dr || dc) this.move(dr, dc); else { this.placeInput(); this.draw(); this.refreshBar() }
        this.inp.focus({ preventScroll: true })
      }

      cancel () {
        this.hideAc()
        this.pointing = null
        this.editing = null
        this.inp.classList.remove('on')
        this.inp.value = ''
        this.placeInput()
        this.refreshBar()
        this.inp.focus({ preventScroll: true })
      }

      clearSel () {
        if (!this.doc.writable) return
        const s = this.sel
        const r2 = Math.min(s.r2, Math.max(this.sheet.maxR, s.r1)); const c2 = Math.min(s.c2, Math.max(this.sheet.maxC, s.c1))
        const over = []
        if ((r2 - s.r1 + 1) * (c2 - s.c1 + 1) <= 20000) for (let r = s.r1; r <= r2; r++) for (let c = s.c1; c <= c2; c++) if (this.shown(r, c) !== '') over.push([r, c, ''])
        this.push([{ op: 'clear', sheet: this.sheet.name, range: `${cellName(s.r1, s.c1)}:${cellName(r2, c2)}` }], { overlay: over })
      }

      tsv () {
        const s = this.sel
        const r2 = Math.min(s.r2, Math.max(this.sheet.maxR, s.r1)); const c2 = Math.min(s.c2, Math.max(this.sheet.maxC, s.c1))
        const rows = []
        for (let r = s.r1; r <= r2; r++) {
          const row = []
          for (let c = s.c1; c <= c2; c++) { const v = this.shown(r, c); row.push(/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v) }
          rows.push(row.join('\t'))
        }
        return rows.join('\n')
      }

      copy (e, cut) {
        if (this.editing) return
        e.preventDefault()
        e.clipboardData.setData('text/plain', this.tsv())
        this.inp.value = ''
        if (cut) this.clearSel()
      }

      paste (e) {
        if (this.editing) return
        e.preventDefault()
        if (!this.doc.writable) return
        // A table copied from a web page or document comes as HTML.
        const html = e.clipboardData.getData('text/html')
        if (html && /<table/i.test(html)) {
          const t = new DOMParser().parseFromString(html, 'text/html').querySelector('table')
          const rows = [...(t?.rows ?? [])].map((tr) => [...tr.cells].map((td) => td.innerText.replace(/\r?\n$/, '').trim()))
          if (rows.length) { this.pasteRows(rows); return }
        }
        this.pasteText(e.clipboardData.getData('text/plain'))
      }

      pasteText (text) {
        if (!text || !this.doc.writable) return
        // Tab-separated rows, with quoted fields that may hold tabs and
        // line breaks.
        const rows = [[]]
        let field = ''; let q = false
        for (let i = 0; i < text.length; i++) {
          const ch = text[i]
          if (q) { if (ch === '"' && text[i + 1] === '"') { field += '"'; i++ } else if (ch === '"') q = false; else field += ch; continue }
          if (ch === '"' && field === '') q = true
          else if (ch === '\t') { rows[rows.length - 1].push(field); field = '' } else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; rows[rows.length - 1].push(field); field = ''; rows.push([]) } else field += ch
        }
        rows[rows.length - 1].push(field)
        if (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop()
        this.pasteRows(rows)
      }

      pasteRows (rows) {
        const s = this.sel
        const over = []
        rows.forEach((row, i) => row.forEach((v, j) => over.push([s.ar + i, s.ac + j, v])))
        this.push([{ op: 'inputs', sheet: this.sheet.name, at: cellName(s.ar, s.ac), rows }], { overlay: over.length <= 20000 ? over : [] })
        this.select(s.ar, s.ac, s.ar + rows.length - 1, s.ac + Math.max(...rows.map((r) => r.length)) - 1, s.ar, s.ac)
      }

      styleSel (props) {
        const s = this.sel
        const r2 = Math.min(s.r2, Math.max(this.sheet.maxR + 50, s.r1)); const c2 = Math.min(s.c2, Math.max(this.sheet.maxC + 10, s.c1))
        if ((r2 - s.r1 + 1) * (c2 - s.c1 + 1) > 200000) { this.frame.toast('选区太大，请缩小后再设置格式。', 'warn'); return }
        this.push([{ op: 'style', sheet: this.sheet.name, range: `${cellName(s.r1, s.c1)}:${cellName(r2, c2)}`, ...props }], { now: true })
      }

      mergeSel () {
        if (!this.doc.writable) return
        const s = this.sel
        const m = this.mergeAt.get(`${s.ar},${s.ac}`)
        if (m && s.r1 === m.r1 && s.r2 === m.r2 && s.c1 === m.c1 && s.c2 === m.c2) this.push([{ op: 'unmerge', sheet: this.sheet.name, range: `${cellName(m.r1, m.c1)}:${cellName(m.r2, m.c2)}` }], { now: true })
        else if (s.r1 !== s.r2 || s.c1 !== s.c2) {
          const r2 = Math.min(s.r2, this.sheet.maxR + 200); const c2 = Math.min(s.c2, this.sheet.maxC + 50)
          this.push([{ op: 'merge', sheet: this.sheet.name, range: `${cellName(s.r1, s.c1)}:${cellName(r2, c2)}` }], { now: true })
        }
      }

      freeze (v) {
        if (!this.doc.writable || !v) return
        const s = this.sel
        const cell = v === 'row' ? 'A2' : v === 'col' ? 'B1' : v === 'none' ? 'A1' : cellName(Math.max(1, s.ar), Math.max(1, s.ac))
        this.push([{ op: 'freeze', sheet: this.sheet.name, cell }], { now: true })
      }

      lines (op) {
        if (!this.doc.writable) return
        const s = this.sel
        const rows = op.endsWith('rows')
        const at = rows ? s.r1 : s.c1
        const count = rows ? Math.min(s.r2, Math.max(this.sheet.maxR, s.r1)) - s.r1 + 1 : Math.min(s.c2, Math.max(this.sheet.maxC, s.c1)) - s.c1 + 1
        this.push([{ op, sheet: this.sheet.name, at, count: Math.max(1, count) }], { now: true })
      }

      // ------------------------------------------------ rows, columns, fill

      setWidths (a, b, px) {
        for (let c = a; c <= b; c++) this.colMap.set(c, px)
        this.push([{ op: 'width', sheet: this.sheet.name, col1: a, col2: b, px }])
        this.layout(); this.draw()
      }

      setHeights (a, b, px) {
        this.sheet.rows = { ...this.sheet.rows }
        for (let r = a; r <= b; r++) this.sheet.rows[r] = px
        this.push([{ op: 'height', sheet: this.sheet.name, row1: a, row2: b, px }])
        this.layout(); this.draw()
      }

      // dataRange is the selection clipped to where the sheet has data.
      dataRange () {
        const s = this.sel
        return { r1: s.r1, c1: s.c1, r2: Math.min(s.r2, Math.max(this.sheet.maxR, s.r1)), c2: Math.min(s.c2, Math.max(this.sheet.maxC, s.c1)) }
      }

      // region is the block of filled cells around a cell, like Excel's
      // current region.
      region (r, c) {
        const filled = (rr, cc) => rr >= 1 && cc >= 1 && this.shown(rr, cc) !== ''
        let r1 = r; let r2 = r; let c1 = c; let c2 = c
        for (let grew = true, n = 0; grew && n < 2000; n++) {
          grew = false
          const rowHas = (rr) => { for (let cc = Math.max(1, c1 - 1); cc <= c2 + 1; cc++) if (filled(rr, cc)) return true; return false }
          const colHas = (cc) => { for (let rr = Math.max(1, r1 - 1); rr <= r2 + 1; rr++) if (filled(rr, cc)) return true; return false }
          if (r1 > 1 && rowHas(r1 - 1)) { r1--; grew = true }
          if (r2 < this.sheet.maxR && rowHas(r2 + 1)) { r2++; grew = true }
          if (c1 > 1 && colHas(c1 - 1)) { c1--; grew = true }
          if (c2 < this.sheet.maxC && colHas(c2 + 1)) { c2++; grew = true }
        }
        return { r1, c1, r2, c2 }
      }

      target () {
        const s = this.sel
        return s.r1 === s.r2 && s.c1 === s.c2 ? this.region(s.ar, s.ac) : this.dataRange()
      }

      fillDrag (e) {
        const s = this.dataRange()
        const move = (ev) => {
          const g = this.hit(ev)
          const r = this.rowAt(g.y); const c = this.colAt(g.x)
          const dr = r > s.r2 ? r - s.r2 : r < s.r1 ? r - s.r1 : 0
          const dc = c > s.c2 ? c - s.c2 : c < s.c1 ? c - s.c1 : 0
          let f
          if (!dr && !dc) f = undefined
          else if (Math.abs(dr) * this.h(r) >= Math.abs(dc) * this.w(c)) f = dr > 0 ? { ...s, r2: r } : { ...s, r1: r }
          else f = dc > 0 ? { ...s, c2: c } : { ...s, c1: c }
          this.fillTo = f
          this.draw()
        }
        const up = () => {
          window.removeEventListener('pointermove', move)
          const f = this.fillTo
          this.fillTo = undefined
          if (!f) { this.draw(); return }
          const range = (x) => `${cellName(x.r1, x.c1)}:${cellName(x.r2, x.c2)}`
          this.push([{ op: 'autofill', sheet: this.sheet.name, from: range(s), to: range(f) }], { now: true })
          this.select(f.r1, f.c1, f.r2, f.c2, this.sel.ar, this.sel.ac)
        }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', up, { once: true })
      }

      // fillDown copies the selection's first row down it (⌘D).
      fillDown () {
        const s = this.dataRange()
        if (s.r2 === s.r1) return
        this.push([{ op: 'autofill', sheet: this.sheet.name, from: `${cellName(s.r1, s.c1)}:${cellName(s.r1, s.c2)}`, to: `${cellName(s.r1, s.c1)}:${cellName(s.r2, s.c2)}` }], { now: true })
      }

      border (kind) {
        if (!this.doc.writable) return
        const s = this.dataRange()
        const r2 = Math.min(this.sel.r2, Math.max(this.sheet.maxR + 50, s.r1)); const c2 = Math.min(this.sel.c2, Math.max(this.sheet.maxC + 10, s.c1))
        if ((r2 - s.r1 + 1) * (c2 - s.c1 + 1) > 200000) { this.frame.toast('选区太大，请缩小后再设置边框。', 'warn'); return }
        this.push([{ op: 'border', sheet: this.sheet.name, range: `${cellName(s.r1, s.c1)}:${cellName(r2, c2)}`, kind }], { now: true })
      }

      decimals (d) {
        if (!this.doc.writable) return
        const st = this.cur()
        let f = st.fmt && st.fmt !== 'General' ? st.fmt : null
        if (!f) { const v = String(this.input(this.sel.ar, this.sel.ac)); const dec = /^-?\d*\.(\d+)$/.exec(v)?.[1]?.length ?? 0; f = dec ? '0.' + '0'.repeat(dec) : '0' }
        if (!/0/.test(f)) return
        const next = f.replace(/(0+)(\.(0*))?/, (all, ints, dot, decs) => { const n = Math.max(0, (decs?.length ?? 0) + d); return ints + (n ? '.' + '0'.repeat(n) : '') })
        this.styleSel({ format: next })
      }

      // ------------------------------------------------ sort and filter

      sortSel (desc, col) {
        if (!this.doc.writable) return
        const af = this.sheet.autoFilter
        let rg = this.target()
        let header
        if (af && this.sel.r1 === this.sel.r2 && this.sel.c1 === this.sel.c2 && this.sel.ar >= af.r1 && this.sel.ar <= af.r2 && this.sel.ac >= af.c1 && this.sel.ac <= af.c2) { rg = af; header = true }
        if (header === undefined) {
          const kinds = (r) => { const out = []; for (let c = rg.c1; c <= rg.c2; c++) out.push(this.cell(r, c)?.[3] ?? (this.shown(r, c) === '' ? -1 : 0)); return out }
          const first = kinds(rg.r1); const second = rg.r2 > rg.r1 ? kinds(rg.r1 + 1) : []
          header = first.every((k) => k === 0) && second.some((k) => k !== 0 && k !== -1)
        }
        col = col ?? Math.max(rg.c1, Math.min(rg.c2, this.sel.ac))
        if (rg.r2 - rg.r1 < (header ? 2 : 1)) return
        this.push([{ op: 'sort', sheet: this.sheet.name, range: `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`, col, desc, header }], { now: true })
      }

      toggleFilter () {
        if (!this.doc.writable) return
        if (this.sheet.autoFilter) { this.push([{ op: 'filter_off', sheet: this.sheet.name }], { now: true }); return }
        const rg = this.target()
        if (rg.r2 <= rg.r1 && this.shown(rg.r1, rg.c1) === '') { this.frame.toast('请先选中带标题行的数据区域。', 'info'); return }
        this.push([{ op: 'filter', sheet: this.sheet.name, range: `${cellName(rg.r1, rg.c1)}:${cellName(Math.max(rg.r2, rg.r1 + 1), rg.c2)}` }], { now: true })
      }

      filterMenu (c, x, y) {
        const af = this.sheet.autoFilter
        if (!af) return
        document.querySelector('.q-menu')?.remove()
        const values = new Map()
        for (let r = af.r1 + 1; r <= af.r2; r++) { const v = this.shown(r, c); values.set(v, (values.get(v) ?? 0) + 1) }
        const list = [...values.keys()].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : a.localeCompare(b, 'zh-CN', { numeric: true })))
        const kept = af.cols?.[c]
        const panel = el('div', { class: 'q-menu q-filter' })
        const ro = !this.doc.writable
        const act = (label, fn) => { const b = el('button', { type: 'button', disabled: ro }, label); b.addEventListener('click', () => { panel.remove(); fn() }); return b }
        panel.append(act('升序排序', () => this.push([{ op: 'sort', sheet: this.sheet.name, range: `${cellName(af.r1, af.c1)}:${cellName(af.r2, af.c2)}`, col: c, desc: false, header: true }], { now: true })))
        panel.append(act('降序排序', () => this.push([{ op: 'sort', sheet: this.sheet.name, range: `${cellName(af.r1, af.c1)}:${cellName(af.r2, af.c2)}`, col: c, desc: true, header: true }], { now: true })))
        panel.append(el('div', { class: 'q-menu-sep' }))
        const search = el('input', { class: 'q-filter-q', placeholder: '搜索', spellcheck: 'false' })
        const box = el('div', { class: 'q-filter-list' })
        const all = el('input', { type: 'checkbox' })
        const checks = list.map((v) => {
          const cb = el('input', { type: 'checkbox' })
          cb.checked = !Array.isArray(kept) || kept.includes(v)
          const row = el('label', { class: 'q-filter-item' }, cb, el('span', {}, v === '' ? '（空白）' : v), el('em', {}, String(values.get(v))))
          row.dataset.v = v
          box.append(row)
          return [v, cb, row]
        })
        const sync = () => { all.checked = checks.every(([, cb]) => cb.checked); all.indeterminate = !all.checked && checks.some(([, cb]) => cb.checked) }
        all.addEventListener('change', () => { for (const [, cb, row] of checks) if (!row.hidden) cb.checked = all.checked; sync() })
        for (const [, cb] of checks) cb.addEventListener('change', sync)
        search.addEventListener('input', () => { const q = search.value.toLowerCase(); for (const [v, , row] of checks) row.hidden = !!q && !v.toLowerCase().includes(q) })
        sync()
        panel.append(search, el('label', { class: 'q-filter-item q-filter-all' }, all, el('span', {}, '全选')), box)
        const ok = el('button', { type: 'button', class: 'q-filter-ok', disabled: ro }, '确定')
        const cancel = el('button', { type: 'button', class: 'q-filter-cancel' }, '取消')
        ok.addEventListener('click', () => {
          panel.remove()
          const chosen = checks.filter(([, cb]) => cb.checked).map(([v]) => v)
          this.push([{ op: 'filter', sheet: this.sheet.name, col: c, values: chosen.length === checks.length ? null : chosen }], { now: true })
        })
        cancel.addEventListener('click', () => panel.remove())
        panel.append(el('div', { class: 'q-filter-btns' }, cancel, ok))
        document.body.append(panel)
        const r = panel.getBoundingClientRect()
        panel.style.left = Math.max(4, Math.min(x, window.innerWidth - r.width - 4)) + 'px'
        panel.style.top = Math.max(4, Math.min(y + 8, window.innerHeight - r.height - 4)) + 'px'
        search.focus()
        const close = (ev) => { if (!panel.contains(ev.target)) { panel.remove(); document.removeEventListener('pointerdown', close, true) } }
        setTimeout(() => document.addEventListener('pointerdown', close, true))
        panel.addEventListener('keydown', (ev) => { if (ev.key === 'Escape') { ev.stopPropagation(); panel.remove(); this.focus() } else if (ev.key === 'Enter') { ev.preventDefault(); ok.click() } })
      }

      // ------------------------------------------------ context menus

      context (e) {
        e.preventDefault()
        const h = this.hit(e)
        const s = this.sel
        if (this.editing) this.commit(0, 0)
        if (h.kind === 'col' || h.kind === 'colEdge') {
          const c = h.c
          if (!(c >= s.c1 && c <= s.c2 && s.r1 === 1 && s.r2 >= this.nr)) this.select(1, c, this.nr, c, 1, c)
          this.headerMenu('c', e.clientX, e.clientY)
        } else if (h.kind === 'row' || h.kind === 'rowEdge') {
          const r = h.r
          if (!(r >= s.r1 && r <= s.r2 && s.c1 === 1 && s.c2 >= this.nc)) this.select(r, 1, r, this.nc, r, 1)
          this.headerMenu('r', e.clientX, e.clientY)
        } else if (h.kind === 'cell' || h.kind === 'filterBtn') {
          if (!(h.r >= s.r1 && h.r <= s.r2 && h.c >= s.c1 && h.c <= s.c2)) this.selectRange({ r: h.r, c: h.c }, { r: h.r, c: h.c })
          this.cellMenu(e.clientX, e.clientY)
        }
      }

      async clip (cut) {
        try { await navigator.clipboard.writeText(this.tsv()) } catch { this.frame.toast(`请用 ⌘${cut ? 'X' : 'C'} ${cut ? '剪切' : '复制'}。`, 'info'); return }
        if (cut) this.clearSel()
      }

      async pasteFromMenu () {
        try { this.pasteText(await navigator.clipboard.readText()) } catch { this.frame.toast('浏览器不允许从菜单读取剪贴板，请用 ⌘V 粘贴。', 'info') }
      }

      cellMenu (x, y) {
        const ro = !this.doc.writable
        const s = this.sel
        const merged = this.mergeAt.has(`${s.ar},${s.ac}`)
        menu(x, y, [
          { label: '标注问题…', kbd: '⇧⌘M', fn: () => this.frame.pinHere() },
          '-',
          { label: '剪切', kbd: '⌘X', disabled: ro, fn: () => this.clip(true) },
          { label: '复制', kbd: '⌘C', fn: () => this.clip(false) },
          { label: '粘贴', kbd: '⌘V', disabled: ro, fn: () => this.pasteFromMenu() },
          '-',
          { label: '在上方插入行', disabled: ro, fn: () => this.lines('insert_rows') },
          { label: '在左侧插入列', disabled: ro, fn: () => this.lines('insert_cols') },
          { label: '删除所在行', disabled: ro, fn: () => this.lines('delete_rows') },
          { label: '删除所在列', disabled: ro, fn: () => this.lines('delete_cols') },
          { label: '清除内容', kbd: '⌫', disabled: ro, fn: () => this.clearSel() },
          '-',
          { label: '升序排序', disabled: ro, fn: () => this.sortSel(false) },
          { label: '降序排序', disabled: ro, fn: () => this.sortSel(true) },
          { label: this.sheet.autoFilter ? '取消筛选' : '筛选', disabled: ro, fn: () => this.toggleFilter() },
          { label: '向下填充', kbd: '⌘D', disabled: ro || s.r1 === s.r2, fn: () => this.fillDown() },
          '-',
          { label: merged ? '取消合并' : '合并单元格', disabled: ro || (!merged && s.r1 === s.r2 && s.c1 === s.c2), fn: () => this.mergeSel() },
          { label: '所有框线', disabled: ro, fn: () => this.border('all') },
          { label: '外侧框线', disabled: ro, fn: () => this.border('outside') },
          { label: '无框线', disabled: ro, fn: () => this.border('none') },
          '-',
          { label: '插入图表…', disabled: ro, fn: () => this.chartTypeMenu(x, y) }
        ])
      }

      // ------------------------------------------------ charts

      chartTypeMenu (x, y) {
        menu(x, y, CHART_TYPES.map(([v, label]) => ({ label, disabled: !this.doc.writable, fn: () => this.insertChart(v) })))
      }

      // insertChart charts the selection (or the block of data around the
      // active cell) and puts the chart to its right.
      insertChart (type) {
        if (!this.doc.writable || this.chartBusy) return
        const t = this.target()
        if (t.r1 === t.r2 && t.c1 === t.c2 && this.shown(t.r1, t.c1) === '') { this.frame.toast('请先选中要做成图表的数据区域。', 'info'); return }
        this.chartBusy = true
        this.newChart = true
        this.push([{ op: 'add_chart', sheet: this.sheet.name, range: `${cellName(t.r1, t.c1)}:${cellName(t.r2, t.c2)}`, type, at: { c: t.c2 + 2, r: t.r1 } }], { now: true })
      }

      buildCharts () {
        this.chartLayer.replaceChildren()
        const charts = this.sheet.charts ?? []
        if (this.newChart && charts.length) { this.selChart = charts[charts.length - 1].k; this.newChart = false }
        const pics = (this.sheet.pics ?? []).map((p) => ({ ...p, k: 'p' + p.k, pk: p.k, pic: true }))
        if (this.newPic && pics.length) { this.selChart = pics[pics.length - 1].k; this.newPic = false }
        const list = [...charts, ...pics]
        this.chartEls = list.map((ch) => {
          const body = el('div', { class: 'qx-chart-body' })
          const box = el('div', { class: 'qx-chart' + (ch.pic ? ' qx-pic' : ''), tabindex: '0', title: ch.pic ? ch.name : ch.editable ? undefined : '这个图表含 Quire 不能改的组合，只能移动、改标题或删除', dataset: { k: String(ch.k) } }, body, ...HANDLES.map((h) => el('i', { class: 'qx-ch-h', dataset: { h } })))
          if (ch.pic) body.append(ch.src ? el('img', { src: ch.src, alt: ch.name ?? '图片', draggable: 'false' }) : el('span', { class: 'qx-pic-x' }, '图片'))
          box.addEventListener('pointerdown', (e) => this.chartDown(e, ch))
          box.addEventListener('dblclick', (e) => { e.stopPropagation(); if (!ch.pic) this.chartTitle(ch) })
          box.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); this.pickChart(ch.k); if (ch.pic) this.picMenu(ch, e.clientX, e.clientY); else this.chartMenu(ch, e.clientX, e.clientY) })
          box.addEventListener('keydown', (e) => this.chartKey(e, ch))
          box.addEventListener('wheel', (e) => { if (mod(e)) return; e.preventDefault(); this.scroll.scrollTop += e.deltaY; this.scroll.scrollLeft += e.deltaX }, { passive: false })
          this.chartLayer.append(box)
          return { ch, box, body, w: 0, h: 0 }
        })
        if (!this.chartEls.some((c) => c.ch.k === this.selChart)) this.selChart = undefined
        for (const c of this.chartEls) c.box.classList.toggle('on', c.ch.k === this.selChart)
        if (this.selChart !== undefined) setTimeout(() => this.chartEls?.find((c) => c.ch.k === this.selChart)?.box.focus({ preventScroll: true }))
      }

      chartRect (ch) {
        const x = this.sx(ch.from.c) + ch.from.dx * this.z; const y = this.sy(ch.from.r) + ch.from.dy * this.z
        const x2 = this.sx(ch.to.c) + ch.to.dx * this.z; const y2 = this.sy(ch.to.r) + ch.to.dy * this.z
        return { x, y, w: Math.max(24, x2 - x), h: Math.max(24, y2 - y) }
      }

      placeCharts () {
        for (const c of this.chartEls ?? []) {
          const r = c.preview ?? this.chartRect(c.ch)
          Object.assign(c.box.style, { left: r.x - HW + 'px', top: r.y - HH + 'px', width: r.w + 'px', height: r.h + 'px' })
          const w = Math.round(r.w); const h = Math.round(r.h)
          if (!c.ch.pic && (w !== c.w || h !== c.h)) { c.w = w; c.h = h; c.body.innerHTML = chartSvg(c.ch, w - 2, h - 2) }
        }
      }

      pickChart (k) {
        this.selChart = k
        for (const c of this.chartEls ?? []) c.box.classList.toggle('on', c.ch.k === k)
        if (k === undefined) return
        this.chartEls.find((c) => c.ch.k === k)?.box.focus({ preventScroll: true })
      }

      // anchorAt is the cell a point on the grid falls in, and the offset
      // into it at 100% zoom.
      anchorAt (x, y) {
        const ax = Math.max(0, x - HW + (this.fc && x - HW < this.FW() ? 0 : this.scroll.scrollLeft))
        const ay = Math.max(0, y - HH + (this.fr && y - HH < this.FH() ? 0 : this.scroll.scrollTop))
        const c = this.find(this.xs, ax, this.nc); const r = this.find(this.ys, ay, this.nr)
        return { c, dx: Math.round((ax - this.X(c)) / this.z), r, dy: Math.round((ay - this.Y(r)) / this.z) }
      }

      chartDown (e, ch) {
        if (e.button !== 0) return
        e.stopPropagation()
        e.preventDefault()
        if (this.editing) this.commit(0, 0)
        this.pickChart(ch.k)
        if (!this.doc.writable || this.chartBusy) return
        const c = this.chartEls.find((x) => x.ch === ch)
        const handle = e.target.dataset?.h
        const start = this.chartRect(ch)
        const ox = e.clientX; const oy = e.clientY
        let moved = false
        const move = (ev) => {
          const dx = ev.clientX - ox; const dy = ev.clientY - oy
          if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return
          moved = true
          let { x, y, w, h } = start
          if (!handle) { x += dx; y += dy } else if (ch.pic && handle.length === 2 && !ev.shiftKey) {
            // Corners keep a picture's proportions.
            const k = Math.max(16 / start.w, (start.w + (handle.includes('e') ? dx : -dx)) / start.w)
            w = start.w * k; h = start.h * k
            if (handle.includes('w')) x = start.x + start.w - w
            if (handle.includes('n')) y = start.y + start.h - h
          } else {
            if (handle.includes('e')) w = Math.max(60, start.w + dx)
            if (handle.includes('s')) h = Math.max(40, start.h + dy)
            if (handle.includes('w')) { w = Math.max(60, start.w - dx); x = start.x + start.w - w }
            if (handle.includes('n')) { h = Math.max(40, start.h - dy); y = start.y + start.h - h }
          }
          c.preview = { x, y, w, h }
          this.placeCharts()
        }
        const up = () => {
          window.removeEventListener('pointermove', move)
          const p = c.preview
          c.preview = undefined
          if (!moved || !p) { this.placeCharts(); return }
          ch.from = this.anchorAt(p.x, p.y)
          ch.to = this.anchorAt(p.x + p.w, p.y + p.h)
          this.chartOp(ch, { op: 'chart_move', from: ch.from, to: ch.to })
        }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', up, { once: true })
      }

      chartKey (e, ch) {
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.deleteChart(ch) } else if (e.key === 'Escape') { e.preventDefault(); this.pickChart(undefined); this.focus() } else if (mod(e) && e.key.toLowerCase() === 'z') { e.preventDefault(); this.doc.step(e.shiftKey ? 'redo' : 'undo') }
      }

      chartMenu (ch, x, y) {
        const ro = !this.doc.writable || this.chartBusy
        const stackable = ['col', 'bar', 'line', 'area'].includes(ch.type)
        menu(x, y, [
          ...CHART_TYPES.map(([v, label]) => ({ label: (ch.type === v ? '✓ ' : '\u3000') + label, disabled: ro || !ch.editable || ch.type === v, fn: () => this.chartOp(ch, { op: 'chart_edit', type: v }) })),
          '-',
          { label: ch.stacked ? '取消堆积' : '堆积', disabled: ro || !ch.editable || !stackable, fn: () => this.chartOp(ch, { op: 'chart_edit', stacked: ch.stacked ? null : 'stacked' }) },
          { label: '编辑标题…', disabled: ro, fn: () => this.chartTitle(ch) },
          { label: '图表选项…', disabled: ro, fn: () => this.chartOptions(ch) },
          { label: '删除图表', kbd: 'Delete', disabled: ro, fn: () => this.deleteChart(ch) }
        ])
      }

      chartTitle (ch) {
        if (!this.doc.writable || this.chartBusy) return
        const t = window.prompt('图表标题（留空则不显示）', ch.title ?? '')
        if (t === null || t === (ch.title ?? '')) return
        this.chartOp(ch, { op: 'chart_edit', title: t })
      }

      // chartOptions sets data labels, the legend, axis titles, the value
      // axis bounds and gridlines in one dialog.
      async chartOptions (ch) {
        if (!this.doc.writable || this.chartBusy) return
        const pie = ch.type === 'pie' || ch.type === 'doughnut'
        const axis = () => !pie
        const v = await dialog({
          title: '图表选项',
          fields: [
            { key: 'labels', label: '数据标签', type: 'select', value: ch.labels ?? 'none', options: [['none', '无'], ['value', '值'], ...(pie ? [['percent', '百分比']] : []), ['category', '类别名称']] },
            { key: 'legend', label: '图例', type: 'select', value: ch.legend ?? 'none', options: [['none', '无'], ['r', '右侧'], ['b', '底部'], ['t', '顶部'], ['l', '左侧'], ['tr', '右上']] },
            { key: 'x_title', label: ch.type === 'scatter' ? 'X 轴标题' : '类别轴标题', value: ch.xTitle ?? '', placeholder: '留空则不显示', show: axis },
            { key: 'y_title', label: ch.type === 'scatter' ? 'Y 轴标题' : '数值轴标题', value: ch.yTitle ?? '', placeholder: '留空则不显示', show: axis },
            { key: 'y_min', label: '数值轴最小值', value: ch.yMin ?? '', placeholder: '自动', show: axis },
            { key: 'y_max', label: '数值轴最大值', value: ch.yMax ?? '', placeholder: '自动', show: axis },
            { key: 'gridlines', label: '主要网格线', type: 'checkbox', value: ch.gridlines !== false, show: axis }
          ],
          check: (x) => {
            if (pie) return ''
            for (const k of ['y_min', 'y_max']) if (x[k].trim() !== '' && !Number.isFinite(+x[k])) return '最小值、最大值要填数字或留空'
            if (x.y_min.trim() !== '' && x.y_max.trim() !== '' && +x.y_min >= +x.y_max) return '最小值要小于最大值'
            return ''
          }
        })
        if (!v) { this.focus(); return }
        const op = { op: 'chart_edit', labels: v.labels, legend: v.legend }
        if (!pie) Object.assign(op, { x_title: v.x_title, y_title: v.y_title, y_min: v.y_min.trim() === '' ? null : +v.y_min, y_max: v.y_max.trim() === '' ? null : +v.y_max, gridlines: v.gridlines })
        this.chartOp(ch, op)
      }

      deleteChart (ch) {
        if (!this.doc.writable || this.chartBusy) return
        const c = this.chartEls.find((x) => x.ch === ch)
        if (c) c.box.hidden = true
        this.selChart = undefined
        this.chartOp(ch, { op: 'chart_delete' })
        this.focus()
      }

      // chartOp sends one chart change at once; charts are numbered by
      // position, so the next waits for the page to come back.
      chartOp (ch, op) {
        if (!this.doc.writable || this.chartBusy) return
        this.chartBusy = true
        if (ch.pic) this.push([{ ...op, op: op.op.replace('chart_', 'pic_'), sheet: this.sheet.name, pic: ch.pk }], { now: true })
        else this.push([{ ...op, sheet: this.sheet.name, chart: ch.k }], { now: true })
      }

      picMenu (ch, x, y) {
        const ro = !this.doc.writable || this.chartBusy
        menu(x, y, [
          { label: '重设大小', disabled: ro || !ch.src, fn: () => { const img = new Image(); img.onload = () => { const to = this.anchorAt(this.sx(ch.from.c) + ch.from.dx * this.z + img.naturalWidth * this.z, this.sy(ch.from.r) + ch.from.dy * this.z + img.naturalHeight * this.z); this.chartOp(ch, { op: 'chart_move', from: ch.from, to }) }; img.src = ch.src } },
          { label: '删除图片', kbd: 'Delete', disabled: ro, fn: () => this.deleteChart(ch) }
        ])
      }

      headerMenu (axis, x, y) {
        const ro = !this.doc.writable
        const s = this.sel
        const rows = axis === 'r'
        const a = rows ? s.r1 : s.c1; const b = rows ? s.r2 : s.c2
        const what = rows ? '行' : '列'
        const size = () => {
          const cur = Math.round((rows ? this.h(a) : this.w(a)) / this.z)
          const v = window.prompt(`${rows ? '行高' : '列宽'}（像素）`, String(cur))
          if (!v || !(+v > 0)) return
          if (rows) this.setHeights(a, b, Math.round(+v)); else this.setWidths(a, b, Math.round(+v))
        }
        menu(x, y, [
          { label: '剪切', kbd: '⌘X', disabled: ro, fn: () => this.clip(true) },
          { label: '复制', kbd: '⌘C', fn: () => this.clip(false) },
          { label: '粘贴', kbd: '⌘V', disabled: ro, fn: () => this.pasteFromMenu() },
          '-',
          { label: rows ? '在上方插入行' : '在左侧插入列', disabled: ro, fn: () => this.lines(rows ? 'insert_rows' : 'insert_cols') },
          { label: `删除${what}`, disabled: ro, fn: () => this.lines(rows ? 'delete_rows' : 'delete_cols') },
          { label: '清除内容', disabled: ro, fn: () => this.clearSel() },
          '-',
          { label: `隐藏${what}`, disabled: ro, fn: () => this.push([{ op: 'hide', sheet: this.sheet.name, axis, from: a, to: b, hidden: true }], { now: true }) },
          { label: `取消隐藏${what}`, disabled: ro, fn: () => this.push([{ op: 'hide', sheet: this.sheet.name, axis, from: a, to: b, hidden: false }], { now: true }) },
          { label: rows ? '行高…' : '列宽…', disabled: ro, fn: size }
        ])
      }

      // ------------------------------------------------ formula help

      suggest () {
        const src = this.editing?.bar ? this.fbar : this.inp
        const v = src.value
        const pos = src.selectionStart ?? v.length
        if (!this.editing || !v.startsWith('=')) { this.hideAc(); return }
        const head = v.slice(0, pos)
        const word = /([A-Za-z][A-Za-z0-9.]*)$/.exec(head)?.[1]
        let names = word && !/^[A-Za-z]{1,3}\d+$/.test(word) ? FN_NAMES.filter((n) => n.startsWith(word.toUpperCase())).slice(0, 8) : []
        if (names.length === 1 && names[0] === word.toUpperCase() && head[head.length] !== '(') names = names.filter(() => !/\($/.test(head))
        const defined = word ? (this.model.names ?? []).map((d) => d.name).filter((n, i, a) => a.indexOf(n) === i && n.toLowerCase().startsWith(word.toLowerCase()) && !/^_xlnm\./i.test(n)) : []
        names = [...defined.slice(0, 3), ...names].slice(0, 8)
        // The function whose parentheses hold the caret.
        let depth = 0; let inside
        for (let i = head.length - 1; i >= 0; i--) {
          if (head[i] === ')') depth++
          else if (head[i] === '(') { if (depth === 0) { inside = /([A-Za-z][A-Za-z0-9.]*)$/.exec(head.slice(0, i))?.[1]?.toUpperCase(); break } depth-- }
        }
        this.acList = names
        this.acIx = 0
        this.acSrc = src
        this.acWord = word
        this.ac.replaceChildren()
        names.forEach((n, i) => {
          const it = el('div', { class: 'qx-ac-item' + (i === 0 ? ' on' : '') }, el('b', {}, n), el('span', {}, FN_HELP[n] ? FN_HELP[n].replace(/^\S+\(.*?\)\s*/, '') : '名称 ' + ((this.model.names ?? []).find((d) => d.name === n)?.ref ?? '')))
          it.addEventListener('mousedown', (e) => { e.preventDefault(); this.acIx = i; this.acPick() })
          this.ac.append(it)
        })
        if (!names.length && inside && FN_HELP[inside]) this.ac.append(el('div', { class: 'qx-ac-tip' }, FN_HELP[inside]))
        if (!this.ac.childNodes.length) { this.hideAc(); return }
        const r = src.getBoundingClientRect()
        const host = this.root.getBoundingClientRect()
        Object.assign(this.ac.style, { left: Math.max(0, r.left - host.left) + 'px', top: r.bottom - host.top + 2 + 'px' })
        this.ac.hidden = false
      }

      hideAc () { this.ac.hidden = true; this.acList = [] }

      acKey (e, src) {
        if (this.ac.hidden || !this.acList?.length || this.acSrc !== src) return false
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault()
          this.acIx = (this.acIx + (e.key === 'ArrowDown' ? 1 : -1) + this.acList.length) % this.acList.length
          ;[...this.ac.children].forEach((n, i) => n.classList.toggle('on', i === this.acIx))
          return true
        }
        if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey)) { e.preventDefault(); this.acPick(); return true }
        if (e.key === 'Escape') { e.preventDefault(); this.hideAc(); return true }
        return false
      }

      acPick () {
        const src = this.acSrc
        const name = this.acList[this.acIx]
        if (!src || !name) return
        const pos = src.selectionStart
        const start = pos - (this.acWord?.length ?? 0)
        const call = FN_HELP[name] ? name + '(' : name
        src.value = src.value.slice(0, start) + call + src.value.slice(pos)
        const at = start + call.length
        src.setSelectionRange(at, at)
        if (src === this.inp) { this.fbar.value = this.inp.value; this.fitInput() } else if (this.editing && !this.editing.bar) this.inp.value = this.fbar.value
        src.focus()
        this.suggest()
      }

      // ------------------------------------------------ formula tools

      // selRef names the selection, whole columns and rows included.
      selRef (abs) {
        const s = this.sel
        const whole = s.r1 === 1 && s.r2 >= this.nr
        const wide = s.c1 === 1 && s.c2 >= this.nc
        const d = abs ? '$' : ''
        const at = (r, c) => d + colName(c) + d + r
        if (whole && !wide) return `${d}${colName(s.c1)}:${d}${colName(s.c2)}`
        if (wide && !whole) return `${d}${s.r1}:${d}${s.r2}`
        return s.r1 === s.r2 && s.c1 === s.c2 ? at(s.ar, s.ac) : `${at(s.r1, s.c1)}:${at(Math.min(s.r2, 1048576), Math.min(s.c2, 16384))}`
      }

      quoteSheet (n) { return /^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'` }

      // pointable: the caret in the formula being typed waits for a
      // reference, so clicking a cell inserts it.
      pointable () {
        const src = this.editing.bar ? this.fbar : this.inp
        const v = src.value
        if (!v.startsWith('=')) return false
        const pos = src.selectionStart ?? v.length
        if (this.pointing && this.pointing.end === pos) return true
        const before = v.slice(0, pos).replace(/\s+$/, '')
        return /[=(,;+\-*/^&<>:]$/.test(before)
      }

      pointRef (h) {
        const src = this.editing.bar ? this.fbar : this.inp
        const pos = src.selectionStart ?? src.value.length
        const start = this.pointing && this.pointing.end === pos ? this.pointing.start : pos
        const anchor = this.pointing && this.pointing.end === pos && this.pointing.anchor && this.pointing.sheet === this.sheet.name ? this.pointing.anchor : { r: h.r, c: h.c }
        const prefix = ''
        const put = (a, b) => {
          const r1 = Math.min(a.r, b.r); const r2 = Math.max(a.r, b.r); const c1 = Math.min(a.c, b.c); const c2 = Math.max(a.c, b.c)
          const ref = prefix + (r1 === r2 && c1 === c2 ? cellName(r1, c1) : `${cellName(r1, c1)}:${cellName(r2, c2)}`)
          const end = this.pointing?.start === start ? this.pointing.end : pos
          src.value = src.value.slice(0, start) + ref + src.value.slice(end)
          this.pointing = { start, end: start + ref.length, anchor: a, sheet: this.sheet.name }
          src.setSelectionRange(start + ref.length, start + ref.length)
          if (src === this.inp) { this.fbar.value = this.inp.value; this.fitInput() } else if (!this.editing.bar) this.inp.value = this.fbar.value
          this.draw()
        }
        this.pointing = { start, end: pos, anchor, sheet: this.sheet.name }
        put({ r: h.r, c: h.c }, { r: h.r, c: h.c })
        src.focus({ preventScroll: true })
        const move = (ev) => { const g = this.hit(ev); if (g.x > HW && g.y > HH) put({ r: h.r, c: h.c }, { r: this.rowAt(g.y), c: this.colAt(g.x) }) }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', () => { window.removeEventListener('pointermove', move); src.focus({ preventScroll: true }) }, { once: true })
      }

      // autoSum writes =SUM() (or another function) over the numbers next
      // to the selection.
      autoSum (fn) {
        if (!this.doc.writable) return
        const s = this.sel
        const isNum = (r, c) => this.cell(r, c)?.[3] === 1
        if (this.editing) { this.startFormula(fn); return }
        if (s.r1 === s.r2 && s.c1 === s.c2) {
          let ref = ''
          let r = s.ar - 1
          while (r >= 1 && isNum(r, s.ac)) r--
          if (r < s.ar - 1) ref = `${cellName(r + 1, s.ac)}:${cellName(s.ar - 1, s.ac)}`
          else {
            let c = s.ac - 1
            while (c >= 1 && isNum(s.ar, c)) c--
            if (c < s.ac - 1) ref = `${cellName(s.ar, c + 1)}:${cellName(s.ar, s.ac - 1)}`
          }
          this.beginEdit('edit')
          this.inp.value = `=${fn}(${ref})`
          const at = this.inp.value.length - 1
          this.inp.setSelectionRange(ref ? fn.length + 2 : at, at)
          this.fbar.value = this.inp.value
          this.fitInput()
          this.suggest()
          this.draw()
          return
        }
        // A block: totals under each column (or in its empty last row).
        const r2 = Math.min(s.r2, Math.max(this.sheet.maxR, s.r1)); const c2 = Math.min(s.c2, Math.max(this.sheet.maxC, s.c1))
        let lastEmpty = r2 > s.r1
        for (let c = s.c1; c <= c2 && lastEmpty; c++) if (this.shown(r2, c) !== '') lastEmpty = false
        const at = lastEmpty ? r2 : r2 + 1
        const top = s.r1; const bottom = lastEmpty ? r2 - 1 : r2
        const ops = []; const over = []
        for (let c = s.c1; c <= c2; c++) {
          let any = false
          for (let r = top; r <= bottom && !any; r++) any = isNum(r, c)
          if (!any) continue
          const text = `=${fn}(${cellName(top, c)}:${cellName(bottom, c)})`
          ops.push({ op: 'input', sheet: this.sheet.name, cell: cellName(at, c), text })
          over.push([at, c, text])
        }
        if (!ops.length) { this.frame.toast('选区里没有数字可以计算。', 'info'); return }
        this.push(ops, { overlay: over, now: true })
      }

      // startFormula puts NAME() into the cell being edited, caret inside.
      startFormula (name) {
        if (!this.doc.writable) return
        if (!this.editing) { this.beginEdit('edit'); this.inp.value = this.inp.value.startsWith('=') ? this.inp.value : '' }
        const src = this.editing.bar ? this.fbar : this.inp
        let pos = src.selectionStart ?? src.value.length
        if (!src.value.startsWith('=')) { src.value = '='; pos = 1 }
        const end = src.selectionEnd ?? pos
        src.value = src.value.slice(0, pos) + name + '()' + src.value.slice(end)
        const at = pos + name.length + 1
        src.focus({ preventScroll: true })
        src.setSelectionRange(at, at)
        if (src === this.inp) { this.fbar.value = this.inp.value; this.fitInput() } else if (!this.editing.bar) this.inp.value = this.fbar.value
        this.suggest()
        this.draw()
      }

      async insertFunction () {
        if (!this.doc.writable) return
        let chosen = FN_CATS[0][1][0]
        const q = el('input', { class: 'q-fn-q', placeholder: '搜索函数，例如 求和、查找、VLOOKUP', spellcheck: 'false' })
        const cat = el('select', { class: 'q-sel' }, ...FN_CATS.map(([c], i) => el('option', { value: String(i) }, c)), el('option', { value: 'all' }, '全部'))
        const list = el('div', { class: 'q-fn-list', role: 'listbox' })
        const help = el('div', { class: 'q-fn-help' })
        const show = () => {
          const t = q.value.trim().toLowerCase()
          const pool = t || cat.value === 'all' ? FN_NAMES : FN_CATS[+cat.value][1]
          const names = pool.filter((n) => !t || n.toLowerCase().includes(t) || FN_HELP[n].toLowerCase().includes(t))
          if (!names.includes(chosen)) chosen = names[0]
          list.replaceChildren(...names.map((n) => {
            const it = el('div', { class: 'q-fn-it' + (n === chosen ? ' on' : ''), role: 'option' }, n)
            it.addEventListener('click', () => { chosen = n; show() })
            it.addEventListener('dblclick', () => { chosen = n; list.closest('form').requestSubmit() })
            return it
          }))
          help.replaceChildren(...(chosen ? [el('b', {}, FN_HELP[chosen].replace(/\)\s.*$/, ')')), el('p', {}, FN_HELP[chosen].replace(/^\S+\(.*?\)\s*/, ''))] : [el('p', {}, '没有找到函数')]))
          list.querySelector('.on')?.scrollIntoView({ block: 'nearest' })
        }
        q.addEventListener('input', show)
        cat.addEventListener('change', show)
        q.addEventListener('keydown', (e) => {
          if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return
          e.preventDefault()
          const names = [...list.children].map((n) => n.textContent)
          chosen = names[Math.max(0, Math.min(names.length - 1, names.indexOf(chosen) + (e.key === 'ArrowDown' ? 1 : -1)))]
          show()
        })
        show()
        const body = el('div', { class: 'q-fn' }, el('div', { class: 'q-fn-top' }, q, cat), list, help)
        const name = await dialog({ title: '插入函数', body, read: () => chosen, check: (v) => v ? '' : '请选择一个函数', wide: true })
        setTimeout(() => q.focus())
        if (name) this.startFormula(name)
        else this.focus()
      }

      async defineName (old) {
        if (!this.doc.writable) return
        const sheets = this.model.sheets.map((x) => x.name)
        const scope = old?.sheet === undefined ? '' : sheets[old.sheet] ?? ''
        const v = await dialog({
          title: old ? '编辑名称' : '新建名称',
          fields: [
            { key: 'name', label: '名称', value: old?.name ?? '', placeholder: '例如 销售额' },
            { key: 'scope', label: '范围', type: 'select', value: scope, options: [['', '工作簿'], ...sheets.map((n) => [n, n])] },
            { key: 'ref', label: '引用位置', value: old ? '=' + old.ref : `=${this.quoteSheet(this.sheet.name)}!${this.selRef(true)}` },
            { key: 'comment', label: '备注', value: old?.comment ?? '' }
          ],
          check: (x) => !/^[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*$/.test(x.name.trim()) || /^[A-Za-z]{1,3}\d+$/.test(x.name.trim()) ? '名称以字母、汉字或下划线开头，不能有空格，也不能像单元格地址' : x.ref.replace(/^=/, '').trim() ? '' : '请填写引用位置'
        })
        this.focus()
        if (!v) return
        this.push([{ op: 'name_set', name: v.name.trim(), ref: v.ref.replace(/^=/, '').trim(), sheet: v.scope || undefined, comment: v.comment.trim() || undefined, old: old?.name, oldSheet: old ? scope || undefined : undefined }], { now: true })
      }

      nameManager () {
        const names = (this.model.names ?? []).filter((d) => !d.hidden && !/^_xlnm\./i.test(d.name))
        const sheets = this.model.sheets.map((x) => x.name)
        const body = (done) => {
          const table = el('div', { class: 'q-names' })
          table.append(el('div', { class: 'q-names-r q-names-h' }, el('span', {}, '名称'), el('span', {}, '引用位置'), el('span', {}, '范围'), el('span', {})))
          for (const d of names) {
            const edit = el('button', { type: 'button', disabled: !this.doc.writable }, '编辑')
            const del = el('button', { type: 'button', disabled: !this.doc.writable }, '删除')
            const row = el('div', { class: 'q-names-r' }, el('b', {}, d.name), el('code', {}, d.ref), el('span', {}, d.sheet === undefined ? '工作簿' : sheets[d.sheet] ?? ''), el('span', { class: 'q-names-a' }, edit, del))
            edit.addEventListener('click', () => { done(null); this.defineName(d) })
            del.addEventListener('click', () => { this.push([{ op: 'name_delete', name: d.name, sheet: d.sheet === undefined ? undefined : sheets[d.sheet] }], { now: true }); row.remove() })
            table.append(row)
          }
          if (!names.length) table.append(el('div', { class: 'q-names-empty' }, '这个工作簿还没有名称。选中区域后点“新建”，公式里就能用名称代替地址。'))
          return table
        }
        dialog({ title: '名称管理器', body, read: () => true, ok: '关闭', wide: true, buttons: this.doc.writable ? [['新建…', (done) => { done(null); this.defineName() }]] : [] }).then(() => this.focus())
      }

      // ------------------------------------------------ conditional formats

      cfMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const x = r.left; const y = r.bottom + 2
        const ro = !this.doc.writable
        const sub = (items) => () => setTimeout(() => menu(x, y, items))
        menu(x, y, [
          { label: '突出显示单元格规则 ▸', disabled: ro, fn: sub([
            ...['greaterThan', 'lessThan', 'between', 'equal'].map((op) => ({ label: CF_OPS[op] + '…', fn: () => this.cfAsk('cellIs', op) })),
            { label: '文本包含…', fn: () => this.cfAsk('containsText') },
            { label: '重复值…', fn: () => this.cfAsk('duplicateValues') },
            { label: '唯一值…', fn: () => this.cfAsk('uniqueValues') }
          ]) },
          { label: '最前/最后规则 ▸', disabled: ro, fn: sub([
            { label: '前 10 项…', fn: () => this.cfAsk('top10') },
            { label: '前 10%…', fn: () => this.cfAsk('top10', 'percent') },
            { label: '最后 10 项…', fn: () => this.cfAsk('top10', 'bottom') },
            { label: '高于平均值…', fn: () => this.cfAsk('aboveAverage') },
            { label: '低于平均值…', fn: () => this.cfAsk('aboveAverage', 'below') }
          ]) },
          { label: '数据条 ▸', disabled: ro, fn: sub(BAR_COLORS.map(([c, l]) => ({ label: l, fn: () => this.cfAdd({ type: 'dataBar', color: c }) }))) },
          { label: '色阶 ▸', disabled: ro, fn: sub(SCALES.map(([colors, l]) => ({ label: l, fn: () => this.cfAdd({ type: 'colorScale', colors }) }))) },
          '-',
          { label: '新建规则（公式）…', disabled: ro, fn: () => this.cfAsk('expression') },
          { label: '管理规则…', fn: () => this.cfManage() },
          '-',
          { label: '清除所选单元格的规则', disabled: ro, fn: () => this.push([{ op: 'cf_clear', sheet: this.sheet.name, range: this.selRef() }], { now: true }) },
          { label: '清除整个工作表的规则', disabled: ro || !this.sheet.cf?.length, fn: () => this.push([{ op: 'cf_clear', sheet: this.sheet.name }], { now: true }) }
        ])
      }

      cfAdd (rule) { this.push([{ op: 'cf_add', sheet: this.sheet.name, range: this.selRef(), rule }], { now: true }) }

      async cfAsk (type, how) {
        const styleField = { key: 'style', label: '设置为', type: 'select', value: 'red', options: CF_STYLES.map(([k, l]) => [k, l]) }
        let fields; let title
        const raw = this.input(this.sel.ar, this.sel.ac)
        const cur = String(raw).startsWith('=') ? this.shown(this.sel.ar, this.sel.ac) : raw
        if (type === 'cellIs') {
          title = `为${CF_OPS[how]}以下值的单元格设置格式`
          fields = [{ key: 'v1', label: how === 'between' ? '最小值' : '值', value: cur, placeholder: '数字、文本或 =B1' }, ...(how === 'between' ? [{ key: 'v2', label: '最大值', value: '' }] : []), styleField]
        } else if (type === 'containsText') {
          title = '为包含以下文本的单元格设置格式'
          fields = [{ key: 'text', label: '文本', value: cur }, styleField]
        } else if (type === 'duplicateValues' || type === 'uniqueValues') {
          title = type === 'duplicateValues' ? '为重复值设置格式' : '为唯一值设置格式'
          fields = [styleField]
        } else if (type === 'top10') {
          title = how === 'bottom' ? '为值最小的那些单元格设置格式' : '为值最大的那些单元格设置格式'
          fields = [{ key: 'rank', label: how === 'percent' ? '前百分比' : '项数', type: 'number', value: '10' }, styleField]
        } else if (type === 'aboveAverage') {
          title = how === 'below' ? '为低于平均值的单元格设置格式' : '为高于平均值的单元格设置格式'
          fields = [styleField]
        } else {
          title = '使用公式确定要设置格式的单元格'
          fields = [{ key: 'f', label: '公式', value: '=', placeholder: `例如 =${cellName(this.sel.r1, this.sel.c1)}>100`, hint: `公式相对于选区左上角 ${cellName(this.sel.r1, this.sel.c1)}，结果为 TRUE 的单元格设置格式` }, styleField]
        }
        const v = await dialog({ title, fields, check: (x) => (type === 'cellIs' && x.v1 === '') || (type === 'containsText' && !x.text) || (type === 'expression' && x.f.replace(/^=/, '').trim() === '') ? '请填写条件' : '' })
        this.focus()
        if (!v) return
        const style = CF_STYLES.find(([k]) => k === v.style)[2]
        const rule = { type, style }
        if (type === 'cellIs') { rule.op = how; rule.f = how === 'between' ? [v.v1, v.v2] : [v.v1] }
        if (type === 'containsText') rule.text = v.text
        if (type === 'top10') { rule.rank = +v.rank || 10; if (how === 'percent') rule.percent = true; if (how === 'bottom') rule.bottom = true }
        if (type === 'aboveAverage' && how === 'below') rule.below = true
        if (type === 'expression') rule.f = [v.f]
        this.cfAdd(rule)
      }

      cfManage () {
        const rules = this.sheet.cf ?? []
        const describe = (x) => {
          if (x.type === 'cellIs') return `单元格值${CF_OPS[x.op] ?? x.op} ${(x.f ?? []).join(' 和 ')}`
          if (x.type === 'containsText') return `包含“${x.text}”`
          if (x.type === 'notContainsText') return `不包含“${x.text}”`
          if (x.type === 'beginsWith') return `开头是“${x.text}”`
          if (x.type === 'endsWith') return `结尾是“${x.text}”`
          if (x.type === 'duplicateValues') return '重复值'
          if (x.type === 'uniqueValues') return '唯一值'
          if (x.type === 'top10') return `${x.bottom ? '最后' : '前'} ${x.rank}${x.percent ? '%' : ' 项'}`
          if (x.type === 'aboveAverage') return x.below ? '低于平均值' : '高于平均值'
          if (x.type === 'expression') return `公式 =${x.f?.[0] ?? ''}`
          if (x.type === 'dataBar') return '数据条'
          if (x.type === 'colorScale') return `${x.colors?.length ?? 2} 色刻度`
          return x.type
        }
        const swatch = (x) => {
          if (x.type === 'dataBar') return el('i', { class: 'q-cf-sw', style: { background: `linear-gradient(90deg,#${x.colors?.[0] ?? '638EC6'},#fff)` } })
          if (x.type === 'colorScale') return el('i', { class: 'q-cf-sw', style: { background: `linear-gradient(90deg,${(x.colors ?? []).map((c) => '#' + c).join(',')})` } })
          return el('i', { class: 'q-cf-sw', style: { background: x.style?.fill ? '#' + x.style.fill : '#fff', color: x.style?.color ? '#' + x.style.color : '#000', fontWeight: x.style?.b ? 700 : 400 } }, 'AaBb')
        }
        const gone = []
        const table = el('div', { class: 'q-names' })
        table.append(el('div', { class: 'q-names-r q-names-h' }, el('span', {}, '规则'), el('span', {}, '应用于'), el('span', {}, '格式'), el('span', {})))
        rules.forEach((x, i) => {
          const del = el('button', { type: 'button', disabled: !this.doc.writable }, '删除')
          const row = el('div', { class: 'q-names-r' }, el('span', {}, describe(x)), el('code', {}, x.sqref), swatch(x), el('span', { class: 'q-names-a' }, del))
          // Earlier deletions in this batch move the later rules up.
          del.addEventListener('click', () => { this.push([{ op: 'cf_delete', sheet: this.sheet.name, index: i - gone.filter((j) => j < i).length, sqref: x.sqref, type: x.type }], { now: true }); gone.push(i); row.remove() })
          table.append(row)
        })
        if (!rules.length) table.append(el('div', { class: 'q-names-empty' }, '这张工作表还没有条件格式。'))
        dialog({ title: '条件格式规则管理器', body: table, read: () => true, ok: '关闭', wide: true }).then(() => this.focus())
      }

      // ------------------------------------------------ data validation

      dvAt (r, c) { return this.sheet.dv?.find((d) => rangeList(d.sqref).some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) }

      dvMenu (x, y) {
        const ro = !this.doc.writable
        menu(x, y, [
          { label: '数据验证…', disabled: ro, fn: () => this.dvAsk() },
          { label: '清除所选单元格的验证', disabled: ro || !this.sheet.dv?.length, fn: () => this.push([{ op: 'dv_clear', sheet: this.sheet.name, range: this.selRef() }], { now: true }) },
          '-',
          { label: '圈释无效数据', disabled: !this.sheet.dv?.length, fn: () => this.circleInvalid() }
        ])
      }

      async dvAsk () {
        const d = this.dvAt(this.sel.ar, this.sel.ac)
        const between = (x) => x.type !== 'list' && x.type !== 'custom' && x.type !== 'any' && /etween/.test(x.op)
        const ranged = (x) => x.type !== 'list' && x.type !== 'custom' && x.type !== 'any'
        const toDate = (n) => { if (n === undefined || n === '') return ''; const t = new Date(Math.round((+n - 25569) * 86400000)); return Number.isNaN(t.getTime()) ? String(n) : t.toISOString().slice(0, 10) }
        const v = await dialog({
          title: '数据验证',
          fields: [
            { key: 'type', label: '允许', type: 'select', value: d?.type ?? 'list', options: [['list', '序列（下拉列表）'], ['whole', '整数'], ['decimal', '小数'], ['date', '日期'], ['textLength', '文本长度'], ['custom', '自定义公式'], ['any', '任何值（清除验证）']] },
            { key: 'op', label: '数据', type: 'select', value: d?.op ?? 'between', options: Object.entries(CF_OPS).map(([k, l]) => [k, l]), show: ranged },
            { key: 'list', label: '来源', value: d?.type === 'list' ? (d.f1?.startsWith('"') ? d.items?.join(',') ?? '' : '=' + (d.f1 ?? '')) : '', placeholder: '是,否  或  =$H$2:$H$9', hint: '用逗号分隔各项，或用 = 引用单元格区域', show: (x) => x.type === 'list' },
            { key: 'f1', label: '最小值', value: d && d.type === 'date' ? toDate(d.v1) : d?.f1 ?? '', show: ranged },
            { key: 'f2', label: '最大值', value: d && d.type === 'date' ? toDate(d.v2) : d?.f2 ?? '', show: between },
            { key: 'formula', label: '公式', value: d?.type === 'custom' ? '=' + (d.f1 ?? '') : '=', show: (x) => x.type === 'custom' },
            { key: 'blank', label: '忽略空值', type: 'checkbox', value: d ? d.blank : true, show: (x) => x.type !== 'any' },
            { key: 'prompt', label: '输入提示', value: d?.prompt ?? '', show: (x) => x.type !== 'any' },
            { key: 'error', label: '出错警告', value: d?.error ?? '', placeholder: '输入无效时显示的信息', show: (x) => x.type !== 'any' },
            { key: 'style', label: '出错时', type: 'select', value: d?.style ?? 'stop', options: [['stop', '停止（拒绝输入）'], ['warning', '警告'], ['information', '仅提示']], show: (x) => x.type !== 'any' }
          ],
          check: (x) => (x.type === 'list' && !x.list.trim()) ? '请填写序列来源' : ranged(x) && x.f1 === '' ? '请填写数值' : between(x) && x.f2 === '' ? '请填写最大值' : x.type === 'date' && (Number.isNaN(dateSerial(x.f1)) || (between(x) && Number.isNaN(dateSerial(x.f2)))) ? '日期写成 2026-01-31 这样' : ''
        })
        this.focus()
        if (!v) return
        if (v.type === 'any') { this.push([{ op: 'dv_clear', sheet: this.sheet.name, range: this.selRef() }], { now: true }); return }
        const num = (t) => v.type === 'date' ? String(dateSerial(t)) : t
        const rule = { type: v.type, op: v.op, blank: v.blank, prompt: v.prompt || undefined, error: v.error || undefined, style: v.style }
        if (v.type === 'list') rule.list = v.list.trim()
        else if (v.type === 'custom') rule.f1 = v.formula
        else { rule.f1 = num(v.f1); if (between(v)) rule.f2 = num(v.f2) }
        this.push([{ op: 'dv_set', sheet: this.sheet.name, range: this.selRef(), rule }], { now: true })
      }

      // invalid checks typed text against the cell's validation.
      invalid (r, c, text) {
        const d = this.dvAt(r, c)
        if (!d || d.type === 'any' || d.type === 'custom' || String(text).startsWith('=')) return null
        const t = String(text).trim()
        const stop = !d.style || d.style === 'stop'
        const say = (fallback) => ({ stop, msg: d.error || fallback })
        if (t === '') return d.blank ? null : say('这个单元格不能为空。')
        if (d.type === 'list') return !d.items || d.items.some((x) => String(x).toLowerCase() === t.toLowerCase()) ? null : say(`请从列表中选择：${(d.items ?? []).slice(0, 8).join('、')}${d.items?.length > 8 ? '…' : ''}`)
        let v = d.type === 'textLength' ? t.length : d.type === 'date' ? dateSerial(t) : Number(t.replace(/,/g, ''))
        if (d.type === 'date' && Number.isNaN(v) && Number.isFinite(Number(t))) v = Number(t)
        const kind = d.type === 'whole' ? '整数' : d.type === 'decimal' ? '数字' : d.type === 'date' ? '日期' : '文本长度'
        if (Number.isNaN(v) || (d.type === 'whole' && !Number.isInteger(v))) return say(`这里只能输入${kind}。`)
        const a = d.v1; const b = d.v2
        if (a === undefined) return null
        const ok = { between: v >= a && v <= b, notBetween: v < a || v > b, equal: v === a, notEqual: v !== a, greaterThan: v > a, lessThan: v < a, greaterThanOrEqual: v >= a, lessThanOrEqual: v <= a }[d.op] ?? true
        const show = (n) => d.type === 'date' ? new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10) : n
        return ok ? null : say(`${kind}应${CF_OPS[d.op] ?? ''} ${show(a)}${/etween/.test(d.op) ? ' 和 ' + show(b) : ''}。`)
      }

      dvList () {
        const s = this.sel
        const d = this.dvAt(s.ar, s.ac)
        if (!d || d.type !== 'list' || !this.doc.writable) return
        const box = this.dvBtn ?? { x: this.sx(s.ac), y: this.sy(s.ar), w: this.w(s.ac), h: this.h(s.ar) }
        const g = this.grid.getBoundingClientRect()
        const cur = this.shown(s.ar, s.ac)
        const m = menu(g.left + this.sx(s.ac), g.top + box.y + box.h + 1, (d.items ?? []).map((it) => ({ label: (it === cur ? '✓ ' : '') + it, fn: () => { this.push([{ op: 'input', sheet: this.sheet.name, cell: cellName(s.ar, s.ac), text: it }], { overlay: [[s.ar, s.ac, it]] }); this.focus() } })))
        m.style.minWidth = Math.max(120, this.w(s.ac) + 17) + 'px'
      }

      showTip () {
        const s = this.sel
        const d = this.dvAt(s.ar, s.ac)
        if (!d?.prompt || (s.r1 !== s.r2 || s.c1 !== s.c2)) { this.tip.hidden = true; return }
        this.tip.replaceChildren(...[d.promptTitle && el('b', {}, d.promptTitle), el('span', {}, d.prompt)].filter(Boolean))
        const b = this.selBox()
        Object.assign(this.tip.style, { left: Math.max(HW, b.x + 8) + 'px', top: Math.max(HH, b.y + b.h + 4) + 'px' })
        this.tip.hidden = false
      }

      // circleInvalid marks cells whose values break their validation.
      circleInvalid () {
        let n = 0; let first
        for (const d of this.sheet.dv ?? []) {
          for (const g of rangeList(d.sqref)) {
            for (let r = g.r1; r <= Math.min(g.r2, this.sheet.maxR); r++) {
              for (let c = g.c1; c <= Math.min(g.c2, this.sheet.maxC); c++) {
                const v = this.input(r, c)
                if (v === '' || String(v).startsWith('=')) continue
                if (this.invalid(r, c, v)) { n++; first = first ?? [r, c] }
              }
            }
          }
        }
        if (!n) { this.frame.toast('没有发现无效数据。', 'info'); return }
        this.frame.toast(`发现 ${n} 个无效数据，已选中第一个。`, 'warn')
        this.select(first[0], first[1], first[0], first[1])
      }

      // ------------------------------------------------ cells: painter, notes, links

      startPainter (sticky) {
        const s = this.sel
        const r2 = Math.min(s.r2, s.r1 + 199); const c2 = Math.min(s.c2, s.c1 + 49)
        const grid = []
        for (let r = s.r1; r <= r2; r++) { const row = []; for (let c = s.c1; c <= c2; c++) row.push(this.cell(r, c)?.[2] ?? 0); grid.push(row) }
        this.painter = { grid, sticky }
        this.ui.painter.classList.add('on')
        this.frame.toast(sticky ? '格式刷：依次选择目标区域，Esc 结束' : '格式刷：选择要套用格式的区域', 'info')
      }

      stopPainter () { this.painter = null; this.ui.painter.classList.remove('on'); this.scroll.style.cursor = 'cell' }

      paintSel () {
        const p = this.painter
        if (!p) return
        const s = this.sel
        const h = p.grid.length; const w = p.grid[0].length
        // A single target cell takes the whole source block.
        const r2 = s.r1 === s.r2 && s.c1 === s.c2 ? s.r1 + h - 1 : Math.min(s.r2, s.r1 + 999)
        const c2 = s.r1 === s.r2 && s.c1 === s.c2 ? s.c1 + w - 1 : Math.min(s.c2, s.c1 + 199)
        const ops = []
        for (let r = s.r1; r <= r2; r++) for (let c = s.c1; c <= c2; c++) {
          const xf = p.grid[(r - s.r1) % h][(c - s.c1) % w]
          if ((this.cell(r, c)?.[2] ?? 0) !== xf) ops.push({ op: 'xf', sheet: this.sheet.name, cell: cellName(r, c), xf })
        }
        if (!p.sticky) this.stopPainter()
        if (ops.length) this.push(ops, { now: true })
      }

      async editNote () {
        if (!this.doc.writable) return
        const s = this.sel
        const cur = this.sheet.notes?.[`${s.ar},${s.ac}`]
        const v = await dialog({ title: `${cur ? '编辑' : '新建'}批注 · ${cellName(s.ar, s.ac)}`, fields: [{ key: 't', label: '内容', type: 'textarea', value: cur?.t ?? '' }], ok: '保存', buttons: cur ? [['删除批注', (done) => { done(null); this.push([{ op: 'note_set', sheet: this.sheet.name, cell: cellName(s.ar, s.ac), text: null }], { now: true }) }]] : [] })
        this.focus()
        if (!v) return
        this.push([{ op: 'note_set', sheet: this.sheet.name, cell: cellName(s.ar, s.ac), text: v.t.trim() || null, author: this.author() }], { now: true })
      }

      author () {
        let a = ''
        try { a = localStorage.getItem('quire.author') ?? '' } catch {}
        return a.trim() || 'Quire'
      }

      deleteNotes () {
        const s = this.sel
        const ops = Object.keys(this.sheet.notes ?? {}).map((k) => k.split(',').map(Number)).filter(([r, c]) => r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2).map(([r, c]) => ({ op: 'note_set', sheet: this.sheet.name, cell: cellName(r, c), text: null }))
        if (!ops.length) { this.frame.toast('所选单元格没有批注。', 'info'); return }
        this.push(ops, { now: true })
      }

      async editLink () {
        if (!this.doc.writable) return
        const s = this.sel
        const cur = this.sheet.links?.[`${s.ar},${s.ac}`]
        const shown = this.shown(s.ar, s.ac)
        const v = await dialog({
          title: cur ? '编辑链接' : '插入链接',
          fields: [
            { key: 'url', label: '地址', value: cur ?? 'https://', placeholder: 'https://… 或 #汇总!A1', hint: '以 # 开头跳到本工作簿的位置，例如 #汇总!A1' },
            { key: 'text', label: '显示文字', value: shown || '', placeholder: '留空则显示地址' }
          ],
          buttons: cur ? [['移除链接', (done) => { done(null); this.push([{ op: 'link_set', sheet: this.sheet.name, cell: cellName(s.ar, s.ac), url: null }, { op: 'style', sheet: this.sheet.name, range: cellName(s.ar, s.ac), underline: false, color: '000000' }], { now: true }) }]] : [],
          check: (x) => /^(https?:\/\/.+|mailto:.+|#.+)$/i.test(x.url.trim()) ? '' : '地址以 http://、https://、mailto: 或 # 开头'
        })
        this.focus()
        if (!v) return
        const cell = cellName(s.ar, s.ac)
        const text = v.text.trim() || v.url.trim()
        const ops = [{ op: 'link_set', sheet: this.sheet.name, cell, url: v.url.trim() }, { op: 'style', sheet: this.sheet.name, range: cell, underline: true, color: '0563C1' }]
        if (text !== this.input(s.ar, s.ac)) ops.unshift({ op: 'input', sheet: this.sheet.name, cell, text })
        this.push(ops, { overlay: [[s.ar, s.ac, text]], now: true })
      }

      followLink (url) {
        if (url.startsWith('#')) {
          const m = /^#(?:'((?:[^']|'')+)'|([^!]+))?!?\$?([A-Za-z]{1,3})\$?(\d+)/.exec(url)
          if (!m) return
          const name = (m[1] ?? m[2] ?? this.sheet.name).replace(/''/g, "'")
          const ix = this.model.sheets.findIndex((x) => x.name === name)
          if (ix >= 0 && ix !== this.sheetIx) this.useSheet(ix, false)
          const g = rangeList(m[3] + m[4])[0]
          if (g) this.select(g.r1, g.c1, g.r1, g.c1)
          return
        }
        window.open(url, '_blank', 'noopener')
      }

      async insertPic () {
        if (!this.doc.writable) return
        const src = await pickImage()
        if (!src) return
        const img = new Image()
        img.onload = () => {
          const k = Math.min(1, 480 / img.naturalWidth, 360 / img.naturalHeight)
          this.newPic = true
          this.push([{ op: 'pic_add', sheet: this.sheet.name, src, at: { r: this.sel.ar, c: this.sel.ac }, w: Math.round(img.naturalWidth * k), h: Math.round(img.naturalHeight * k) }], { now: true })
        }
        img.onerror = () => this.frame.toast('读不出这张图片。', 'warn')
        img.src = src
      }

      // ------------------------------------------------ data tools

      // headerNames names a range's columns by its first row when that row
      // looks like headings.
      headerNames (rg) {
        const first = []; const second = []
        for (let c = rg.c1; c <= rg.c2; c++) { first.push(this.cell(rg.r1, c)?.[3] ?? (this.shown(rg.r1, c) === '' ? -1 : 0)); second.push(rg.r2 > rg.r1 ? this.cell(rg.r1 + 1, c)?.[3] ?? -1 : -1) }
        const header = first.every((k) => k === 0) && (second.some((k) => k !== 0 && k !== -1) || rg.r2 > rg.r1)
        const names = []
        for (let c = rg.c1; c <= rg.c2; c++) names.push([c, header && this.shown(rg.r1, c) ? `${this.shown(rg.r1, c)}（${colName(c)} 列）` : `${colName(c)} 列`])
        return { header, names }
      }

      async sortDialog () {
        if (!this.doc.writable) return
        const rg = this.target()
        if (rg.r2 <= rg.r1) { this.frame.toast('请先选中要排序的数据。', 'info'); return }
        const { header, names } = this.headerNames(rg)
        const levels = []
        const list = el('div', { class: 'q-sort' })
        const hdr = el('input', { type: 'checkbox' })
        hdr.checked = header
        const addLevel = (col, desc) => {
          const cs = el('select', { class: 'q-sel' }, ...names.map(([c, l]) => el('option', { value: String(c), selected: c === col }, l)))
          const os = el('select', { class: 'q-sel' }, el('option', { value: '0', selected: !desc }, '升序'), el('option', { value: '1', selected: !!desc }, '降序'))
          const del = el('button', { type: 'button', title: '删除这一级' }, '×')
          const row = el('div', { class: 'q-sort-r' }, el('span', {}, levels.length ? '次要关键字' : '主要关键字'), cs, os, del)
          const lvl = { cs, os, row }
          del.addEventListener('click', () => { if (levels.length > 1) { levels.splice(levels.indexOf(lvl), 1); row.remove(); levels.forEach((l, i) => { l.row.firstChild.textContent = i ? '次要关键字' : '主要关键字' }) } })
          levels.push(lvl)
          list.append(row)
        }
        addLevel(Math.max(rg.c1, Math.min(rg.c2, this.sel.ac)), false)
        const body = el('div', {}, list, el('label', { class: 'q-dlg-row q-dlg-check' }, el('span', {}, '数据包含标题'), hdr))
        const keys = await dialog({ title: `排序 · ${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`, body, wide: true, buttons: [['添加条件', () => { if (levels.length < 6) addLevel(rg.c1, false) }]], read: () => levels.map((l) => ({ col: +l.cs.value, desc: l.os.value === '1' })) })
        this.focus()
        if (!keys) return
        this.push([{ op: 'sort', sheet: this.sheet.name, range: `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`, keys, header: hdr.checked }], { now: true })
      }

      async dedupeDialog () {
        if (!this.doc.writable) return
        const rg = this.target()
        if (rg.r2 <= rg.r1) { this.frame.toast('请先选中要处理的数据。', 'info'); return }
        const { header, names } = this.headerNames(rg)
        const v = await dialog({
          title: '删除重复项',
          fields: [{ key: 'header', label: '数据包含标题', type: 'checkbox', value: header }, ...names.map(([c, l]) => ({ key: 'c' + c, label: l, type: 'checkbox', value: true }))],
          ok: '删除重复项',
          check: (x) => names.some(([c]) => x['c' + c]) ? '' : '至少选一列'
        })
        this.focus()
        if (!v) return
        const cols = names.filter(([c]) => v['c' + c]).map(([c]) => c)
        const r1 = v.header ? rg.r1 + 1 : rg.r1
        const seen = new Set(); let n = 0
        for (let r = r1; r <= Math.min(rg.r2, this.sheet.maxR); r++) { const k = JSON.stringify(cols.map((c) => String(this.shown(r, c)).toLowerCase())); if (seen.has(k)) n++; else seen.add(k) }
        if (!n) { this.frame.toast('没有发现重复值。', 'info'); return }
        this.push([{ op: 'dedupe', sheet: this.sheet.name, range: `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`, cols, header: v.header }], { now: true })
        this.frame.toast(`删除了 ${n} 个重复行，保留 ${seen.size} 个唯一值。`, 'info')
      }

      async splitColumn () {
        if (!this.doc.writable) return
        const s = this.sel
        if (s.c1 !== s.c2) { this.frame.toast('分列一次只能处理一列。', 'info'); return }
        const r2 = Math.min(s.r2, Math.max(this.sheet.maxR, s.r1))
        const v = await dialog({
          title: `分列 · ${colName(s.c1)} 列`,
          fields: [
            { key: 'sep', label: '分隔符', type: 'select', value: ',', options: [[',', '逗号'], ['\t', '制表符'], [' ', '空格'], [';', '分号'], ['、', '顿号'], ['other', '其他']] },
            { key: 'other', label: '其他', value: '', show: (x) => x.sep === 'other' },
            { key: 'merge', label: '连续分隔符视为一个', type: 'checkbox', value: true }
          ],
          check: (x) => x.sep === 'other' && !x.other ? '请填写分隔符' : ''
        })
        this.focus()
        if (!v) return
        const sep = v.sep === 'other' ? v.other : v.sep
        const rows = []
        let width = 1
        for (let r = s.r1; r <= r2; r++) {
          const t = String(this.input(r, s.c1))
          let parts = t.startsWith('=') ? [t] : t.split(sep === ',' ? /[,，]/ : sep)
          if (v.merge) parts = parts.filter((x, i) => x !== '' || i === 0)
          parts = parts.map((x) => x.trim())
          width = Math.max(width, parts.length)
          rows.push(parts)
        }
        if (width < 2) { this.frame.toast('没有找到分隔符。', 'info'); return }
        let busy = false
        for (let r = s.r1; r <= r2 && !busy; r++) for (let c = s.c1 + 1; c < s.c1 + width && !busy; c++) busy = this.shown(r, c) !== ''
        if (busy && !window.confirm('右侧单元格已有内容，分列会覆盖它们。继续吗？')) return
        const over = []
        rows.forEach((row, i) => row.forEach((x, j) => over.push([s.r1 + i, s.c1 + j, x])))
        this.push([{ op: 'inputs', sheet: this.sheet.name, at: cellName(s.r1, s.c1), rows }], { overlay: over, now: true })
      }

      outline (delta) {
        if (!this.doc.writable) return
        const s = this.sel
        const cols = s.r1 === 1 && s.r2 >= this.nr && !(s.c1 === 1 && s.c2 >= this.nc)
        if (cols) this.push([{ op: 'outline', sheet: this.sheet.name, axis: 'c', from: s.c1, to: Math.min(s.c2, 16384), delta }], { now: true })
        else this.push([{ op: 'outline', sheet: this.sheet.name, axis: 'r', from: s.r1, to: Math.min(s.r2, Math.max(this.sheet.maxR, s.r1)), delta }], { now: true })
      }

      // detail hides or shows the group around the selection.
      detail (hide) {
        if (!this.doc.writable) return
        const s = this.sel
        const lv = this.sheet.lv ?? {}
        const cols = s.r1 === 1 && s.r2 >= this.nr && !(s.c1 === 1 && s.c2 >= this.nc)
        const map = cols ? lv.c ?? {} : lv.r ?? {}
        let a = cols ? s.c1 : s.r1; let b = cols ? s.c2 : s.r2
        const level = map[a] ?? map[b] ?? map[a - 1] ?? map[b + 1]
        if (!level) { this.frame.toast('所选位置没有组合。先用“组合”把明细行或列组合起来。', 'info'); return }
        while (map[a - 1] >= level) a--
        while (map[b + 1] >= level) b++
        this.push([{ op: 'hide', sheet: this.sheet.name, axis: cols ? 'c' : 'r', from: a, to: b, hidden: hide }], { now: true })
      }

      // ------------------------------------------------ protection and page

      allowed (ops) {
        for (const o of ops) {
          if (o.op === 'protect' || o.op === 'unprotect') continue
          if (o.op === 'input' || o.op === 'inputs') {
            const at = rangeList(o.cell ?? o.at)[0]
            const rows = o.op === 'input' ? [[o.text]] : o.rows
            for (let i = 0; i < rows.length; i++) for (let j = 0; j < rows[i].length; j++) if (!this.styleOf(at.r1 + i, at.c1 + j).unlocked) return false
            continue
          }
          if (o.op === 'clear') {
            const g = rangeList(o.range)[0]
            for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 500); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 100); c++) if (this.shown(r, c) !== '' && !this.styleOf(r, c).unlocked) return false
            continue
          }
          return false
        }
        return true
      }

      async toggleProtect () {
        if (!this.doc.writable) return
        const on = !!this.sheet.protect
        const v = await dialog({ title: on ? '撤销工作表保护' : '保护工作表', fields: [{ key: 'pw', label: '密码', type: 'password', value: '', placeholder: on ? '没有密码就留空' : '可不填', hint: on ? '' : '只有未锁定的单元格可以修改。先选中要放开的单元格，在“审阅”里取消“锁定单元格”。' }, ...(on ? [] : [{ key: 'pw2', label: '确认密码', type: 'password', value: '' }])], check: (x) => !on && x.pw !== x.pw2 ? '两次输入的密码不一样' : '' })
        this.focus()
        if (!v) return
        this.push([{ op: on ? 'unprotect' : 'protect', sheet: this.sheet.name, password: v.pw || undefined }], { now: true })
      }

      page (props) {
        if (!this.doc.writable) return
        this.push([{ op: 'page', sheet: this.sheet.name, ...props }], { now: true })
      }

      async printTitles () {
        if (!this.doc.writable) return
        const cur = this.sheet.page?.titles?.replace(/^.*!/, '') ?? ''
        const v = await dialog({ title: '打印标题', fields: [{ key: 't', label: '顶端标题行', value: cur || '$1:$1', placeholder: '$1:$1', hint: '这些行会出现在打印的每一页顶端；留空则取消' }] })
        this.focus()
        if (!v) return
        this.page({ titles: v.t.trim() || null })
      }

      // ------------------------------------------------ find

      findAll (q) {
        const ql = q.toLowerCase()
        const out = []
        const keys = Object.keys(this.sheet.cells ?? {}).map((k) => k.split(',').map(Number)).sort((a, b) => a[0] - b[0] || a[1] - b[1])
        for (const [r, c] of keys) {
          if (!this.h(r) || !this.w(c)) continue
          if (String(this.shown(r, c)).toLowerCase().includes(ql) || String(this.input(r, c)).toLowerCase().includes(ql)) out.push({ sheet: this.sheetIx, r, c })
        }
        return out
      }

      showMatch (m) {
        if (m.sheet !== this.sheetIx) this.useSheet(m.sheet, false)
        this.select(m.r, m.c, m.r, m.c)
      }

      replaceMatch (m, q, rep) {
        const cur = String(this.input(m.r, m.c))
        const next = cur.replace(new RegExp(reEsc(q), 'gi'), () => rep)
        if (next !== cur) this.push([{ op: 'input', sheet: this.sheet.name, cell: cellName(m.r, m.c), text: next }], { overlay: [[m.r, m.c, next]], now: true })
      }

      replaceAll (q, rep) {
        const ops = []; const over = []
        const re = new RegExp(reEsc(q), 'gi')
        for (const m of this.findAll(q)) {
          const cur = String(this.input(m.r, m.c))
          const next = cur.replace(re, () => rep)
          if (next !== cur) { ops.push({ op: 'input', sheet: this.sheet.name, cell: cellName(m.r, m.c), text: next }); over.push([m.r, m.c, next]) }
        }
        if (ops.length) this.push(ops, { overlay: over, now: true })
        return ops.length
      }

      selectedText () { return String(this.shown(this.sel.ar, this.sel.ac)) }

      // ------------------------------------------------ saving

      push (ops, { overlay = [], now, switchTo } = {}) {
        if (!this.doc.writable) return
        if (this.sheet.protect && !this.allowed(ops)) { this.frame.toast('工作表受保护，不能这样修改。到“审阅”里撤销保护后再改。', 'warn'); return }
        const seq = ++this.seq
        for (const [r, c, text] of overlay) this.overlay.set(`${this.sheet.name}|${r},${c}`, { text, seq })
        for (const op of ops) this.queue.push({ op, seq })
        if (switchTo) this.switchTo = switchTo
        this.draw()
        this.refreshBar()
        this.doc.dirty(now || ops.some((o) => o.op !== 'input' && o.op !== 'inputs' && o.op !== 'clear'))
      }

      hasChanges () { return this.queue.length > 0 }

      collect () {
        if (!this.queue.length) return null
        const items = this.queue.splice(0)
        // Ops pushed together are one action, one undo step.
        const steps = []
        let last
        for (const x of items) { if (x.seq === last) steps[steps.length - 1]++; else { steps.push(1); last = x.seq } }
        return { ops: items.map((x) => x.op), seq: items[items.length - 1].seq, steps }
      }

      shortcuts () {
        return Object.assign([['Enter / Tab', '确认并移到下方 / 右侧单元格'], ['F2', '编辑单元格'], ['⇧F2', '编辑批注'], ['⇧F3', '插入函数'], ['Delete', '清除内容'], ['⌘C / ⌘X / ⌘V', '复制 / 剪切 / 粘贴'], ['⌘D', '向下填充'], ['⌘A', '全选'], ['⌘B / ⌘I / ⌘U', '加粗 / 倾斜 / 下划线'], ['⌘Home', '回到 A1'], ['PageUp / PageDown', '翻页'], ['⇧F10', '单元格菜单'], ['⌘滚轮', '缩放'], ['Esc', '取消编辑']], { title: '表格' })
      }

      saved (model, batch) {
        this.chartBusy = false
        for (const [k, o] of [...this.overlay]) if (o.seq <= batch.seq) this.overlay.delete(k)
        const name = this.switchTo
        this.switchTo = undefined
        const sel = this.sel
        const top = this.scroll.scrollTop; const left = this.scroll.scrollLeft
        this.model = model
        let ix = model.sheets.findIndex((s) => s.name === (name ?? this.sheet?.name))
        if (ix < 0) ix = Math.min(this.sheetIx, model.sheets.length - 1)
        const same = ix === this.sheetIx && !name
        this.useSheet(ix, same)
        if (same) { this.sel = sel; this.scroll.scrollTop = top; this.scroll.scrollLeft = left }
        this.draw(); this.placeInput(); this.refreshBar()
      }

      failed () { this.overlay.clear(); this.chartBusy = false; this.newChart = false; this.buildCharts(); this.draw() }
    }

    // ---------------------------------------------------------- document

    const RUN_KEYS = ['b', 'i', 'u', 's', 'color', 'size', 'highlight', 'vert', 'link', 'font', 'rev']
    const LINE_OPTS = [{ value: '', label: '行距' }, ...[[240, '1.0'], [276, '1.15'], [360, '1.5'], [480, '2.0'], [600, '2.5'], [720, '3.0']].map(([v, l]) => ({ value: String(v), label: l + ' 倍' }))]
    const SPACE_OPTS = [{ value: '', label: '段间距' }, { value: '0,0', label: '无间距' }, { value: '0,120', label: '段后 6 磅' }, { value: '0,240', label: '段后 12 磅' }, { value: '120,120', label: '前后各 6 磅' }, { value: '240,240', label: '前后各 12 磅' }]
    const PAGE_LABEL = { A4: 'A4', A3: 'A3', A5: 'A5', B5: 'B5', Letter: 'Letter', Legal: 'Legal' }
    const MARGIN_LABEL = { normal: '普通', narrow: '窄', moderate: '适中', wide: '宽' }
    const TRI = { b: 'b', i: 'i', u: 'u', s: 's' }
    const ALIGN_CSS = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }
    const BUILTIN_STYLES = [
      { value: '', label: '正文' },
      { value: '@h1', label: '标题 1' },
      { value: '@h2', label: '标题 2' },
      { value: '@h3', label: '标题 3' },
      { value: '@quote', label: '引用' },
      { value: '@code', label: '代码' }
    ]
    const HEADING_LOOK = { 1: { b: true, size: 32 }, 2: { b: true, size: 28 }, 3: { b: true, size: 24 }, 4: { b: true, size: 22 }, 5: { b: true, size: 21 }, 6: { b: true, size: 21 } }

    class DocxEditor {
      constructor (frame, doc) {
        this.frame = frame
        this.doc = doc
        this.k = 0
        this.base = new Map()
        this.dirtyBlocks = new Set()
        this.dirtyCells = new Set()
        this.struct = []
        this.extra = []
        this.zoom = 0
        this.touched = false
        this.build()
      }

      build () {
        const tb = new Toolbar(this.frame.tools, this.frame.strip)
        fileTab(tb, this.frame)
        this.ui = {}
        tb.tab('开始')
        this.ui.style = tb.select('段落样式', BUILTIN_STYLES, (v) => this.paraProps((d) => this.setStyle(d, v)), { width: '84px' })
        this.ui.font = tb.select('字体', fontOptions(), (v) => this.format((s) => { if (v) s.dataset.font = v; else delete s.dataset.font }), { width: '86px' })
        this.ui.size = tb.select('字号', [{ value: '', label: '字号' }, ...[8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72].map((n) => ({ value: String(n), label: String(n) }))], (v) => { if (v) this.format((s) => { s.dataset.size = String(Math.round(+v * 2)) }) }, { width: '58px' })
        tb.sep()
        this.ui.b = tb.button('B', '加粗（⌘B）', () => this.toggle('b'), { cls: 'q-b' })
        this.ui.i = tb.button('I', '倾斜（⌘I）', () => this.toggle('i'), { cls: 'q-i' })
        this.ui.u = tb.button('U', '下划线（⌘U）', () => this.toggle('u'), { cls: 'q-u' })
        this.ui.s = tb.button('S', '删除线', () => this.toggle('s'), { cls: 'q-s' })
        this.ui.sup = tb.button('', '上标', () => this.vert('sup'), { ico: 'sup' })
        this.ui.sub = tb.button('', '下标', () => this.vert('sub'), { ico: 'sub' })
        tb.color('文字颜色', 'A', (c) => this.format((s) => { s.dataset.color = c }), 'C00000')
        tb.color('突出显示', '▮', (c) => this.format((s) => { s.dataset.hl = c }), 'FFFF00')
        tb.button('', '清除格式', () => this.format((s) => { for (const k of ['b', 'i', 'u', 's', 'color', 'size', 'hl', 'vert', 'font']) delete s.dataset[k] }), { ico: 'trash', cls: 'q-mini' })
        tb.sep()
        this.ui.al = tb.button('', '左对齐', () => this.paraProps((d) => this.setAlign(d, 'left')), { ico: 'left' })
        this.ui.ac = tb.button('', '居中', () => this.paraProps((d) => this.setAlign(d, 'center')), { ico: 'center' })
        this.ui.ar = tb.button('', '右对齐', () => this.paraProps((d) => this.setAlign(d, 'right')), { ico: 'right' })
        this.ui.aj = tb.button('', '两端对齐', () => this.paraProps((d) => this.setAlign(d, 'both')), { ico: 'justify' })
        this.ui.bul = tb.button('', '项目符号', () => this.listToggle(false), { ico: 'bullet' })
        this.ui.num = tb.button('', '编号', () => this.listToggle(true), { ico: 'number' })
        this.ui.listStyle = tb.button('▾', '列表样式：项目符号、编号和多级列表', (e) => this.listMenu(e), { cls: 'q-mini' })
        tb.button('', '减少缩进', () => this.indent(-1), { ico: 'outdent' })
        tb.button('', '增加缩进', () => this.indent(1), { ico: 'indent' })
        this.ui.fl = tb.button('', '首行缩进 2 字符', () => this.firstLine(), { ico: 'firstLine' })
        this.ui.line = tb.select('行距', LINE_OPTS, (v) => this.paraProps((d) => { if (v) d.dataset.line = v; else delete d.dataset.line; this.dress(d) }), { width: '62px' })
        this.ui.space = tb.select('段间距', SPACE_OPTS, (v) => { if (!v) return; const [b, a] = v.split(','); this.paraProps((d) => { d.dataset.before = b; d.dataset.after = a; this.dress(d) }) }, { width: '70px' })
        this.ui.pborder = tb.button('边框 ▾', '段落边框', (e) => this.paraBorderMenu(e))
        this.ui.pshade = tb.button('底纹 ▾', '段落底纹', (e) => this.paraShadeMenu(e))
        tb.sep()
        this.ui.painter = tb.button('格式刷', '格式刷：单击刷一次，双击连续刷，Esc 退出', () => this.startPainter(false))
        this.ui.painter.addEventListener('dblclick', () => this.startPainter(true))
        tb.tab('插入')
        tb.button('图片', '插入图片', () => this.pickPicture(), { ico: 'image' })
        tb.button('表格', '插入表格（3×3）', () => this.insertTable(), { ico: 'table' })
        tb.button('分页符', '插入分页符', () => this.pageBreak(), { ico: 'pagebreak' })
        tb.sep()
        this.ui.link = tb.button('链接', '链接（⌘K）', () => this.linkSel(), { ico: 'link' })
        this.ui.toc = tb.button('目录', '插入目录；已有目录时更新或删除', (e) => this.tocMenu(e))
        this.ui.pageNo = tb.button('#', '页码（页脚居中）', () => this.togglePageNumber())
        this.ui.cmIns = tb.button('批注', '插入批注：先选中文字', () => this.newComment(), { ico: 'commentAdd' })
        this.ui.dropCap = tb.button('首字下沉 ▾', '首字下沉', (e) => this.dropMenu(e))
        tb.sep()
        this.ui.fnIns = tb.button('脚注', '插入脚注（⌘⌥F）', () => this.insertNote('fn'))
        this.ui.enIns = tb.button('尾注', '插入尾注（⌘⌥D）', () => this.insertNote('en'))
        tb.tab('表格')
        this.ui.rowIns = tb.button('', '在下方插入行', () => this.tableRow('below'), { ico: 'rowIns' })
        this.ui.rowDel = tb.button('', '删除所在行', () => this.tableRow('delete'), { ico: 'rowDel' })
        this.ui.colInsL = tb.button('', '在左侧插入列', () => this.tableCol('left'), { ico: 'colInsL' })
        this.ui.colInsR = tb.button('', '在右侧插入列', () => this.tableCol('right'), { ico: 'colInsR' })
        this.ui.colDel = tb.button('', '删除所在列', () => this.tableCol('delete'), { ico: 'colDel' })
        this.ui.tblDel = tb.button('', '删除表格', () => this.deleteTable(), { ico: 'trash' })
        tb.sep()
        this.ui.merge = tb.button('合并', '合并单元格：先拖选多个单元格', () => this.mergeCells())
        this.ui.split = tb.button('拆分', '拆分单元格；合并过的单元格取消合并', () => this.splitCells())
        this.ui.shade = tb.button('底纹 ▾', '单元格底纹', (e) => this.shadeMenu(e))
        this.ui.border = tb.button('边框 ▾', '单元格边框', (e) => this.borderMenu(e))
        this.ui.calign = tb.button('对齐 ▾', '单元格内文字对齐', (e) => this.cellAlignMenu(e))
        this.ui.tstyle = tb.button('样式 ▾', '表格样式', (e) => this.tableStyleMenu(e))
        this.ui.even = tb.button('平均列宽', '平均分布各列', () => this.evenColumns())
        this.tableHint = el('span', { class: 'q-hint' }, '把光标放进表格后可用')
        tb.root.append(this.tableHint)
        tb.tab('布局')
        this.ui.pageSetup = tb.button('页面', '页面设置：纸张、方向、页边距', (e) => this.pageMenu(e))
        this.ui.paper = tb.button('纸张', '纸张大小', (e) => this.pageMenu(e, 'size'))
        this.ui.orient = tb.button('方向', '纸张方向', (e) => this.pageMenu(e, 'orient'))
        this.ui.margins = tb.button('页边距', '页边距', (e) => this.pageMenu(e, 'margin'))
        tb.sep()
        tb.button('分隔符', '分隔符：插入分页符', () => this.pageBreak())
        this.ui.cols = tb.button('分栏 ▾', '分栏', (e) => this.pageMenu(e, 'cols'))
        tb.tab('审阅')
        this.ui.cmNew = tb.button('新建批注', '新建批注（⌘⌥M）：先选中文字', () => this.newComment(), { ico: 'commentAdd' })
        this.paneBtn = tb.button('批注栏', '批注栏', () => { this.paneOpen = !this.paneOpen; this.renderPane() }, { ico: 'comment' })
        tb.sep()
        this.ui.track = tb.button('修订', '修订：打开后，所做的修改会记录为修订', () => this.toggleTrack())
        this.ui.accept = tb.button('接受 ▾', '接受修订', (e) => this.revMenu(e, 'accept'))
        this.ui.reject = tb.button('拒绝 ▾', '拒绝修订', (e) => this.revMenu(e, 'reject'))
        tb.button('上一处', '上一处修订', () => this.revJump(-1))
        tb.button('下一处', '下一处修订', () => this.revJump(1))
        tb.tab('视图')
        this.zoomSel = tb.select('缩放', [{ value: '', label: '适应宽度' }, ...[50, 75, 100, 125, 150, 200].map((n) => ({ value: String(n), label: n + '%' }))], (v) => { this.zoom = v ? +v / 100 : 0; this.resize() }, { width: '78px' })
        this.paneBtn2 = tb.button('批注栏', '显示或隐藏批注栏', () => { this.paneOpen = !this.paneOpen; this.renderPane() })
        this.navBtn = tb.button('导航窗格', '导航窗格：按标题浏览和跳转', () => this.toggleNav(), { ico: 'nav' })
        tb.restore('docx')
        this.buttons = Object.values(this.ui)

        this.page = el('div', { class: 'qd-page', spellcheck: 'false' })
        this.hdr = el('div', { class: 'qd-hf qd-hdr', spellcheck: 'false', 'data-ph': '页眉' })
        this.ftr = el('div', { class: 'qd-hf qd-ftr', spellcheck: 'false', 'data-ph': '页脚' })
        this.sheets = el('div', { class: 'qd-pages', 'aria-hidden': 'true' })
        this.notesBox = el('div', { class: 'qd-notes', hidden: true })
        this.sheet = el('div', { class: 'qd-sheet' }, this.sheets, this.hdr, this.page, this.notesBox, this.ftr)
        this.sheets.addEventListener('dblclick', (e) => {
          const c = e.target.closest('.qd-hfc')
          if (!c) return
          const box = c.dataset.kind === 'header' ? this.hdr : this.ftr
          box.scrollIntoView({ block: 'center' })
          box.focus()
        })
        this.zoomer = el('div', { class: 'qd-zoom' }, this.sheet)
        for (const [box, kind] of [[this.hdr, 'header'], [this.ftr, 'footer']]) {
          box.addEventListener('blur', () => this.hfDone(box, kind))
          box.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); box.textContent = this.model[kind].text; box.blur() } })
          box.addEventListener('paste', (e) => { e.preventDefault(); document.execCommand('insertText', false, (e.clipboardData.getData('text/plain') || '').replace(/\r\n?/g, '\n')) })
        }
        this.scroll = el('div', { class: 'qd-scroll' }, this.zoomer)
        this.cpane = el('div', { class: 'qd-cpane', hidden: true })
        this.statusText = el('span', { class: 'qd-stxt' })
        this.zoomer = zoomSlider((v) => this.setZoom(v), () => this.setZoom(0))
        this.navSb = el('button', { type: 'button', class: 'q-sbtn', title: '导航窗格', 'aria-label': '导航窗格', html: icon('nav') + '<span>导航</span>' })
        this.navSb.addEventListener('click', () => this.toggleNav())
        this.statusBar = el('div', { class: 'qd-status' }, this.navSb, this.statusText, el('span', { class: 'q-space' }), this.zoomer.node)
        this.navList = el('div', { class: 'qd-nav-l', role: 'tree', 'aria-label': '标题' })
        this.navCount = el('span', { class: 'qd-nav-n' })
        const navX = el('button', { type: 'button', class: 'qd-nav-x', title: '关闭导航窗格', 'aria-label': '关闭导航窗格' }, '✕')
        navX.addEventListener('click', () => this.toggleNav(false))
        this.nav = el('nav', { class: 'qd-nav', hidden: true, 'aria-label': '导航窗格' }, el('div', { class: 'qd-nav-h' }, '导航', this.navCount, navX), this.navList)
        this.navShut = new Set()
        this.frame.main.append(el('div', { class: 'qd-body' }, this.nav, this.scroll, this.cpane), this.statusBar)
        this.scroll.addEventListener('scroll', () => this.status())

        this.page.addEventListener('beforeinput', (e) => this.beforeInput(e))
        this.page.addEventListener('keydown', (e) => this.key(e))
        this.page.addEventListener('paste', (e) => this.paste(e))
        this.page.addEventListener('cut', (e) => this.cut(e))
        this.page.addEventListener('dragstart', (e) => e.preventDefault())
        this.page.addEventListener('dragover', (e) => { if (imageFile(e.dataTransfer) || [...(e.dataTransfer?.items ?? [])].some((i) => i.kind === 'file')) e.preventDefault() })
        this.page.addEventListener('drop', (e) => this.drop(e))
        this.page.addEventListener('contextmenu', (e) => this.context(e))
        this.page.addEventListener('mousedown', (e) => this.tableDown(e))
        this.page.addEventListener('mousemove', (e) => this.tableHover(e))
        this.page.addEventListener('compositionstart', () => { this.composing = true })
        this.page.addEventListener('compositionend', () => { this.composing = false; this.changed() })
        this.page.addEventListener('input', () => { this.adoptStray(); this.changed() })
        this.page.addEventListener('click', (e) => {
          const link = e.target.closest?.('[data-link]')
          if (link && mod(e) && /^https?:/i.test(link.dataset.link)) window.open(link.dataset.link, '_blank', 'noopener')
          const cm = e.target.closest?.('.qd-cm')
          if (cm) { this.paneOpen = true; this.activate(cm.dataset.id) }
          const toc = e.target.closest?.('.qd-p[data-toc]')
          if (toc && +toc.dataset.toc > 0) this.jumpToHeading(toc.textContent.replace(/\t[\s\S]*$/, ''))
          const fn = e.target.closest?.('.qd-fn')
          if (fn) { const box = this.notesBox.querySelector(`[data-note="${fn.dataset.note}"][data-id="${fn.dataset.id}"] .qd-note-t`); if (box) { box.scrollIntoView({ block: 'center' }); box.focus() } }
          if (e.target.tagName === 'IMG' && e.target.classList.contains('qd-img')) this.pickImg(e.target); else this.dropImg()
        })
        this.scroll.addEventListener('scroll', () => this.placeImgHandle())
        this.onSel = () => { this.refreshBar(); this.status(); this.freshenRev() }
        document.addEventListener('selectionchange', this.onSel)
        this.mo = new MutationObserver((recs) => this.observe(recs))
        this.ro = new ResizeObserver(() => this.resize())
        this.ro.observe(this.scroll)
      }

      dispose () {
        this.dropImg()
        document.removeEventListener('selectionchange', this.onSel)
        window.CSS.highlights?.delete?.('q-comment')
        window.CSS.highlights?.delete?.('q-comment-on')
        pinRanges.delete(this)
        paintPinRanges()
        this.mo.disconnect()
        this.ro.disconnect()
      }

      // ------------------------------------------------ rendering

      load (model, { keep } = {}) {
        const top = this.scroll.scrollTop
        this.model = model
        this.mo.disconnect()
        const pg = model.page
        this.sheet.style.width = pg.width + 'px'
        Object.assign(this.page.style, { padding: `0 ${pg.right}px 0 ${pg.left}px`, marginTop: pg.top + 'px', minHeight: Math.max(100, pg.height - pg.top - pg.bottom) + 'px' })
        const cols = pg.cols > 1 ? pg.cols : 0
        Object.assign(this.page.style, { display: cols ? 'block' : '', columnCount: cols ? String(cols) : '', columnGap: cols ? (pg.colSpace ?? 28) + 'px' : '' })
        for (const box of [this.hdr, this.ftr]) Object.assign(box.style, { paddingLeft: pg.left + 'px', paddingRight: pg.right + 'px' })
        Object.assign(this.hdr.style, { height: pg.top + 'px', top: '0px' })
        Object.assign(this.ftr.style, { height: pg.bottom + 'px', top: pg.height - pg.bottom + 'px' })
        const def = model.looks[''] ?? {}
        this.sheet.style.fontSize = (def.size ? def.size / 2 : 10.5) + 'pt'
        this.showHf()
        this.page.replaceChildren()
        this.base = new Map()
        for (const b of model.blocks) {
          const node = b.kind === 'p' ? this.para(b) : this.table(b)
          this.page.append(node)
        }
        if (!model.blocks.length) this.page.append(this.para({ kind: 'p', runs: [] }))
        this.renumber()
        model.blocks.forEach((b, i) => {
          const node = this.page.children[i]
          this.base.set(node.dataset.k, this.baseline(node, b.ref))
        })
        this.dirtyBlocks.clear()
        this.dirtyCells.clear()
        this.touched = false
        this.page.contentEditable = this.doc.writable ? 'true' : 'false'
        const styles = this.ui.style
        styles.replaceChildren(...BUILTIN_STYLES.map((o) => el('option', { value: o.value }, o.label)))
        const named = new Set(['标题 1', '标题 2', '标题 3', 'heading 1', 'heading 2', 'heading 3', 'normal', 'quote', 'title'])
        for (const s of model.styles ?? []) if (!named.has(String(s.name).toLowerCase())) styles.append(el('option', { value: s.id }, s.name || s.id))
        for (const b of this.buttons) b.disabled = !this.doc.writable
        this.mo.observe(this.page, { childList: true, characterData: true, subtree: true, attributes: true, attributeFilter: ['data-b', 'data-i', 'data-u', 'data-s', 'data-color', 'data-size', 'data-hl', 'data-vert', 'data-link', 'data-font', 'data-style', 'data-align', 'data-list', 'data-lv', 'data-line', 'data-before', 'data-after', 'data-ind', 'data-fl', 'data-bdr', 'data-shade', 'data-num'] })
        this.resize()
        if (keep) this.scroll.scrollTop = top
        if (this.paneOpen === undefined) this.paneOpen = (model.comments ?? []).some((c) => !c.done)
        if (this.draft && !keep) this.draft = null
        this.cmEdit = null
        this.ui.toc.textContent = model.toc ? '目录 ▾' : '目录'
        this.revStale = null
        this.dropImg()
        this.renderNotes()
        this.paginate()
        for (const img of this.page.querySelectorAll('img')) if (!img.complete) img.addEventListener('load', () => this.later(), { once: true })
        document.fonts?.ready?.then(() => this.later())
        this.paintComments()
        this.renderPane()
        this.refreshBar()
      }

      nextK () { return 'k' + (++this.k) }

      para (b) {
        const d = el('div', { class: 'qd-p', dataset: { k: this.nextK() } })
        if (b.style) d.dataset.style = b.style
        if (b.heading) d.dataset.h = String(b.heading)
        if (b.align) d.dataset.align = b.align
        if (b.list) { d.dataset.list = b.list.ordered ? 'o' : 'b'; d.dataset.lv = String(b.list.level) }
        if (b.spacing?.line) d.dataset.line = String(b.spacing.line)
        if (b.spacing?.before !== undefined) d.dataset.before = String(b.spacing.before)
        if (b.spacing?.after !== undefined) d.dataset.after = String(b.spacing.after)
        if (b.indent?.left !== undefined) d.dataset.ind = String(b.indent.left)
        if (b.indent?.firstLine !== undefined) d.dataset.fl = String(b.indent.firstLine)
        if (b.list?.num) d.dataset.num = b.list.num
        if (b.border) d.dataset.bdr = b.border
        if (b.shade) d.dataset.shade = b.shade
        if (b.drop) { d.dataset.drop = String(b.drop); d.classList.add('qd-drop') }
        if (b.locked) { d.dataset.locked = '1'; d.contentEditable = 'false'; d.title = '这一段含域或修订，请在 Word/WPS 中编辑' }
        if (b.toc !== undefined) { d.dataset.toc = String(b.toc); d.title = b.toc ? '目录：点击跳到这一节；用工具栏“目录”更新' : '目录' }
        let pic = 0
        for (const r of b.runs ?? []) {
          if (r.t === 'text') d.append(this.run(r))
          else if (r.t === 'del') {
            const rv = this.model.revs?.[r.id]
            d.append(el('span', { class: 'qd-del', contenteditable: 'false', dataset: { del: String(r.id) }, title: rv ? `${rv.author || '未知作者'} 删除 · ${String(rv.date).slice(0, 16).replace('T', ' ')}` : '已删除' }, r.text))
          } else if (r.t === 'fn' || r.t === 'en') d.append(el('span', { class: 'qd-fn', contenteditable: 'false', dataset: { note: r.t, id: String(r.id) } }))
          else if (r.t === 'img') {
            const img = el('img', { class: 'qd-img', contenteditable: 'false', src: r.src || undefined, alt: r.src ? '' : '图片', dataset: { img: String(r.n ?? pic), from: b.ref, key: this.nextK() } })
            if (r.w) img.style.width = r.w + 'px'
            if (r.h) img.style.height = r.h + 'px'
            if (r.wrap) { img.dataset.wrap = r.wrap; img.dataset.align = r.align ?? 'left' }
            this.dressImg(img)
            pic++
            d.append(img)
          } else if (r.t === 'pb') d.append(el('span', { class: 'qd-pb', contenteditable: 'false', 'data-pb': '1' }))
          else if (r.t === 'cs' || r.t === 'ce' || r.t === 'cr') d.append(this.marker(r.t[1], r.id))
        }
        if (!d.childNodes.length) d.append(el('br'))
        // A contents entry with a page number: the tab before it is a
        // dotted leader.
        if (b.toc > 0) {
          const tab = [...d.querySelectorAll(':scope > .qd-r')].reverse().find((x) => x.textContent === '\t')
          if (tab) { tab.classList.add('qd-leader'); d.classList.add('qd-tocp') }
        }
        this.dress(d)
        return d
      }

      // marker stands for where a comment starts (s), ends (e) and is
      // referenced (r); the server keeps them in step with comments.xml.
      marker (kind, id) {
        return el('span', { class: 'qd-cm', contenteditable: 'false', title: kind === 'r' ? '批注' : undefined, dataset: { cm: kind, id: String(id) } })
      }

      run (r) {
        const s = el('span', { class: 'qd-r' }, r.text)
        for (const k of Object.keys(TRI)) if (r[k] !== undefined) s.dataset[k] = r[k] ? '1' : '0'
        if (r.color) s.dataset.color = r.color
        if (r.size) s.dataset.size = String(r.size)
        if (r.highlight) s.dataset.hl = r.highlight
        if (r.vert) s.dataset.vert = r.vert
        if (r.link) s.dataset.link = r.link
        if (r.font) s.dataset.font = r.font
        if (r.rev) {
          s.dataset.rev = r.rev
          const rv = this.model.revs?.[r.rev.slice(4)]
          s.title = rv ? `${rv.author || '未知作者'} 插入 · ${String(rv.date).slice(0, 16).replace('T', ' ')}` : '插入'
        }
        this.paint(s)
        return s
      }

      // paint shows a run's own properties.
      paint (s) {
        const d = s.dataset
        const st = s.style
        st.fontWeight = d.b === '1' ? '700' : d.b === '0' ? '400' : ''
        st.fontStyle = d.i === '1' ? 'italic' : d.i === '0' ? 'normal' : ''
        const deco = [d.u === '1' || d.link ? 'underline' : '', d.s === '1' ? 'line-through' : ''].filter(Boolean).join(' ')
        st.textDecoration = deco || (d.u === '0' || d.s === '0' ? 'none' : '')
        st.color = d.color ? '#' + d.color : d.link ? '#1155cc' : ''
        st.fontSize = d.size ? (+d.size / 2) + 'pt' : d.vert ? '0.7em' : ''
        st.background = d.hl ? '#' + d.hl : ''
        st.verticalAlign = d.vert === 'sup' ? 'super' : d.vert === 'sub' ? 'sub' : ''
        st.fontFamily = d.font ? fontCss(d.font) : ''
        if (d.rev) { st.textDecoration = [st.textDecoration === 'none' ? '' : st.textDecoration, 'underline'].filter(Boolean).join(' '); if (!d.color) st.color = '#c2185b' }
      }

      // dress shows a paragraph's style, alignment and list.
      dress (d) {
        const style = d.dataset.style
        const h = style?.startsWith('@h') ? +style.slice(2) : +d.dataset.h || 0
        const look = { ...(h ? HEADING_LOOK[h] : {}), ...(style && !style.startsWith('@') ? this.model.looks[style] : style ? {} : this.model.looks['']) }
        if (style === '@quote') Object.assign(look, { i: true, color: '595959', indent: 720 })
        const st = d.style
        st.fontWeight = look.b ? '700' : ''
        st.fontStyle = look.i ? 'italic' : ''
        st.color = look.color ? '#' + look.color : ''
        st.fontSize = look.size ? look.size / 2 + 'pt' : ''
        st.textAlign = ALIGN_CSS[d.dataset.align ?? look.align] ?? ''
        const ds = d.dataset
        st.setProperty('--mt', (ds.before !== undefined ? +ds.before / 20 : look.before !== undefined ? look.before / 20 : h ? 12 : 0) + 'pt')
        st.marginBottom = (ds.after !== undefined ? +ds.after / 20 : look.after !== undefined ? look.after / 20 : 6) + 'pt'
        st.lineHeight = ds.line ? String(+ds.line / 240 * 1.2) : ''
        st.textIndent = ds.fl !== undefined && !ds.list ? +ds.fl / 15 + 'px' : ''
        st.background = ds.shade ? '#' + ds.shade : look.shade ? '#' + look.shade : style === '@code' ? '#f4f4f4' : ''
        st.fontFamily = style === '@code' ? 'Menlo, Consolas, monospace' : ''
        const lv = +d.dataset.lv || 0
        const ind = ds.ind !== undefined ? +ds.ind / 15 : look.indent ? look.indent / 15 : 0
        const hang = ds.fl !== undefined && +ds.fl < 0 && !ds.list ? -ds.fl / 15 : 0
        const bd = ds.bdr ?? ''
        const line = '1px solid #000'
        st.borderTop = bd.includes('t') ? line : ''
        st.borderBottom = bd.includes('b') ? line : ''
        st.borderLeft = bd.includes('l') ? line : ''
        st.borderRight = bd.includes('r') ? line : ''
        st.paddingTop = bd.includes('t') ? '1px' : ''
        st.paddingBottom = bd.includes('b') ? '1px' : ''
        st.paddingRight = bd.includes('r') ? '4px' : ''
        const extra = bd.includes('l') ? 4 : 0
        const mw = Math.max(24, (+ds.mw || 20) + 4)
        st.setProperty('--mw', (+ds.mw || 20) + 'px')
        st.paddingLeft = d.dataset.list ? `${lv * 24 + mw + ind + extra}px` : ind || hang || extra ? Math.max(ind, hang) + extra + 'px' : ''
        if (ds.drop) { st.setProperty('--mt', '0pt'); st.marginBottom = '0' }
      }

      table (b) {
        const wrap = el('div', { class: 'qd-t', contenteditable: 'false', dataset: { k: this.nextK() } })
        const t = el('table')
        const width = Math.max(...b.rows.map((r) => r.reduce((n, c) => n + (c.span || 1), 0)), 1)
        const T = b.borders ?? {}
        const line = (v) => (!v ? '1px dashed #d5d9de' : v === 'none' ? '1px dashed #e3e6ea' : `${v.endsWith(':2') ? 2 : 1}px solid #${v.slice(0, 6)}`)
        const VA = { center: 'middle', bottom: 'bottom' }
        const HA = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }
        // Rows spanned by vertically merged cells.
        const grid = b.rows.map((r) => { let g = 0; return r.map((c) => { const at = g; g += c.span || 1; return at }) })
        b.rows.forEach((row, ri) => {
          const tr = el('tr')
          row.forEach((c, ci) => {
            if (c.vmerge === 'continue') return
            const g = c.g ?? grid[ri][ci]
            const td = el('td', { contenteditable: c.locked || !this.doc.writable ? 'false' : 'true', dataset: { r: String(ri + 1), c: String(ci + 1), g: String(g) } })
            if (c.span > 1) td.colSpan = c.span
            if (c.vmerge === 'restart') {
              let n = 1
              for (let k = ri + 1; k < b.rows.length; k++) { const j = grid[k].indexOf(grid[ri][ci]); if (j >= 0 && b.rows[k][j].vmerge === 'continue') n++; else break }
              td.rowSpan = n
            }
            if (c.fill && /^[0-9A-F]{6}$/i.test(c.fill)) td.style.background = '#' + c.fill
            if (b.header && ri === 0) td.classList.add('qd-th')
            const span = c.span || 1
            const last = ri + (td.rowSpan || 1) - 1 >= b.rows.length - 1
            const bd = c.bd ?? {}
            Object.assign(td.style, {
              borderTop: line(bd.top ?? (ri === 0 ? T.top : T.insideH)),
              borderBottom: line(bd.bottom ?? (last ? T.bottom : T.insideH)),
              borderLeft: line(bd.left ?? (g === 0 ? T.left : T.insideV)),
              borderRight: line(bd.right ?? (g + span >= width ? T.right : T.insideV)),
              verticalAlign: VA[c.valign] ?? '',
              textAlign: HA[c.align] ?? '',
              color: c.color ? '#' + c.color : '',
              fontWeight: c.b === true ? '700' : c.b === false ? '400' : ''
            })
            if (c.locked) td.title = '这个单元格含复杂内容，请在 Word/WPS 中编辑'
            td.textContent = c.text
            tr.append(td)
          })
          t.append(tr)
        })
        wrap.gridN = width
        wrap.append(t)
        if (b.widths?.length === width) {
          wrap.widths = [...b.widths]
          this.applyWidths(wrap)
        } else t.style.minWidth = Math.min(width * 60, 600) + 'px'
        return wrap
      }

      later () { clearTimeout(this.pageTimer); this.pageTimer = setTimeout(() => this.paginate(), 60) }

      // paginate lays the page out on sheets of paper: a block that would
      // cross the bottom margin moves to the next sheet (blocks taller
      // than a sheet stay where they are), a page break starts a new one,
      // and every sheet shows the header and footer.
      paginate () {
        const pg = this.model?.page
        if (!pg || !this.page.isConnected) return
        const GAP = 16
        const H = pg.height
        const contentH = Math.max(100, pg.height - pg.top - pg.bottom)
        const stride = H + GAP
        const all = [...this.page.children]
        for (const b of all) { b.style.removeProperty('--push'); b.classList.remove('qd-dropnext') }
        const k = this.page.offsetWidth / pg.width || 1
        // A drop cap floats at the start of the paragraph after it.
        const drops = all.filter((b) => b.dataset.drop)
        for (const d of drops) {
          const next = d.nextElementSibling
          if (!next?.classList.contains('qd-p')) continue
          next.classList.add('qd-dropnext')
          next.style.setProperty('--dw', d.offsetWidth / k + 4 + 'px')
          next.style.setProperty('--dh', d.offsetHeight / k + 'px')
        }
        const blocks = all.filter((b) => !b.dataset.drop)
        const multi = pg.cols > 1
        const boxes = blocks.map((b) => [b.offsetTop / k, b.offsetHeight / k])
        let shift = 0; let last = 0; let breakNext = false
        for (const [i, b] of blocks.entries()) {
          const h = boxes[i][1]
          const y = boxes[i][0] + shift
          const n = Math.floor(y / stride)
          const within = y - n * stride
          let push = 0
          if (multi) { last = Math.max(last, y + h); continue }
          if ((breakNext && within > 1) || within >= contentH || (within > 1 && within + h > contentH && h <= contentH)) push = (n + 1) * stride - y
          if (push > 0.5) { b.style.setProperty('--push', push + 'px'); shift += push }
          breakNext = !!b.querySelector?.('.qd-pb')
          last = y + push + h
        }
        let count = Math.max(1, Math.floor(Math.max(0, last - 1) / stride) + 1 + (breakNext ? 1 : 0))
        if (multi) count = Math.max(1, Math.ceil(last / contentH))
        for (const d of drops) {
          const next = d.nextElementSibling
          if (next?.classList.contains('qd-dropnext')) Object.assign(d.style, { top: next.offsetTop / k + 'px', left: pg.left + 'px' })
        }
        if (!this.notesBox.hidden) {
          const nh = this.notesBox.offsetHeight / k
          let top = (multi ? this.page.offsetHeight / k : last) + 18
          const n = Math.floor(top / stride)
          if (!multi && top - n * stride + nh > contentH && nh <= contentH) top = (n + 1) * stride
          Object.assign(this.notesBox.style, { top: pg.top + top + 'px', left: pg.left + 'px', right: pg.right + 'px' })
          count = Math.max(count, Math.floor((top + nh) / stride) + 1)
        }
        this.pageCount = count
        this.page.style.minHeight = multi ? '' : (count - 1) * stride + contentH + 'px'
        this.sheet.style.height = count * H + (count - 1) * GAP + 'px'
        const hf = (kind, i) => {
          const info = this.model[kind] ?? { text: '' }
          let t = info.text ?? ''
          if (info.page) t = t.replace('第  页', `第 ${i + 1} 页`).replace('共  页', `共 ${count} 页`)
          return t
        }
        const sheets = []
        for (let i = 0; i < count; i++) {
          const top = i * stride
          sheets.push(el('div', { class: 'qd-pg', style: { top: top + 'px', height: H + 'px' } }))
          if (i === 0) continue
          sheets.push(el('div', { class: 'qd-hfc', dataset: { kind: 'header' }, title: '双击编辑页眉', style: { top: top + 'px', height: pg.top + 'px', paddingLeft: pg.left + 'px', paddingRight: pg.right + 'px' } }, hf('header', i)))
          sheets.push(el('div', { class: 'qd-hfc qd-hfc-f', dataset: { kind: 'footer' }, title: '双击编辑页脚', style: { top: top + H - pg.bottom + 'px', height: pg.bottom + 'px', paddingLeft: pg.left + 'px', paddingRight: pg.right + 'px' } }, hf('footer', i)))
        }
        this.sheets.replaceChildren(...sheets)
        if (this.model.footer?.page && document.activeElement !== this.ftr) this.ftr.textContent = hf('footer', 0)
        this.countWords()
        this.paintPins()
        this.renderNav()
        this.status()
      }

      countWords () {
        const t = this.page.textContent.replace(/\u200b/g, '')
        const cjk = (t.match(/[\u3400-\u9fff\uf900-\ufaff\u3000-\u303f\uff00-\uffef]/g) ?? []).length
        const words = (t.replace(/[\u3400-\u9fff\uf900-\ufaff\u3000-\u303f\uff00-\uffef]/g, ' ').match(/[A-Za-z0-9]+(?:['’.-][A-Za-z0-9]+)*/g) ?? []).length
        this.words = cjk + words
      }

      shortcuts () {
        return Object.assign([['⌘B / ⌘I / ⌘U', '加粗 / 倾斜 / 下划线'], ['⌘K', '插入超链接'], ['⌥⌘M', '新建批注'], ['⌥⌘F', '插入脚注'], ['⌥⌘D', '插入尾注'], ['Tab / ⇧Tab', '列表升降级；表格中移到下一 / 上一格'], ['⌘A', '全选（表格内选中单元格）'], ['Esc', '退出格式刷']], { title: '文档' })
      }

      // status shows the page in view, the page count and the word count.
      status () {
        if (!this.model || !this.statusBar) return
        const pg = this.model.page
        const z = parseFloat(this.sheet.style.zoom) || 1
        const stride = (pg.height + 16) * z
        const cur = Math.min(this.pageCount ?? 1, Math.floor((this.scroll.scrollTop + this.scroll.clientHeight / 3) / stride) + 1)
        const range = currentRange(this.page)
        const sel = range && !range.collapsed ? range.toString().replace(/\s/g, '').length : 0
        this.statusText.textContent = `第 ${cur} 页，共 ${this.pageCount ?? 1} 页　　${sel ? `已选 ${sel} 字 / ` : ''}${this.words ?? 0} 字`
        this.navSpot()
      }

      // renumber shows list markers.
      // renumber shows list marks as Word numbers them: a list in the file
      // counts on through the document; a new one until text breaks it.
      renumber () {
        let nf = 0; let ne = 0
        for (const m of this.page.querySelectorAll('.qd-fn')) {
          const t = m.dataset.note === 'en' ? numText(++ne, 'lowerRoman') : String(++nf)
          if (m.textContent !== t) m.textContent = t
        }
        const all = new Map()
        for (const d of this.page.children) {
          if (!d.classList.contains('qd-p')) continue
          if (!d.dataset.list) { if (d.textContent.replace(/​/g, '').trim()) { all.delete('newo'); all.delete('newb') } delete d.dataset.mark; continue }
          const lv = +d.dataset.lv || 0
          const lvls = d.dataset.num ? this.model.numbering?.[d.dataset.num] : undefined
          const key = lvls ? 'n' + d.dataset.num : 'new' + d.dataset.list
          let c = all.get(key)
          if (!c) { c = []; all.set(key, c) }
          c.length = lv + 1
          c[lv] = (c[lv] ?? ((lvls?.[lv]?.start ?? 1) - 1)) + 1
          let mark
          const L = lvls?.[lv]
          if (L) {
            if (L.fmt === 'bullet') mark = L.text || '•'
            else if (L.fmt === 'none') mark = ''
            else mark = L.text.replace(/%(\d)/g, (_, k) => numText(c[k - 1] ?? lvls[k - 1]?.start ?? 1, lvls[k - 1]?.fmt ?? 'decimal'))
            if (L.fmt === 'bullet' && /[\uf000-\uf0ff]/.test(mark)) mark = ['•', '◦', '▪'][lv % 3]
          } else mark = d.dataset.list === 'o' ? `${c[lv]}.` : ['•', '◦', '▪'][lv % 3]
          if (d.dataset.mark !== mark) d.dataset.mark = mark
          const mw = String(Math.max(20, [...mark].reduce((a, ch) => a + (ch.charCodeAt(0) > 255 ? 13 : 7), 0) + 2))
          if (d.dataset.mw !== mw) { d.dataset.mw = mw; this.dress(d) }
        }
      }

      resize () {
        const w = this.scroll.clientWidth
        const pw = this.model?.page.width ?? 794
        const z = this.zoom || Math.min(1, Math.max(0.3, (w - 24) / pw))
        this.sheet.style.zoom = String(z)
        this.zoomer?.update(z * 100)
      }

      setZoom (v) { this.zoom = v ? v / 100 : 0; syncZoomSel(this.zoomSel, v); this.resize(); this.status() }

      // ------------------------------------------------ reading the page

      blockOf (n) {
        while (n && n.parentNode !== this.page) n = n.parentNode
        return n && n.nodeType === 1 ? n : undefined
      }

      cellOf (n) { return (n?.nodeType === 1 ? n : n?.parentElement)?.closest?.('td') ?? undefined }

      ser (d) {
        const runs = []
        const add = (text, props) => {
          text = text.replace(/​/g, '').replace(/ /g, ' ')
          if (!text) return
          const last = runs[runs.length - 1]
          if (last?.t === 'text' && RUN_KEYS.every((k) => last[k] === props[k])) last.text += text
          else runs.push({ t: 'text', text, ...props })
        }
        const walk = (node, props) => {
          for (const c of node.childNodes) {
            if (c.nodeType === 3) add(c.data, props)
            else if (c.nodeType !== 1) continue
            else if (c.tagName === 'IMG' && c.dataset.img !== undefined) runs.push({ t: 'img', n: +c.dataset.img, from: c.dataset.from, key: c.dataset.key, w: Math.round(parseFloat(c.style.width)) || 0, h: Math.round(parseFloat(c.style.height)) || 0, ...(c.dataset.wrap ? { wrap: c.dataset.wrap, align: c.dataset.align } : {}) })
            else if (c.hasAttribute('data-pb')) runs.push({ t: 'pb' })
            else if (c.classList.contains('qd-cm')) runs.push({ t: 'c' + c.dataset.cm, id: c.dataset.id })
            else if (c.classList.contains('qd-del')) runs.push({ t: 'del', id: c.dataset.del })
            else if (c.classList.contains('qd-fn')) runs.push({ t: c.dataset.note, id: c.dataset.id })
            else if (c.tagName === 'BR') continue
            else walk(c, this.propsOf(c, props))
          }
        }
        walk(d, {})
        const ds = d.dataset
        const num = (v) => (v === undefined ? null : +v)
        return {
          runs,
          style: ds.style ?? null,
          align: ds.align ?? null,
          list: ds.list ? { ordered: ds.list === 'o', level: +ds.lv || 0, ...(ds.num ? { numId: ds.num } : {}) } : null,
          spacing: { line: num(ds.line), before: num(ds.before), after: num(ds.after) },
          indent: { left: num(ds.ind), firstLine: num(ds.fl) },
          border: ds.bdr ?? null,
          shade: ds.shade ?? null
        }
      }

      propsOf (span, inherited) {
        const p = { ...inherited }
        const d = span.dataset
        for (const k of Object.keys(TRI)) if (d[k] !== undefined) p[k] = d[k] === '1'
        if (d.color) p.color = d.color
        if (d.size) p.size = +d.size
        if (d.hl) p.highlight = d.hl
        if (d.vert) p.vert = d.vert
        if (d.link) p.link = d.link
        if (d.font) p.font = d.font
        if (d.rev) p.rev = d.rev
        for (const k of RUN_KEYS) if (p[k] === undefined) delete p[k]
        return p
      }

      sig (s) { return JSON.stringify({ ...s, runs: s.runs.map((r) => (r.t === 'img' ? { t: 'img', key: r.key, w: r.w, h: r.h, wrap: r.wrap, align: r.align } : r)) }) }

      baseline (node, ref) {
        if (node.classList.contains('qd-t')) return { ref, kind: 't' }
        const s = this.ser(node)
        return { ref, kind: 'p', sig: this.sig(s), props: this.pprops(s) }
      }

      pprops (s) { return { style: s.style, align: s.align, list: s.list, spacing: s.spacing, indent: s.indent, border: s.border, shade: s.shade } }

      // pdiff is what changed in a paragraph's properties; spacing and
      // indents name only the parts that changed.
      pdiff (now, was) {
        const props = {}
        for (const key of ['style', 'align', 'list', 'border', 'shade']) if (JSON.stringify(now[key] ?? null) !== JSON.stringify(was[key] ?? null)) props[key] = now[key]
        for (const key of ['spacing', 'indent']) {
          const part = {}
          for (const [k, v] of Object.entries(now[key] ?? {})) if (v !== (was[key]?.[k] ?? null)) part[k] = v
          if (Object.keys(part).length) props[key] = part
        }
        return props
      }

      // adoptStray turns what the browser added on its own into runs the
      // editor understands.
      adoptStray () {
        for (const n of [...this.page.childNodes]) {
          if (n.nodeType === 1 && (n.classList.contains('qd-p') || n.classList.contains('qd-t'))) continue
          if (n.nodeType === 3 && !n.data.trim()) { n.remove(); continue }
          const d = this.para({ kind: 'p', runs: [] })
          d.replaceChildren()
          n.replaceWith(d)
          d.append(n)
          if (n.nodeType === 1 && n.tagName === 'DIV') n.replaceWith(...n.childNodes)
        }
        for (const s of this.page.querySelectorAll('.qd-p b, .qd-p strong, .qd-p i, .qd-p em, .qd-p u, .qd-p s, .qd-p strike, .qd-p font, .qd-p span:not(.qd-r):not(.qd-pb):not(.qd-cm)')) {
          const span = el('span', { class: 'qd-r' })
          const t = s.tagName
          if (t === 'B' || t === 'STRONG') span.dataset.b = '1'
          if (t === 'I' || t === 'EM') span.dataset.i = '1'
          if (t === 'U') span.dataset.u = '1'
          if (t === 'S' || t === 'STRIKE') span.dataset.s = '1'
          const cs = s.style
          if (cs.fontWeight) span.dataset.b = +cs.fontWeight >= 600 || cs.fontWeight === 'bold' ? '1' : '0'
          if (cs.fontStyle) span.dataset.i = cs.fontStyle === 'italic' ? '1' : '0'
          if (/underline/.test(cs.textDecoration || cs.textDecorationLine)) span.dataset.u = '1'
          if (/line-through/.test(cs.textDecoration || cs.textDecorationLine)) span.dataset.s = '1'
          const c = rgbHex(cs.color) ?? (s.getAttribute('color') || '').replace('#', '').toUpperCase()
          if (c && /^[0-9A-F]{6}$/.test(c)) span.dataset.color = c
          const bg = rgbHex(cs.backgroundColor)
          if (bg) span.dataset.hl = bg
          if (cs.fontSize && /px$/.test(cs.fontSize)) span.dataset.size = String(Math.round(parseFloat(cs.fontSize) * 1.5))
          if (cs.fontSize && /pt$/.test(cs.fontSize)) span.dataset.size = String(Math.round(parseFloat(cs.fontSize) * 2))
          span.append(...s.childNodes)
          s.replaceWith(span)
          this.paint(span)
        }
      }

      observe (recs) {
        for (const r of recs) {
          const cell = this.cellOf(r.target)
          if (cell) { this.dirtyCells.add(cell); continue }
          const b = this.blockOf(r.target)
          if (b) this.dirtyBlocks.add(b)
          if (r.target === this.page) for (const n of r.addedNodes) if (n.nodeType === 1) this.dirtyBlocks.add(n)
        }
      }

      changed () {
        if (!this.doc.writable) return
        this.touched = true
        clearTimeout(this.numTimer)
        this.numTimer = setTimeout(() => { this.renumber(); this.paintComments(); this.paginate() }, 120)
        this.doc.dirty()
      }

      // ------------------------------------------------ saving

      hasChanges () { return this.touched || this.struct.length > 0 || this.extra.length > 0 }

      collect () {
        if (this.composing) { this.composeWait = true; return null }
        this.observe(this.mo.takeRecords())
        this.renumber()
        const ops = []
        const snap = []
        const sigs = []
        const seen = new Set()
        let prev = null
        const orphans = []
        const inserts = []
        const revNodes = new Set()
        for (const node of [...this.page.children]) {
          let k = node.dataset.k
          if (k && seen.has(k)) { k = undefined; delete node.dataset.k }
          const b = k ? this.base.get(k) : undefined
          snap.push(node)
          if (b) {
            seen.add(k)
            for (const o of orphans.splice(0)) o.before = b.ref
            if (b.kind === 'p') {
              if (this.dirtyBlocks.has(node)) {
                const s = this.ser(node)
                const sig = this.sig(s)
                sigs.push({ sig, props: this.pprops(s) })
                if (sig !== b.sig) {
                  const props = this.pdiff(s, b.props)
                  ops.push({ op: 'para', block: b.ref, runs: this.wire(s.runs), ...(Object.keys(props).length ? { props } : {}), ...(this.model.track ? { author: this.author() } : {}) })
                  revNodes.add(node)
                }
              } else sigs.push({ sig: b.sig, props: b.props })
            } else sigs.push(null)
            prev = b
          } else if (node.classList.contains('qd-p')) {
            const s = this.ser(node)
            sigs.push({ sig: this.sig(s), props: this.pprops(s) })
            const op = { op: 'insert', runs: this.wire(s.runs), props: this.pprops(s), ...(this.model.track ? { author: this.author() } : {}) }
            revNodes.add(node)
            if (prev) { op.after = prev.ref; if (prev.kind === 'p') op.like = prev.ref } else orphans.push(op)
            inserts.push(op)
          } else { snap.pop() }
        }
        ops.push(...inserts)
        const gone = [...this.base.entries()].filter(([k]) => !seen.has(k)).map(([, b]) => b.ref)
        if (gone.length) ops.push({ op: 'delete', blocks: gone, ...(this.model.track ? { author: this.author() } : {}) })
        for (const td of this.dirtyCells) {
          if (!td.isConnected) continue
          const b = this.base.get(this.blockOf(td)?.dataset.k)
          if (b?.kind === 't') ops.push({ op: 'cell', table: b.ref, row: +td.dataset.r, col: +td.dataset.c, text: this.cellText(td) })
        }
        const struct = this.struct.splice(0)
        for (const s of struct) ops.push(s(this))
        ops.push(...this.extra.splice(0))
        this.dirtyBlocks.clear()
        this.dirtyCells.clear()
        this.touched = false
        if (!ops.length) return null
        if (struct.length) this.page.contentEditable = 'false'
        return { ops, snap, sigs, struct: struct.length > 0, revNodes }
      }

      wire (runs) { return runs.map((r) => (r.t === 'img' ? { t: 'img', n: r.n, from: r.from, ...(r.w && r.h ? { w: r.w, h: r.h } : {}), ...(r.wrap ? { wrap: r.wrap, align: r.align } : {}) } : r)) }

      cellText (td) {
        let s = ''
        const walk = (n) => { for (const c of n.childNodes) { if (c.nodeType === 3) s += c.data; else if (c.tagName === 'BR') s += '\n'; else if (c.nodeType === 1) { if (c.tagName === 'DIV' || c.tagName === 'P') { if (s && !s.endsWith('\n')) s += '\n' } walk(c) } } }
        walk(td)
        return s.replace(/​/g, '').replace(/ /g, ' ').replace(/\n+$/, '')
      }

      saved (model, batch) {
        if (batch.struct || model.blocks.length !== batch.snap.length) {
          const pending = this.touched
          this.load(model, { keep: true })
          if (pending) this.frame.toast('文档结构已变化，刚才保存期间的少量输入可能需要重做。', 'warn')
          // A new contents moves the headings after it: count again once
          // the pages settle.
          if (this.tocCheck) {
            this.tocCheck = false
            setTimeout(() => { if (this.model.toc && this.tocDrift()) this.tocOp({ op: 'toc_update' }) }, 400)
          }
          return
        }
        this.model = model
        this.showHf()
        // With changes tracked, saved paragraphs show the revisions the file
        // now holds; the one being typed in follows once the caret leaves.
        if (model.track && batch.revNodes?.size) {
          const here = this.blockOf(currentRange(this.page)?.startContainer)
          batch.snap.forEach((node, i) => {
            const b = model.blocks[i]
            if (b?.kind !== 'p' || !batch.revNodes.has(node) || !node.isConnected) return
            if (node === here) { this.revStale = node; return }
            const fresh = this.para(b)
            node.replaceWith(fresh)
            batch.snap[i] = fresh
            const s = this.ser(fresh)
            batch.sigs[i] = { sig: this.sig(s), props: this.pprops(s) }
          })
          this.renumber()
        }
        this.renderNotes()
        this.base = new Map()
        batch.snap.forEach((node, i) => {
          const b = model.blocks[i]
          if (!node.dataset.k) node.dataset.k = this.nextK()
          node.querySelectorAll('img[data-img]').forEach((im, j) => { im.dataset.from = b.ref; im.dataset.img = String(j) })
          if (b.kind === 't') { this.base.set(node.dataset.k, { ref: b.ref, kind: 't' }); return }
          const s = batch.sigs[i]
          this.base.set(node.dataset.k, { ref: b.ref, kind: 'p', sig: s.sig, props: s.props })
        })
        this.observe(this.mo.takeRecords())
      }

      failed () {
        // The page still holds the edits; send everything again next time.
        for (const n of this.page.children) this.dirtyBlocks.add(n)
        this.page.contentEditable = this.doc.writable ? 'true' : 'false'
      }

      // ------------------------------------------------ editing

      beforeInput (e) {
        if (!this.doc.writable) { e.preventDefault(); return }
        const t = e.inputType
        if (t.startsWith('format') || t === 'historyUndo' || t === 'historyRedo' || t === 'insertFromDrop') { e.preventDefault(); return }
        const range = currentRange(this.page)
        if (!range) return
        if (t === 'insertParagraph') { e.preventDefault(); this.split(); return }
        if (t === 'insertLineBreak') { e.preventDefault(); this.insertText('\n'); return }
        if (t.startsWith('delete')) {
          if (!range.collapsed) { e.preventDefault(); this.deleteRange(range); return }
          const cell = this.cellOf(range.startContainer)
          if (cell) return
          const block = this.blockOf(range.startContainer)
          if (!block || !block.classList.contains('qd-p')) return
          const back = t.endsWith('Backward')
          if (back && this.atEdge(range, block, true)) { e.preventDefault(); this.joinBack(block) } else if (!back && this.atEdge(range, block, false)) { e.preventDefault(); const next = block.nextElementSibling; if (next?.classList.contains('qd-p') && !next.dataset.locked) this.joinBack(next, true) }
          return
        }
        if (t === 'insertText' || t === 'insertCompositionText' || t === 'insertReplacementText') {
          if (!range.collapsed && this.blockOf(range.startContainer) !== this.blockOf(range.endContainer)) this.deleteRange(range)
        }
      }

      key (e) {
        if (e.key === 'Escape' && this.painter) { e.preventDefault(); this.stopPainter(); return }
        if (e.key === 'Tab' && !mod(e) && !e.altKey) {
          const range = currentRange(this.page)
          const td = range && this.cellOf(range.startContainer)
          if (td) {
            e.preventDefault()
            const tds = [...td.closest('.qd-t').querySelectorAll('td[contenteditable=true]')]
            const next = tds[tds.indexOf(td) + (e.shiftKey ? -1 : 1)]
            if (next) { next.focus(); const r = document.createRange(); r.selectNodeContents(next); const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r) }
            return
          }
          const b = range && this.blockOf(range.startContainer)
          if (b?.dataset.list) { e.preventDefault(); this.indent(e.shiftKey ? -1 : 1); return }
          if (b?.classList.contains('qd-p') && !e.shiftKey) { e.preventDefault(); this.insertText('\t'); return }
          return
        }
        if (!mod(e)) return
        const k = e.key.toLowerCase()
        if (k === 'a') {
          e.preventDefault()
          const cell = this.cellOf(currentRange(this.page)?.startContainer)
          const r = document.createRange(); r.selectNodeContents(cell ?? this.page)
          const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r)
        } else if (e.altKey && e.code === 'KeyM') { e.preventDefault(); this.newComment() } else if (e.altKey && e.code === 'KeyF') { e.preventDefault(); this.insertNote('fn') } else if (e.altKey && e.code === 'KeyD') { e.preventDefault(); this.insertNote('en') } else if (k === 'z') { e.preventDefault(); this.doc.step(e.shiftKey ? 'redo' : 'undo') } else if (k === 'y') { e.preventDefault(); this.doc.step('redo') } else if (k === 'b' || k === 'i' || k === 'u') { e.preventDefault(); this.toggle(k) } else if (k === 'k') { e.preventDefault(); this.linkSel() }
      }

      atEdge (range, block, start) {
        const r = document.createRange()
        if (start) { r.setStart(block, 0); r.setEnd(range.startContainer, range.startOffset) } else { r.setStart(range.endContainer, range.endOffset); r.setEnd(block, block.childNodes.length) }
        return !r.toString().replace(/​/g, '') && !r.cloneContents().querySelector?.('img,[data-pb]')
      }

      tidy (block) {
        for (const s of block.querySelectorAll('span.qd-r')) if (!s.textContent && !s.querySelector('img')) s.remove()
        const has = block.textContent || block.querySelector('img,[data-pb]')
        for (const br of block.querySelectorAll(':scope > br')) if (has || block.childNodes.length > 1) br.remove()
        if (!block.childNodes.length) block.append(el('br'))
      }

      deleteRange (range) {
        const sb = this.blockOf(range.startContainer)
        const eb = this.blockOf(range.endContainer)
        if (!sb || !eb) return
        if (this.cellOf(range.startContainer) || this.cellOf(range.endContainer)) {
          if (this.cellOf(range.startContainer) === this.cellOf(range.endContainer)) { range.deleteContents(); this.changed() }
          return
        }
        range.deleteContents()
        if (sb !== eb && sb.isConnected && eb.isConnected && sb.classList.contains('qd-p') && eb.classList.contains('qd-p') && !sb.dataset.locked && !eb.dataset.locked) {
          for (const br of sb.querySelectorAll(':scope > br')) br.remove()
          const at = sb.childNodes.length
          while (eb.firstChild) sb.append(eb.firstChild)
          eb.remove()
          this.tidy(sb)
          setCaret(sb, Math.min(at, sb.childNodes.length))
        } else if (sb.isConnected && sb.classList.contains('qd-p')) {
          this.tidy(sb)
          if (eb !== sb && eb.isConnected && eb.classList.contains('qd-p')) this.tidy(eb)
          const r = currentRange(this.page)
          if (!r || !sb.contains(r.startContainer)) setCaret(sb, sb.childNodes.length)
        }
        this.changed()
      }

      // joinBack merges the block into the one before it; forward keeps the
      // caret where the earlier block ends.
      joinBack (block, forward) {
        if (block.dataset.list && !forward) { delete block.dataset.list; delete block.dataset.lv; delete block.dataset.num; this.dress(block); this.changed(); return }
        const prev = block.previousElementSibling
        const empty = !block.textContent.replace(/​/g, '') && !block.querySelector('img,[data-pb]')
        if (!prev || !prev.classList.contains('qd-p') || prev.dataset.locked) {
          if (prev && empty) { const next = block.nextElementSibling; block.remove(); if (next?.classList.contains('qd-p')) setCaret(next, 0); this.changed() }
          return
        }
        for (const br of prev.querySelectorAll(':scope > br')) br.remove()
        const at = prev.childNodes.length
        if (!empty) while (block.firstChild) prev.append(block.firstChild)
        block.remove()
        this.tidy(prev)
        const last = prev.childNodes[at - 1]
        if (last?.nodeType === 3) setCaret(last, last.length)
        else if (last?.nodeType === 1 && last.lastChild?.nodeType === 3) setCaret(last.lastChild, last.lastChild.length)
        else setCaret(prev, Math.min(at, prev.childNodes.length))
        this.changed()
      }

      split () {
        let range = currentRange(this.page)
        if (!range) return
        if (this.cellOf(range.startContainer)) { this.insertText('\n'); return }
        if (!range.collapsed) { this.deleteRange(range); range = currentRange(this.page); if (!range) return }
        const block = this.blockOf(range.startContainer)
        if (!block || !block.classList.contains('qd-p') || block.dataset.locked) return
        if (block.dataset.list && !block.textContent.replace(/​/g, '') && !block.querySelector('img')) { delete block.dataset.list; delete block.dataset.lv; delete block.dataset.num; this.dress(block); this.changed(); return }
        const r = document.createRange()
        r.setStart(range.startContainer, range.startOffset)
        r.setEnd(block, block.childNodes.length)
        const frag = r.extractContents()
        const nb = block.cloneNode(false)
        delete nb.dataset.k
        delete nb.dataset.mark
        delete nb.dataset.drop
        nb.classList.remove('qd-drop', 'qd-dropnext')
        const style = block.dataset.style
        const heading = block.dataset.h || style?.startsWith('@h') || /heading|标题|title/i.test(this.model.styles?.find((s) => s.id === style)?.name ?? '')
        if (heading) { delete nb.dataset.style; delete nb.dataset.h }
        nb.append(frag)
        const carrier = nb.querySelector('span.qd-r')
        for (const s of nb.querySelectorAll('span.qd-r')) if (!s.textContent.replace(/​/g, '') && !s.querySelector('img')) s.remove()
        let caretNode
        if (!nb.textContent.replace(/​/g, '') && !nb.querySelector('img,[data-pb]')) {
          nb.replaceChildren()
          const c = carrier && !heading ? carrier.cloneNode(false) : el('span', { class: 'qd-r' })
          c.textContent = ZW
          nb.append(c)
          caretNode = c.firstChild
        }
        this.tidy(block)
        block.after(nb)
        this.dress(nb)
        if (caretNode) setCaret(caretNode, 1)
        else {
          const w = document.createTreeWalker(nb, NodeFilter.SHOW_TEXT)
          const first = w.nextNode()
          if (first) setCaret(first, 0); else setCaret(nb, 0)
        }
        nb.scrollIntoView?.({ block: 'nearest' })
        this.changed()
      }

      insertText (text) {
        let range = currentRange(this.page)
        if (!range) return
        if (!range.collapsed) { this.deleteRange(range); range = currentRange(this.page); if (!range) return }
        const c = range.startContainer
        const block = this.blockOf(c)
        if (block?.dataset.locked) return
        if (c.nodeType === 3) {
          c.insertData(range.startOffset, text)
          const end = range.startOffset + text.length
          // A line break at the very end needs something after it to show.
          if (text.endsWith('\n') && end === c.length) { c.appendData(ZW) }
          setCaret(c, end)
        } else {
          const node = document.createTextNode(text.endsWith('\n') ? text + ZW : text)
          const ref = c.childNodes[range.startOffset] ?? null
          if (ref?.nodeName === 'BR') ref.replaceWith(node); else c.insertBefore(node, ref)
          setCaret(node, text.length)
        }
        this.changed()
      }

      paste (e) {
        e.preventDefault()
        if (!this.doc.writable) return
        const img = imageFile(e.clipboardData)
        if (img) { const at = this.spot(); readImage(img).then((src) => this.insertPicture(src, at), (err) => this.frame.toast(err.message, 'warn')); return }
        const html = e.clipboardData.getData('text/html')
        const range = currentRange(this.page)
        if (html && range && !this.cellOf(range.startContainer)) {
          const t = htmlTable(html)
          if (t) { this.pasteTable(t); return }
          const paras = parseHtml(html)
          if (paras.length) { this.pasteParas(paras); return }
        }
        this.pasteText(e.clipboardData.getData('text/plain') || '')
      }

      // pasteParas puts formatted paragraphs at the caret: the first joins
      // the paragraph there, the rest follow as paragraphs of their own.
      pasteParas (paras) {
        let range = currentRange(this.page)
        if (!range) return
        if (!range.collapsed) { this.deleteRange(range); range = currentRange(this.page); if (!range) return }
        const block = this.blockOf(range.startContainer)
        if (!block?.classList.contains('qd-p') || block.dataset.locked) return
        const empty = !block.textContent.replace(/​/g, '') && !block.querySelector('img,[data-pb]')
        const spans = (p) => p.runs.filter((r) => r.text).map((r) => this.run({ t: 'text', ...r }))
        // Split the run under the caret so pasted runs sit beside it.
        const at = () => {
          const r = currentRange(this.page)
          let node = r.startContainer; let off = r.startOffset
          if (node.nodeType === 3) {
            const span = node.parentElement.closest('span.qd-r')
            const tail = node.splitText(off)
            if (span && span.parentNode === block) {
              const right = span.cloneNode(false)
              let n = tail
              const move = []
              while (n) { move.push(n); n = n.nextSibling }
              let up = tail.parentNode
              right.append(...move)
              while (up !== span) { up = up.parentNode }
              span.after(right)
              return { parent: block, before: right }
            }
            return { parent: tail.parentNode, before: tail }
          }
          return { parent: node, before: node.childNodes[off] ?? null }
        }
        const first = paras[0]
        if (empty && first.style) this.setStyle(block, first.style)
        if (empty && first.list) { block.dataset.list = first.list.ordered ? 'o' : 'b'; block.dataset.lv = String(first.list.level); this.dress(block) }
        const pos = at()
        const firstSpans = spans(first)
        for (const n of firstSpans) pos.parent.insertBefore(n, pos.before)
        for (const br of block.querySelectorAll(':scope > br')) br.remove()
        const last = firstSpans[firstSpans.length - 1]
        if (paras.length === 1) {
          if (last?.lastChild) setCaret(last.lastChild, last.lastChild.length)
        } else {
          if (last?.lastChild) setCaret(last.lastChild, last.lastChild.length); else setCaret(block, 0)
          this.split()
          const tail = this.blockOf(currentRange(this.page).startContainer)
          for (const p of paras.slice(1, -1)) {
            const d = this.para({ kind: 'p', runs: p.runs.map((r) => ({ t: 'text', ...r })), style: p.style, align: p.align, list: p.list })
            tail.before(d)
          }
          const end = paras[paras.length - 1]
          const endSpans = spans(end)
          for (const br of tail.querySelectorAll(':scope > br')) br.remove()
          for (const s of tail.querySelectorAll('span.qd-r')) if (s.textContent === ZW) s.remove()
          tail.prepend(...endSpans)
          if (!tail.textContent.replace(/​/g, '') && end.style) this.setStyle(tail, end.style)
          this.tidy(tail)
          const e = endSpans[endSpans.length - 1]
          if (e?.lastChild) setCaret(e.lastChild, e.lastChild.length); else setCaret(tail, 0)
        }
        this.tidy(block)
        this.renumber()
        this.changed()
      }

      pasteText (raw) {
        const text = raw.replace(/\r\n?/g, '\n')
        if (!text) return
        const range = currentRange(this.page)
        if (range && this.cellOf(range.startContainer)) { this.insertText(text); return }
        const lines = text.split('\n')
        lines.forEach((line, i) => { if (i) this.split(); if (line) this.insertText(line.replace(/\t/g, '\t')) })
      }

      cut (e) {
        const range = currentRange(this.page)
        if (!range || range.collapsed) return
        e.preventDefault()
        e.clipboardData.setData('text/plain', range.toString().replace(/​/g, ''))
        if (this.doc.writable) this.deleteRange(range)
      }

      // ------------------------------------------------ formatting

      editableText (t) {
        const b = this.blockOf(t)
        return b?.classList.contains('qd-p') && !b.dataset.locked
      }

      format (apply) {
        if (!this.doc.writable) return
        const range = currentRange(this.page)
        if (!range) return
        if (range.collapsed) {
          const b = this.blockOf(range.startContainer)
          if (!b?.classList.contains('qd-p') || b.dataset.locked) return
          const span = caretRun(range, b, () => el('span', { class: 'qd-r' }))
          apply(span)
          this.paint(span)
          this.changed()
          this.refreshBar()
          return
        }
        const texts = textsIn(range, this.page, (t) => this.editableText(t))
        for (const t of texts) { const s = isolate(t, this.blockOf(t), 'qd-r'); apply(s); this.paint(s) }
        selectTexts(texts)
        this.changed()
        this.refreshBar()
      }

      // toggle switches a property on, or off when all selected text has it.
      toggle (k) {
        const range = currentRange(this.page)
        if (!range) return
        const test = {
          b: (cs) => +cs.fontWeight >= 600,
          i: (cs) => cs.fontStyle === 'italic',
          u: (cs) => /underline/.test(cs.textDecorationLine),
          s: (cs) => /line-through/.test(cs.textDecorationLine)
        }[k]
        let on
        if (range.collapsed) {
          const n = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer
          on = n && test(getComputedStyle(n))
        } else {
          const nodes = textsIn(range.cloneRange(), this.page, (t) => this.editableText(t))
          on = nodes.length > 0 && nodes.every((t) => test(getComputedStyle(t.parentElement)))
        }
        this.format((s) => { s.dataset[k] = on ? '0' : '1' })
      }

      selectedBlocks () {
        const range = currentRange(this.page)
        if (!range) return []
        const out = []
        for (const d of this.page.children) if (d.classList.contains('qd-p') && !d.dataset.locked && (range.intersectsNode(d) || d.contains(range.startContainer))) out.push(d)
        return out
      }

      paraProps (fn) {
        if (!this.doc.writable) return
        const blocks = this.selectedBlocks()
        for (const d of blocks) fn(d)
        if (blocks.length) { this.renumber(); this.changed(); this.refreshBar() }
      }

      setStyle (d, v) {
        if (v) d.dataset.style = v; else delete d.dataset.style
        delete d.dataset.h
        if (v?.startsWith('@h')) d.dataset.h = v.slice(2)
        this.dress(d)
      }

      setAlign (d, a) {
        if (d.dataset.align === a) delete d.dataset.align; else d.dataset.align = a
        this.dress(d)
      }

      listToggle (ordered) {
        const blocks = this.selectedBlocks()
        const all = blocks.length && blocks.every((d) => d.dataset.list === (ordered ? 'o' : 'b'))
        this.paraProps((d) => {
          if (all) { delete d.dataset.list; delete d.dataset.lv; delete d.dataset.num } else { if (d.dataset.list !== (ordered ? 'o' : 'b')) delete d.dataset.num; d.dataset.list = ordered ? 'o' : 'b'; d.dataset.lv = d.dataset.lv ?? '0' }
          this.dress(d)
        })
      }

      pageBreak () {
        if (!this.doc.writable) return
        const range = currentRange(this.page)
        if (!range) return
        const b = this.blockOf(range.startContainer)
        if (!b?.classList.contains('qd-p') || b.dataset.locked) return
        range.collapse(false)
        const pb = el('span', { class: 'qd-pb', contenteditable: 'false', 'data-pb': '1' })
        range.insertNode(pb)
        this.tidy(b)
        this.split()
      }

      // Table changes go out after what is pending is saved, so the blocks
      // they name have their current references; the page reloads after.
      spot () {
        const range = currentRange(this.page)
        return { node: range ? this.blockOf(range.startContainer) : this.page.lastElementChild, cell: range ? this.cellOf(range.startContainer) : undefined }
      }

      async structural (make, at = this.spot()) {
        if (!this.doc.writable) return
        const { node, cell } = at
        await this.doc.flush()
        if (!node?.isConnected) return
        const b = this.base.get(node.dataset.k)
        if (!b) { this.frame.toast('请先把光标放在已保存的段落或表格里。', 'info'); return }
        const op = make(b, cell)
        if (!op) return
        this.struct.push(() => op)
        this.doc.dirty(true)
      }

      insertTable () { this.structural((b) => ({ op: 'table', after: b.ref, rows: 3, cols: 3 })) }
      // pasteTable puts a pasted table after the block with the caret.
      pasteTable (t) {
        const cols = Math.max(...t.rows.map((r) => r.length))
        this.structural((b) => ({ op: 'table', after: b.ref, rows: t.rows.length, cols, cells: t.rows, header: t.header }))
      }
      tableRow (action) {
        this.structural((b, cell) => {
          if (b.kind !== 't' || !cell) { this.frame.toast('请先把光标放在表格的单元格里。', 'info'); return undefined }
          return { op: 'row', table: b.ref, row: +cell.dataset.r, action }
        })
      }

      // ------------------------------------------------ table cells

      applyWidths (wrap) {
        const t = wrap.querySelector('table') ?? wrap.firstChild
        if (!t || !wrap.widths) return
        let cg = t.querySelector('colgroup')
        if (!cg) { cg = el('colgroup'); t.prepend(cg) }
        // In columns a table shows no wider than its column.
        const pg = this.model?.page
        const room = pg?.cols > 1 ? (pg.width - pg.left - pg.right - (pg.colSpace ?? 28) * (pg.cols - 1)) / pg.cols : Infinity
        const sum = wrap.widths.reduce((a, b) => a + b, 0)
        const f = sum > room ? room / sum : 1
        cg.replaceChildren(...wrap.widths.map((w) => el('col', { style: { width: w * f + 'px' } })))
        Object.assign(t.style, { tableLayout: 'fixed', width: sum * f + 'px', minWidth: '' })
      }

      // measureWidths reads the grid columns' widths off the screen, for a
      // table whose file gives none.
      measureWidths (wrap) {
        const n = wrap.gridN || 1
        const k = (this.page.offsetWidth / this.model.page.width) || 1
        const out = Array(n).fill(0)
        for (const td of wrap.querySelectorAll('td')) if (td.colSpan === 1) out[+td.dataset.g] ||= td.offsetWidth / k
        const known = out.filter(Boolean)
        const avg = known.length ? known.reduce((a, b) => a + b, 0) / known.length : (wrap.offsetWidth / k) / n
        return out.map((w) => Math.round(w || avg))
      }

      edgeOf (e) {
        const td = e.target.closest?.('.qd-t td')
        if (!td || !this.doc.writable) return undefined
        const r = td.getBoundingClientRect()
        if (e.clientX > r.right - 5) return { td, at: +td.dataset.g + td.colSpan }
        if (e.clientX < r.left + 4 && +td.dataset.g > 0) return { td, at: +td.dataset.g }
        return undefined
      }

      tableHover (e) {
        if (e.buttons) return
        const edge = this.edgeOf(e)
        const td = e.target.closest?.('.qd-t td')
        if (this.rzTd && this.rzTd !== td) this.rzTd.classList.remove('qd-rz')
        if (td) td.classList.toggle('qd-rz', !!edge)
        this.rzTd = td
      }

      tableDown (e) {
        if (e.button !== 0) return
        const edge = this.edgeOf(e)
        const td = e.target.closest?.('.qd-t td')
        if (edge) { e.preventDefault(); this.resizeColumn(e, edge); return }
        if (e.shiftKey && td && this.csel?.anchor && this.csel.anchor.closest('.qd-t') === td.closest('.qd-t')) { e.preventDefault(); this.selectCells(this.csel.anchor, td); return }
        if (e.shiftKey && td) {
          const at = this.cellOf(currentRange(this.page)?.startContainer)
          if (at && at !== td && at.closest('.qd-t') === td.closest('.qd-t')) { e.preventDefault(); this.selectCells(at, td); return }
        }
        this.clearCells()
        if (!td) return
        const wrap = td.closest('.qd-t')
        const move = (ev) => {
          const over = document.elementFromPoint(ev.clientX, ev.clientY)?.closest?.('td')
          if (!over || over.closest('.qd-t') !== wrap) return
          if (over === td && !this.csel) return
          window.getSelection()?.removeAllRanges()
          this.selectCells(td, over)
        }
        const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); this.page.classList.remove('qd-csel-on') }
        document.addEventListener('mousemove', move)
        document.addEventListener('mouseup', up)
      }

      // selectCells selects the rectangle between two cells, grown to hold
      // every merged cell it touches, as Word does.
      selectCells (a, b) {
        const wrap = a.closest('.qd-t')
        const box = (td) => [+td.dataset.r, +td.dataset.r + td.rowSpan - 1, +td.dataset.g, +td.dataset.g + td.colSpan - 1]
        let [r1, r2, g1, g2] = box(a)
        const [s1, s2, h1, h2] = box(b)
        r1 = Math.min(r1, s1); r2 = Math.max(r2, s2); g1 = Math.min(g1, h1); g2 = Math.max(g2, h2)
        const tds = [...wrap.querySelectorAll('td')]
        for (let grew = true; grew;) {
          grew = false
          for (const td of tds) {
            const [t1, t2, c1, c2] = box(td)
            if (t2 < r1 || t1 > r2 || c2 < g1 || c1 > g2) continue
            if (t1 < r1) { r1 = t1; grew = true }
            if (t2 > r2) { r2 = t2; grew = true }
            if (c1 < g1) { g1 = c1; grew = true }
            if (c2 > g2) { g2 = c2; grew = true }
          }
        }
        for (const td of tds) {
          const [t1, t2, c1, c2] = box(td)
          td.classList.toggle('qd-csel', !(t2 < r1 || t1 > r2 || c2 < g1 || c1 > g2))
        }
        this.page.classList.add('qd-csel-on')
        this.csel = { wrap, anchor: a, td: a, r1, r2, g1, g2 }
        this.status()
      }

      clearCells () {
        if (!this.csel) return
        for (const td of this.csel.wrap.querySelectorAll('td.qd-csel')) td.classList.remove('qd-csel')
        this.csel = undefined
      }

      cellRect () {
        if (this.csel?.wrap.isConnected) return this.csel
        const td = this.cellOf(currentRange(this.page)?.startContainer)
        if (!td) return undefined
        const r = +td.dataset.r; const g = +td.dataset.g
        return { wrap: td.closest('.qd-t'), td, r1: r, r2: r + td.rowSpan - 1, g1: g, g2: g + td.colSpan - 1 }
      }

      tableAct (make) {
        const s = this.cellRect()
        if (!s) { this.frame.toast('请先把光标放在表格的单元格里。', 'info'); return }
        this.structural((b) => (b.kind === 't' ? make(b, s) : undefined), { node: s.wrap, cell: s.td })
      }

      mergeCells () {
        this.tableAct((b, s) => {
          if (s.r1 === s.r2 && s.g1 === s.g2) { this.frame.toast('先拖选两个以上的单元格再合并。', 'info'); return undefined }
          return { op: 'merge', table: b.ref, r1: s.r1, r2: s.r2, g1: s.g1, g2: s.g2 }
        })
      }

      splitCells () {
        const s = this.cellRect()
        const td = s?.td
        if (!td) { this.frame.toast('请先把光标放在要拆分的单元格里。', 'info'); return }
        if (td.rowSpan > 1 || td.colSpan > 1) { this.tableAct((b) => ({ op: 'unmerge', table: b.ref, row: +td.dataset.r, g: +td.dataset.g })); return }
        const ans = window.prompt('拆分为几列、几行？（例如 2,1 表示 2 列 1 行）', '2,1')
        if (!ans) return
        const [cols, rows] = ans.split(/[,，x×\s]+/).map((n) => parseInt(n, 10))
        if (!(cols >= 1) || !((rows || 1) >= 1) || (cols === 1 && (rows || 1) === 1)) { this.frame.toast('请输入列数和行数，例如 2,1。', 'info'); return }
        this.tableAct((b) => ({ op: 'split', table: b.ref, row: +td.dataset.r, g: +td.dataset.g, cols, rows: rows || 1 }))
      }

      cellFormat (props) {
        this.tableAct((b, s) => ({ op: 'tcell', table: b.ref, r1: s.r1, r2: s.r2, g1: s.g1, g2: s.g2, ...props }))
      }

      shadeMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const colors = [['无底纹', ''], ['浅灰', 'F2F2F2'], ['灰色', 'D9D9D9'], ['浅蓝', 'DEEAF6'], ['浅绿', 'E2EFDA'], ['浅黄', 'FFF2CC'], ['浅橙', 'FCE4D6'], ['蓝色', '4472C4'], ['深蓝', '1F3864']]
        menu(r.left, r.bottom + 2, [
          ...colors.map(([label, hex]) => ({ label: (hex ? '■ ' : '□ ') + label, fn: () => this.cellFormat({ fill: hex }) })),
          '-',
          { label: '其他颜色…', fn: () => { const i = el('input', { type: 'color', value: '#ffff00' }); i.addEventListener('change', () => this.cellFormat({ fill: i.value.slice(1).toUpperCase() })); i.click() } }
        ])
      }

      borderMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const items = [['所有框线', 'all'], ['外侧框线', 'outer'], ['内部框线', 'inner'], ['上框线', 'top'], ['下框线', 'bottom'], ['左框线', 'left'], ['右框线', 'right'], ['无框线', 'none']]
        menu(r.left, r.bottom + 2, items.map(([label, border]) => ({ label, fn: () => this.cellFormat({ border }) })))
      }

      cellAlignMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const v = [['top', '靠上'], ['center', '居中'], ['bottom', '靠下']]
        const h = [['left', '左对齐'], ['center', '居中'], ['right', '右对齐']]
        const items = []
        for (const [va, vl] of v) {
          for (const [ha, hl] of h) items.push({ label: `${vl} · ${hl}`, fn: () => this.cellFormat({ valign: va, align: ha }) })
          items.push('-')
        }
        menu(r.left, r.bottom + 2, items)
      }

      tableStyleMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const items = [['网格', 'grid'], ['无框线', 'plain'], ['浅色表头', 'light'], ['镶边行（蓝）', 'banded'], ['深色表头', 'dark']]
        menu(r.left, r.bottom + 2, items.map(([label, style]) => ({ label, fn: () => this.tableAct((b) => ({ op: 'tstyle', table: b.ref, style })) })))
      }

      evenColumns () {
        this.tableAct((b, s) => {
          const ws = s.wrap.widths ?? this.measureWidths(s.wrap)
          const each = ws.reduce((a, c) => a + c, 0) / ws.length
          return { op: 'twidth', table: b.ref, widths: ws.map(() => Math.round(each * 15)) }
        })
      }

      // resizeColumn drags a column border; the column to its right gives
      // up the width, and the last border widens the table.
      resizeColumn (e, { td, at }) {
        const wrap = td.closest('.qd-t')
        const ws = wrap.widths ? [...wrap.widths] : this.measureWidths(wrap)
        const start = [...ws]
        const k = (this.page.offsetWidth / this.model.page.width) || 1
        const x0 = e.clientX
        const guide = el('div', { class: 'qd-colguide' })
        document.body.append(guide)
        const place = (x) => { const r = wrap.getBoundingClientRect(); Object.assign(guide.style, { left: x + 'px', top: r.top + 'px', height: r.height + 'px' }) }
        place(x0)
        const move = (ev) => {
          let dx = (ev.clientX - x0) / k
          const i = at - 1
          if (i < 0) return
          if (at < start.length) dx = Math.max(-(start[i] - 20), Math.min(start[at] - 20, dx))
          else dx = Math.max(-(start[i] - 20), dx)
          ws[i] = start[i] + dx
          if (at < start.length) ws[at] = start[at] - dx
          wrap.widths = ws.map(Math.round)
          this.applyWidths(wrap)
          place(x0 + dx * k)
        }
        const up = () => {
          document.removeEventListener('mousemove', move)
          document.removeEventListener('mouseup', up)
          guide.remove()
          if (ws.every((w, i) => Math.abs(w - start[i]) < 1)) return
          this.structural((b) => (b.kind === 't' ? { op: 'twidth', table: b.ref, widths: ws.map((w) => Math.round(w * 15)) } : undefined), { node: wrap, cell: td })
        }
        document.addEventListener('mousemove', move)
        document.addEventListener('mouseup', up)
      }

      deleteTable () {
        this.structural((b) => {
          if (b.kind !== 't') { this.frame.toast('请先把光标放在表格里。', 'info'); return undefined }
          return { op: 'delete', blocks: [b.ref] }
        })
      }

      tableCol (action) {
        this.structural((b, cell) => {
          if (b.kind !== 't' || !cell) { this.frame.toast('请先把光标放在表格的单元格里。', 'info'); return undefined }
          return { op: 'col', table: b.ref, col: +cell.dataset.c, action }
        })
      }

      // ------------------------------------------------ pictures

      async pickPicture () {
        if (!this.doc.writable) return
        const at = this.spot()
        const src = await pickImage()
        if (src) this.insertPicture(src, at)
      }

      insertPicture (src, at) {
        this.structural((b) => ({ op: 'image', after: b.ref, src }), at)
      }

      drop (e) {
        e.preventDefault()
        if (!this.doc.writable) return
        const f = imageFile(e.dataTransfer)
        if (!f) { if (e.dataTransfer?.files?.length) this.frame.toast('只能拖入 PNG、JPEG、GIF 图片。', 'info'); return }
        const pos = document.caretRangeFromPoint?.(e.clientX, e.clientY)
        if (pos && this.page.contains(pos.startContainer)) { const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(pos) }
        const at = this.spot()
        readImage(f).then((src) => this.insertPicture(src, at), (err) => this.frame.toast(err.message, 'warn'))
      }

      // ------------------------------------------------ runs and paragraphs

      vert (v) {
        const range = currentRange(this.page)
        const n = range && (range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer)
        const on = n?.closest?.('[data-vert]')?.dataset.vert === v
        this.format((s) => { if (on) delete s.dataset.vert; else s.dataset.vert = v })
      }

      linkSel () {
        if (!this.doc.writable) return
        let range = currentRange(this.page)
        if (!range) return
        const n = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer
        const cur = n?.closest?.('[data-link]')
        if (range.collapsed) {
          if (!cur) { this.frame.toast('请先选中要加链接的文字。', 'info'); return }
          const r = document.createRange(); r.selectNodeContents(cur)
          const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r)
          range = r
        }
        const url = window.prompt('链接地址（留空则移除链接）', cur?.dataset.link ?? 'https://')
        if (url === null) return
        const v = url.trim()
        if (v && !/^(https?:\/\/|mailto:|#)/i.test(v)) { this.frame.toast('链接需以 http://、https://、mailto: 或 # 开头。', 'warn'); return }
        this.format((s) => { if (v && v !== 'https://') s.dataset.link = v; else delete s.dataset.link })
      }

      indent (d) {
        const step = 420
        this.paraProps((p) => {
          if (p.dataset.list) p.dataset.lv = String(Math.max(0, Math.min(8, (+p.dataset.lv || 0) + d)))
          else {
            const left = Math.max(0, (+p.dataset.ind || 0) + d * step)
            if (left || p.dataset.ind !== undefined) p.dataset.ind = String(left)
          }
          this.dress(p)
        })
      }

      firstLine () {
        const size = this.model.looks['']?.size ?? 21
        const blocks = this.selectedBlocks()
        const on = blocks.length && blocks.every((p) => +p.dataset.fl > 0)
        this.paraProps((p) => { if (on) p.dataset.fl = '0'; else p.dataset.fl = String(size * 20); this.dress(p) })
      }

      // ------------------------------------------------ lists, borders, drop caps

      listMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const pick = (style) => this.listStyle(style)
        menu(r.left, r.bottom + 2, [
          { label: '项目符号：•  ○  ■', fn: () => pick('disc') },
          { label: '○ 空心圆', fn: () => pick('circle') },
          { label: '■ 方块', fn: () => pick('square') },
          { label: '◆ 菱形', fn: () => pick('diamond') },
          { label: '✓ 对勾', fn: () => pick('check') },
          { label: '➢ 箭头', fn: () => pick('arrow') },
          '-',
          { label: '1.  2.  3.', fn: () => pick('decimal') },
          { label: '1)  2)  3)', fn: () => pick('paren') },
          { label: '(1) (2) (3)', fn: () => pick('fullparen') },
          { label: '一、二、三、', fn: () => pick('cn') },
          { label: '(一) (二) (三)', fn: () => pick('cnparen') },
          { label: 'a.  b.  c.', fn: () => pick('lowerLetter') },
          { label: 'A.  B.  C.', fn: () => pick('upperLetter') },
          { label: 'i.  ii.  iii.', fn: () => pick('lowerRoman') },
          { label: 'I.  II.  III.', fn: () => pick('upperRoman') },
          { label: '①  ②  ③', fn: () => pick('circled') },
          '-',
          { label: '多级：1. / 1.1 / 1.1.1', fn: () => pick('multi') },
          { label: '多级：一、/（一）/ 1. / (1)', fn: () => pick('outline') },
          '-',
          { label: 'Tab / Shift+Tab 调整列表级别', disabled: true }
        ])
      }

      // listStyle gives the selected paragraphs a new list of a style; the
      // page reloads so they carry the list the file now has.
      async listStyle (style) {
        if (!this.doc.writable) return
        const picked = this.selectedBlocks()
        if (!picked.length) { this.frame.toast('请先把光标放在段落里。', 'info'); return }
        await this.doc.flush()
        const refs = picked.filter((d) => d.isConnected).map((d) => this.base.get(d.dataset.k)?.ref).filter(Boolean)
        if (!refs.length) return
        this.struct.push(() => ({ op: 'liststyle', blocks: refs, style }))
        this.doc.dirty(true)
      }

      paraBorderMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const set = (v) => this.paraProps((d) => { if (v) d.dataset.bdr = v; else delete d.dataset.bdr; this.dress(d) })
        menu(r.left, r.bottom + 2, [
          { label: '方框', fn: () => set('tlbr') },
          { label: '上框线', fn: () => set('t') },
          { label: '下框线', fn: () => set('b') },
          { label: '上下框线', fn: () => set('tb') },
          { label: '左框线', fn: () => set('l') },
          '-',
          { label: '无框线', fn: () => set('') }
        ])
      }

      paraShadeMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const set = (v) => this.paraProps((d) => { if (v) d.dataset.shade = v; else delete d.dataset.shade; this.dress(d) })
        const colors = [['浅灰', 'F2F2F2'], ['灰色', 'D9D9D9'], ['浅蓝', 'DEEAF6'], ['浅绿', 'E2EFDA'], ['浅黄', 'FFF2CC'], ['浅橙', 'FCE4D6']]
        menu(r.left, r.bottom + 2, [
          ...colors.map(([label, hex]) => ({ label, fn: () => set(hex) })),
          { label: '其他颜色…', fn: () => { const i = el('input', { type: 'color', value: '#fff2cc' }); i.addEventListener('change', () => set(i.value.slice(1).toUpperCase())); i.click() } },
          '-',
          { label: '无底纹', fn: () => set('') }
        ])
      }

      dropMenu (e) {
        const r = e.currentTarget.getBoundingClientRect()
        const range = currentRange(this.page)
        let b = range ? this.blockOf(range.startContainer) : undefined
        if (b && !b.dataset.drop && b.previousElementSibling?.dataset.drop) b = b.previousElementSibling
        const set = (lines) => this.structural((x) => (x.kind === 'p' ? { op: 'dropcap', block: x.ref, lines } : undefined), { node: b, cell: undefined })
        const on = !!b?.dataset.drop
        menu(r.left, r.bottom + 2, [
          { label: '无', disabled: !on, fn: () => set(0) },
          { label: '下沉 2 行', disabled: !b, fn: () => set(2) },
          { label: '下沉 3 行', disabled: !b, fn: () => set(3) },
          { label: '下沉 4 行', disabled: !b, fn: () => set(4) }
        ])
      }

      // ------------------------------------------------ revisions

      toggleTrack () {
        if (!this.doc.writable) return
        const on = !this.model.track
        this.model.track = on
        this.extra.push({ op: 'track', on })
        this.ui.track.classList.toggle('on', on)
        this.frame.toast(on ? '修订已打开：之后的修改会记录为修订。' : '修订已关闭。', 'info')
        this.doc.dirty(true)
      }

      async revOp (op) {
        if (!this.doc.writable) return
        await this.doc.flush()
        this.struct.push(() => op)
        this.doc.dirty(true)
      }

      revAt () {
        const range = currentRange(this.page)
        if (!range) return undefined
        const n = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer
        const hit = n?.closest?.('[data-rev]')
        if (hit) return hit.dataset.rev.slice(4)
        const kids = n?.childNodes ? [...n.childNodes] : []
        const near = (range.startContainer.nodeType === 1 ? [kids[range.startOffset - 1], kids[range.startOffset]] : [n?.previousSibling, n?.nextSibling]).find((x) => x?.classList?.contains('qd-del'))
        return near?.dataset.del
      }

      revMenu (e, action) {
        const r = e.currentTarget.getBoundingClientRect()
        const id = this.revAt()
        const range = currentRange(this.page)
        const block = range ? this.blockOf(range.startContainer) : undefined
        const word = action === 'accept' ? '接受' : '拒绝'
        menu(r.left, r.bottom + 2, [
          { label: `${word}此处修订`, disabled: !id, fn: () => this.revOp({ op: 'rev', action, id }) },
          { label: `${word}本段修订`, disabled: !block?.classList.contains('qd-p'), fn: () => this.structural((b) => (b.kind === 'p' ? { op: 'rev', action, block: b.ref } : undefined), { node: block }) },
          { label: `${word}所有修订`, fn: () => this.revOp({ op: 'rev', action }) }
        ])
      }

      revJump (dir) {
        const marks = [...this.page.querySelectorAll('[data-rev],.qd-del')]
        if (!marks.length) { this.frame.toast('文档里没有修订。', 'info'); return }
        const range = currentRange(this.page)
        const pos = (n) => (range ? range.comparePoint(n, 0) : -1)
        const next = dir > 0
          ? marks.find((n) => pos(n) > 0 && !n.contains(range?.startContainer)) ?? marks[0]
          : [...marks].reverse().find((n) => pos(n) < 0 && !n.contains(range?.startContainer)) ?? marks[marks.length - 1]
        next.scrollIntoView({ block: 'center' })
        const r = document.createRange()
        r.selectNodeContents(next)
        this.page.focus({ preventScroll: true })
        const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r)
      }

      // freshenRev redraws the paragraph last typed in once the caret
      // leaves it, so its new revisions show.
      freshenRev () {
        const node = this.revStale
        if (!node) return
        if (!node.isConnected) { this.revStale = null; return }
        const range = currentRange(this.page)
        if (range && node.contains(range.startContainer)) return
        if (this.dirtyBlocks.has(node) || this.touched) return
        const base = this.base.get(node.dataset.k)
        const b = base && this.model.blocks.find((x) => x.ref === base.ref)
        this.revStale = null
        if (!b || b.kind !== 'p') return
        const fresh = this.para(b)
        fresh.dataset.k = node.dataset.k
        this.observe(this.mo.takeRecords())
        node.replaceWith(fresh)
        this.mo.takeRecords()
        const s = this.ser(fresh)
        this.base.set(fresh.dataset.k, { ...base, sig: this.sig(s), props: this.pprops(s) })
        this.renumber()
      }

      // ------------------------------------------------ footnotes and endnotes

      insertNote (kind) {
        if (!this.doc.writable) return
        const range = currentRange(this.page)
        const block = range && this.blockOf(range.startContainer)
        if (!block?.classList.contains('qd-p') || block.dataset.locked || this.cellOf(range.startContainer)) { this.frame.toast('请把光标放在正文段落里再插入' + (kind === 'fn' ? '脚注。' : '尾注。'), 'info'); return }
        const text = window.prompt(kind === 'fn' ? '脚注内容' : '尾注内容', '')
        if (text === null) return
        const id = 'n-' + Math.random().toString(36).slice(2, 8)
        const r = currentRange(this.page) ?? range
        r.collapse(false)
        const mark = el('span', { class: 'qd-fn', contenteditable: 'false', dataset: { note: kind, id } })
        r.insertNode(mark)
        setCaret(mark.parentNode, [...mark.parentNode.childNodes].indexOf(mark) + 1)
        this.tidy(block)
        this.renumber()
        this.changed()
        this.struct.push(() => ({ op: 'note_new', kind, id, text }))
        this.doc.dirty(true)
      }

      renderNotes () {
        const notes = this.model.notes ?? []
        const byKey = new Map(notes.map((n) => [n.kind + ':' + n.id, n]))
        const refs = [...this.page.querySelectorAll('.qd-fn')]
        const sections = []
        for (const kind of ['fn', 'en']) {
          const items = refs.filter((r) => r.dataset.note === kind).map((r) => ({ ref: r, note: byKey.get(kind + ':' + r.dataset.id) })).filter((x) => x.note)
          if (!items.length) continue
          sections.push(el('div', { class: 'qd-notes-h' }, kind === 'fn' ? '脚注' : '尾注'))
          for (const { ref, note } of items) {
            const t = el('div', { class: 'qd-note-t', contenteditable: this.doc.writable && !note.locked ? 'plaintext-only' : 'false' }, note.text)
            t.addEventListener('blur', () => {
              const text = (t.innerText ?? t.textContent).replace(/\n$/, '')
              if (text === note.text) return
              note.text = text
              this.extra.push({ op: 'note_edit', kind, id: note.id, text })
              this.doc.dirty(true)
            })
            t.addEventListener('keydown', (e) => { if (e.key === 'Escape') { t.textContent = note.text; t.blur() } })
            const num = el('span', { class: 'qd-note-n', title: '跳到正文中的位置' }, ref.textContent)
            num.addEventListener('click', () => { this.page.querySelector(`.qd-fn[data-note="${kind}"][data-id="${note.id}"]`)?.scrollIntoView({ block: 'center' }) })
            sections.push(el('div', { class: 'qd-note', dataset: { note: kind, id: note.id } }, num, t))
          }
        }
        if (this.notesBox.contains(document.activeElement)) return
        this.notesBox.replaceChildren(...sections)
        this.notesBox.hidden = !sections.length
      }

      // ------------------------------------------------ pictures in text

      dressImg (img) {
        const w = img.dataset.wrap
        const a = img.dataset.align ?? 'left'
        const st = img.style
        st.float = ''; st.display = ''; st.margin = ''; st.position = ''; st.zIndex = ''; st.opacity = ''
        if (!w || w === 'inline') return
        if (w === 'topBottom') { st.display = 'block'; st.margin = a === 'center' ? '4px auto' : a === 'right' ? '4px 0 4px auto' : '4px auto 4px 0'; return }
        st.float = a === 'right' ? 'right' : 'left'
        st.margin = a === 'right' ? '2px 0 4px 10px' : '2px 10px 4px 0'
        if (w === 'front') { st.position = 'relative'; st.zIndex = '3' }
        if (w === 'behind') { st.position = 'relative'; st.zIndex = '0'; st.opacity = '0.9' }
      }

      pickImg (img) {
        if (this.imgSel === img) return
        this.dropImg()
        this.imgSel = img
        img.classList.add('qd-imgsel')
        if (!this.doc.writable) return
        const h = el('div', { class: 'qd-imgh', title: '拖动调整大小（按住 Shift 不保持比例）' })
        document.body.append(h)
        this.imgHandle = h
        this.placeImgHandle()
        h.addEventListener('mousedown', (e) => this.resizeImg(e))
        this.onImgKey = (e) => {
          if (!this.imgSel?.isConnected) { this.dropImg(); return }
          if ((e.key === 'Delete' || e.key === 'Backspace') && this.page.contains(document.activeElement ?? null)) { e.preventDefault(); e.stopPropagation(); this.removeImg(this.imgSel) }
          if (e.key === 'Escape') this.dropImg()
        }
        document.addEventListener('keydown', this.onImgKey, true)
      }

      removeImg (img) {
        const b = this.blockOf(img)
        this.dropImg()
        img.remove()
        if (b) { this.tidy(b); this.dirtyBlocks.add(b) }
        this.changed()
      }

      dropImg () {
        this.imgSel?.classList.remove('qd-imgsel')
        this.imgSel = null
        this.imgHandle?.remove()
        this.imgHandle = null
        if (this.onImgKey) document.removeEventListener('keydown', this.onImgKey, true)
        this.onImgKey = null
      }

      placeImgHandle () {
        if (!this.imgHandle || !this.imgSel) return
        if (!this.imgSel.isConnected) { this.dropImg(); return }
        const r = this.imgSel.getBoundingClientRect()
        Object.assign(this.imgHandle.style, { left: r.right - 6 + 'px', top: r.bottom - 6 + 'px' })
      }

      imgTouched (img) {
        const b = this.blockOf(img)
        if (b) this.dirtyBlocks.add(b)
        this.changed()
        this.placeImgHandle()
      }

      resizeImg (e) {
        e.preventDefault(); e.stopPropagation()
        const img = this.imgSel
        const r0 = img.getBoundingClientRect()
        const w0 = parseFloat(img.style.width) || img.offsetWidth
        const h0 = parseFloat(img.style.height) || img.offsetHeight
        const z = r0.width / w0 || 1
        const x0 = e.clientX; const y0 = e.clientY
        const move = (ev) => {
          let w = Math.max(16, w0 + (ev.clientX - x0) / z)
          let h = Math.max(16, h0 + (ev.clientY - y0) / z)
          if (!ev.shiftKey) { const f = Math.max(w / w0, h / h0); w = w0 * f; h = h0 * f }
          img.style.width = Math.round(w) + 'px'
          img.style.height = Math.round(h) + 'px'
          this.placeImgHandle()
        }
        const up = () => {
          document.removeEventListener('mousemove', move)
          document.removeEventListener('mouseup', up)
          this.imgTouched(img)
        }
        document.addEventListener('mousemove', move)
        document.addEventListener('mouseup', up)
      }

      imgMenu (img, x, y) {
        const ro = !this.doc.writable
        const w = img.dataset.wrap ?? 'inline'
        const mark = (on, label) => (on ? '✓ ' : '  ') + label
        const set = (wrap, align) => {
          img.dataset.wrap = wrap
          img.dataset.align = align ?? img.dataset.align ?? 'left'
          this.dressImg(img)
          this.imgTouched(img)
        }
        const floating = w !== 'inline'
        menu(x, y, [
          { label: mark(w === 'inline', '嵌入型'), disabled: ro, fn: () => set('inline') },
          { label: mark(w === 'square', '四周型'), disabled: ro, fn: () => set('square') },
          { label: mark(w === 'tight', '紧密型'), disabled: ro, fn: () => set('tight') },
          { label: mark(w === 'topBottom', '上下型'), disabled: ro, fn: () => set('topBottom') },
          { label: mark(w === 'behind', '衬于文字下方'), disabled: ro, fn: () => set('behind') },
          { label: mark(w === 'front', '浮于文字上方'), disabled: ro, fn: () => set('front') },
          '-',
          { label: mark(floating && img.dataset.align === 'left', '靠左'), disabled: ro || !floating, fn: () => set(w, 'left') },
          { label: mark(floating && img.dataset.align === 'center', '居中'), disabled: ro || !floating, fn: () => set(w, 'center') },
          { label: mark(floating && img.dataset.align === 'right', '靠右'), disabled: ro || !floating, fn: () => set(w, 'right') },
          '-',
          {
            label: '重设大小',
            disabled: ro || !img.naturalWidth,
            fn: () => {
              const pg = this.model.page
              const f = Math.min(1, (pg.width - pg.left - pg.right) / img.naturalWidth)
              img.style.width = Math.round(img.naturalWidth * f) + 'px'
              img.style.height = Math.round(img.naturalHeight * f) + 'px'
              this.imgTouched(img)
            }
          },
          { label: '删除图片', disabled: ro, fn: () => this.removeImg(img) }
        ])
      }

      // ------------------------------------------------ format painter

      startPainter (sticky) {
        if (!this.doc.writable) return
        const range = currentRange(this.page)
        const n = range ? (range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer) : undefined
        const block = range ? this.blockOf(range.startContainer) : undefined
        if (!block?.classList.contains('qd-p')) { this.frame.toast('先把光标放在要复制格式的文字里。', 'info'); return }
        const run = {}
        const keys = ['b', 'i', 'u', 's', 'color', 'size', 'hl', 'vert', 'font']
        for (const k of keys) { const owner = n?.closest?.(`[data-${k}]`); if (owner && block.contains(owner)) run[k] = owner.dataset[k] }
        const para = {}
        for (const k of ['style', 'h', 'align', 'line', 'before', 'after', 'ind', 'fl', 'bdr', 'shade']) if (block.dataset[k] !== undefined) para[k] = block.dataset[k]
        this.painter = { run, para, keys, sticky }
        this.ui.painter.classList.add('on')
        this.page.classList.add('qd-painting')
        if (!this.onPaintUp) {
          this.onPaintUp = () => setTimeout(() => this.paintUp(), 0)
          this.page.addEventListener('mouseup', this.onPaintUp)
        }
      }

      paintUp () {
        const p = this.painter
        if (!p) return
        const range = currentRange(this.page)
        if (!range) return
        if (!range.collapsed) {
          this.format((s) => { for (const k of p.keys) { if (p.run[k] !== undefined) s.dataset[k] = p.run[k]; else delete s.dataset[k] } })
        }
        this.paraProps((d) => {
          for (const k of ['style', 'h', 'align', 'line', 'before', 'after', 'ind', 'fl', 'bdr', 'shade']) { if (p.para[k] !== undefined) d.dataset[k] = p.para[k]; else delete d.dataset[k] }
          this.dress(d)
        })
        if (!p.sticky) this.stopPainter()
      }

      stopPainter () {
        this.painter = null
        this.ui.painter.classList.remove('on')
        this.page.classList.remove('qd-painting')
      }

      // ------------------------------------------------ header, footer, page

      showHf () {
        for (const [box, kind] of [[this.hdr, 'header'], [this.ftr, 'footer']]) {
          if (document.activeElement === box) continue
          const info = this.model[kind] ?? { text: '', editable: true }
          const ok = this.doc.writable && info.editable
          box.contentEditable = ok ? 'plaintext-only' : 'false'
          box.classList.toggle('qd-hf-ro', !ok)
          box.textContent = info.page ? (info.text.trim() ? info.text.replace('第  页', '第 # 页').replace('共  页', '共 # 页') : '') : info.text
          box.title = info.page ? '页脚含页码，用工具栏的 # 按钮移除' : !info.editable ? '页眉页脚含复杂内容，请在 Word/WPS 中编辑' : `点击编辑${kind === 'header' ? '页眉' : '页脚'}`
        }
        this.ui.pageNo.classList.toggle('on', !!this.model.footer?.page)
      }

      hfText (box) { return (box.innerText ?? box.textContent).replace(/\n$/, '') }

      hfDone (box, kind) {
        const info = this.model[kind]
        if (!this.doc.writable || !info?.editable || info.page) return
        const text = this.hfText(box)
        if (text === info.text || (!info.exists && !text.trim())) return
        this.extra.push({ op: 'hf', kind, text })
        this.model[kind] = { ...info, text, exists: true }
        this.doc.dirty(true)
      }

      togglePageNumber () {
        if (!this.doc.writable) return
        const f = this.model.footer
        if (!f.editable && !f.page) { this.frame.toast('页脚含复杂内容，请在 Word/WPS 中编辑。', 'info'); return }
        this.extra.push(f.page ? { op: 'hf', kind: 'footer', text: '' } : { op: 'hf', kind: 'footer', pageNumber: true })
        this.doc.dirty(true)
      }

      pageMenu (e, part) {
        const pg = this.model.page
        const r = (e?.currentTarget ?? this.ui.pageSetup).getBoundingClientRect()
        const mark = (on, label) => (on ? '✓ ' : '  ') + label
        const set = (props) => { this.struct.push(() => ({ op: 'page', ...props })); this.doc.dirty(true) }
        const parts = {
          cols: [1, 2, 3].map((n) => ({ label: mark((pg.cols ?? 1) === n, ['一栏', '两栏', '三栏'][n - 1]), fn: () => set({ cols: n }) })),
          size: Object.entries(PAGE_LABEL).map(([k, l]) => ({ label: mark(pg.size === k, (part ? '' : '纸张 ') + l), fn: () => set({ size: k }) })),
          orient: [
            { label: mark(pg.orient !== 'landscape', '纵向'), fn: () => set({ orient: 'portrait' }) },
            { label: mark(pg.orient === 'landscape', '横向'), fn: () => set({ orient: 'landscape' }) }
          ],
          margin: Object.entries(MARGIN_LABEL).map(([k, l]) => ({ label: mark(pg.margin === k, (part ? '' : '页边距 ') + l), fn: () => set({ margin: k }) }))
        }
        const items = part ? parts[part] : [...parts.size, '-', ...parts.orient, '-', ...parts.margin]
        menu(r.left, r.bottom + 4, items.map((x) => (x === '-' ? x : { ...x, disabled: !this.doc.writable })))
      }

      // ------------------------------------------------ context menu

      context (e) {
        e.preventDefault()
        if (e.target.tagName === 'IMG' && e.target.classList.contains('qd-img')) { this.pickImg(e.target); this.imgMenu(e.target, e.clientX, e.clientY); return }
        const revEl = e.target.closest?.('[data-rev],.qd-del')
        if (revEl && this.doc.writable) {
          const id = revEl.dataset.rev ? revEl.dataset.rev.slice(4) : revEl.dataset.del
          menu(e.clientX, e.clientY, [
            { label: '接受此修订', fn: () => this.revOp({ op: 'rev', action: 'accept', id }) },
            { label: '拒绝此修订', fn: () => this.revOp({ op: 'rev', action: 'reject', id }) },
            '-',
            { label: '接受所有修订', fn: () => this.revOp({ op: 'rev', action: 'accept' }) },
            { label: '拒绝所有修订', fn: () => this.revOp({ op: 'rev', action: 'reject' }) }
          ])
          return
        }
        const range = currentRange(this.page)
        const ro = !this.doc.writable
        const sel = range && !range.collapsed
        const block = range ? this.blockOf(range.startContainer) : this.csel?.wrap
        const inTable = block?.classList.contains('qd-t') || !!e.target.closest?.('.qd-t')
        menu(e.clientX, e.clientY, [
          { label: '标注问题…', kbd: '⇧⌘M', fn: () => this.frame.pinHere() },
          '-',
          { label: '剪切', kbd: '⌘X', disabled: ro || !sel, fn: () => { document.execCommand('copy'); const r = currentRange(this.page); if (r) this.deleteRange(r) } },
          { label: '复制', kbd: '⌘C', disabled: !sel, fn: () => document.execCommand('copy') },
          { label: '粘贴', kbd: '⌘V', disabled: ro, fn: async () => { try { const t = await navigator.clipboard.readText(); if (t) this.pasteText(t) } catch { this.frame.toast('浏览器不允许从菜单读取剪贴板，请用 ⌘V 粘贴。', 'info') } } },
          '-',
          { label: '链接…', kbd: '⌘K', disabled: ro, fn: () => this.linkSel() },
          { label: '首行缩进', disabled: ro, fn: () => this.firstLine() },
          { label: '插入图片…', disabled: ro, fn: () => this.pickPicture() },
          { label: '插入表格', disabled: ro, fn: () => this.insertTable() },
          { label: '插入分页符', disabled: ro, fn: () => this.pageBreak() },
          '-',
          { label: '新建批注', kbd: '⌘⌥M', disabled: ro || !sel || inTable, fn: () => this.newComment() },
          { label: this.model.toc ? '更新目录' : '插入目录', disabled: ro, fn: () => (this.model.toc ? this.tocOp({ op: 'toc_update' }) : this.insertToc()) },
          ...(inTable ? ['-',
            { label: '在下方插入行', disabled: ro, fn: () => this.tableRow('below') },
            { label: '在左侧插入列', disabled: ro, fn: () => this.tableCol('left') },
            { label: '在右侧插入列', disabled: ro, fn: () => this.tableCol('right') },
            { label: '删除所在行', disabled: ro, fn: () => this.tableRow('delete') },
            { label: '删除所在列', disabled: ro, fn: () => this.tableCol('delete') },
            '-',
            { label: '合并单元格', disabled: ro || !this.csel, fn: () => this.mergeCells() },
            { label: '拆分单元格…', disabled: ro, fn: () => this.splitCells() },
            { label: '平均分布各列', disabled: ro, fn: () => this.evenColumns() },
            { label: '删除表格', disabled: ro, fn: () => this.deleteTable() }] : [])
        ])
      }

      // ------------------------------------------------ find

      findHosts () {
        const out = []
        for (const b of this.page.children) {
          if (b.classList.contains('qd-p')) { if (!b.dataset.locked) out.push(b) } else for (const td of b.querySelectorAll('td')) out.push(td)
        }
        return out
      }

      texts (host) {
        const out = []
        const w = document.createTreeWalker(host, NodeFilter.SHOW_TEXT)
        for (let t = w.nextNode(); t; t = w.nextNode()) out.push(t)
        return out
      }

      findAll (q) {
        const ql = q.toLowerCase()
        const out = []
        for (const host of this.findHosts()) {
          const text = this.texts(host).map((t) => t.data).join('').toLowerCase()
          for (let i = text.indexOf(ql); i >= 0; i = text.indexOf(ql, i + ql.length)) out.push({ host, start: i, end: i + ql.length })
        }
        return out
      }

      rangeOf (m) {
        if (!m.host.isConnected) return undefined
        const r = document.createRange()
        let pos = 0; let started = false
        for (const t of this.texts(m.host)) {
          const end = pos + t.length
          if (!started && m.start < end) { r.setStart(t, m.start - pos); started = true }
          if (started && m.end <= end) { r.setEnd(t, m.end - pos); return r }
          pos = end
        }
        return undefined
      }

      showMatch (m) {
        const r = this.rangeOf(m)
        if (!r) return
        if (window.CSS?.highlights && window.Highlight) window.CSS.highlights.set('q-find', new Highlight(r))
        else { const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r) }
        const box = r.getBoundingClientRect(); const view = this.scroll.getBoundingClientRect()
        if (box.top < view.top + 20 || box.bottom > view.bottom - 20) this.scroll.scrollTop += box.top - view.top - view.height / 3
      }

      replaceIn (m, q, rep) {
        const r = this.rangeOf(m)
        if (!r || r.toString().toLowerCase() !== q.toLowerCase()) return false
        const block = this.blockOf(m.host)
        if (!block || block.dataset.locked || m.host.isContentEditable === false) return false
        const start = r.startContainer; const off = r.startOffset
        r.deleteContents()
        if (start.nodeType === 3 && start.isConnected) start.insertData(off, rep)
        else r.insertNode(document.createTextNode(rep))
        if (block.classList.contains('qd-p')) this.tidy(block)
        return true
      }

      replaceMatch (m, q, rep) { if (this.doc.writable && this.replaceIn(m, q, rep)) this.changed() }

      replaceAll (q, rep) {
        if (!this.doc.writable) return 0
        let n = 0
        for (const m of this.findAll(q).reverse()) if (this.replaceIn(m, q, rep)) n++
        if (n) this.changed()
        return n
      }

      selectedText () {
        const r = currentRange(this.page)
        return r && !r.collapsed ? r.toString().replace(/​/g, '') : ''
      }

      focus () { window.CSS.highlights?.delete?.('q-find'); this.page.focus({ preventScroll: true }) }

      // ------------------------------------------------ comments

      author () {
        let a = ''
        try { a = localStorage.getItem('quire.author') ?? '' } catch {}
        return a.trim()
      }

      // paintComments marks what each open comment covers.
      paintComments () {
        const hl = window.CSS.highlights
        if (!hl || typeof Highlight === 'undefined') return
        const done = new Set((this.model?.comments ?? []).filter((c) => c.done).map((c) => c.id))
        const all = []; const on = []
        const starts = new Map()
        for (const m of this.page.querySelectorAll('.qd-cm')) {
          if (m.dataset.cm === 's') { starts.set(m.dataset.id, m); continue }
          if (m.dataset.cm !== 'e' || done.has(m.dataset.id)) continue
          const s = starts.get(m.dataset.id)
          if (!s) continue
          const r = document.createRange()
          r.setStartAfter(s); r.setEndBefore(m)
          ;(m.dataset.id === this.activeCm ? on : all).push(r)
        }
        hl.set('q-comment', new Highlight(...all))
        hl.set('q-comment-on', new Highlight(...on))
      }

      activate (id) {
        this.activeCm = id
        this.paintComments()
        this.renderPane()
        const m = this.page.querySelector(`.qd-cm[data-cm="s"][data-id="${CSS_ID(id)}"]`) ?? this.page.querySelector(`.qd-cm[data-id="${CSS_ID(id)}"]`)
        m?.parentElement?.scrollIntoView({ block: 'nearest', behavior: 'smooth' })
        this.cpane.querySelector(`.qd-card[data-id="${CSS_ID(id)}"]`)?.scrollIntoView({ block: 'nearest' })
      }

      newComment () {
        if (!this.doc.writable) return
        const range = currentRange(this.page)
        if (!range || range.collapsed || !range.toString().trim()) { this.frame.toast('请先选中要批注的文字。', 'info'); return }
        const a = this.blockOf(range.startContainer); const b = this.blockOf(range.endContainer)
        if (!a?.classList.contains('qd-p') || !b?.classList.contains('qd-p') || a.dataset.locked || b.dataset.locked || a.dataset.toc || b.dataset.toc) {
          this.frame.toast('批注只能加在正文段落的文字上，表格和目录里暂不支持。', 'info')
          return
        }
        this.draft = { range: range.cloneRange(), quote: range.toString().replace(/\s+/g, ' ').slice(0, 120) }
        this.cmEdit = null
        this.paneOpen = true
        this.renderPane()
        this.cpane.querySelector('.qd-draft textarea')?.focus()
      }

      commitDraft (text) {
        const d = this.draft
        this.draft = null
        text = text.trim()
        const range = d?.range
        if (!text || !range) { this.renderPane(); return }
        if (!range.startContainer.isConnected || !range.endContainer.isConnected || range.collapsed) {
          this.frame.toast('选中的文字已经变了，请重新选择后再批注。', 'warn')
          this.renderPane()
          return
        }
        const id = 'n-' + Math.random().toString(36).slice(2, 10)
        const end = range.cloneRange()
        end.collapse(false)
        end.insertNode(this.marker('r', id))
        end.insertNode(this.marker('e', id))
        const start = range.cloneRange()
        start.collapse(true)
        start.insertNode(this.marker('s', id))
        const author = this.author() || '用户'
        this.struct.push(() => ({ op: 'comment_new', id, text, author }))
        this.changed()
        this.doc.dirty(true)
      }

      // commentOp sends a change to a comment; the page reloads after.
      commentOp (op) {
        if (!this.doc.writable) return
        this.cmEdit = null
        this.struct.push(() => op)
        this.doc.dirty(true)
      }

      renderPane () {
        const pane = this.cpane
        this.paneBtn.classList.toggle('on', !!this.paneOpen)
        this.paneBtn2.classList.toggle('on', !!this.paneOpen)
        pane.hidden = !this.paneOpen
        if (!this.paneOpen) return
        const ro = !this.doc.writable
        const list = this.model?.comments ?? []
        const top = list.filter((c) => !c.parent || !list.some((p) => p.id === c.parent))
        const who = el('input', { class: 'qd-cwho', placeholder: '署名（批注作者）', value: this.author(), title: '新批注和回复使用的署名' })
        who.addEventListener('change', () => { try { localStorage.setItem('quire.author', who.value.trim()) } catch {} })
        const head = el('div', { class: 'qd-chead' }, el('b', {}, `批注 ${top.length}`), who, el('button', { type: 'button', class: 'q-btn q-mini', title: '关闭批注栏', html: '×', onclick: () => { this.paneOpen = false; this.renderPane() } }))
        const nodes = [head]
        const editor = (value, ok, okLabel) => {
          const ta = el('textarea', { class: 'qd-cta', rows: '3', placeholder: '输入批注…' })
          ta.value = value
          const done = () => ok(ta.value)
          ta.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && mod(e)) { e.preventDefault(); done() }
            if (e.key === 'Escape') { e.preventDefault(); this.draft = null; this.cmEdit = null; this.renderPane() }
          })
          return el('div', { class: 'qd-ced' }, ta, el('div', { class: 'qd-cbtns' },
            el('button', { type: 'button', class: 'qd-cok', onclick: done }, okLabel),
            el('button', { type: 'button', onclick: () => { this.draft = null; this.cmEdit = null; this.renderPane() } }, '取消')))
        }
        if (this.draft) {
          nodes.push(el('div', { class: 'qd-card qd-draft on' },
            el('div', { class: 'qd-cmeta' }, el('b', {}, this.author() || '用户'), el('span', {}, '新批注')),
            el('div', { class: 'qd-cquote' }, this.draft.quote),
            editor('', (t) => this.commitDraft(t), '批注')))
        }
        const when = (d) => (d ? String(d).replace('T', ' ').replace(/:\d\d(\.\d+)?Z?$/, '') : '')
        const item = (c, isTop) => {
          const editing = this.cmEdit?.id === c.id && this.cmEdit.mode === 'edit'
          const acts = ro ? null : el('div', { class: 'qd-cacts' },
            isTop ? el('button', { type: 'button', onclick: (e) => { e.stopPropagation(); this.cmEdit = { id: c.id, mode: 'reply' }; this.renderPane(); this.cpane.querySelector('.qd-cta')?.focus() } }, '回复') : null,
            el('button', { type: 'button', onclick: (e) => { e.stopPropagation(); this.cmEdit = { id: c.id, mode: 'edit' }; this.renderPane(); this.cpane.querySelector('.qd-cta')?.focus() } }, '编辑'),
            isTop ? el('button', { type: 'button', onclick: (e) => { e.stopPropagation(); this.commentOp({ op: 'comment_done', id: c.id, done: !c.done }) } }, c.done ? '重新打开' : '解决') : null,
            el('button', { type: 'button', onclick: (e) => { e.stopPropagation(); this.commentOp({ op: 'comment_delete', id: c.id }) } }, '删除'))
          return [
            el('div', { class: 'qd-cmeta' }, el('b', {}, c.author || '未署名'), el('span', {}, when(c.date)), c.done && isTop ? el('em', {}, '已解决') : null),
            isTop && c.quote ? el('div', { class: 'qd-cquote' }, c.quote) : null,
            editing ? editor(c.text, (t) => { if (t.trim() && t !== c.text) this.commentOp({ op: 'comment_edit', id: c.id, text: t }); else { this.cmEdit = null; this.renderPane() } }, '保存') : el('div', { class: 'qd-ctext' }, c.text),
            editing ? null : acts
          ]
        }
        for (const c of top) {
          const replies = list.filter((r) => r.parent === c.id)
          const card = el('div', { class: 'qd-card' + (c.done ? ' done' : '') + (this.activeCm === c.id ? ' on' : ''), dataset: { id: c.id } },
            ...item(c, true),
            ...replies.map((r) => el('div', { class: 'qd-creply' }, ...item(r, false))),
            this.cmEdit?.id === c.id && this.cmEdit.mode === 'reply' ? editor('', (t) => { if (t.trim()) this.commentOp({ op: 'comment_reply', parent: c.id, text: t.trim(), author: this.author() || '用户' }) }, '回复') : null)
          card.addEventListener('click', (e) => { if (!e.target.closest('button,textarea')) this.activate(c.id) })
          nodes.push(card)
        }
        if (!top.length && !this.draft) nodes.push(el('div', { class: 'qd-cnone' }, ro ? '没有批注。' : '没有批注。选中文字后点工具栏的批注按钮或按 ⌘⌥M 添加。'))
        pane.replaceChildren(...nodes)
      }

      // ------------------------------------------------ table of contents

      tocMenu (e) {
        if (!this.model.toc) { this.insertToc(); return }
        const r = e.currentTarget.getBoundingClientRect()
        menu(r.left, r.bottom + 2, [
          { label: '更新目录', fn: () => this.tocOp({ op: 'toc_update' }) },
          { label: '删除目录', fn: () => this.tocOp({ op: 'toc_delete' }) }
        ])
      }

      // insertToc puts a contents list before the paragraph with the caret,
      // or at the top of the document.
      insertToc () {
        const range = currentRange(this.page)
        this.tocCheck = true
        this.structural((b) => ({ op: 'toc', before: b.ref, levels: 3, pages: this.headingPages() }), range ? this.spot() : { node: this.page.firstElementChild })
      }

      async tocOp (op) {
        if (!this.doc.writable) return
        await this.doc.flush()
        this.struct.push(() => (op.op === 'toc_update' ? { ...op, pages: this.headingPages() } : op))
        this.doc.dirty(true)
      }

      // headingPages gives the page each saved heading is laid out on,
      // keyed by block ref, for the contents' page numbers.
      headingPages () {
        const pg = this.model?.page
        if (!pg) return undefined
        const k = this.page.offsetWidth / pg.width || 1
        const per = pg.cols > 1 ? Math.max(100, pg.height - pg.top - pg.bottom) : pg.height + 16
        const out = {}
        for (const b of this.page.children) {
          const ref = this.headingLevel(b) && this.base.get(b.dataset.k)?.ref
          if (ref) out[ref] = Math.floor(b.offsetTop / k / per) + 1
        }
        return out
      }

      // tocDrift is true when the contents' page numbers disagree with
      // where the headings now sit (a new contents pushes them down).
      tocDrift () {
        const shown = [...this.page.querySelectorAll('.qd-tocp')].map((p) => +p.textContent.split('\t').pop())
        const pages = this.headingPages() ?? {}
        const want = [...this.page.children].filter((b) => { const l = this.headingLevel(b); return l && l <= 3 && this.base.get(b.dataset.k) }).map((b) => pages[this.base.get(b.dataset.k).ref])
        return shown.length > 0 && shown.length === want.length && shown.some((n, i) => n !== want[i])
      }

      // ------------------------------------------------ navigation pane
      // Headings in document order, like Word's navigation pane: click to
      // go there, the section in view is lit, levels fold. Open by default
      // when the document has headings, and as last set after that.

      headingLevel (b) {
        if (!b.classList?.contains('qd-p') || b.dataset.toc !== undefined) return 0
        const h = +b.dataset.h || (b.dataset.style?.startsWith('@h') ? +b.dataset.style.slice(2) : 0)
        return h > 0 ? h : 0
      }

      toggleNav (open = !this.navOpen) {
        this.navOpen = open
        try { localStorage.setItem('quire.docx.nav', open ? '1' : '0') } catch {}
        this.nav.hidden = !open
        this.navBtn.classList.toggle('on', open)
        this.navSb.classList.toggle('on', open)
        this.navSig = null
        if (open) this.renderNav()
      }

      renderNav () {
        if (!this.nav) return
        const hs = []
        for (const b of this.page.children) {
          const lv = this.headingLevel(b)
          const t = lv && oneLine(b.textContent)
          if (t) hs.push({ b, lv, t })
        }
        this.navHeads = hs
        if (this.navOpen === undefined) {
          let pref = null
          try { pref = localStorage.getItem('quire.docx.nav') } catch {}
          this.navOpen = pref ? pref === '1' : hs.length > 0
          this.nav.hidden = !this.navOpen
          this.navBtn.classList.toggle('on', this.navOpen)
          this.navSb.classList.toggle('on', this.navOpen)
        }
        if (!this.navOpen) return
        const sig = hs.map((h) => h.b.dataset.k + h.lv + h.t).join('\n')
        if (sig !== this.navSig) {
          this.navSig = sig
          this.navAt = undefined
          this.navCount.textContent = hs.length ? String(hs.length) : ''
          if (!hs.length) this.navList.replaceChildren(el('p', { class: 'qd-nav-e' }, '文档里还没有标题。给段落套用「标题 1」「标题 2」等样式，标题就会列在这里，点一下即可跳转。'))
          else {
            const min = Math.min(...hs.map((h) => h.lv))
            this.navList.replaceChildren(...hs.map((h, i) => {
              const kids = i + 1 < hs.length && hs[i + 1].lv > h.lv
              const tw = el('span', { class: 'qd-nav-tw' + (kids ? '' : ' none'), title: kids ? '折叠或展开' : undefined })
              const it = el('div', { class: 'qd-nav-i', role: 'treeitem', title: h.t, dataset: { lv: String(h.lv), i: String(i) }, style: { paddingLeft: 2 + (h.lv - min) * 14 + 'px' } }, tw, el('span', { class: 'qd-nav-t' }, h.t))
              if (kids) {
                tw.addEventListener('click', (e) => {
                  e.stopPropagation()
                  const k = h.b.dataset.k
                  if (this.navShut.has(k)) this.navShut.delete(k); else this.navShut.add(k)
                  this.navFold()
                })
              }
              it.addEventListener('mousedown', (e) => e.preventDefault())
              it.addEventListener('click', () => this.navGo(i))
              return it
            }))
            this.navFold()
          }
        }
        this.navSpot()
      }

      // navFold hides headings under a folded one.
      navFold () {
        let below = Infinity
        this.navHeads.forEach((h, i) => {
          const it = this.navList.children[i]
          if (!it) return
          const shut = this.navShut.has(h.b.dataset.k)
          it.classList.toggle('shut', shut)
          if (h.lv > below) { it.hidden = true; return }
          it.hidden = false
          below = shut ? h.lv : Infinity
        })
      }

      navGo (i) {
        const h = this.navHeads?.[i]
        if (!h?.b.isConnected) return
        const top = h.b.getBoundingClientRect().top - this.scroll.getBoundingClientRect().top + this.scroll.scrollTop - 16
        this.scroll.scrollTop = Math.max(0, top)
        // The heading picked stays lit until the page scrolls on.
        this.navJump = { i, top: this.scroll.scrollTop }
        const w = document.createTreeWalker(h.b, NodeFilter.SHOW_TEXT)
        const t = w.nextNode()
        setCaret(t ?? h.b, 0)
        this.page.focus({ preventScroll: true })
        this.navSpot()
      }

      // navSpot lights the heading of the part in view.
      navSpot () {
        const hs = this.navHeads
        if (!this.navOpen || !hs?.length || this.nav.hidden) return
        let at = 0
        if (this.navJump && Math.abs(this.scroll.scrollTop - this.navJump.top) < 2) at = this.navJump.i
        else {
          this.navJump = null
          const line = this.scroll.getBoundingClientRect().top + 32
          let lo = 0; let hi = hs.length - 1
          while (lo <= hi) {
            const mid = (lo + hi) >> 1
            if (hs[mid].b.getBoundingClientRect().top <= line) { at = mid; lo = mid + 1 } else hi = mid - 1
          }
        }
        if (at === this.navAt) return
        this.navAt = at
        for (const x of this.navList.querySelectorAll('.qd-nav-i.on')) x.classList.remove('on')
        // A folded heading stands in for the parts inside it.
        let i = at
        while (i > 0 && this.navList.children[i]?.hidden) i--
        const it = this.navList.children[i]
        if (!it) return
        it.classList.add('on')
        const lr = this.navList.getBoundingClientRect(); const ir = it.getBoundingClientRect()
        if (ir.top < lr.top || ir.bottom > lr.bottom) this.navList.scrollTop += ir.top - lr.top - lr.height / 3
      }

      // ------------------------------------------------ pins

      // pinAt selects the text a box on screen spans: from the line at
      // its top left to the line at its bottom right, whole blocks when
      // a corner falls outside them.
      pinAt (b) {
        const hit = []
        for (const n of this.page.children) {
          if (n.nodeType !== 1 || n.offsetParent === null) continue
          const r = n.getBoundingClientRect()
          if (r.bottom < b.top) continue
          if (r.top > b.bottom) break
          if (r.right >= b.left && r.left <= b.right) hit.push(n)
        }
        if (!hit.length) return null
        const A = hit[0]; const B = hit[hit.length - 1]
        const caret = (x, y, n) => {
          const c = document.caretRangeFromPoint?.(x, y)
          return c && n.contains(c.startContainer) ? c : null
        }
        const range = document.createRange()
        if (b.right - b.left < 4 && b.bottom - b.top < 4) {
          const c = caret(b.left, b.top, A)
          if (c) range.setStart(c.startContainer, c.startOffset); else range.setStart(A, 0)
          range.collapse(true)
        } else {
          const a = caret(b.left, b.top, A); const z = caret(b.right, b.bottom, B)
          if (a) range.setStart(a.startContainer, a.startOffset); else range.setStart(A, 0)
          if (z) range.setEnd(z.startContainer, z.startOffset); else range.setEnd(B, B.childNodes.length)
          if (range.collapsed) { range.setStart(A, 0); range.setEnd(B, B.childNodes.length) }
        }
        const sel = window.getSelection()
        sel.removeAllRanges()
        sel.addRange(range)
        return this.pinTarget()
      }

      pinTarget () {
        const range = currentRange(this.page)
        if (!range) return null
        const block = this.blockOf(range.startContainer)
        if (!block) return null
        const td = this.cellOf(range.startContainer)
        const inCell = td && block.contains(td) && td.dataset.r
        let where = this.refNear(block)
        const last = this.blockOf(range.endContainer)
        // A span over several blocks is named by its first and last.
        if (last && last !== block && !range.collapsed) where += '–' + this.refNear(last).replace(' 之后', '')
        else if (inCell) where += ` 第 ${td.dataset.r} 行第 ${td.dataset.c} 列`
        const own = inCell ? td : block
        const whole = () => { const r = document.createRange(); r.selectNodeContents(own); return r }
        const hl = range.collapsed ? whole() : range.cloneRange()
        let rect = hl.getBoundingClientRect()
        if (!rect.width && !rect.height) rect = own.getBoundingClientRect()
        const ob = own.getBoundingClientRect()
        const shot = { left: Math.min(ob.left, rect.left), top: Math.min(ob.top, rect.top), right: Math.max(ob.right, rect.right), bottom: Math.max(ob.bottom, rect.bottom) }
        return { where, quote: oneLine(hl.toString()), rect, range: hl, shot }
      }

      shoot (rect) {
        const nodes = []
        for (const n of this.page.children) {
          if (n.nodeType !== 1 || n.offsetParent === null) continue
          const r = n.getBoundingClientRect()
          if (r.bottom < rect.top) continue
          if (r.top > rect.bottom) break
          nodes.push(n)
        }
        return domSnap(nodes, rect, getComputedStyle(this.page).backgroundColor)
      }

      // refNear is the block's number as office_read gives it; a block
      // not saved yet is named after the one before it.
      refNear (block) {
        let after = false
        for (let n = block; n; n = n.previousElementSibling) {
          const b = n.dataset?.k && this.base.get(n.dataset.k)
          if (b) return `[${b.ref}]${after ? ' 之后' : ''}`
          after = true
        }
        return '[开头]'
      }

      pinSpot (p, byRef) {
        const m = /^\[([pt]\d+)\](?: 之后)?(?:–\[([pt]\d+)\])?(?: 第 (\d+) 行第 (\d+) 列)?/.exec(p.where)
        let node = m && byRef.get(m[1])
        if (!node) return null
        const end = m[2] && byRef.get(m[2])
        if (end) {
          const range = document.createRange()
          range.setStart(node, 0)
          range.setEnd(end, end.childNodes.length)
          return { node, range }
        }
        if (m[3]) node = node.querySelector(`td[data-r="${m[3]}"][data-c="${m[4]}"]`) ?? node
        let range = p.quote && p.quote !== '空' ? textRange(node, p.quote) : null
        if (!range) { range = document.createRange(); range.selectNodeContents(node) }
        return { node, range }
      }

      refMap () {
        const byRef = new Map()
        for (const n of this.page.children) { const b = n.dataset?.k && this.base.get(n.dataset.k); if (b) byRef.set(b.ref, n) }
        return byRef
      }

      showPins (pins) { this.pinList = pins; this.paintPins() }

      paintPins () {
        const ranges = []
        if (this.pinList?.length) {
          const byRef = this.refMap()
          for (const p of this.pinList) { const x = this.pinSpot(p, byRef); if (x) ranges.push(x.range) }
        }
        if (ranges.length) pinRanges.set(this, ranges); else pinRanges.delete(this)
        paintPinRanges()
      }

      revealPin (p) {
        const x = this.pinSpot(p, this.refMap())
        if (!x) return
        x.node.scrollIntoView({ block: 'center' })
        const sel = window.getSelection()
        sel.removeAllRanges()
        sel.addRange(x.range)
        this.page.focus({ preventScroll: true })
      }

      jumpToHeading (text) {
        const want = text.replace(/\s+/g, '')
        const h = [...this.page.querySelectorAll('.qd-p[data-h], .qd-p[data-style^="@h"]')].find((x) => !x.dataset.toc && x.textContent.replace(/\s+/g, '') === want)
        h?.scrollIntoView({ block: 'start', behavior: 'smooth' })
      }

      refreshBar () {
        const range = currentRange(this.page)
        if (!range) return
        const n = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer
        const block = this.blockOf(range.startContainer)
        if (!n || !block) return
        const cs = getComputedStyle(n)
        this.ui.b.classList.toggle('on', +cs.fontWeight >= 600)
        this.ui.i.classList.toggle('on', cs.fontStyle === 'italic')
        this.ui.u.classList.toggle('on', /underline/.test(cs.textDecorationLine))
        this.ui.s.classList.toggle('on', /line-through/.test(cs.textDecorationLine))
        const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2
        this.ui.size.value = [...this.ui.size.options].some((o) => o.value === String(pt)) ? String(pt) : ''
        const isP = block.classList.contains('qd-p')
        const style = isP ? block.dataset.style ?? '' : ''
        if (isP && ![...this.ui.style.options].some((o) => o.value === style)) this.ui.style.append(el('option', { value: style }, this.model.styles?.find((s) => s.id === style)?.name ?? style))
        this.ui.style.value = style
        const a = block.dataset.align
        this.ui.al.classList.toggle('on', a === 'left' || a === 'start')
        this.ui.ac.classList.toggle('on', a === 'center')
        this.ui.ar.classList.toggle('on', a === 'right' || a === 'end')
        this.ui.aj.classList.toggle('on', a === 'both')
        this.ui.bul.classList.toggle('on', block.dataset.list === 'b')
        this.ui.num.classList.toggle('on', block.dataset.list === 'o')
        this.ui.track.classList.toggle('on', !!this.model?.track)
        const inTable = block.classList.contains('qd-t')
        if (this.doc.writable) for (const k of ['rowIns', 'rowDel', 'colInsL', 'colInsR', 'colDel', 'tblDel', 'merge', 'split', 'shade', 'border', 'calign', 'tstyle', 'even']) this.ui[k].disabled = !inTable
        this.tableHint.hidden = inTable
        const font = n.closest?.('[data-font]')?.dataset.font ?? ''
        if (font && ![...this.ui.font.options].some((o) => o.value === font)) this.ui.font.append(el('option', { value: font }, font))
        this.ui.font.value = font
        const vert = n.closest?.('[data-vert]')?.dataset.vert
        this.ui.sup.classList.toggle('on', vert === 'sup')
        this.ui.sub.classList.toggle('on', vert === 'sub')
        this.ui.link.classList.toggle('on', !!n.closest?.('[data-link]'))
        this.ui.fl.classList.toggle('on', isP && +block.dataset.fl > 0)
        this.ui.line.value = isP && [...this.ui.line.options].some((o) => o.value === block.dataset.line) ? block.dataset.line : ''
        const sp = isP && block.dataset.before !== undefined && block.dataset.after !== undefined ? `${block.dataset.before},${block.dataset.after}` : ''
        this.ui.space.value = [...this.ui.space.options].some((o) => o.value === sp) ? sp : ''
      }
    }

    // ---------------------------------------------------------- presentation

    // Effect keys: entrance effects by name, the other classes as cls:name.
    const ANIM_FX = [['none', '无'], ['appear', '进入 · 出现'], ['fade', '进入 · 淡入'], ['fly', '进入 · 飞入'], ['wipe', '进入 · 擦除'], ['zoom', '进入 · 缩放'],
      ['emph:pulse', '强调 · 脉冲'], ['emph:grow', '强调 · 放大/缩小'], ['emph:spin', '强调 · 陀螺旋'], ['emph:transparency', '强调 · 透明'],
      ['exit:appear', '退出 · 消失'], ['exit:fade', '退出 · 淡出'], ['exit:fly', '退出 · 飞出'], ['exit:wipe', '退出 · 擦除'], ['exit:zoom', '退出 · 缩放'],
      ['path:line', '路径 · 直线']]
    const ANIM_NAME = Object.fromEntries(ANIM_FX.map(([k, l]) => [k, l.replace(/^.+ · /, '')]))
    const animKey = (a) => (a.cls && a.cls !== 'entr' ? `${a.cls}:${a.effect}` : a.effect)
    const animOf = (key) => { const [c, e] = key.includes(':') ? key.split(':') : ['entr', key]; return { cls: c, effect: e } }
    const animHasDir = (a) => a && (a.cls === 'path' || (a.cls !== 'emph' && ['fly', 'wipe'].includes(a.effect)))
    const ANIM_DIRS = [['b', '自底部'], ['t', '自顶部'], ['l', '自左侧'], ['r', '自右侧']]
    const EXIT_DIRS = [['b', '到底部'], ['t', '到顶部'], ['l', '到左侧'], ['r', '到右侧']]
    const PATH_DIRS = [['b', '向下'], ['t', '向上'], ['l', '向左'], ['r', '向右']]
    // A shape starts hidden when its first effect is an entrance, and ends
    // hidden when its last is an exit.
    const animStartsHidden = (list, spid) => { const a = (list ?? []).find((x) => x.spid === spid); return !!a && (a.cls ?? 'entr') === 'entr' }
    const animEndsHidden = (list, spid) => { const a = (list ?? []).filter((x) => x.spid === spid).pop(); return a?.cls === 'exit' }
    const ANIM_START = [['click', '单击时'], ['with', '与上一动画同时'], ['after', '上一动画之后']]
    const TRANS_FX = [['none', '无'], ['fade', '淡入淡出'], ['push', '推入'], ['wipe', '擦除'], ['split', '分割'], ['cover', '覆盖'], ['pull', '揭开'], ['zoom', '缩放'], ['dissolve', '溶解']]
    const TRANS_DIRS = [['l', '自右侧'], ['r', '自左侧'], ['u', '自底部'], ['d', '自顶部']]
    const SPLIT_DIRS = [['horz-out', '水平向外'], ['horz-in', '水平向内'], ['vert-out', '垂直向外'], ['vert-in', '垂直向内']]
    const TRANS_SPD = { fast: 500, med: 750, slow: 1000 }
    const BULLETS = [['c:•', '•  实心圆点'], ['c:○', '○  空心圆点'], ['c:■', '■  实心方块'], ['c:◆', '◆  菱形'], ['c:➢', '➢  箭头'], ['c:✓', '✓  对勾'], ['c:–', '–  短横线']]
    const NUMBERS = [['a:arabicPeriod', '1.  2.  3.'], ['a:arabicParenR', '1)  2)  3)'], ['a:alphaUcPeriod', 'A.  B.  C.'], ['a:alphaLcParenR', 'a)  b)  c)'], ['a:romanUcPeriod', 'I.  II.  III.'], ['a:ea1ChsPeriod', '一．二．三．'], ['a:circleNumDbPlain', '①  ②  ③']]
    const LINE_SP = [100, 115, 150, 200, 250, 300]
    const TEXT_FX_LIST = [['shadow', '阴影'], ['outline', '白色填充 · 彩色轮廓'], ['glow', '发光'], ['gradient', '渐变填充'], ['fire', '橙红渐变 · 阴影'], ['neon', '霓虹'], ['none', '清除艺术字效果']]
    const GRADIENTS = [['浅蓝', ['DEEBF7', '5B9BD5', 90]], ['橙黄', ['FFE699', 'ED7D31', 90]], ['青绿', ['E2F0D9', '70AD47', 90]], ['紫红', ['F4B6D2', '7030A0', 45]], ['银灰', ['FFFFFF', 'A6A6A6', 90]], ['深夜', ['2F5597', '0D1A33', 90]], ['日落', ['FFC000', 'C00000', 0]]]
    const THEME_COLORS = [
      ['Office', { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72' }],
      ['Office 2007', { dk1: '000000', lt1: 'FFFFFF', dk2: '1F497D', lt2: 'EEECE1', accent1: '4F81BD', accent2: 'C0504D', accent3: '9BBB59', accent4: '8064A2', accent5: '4BACC6', accent6: 'F79646', hlink: '0000FF', folHlink: '800080' }],
      ['蓝色', { dk1: '000000', lt1: 'FFFFFF', dk2: '17406D', lt2: 'DBEFF9', accent1: '0F6FC6', accent2: '009DD9', accent3: '0BD0D9', accent4: '10CF9B', accent5: '7CCA62', accent6: 'A5C249', hlink: 'F49100', folHlink: '85DFD0' }],
      ['绿色', { dk1: '000000', lt1: 'FFFFFF', dk2: '455F51', lt2: 'E3DED1', accent1: '549E39', accent2: '8AB833', accent3: 'C0CF3A', accent4: '029676', accent5: '4AB5C4', accent6: '0989B1', hlink: '6B9F25', folHlink: 'BA6906' }],
      ['橙色', { dk1: '000000', lt1: 'FFFFFF', dk2: '637052', lt2: 'CCDDEA', accent1: 'E48312', accent2: 'BD582C', accent3: '865640', accent4: '9B8357', accent5: 'C2BC80', accent6: '94A088', hlink: '2998E3', folHlink: '8C8C8C' }],
      ['红色', { dk1: '000000', lt1: 'FFFFFF', dk2: '323232', lt2: 'E5C243', accent1: 'A5300F', accent2: 'D55816', accent3: 'E19825', accent4: 'B19C7D', accent5: '7F5F52', accent6: 'B27D49', hlink: '6B9F25', folHlink: 'B26B02' }],
      ['紫罗兰', { dk1: '000000', lt1: 'FFFFFF', dk2: '373545', lt2: 'DCD8DC', accent1: 'AD84C6', accent2: '8784C7', accent3: '5D739A', accent4: '6997AF', accent5: '84ACB6', accent6: '6F8183', hlink: '69A020', folHlink: '8C8C8C' }],
      ['灰度', { dk1: '000000', lt1: 'FFFFFF', dk2: '000000', lt2: 'F8F8F8', accent1: 'DDDDDD', accent2: 'B2B2B2', accent3: '969696', accent4: '808080', accent5: '5F5F5F', accent6: '4D4D4D', hlink: '5F5F5F', folHlink: '919191' }]
    ]
    const THEME_FONTS = [
      ['Office · 等线', { major: { latin: 'Calibri Light', ea: '等线 Light' }, minor: { latin: 'Calibri', ea: '等线' } }],
      ['微软雅黑', { major: { latin: 'Arial', ea: '微软雅黑' }, minor: { latin: 'Arial', ea: '微软雅黑' } }],
      ['黑体 · 宋体', { major: { latin: 'Arial', ea: '黑体' }, minor: { latin: 'Times New Roman', ea: '宋体' } }],
      ['楷体 · 仿宋', { major: { latin: 'Georgia', ea: '楷体' }, minor: { latin: 'Georgia', ea: '仿宋' } }],
      ['苹方', { major: { latin: 'Helvetica Neue', ea: 'PingFang SC' }, minor: { latin: 'Helvetica Neue', ea: 'PingFang SC' } }]
    ]

    const alphaNum = (n) => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26) } return s }
    const romanNum = (n) => { let s = ''; for (const [v, r] of [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]) while (n >= v) { s += r; n -= v } return s }
    const chsNum = (n) => { const d = '零一二三四五六七八九'; return n < 10 ? d[n] : n < 20 ? '十' + (n % 10 ? d[n % 10] : '') : n < 100 ? d[Math.floor(n / 10)] + '十' + (n % 10 ? d[n % 10] : '') : String(n) }
    const NUMFMT = {
      arabicPeriod: (n) => n + '.', arabicParenR: (n) => n + ')', arabicParenBoth: (n) => `(${n})`, arabicPlain: (n) => String(n),
      alphaLcPeriod: (n) => alphaNum(n) + '.', alphaUcPeriod: (n) => alphaNum(n).toUpperCase() + '.', alphaLcParenR: (n) => alphaNum(n) + ')', alphaUcParenR: (n) => alphaNum(n).toUpperCase() + ')',
      romanLcPeriod: (n) => romanNum(n).toLowerCase() + '.', romanUcPeriod: (n) => romanNum(n) + '.',
      ea1ChsPeriod: (n) => chsNum(n) + '．', ea1ChsPlain: (n) => chsNum(n), circleNumDbPlain: (n) => n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : `(${n})`
    }

    const unionBox = (bs) => {
      const x1 = Math.min(...bs.map((b) => b.x)); const y1 = Math.min(...bs.map((b) => b.y))
      return { x: x1, y: y1, cx: Math.max(...bs.map((b) => b.x + b.cx)) - x1, cy: Math.max(...bs.map((b) => b.y + b.cy)) - y1 }
    }

    // resizeBox drags a handle of a box that may be turned: the pointer's
    // movement is read in the box's own frame and the opposite side stays put.
    function resizeBox (s, h, dx, dy, keep) {
      const r = (s.rot || 0) * Math.PI / 180
      const lx = dx * Math.cos(r) + dy * Math.sin(r); const ly = -dx * Math.sin(r) + dy * Math.cos(r)
      let x = 0; let y = 0; let w = s.cx; let hh = s.cy
      if (h.includes('w')) { const d = Math.min(lx, s.cx - 1); x = d; w = s.cx - d }
      if (h.includes('e')) w = Math.max(1, s.cx + lx)
      if (h.includes('n')) { const d = Math.min(ly, s.cy - 1); y = d; hh = s.cy - d }
      if (h.includes('s')) hh = Math.max(1, s.cy + ly)
      if (keep && h.length === 2 && s.cx && s.cy) {
        const ratio = s.cx / s.cy
        if (w / hh > ratio) { const nw = hh * ratio; if (h.includes('w')) x += w - nw; w = nw } else { const nh = w / ratio; if (h.includes('n')) y += hh - nh; hh = nh }
      }
      const cxl = x + w / 2 - s.cx / 2; const cyl = y + hh / 2 - s.cy / 2
      const ccx = s.x + s.cx / 2 + cxl * Math.cos(r) - cyl * Math.sin(r)
      const ccy = s.y + s.cy / 2 + cxl * Math.sin(r) + cyl * Math.cos(r)
      return { x: ccx - w / 2, y: ccy - hh / 2, cx: w, cy: hh, rot: s.rot }
    }

    // animGroups splits a slide's effects into clicks; each click holds
    // steps that start together, at a delay after the click.
    function animGroups (list) {
      const groups = []
      for (const a of list ?? []) {
        if (a.start === 'click' || !groups.length) groups.push({ auto: !groups.length && a.start !== 'click', steps: [{ at: 0, items: [] }], end: 0 })
        const g = groups[groups.length - 1]
        if (a.start === 'after') g.steps.push({ at: g.end, items: [] })
        const st = g.steps[g.steps.length - 1]
        st.items.push(a)
        g.end = Math.max(g.end, st.at + (a.dur || 500))
      }
      return groups
    }
    const animTargets = (root, spid) => root ? [...root.querySelectorAll(`[data-id="${spid}"],[data-gid="${spid}"]`)] : []

    // runEffect plays one effect on a shape of a slide: an entrance shows
    // it, an exit hides it at the end, emphasis and paths hold where they
    // finish.
    function runEffect (n, a, delay) {
      const base = n.dataset.base ?? (n.dataset.base = n.style.transform || '')
      const slide = n.closest('.slide')
      const W = slide?.offsetWidth ?? 960; const H = slide?.offsetHeight ?? 540
      const cls = a.cls ?? 'entr'
      const d = a.dir ?? 'b'
      const dur = a.effect === 'appear' ? 1 : Math.max(50, a.dur || 500)
      const ease = { easing: 'ease-in-out', delay, duration: dur, fill: 'both' }
      if (cls === 'emph') {
        const frames = a.effect === 'pulse' ? [{ transform: `${base} scale(1)` }, { transform: `${base} scale(1.05)` }, { transform: `${base} scale(1)` }]
          : a.effect === 'grow' ? [{ transform: `${base} scale(1)` }, { transform: `${base} scale(1.5)` }]
            : a.effect === 'spin' ? [{ transform: `${base} rotate(0deg)` }, { transform: `${base} rotate(360deg)` }]
              : [{ opacity: 1 }, { opacity: 0.5 }]
        return n.animate(frames, { ...ease, composite: 'replace' })
      }
      if (cls === 'path') {
        const v = { b: [0, H / 4], t: [0, -H / 4], r: [W / 4, 0], l: [-W / 4, 0] }[d] ?? [0, H / 4]
        return n.animate([{ transform: `translate(0px,0px) ${base}` }, { transform: `translate(${v[0]}px,${v[1]}px) ${base}` }], ease)
      }
      let frames = [{ opacity: 0 }, { opacity: 1 }]
      if (a.effect === 'fly') {
        const dx = d === 'l' ? -(n.offsetLeft + n.offsetWidth) : d === 'r' ? W - n.offsetLeft : 0
        const dy = d === 't' ? -(n.offsetTop + n.offsetHeight) : d === 'b' ? H - n.offsetTop : 0
        frames = [{ transform: `translate(${dx}px,${dy}px) ${base}` }, { transform: `translate(0px,0px) ${base}` }]
      } else if (a.effect === 'wipe') {
        const edge = { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }
        // An exit wipes away toward its edge.
        const from = cls === 'exit' ? { b: edge.t, t: edge.b, l: edge.r, r: edge.l }[d] : edge[d]
        frames = [{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }]
      } else if (a.effect === 'zoom') frames = [{ transform: `${base} scale(.1)`, opacity: 0 }, { transform: `${base} scale(1)`, opacity: 1 }]
      else if (a.effect === 'appear') frames = [{ opacity: 1 }, { opacity: 1 }]
      if (cls === 'exit') {
        const x = n.animate(frames.reverse(), { duration: dur, delay, fill: 'forwards', easing: a.effect === 'fly' ? 'cubic-bezier(.7,0,.8,.3)' : 'ease-in' })
        x.finished.then(() => { n.style.visibility = 'hidden'; x.cancel() }, () => {})
        return x
      }
      n.style.visibility = 'visible'
      return n.animate(frames, { duration: dur, delay, fill: 'backwards', easing: a.effect === 'fly' ? 'cubic-bezier(.2,.7,.3,1)' : 'ease-out' })
    }

    // transitionIn brings slide cur in over old with a slide transition.
    function transitionIn (cur, old, t) {
      const d = TRANS_SPD[t.spd] ?? 750
      const W = cur.offsetWidth || 960; const H = cur.offsetHeight || 540
      const v = { l: [W, 0], r: [-W, 0], u: [0, H], d: [0, -H] }[t.dir ?? 'l'] ?? [W, 0]
      const opt = { duration: d, easing: 'ease-in-out', fill: 'both' }
      const anims = []
      const tr = (x, y) => `translate(${x}px,${y}px)`
      if (t.type === 'fade') anims.push(cur.animate([{ opacity: 0 }, { opacity: 1 }], opt))
      else if (t.type === 'dissolve') anims.push(cur.animate([{ opacity: 0 }, { opacity: 1 }], { ...opt, easing: 'steps(8)' }))
      else if (t.type === 'push') { anims.push(cur.animate([{ transform: tr(...v) }, { transform: tr(0, 0) }], opt)); if (old) anims.push(old.animate([{ transform: tr(0, 0) }, { transform: tr(-v[0], -v[1]) }], opt)) }
      else if (t.type === 'cover') anims.push(cur.animate([{ transform: tr(...v) }, { transform: tr(0, 0) }], opt))
      else if (t.type === 'pull' && old) { old.parentElement?.append(old); anims.push(old.animate([{ transform: tr(0, 0) }, { transform: tr(-v[0], -v[1]) }], opt)) }
      else if (t.type === 'wipe') anims.push(cur.animate([{ clipPath: { l: 'inset(0 0 0 100%)', r: 'inset(0 100% 0 0)', u: 'inset(100% 0 0 0)', d: 'inset(0 0 100% 0)' }[t.dir ?? 'l'] ?? 'inset(0 0 0 100%)' }, { clipPath: 'inset(0 0 0 0)' }], opt))
      else if (t.type === 'split') {
        const vert = t.orient === 'vert'
        const shut = vert ? 'inset(0 50% 0 50%)' : 'inset(50% 0 50% 0)'
        if (t.dir === 'in' && old) { old.parentElement?.append(old); anims.push(old.animate([{ clipPath: 'inset(0 0 0 0)' }, { clipPath: shut }], opt)) } else anims.push(cur.animate([{ clipPath: shut }, { clipPath: 'inset(0 0 0 0)' }], opt))
      } else if (t.type === 'zoom') anims.push(cur.animate([{ transform: 'scale(.3)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }], opt))
      return Promise.all(anims.map((x) => x.finished.catch(() => {}))).then(() => { for (const x of anims) x.cancel() })
    }

    // contentHeight measures the text of a shape, without its insets.
    const contentHeight = (n) => {
      const ps = [...n.children].filter((c) => c.tagName === 'P')
      if (!ps.length) return 0
      const last = ps[ps.length - 1]
      return last.offsetTop + last.offsetHeight - ps[0].offsetTop
    }

    // bgCss turns a slide's background declaration into a CSS value.
    const bgCss = (b) => String(b || '').replace(/^\s*background:\s*/, '').replace(/;\s*$/, '') || '#fff'
    const swatch = (colors) => el('span', { class: 'q-swatches' }, ...colors.map((c) => el('i', { style: { background: '#' + c } })))
    const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
    const SHAPES = [['rect', '矩形'], ['roundRect', '圆角矩形'], ['ellipse', '椭圆'], ['triangle', '三角形'], ['diamond', '菱形'], ['pentagon', '五边形'], ['star5', '五角星'], ['rightArrow', '右箭头'], ['leftArrow', '左箭头']]
    const ARRANGE = [{ value: '', label: '排列' }, { value: 'front', label: '置于顶层' }, { value: 'forward', label: '上移一层' }, { value: 'backward', label: '下移一层' }, { value: 'back', label: '置于底层' }, { value: 'l', label: '左对齐（幻灯片）' }, { value: 'c', label: '水平居中' }, { value: 'r', label: '右对齐（幻灯片）' }, { value: 't', label: '顶端对齐' }, { value: 'm', label: '垂直居中' }, { value: 'b', label: '底端对齐' }]

    class PptxEditor {
      constructor (frame, doc) {
        this.frame = frame
        this.doc = doc
        this.queue = []
        this.zoom = 0
        this.ix = 0
        this.selected = null
        this.editing = null
        this.build()
      }

      build () {
        const tb = new Toolbar(this.frame.tools, this.frame.strip)
        fileTab(tb, this.frame)
        this.tb = tb
        this.ui = {}
        this.mode = 'slides'
        this.mix = 0
        this.multi = []
        this.fitAfter = null
        tb.onShow = (name) => { this.tabName = name; this.paintBadges() }
        tb.tab('开始')
        this.ui.layout = tb.select('新建幻灯片', [{ value: '', label: '＋ 新建幻灯片' }], (v) => { if (v) this.newSlide(v); this.ui.layout.value = '' }, { width: '108px' })
        this.ui.dup = tb.button('', '复制幻灯片', () => this.slideOp({ op: 'duplicate_slide' }, 1), { ico: 'copy' })
        this.ui.del = tb.button('', '删除幻灯片', () => { if (this.model.slides.length > 1 && window.confirm(`删除第 ${this.ix + 1} 张幻灯片？`)) this.slideOp({ op: 'delete_slide' }, 0) }, { ico: 'trash' })
        this.ui.up = tb.button('', '上移', () => { if (this.ix > 0) this.slideOp({ op: 'move_slide', to: this.ix }, -1) }, { ico: 'up' })
        this.ui.down = tb.button('', '下移', () => { if (this.ix < this.model.slides.length - 1) this.slideOp({ op: 'move_slide', to: this.ix + 2 }, 1) }, { ico: 'down' })
        tb.sep()
        this.ui.font = tb.select('字体', fontOptions(), (v) => this.format((s) => { if (v) s.dataset.font = v; else delete s.dataset.font }), { width: '86px' })
        this.ui.size = tb.select('字号', [{ value: '', label: '字号' }, ...[10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 60, 72, 96].map((n) => ({ value: String(n), label: String(n) }))], (v) => { if (v) this.format((s) => { s.dataset.sz = String(+v * 100) }) }, { width: '58px' })
        this.ui.b = tb.button('B', '加粗（⌘B）', () => this.toggle('b'), { cls: 'q-b' })
        this.ui.i = tb.button('I', '倾斜（⌘I）', () => this.toggle('i'), { cls: 'q-i' })
        this.ui.u = tb.button('U', '下划线（⌘U）', () => this.toggle('u'), { cls: 'q-u' })
        tb.color('文字颜色', 'A', (c) => this.format((s) => { s.dataset.c = c }), 'C00000')
        tb.sep()
        this.ui.bu = tb.button('', '项目符号', (e) => this.bulletMenu(e, 'c'), { ico: 'bullet' })
        this.ui.num = tb.button('', '编号', (e) => this.bulletMenu(e, 'a'), { ico: 'number' })
        this.ui.outdent = tb.button('', '降低列表级别（⇧Tab）', () => this.indent(-1), { ico: 'outdent' })
        this.ui.indent = tb.button('', '提高列表级别（Tab）', () => this.indent(1), { ico: 'indent' })
        this.ui.al = tb.button('', '左对齐', () => this.align('l'), { ico: 'left' })
        this.ui.ac = tb.button('', '居中', () => this.align('ctr'), { ico: 'center' })
        this.ui.ar = tb.button('', '右对齐', () => this.align('r'), { ico: 'right' })
        this.ui.ls = tb.select('行距', [{ value: '', label: '行距' }, ...LINE_SP.map((v) => ({ value: String(v), label: String(v / 100) }))], (v) => { this.ui.ls.value = ''; if (v) this.lineSpacing(+v) }, { width: '54px' })
        tb.tab('插入')
        this.ui.box = tb.button('文本框', '插入文本框', () => { this.selectLast = 'edit'; this.push({ op: 'add_textbox', ...this.at(), sz: 1800 }) }, { ico: 'text' })
        this.ui.shape = tb.select('插入形状', [{ value: '', label: '形状' }, ...SHAPES.map(([v, l]) => ({ value: v, label: l }))], (v) => { this.ui.shape.value = ''; if (v) this.addShape(v) }, { width: '62px' })
        this.ui.pic = tb.button('图片', '插入图片', () => this.pickPicture(), { ico: 'image' })
        this.ui.tbl = tb.button('表格', '插入表格（3×3）', () => { this.selectLast = 'select'; this.push({ op: 'add_table', ...this.at(), rows: 3, cols: 3 }) }, { ico: 'table' })
        tb.tab('设计')
        this.ui.layoutSel = tb.select('更换当前幻灯片的版式', [], (v) => { if (v) this.changeLayout(v) }, { width: '112px' })
        this.ui.themeC = tb.button('颜色 ▾', '主题颜色', (e) => this.themeMenu(e, 'colors'))
        this.ui.themeF = tb.button('字体 ▾', '主题字体', (e) => this.themeMenu(e, 'fonts'))
        this.ui.bg = tb.button('背景格式', '设置背景：纯色、渐变、图片', () => this.bgDialog(), { ico: 'image' })
        tb.sep()
        this.ui.masterBtn = tb.button('幻灯片母版', '编辑母版与版式；改动影响用到它们的所有幻灯片', () => this.setMode(this.mode === 'master' ? 'slides' : 'master'))
        tb.tab('形状格式')
        this.fillBtn = tb.color('填充颜色', '◼', (c) => this.styleShape({ fill: c }), '4472C4')
        this.ui.grad = tb.button('渐变 ▾', '渐变填充', (e) => this.gradMenu(e))
        this.lineBtn = tb.color('轮廓颜色', '▢', (c) => this.styleShape({ line: c }), '2F5597')
        this.ui.lineW = tb.select('轮廓粗细', [{ value: '', label: '粗细' }, ...[0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((v) => ({ value: String(v), label: v + ' 磅' }))], (v) => { this.ui.lineW.value = ''; if (v) this.styleShape({ lineW: +v }) }, { width: '58px' })
        this.ui.noFill = tb.button('', '无填充、无轮廓', () => this.styleShape({ fill: null, line: null }), { ico: 'noFill' })
        this.ui.shadow = tb.button('阴影', '添加或去掉阴影', () => this.toggleShadow())
        this.ui.wordart = tb.button('艺术字 ▾', '文字效果：阴影、轮廓、发光、渐变', (e) => this.wordArtMenu(e))
        this.ui.bodyBtn = tb.button('文本框…', '文本框选项：边距、自动调整、文字方向', () => this.bodyDialog())
        tb.sep()
        this.ui.rotate = tb.button('旋转 ▾', '旋转；也可拖动形状上方的圆形控点', (e) => this.rotateMenu(e))
        this.ui.alignBtn = tb.button('对齐 ▾', '对齐与分布', (e) => this.alignMenu(e))
        this.ui.arrange = tb.select('排列', ARRANGE.slice(0, 5), (v) => { this.ui.arrange.value = ''; if (v) this.arrange(v) }, { width: '62px' })
        this.ui.groupBtn = tb.button('组合', '组合所选形状（⌘G）', () => this.group(), { ico: 'group' })
        this.ui.ungroupBtn = tb.button('取消组合', '取消组合（⌘⇧G）', () => this.ungroup())
        this.ui.cropBtn = tb.button('裁剪', '裁剪图片：拖动黑色控点，Enter 完成', () => this.startCrop())
        this.ui.cropAsp = tb.button('比例 ▾', '按比例裁剪图片', (e) => this.cropMenu(e))
        this.ui.rm = tb.button('', '删除形状（Delete）', () => this.deleteShape(), { ico: 'colDel' })
        tb.tab('切换')
        this.ui.trans = tb.select('切换效果', TRANS_FX.map(([v, l]) => ({ value: v, label: l })), (v) => this.setTransition({ type: v }), { width: '86px' })
        this.ui.transDir = tb.select('效果选项', [], (v) => this.setTransition({ dir: v }), { width: '80px' })
        this.ui.transSpd = tb.select('持续时间', [['fast', '快速 0.5 秒'], ['med', '中速 0.75 秒'], ['slow', '慢速 1 秒']].map(([v, l]) => ({ value: v, label: l })), (v) => this.setTransition({ spd: v }), { width: '98px' })
        this.ui.adv = tb.select('换片方式', [['', '单击鼠标时'], ['1', '1 秒后自动'], ['2', '2 秒后自动'], ['3', '3 秒后自动'], ['5', '5 秒后自动'], ['10', '10 秒后自动'], ['30', '30 秒后自动']].map(([v, l]) => ({ value: v, label: l })), (v) => this.setTransition({ adv: v === '' ? null : +v }), { width: '98px' })
        this.ui.transAll = tb.button('全部应用', '把当前幻灯片的切换效果应用到所有幻灯片', () => this.setTransition({}, true))
        this.ui.transPrev = tb.button('预览', '预览切换效果', () => this.previewTransition(), { ico: 'play' })
        tb.tab('动画')
        this.ui.anim = tb.select('动画效果', ANIM_FX.map(([v, l]) => ({ value: v, label: l })), (v) => this.setAnim(v === 'none' ? { effect: 'none' } : animOf(v)), { width: '112px' })
        this.ui.animDir = tb.select('效果选项', ANIM_DIRS.map(([v, l]) => ({ value: v, label: l })), (v) => this.setAnim({ dir: v }), { width: '72px' })
        this.ui.animStart = tb.select('开始', ANIM_START.map(([v, l]) => ({ value: v, label: l })), (v) => this.setAnim({ start: v }), { width: '110px' })
        this.ui.animDur = tb.select('持续时间', [250, 500, 750, 1000, 1500, 2000, 3000].map((v) => ({ value: String(v), label: v / 1000 + ' 秒' })), (v) => this.setAnim({ dur: +v }), { width: '66px' })
        this.ui.animUp = tb.button('', '向前移动动画', () => this.moveAnim(-1), { ico: 'up' })
        this.ui.animDown = tb.button('', '向后移动动画', () => this.moveAnim(1), { ico: 'down' })
        this.paneBtn = tb.button('动画窗格', '显示或隐藏动画窗格', () => { this.apane.hidden = !this.apane.hidden; this.paneBtn.classList.toggle('on', !this.apane.hidden); this.paintPane() })
        this.animPrev = tb.button('预览', '在编辑区预览本页动画', () => this.previewAnims(), { ico: 'play' })
        tb.tab('放映')
        tb.button('从头开始', '从头开始放映（F5）', () => this.play(0))
        tb.button('从当前开始', '从当前幻灯片放映（⇧F5）', () => this.play(this.ix), { ico: 'play' })
        tb.button('演示者视图', '放映时显示备注、计时和下一张；可另开观众窗口（⌥F5）', () => this.play(this.ix, { presenter: true }))
        this.ui.hide = tb.button('隐藏幻灯片', '放映时隐藏', () => this.slideOp({ op: 'hide_slide', hidden: !this.slide().hidden }, 0), { ico: 'eye' })
        tb.tab('视图')
        this.normalBtn = tb.button('普通', '普通视图', () => this.setMode('slides'))
        this.masterView = tb.button('幻灯片母版', '母版视图', () => this.setMode('master'))
        tb.sep()
        this.zoomSel = tb.select('缩放', [{ value: '', label: '适应' }, ...[50, 75, 100, 125, 150, 200].map((n) => ({ value: String(n), label: n + '%' }))], (v) => { this.zoom = v ? +v / 100 : 0; this.resize() }, { width: '62px' })
        this.notesBtn = tb.button('备注', '显示或隐藏备注栏', () => { this.notes.hidden = !this.notes.hidden; this.notesBtn.classList.toggle('on', !this.notes.hidden); this.resize() })
        this.notesBtn.classList.add('on')
        tb.restore('pptx')
        this.buttons = Object.values(this.ui)

        this.mbar = el('div', { class: 'qp-mbar', hidden: true }, el('b', {}, '幻灯片母版'))
        const closeM = el('button', { type: 'button' }, '关闭母版视图')
        closeM.addEventListener('click', () => this.setMode('slides'))
        this.mbar.append(closeM)
        this.thumbList = el('div', { class: 'qp-tlist' })
        this.thumbs = el('div', { class: 'qp-thumbs' }, this.mbar, this.thumbList)
        this.rotH = el('i', { class: 'qp-rot', dataset: { h: 'rot' }, title: '旋转（按住 Shift 以 15° 为步长）' })
        this.box = el('div', { class: 'qp-box', hidden: true }, ...HANDLES.map((h) => el('i', { class: 'qp-h qp-' + h, dataset: { h } })), this.rotH)
        this.slideEl = el('div', { class: 'slide' })
        this.outl = el('div', { class: 'qp-outl' })
        this.guides = el('div', { class: 'qp-guides' })
        this.badges = el('div', { class: 'qp-badges' })
        this.pinLayer = el('div', { class: 'qp-pins' })
        this.scaler = el('div', { class: 'qp-scaler' }, this.slideEl, this.pinLayer, this.outl, this.box, this.guides, this.badges)
        this.stage = el('div', { class: 'qp-stage', tabindex: '0' }, this.scaler)
        this.notes = el('textarea', { class: 'qp-notes', placeholder: '单击此处添加备注', spellcheck: 'false' })
        this.apane = el('div', { class: 'qp-apane', hidden: true })
        this.center = el('div', { class: 'qp-center' }, this.stage, this.notes)
        this.root = el('div', { class: 'qp' }, this.thumbs, this.center, this.apane)
        this.statusText = el('span', { class: 'qp-stxt' })
        this.zoomer = zoomSlider((v) => this.setZoom(v), () => this.setZoom(0))
        this.statusBar = el('div', { class: 'qp-status' }, this.statusText, el('span', { class: 'q-space' }), this.zoomer.node)
        this.frame.main.append(this.root, this.statusBar)

        this.stage.addEventListener('pointerdown', (e) => this.down(e))
        this.stage.addEventListener('dblclick', (e) => this.dbl(e))
        this.stage.addEventListener('keydown', (e) => this.key(e))
        this.stage.addEventListener('paste', (e) => this.paste(e))
        this.stage.addEventListener('contextmenu', (e) => this.context(e))
        this.stage.addEventListener('dragover', (e) => { if ([...(e.dataTransfer?.items ?? [])].some((i) => i.kind === 'file')) e.preventDefault() })
        this.stage.addEventListener('drop', (e) => { e.preventDefault(); const f = imageFile(e.dataTransfer); if (f && this.doc.writable) readImage(f).then((src) => this.addPicture(src), (err) => this.frame.toast(err.message, 'warn')) })
        this.thumbs.addEventListener('contextmenu', (e) => this.thumbMenu(e))
        this.notes.addEventListener('input', () => { clearTimeout(this.notesTimer); this.notesTimer = setTimeout(() => this.saveNotes(), 800) })
        this.notes.addEventListener('blur', () => this.saveNotes())
        this.notes.addEventListener('keydown', (e) => { if (mod(e) && e.key.toLowerCase() === 'z') e.stopPropagation() })
        this.onSel = () => this.refreshBar()
        document.addEventListener('selectionchange', this.onSel)
        this.ro = new ResizeObserver(() => this.resize())
        this.ro.observe(this.stage)
      }

      dispose () {
        document.removeEventListener('selectionchange', this.onSel)
        this.ro.disconnect()
      }

      // ------------------------------------------------ pins

      shapeWords (n) { return oneLine(n.innerText ?? n.textContent) || ({ pic: '图片', tbl: '表格' }[n.dataset.t] ?? '形状') }

      pinTarget () {
        if (this.mode === 'master' || !this.model) return null
        const sl = this.slide()
        if (!sl) return null
        const page = `第 ${sl.n} 页`
        const sel = window.getSelection()
        if (this.editing && sel?.rangeCount && !sel.isCollapsed && this.slideEl.contains(sel.anchorNode)) {
          const r = sel.getRangeAt(0)
          const n = (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement)?.closest?.('[data-k]')
          return { where: n?.dataset.k ? `${page} [s${sl.n}.${n.dataset.k}]` : page, quote: r.toString(), rect: r.getBoundingClientRect() }
        }
        const key = this.selected
        const els = key ? this.itemEls(key) : []
        if (els.length) {
          const rs = els.map((n) => n.getBoundingClientRect())
          const rect = { left: Math.min(...rs.map((r) => r.left)), right: Math.max(...rs.map((r) => r.right)), top: Math.min(...rs.map((r) => r.top)), bottom: Math.max(...rs.map((r) => r.bottom)) }
          return { where: String(key).startsWith('g') ? page : `${page} [s${sl.n}.${key}]`, quote: els.map((n) => this.shapeWords(n)).join(' '), rect }
        }
        const title = [...this.slideEl.querySelectorAll('[data-k]')].map((n) => oneLine(n.innerText ?? n.textContent)).find(Boolean) ?? ''
        const r = this.slideEl.getBoundingClientRect()
        return { where: page, quote: title, rect: { left: r.left + 16, right: r.left + 16, top: r.top + 16, bottom: r.top + 16 }, shot: { left: r.left, top: r.top, right: r.right, bottom: r.bottom } }
      }

      // pinAt marks the shapes a box on screen takes in: those it mostly
      // covers, else the one it sits inside. A box over empty slide is
      // named as a part of the slide, in percent of its width and height.
      pinAt (b) {
        if (this.mode === 'master' || !this.model) return null
        const sl = this.slide()
        if (!sl) return null
        const S = this.slideEl.getBoundingClientRect()
        const x1 = Math.max(b.left, S.left); const y1 = Math.max(b.top, S.top); const x2 = Math.min(b.right, S.right); const y2 = Math.min(b.bottom, S.bottom)
        if (x2 < x1 || y2 < y1) return null
        const page = `第 ${sl.n} 页`
        const items = [...new Set([...this.slideEl.querySelectorAll('[data-k]')].map((n) => this.keyOf(n)))].map((key) => {
          const els = this.itemEls(key)
          const rs = els.map((n) => n.getBoundingClientRect())
          return { key, els, r: { left: Math.min(...rs.map((r) => r.left)), right: Math.max(...rs.map((r) => r.right)), top: Math.min(...rs.map((r) => r.top)), bottom: Math.max(...rs.map((r) => r.bottom)) } }
        }).filter((x) => x.els.length)
        const area = (r) => Math.max(0, r.right - r.left) * Math.max(0, r.bottom - r.top)
        const meet = (r) => area({ left: Math.max(r.left, x1), right: Math.min(r.right, x2), top: Math.max(r.top, y1), bottom: Math.min(r.bottom, y2) })
        const inside = (r) => r.left <= x1 && r.right >= x2 && r.top <= y1 && r.bottom >= y2
        let picked
        if (x2 - x1 < 4 && y2 - y1 < 4) picked = items.filter((x) => inside(x.r)).slice(-1)
        else {
          picked = items.filter((x) => meet(x.r) >= area(x.r) * 0.5)
          if (!picked.length) picked = items.filter((x) => inside(x.r)).sort((a, z) => area(a.r) - area(z.r)).slice(0, 1)
        }
        const keys = picked.map((x) => x.key)
        this.multi = keys
        this.selected = keys[keys.length - 1] ?? null
        this.placeBox(); this.refreshBar(); this.paintPane?.()
        const refs = picked.filter((x) => !String(x.key).startsWith('g')).map((x) => `[s${sl.n}.${x.key}]`)
        const quote = picked.flatMap((x) => x.els.map((n) => this.shapeWords(n))).join(' ')
        const rect = { left: x1, right: x2, top: y1, bottom: y2 }
        if (refs.length) return { where: `${page} ${refs.join(' ')}`, quote, rect }
        if (!picked.length && x2 - x1 < 4 && y2 - y1 < 4) return this.pinTarget()
        const pc = (v, a, w) => Math.round((v - a) / w * 100)
        return { where: `${page} 区域 ${pc(x1, S.left, S.width)}%,${pc(y1, S.top, S.height)}%–${pc(x2, S.left, S.width)}%,${pc(y2, S.top, S.height)}%`, quote, rect }
      }

      shoot (rect) {
        const S = this.slideEl.getBoundingClientRect()
        const r = { left: Math.max(rect.left, S.left), top: Math.max(rect.top, S.top), right: Math.min(rect.right, S.right), bottom: Math.min(rect.bottom, S.bottom) }
        return domSnap([this.slideEl], r, '#fff')
      }

      // pinMark reads a slide pin's spot: the slide, its shapes, a part.
      pinMark (where) {
        const m = /^第 (\d+) 页/.exec(where)
        if (!m) return null
        const ks = [...where.matchAll(/\[s\d+\.(\d+)\]/g)].map((x) => x[1])
        const a = /区域 (\d+)%,(\d+)%–(\d+)%,(\d+)%/.exec(where)
        return { n: +m[1], ks, area: a && a.slice(1).map(Number) }
      }

      showPins (pins) { this.pinList = pins; this.paintPins() }

      paintPins () {
        if (!this.pinLayer || !this.model) return
        const marks = (this.pinList ?? []).map((p) => this.pinMark(p.where)).filter(Boolean)
        const kids = []
        const sl = this.mode === 'master' ? null : this.slide()
        const k = this.model.emu * (this.scale || 1)
        for (const m of marks) {
          if (!sl || m.n !== sl.n) continue
          if (m.area) {
            const [l, t, r, btm] = m.area
            kids.push(el('div', { class: 'qp-pin', style: { left: l + '%', top: t + '%', width: r - l + '%', height: btm - t + '%' } }))
            continue
          }
          if (!m.ks.length) { kids.push(el('div', { class: 'qp-pin qp-pin-all' })); continue }
          for (const key of m.ks) {
            const n = this.shapeEl(key)
            if (!n) continue
            const b = this.boxOf(n)
            kids.push(el('div', { class: 'qp-pin', style: { left: b.x * k - 4 + 'px', top: b.y * k - 4 + 'px', width: b.cx * k + 8 + 'px', height: b.cy * k + 8 + 'px' } }))
          }
        }
        this.pinLayer.replaceChildren(...kids)
        const has = new Set(marks.map((m) => m.n))
        if (this.mode !== 'master') this.model.slides.forEach((x, i) => this.thumbList.children[i]?.classList.toggle('q-pinned', has.has(x.n)))
      }

      revealPin (p) {
        const m = this.pinMark(p.where)
        if (!m) return
        if (this.mode === 'master') this.setMode('slides')
        const i = this.model.slides.findIndex((x) => x.n === m.n)
        if (i < 0) return
        this.show(i)
        if (m.ks[0] && this.shapeEl(m.ks[0])) this.select(m.ks[0])
        this.stage.focus({ preventScroll: true })
      }

      slide () { return this.model.slides[this.ix] }
      cur () { return this.mode === 'master' ? this.model.masters?.[this.mix] : this.model.slides[this.ix] }
      items () { return this.mode === 'master' ? (this.model.masters ?? []) : this.model.slides }
      pos () { return this.mode === 'master' ? this.mix : this.ix }
      // at names where shape edits go: the slide on stage, or the master
      // or layout part in the master view.
      at () { return this.mode === 'master' ? { part: this.cur().part } : { slide: this.slide().n } }
      shapeEl (k) { return this.slideEl.querySelector(`[data-k="${k}"]`) }

      // ------------------------------------------------ rendering

      load (model, { keep } = {}) {
        if (this.editing) { this.endEdit(false) }
        this.model = model
        if (!document.getElementById('quire-pptx-css')) {
          // The deck's own CSS, kept to this editor.
          const scoped = model.css.replace(/(^|\})\s*([^{}]+)\{/g, (m, a, sel) => a + sel.split(',').map((s) => '.qp ' + s.trim()).join(',') + '{')
          document.head.append(el('style', { id: 'quire-pptx-css' }, scoped))
        }
        this.fillLayouts()
        if (!keep) { this.ix = 0; this.mix = 0 }
        if (this.mode === 'master' && !model.masters?.length) this.mode = 'slides'
        this.ix = Math.max(0, Math.min(this.ix, model.slides.length - 1))
        this.mix = Math.max(0, Math.min(this.mix, (model.masters?.length ?? 1) - 1))
        // The notes box shows what the file holds now; nothing typed there
        // is pending once a new model arrives.
        this.notesShown = undefined
        for (const b of this.buttons) b.disabled = !this.doc.writable
        this.notes.readOnly = !this.doc.writable
        this.renderThumbs()
        this.show(this.pos(), keep)
      }

      fillLayouts () {
        const m = this.model
        this.ui.layout.replaceChildren(el('option', { value: '' }, '＋ 新建幻灯片'), ...m.layouts.map((l) => el('option', { value: l.part }, l.name)))
        this.ui.layoutSel.replaceChildren(el('option', { value: '' }, '版式'), ...m.layouts.map((l) => el('option', { value: l.part }, l.name)))
      }

      renderThumbs () {
        const m = this.model
        const master = this.mode === 'master'
        this.root.classList.toggle('qp-master', master)
        this.mbar.hidden = !master
        const tw = 136
        const s = tw / m.width
        this.thumbList.replaceChildren()
        let n = 0
        this.items().forEach((sl, i) => {
          const inner = el('div', { class: 'slide', html: sl.html, style: { width: m.width + 'px', height: m.height + 'px', background: bgCss(sl.bg), transform: `scale(${s})`, transformOrigin: '0 0' } })
          const lay = master && sl.kind === 'layout'
          const w = lay ? 112 : tw
          if (lay) inner.style.transform = `scale(${w / m.width})`
          const t = el('button', { type: 'button', class: 'qp-thumb' + (i === this.pos() ? ' on' : '') + (sl.hidden ? ' hidden' : '') + (lay ? ' layout' : ''), title: master ? (sl.kind === 'master' ? '母版' : `版式：${sl.name}`) : sl.layout || '' },
            el('span', { class: 'qp-num' }, master ? (lay ? '' : '◆') : String(++n)),
            el('span', { class: 'qp-mini', style: { width: w + 'px', height: Math.round(m.height * w / m.width) + 'px' } }, inner),
            !master && (sl.transition?.type && sl.transition.type !== 'none' || sl.anims?.length) ? el('span', { class: 'qp-star', title: '有切换或动画' }, '★') : null)
          t.addEventListener('click', () => this.show(i))
          t.dataset.i = String(i)
          this.thumbList.append(t)
        })
        this.paintPins()
      }

      show (i, keep) {
        if (this.editing) this.endEdit(true)
        if (this.crop) this.endCrop(true)
        this.saveNotes()
        const sel = keep ? this.selected : null
        const multi = keep ? [...this.multi] : []
        if (this.mode === 'master') this.mix = i; else this.ix = i
        if (!keep) this.animSel = undefined
        const sl = this.cur()
        const m = this.model
        Object.assign(this.slideEl.style, { width: m.width + 'px', height: m.height + 'px', background: bgCss(sl.bg) })
        this.slideEl.innerHTML = sl.missing ? '<div class="qp-missing">这张幻灯片的内容缺失</div>' : sl.html
        for (const n of this.slideEl.querySelectorAll('[data-fs]')) n.style.setProperty('--qfs', n.dataset.fs + 'px')
        if (this.mode === 'slides') {
          this.notes.value = sl.notes ?? ''
          this.notesShown = this.notes.value
          this.notesSlide = sl.n
        } else this.notesShown = undefined
        ;[...this.thumbList.children].forEach((t, j) => t.classList.toggle('on', j === i))
        this.thumbList.children[i]?.scrollIntoView?.({ block: 'nearest' })
        const live = multi.filter((k) => this.itemEls(k).length)
        this.multi = live.length ? live : []
        this.selected = sel && this.itemEls(sel).length ? sel : this.multi[this.multi.length - 1] ?? null
        if (this.selected && !this.multi.includes(this.selected)) this.multi = [this.selected]
        this.resize()
        this.refreshBar()
        this.paintPane()
        this.paintPins()
      }

      resize () {
        if (!this.model) return
        const w = this.stage.clientWidth - 32
        const h = this.stage.clientHeight - 32
        const s = this.zoom || Math.max(0.1, Math.min(w / this.model.width, h / this.model.height))
        this.stage.classList.toggle('qp-zoomed', !!this.zoom)
        this.scale = s
        this.slideEl.style.transform = `scale(${s})`
        this.slideEl.style.transformOrigin = '0 0'
        this.scaler.style.width = this.model.width * s + 'px'
        this.scaler.style.height = this.model.height * s + 'px'
        this.zoomer?.update(s * 100)
        this.placeBox()
        this.paintBadges()
        this.paintPins()
      }

      setZoom (v) { this.zoom = v ? v / 100 : 0; syncZoomSel(this.zoomSel, v); this.resize() }

      paintStatus () {
        if (!this.statusText || !this.model) return
        const cur = this.cur()
        const parts = [this.mode === 'master' ? `母版视图 · ${cur?.kind === 'master' ? '幻灯片母版' : `版式：${cur?.name ?? ''}`}` : `幻灯片 ${this.ix + 1} / ${this.model.slides.length}`]
        if (this.mode === 'slides') { const sl = this.slide(); if (sl?.layout) parts.push(sl.layout); if (sl?.hidden) parts.push('放映时隐藏') }
        if (this.multi.length) parts.push(`已选择 ${this.multi.length} 个对象`)
        this.statusText.textContent = parts.join('　·　')
      }

      // ------------------------------------------------ selection

      boxOf (n) { const [x, y, cx, cy] = (n.dataset.box || '0,0,0,0').split(',').map(Number); return { x, y, cx, cy } }

      // Keys name what is selected: 'N' is the shape with data-k N, 'gM'
      // the top-level group M (moved and resized as one).
      itemEls (key) {
        if (!key) return []
        if (String(key).startsWith('g')) return [...this.slideEl.querySelectorAll(`[data-g="${key.slice(1)}"]`)]
        const n = this.shapeEl(key)
        return n ? [n] : []
      }

      itemBox (key) {
        const n = this.itemEls(key)[0]
        if (!n) return null
        if (String(key).startsWith('g')) { const [x, y, cx, cy, rot] = (n.dataset.gbox || '0,0,0,0,0').split(',').map(Number); return { x, y, cx, cy, rot: rot || 0 } }
        return { ...this.boxOf(n), rot: +(n.dataset.rot || 0) }
      }

      keyOf (n) { return n.dataset.g && n.dataset.g !== '0' ? 'g' + n.dataset.g : n.dataset.k }
      free (key) { if (!key) return false; if (String(key).startsWith('g')) return true; const n = this.shapeEl(key); return !!n && !n.dataset.g }
      spidOf (key) { const n = this.itemEls(key)[0]; return n ? (String(key).startsWith('g') ? n.dataset.gid : n.dataset.id) : undefined }

      select (key, { add } = {}) {
        if (add && key) {
          const i = this.multi.indexOf(key)
          if (i >= 0) this.multi.splice(i, 1); else this.multi.push(key)
          this.selected = this.multi[this.multi.length - 1] ?? null
        } else {
          if (key !== this.selected) this.animSel = undefined
          this.selected = key
          this.multi = key ? [key] : []
        }
        this.placeBox()
        this.refreshBar()
        this.paintPane()
      }

      placeBox () {
        this.multi = this.multi.filter((k) => this.itemEls(k).length)
        if (this.selected && !this.itemEls(this.selected).length) this.selected = this.multi[this.multi.length - 1] ?? null
        this.outl.replaceChildren()
        if (!this.multi.length || !this.model) { this.box.hidden = true; return }
        const k = this.model.emu * this.scale
        const boxes = this.multi.map((x) => this.itemBox(x))
        const single = boxes.length === 1
        const b = single ? boxes[0] : { ...unionBox(boxes), rot: 0 }
        if (!single) for (const o of boxes) this.outl.append(el('i', { class: 'qp-ol', style: { left: o.x * k + 'px', top: o.y * k + 'px', width: o.cx * k + 'px', height: o.cy * k + 'px', transform: o.rot ? `rotate(${o.rot}deg)` : '' } }))
        Object.assign(this.box.style, { left: b.x * k + 'px', top: b.y * k + 'px', width: b.cx * k + 'px', height: b.cy * k + 'px', transform: b.rot ? `rotate(${b.rot}deg)` : '' })
        this.box.hidden = false
        const key = this.multi[0]
        this.box.classList.toggle('fixed', !this.doc.writable || (single && !this.free(key)))
        this.box.classList.toggle('multi', !single)
        this.box.classList.toggle('grp', single && String(key).startsWith('g'))
        this.box.classList.toggle('editing', !!this.editing)
        this.box.classList.toggle('cropping', !!this.crop)
      }

      target (e) {
        const n = e.target.closest?.('[data-k]')
        if (!n || !this.slideEl.contains(n)) return null
        return n
      }

      down (e) {
        if (e.button !== 0) return
        if (this.cellEdit?.contains(e.target)) return
        if (this.cellEdit) this.cellEdit.blur()
        if (this.editing && this.editing.node.contains(e.target)) return
        if (this.editing) this.endEdit(true)
        const handle = e.target.dataset?.h
        if (this.crop) {
          if (handle && handle !== 'rot') { this.cropDrag(e, handle); return }
          this.endCrop(true)
        }
        this.stage.focus({ preventScroll: true })
        let key
        let drill = null
        if (handle) key = this.selected
        else {
          const n = this.target(e)
          if (n) {
            key = this.keyOf(n)
            // A second click on a selected group picks the shape inside it;
            // a drag still moves the whole group.
            if (key.startsWith('g') && !e.shiftKey && (this.selected === key || this.shapeEl(this.selected)?.dataset.g === key.slice(1))) drill = n.dataset.k
          } else if (this.box.contains(e.target) && !this.box.hidden) key = this.selected
        }
        // Shift-clicks and drags must not start the browser's own text selection.
        if (e.shiftKey || !key) { e.preventDefault(); window.getSelection()?.removeAllRanges() }
        if (!key) { if (!e.shiftKey) this.select(null); this.marquee(e); return }
        if (e.shiftKey && !handle) { this.select(key, { add: true }); return }
        if (!handle && !this.multi.includes(key)) this.select(key)
        else if (!handle && this.selected !== key && this.multi.length === 1) this.select(key)
        if (!this.doc.writable) { if (drill) this.select(drill); return }
        const keys = handle ? [this.selected] : this.multi.filter((x) => this.free(x))
        if (!keys.length || !this.free(keys[0])) { if (drill) this.select(drill); return }
        e.preventDefault()
        const k = this.model.emu * this.scale
        const ox = e.clientX; const oy = e.clientY
        const starts = new Map(keys.map((x) => [x, this.itemBox(x)]))
        if (handle === 'rot') { this.rotateDrag(keys[0], starts.get(keys[0])); return }
        const u0 = unionBox([...starts.values()])
        const targets = this.snapTargets(new Set(this.multi))
        const pic = this.shapeEl(keys[0])?.dataset.t === 'pic'
        let moved = false
        let nexts = starts
        const move = (ev) => {
          if (!moved && Math.abs(ev.clientX - ox) + Math.abs(ev.clientY - oy) < 3) return
          moved = true
          let dx = (ev.clientX - ox) / k; let dy = (ev.clientY - oy) / k
          if (!handle) {
            if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0 }
            const snap = ev.altKey ? { dx: 0, dy: 0, lines: [] } : this.snap({ ...u0, x: u0.x + dx, y: u0.y + dy }, targets, 'move')
            dx += snap.dx; dy += snap.dy
            this.showGuides(snap.lines)
            nexts = new Map([...starts].map(([x, b]) => [x, { ...b, x: b.x + dx, y: b.y + dy }]))
          } else {
            const s = starts.get(keys[0])
            let nb = resizeBox(s, handle, dx, dy, ev.shiftKey !== (pic && handle.length === 2))
            if (!s.rot && !ev.altKey) {
              const snap = this.snap(nb, targets, handle)
              if (handle.includes('e')) nb.cx += snap.dx
              if (handle.includes('w')) { nb.x += snap.dx; nb.cx -= snap.dx }
              if (handle.includes('s')) nb.cy += snap.dy
              if (handle.includes('n')) { nb.y += snap.dy; nb.cy -= snap.dy }
              nb = { ...nb, cx: Math.max(1, nb.cx), cy: Math.max(1, nb.cy) }
              this.showGuides(snap.lines)
            }
            nexts = new Map([[keys[0], nb]])
          }
          for (const [x, b] of nexts) this.setItemBox(x, b)
          this.placeBox()
        }
        const up = () => {
          window.removeEventListener('pointermove', move)
          this.showGuides([])
          if (moved) for (const [x, b] of nexts) this.commitItem(x, b)
          else if (drill) this.select(drill)
        }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', up, { once: true })
      }

      rotateDrag (key, start) {
        const c = this.box.getBoundingClientRect()
        const cx = c.left + c.width / 2; const cy = c.top + c.height / 2
        let rot = start.rot || 0
        const move = (ev) => {
          let a = Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180 / Math.PI + 90
          a = ((a % 360) + 360) % 360
          if (ev.shiftKey) a = Math.round(a / 15) * 15
          else for (const s of [0, 90, 180, 270, 360]) if (Math.abs(a - s) < 3) a = s
          rot = a % 360
          this.setItemBox(key, { ...start, rot })
          this.placeBox()
          this.showGuides([])
        }
        const up = () => { window.removeEventListener('pointermove', move); if (rot !== (start.rot || 0)) this.commitItem(key, { ...start, rot }) }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', up, { once: true })
      }

      // marquee selects the shapes a dragged rectangle fully covers.
      marquee (e) {
        const r0 = this.scaler.getBoundingClientRect()
        const k = this.model.emu * this.scale
        const ox = e.clientX; const oy = e.clientY
        const add = e.shiftKey
        let rect = null
        const move = (ev) => {
          if (!rect && Math.abs(ev.clientX - ox) + Math.abs(ev.clientY - oy) < 4) return
          if (!rect) { rect = el('div', { class: 'qp-marq' }); this.scaler.append(rect) }
          const x1 = Math.min(ox, ev.clientX) - r0.left; const y1 = Math.min(oy, ev.clientY) - r0.top
          Object.assign(rect.style, { left: x1 + 'px', top: y1 + 'px', width: Math.abs(ev.clientX - ox) + 'px', height: Math.abs(ev.clientY - oy) + 'px' })
        }
        const up = (ev) => {
          window.removeEventListener('pointermove', move)
          if (!rect) return
          rect.remove()
          const x1 = (Math.min(ox, ev.clientX) - r0.left) / k; const y1 = (Math.min(oy, ev.clientY) - r0.top) / k
          const x2 = (Math.max(ox, ev.clientX) - r0.left) / k; const y2 = (Math.max(oy, ev.clientY) - r0.top) / k
          const keys = [...new Set([...this.slideEl.querySelectorAll('[data-k]')].map((n) => this.keyOf(n)))].filter((key) => { const b = this.itemBox(key); return b && b.x >= x1 && b.y >= y1 && b.x + b.cx <= x2 && b.y + b.cy <= y2 })
          this.multi = add ? [...new Set([...this.multi, ...keys])] : keys
          this.selected = this.multi[this.multi.length - 1] ?? null
          this.placeBox(); this.refreshBar(); this.paintPane()
        }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', up, { once: true })
      }

      snapTargets (skip) {
        const W = this.model.width / this.model.emu; const H = this.model.height / this.model.emu
        const x = [0, W / 2, W]; const y = [0, H / 2, H]
        const seen = new Set()
        for (const n of this.slideEl.querySelectorAll('[data-k]')) {
          const key = this.keyOf(n)
          if (skip.has(key) || seen.has(key)) continue
          seen.add(key)
          const b = this.itemBox(key)
          if (!b || b.rot) continue
          x.push(b.x, b.x + b.cx / 2, b.x + b.cx); y.push(b.y, b.y + b.cy / 2, b.y + b.cy)
        }
        return { x, y }
      }

      // snap finds the nearest guide within a few pixels for the edges a
      // move or a handle drags.
      snap (b, t, mode) {
        const tol = 6 / (this.model.emu * this.scale)
        const xs = mode === 'move' ? [b.x, b.x + b.cx / 2, b.x + b.cx] : [mode.includes('w') ? b.x : null, mode.includes('e') ? b.x + b.cx : null]
        const ys = mode === 'move' ? [b.y, b.y + b.cy / 2, b.y + b.cy] : [mode.includes('n') ? b.y : null, mode.includes('s') ? b.y + b.cy : null]
        const best = (cands, lines) => {
          let bd = Infinity; let bv = null
          for (const c of cands) { if (c === null) continue; for (const v of lines) { const d = v - c; if (Math.abs(d) < Math.abs(bd)) { bd = d; bv = v } } }
          return Math.abs(bd) <= tol ? { d: bd, v: bv } : null
        }
        const sx = best(xs, t.x); const sy = best(ys, t.y)
        const lines = []
        if (sx) lines.push({ x: sx.v })
        if (sy) lines.push({ y: sy.v })
        return { dx: sx ? sx.d : 0, dy: sy ? sy.d : 0, lines }
      }

      showGuides (lines) {
        const k = this.model.emu * this.scale
        this.guides.replaceChildren(...lines.map((l) => el('i', { class: 'qp-guide ' + (l.x !== undefined ? 'v' : 'h'), style: l.x !== undefined ? { left: l.x * k + 'px' } : { top: l.y * k + 'px' } })))
      }

      moveTo (n, b, quiet) {
        const k = this.model.emu
        Object.assign(n.style, { left: b.x * k + 'px', top: b.y * k + 'px', width: b.cx * k + 'px', height: b.cy * k + 'px' })
        if (b.rot !== undefined) {
          const flip = /scaleX\(-1\)/.test(n.style.transform)
          n.style.transform = (b.rot ? `rotate(${b.rot}deg)` : '') + (flip ? ' scaleX(-1)' : '')
          if (b.rot) n.dataset.rot = String(+(+b.rot).toFixed(2)); else delete n.dataset.rot
        }
        n.dataset.box = [b.x, b.y, b.cx, b.cy].map(Math.round).join(',')
        if (!quiet) this.placeBox()
      }

      // setItemBox shows a shape or a whole group at a new box.
      setItemBox (key, b) {
        if (!String(key).startsWith('g')) { const n = this.shapeEl(key); if (n) this.moveTo(n, b, true); return }
        const old = this.itemBox(key)
        if (!old) return
        const sx = old.cx ? b.cx / old.cx : 1; const sy = old.cy ? b.cy / old.cy : 1
        for (const n of this.itemEls(key)) {
          const c = this.boxOf(n)
          this.moveTo(n, { x: b.x + (c.x - old.x) * sx, y: b.y + (c.y - old.y) * sy, cx: c.cx * sx, cy: c.cy * sy }, true)
          n.dataset.gbox = [b.x, b.y, b.cx, b.cy].map(Math.round).join(',') + ',' + (old.rot || 0)
        }
      }

      commitItem (key, b) {
        const r = [b.x, b.y, b.cx, b.cy].map(Math.round)
        const op = { op: 'xfrm', ...this.at(), x: r[0], y: r[1], cx: r[2], cy: r[3] }
        if (String(key).startsWith('g')) op.group = +key.slice(1); else { op.shape = +key; if (b.rot !== undefined) op.rot = +(+b.rot).toFixed(2) }
        this.push(op, { quiet: true })
      }

      commitBox (n, b) { this.commitItem(n.dataset.k, b) }

      key (e) {
        if (this.crop) {
          if (e.key === 'Enter') { e.preventDefault(); this.endCrop(true) } else if (e.key === 'Escape') { e.preventDefault(); this.endCrop(false) }
          return
        }
        if (this.editing) {
          if (e.key === 'Escape') { e.preventDefault(); this.endEdit(true); this.stage.focus() }
          if (e.key === 'Tab') { e.preventDefault(); this.indent(e.shiftKey ? -1 : 1); return }
          if (mod(e) && e.key.toLowerCase() === 'a') {
            e.preventDefault()
            const r = document.createRange(); r.selectNodeContents(this.editing.node)
            const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r)
          }
          if (mod(e) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) { e.preventDefault(); this.toggle(e.key.toLowerCase()) }
          if (mod(e) && ['z', 'y'].includes(e.key.toLowerCase())) e.preventDefault()
          return
        }
        if (this.cellEdit) return
        if (mod(e) && e.key.toLowerCase() === 'z') { e.preventDefault(); this.doc.step(e.shiftKey ? 'redo' : 'undo'); return }
        if (mod(e) && e.key.toLowerCase() === 'y') { e.preventDefault(); this.doc.step('redo'); return }
        if (e.key === 'F5' && e.altKey) { e.preventDefault(); this.play(this.ix, { presenter: true }); return }
        if (e.key === 'F5' && e.shiftKey) { e.preventDefault(); this.play(this.ix); return }
        if (e.key === 'F5') { e.preventDefault(); this.play(0); return }
        if (mod(e) && !e.altKey) {
          const k = e.key.toLowerCase()
          if (k === 'a') { e.preventDefault(); this.multi = [...new Set([...this.slideEl.querySelectorAll('[data-k]')].map((n) => this.keyOf(n)))]; this.selected = this.multi[this.multi.length - 1] ?? null; this.placeBox(); this.refreshBar(); return }
          if (k === 'g' && this.doc.writable) { e.preventDefault(); if (e.shiftKey) this.ungroup(); else this.group(); return }
          if (k === 'c' && this.selected) { e.preventDefault(); this.copyShape(false); return }
          if (k === 'x' && this.selected && this.doc.writable) { e.preventDefault(); this.copyShape(true); return }
          if (k === 'd' && this.selected && this.doc.writable) { e.preventDefault(); this.duplicateShape(); return }
          if (k === 'v' && this.doc.writable) {
            // The paste event follows with what the system clipboard holds;
            // without one, paste the copied shape.
            this.pasteSeen = false
            setTimeout(() => { if (!this.pasteSeen && this.clip) this.pasteShape() }, 80)
            return
          }
        }
        const last = this.items().length - 1
        if (e.key === 'PageDown' || (!this.selected && (e.key === 'ArrowDown' || e.key === 'ArrowRight'))) { e.preventDefault(); if (this.pos() < last) this.show(this.pos() + 1); return }
        if (e.key === 'PageUp' || (!this.selected && (e.key === 'ArrowUp' || e.key === 'ArrowLeft'))) { e.preventDefault(); if (this.pos() > 0) this.show(this.pos() - 1); return }
        if (!this.multi.length) return
        if (e.key === 'Escape') { this.select(null); return }
        if (!this.doc.writable) return
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.deleteShape(); return }
        const n = this.selected && this.shapeEl(this.selected)
        if ((e.key === 'Enter' || e.key === 'F2') && n) { e.preventDefault(); this.beginEdit(n); return }
        const arrows = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }
        const keys = this.multi.filter((x) => this.free(x))
        if (arrows[e.key] && keys.length) {
          e.preventDefault()
          const step = (e.shiftKey ? 10 : 1) / this.model.emu
          for (const x of keys) { const b = this.itemBox(x); b.x += arrows[e.key][0] * step; b.y += arrows[e.key][1] * step; this.setItemBox(x, b) }
          this.placeBox()
          clearTimeout(this.nudge)
          this.nudge = setTimeout(() => { for (const x of keys) if (this.itemEls(x).length) this.commitItem(x, this.itemBox(x)) }, 400)
        }
      }

      // ------------------------------------------------ text

      dbl (e) {
        const n = this.target(e)
        if (!n || !this.doc.writable) return
        if (n.dataset.t === 'tbl') { const td = e.target.closest?.('td[data-r]'); if (td && n.contains(td)) this.editCell(n, td); return }
        // A SmartArt node's text is edited in place, like a text box.
        if (n.dataset.saFrame) { const sa = e.target.closest?.('[data-sa]'); if (sa && n.contains(sa)) this.beginEdit(sa, e); return }
        this.beginEdit(n, e)
      }

      beginEdit (n, e) {
        if ((n.dataset.t !== 'txt' && !n.dataset.sa) || this.editing) return
        this.select(n.dataset.k ?? n.closest('[data-k]').dataset.k)
        if (!n.querySelector('p')) n.append(el('p', {}, ZW))
        for (const b of n.querySelectorAll('.bu')) b.contentEditable = 'false'
        this.editing = { node: n, before: JSON.stringify(this.paras(n)) }
        n.classList.add('qp-editing')
        n.contentEditable = 'true'
        n.spellcheck = false
        n.addEventListener('paste', this.onPaste = (ev) => {
          ev.preventDefault()
          const t = ev.clipboardData.getData('text/plain').replace(/\r\n?/g, '\n')
          document.execCommand('insertText', false, t)
        })
        n.focus({ preventScroll: true })
        const pos = e && document.caretPositionFromPoint?.(e.clientX, e.clientY)
        if (pos && n.contains(pos.offsetNode)) setCaret(pos.offsetNode, pos.offset)
        else if (e && document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(e.clientX, e.clientY); if (r && n.contains(r.startContainer)) setCaret(r.startContainer, r.startOffset) } else {
          const r = document.createRange(); r.selectNodeContents(n); const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r)
        }
        this.placeBox()
        this.refreshBar()
      }

      endEdit (save) {
        const ed = this.editing
        if (!ed) return
        this.editing = null
        const n = ed.node
        n.contentEditable = 'false'
        n.classList.remove('qp-editing')
        n.removeEventListener('paste', this.onPaste)
        window.getSelection()?.removeAllRanges()
        const paras = this.paras(n)
        if (save && JSON.stringify(paras) !== ed.before && n.isConnected && n.dataset.sa) {
          this.push({ op: 'smart_text', ...this.at(), shape: +n.closest('[data-k]').dataset.k, node: n.dataset.sa, paras }, { quiet: true })
        } else if (save && JSON.stringify(paras) !== ed.before && n.isConnected) {
          const empty = paras.every((p) => !p.runs.some((r) => r.text))
          n.toggleAttribute('data-empty', empty)
          n.toggleAttribute('data-filled', !empty)
          this.push({ op: 'shape_text', ...this.at(), shape: +n.dataset.k, paras }, { quiet: true })
          this.autoFit(n)
        }
        this.placeBox()
      }

      // autoFit keeps text that is set to fit: the shape grows or shrinks
      // with its text, or the text shrinks into the shape.
      autoFit (n) {
        if (!n?.isConnected || !this.doc.writable) return
        const cs = getComputedStyle(n)
        const pad = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom)
        const need = contentHeight(n)
        if (n.dataset.fit === 'shape' && !n.dataset.vert) {
          const b = this.boxOf(n)
          const cy = Math.round((need + pad) / this.model.emu)
          if (Math.abs(cy - b.cy) > 6000) { const nb = { ...b, cy, rot: +(n.dataset.rot || 0) }; this.moveTo(n, nb); this.commitItem(n.dataset.k, nb) }
        } else if (n.dataset.fit === 'shrink') {
          const s = +(n.dataset.scale || 1)
          const avail = n.clientHeight - pad
          if (need > avail + 1 || (s < 1 && need < avail * 0.85)) {
            const ns = Math.max(0.25, Math.min(1, s * avail / Math.max(1, need) * 0.97))
            if (Math.abs(ns - s) > 0.02) { this.renderAfter = true; this.selectAfter = n.dataset.k; this.push({ op: 'body', ...this.at(), shape: +n.dataset.k, scale: +ns.toFixed(3) }, { quiet: true }) }
          }
        }
      }

      // paras reads a text shape back: paragraphs of runs with the
      // properties each run sets itself.
      paras (n) {
        const out = []
        let cur
        const para = (p) => {
          cur = { runs: [] }
          if (p?.dataset?.algn) cur.algn = p.dataset.algn
          cur.lvl = +(p?.dataset?.lv ?? 0)
          if (p?.dataset?.buSet) cur.bu = p.dataset.bu
          if (p?.dataset?.lsSet) cur.ls = +p.dataset.ls
          out.push(cur)
        }
        const add = (text, props) => {
          text = text.replace(/​/g, '').replace(/ /g, ' ')
          if (!text) return
          if (!cur) para()
          const last = cur.runs[cur.runs.length - 1]
          if (last && ['b', 'i', 'u', 'sz', 'color', 'font'].every((k) => last[k] === props[k])) last.text += text
          else cur.runs.push({ text, ...props })
        }
        const propsOf = (s, base) => {
          const p = { ...base }
          const d = s.dataset
          if (d.b !== undefined) p.b = d.b === '1' || d.b === 'true'
          if (d.i !== undefined) p.i = d.i === '1' || d.i === 'true'
          if (d.u !== undefined) p.u = d.u === '1'
          if (d.sz !== undefined) p.sz = +d.sz
          if (d.c) p.color = d.c.toUpperCase()
          if (d.font) p.font = d.font
          return p
        }
        const walk = (node, props) => {
          for (const c of node.childNodes) {
            if (c.nodeType === 3) add(c.data, props)
            else if (c.nodeType !== 1 || c.classList.contains('bu')) continue
            else if (c.tagName === 'BR') add('\n', props)
            else if (c.tagName === 'P' || c.tagName === 'DIV') { para(c); walk(c, props) } else walk(c, propsOf(c, props))
          }
        }
        walk(n, {})
        if (!out.length) out.push({ runs: [], lvl: 0 })
        for (const p of out) {
          for (const r of p.runs) for (const k of Object.keys(r)) if (r[k] === undefined) delete r[k]
          const last = p.runs[p.runs.length - 1]
          if (last && last.text.endsWith('\n') && p.runs.length) last.text = last.text.replace(/\n$/, '')
        }
        return out
      }

      paintRun (s) {
        const d = s.dataset
        s.style.fontWeight = d.b === '1' ? '700' : d.b === '0' ? '400' : s.style.fontWeight
        s.style.fontStyle = d.i === '1' ? 'italic' : d.i === '0' ? 'normal' : s.style.fontStyle
        s.style.textDecoration = d.u === '1' ? 'underline' : d.u === '0' ? 'none' : s.style.textDecoration
        if (d.sz) s.style.fontSize = (+d.sz / 100 * this.model.emu * 12700).toFixed(1) + 'px'
        if (d.c) s.style.color = '#' + d.c
        if (d.font) s.style.fontFamily = fontCss(d.font)
      }

      // format changes the selected text while editing, or all the text of
      // the selected shape.
      format (apply) {
        if (!this.doc.writable) return
        if (this.editing) {
          const n = this.editing.node
          const range = currentRange(n)
          if (!range) return
          if (range.collapsed) {
            const p = (range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer).closest('p') ?? n
            const s = caretRun(range, p, () => el('span'))
            apply(s); this.paintRun(s)
            return
          }
          const texts = textsIn(range, n, (t) => !t.parentElement.closest('.bu'))
          for (const t of texts) { const s = isolate(t, t.parentElement.closest('p') ?? n, ''); apply(s); this.paintRun(s) }
          selectTexts(texts)
          this.refreshBar()
          return
        }
        const n = this.selected && this.shapeEl(this.selected)
        if (!n || n.dataset.t !== 'txt') return
        const walk = document.createTreeWalker(n, NodeFilter.SHOW_TEXT)
        const texts = []
        for (let t = walk.nextNode(); t; t = walk.nextNode()) if (t.data.replace(/​/g, '') && !t.parentElement.closest('.bu')) texts.push(t)
        if (!texts.length) return
        for (const t of texts) { const s = isolate(t, t.parentElement.closest('p') ?? n, ''); apply(s); this.paintRun(s) }
        this.push({ op: 'shape_text', ...this.at(), shape: +n.dataset.k, paras: this.paras(n) }, { quiet: true })
      }

      toggle (k) {
        const test = { b: (cs) => +cs.fontWeight >= 600, i: (cs) => cs.fontStyle === 'italic', u: (cs) => /underline/.test(cs.textDecorationLine) }[k]
        let nodes = []
        const root = this.editing?.node ?? (this.selected && this.shapeEl(this.selected))
        if (!root) return
        const range = this.editing && currentRange(root)
        if (range && !range.collapsed) nodes = textsIn(range.cloneRange(), root, (t) => !t.parentElement.closest('.bu'))
        else if (range) nodes = [range.startContainer]
        else { const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); for (let t = w.nextNode(); t; t = w.nextNode()) if (t.data.trim()) nodes.push(t) }
        const on = nodes.length > 0 && nodes.every((t) => test(getComputedStyle(t.nodeType === 3 ? t.parentElement : t)))
        this.format((s) => { s.dataset[k] = on ? '0' : '1' })
      }

      align (a) {
        const n = this.editing?.node ?? (this.selected && this.shapeEl(this.selected))
        if (!n || n.dataset.t !== 'txt' || !this.doc.writable) return
        let ps = [...n.querySelectorAll('p')]
        const range = this.editing && currentRange(n)
        if (range) ps = ps.filter((p) => range.intersectsNode(p))
        const css = { l: 'left', ctr: 'center', r: 'right' }[a]
        for (const p of ps) { p.dataset.algn = a; p.style.textAlign = css }
        if (!this.editing) this.push({ op: 'shape_text', ...this.at(), shape: +n.dataset.k, paras: this.paras(n) }, { quiet: true })
        this.refreshBar()
      }

      refreshBar () {
        if (!this.model) return
        this.paintStatus()
        const ro = !this.doc.writable
        const master = this.mode === 'master'
        const n = this.selected && this.shapeEl(this.selected)
        const txt = n?.dataset.t === 'txt'
        const many = this.multi.length
        for (const k of ['b', 'i', 'u', 'size', 'al', 'ac', 'ar', 'font', 'bu', 'num', 'outdent', 'indent', 'ls', 'wordart', 'bodyBtn']) this.ui[k].disabled = ro || !txt
        this.ui.rm.disabled = ro || !many
        const styled = !ro && n && (n.dataset.t === 'txt' || n.dataset.t === 'shape')
        this.ui.noFill.disabled = !styled
        this.ui.grad.disabled = !styled
        this.ui.lineW.disabled = ro || !n || !['txt', 'shape', 'pic'].includes(n.dataset.t)
        this.ui.shadow.disabled = ro || !n || !['txt', 'shape', 'pic'].includes(n.dataset.t)
        this.ui.shadow.classList.toggle('on', !!n?.style.boxShadow)
        this.ui.arrange.disabled = ro || many !== 1 || !this.free(this.selected)
        this.ui.alignBtn.disabled = ro || !many || this.multi.some((k) => !this.free(k))
        this.ui.rotate.disabled = ro || !this.multi.some((k) => !String(k).startsWith('g') && this.free(k))
        this.ui.groupBtn.disabled = ro || many < 2
        this.ui.ungroupBtn.disabled = ro || !String(this.selected ?? '').startsWith('g')
        this.ui.cropBtn.disabled = ro || n?.dataset.t !== 'pic' || many > 1
        this.ui.cropAsp.disabled = this.ui.cropBtn.disabled
        this.ui.cropBtn.classList.toggle('on', !!this.crop)
        for (const b of [this.fillBtn, this.lineBtn]) b.classList.toggle('q-off', !styled)
        // Slide-only controls rest while the master is on stage.
        for (const k of ['layout', 'dup', 'del', 'up', 'down', 'hide', 'layoutSel', 'trans', 'transDir', 'transSpd', 'adv', 'transAll']) this.ui[k].disabled = ro || master
        if (!master) {
          this.ui.del.disabled = ro || this.model.slides.length < 2
          this.ui.up.disabled = ro || this.ix === 0
          this.ui.down.disabled = ro || this.ix >= this.model.slides.length - 1
        }
        this.ui.hide.classList.toggle('on', !!this.slide()?.hidden)
        this.ui.masterBtn.classList.toggle('on', master)
        this.masterView.classList.toggle('on', master)
        this.normalBtn.classList.toggle('on', !master)
        this.ui.layoutSel.value = master ? '' : this.slide()?.layoutPart ?? ''
        this.refreshMotion()
        if (!txt) { for (const k of ['b', 'i', 'u', 'al', 'ac', 'ar']) this.ui[k].classList.remove('on'); return }
        let probe = n.querySelector('span:not(.bu)') ?? n
        const range = this.editing && currentRange(n)
        if (range) probe = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer
        const cs = getComputedStyle(probe)
        this.ui.b.classList.toggle('on', +cs.fontWeight >= 600)
        this.ui.i.classList.toggle('on', cs.fontStyle === 'italic')
        this.ui.u.classList.toggle('on', /underline/.test(cs.textDecorationLine))
        const pt = Math.round(parseFloat(cs.fontSize) / (this.model.emu * 12700))
        this.ui.size.value = [...this.ui.size.options].some((o) => o.value === String(pt)) ? String(pt) : ''
        const font = probe.closest?.('[data-font]')?.dataset.font ?? ''
        if (font && ![...this.ui.font.options].some((o) => o.value === font)) this.ui.font.append(el('option', { value: font }, font))
        this.ui.font.value = font
        const p = probe.closest?.('p')
        const ta = p ? getComputedStyle(p).textAlign : ''
        this.ui.al.classList.toggle('on', ta === 'left' || ta === 'start')
        this.ui.ac.classList.toggle('on', ta === 'center')
        this.ui.ar.classList.toggle('on', ta === 'right' || ta === 'end')
        const bu = p?.dataset.bu ?? ''
        this.ui.bu.classList.toggle('on', bu.startsWith('c:'))
        this.ui.num.classList.toggle('on', bu.startsWith('a:'))
      }

      // refreshMotion shows the transition and the selected effect.
      refreshMotion () {
        const ro = !this.doc.writable
        const master = this.mode === 'master'
        const sl = this.slide()
        const t = sl?.transition ?? { type: 'none' }
        const known = TRANS_FX.some(([v]) => v === t.type)
        if (!known && t.type && ![...this.ui.trans.options].some((o) => o.value === t.type)) this.ui.trans.append(el('option', { value: t.type }, 'PowerPoint 效果'))
        this.ui.trans.value = t.type ?? 'none'
        const dirs = t.type === 'split' ? SPLIT_DIRS : ['push', 'wipe', 'cover', 'pull'].includes(t.type) ? TRANS_DIRS : []
        this.ui.transDir.replaceChildren(...dirs.map(([v, l]) => el('option', { value: v }, l)))
        this.ui.transDir.value = t.type === 'split' ? `${t.orient ?? 'horz'}-${t.dir ?? 'out'}` : t.dir ?? 'l'
        this.ui.transDir.disabled = ro || master || !dirs.length
        this.ui.transSpd.value = t.spd ?? 'med'
        this.ui.adv.value = t.adv === undefined ? '' : String(t.adv)
        if (t.adv !== undefined && this.ui.adv.value !== String(t.adv)) { this.ui.adv.append(el('option', { value: String(t.adv) }, `${t.adv} 秒后自动`)); this.ui.adv.value = String(t.adv) }
        this.transPrevOk = !master
        this.ui.transPrev.disabled = master || !known || t.type === 'none'
        const a = this.curAnim()
        const can = !ro && !master && !sl?.animLocked && !!this.multi.length
        this.ui.anim.value = a ? animKey(a) : 'none'
        this.ui.anim.disabled = !can
        const adirs = a?.cls === 'exit' ? EXIT_DIRS : a?.cls === 'path' ? PATH_DIRS : ANIM_DIRS
        this.ui.animDir.querySelectorAll('option').forEach((o, i) => { o.textContent = adirs[i]?.[1] ?? o.textContent })
        this.ui.animDir.value = a?.dir ?? 'b'
        this.ui.animDir.disabled = !can || !animHasDir(a)
        this.ui.animStart.value = a?.start ?? 'click'
        this.ui.animDur.value = String(a?.dur ?? 500)
        if (a && this.ui.animDur.value !== String(a.dur)) { this.ui.animDur.append(el('option', { value: String(a.dur) }, a.dur / 1000 + ' 秒')); this.ui.animDur.value = String(a.dur) }
        for (const k of ['animStart', 'animDur', 'animUp', 'animDown']) this.ui[k].disabled = !can || !a
        this.animPrev.disabled = master || !sl?.anims?.length
        this.paneBtn.disabled = master
      }

      // curAnim is the effect the animation controls edit: the one picked in
      // the pane, else the first one of the selected shape.
      curAnim () {
        const list = this.slide()?.anims ?? []
        if (this.mode === 'master') return undefined
        if (this.animSel !== undefined && list[this.animSel]) return list[this.animSel]
        const spid = this.selected && this.spidOf(this.selected)
        const i = list.findIndex((a) => a.spid === spid)
        return i >= 0 ? list[i] : undefined
      }

      // ------------------------------------------------ inserting

      addShape (prst) {
        const W = this.model.width / this.model.emu; const H = this.model.height / this.model.emu
        const cx = Math.round(W * 0.2); const cy = Math.round(prst.endsWith('Arrow') ? H * 0.16 : Math.min(H * 0.3, cx))
        this.selectLast = 'select'
        this.push({ op: 'add_shape', ...this.at(), prst, x: Math.round((W - cx) / 2), y: Math.round((H - cy) / 2), cx, cy })
      }

      async pickPicture () {
        if (!this.doc.writable) return
        const src = await pickImage()
        if (src) this.addPicture(src)
      }

      addPicture (src) {
        // Keep the picture's proportions, fitted to the middle of the slide.
        const img = new Image()
        img.onload = () => {
          const W = this.model.width / this.model.emu; const H = this.model.height / this.model.emu
          const r = img.naturalWidth / Math.max(1, img.naturalHeight)
          let cx = W * 0.6; let cy = cx / r
          if (cy > H * 0.7) { cy = H * 0.7; cx = cy * r }
          this.selectLast = 'select'
          this.push({ op: 'add_image', ...this.at(), src, x: Math.round((W - cx) / 2), y: Math.round((H - cy) / 2), cx: Math.round(cx), cy: Math.round(cy) })
        }
        img.onerror = () => this.frame.toast('无法读取这张图片。', 'warn')
        img.src = src
      }

      styleShape (props) {
        const n = this.selected && this.shapeEl(this.selected)
        const picOk = n?.dataset.t === 'pic' && Object.keys(props).every((k) => ['line', 'lineW', 'shadow'].includes(k))
        if (!n || !this.doc.writable || !(['txt', 'shape'].includes(n.dataset.t) || picOk)) { this.frame.toast('请先选中一个形状或文本框。', 'info'); return }
        this.push({ op: 'shape_style', ...this.at(), shape: +n.dataset.k, ...props })
        this.selectAfter = n.dataset.k
      }

      toggleShadow () {
        const n = this.selected && this.shapeEl(this.selected)
        if (n) this.styleShape({ shadow: !n.style.boxShadow })
      }

      // menuAt opens a menu under the toolbar button that was clicked.
      menuAt (e, items) {
        const r = e?.currentTarget?.getBoundingClientRect?.() ?? e?.target?.getBoundingClientRect?.()
        return menu(r ? r.left : e.clientX, r ? r.bottom + 2 : e.clientY, items)
      }

      gradMenu (e) {
        const n = this.selected && this.shapeEl(this.selected)
        if (!n) return
        const fill = this.fillBtn.querySelector('input').value.slice(1).toUpperCase()
        const light = fill.match(/../g).map((x) => Math.round(parseInt(x, 16) + (255 - parseInt(x, 16)) * 0.6).toString(16).padStart(2, '0')).join('').toUpperCase()
        this.menuAt(e, [
          { label: '由填充色生成', pre: swatch([light, fill]), fn: () => this.styleShape({ gradient: [light, fill, 90] }) },
          '-',
          ...GRADIENTS.map(([l, g]) => ({ label: l, pre: swatch(g.slice(0, 2)), fn: () => this.styleShape({ gradient: g }) })),
          '-',
          { label: '纯色（去掉渐变）', fn: () => this.styleShape({ fill }) }
        ])
      }

      wordArtMenu (e) {
        const n = this.selected && this.shapeEl(this.selected)
        if (!n || n.dataset.t !== 'txt') return
        this.menuAt(e, TEXT_FX_LIST.map(([v, l]) => ({ label: l, pre: el('b', { class: 'qp-fx qp-fx-' + v }, 'A'), fn: () => { this.renderAfter = true; this.selectAfter = n.dataset.k; this.push({ op: 'text_effect', ...this.at(), shape: +n.dataset.k, preset: v }) } })))
      }

      async bodyDialog () {
        const n = this.selected && this.shapeEl(this.selected)
        if (!n || n.dataset.t !== 'txt' || !this.doc.writable) return
        const ins = (n.dataset.ins || '91440,45720,91440,45720').split(',').map((v) => (+v / 360000).toFixed(2))
        const v = await dialog({
          title: '文本框选项',
          fields: [
            { key: 'fit', label: '自动调整', type: 'select', value: n.dataset.fit || 'none', options: [['none', '不自动调整'], ['shrink', '溢出时缩小文字'], ['shape', '根据文字调整形状大小']] },
            { key: 'anchor', label: '垂直对齐', type: 'select', value: n.dataset.anchor || 't', options: [['t', '顶端'], ['ctr', '中部'], ['b', '底端']] },
            { key: 'vert', label: '文字方向', type: 'select', value: n.dataset.vert ? 'eaVert' : 'horz', options: [['horz', '横排'], ['eaVert', '竖排']] },
            { key: 'l', label: '左边距', type: 'number', value: ins[0], hint: '厘米' },
            { key: 'r', label: '右边距', type: 'number', value: ins[2] },
            { key: 't', label: '上边距', type: 'number', value: ins[1] },
            { key: 'b', label: '下边距', type: 'number', value: ins[3] },
            { key: 'wrap', label: '自动换行', type: 'checkbox', value: !n.dataset.nowrap }
          ],
          check: (x) => ['l', 't', 'r', 'b'].some((k) => !(+x[k] >= 0 && +x[k] <= 10)) ? '边距应在 0 到 10 厘米之间' : ''
        })
        if (!v) return
        this.renderAfter = true
        this.selectAfter = n.dataset.k
        if (v.fit !== 'none') this.fitAfter = n.dataset.k
        this.push({ op: 'body', ...this.at(), shape: +n.dataset.k, fit: v.fit, anchor: v.anchor, vert: v.vert, wrap: v.wrap, ins: [v.l, v.t, v.r, v.b].map((x) => Math.round(+x * 360000)) })
      }

      // ------------------------------------------------ paragraphs

      paraTargets () {
        const n = this.editing?.node ?? (this.selected && this.shapeEl(this.selected))
        if (!n || n.dataset.t !== 'txt' || !this.doc.writable) return null
        if (!n.querySelector('p')) n.append(el('p', {}, ZW))
        let ps = [...n.querySelectorAll('p')]
        const range = this.editing && currentRange(n)
        if (range) ps = ps.filter((p) => range.intersectsNode(p))
        return ps.length ? { n, ps } : null
      }

      paraDone (n) {
        if (!this.editing) { this.renderAfter = true; this.selectAfter = n.dataset.k; this.push({ op: 'shape_text', ...this.at(), shape: +n.dataset.k, paras: this.paras(n) }, { quiet: true }) }
        this.refreshBar()
      }

      bulletMenu (e, kind) {
        const list = kind === 'c' ? BULLETS : NUMBERS
        this.menuAt(e, [{ label: '无', fn: () => this.setBullets('none') }, '-', ...list.map(([v, l]) => ({ label: l, fn: () => this.setBullets(v) }))])
      }

      setBullets (bu) {
        const t = this.paraTargets()
        if (!t) return
        for (const p of t.ps) { p.dataset.bu = bu; p.dataset.buSet = '1' }
        this.paintBullets(t.n)
        this.paraDone(t.n)
      }

      // paintBullets redraws the bullets paragraphs set themselves.
      paintBullets (n) {
        const em = this.model.emu
        const counters = new Map()
        for (const p of n.querySelectorAll('p')) {
          const lv = +(p.dataset.lv || 0)
          for (const key of [...counters.keys()]) if (key > lv) counters.delete(key)
          const bu = p.dataset.bu
          if (!bu) continue
          p.querySelector('.bu')?.remove()
          let glyph = ''
          if (bu.startsWith('c:')) { glyph = bu.slice(2); counters.delete(lv) } else if (bu.startsWith('a:')) {
            const k = (counters.get(lv) ?? 0) + 1
            counters.set(lv, k)
            glyph = (NUMFMT[bu.slice(2)] ?? NUMFMT.arabicPeriod)(k)
          } else counters.delete(lv)
          if (glyph) {
            const s = el('span', { class: 'bu' }, glyph)
            if (this.editing) s.contentEditable = 'false'
            p.prepend(s)
            if (!parseFloat(p.style.paddingLeft)) { p.style.paddingLeft = (342900 + lv * 457200) * em + 'px'; p.style.textIndent = -342900 * em + 'px' }
          } else p.style.textIndent = '0px'
        }
      }

      indent (d) {
        const t = this.paraTargets()
        if (!t) return
        const step = 457200 * this.model.emu
        for (const p of t.ps) {
          const lv = Math.max(0, Math.min(8, +(p.dataset.lv || 0) + d))
          if (lv === +(p.dataset.lv || 0)) continue
          p.dataset.lv = String(lv)
          p.style.paddingLeft = Math.max(0, (parseFloat(p.style.paddingLeft) || 0) + d * step) + 'px'
        }
        this.paintBullets(t.n)
        this.paraDone(t.n)
      }

      lineSpacing (v) {
        const t = this.paraTargets()
        if (!t) return
        for (const p of t.ps) { p.dataset.ls = String(v); p.dataset.lsSet = '1'; p.style.lineHeight = (1.2 * v / 100).toFixed(3) }
        this.paraDone(t.n)
      }

      // ------------------------------------------------ rotate, align, group

      rotateMenu (e) {
        const turn = (d, abs) => {
          for (const k of this.multi.filter((x) => !String(x).startsWith('g') && this.free(x))) {
            const b = this.itemBox(k)
            const rot = abs !== undefined ? abs : (((b.rot || 0) + d) % 360 + 360) % 360
            this.setItemBox(k, { ...b, rot })
            this.commitItem(k, { ...b, rot })
          }
          this.placeBox()
        }
        this.menuAt(e, [
          { label: '向右旋转 90°', fn: () => turn(90) },
          { label: '向左旋转 90°', fn: () => turn(-90) },
          { label: '旋转 180°', fn: () => turn(180) },
          { label: '重置旋转', fn: () => turn(0, 0) },
          '-',
          { label: '指定角度…', fn: async () => { const v = await dialog({ title: '旋转角度', fields: [{ key: 'a', label: '角度（°）', type: 'number', value: String(this.itemBox(this.selected)?.rot ?? 0) }], check: (x) => Number.isFinite(+x.a) && x.a !== '' ? '' : '请输入数字' }); if (v) turn(0, ((+v.a % 360) + 360) % 360) } }
        ])
      }

      alignMenu (e) {
        const many = this.multi.length > 1
        const to = many ? '所选对象' : '幻灯片'
        this.menuAt(e, [
          { label: `左对齐（${to}）`, fn: () => this.arrange('l') },
          { label: '水平居中', fn: () => this.arrange('c') },
          { label: '右对齐', fn: () => this.arrange('r') },
          '-',
          { label: '顶端对齐', fn: () => this.arrange('t') },
          { label: '垂直居中', fn: () => this.arrange('m') },
          { label: '底端对齐', fn: () => this.arrange('b') },
          '-',
          { label: '横向分布', disabled: this.multi.length < 3, fn: () => this.arrange('dh') },
          { label: '纵向分布', disabled: this.multi.length < 3, fn: () => this.arrange('dv') }
        ])
      }

      group () {
        const keys = this.multi.filter((k) => this.free(k))
        if (keys.length < 2) { this.frame.toast('按住 Shift 选中两个以上的形状再组合。', 'info'); return }
        if (keys.some((k) => !String(k).startsWith('g') && this.shapeEl(k)?.dataset.ph)) { this.frame.toast('占位符不能组合。', 'warn'); return }
        this.reselect = keys.flatMap((k) => this.itemEls(k).map((n) => n.dataset.id)).filter(Boolean)
        this.push({ op: 'group', ...this.at(), items: keys.map((k) => String(k).startsWith('g') ? { group: +k.slice(1) } : { shape: +k }) })
      }

      ungroup () {
        const key = this.selected
        if (!String(key ?? '').startsWith('g')) return
        this.reselect = this.itemEls(key).map((n) => n.dataset.id).filter(Boolean)
        this.push({ op: 'ungroup', ...this.at(), group: +key.slice(1) })
      }

      // ------------------------------------------------ crop

      startCrop () {
        const n = this.selected && this.shapeEl(this.selected)
        if (!n || n.dataset.t !== 'pic' || !this.doc.writable || this.crop) return
        if (+(n.dataset.rot || 0)) { this.frame.toast('旋转过的图片请先重置旋转再裁剪。', 'info'); return }
        const [l, t, r, b] = (n.dataset.crop || '0,0,0,0').split(',').map(Number)
        const box = this.boxOf(n)
        const fw = box.cx / Math.max(0.01, 1 - l - r); const fh = box.cy / Math.max(0.01, 1 - t - b)
        this.crop = { n, full: { x: box.x - l * fw, y: box.y - t * fh, w: fw, h: fh }, box: { ...box }, start: { ...box } }
        n.classList.add('qp-cropping')
        this.placeBox(); this.refreshBar()
        this.frame.toast('拖动黑色控点裁剪，按 Enter 完成，Esc 取消。', 'info')
      }

      cropDrag (e, h) {
        e.preventDefault()
        const c = this.crop
        const k = this.model.emu * this.scale
        const ox = e.clientX; const oy = e.clientY
        const s = { ...c.box }; const F = c.full
        const min = 20000
        const move = (ev) => {
          const dx = (ev.clientX - ox) / k; const dy = (ev.clientY - oy) / k
          const b = { ...s }
          if (h.includes('w')) { const x = Math.max(F.x, Math.min(s.x + s.cx - min, s.x + dx)); b.cx = s.x + s.cx - x; b.x = x }
          if (h.includes('e')) b.cx = Math.max(min, Math.min(F.x + F.w - s.x, s.cx + dx))
          if (h.includes('n')) { const y = Math.max(F.y, Math.min(s.y + s.cy - min, s.y + dy)); b.cy = s.y + s.cy - y; b.y = y }
          if (h.includes('s')) b.cy = Math.max(min, Math.min(F.y + F.h - s.y, s.cy + dy))
          c.box = b
          this.paintCrop()
        }
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', () => window.removeEventListener('pointermove', move), { once: true })
      }

      paintCrop () {
        const c = this.crop
        const F = c.full; const b = c.box
        this.moveTo(c.n, b, true)
        const img = c.n.querySelector('img')
        if (img) Object.assign(img.style, { left: (F.x - b.x) / b.cx * 100 + '%', top: (F.y - b.y) / b.cy * 100 + '%', width: F.w / b.cx * 100 + '%', height: F.h / b.cy * 100 + '%' })
        this.placeBox()
      }

      endCrop (save) {
        const c = this.crop
        if (!c) return
        const F = c.full
        if (!save) { c.box = { ...c.start }; this.paintCrop() }
        const b = c.box
        const v = { l: (b.x - F.x) / F.w, t: (b.y - F.y) / F.h, r: (F.x + F.w - b.x - b.cx) / F.w, b: (F.y + F.h - b.y - b.cy) / F.h }
        this.crop = null
        c.n.classList.remove('qp-cropping')
        const changed = ['x', 'y', 'cx', 'cy'].some((k) => Math.abs(b[k] - c.start[k]) > 1)
        if (save && changed) {
          c.n.dataset.crop = [v.l, v.t, v.r, v.b].map((x) => +x.toFixed(5)).join(',')
          this.push({ op: 'crop', ...this.at(), shape: +c.n.dataset.k, ...v, x: Math.round(b.x), y: Math.round(b.y), cx: Math.round(b.cx), cy: Math.round(b.cy) }, { quiet: true })
        }
        this.placeBox(); this.refreshBar()
      }

      cropMenu (e) {
        const n = this.selected && this.shapeEl(this.selected)
        if (!n || n.dataset.t !== 'pic') return
        const apply = (ratio) => {
          const [l, t, r, b] = (n.dataset.crop || '0,0,0,0').split(',').map(Number)
          const box = this.boxOf(n)
          const fw = box.cx / Math.max(0.01, 1 - l - r); const fh = box.cy / Math.max(0.01, 1 - t - b)
          const F = { x: box.x - l * fw, y: box.y - t * fh, w: fw, h: fh }
          let w = F.w; let h = F.h
          if (ratio) { if (w / h > ratio) w = h * ratio; else h = w / ratio }
          const nb = { x: F.x + (F.w - w) / 2, y: F.y + (F.h - h) / 2, cx: w, cy: h }
          this.renderAfter = true
          this.selectAfter = n.dataset.k
          this.push({ op: 'crop', ...this.at(), shape: +n.dataset.k, l: (nb.x - F.x) / F.w, t: (nb.y - F.y) / F.h, r: (F.x + F.w - nb.x - nb.cx) / F.w, b: (F.y + F.h - nb.y - nb.cy) / F.h, x: Math.round(nb.x), y: Math.round(nb.y), cx: Math.round(nb.cx), cy: Math.round(nb.cy) })
        }
        this.menuAt(e, [['1:1 正方形', 1], ['4:3', 4 / 3], ['3:4', 3 / 4], ['16:9', 16 / 9], ['9:16', 9 / 16], ['3:2', 3 / 2]].map(([l, r]) => ({ label: l, fn: () => apply(r) })).concat(['-', { label: '还原整张图片', fn: () => apply(0) }]))
      }

      // ------------------------------------------------ design

      changeLayout (part) {
        this.ui.layoutSel.value = this.slide()?.layoutPart ?? ''
        if (!this.doc.writable || this.mode === 'master' || part === this.slide()?.layoutPart) return
        this.push({ op: 'slide_layout', slide: this.slide().n, layout: part })
      }

      themeMenu (e, kind) {
        if (!this.doc.writable) return
        const th = this.model.theme ?? {}
        if (kind === 'colors') {
          this.menuAt(e, [
            ...THEME_COLORS.map(([name, c]) => ({ label: (th.name === name ? '✓ ' : '') + name, pre: swatch(['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((k) => c[k])), fn: () => this.push({ op: 'theme', name, colors: c }) })),
            '-',
            { label: '自定义颜色…', fn: () => this.customColors() }
          ])
        } else {
          this.menuAt(e, THEME_FONTS.map(([name, f]) => ({ label: (th.minorEa === f.minor.ea && th.majorEa === f.major.ea ? '✓ ' : '') + name, pre: el('b', { class: 'qp-fontpre', style: { fontFamily: `'${f.major.latin}','${f.major.ea}'` } }, 'Aa 文'), fn: () => this.push({ op: 'theme', fonts: f }) })))
        }
      }

      async customColors () {
        const c = this.model.theme?.colors ?? {}
        const names = [['dk2', '深色文字'], ['lt2', '浅色背景'], ['accent1', '着色 1'], ['accent2', '着色 2'], ['accent3', '着色 3'], ['accent4', '着色 4'], ['accent5', '着色 5'], ['accent6', '着色 6'], ['hlink', '超链接']]
        const v = await dialog({ title: '自定义主题颜色', fields: names.map(([k, l]) => ({ key: k, label: l, type: 'color', value: '#' + (c[k] ?? '888888') })) })
        if (!v) return
        this.push({ op: 'theme', name: '自定义', colors: Object.fromEntries(names.map(([k]) => [k, v[k].slice(1).toUpperCase()])) })
      }

      async bgDialog () {
        if (!this.doc.writable) return
        let src = null
        const pick = el('button', { type: 'button', class: 'q-dlg-x' }, '选择图片…')
        const label = el('small', { class: 'qp-bgname' }, '未选择')
        const row = el('div', { class: 'q-dlg-row qp-bgpic' }, el('span', {}, '图片'), el('div', {}, pick, ' ', label))
        pick.addEventListener('click', async () => { const s = await pickImage(); if (s) { src = s; label.textContent = '已选择图片' } })
        // The picture row belongs to the image fill only.
        setTimeout(() => { const kind = row.closest('form')?.querySelector('[name=kind]'); if (!kind) return; const upd = () => { row.hidden = kind.value !== 'image' }; kind.addEventListener('change', upd); upd() })
        const v = await dialog({
          title: this.mode === 'master' ? '母版背景' : '设置背景格式',
          fields: [
            { key: 'kind', label: '填充', type: 'select', value: 'solid', options: [['solid', '纯色'], ['gradient', '渐变'], ['image', '图片'], ['reset', '恢复默认']] },
            { key: 'c1', label: '颜色', type: 'color', value: '#FFFFFF', show: (x) => x.kind === 'solid' || x.kind === 'gradient' },
            { key: 'c2', label: '结束颜色', type: 'color', value: '#DDEBF7', show: (x) => x.kind === 'gradient' },
            { key: 'ang', label: '方向', type: 'select', value: '90', options: [['90', '从上到下'], ['0', '从左到右'], ['45', '左上到右下'], ['135', '右上到左下']], show: (x) => x.kind === 'gradient' },
            { key: 'all', label: '应用到全部', type: 'checkbox', value: false, show: () => this.mode !== 'master' }
          ],
          body: row,
          check: (x) => x.kind === 'image' && !src ? '请先选择图片' : ''
        })
        if (!v) return
        const op = { op: 'background', ...(v.all && this.mode !== 'master' ? { all: true } : this.at()) }
        if (v.kind === 'solid') op.fill = v.c1.slice(1).toUpperCase()
        else if (v.kind === 'gradient') op.gradient = [v.c1.slice(1).toUpperCase(), v.c2.slice(1).toUpperCase(), +v.ang]
        else if (v.kind === 'image') op.image = src
        else op.reset = true
        this.push(op)
      }

      setMode (mode) {
        if (mode === this.mode || !this.model) return
        if (mode === 'master' && !this.model.masters?.length) { this.frame.toast('这个演示文稿没有母版。', 'warn'); return }
        if (this.editing) this.endEdit(true)
        if (this.crop) this.endCrop(true)
        this.saveNotes()
        this.mode = mode
        this.selected = null; this.multi = []
        if (mode === 'master') {
          const lp = this.slide()?.layoutPart
          const i = this.model.masters.findIndex((x) => x.part === lp)
          this.mix = i >= 0 ? i : 0
        }
        this.renderThumbs()
        this.show(this.pos())
      }

      // ------------------------------------------------ motion

      setTransition (patch, all) {
        if (this.mode === 'master' || !this.doc.writable) return
        const sl = this.slide()
        const cur = { type: 'none', spd: 'med', ...(sl.transition ?? {}) }
        if (!TRANS_FX.some(([v]) => v === cur.type) && !patch.type) { this.frame.toast('当前切换效果由 PowerPoint 设置，请先选一种效果。', 'info'); this.refreshMotion(); return }
        const t = { ...cur, ...patch }
        if (patch.type && patch.type !== cur.type) { delete t.dir; delete t.orient }
        if (t.type === 'split' && patch.dir) { const [o, d] = patch.dir.split('-'); t.orient = o; t.dir = d }
        if (t.type !== 'split') delete t.orient
        this.push({ op: 'transition', ...(all ? { all: true } : { slide: sl.n }), type: t.type, spd: t.spd, ...(t.dir ? { dir: t.dir } : {}), ...(t.orient ? { orient: t.orient } : {}), adv: t.adv ?? null })
        if (all) this.frame.toast('已应用到所有幻灯片。', 'info')
        if (patch.type && patch.type !== 'none') this.previewNext = true
      }

      previewTransition () {
        const sl = this.slide()
        const t = sl?.transition
        if (!t || t.type === 'none' || this.mode === 'master') return
        const m = this.model
        const prev = this.model.slides[this.ix - 1]
        const wrap = el('div', { class: 'qp-tprev', style: { width: m.width + 'px', height: m.height + 'px', transform: `scale(${this.scale})` } })
        const old = el('div', { class: 'slide', html: prev?.html ?? '', style: { width: m.width + 'px', height: m.height + 'px', background: prev ? bgCss(prev.bg) : '#000' } })
        const cur = el('div', { class: 'slide', html: sl.html, style: { width: m.width + 'px', height: m.height + 'px', background: bgCss(sl.bg) } })
        wrap.append(old, cur)
        this.scaler.append(wrap)
        transitionIn(cur, old, t).then(() => setTimeout(() => wrap.remove(), 250))
      }

      setAnim (patch) {
        if (this.mode === 'master' || !this.doc.writable) return
        const sl = this.slide()
        if (sl.animLocked) { this.frame.toast('这张幻灯片含有本编辑器不支持的动画（触发器或曲线路径等），请在 PowerPoint 中修改。', 'warn'); this.refreshMotion(); return }
        const list = (sl.anims ?? []).map((a) => ({ ...a }))
        const spids = this.multi.map((k) => this.spidOf(k)).filter(Boolean)
        if (!spids.length && this.animSel === undefined) { this.frame.toast('先选中要添加动画的形状。', 'info'); this.refreshMotion(); return }
        const picked = this.animSel !== undefined && list[this.animSel] ? [this.animSel] : list.map((a, i) => (spids.includes(a.spid) ? i : -1)).filter((i) => i >= 0)
        let next = list
        if (patch.effect === 'none') { next = list.filter((a, i) => !picked.includes(i)); this.animSel = undefined } else if (patch.effect && !picked.length) {
          for (const spid of spids) next.push({ spid, cls: patch.cls, effect: patch.effect, start: 'click', dur: patch.cls === 'emph' && patch.effect !== 'pulse' ? 1000 : patch.cls === 'path' ? 2000 : 500, dir: 'b' })
        } else if (!picked.length) { this.frame.toast('先为所选形状选择一种动画效果。', 'info'); this.refreshMotion(); return } else for (const i of picked) Object.assign(next[i], patch)
        this.pushAnims(next)
      }

      pushAnims (list) {
        this.keepMulti = [...this.multi]
        this.push({ op: 'animations', slide: this.slide().n, list: list.map(({ spid, cls, effect, start, dur, dir }) => ({ spid, ...(cls && cls !== 'entr' ? { cls } : {}), effect, start, dur, ...(dir ? { dir } : {}) })) })
      }

      moveAnim (d) {
        const list = (this.slide()?.anims ?? []).map((a) => ({ ...a }))
        const a = this.curAnim()
        const i = list.findIndex((x) => x.spid === a?.spid && animKey(x) === (a && animKey(a)) && x.start === a?.start && x.dur === a?.dur)
        const j = i + d
        if (i < 0 || j < 0 || j >= list.length) return
        ;[list[i], list[j]] = [list[j], list[i]]
        this.animSel = j
        this.pushAnims(list)
      }

      // paintBadges numbers the animated shapes while the 动画 tab is open.
      paintBadges () {
        if (!this.badges) return
        this.badges.replaceChildren()
        if (this.tabName !== '动画' || this.mode === 'master' || !this.model) return
        const sl = this.slide()
        if (!sl?.anims?.length) return
        const label = new Map()
        let c = 0
        for (const g of animGroups(sl.anims)) { const l = g.auto ? '0' : String(++c); for (const st of g.steps) for (const a of st.items) label.set(a, l) }
        const per = new Map()
        for (const a of sl.anims) { const arr = per.get(a.spid) ?? []; if (!arr.includes(label.get(a))) arr.push(label.get(a)); per.set(a.spid, arr) }
        const k = this.model.emu * this.scale
        for (const [spid, labels] of per) {
          const n = animTargets(this.slideEl, spid)[0]
          if (!n) continue
          const b = n.dataset.gid === spid ? this.itemBox('g' + n.dataset.g) : this.boxOf(n)
          this.badges.append(el('span', { class: 'qp-badge', style: { left: Math.max(0, b.x * k - 18) + 'px', top: b.y * k + 'px' } }, labels.join(',')))
        }
      }

      paintPane () {
        if (!this.apane || this.apane.hidden || !this.model) return
        const sl = this.slide()
        const close = el('button', { type: 'button', title: '关闭' }, '✕')
        close.addEventListener('click', () => { this.apane.hidden = true; this.paneBtn.classList.remove('on') })
        const playAll = el('button', { type: 'button' }, '▶ 播放')
        playAll.addEventListener('click', () => this.previewAnims())
        const rows = []
        if (this.mode === 'master') rows.push(el('p', { class: 'qp-anote' }, '母版视图没有动画。'))
        else if (sl?.animLocked) rows.push(el('p', { class: 'qp-anote' }, '这张幻灯片含有本编辑器不支持的动画，只能查看。'))
        else if (!sl?.anims?.length) rows.push(el('p', { class: 'qp-anote' }, '选中形状后，在“动画效果”里选一种效果。'))
        let c = 0
        const lab = new Map()
        for (const g of animGroups(sl.anims ?? [])) { const l = g.auto ? '0' : String(++c); for (const st of g.steps) for (const a of st.items) lab.set(a, l) }
        ;(this.mode === 'master' ? [] : sl.anims ?? []).forEach((a, i) => {
          const n = animTargets(this.slideEl, a.spid)[0]
          const name = n ? (n.dataset.t === 'pic' ? '图片' : (n.innerText || '').replace(/\u200b/g, '').trim().slice(0, 14) || (n.dataset.gid === a.spid ? '组合' : '形状')) : `形状 ${a.spid}`
          const row = el('div', { class: 'qp-arow' + (i === this.animSel ? ' on' : '') },
            el('b', {}, a.start === 'click' ? lab.get(a) : a.start === 'with' ? '↔' : '↳'),
            el('span', { class: 'qp-afx', dataset: { cls: a.cls ?? 'entr' } }, ANIM_NAME[animKey(a)] ?? a.effect),
            el('span', { class: 'qp-aname' }, name))
          const del = el('button', { type: 'button', title: '删除此动画' }, '✕')
          del.addEventListener('click', (ev) => { ev.stopPropagation(); if (sl.animLocked || !this.doc.writable) return; this.animSel = undefined; this.pushAnims((sl.anims ?? []).filter((x, j) => j !== i)) })
          row.append(del)
          row.addEventListener('click', () => {
            this.animSel = i
            const key = n ? this.keyOf(n) : null
            this.selected = key; this.multi = key ? [key] : []
            this.placeBox(); this.refreshBar(); this.paintPane()
          })
          rows.push(row)
        })
        this.apane.replaceChildren(el('div', { class: 'qp-ahead' }, el('b', {}, '动画窗格'), el('span', { class: 'q-space' }), playAll, close), ...rows)
      }

      // previewAnims plays the slide's effects on stage, one click after
      // another.
      previewAnims () {
        const sl = this.slide()
        if (this.mode === 'master' || !sl?.anims?.length) return
        if (this.editing) this.endEdit(true)
        const root = this.slideEl
        const touched = new Set()
        for (const a of sl.anims) for (const n of animTargets(root, a.spid)) { if (animStartsHidden(sl.anims, a.spid)) n.style.visibility = 'hidden'; touched.add(n) }
        let at = 300
        for (const g of animGroups(sl.anims)) {
          const start = at
          setTimeout(() => { for (const st of g.steps) for (const a of st.items) for (const n of animTargets(root, a.spid)) runEffect(n, a, st.at) }, start)
          at = start + g.end + 400
        }
        // Back to the slide as edited: shown, held effects let go.
        setTimeout(() => { for (const n of touched) { for (const x of n.getAnimations()) x.cancel(); n.style.visibility = '' } }, at + 600)
      }

      arrange (v) {
        const keys = this.multi.filter((k) => this.free(k))
        if (!keys.length || !this.doc.writable) return
        if (['front', 'back', 'forward', 'backward'].includes(v)) {
          const key = this.selected
          if (!this.free(key)) return
          if (String(key).startsWith('g')) { this.push({ op: 'order', ...this.at(), group: +key.slice(1), to: v }); this.reselect = this.itemEls(key).map((n) => n.dataset.id); return }
          const k = +key
          const all = [...this.slideEl.querySelectorAll('[data-k]')].map((x) => +x.dataset.k)
          const top = Math.max(...all)
          this.reselect = [this.shapeEl(key)?.dataset.id].filter(Boolean)
          if (!this.reselect.length) this.selectAfter = String({ front: top, back: 1, forward: Math.min(top, k + 1), backward: Math.max(1, k - 1) }[v])
          this.push({ op: 'order', ...this.at(), shape: k, to: v })
          return
        }
        const W = this.model.width / this.model.emu; const H = this.model.height / this.model.emu
        const boxes = new Map(keys.map((k) => [k, this.itemBox(k)]))
        const ref = keys.length > 1 ? unionBox([...boxes.values()]) : { x: 0, y: 0, cx: W, cy: H }
        if (v === 'dh' || v === 'dv') {
          const h = v === 'dh'
          const sorted = [...boxes].sort((a, b) => h ? a[1].x - b[1].x : a[1].y - b[1].y)
          const total = sorted.reduce((s, [, b]) => s + (h ? b.cx : b.cy), 0)
          const gap = ((h ? ref.cx : ref.cy) - total) / Math.max(1, sorted.length - 1)
          let pos = h ? ref.x : ref.y
          for (const [, b] of sorted) { if (h) b.x = pos; else b.y = pos; pos += (h ? b.cx : b.cy) + gap }
        } else {
          for (const b of boxes.values()) {
            if (v === 'l') b.x = ref.x; else if (v === 'c') b.x = ref.x + (ref.cx - b.cx) / 2; else if (v === 'r') b.x = ref.x + ref.cx - b.cx
            else if (v === 't') b.y = ref.y; else if (v === 'm') b.y = ref.y + (ref.cy - b.cy) / 2; else if (v === 'b') b.y = ref.y + ref.cy - b.cy
          }
        }
        for (const [k, b] of boxes) { this.setItemBox(k, b); this.commitItem(k, b) }
        this.placeBox()
      }

      // ------------------------------------------------ shape clipboard

      copyShape (cut) {
        const key = this.selected
        if (!this.free(key)) return
        this.slideEl.querySelector('.qp-cut')?.classList.remove('qp-cut')
        const els = this.itemEls(key)
        const text = els.map((n) => (n.innerText ?? '').trim()).filter(Boolean).join('\n')
        const g = String(key).startsWith('g')
        this.clip = { from: this.at(), ...(g ? { group: +key.slice(1) } : { shape: +key }), box: this.itemBox(key), cut, text, times: 0 }
        if (cut) for (const n of els) n.classList.add('qp-cut')
        navigator.clipboard?.writeText(text).catch(() => {})
        this.frame.toast(cut ? '已剪切形状，到目标位置按 ⌘V 粘贴。' : '已复制形状。', 'info')
      }

      pasteShape () {
        const c = this.clip
        if (!c || !this.doc.writable) return
        const here = this.at()
        const same = c.from.slide === here.slide && c.from.part === here.part
        c.times++
        const off = same && !c.cut ? Math.round(this.model.width / this.model.emu * 0.02) * c.times : 0
        const box = { x: c.box.x, y: c.box.y, cx: c.box.cx, cy: c.box.cy }
        this.selectLast = 'select'
        this.push({ op: 'copy_shape', from: c.from.slide, fromPart: c.from.part, ...(c.group ? { group: c.group } : { shape: c.shape }), ...here, box, dx: off, dy: off })
        if (c.cut) { this.push({ op: 'delete_shape', ...(c.from.part ? { part: c.from.part } : { slide: c.from.slide }), ...(c.group ? { group: c.group } : { shape: c.shape }) }); this.clip = undefined }
      }

      duplicateShape () {
        const key = this.selected
        if (!this.free(key)) return
        const off = Math.round(this.model.width / this.model.emu * 0.02)
        const b = this.itemBox(key)
        const here = this.at()
        this.selectLast = 'select'
        this.push({ op: 'copy_shape', from: here.slide, fromPart: here.part, ...(String(key).startsWith('g') ? { group: +key.slice(1) } : { shape: +key }), ...here, box: { x: b.x, y: b.y, cx: b.cx, cy: b.cy }, dx: off, dy: off })
      }

      paste (e) {
        this.pasteSeen = true
        if (this.editing || this.cellEdit || !this.doc.writable) return
        e.preventDefault()
        const img = imageFile(e.clipboardData)
        if (img) { readImage(img).then((src) => this.addPicture(src), (err) => this.frame.toast(err.message, 'warn')); return }
        const text = (e.clipboardData.getData('text/plain') || '').replace(/\r\n?/g, '\n')
        if (this.clip && text.trim() === this.clip.text) { this.pasteShape(); return }
        if (text.trim()) { this.selectLast = 'select'; this.push({ op: 'add_textbox', ...this.at(), sz: 1800, text: text.trim() }) } else if (this.clip) this.pasteShape()
      }

      // ------------------------------------------------ table cells

      // tableItems are the row and column entries of a table's menu, for
      // the cell under the pointer.
      tableItems (n, e, ro) {
        const td = n.dataset.t === 'tbl' && e.target.closest?.('td[data-r]')
        if (!td || !n.contains(td)) return []
        const r = +td.dataset.r; const c = +td.dataset.c
        const trs = [...n.querySelectorAll('tr')]
        const cols = Math.max(...trs.map((tr) => [...tr.children].reduce((a, x) => a + (x.colSpan || 1), 0)))
        const go = (axis, at, count) => () => { this.renderAfter = true; this.selectAfter = n.dataset.k; this.push({ op: 'table_lines', ...this.at(), shape: +n.dataset.k, axis, at, count }) }
        return ['-',
          { label: '在上方插入行', disabled: ro, fn: go('r', r, 1) },
          { label: '在下方插入行', disabled: ro, fn: go('r', r + td.rowSpan, 1) },
          { label: '在左侧插入列', disabled: ro, fn: go('c', c, 1) },
          { label: '在右侧插入列', disabled: ro, fn: go('c', c + td.colSpan, 1) },
          { label: '删除行', disabled: ro || trs.length < 2, fn: go('r', r, -1) },
          { label: '删除列', disabled: ro || cols < 2, fn: go('c', c, -1) }]
      }

      editCell (n, td) {
        if (!this.doc.writable) return
        this.select(n.dataset.k)
        const before = td.innerText
        td.contentEditable = 'plaintext-only'
        td.classList.add('qp-cell-edit')
        this.cellEdit = td
        td.focus()
        const r = document.createRange(); r.selectNodeContents(td)
        const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r)
        const done = (save) => {
          td.removeEventListener('blur', onBlur)
          td.removeEventListener('keydown', onKey)
          td.contentEditable = 'false'
          td.classList.remove('qp-cell-edit')
          this.cellEdit = null
          const text = td.innerText.replace(/\n$/, '')
          if (!save) td.innerText = before
          else if (text !== before.replace(/\n$/, '') && n.isConnected) this.push({ op: 'table_cell', ...this.at(), shape: +n.dataset.k, row: +td.dataset.r, col: +td.dataset.c, text }, { quiet: true })
        }
        const onBlur = () => done(true)
        const onKey = (e) => {
          e.stopPropagation()
          if (e.key === 'Escape') { e.preventDefault(); done(false); this.stage.focus() } else if (e.key === 'Tab') {
            e.preventDefault()
            const cells = [...n.querySelectorAll('td[data-r]')]
            const next = cells[cells.indexOf(td) + (e.shiftKey ? -1 : 1)]
            td.blur()
            if (next) this.editCell(n, next)
          }
        }
        td.addEventListener('blur', onBlur)
        td.addEventListener('keydown', onKey)
      }

      // ------------------------------------------------ slideshow

      play (from, { presenter } = {}) {
        if (this.mode === 'master') this.setMode('slides')
        if (this.editing) this.endEdit(true)
        const m = this.model
        const order = m.slides.map((s, i) => i).filter((i) => !m.slides[i].hidden || i === from)
        if (!order.length) return
        let at = Math.max(0, order.indexOf(from))
        let step = 0
        let groups = []
        let cur = null
        let advTimer
        let aud = null
        const frame = el('div', { class: 'qp-frame' })
        const end = el('div', { class: 'qp-show-end', hidden: true }, '放映结束，单击或按 Esc 退出')
        const screen = el('div', { class: 'qp-screen' }, frame, end)
        const ov = el('div', { class: 'qp qp-show' + (presenter ? ' qp-presenter' : ''), tabindex: '0' }, screen)
        // The presenter's side: time, counter, the next slide and notes.
        const pv = {}
        if (presenter) {
          pv.time = el('b', { class: 'qp-pv-time' }, '00:00')
          pv.clock = el('span', { class: 'qp-pv-clock' })
          pv.count = el('span', { class: 'qp-pv-count' })
          pv.next = el('div', { class: 'qp-pv-next' })
          pv.notes = el('div', { class: 'qp-pv-notes' })
          const btn = (label, fn, title) => { const b = el('button', { type: 'button', title: title ?? label }, label); b.addEventListener('click', (ev) => { ev.stopPropagation(); fn() }); return b }
          let paused = false; let elapsed = 0
          pv.pause = btn('暂停', () => { paused = !paused; pv.pause.textContent = paused ? '继续' : '暂停' })
          pv.tick = setInterval(() => {
            if (!paused) elapsed++
            pv.time.textContent = `${String(Math.floor(elapsed / 60)).padStart(2, '0')}:${String(elapsed % 60).padStart(2, '0')}`
            pv.clock.textContent = new Date().toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' })
          }, 1000)
          const side = el('div', { class: 'qp-pv' },
            el('div', { class: 'qp-pv-top' }, pv.time, pv.pause, btn('重置', () => { elapsed = 0; pv.time.textContent = '00:00' }), el('span', { class: 'q-space' }), pv.clock),
            pv.count,
            el('div', { class: 'qp-pv-h' }, '下一张'), pv.next,
            el('div', { class: 'qp-pv-h' }, '备注'), pv.notes,
            el('div', { class: 'qp-pv-btns' }, btn('◀ 上一步', () => prev()), btn('下一步 ▶', () => next()), btn('观众窗口', () => openAud(), '在另一个窗口显示幻灯片，可拖到投影屏幕上全屏'), btn('结束放映', () => close())))
          side.addEventListener('click', (ev) => ev.stopPropagation())
          ov.append(side)
        }
        document.body.append(ov)
        const fitIn = (box, f) => {
          const r = box.getBoundingClientRect()
          const s = Math.min(r.width / m.width, r.height / m.height)
          Object.assign(f.style, { width: m.width + 'px', height: m.height + 'px', transform: `translate(${(r.width - m.width * s) / 2}px, ${(r.height - m.height * s) / 2}px) scale(${s})` })
        }
        const fit = () => { fitIn(screen, frame); fitAud() }
        const make = (sl, doc = document) => {
          const n = doc.createElement('div')
          n.className = 'slide'
          n.innerHTML = sl.html
          Object.assign(n.style, { width: m.width + 'px', height: m.height + 'px', background: bgCss(sl.bg) })
          for (const x of n.querySelectorAll('[data-fs]')) x.style.setProperty('--qfs', x.dataset.fs + 'px')
          // Empty placeholders do not show; an empty text box shows blank.
          for (const x of n.querySelectorAll('[data-prompt]:not([data-filled])')) { if (x.dataset.ph) x.remove(); else x.removeAttribute('data-prompt') }
          return n
        }
        // Without a second window the presenter can show the audience view
        // here, full screen, and switch back.
        const solo = (on) => { ov.classList.toggle('qp-solo', on); requestAnimationFrame(fit); ov.focus() }
        if (presenter) {
          const back = el('button', { type: 'button', class: 'qp-solo-back', title: '回到带备注和计时的演示者视图' }, '演示者视图')
          back.addEventListener('click', (ev) => { ev.stopPropagation(); solo(false) })
          ov.append(back)
        }
        const audFrame = () => aud && !aud.closed ? aud.document.querySelector('.qp-frame') : null
        const fitAud = () => { const f = audFrame(); if (f) fitIn(aud.document.body, f) }
        const syncAud = () => { const f = audFrame(); if (!f) return; f.replaceChildren(); if (cur) f.append(aud.document.importNode(cur, true)); fitAud() }
        const openAud = () => {
          if (aud && !aud.closed) { aud.focus(); return }
          aud = window.open('', 'quire-audience', 'width=1280,height=760')
          if (!aud) { solo(true); this.frame.toast('浏览器拦截了弹出窗口：已改在本窗口全屏显示幻灯片（观众看到的画面），点右下角「演示者视图」切回。', 'warn'); return }
          const css = document.getElementById('quire-pptx-css')?.textContent ?? ''
          aud.document.open()
          aud.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>放映</title><style>html,body{margin:0;width:100%;height:100%;background:#000;overflow:hidden}.qp{position:fixed;inset:0}.qp-frame{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden}.qp-frame>.slide{position:absolute;left:0;top:0;box-shadow:none}${css}</style></head><body><div class="qp"><div class="qp-frame"></div></div></body></html>`)
          aud.document.close()
          aud.addEventListener('resize', fitAud)
          aud.addEventListener('keydown', (ev) => onKey(ev))
          aud.document.addEventListener('click', () => next())
          syncAud()
        }
        const paintPv = () => {
          if (!presenter) return
          const sl = m.slides[order[at]]
          pv.count.textContent = at < order.length ? `第 ${at + 1} / ${order.length} 张${groups.length ? ` · 动画 ${Math.min(step, groups.length)} / ${groups.length}` : ''}` : '放映结束'
          pv.notes.textContent = sl?.notes?.trim() || '（无备注）'
          pv.next.replaceChildren()
          const nx = m.slides[order[at + 1]]
          if (at < order.length && nx) { const f = el('div', { class: 'qp-frame' }); f.append(make(nx)); pv.next.append(f); requestAnimationFrame(() => fitIn(pv.next, f)) } else pv.next.append(el('span', {}, at + 1 >= order.length ? '放映结束' : ''))
        }
        const scheduleAdv = () => {
          clearTimeout(advTimer)
          const adv = m.slides[order[at]]?.transition?.adv
          if (adv !== undefined && step >= groups.length) advTimer = setTimeout(() => next(), adv * 1000)
        }
        const playStep = () => {
          const g = groups[step++]
          const roots = [cur, audFrame()?.firstElementChild].filter(Boolean)
          for (const st of g.steps) for (const a of st.items) for (const r of roots) for (const n of animTargets(r, a.spid)) runEffect(n, a, st.at)
          paintPv()
          clearTimeout(advTimer)
          advTimer = setTimeout(scheduleAdv, g.end)
        }
        const enter = (i, { built, animate } = {}) => {
          clearTimeout(advTimer)
          end.hidden = i < order.length
          if (i >= order.length) { for (const x of [...frame.children]) x.remove(); cur = null; groups = []; syncAud(); paintPv(); return }
          const sl = m.slides[order[i]]
          const n = make(sl)
          groups = animGroups(sl.anims)
          step = built ? groups.length : 0
          for (const a of sl.anims ?? []) if (built ? animEndsHidden(sl.anims, a.spid) : animStartsHidden(sl.anims, a.spid)) for (const x of animTargets(n, a.spid)) x.style.visibility = 'hidden'
          const old = cur
          const t = sl.transition
          frame.append(n)
          cur = n
          for (const x of [...frame.children]) if (x !== n && x !== old) x.remove()
          if (animate && old && t && TRANS_FX.some(([v]) => v === t.type) && t.type !== 'none') transitionIn(n, old, t).then(() => old.remove())
          else old?.remove()
          syncAud()
          paintPv()
          if (!built && groups[0]?.auto) playStep(); else scheduleAdv()
        }
        const next = () => {
          if (at >= order.length) { close(); return }
          if (step < groups.length) { playStep(); return }
          at++
          enter(at, { animate: true })
        }
        const prev = () => { if (at === 0 && step === 0) return; at = Math.max(0, Math.min(at, order.length) - (at >= order.length ? 1 : step > 0 && !groups[0]?.auto ? 0 : 1)); enter(at, { built: true }) }
        const close = () => {
          clearTimeout(advTimer)
          clearInterval(pv.tick)
          window.removeEventListener('resize', fit)
          document.removeEventListener('fullscreenchange', onFs)
          ov.remove()
          if (aud && !aud.closed) aud.close()
          if (document.fullscreenElement === ov) document.exitFullscreen().catch(() => {})
          const ix = order[Math.min(at, order.length - 1)]
          if (ix !== this.ix) this.show(ix)
          this.stage.focus({ preventScroll: true })
        }
        const onFs = () => { if (!document.fullscreenElement) close() }
        const onKey = (e) => {
          e.preventDefault(); e.stopPropagation()
          if (e.key === 'Escape') close()
          else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) next()
          else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key)) prev()
          else if (e.key === 'Home') { at = 0; enter(0) } else if (e.key === 'End') { at = order.length - 1; enter(at, { built: true }) }
        }
        ov.addEventListener('keydown', onKey)
        screen.addEventListener('click', () => next())
        ov.addEventListener('contextmenu', (e) => { e.preventDefault(); prev() })
        window.addEventListener('resize', fit)
        fit()
        enter(at)
        ov.focus()
        ov.requestFullscreen?.().then(() => document.addEventListener('fullscreenchange', onFs)).catch(() => {})
        this.showing = { next, prev, close, ov }
      }

      // ------------------------------------------------ context menus

      context (e) {
        e.preventDefault()
        if (this.editing) return
        const ro = !this.doc.writable
        const n = this.target(e)
        if (n) { const key = this.keyOf(n); if (!this.multi.includes(key) && !this.multi.includes(n.dataset.k)) this.select(key) } else this.select(null)
        this.stage.focus({ preventScroll: true })
        const key = this.selected
        const free = this.free(key)
        const master = this.mode === 'master'
        menu(e.clientX, e.clientY, n ? [
          { label: '标注问题…', kbd: '⇧⌘M', disabled: master, fn: () => this.frame.pinHere() },
          '-',
          { label: '剪切', kbd: '⌘X', disabled: ro || !free, fn: () => this.copyShape(true) },
          { label: '复制', kbd: '⌘C', disabled: !free, fn: () => this.copyShape(false) },
          { label: '粘贴', kbd: '⌘V', disabled: ro || !this.clip, fn: () => this.pasteShape() },
          { label: '创建副本', kbd: '⌘D', disabled: ro || !free, fn: () => this.duplicateShape() },
          { label: '删除', kbd: '⌫', disabled: ro, fn: () => this.deleteShape() },
          '-',
          n.dataset.t === 'txt' && this.selected === n.dataset.k ? { label: '编辑文字', kbd: '⏎', disabled: ro, fn: () => this.beginEdit(n) } : null,
          n.dataset.t === 'txt' && this.selected === n.dataset.k ? { label: '文本框选项…', disabled: ro, fn: () => this.bodyDialog() } : null,
          n.dataset.t === 'pic' ? { label: '裁剪', disabled: ro, fn: () => this.startCrop() } : null,
          ...this.tableItems(n, e, ro),
          this.multi.length > 1 ? { label: '组合', kbd: '⌘G', disabled: ro, fn: () => this.group() } : null,
          String(key).startsWith('g') ? { label: '取消组合', kbd: '⇧⌘G', disabled: ro, fn: () => this.ungroup() } : null,
          '-',
          { label: '置于顶层', disabled: ro || !free, fn: () => this.arrange('front') },
          { label: '上移一层', disabled: ro || !free, fn: () => this.arrange('forward') },
          { label: '下移一层', disabled: ro || !free, fn: () => this.arrange('backward') },
          { label: '置于底层', disabled: ro || !free, fn: () => this.arrange('back') },
          '-',
          { label: '水平居中', disabled: ro || !free, fn: () => this.arrange('c') },
          { label: '垂直居中', disabled: ro || !free, fn: () => this.arrange('m') },
          master ? null : '-',
          master ? null : { label: '添加动画：淡入', disabled: ro || !!this.slide().animLocked, fn: () => { this.animSel = undefined; this.setAnim({ effect: 'fade' }) } }
        ] : [
          { label: '标注这张幻灯片…', kbd: '⇧⌘M', disabled: master, fn: () => this.frame.pinHere() },
          '-',
          { label: '粘贴', kbd: '⌘V', disabled: ro || !this.clip, fn: () => this.pasteShape() },
          { label: '全选', kbd: '⌘A', fn: () => this.key({ key: 'a', metaKey: true, ctrlKey: true, preventDefault () {} }) },
          '-',
          { label: '插入文本框', disabled: ro, fn: () => this.ui.box.click() },
          { label: '插入图片…', disabled: ro, fn: () => this.pickPicture() },
          { label: '插入表格', disabled: ro, fn: () => this.ui.tbl.click() },
          '-',
          { label: '设置背景格式…', disabled: ro, fn: () => this.bgDialog() },
          master ? { label: '关闭母版视图', fn: () => this.setMode('slides') } : { label: '从当前幻灯片放映', kbd: '⇧F5', fn: () => this.play(this.ix) }
        ])
      }

      thumbMenu (e) {
        const t = e.target.closest?.('.qp-thumb')
        if (!t) return
        e.preventDefault()
        const i = +t.dataset.i
        if (i !== this.pos()) this.show(i)
        if (this.mode === 'master') { menu(e.clientX, e.clientY, [{ label: '设置背景格式…', disabled: !this.doc.writable, fn: () => this.bgDialog() }, { label: '关闭母版视图', fn: () => this.setMode('slides') }]); return }
        const ro = !this.doc.writable
        const sl = this.slide()
        const layout = this.model.layouts[0]
        menu(e.clientX, e.clientY, [
          { label: '新建幻灯片', disabled: ro || !layout, fn: () => this.newSlide(layout.part) },
          { label: '复制幻灯片', disabled: ro, fn: () => this.slideOp({ op: 'duplicate_slide' }, 1) },
          { label: '删除幻灯片', disabled: ro || this.model.slides.length < 2, fn: () => this.ui.del.click() },
          '-',
          { label: '上移', disabled: ro || i === 0, fn: () => this.ui.up.click() },
          { label: '下移', disabled: ro || i >= this.model.slides.length - 1, fn: () => this.ui.down.click() },
          { label: sl.hidden ? '取消隐藏' : '放映时隐藏', disabled: ro, fn: () => this.ui.hide.click() },
          '-',
          { label: '设置背景格式…', disabled: ro, fn: () => this.bgDialog() },
          '-',
          { label: '从这张开始放映', fn: () => this.play(i) },
          { label: '从这张开始（演示者视图）', fn: () => this.play(i, { presenter: true }) }
        ])
      }

      // ------------------------------------------------ find

      hosts (root) {
        const out = []
        for (const n of root.querySelectorAll('[data-k]')) {
          if (n.dataset.t === 'txt') out.push({ node: n, k: +n.dataset.k })
          else if (n.dataset.t === 'tbl') for (const td of n.querySelectorAll('td[data-r]')) out.push({ node: td, k: +n.dataset.k, r: +td.dataset.r, c: +td.dataset.c })
        }
        return out
      }

      texts (node) {
        const out = []
        const w = document.createTreeWalker(node, NodeFilter.SHOW_TEXT)
        for (let t = w.nextNode(); t; t = w.nextNode()) if (!t.parentElement.closest('.bu')) out.push(t)
        return out
      }

      slideDom (i) { return i === this.ix && this.mode === 'slides' ? this.slideEl : el('div', { html: this.model.slides[i].html }) }

      findAll (q) {
        const ql = q.toLowerCase()
        const out = []
        this.model.slides.forEach((sl, i) => {
          if (sl.missing) return
          const hs = this.hosts(this.slideDom(i))
          hs.forEach((h, j) => {
            const text = this.texts(h.node).map((t) => t.data).join('').toLowerCase()
            for (let x = text.indexOf(ql); x >= 0; x = text.indexOf(ql, x + ql.length)) out.push({ slide: i, host: j, start: x, end: x + ql.length })
          })
        })
        return out
      }

      rangeIn (node, m) {
        const r = document.createRange()
        let pos = 0; let started = false
        for (const t of this.texts(node)) {
          const end = pos + t.length
          if (!started && m.start < end) { r.setStart(t, m.start - pos); started = true }
          if (started && m.end <= end) { r.setEnd(t, m.end - pos); return r }
          pos = end
        }
        return undefined
      }

      showMatch (m) {
        if (this.mode === 'master') this.setMode('slides')
        if (this.editing) this.endEdit(true)
        if (m.slide !== this.ix) this.show(m.slide)
        const h = this.hosts(this.slideEl)[m.host]
        if (!h) return
        this.select(String(h.k))
        const r = this.rangeIn(h.node, m)
        if (r && window.CSS?.highlights && window.Highlight) window.CSS.highlights.set('q-find', new Highlight(r))
      }

      // replaceIn changes the text in a slide's DOM and returns the op that
      // saves it.
      replaceIn (dom, i, m, q, rep) {
        const h = this.hosts(dom)[m.host]
        const r = h && this.rangeIn(h.node, m)
        if (!r || r.toString().toLowerCase() !== q.toLowerCase()) return undefined
        const start = r.startContainer; const off = r.startOffset
        r.deleteContents()
        if (start.nodeType === 3 && start.isConnected !== false) start.insertData(off, rep)
        return h
      }

      opFor (h, i) {
        const n = this.model.slides[i].n
        return h.r ? { op: 'table_cell', slide: n, shape: h.k, row: h.r, col: h.c, text: h.node.innerText ?? h.node.textContent } : { op: 'shape_text', slide: n, shape: h.k, paras: this.paras(h.node) }
      }

      replaceMatch (m, q, rep) {
        if (!this.doc.writable) return
        if (this.mode === 'master') this.setMode('slides')
        if (m.slide !== this.ix) this.show(m.slide)
        const h = this.replaceIn(this.slideEl, m.slide, m, q, rep)
        if (h) this.push(this.opFor(h, m.slide), { quiet: true })
      }

      replaceAll (q, rep) {
        if (!this.doc.writable) return 0
        if (this.editing) this.endEdit(true)
        const all = this.findAll(q)
        const doms = new Map()
        const touched = new Map()
        let n = 0
        for (const m of all.reverse()) {
          if (!doms.has(m.slide)) doms.set(m.slide, this.slideDom(m.slide))
          const h = this.replaceIn(doms.get(m.slide), m.slide, m, q, rep)
          if (h) { n++; touched.set(`${m.slide}|${m.host}`, [h, m.slide]) }
        }
        if (touched.size) this.mark()
        for (const [h, i] of touched.values()) this.queue.push(this.opFor(h, i))
        if (n) { this.doc.dirty(true); this.renderAfter = true }
        return n
      }

      selectedText () {
        if (this.editing) { const r = currentRange(this.editing.node); return r ? r.toString() : '' }
        return ''
      }

      focus () { window.CSS.highlights?.delete?.('q-find'); this.stage.focus({ preventScroll: true }) }

      // ------------------------------------------------ slides

      newSlide (layout) {
        this.switchTo = this.ix + 1
        this.push({ op: 'new_slide', layout, at: this.ix + 2 })
      }

      slideOp (op, shift) {
        if (!this.doc.writable) return
        this.switchTo = this.ix + shift
        this.push({ ...op, slide: this.slide().n })
      }

      deleteShape () {
        if (!this.doc.writable) return
        const keys = this.multi.filter((k) => this.itemEls(k).length)
        if (!keys.length) return
        // Later shapes first: deleting one renumbers those after it.
        const first = (k) => Math.min(...this.itemEls(k).map((n) => +n.dataset.k))
        for (const k of keys.sort((a, b) => first(b) - first(a))) {
          this.push({ op: 'delete_shape', ...this.at(), ...(String(k).startsWith('g') ? { group: +k.slice(1) } : { shape: +k }) })
          for (const n of this.itemEls(k)) n.remove()
        }
        this.select(null)
      }

      saveNotes () {
        clearTimeout(this.notesTimer)
        if (!this.model || !this.doc.writable || this.notesShown === undefined || this.notes.value === this.notesShown) return
        this.notesShown = this.notes.value
        this.push({ op: 'notes', slide: this.notesSlide, text: this.notes.value }, { quiet: true })
      }

      // ------------------------------------------------ saving

      push (op, { quiet } = {}) {
        if (!this.doc.writable) return
        if (!quiet && this.editing) { this.endEdit(true); this.stepOpen = false }
        this.mark()
        this.queue.push(op)
        this.doc.dirty(true)
      }

      // mark starts an undo step: ops pushed in one go (one action) share it.
      mark () {
        if (this.stepOpen) return
        this.stepOpen = true
        ;(this.marks ??= []).push(this.queue.length)
        queueMicrotask(() => { this.stepOpen = false })
      }

      hasChanges () { return this.queue.length > 0 }

      shortcuts () {
        return Object.assign([['⌘B / ⌘I / ⌘U', '加粗 / 倾斜 / 下划线'], ['Enter / F2', '编辑所选文本框的文字'], ['Esc', '结束编辑 / 取消选择'], ['Tab / ⇧Tab', '编辑时提高 / 降低列表级别'], ['⌘C / ⌘X / ⌘V', '复制 / 剪切 / 粘贴形状'], ['⌘D', '创建副本'], ['⌘A', '选中所有形状'], ['⌘G / ⇧⌘G', '组合 / 取消组合'], ['Shift+单击', '多选；拖动空白处框选'], ['方向键 / ⇧方向键', '微移 1 / 10 像素'], ['Delete', '删除所选'], ['PageUp / PageDown', '上一张 / 下一张'], ['F5 / ⇧F5', '从头放映 / 从当前放映'], ['⌥F5', '演示者视图'], ['拖动 + Shift', '水平或垂直移动；角点保持比例；旋转按 15°'], ['拖动 + ⌥', '不吸附参考线'], ['Enter / Esc', '裁剪时完成 / 取消']], { title: '演示文稿' })
      }

      collect () {
        if (!this.queue.length) return null
        const ops = this.queue.splice(0)
        const cuts = [...new Set([0, ...(this.marks ?? []).filter((i) => i > 0 && i < ops.length)])].sort((a, b) => a - b)
        this.marks = []
        return { ops, steps: cuts.map((c, i) => (cuts[i + 1] ?? ops.length) - c) }
      }

      saved (model, batch) {
        const shapeOps = !this.renderAfter && batch.ops.every((o) => o.op === 'xfrm' || o.op === 'shape_text' || o.op === 'notes' || o.op === 'table_cell' || o.op === 'crop')
        this.renderAfter = false
        const prevCount = this.items().length
        this.model = model
        if (this.mode === 'master' && !model.masters?.length) this.mode = 'slides'
        this.fillLayouts()
        if (this.switchTo !== undefined) { this.ix = Math.max(0, Math.min(this.switchTo, model.slides.length - 1)); this.switchTo = undefined }
        this.ix = Math.min(this.ix, model.slides.length - 1)
        this.mix = Math.min(this.mix, (model.masters?.length ?? 1) - 1)
        // Thumbnails always follow; the slide on stage is left alone while
        // its text is being edited, and when only shapes moved or text
        // changed, as the page already shows that.
        this.renderThumbs()
        if (this.editing || this.crop) return
        if (shapeOps && prevCount === this.items().length) {
          if (this.mode === 'slides') this.notesShown = model.slides[this.ix].notes ?? ''
          const fresh = el('div', { html: this.cur().html })
          // Pick up new boxes and empty markers without redrawing.
          for (const n of fresh.querySelectorAll('[data-k]')) { const cur = this.shapeEl(n.dataset.k); if (cur) { cur.dataset.box = n.dataset.box; if (n.dataset.gbox) cur.dataset.gbox = n.dataset.gbox } }
          this.placeBox()
          this.refreshBar()
          this.paintBadges()
          return
        }
        const keep = this.keepMulti
        this.keepMulti = undefined
        this.show(this.pos(), true)
        if (this.selectLast) {
          const mode = this.selectLast
          this.selectLast = false
          const all = [...this.slideEl.querySelectorAll('[data-k]')]
          const last = all.sort((a, b) => +a.dataset.k - +b.dataset.k).pop()
          if (last) { this.select(this.keyOf(last)); if (mode === 'edit') this.beginEdit(last) }
        } else if (this.reselect) {
          const keys = [...new Set(this.reselect.map((id) => this.slideEl.querySelector(`[data-id="${id}"]`)).filter(Boolean).map((n) => this.keyOf(n)))]
          this.reselect = undefined
          this.multi = keys
          this.selected = keys[keys.length - 1] ?? null
          this.placeBox(); this.refreshBar()
        } else if (keep?.length) {
          this.multi = keep.filter((k) => this.itemEls(k).length)
          this.selected = this.multi[this.multi.length - 1] ?? null
          this.placeBox(); this.refreshBar()
        } else if (this.selectAfter) {
          const k = this.selectAfter
          this.selectAfter = undefined
          if (this.shapeEl(k)) this.select(k)
        }
        this.paintBadges()
        this.paintPane()
        if (this.fitAfter) { const n = this.shapeEl(this.fitAfter); this.fitAfter = null; if (n) this.autoFit(n) }
        if (this.previewNext) { this.previewNext = false; this.previewTransition() }
      }

      failed () { this.renderAfter = false; this.reselect = undefined; this.show(this.pos(), true) }
    }

    // ------------------------------------------------------------ dsh glue

    async function start (host, address, onReady, onFail) {
      const file = parseFileAddress(address)
      const frame = new Frame(host)
      const offPins = pinStore.subscribe((sid) => { if (sid === file?.sessionId) frame.syncPins() })
      pinStore.frames.add(frame)
      if (!file) { frame.fail('这个文件不属于当前会话，Quire 无法打开。'); onFail(); return { dispose: () => frame.dispose() } }
      const doc = new Doc(frame, file)
      frame.doc = doc
      frame.main.append(el('div', { class: 'q-loading' }, '正在打开…'))
      const ctl = {
        doc,
        disposed: false,
        dispose () {
          this.disposed = true
          doc.disposed = true
          // Save what is pending before the tab goes.
          if (doc.editor?.hasChanges?.()) doc.kick()
          doc.editor?.dispose?.()
          offPins()
          pinStore.frames.delete(frame)
          frame.dispose()
        }
      }
      try {
        const r = await doc.open()
        if (ctl.disposed) return ctl
        const Editor = { xlsx: XlsxEditor, docx: DocxEditor, pptx: PptxEditor }[r.kind]
        frame.main.replaceChildren()
        const editor = new Editor(frame, doc)
        doc.editor = editor
        frame.editor = editor
        frame.readOnly(r.writable ? undefined : r.why)
        editor.load(r.model, {})
        frame.syncPins()
        pinStore.load(file.sessionId)
        onReady()
      } catch (err) {
        frame.fail(`打不开这个文件：${err.message}`)
        onFail()
      }
      return ctl
    }

    function QuireBody (props) {
      const hostRef = useRef(null)
      const ctlRef = useRef(null)
      const content = props.content
      const request = content?.kind === 'renderer' ? content : undefined
      const meta = props.useResource ? props.useResource(props.resourceAddress) : undefined
      const metaVersion = meta?.value?.version
      const metaRef = useRef(metaVersion)
      metaRef.current = metaVersion
      const reqRef = useRef(request)
      reqRef.current = request
      useEffect(() => {
        let ctl
        let gone = false
        start(hostRef.current, props.resourceAddress,
          () => reqRef.current?.loaded(String(metaRef.current ?? 'quire')),
          () => reqRef.current?.failed()
        ).then((c) => { if (gone) c.dispose(); else { ctl = c; ctlRef.current = c } })
        return () => { gone = true; ctl?.dispose(); ctlRef.current = null }
      }, [props.resourceAddress])

      // The preview reloads on file changes, ours included: keep the
      // editor, and only show content that differs from what it has.
      useEffect(() => {
        const ctl = ctlRef.current
        if (!ctl || !request) return
        ctl.doc.remote().finally(() => request.loaded(String(metaRef.current ?? 'quire')))
      }, [request?.revision])

      useEffect(() => {
        const ctl = ctlRef.current
        if (!ctl || metaVersion === undefined) return
        ctl.doc.remote().finally(() => reqRef.current?.loaded(String(metaVersion)))
      }, [metaVersion])

      return React.createElement('div', { ref: hostRef, className: 'quire-host' })
    }

    // QuirePinDock: the pins waiting in this Session, folded into one pill
    // above the composer. The pill opens the list: each pin with its
    // picture and note, which can be edited or the pin removed; clicking
    // one shows the spot in an open editor. They go out with the next
    // message.
    function QuirePinDock (props) {
      const h = React.createElement
      const sid = props.sessionId
      const [, tick] = React.useReducer((x) => x + 1, 0)
      const [open, setOpen] = React.useState(false)
      const [editing, setEditing] = React.useState(null)
      const [text, setText] = React.useState('')
      const [error, setError] = React.useState('')
      const root = useRef(null)
      useEffect(() => pinStore.subscribe((s) => { if (s === sid) tick() }), [sid])
      // The server takes the pins when a sent message reaches the Agent;
      // look again whenever the composer or the Session moves on.
      const input = props.input
      const session = props.session
      useEffect(() => { pinStore.load(sid) }, [sid, input?.phase, input?.queue?.length, session?.running, session?.pendingSubmissions?.length])
      const items = pinStore.get(sid).items
      const count = items.length
      useEffect(() => { if (!count) { setOpen(false); setEditing(null) } }, [count])
      // An open list folds away on a click elsewhere or Esc.
      useEffect(() => {
        if (!open) return
        const away = (e) => { if (root.current && !root.current.contains(e.target)) { setOpen(false); setEditing(null) } }
        const esc = (e) => { if (e.key === 'Escape' && !e.defaultPrevented) setOpen(false) }
        document.addEventListener('pointerdown', away, true)
        document.addEventListener('keydown', esc)
        return () => { document.removeEventListener('pointerdown', away, true); document.removeEventListener('keydown', esc) }
      }, [open])
      if (!count) return null
      const act = (op, extra) => pinStore.call(sid, op, extra).then(() => setError(''), (err) => setError(err.message))
      const save = (id) => { const note = text.trim(); if (!note) return; act('edit', { id, note }); setEditing(null) }
      const empty = !String(input?.draft ?? '').trim() && !(input?.attachmentIds?.length)
      const send = () => {
        const a = props.inputActions
        if (!a) return
        setOpen(false)
        a.setDraft('请按这些标注处理。')
        setTimeout(() => a.submit())
      }
      const card = (p) => {
        const on = editing === p.id
        return h('div', { key: p.id, className: 'qpd-card' + (on ? ' on' : ''), title: on ? undefined : '在编辑器里查看这处', onClick: () => { if (!on) pinStore.reveal(sid, p) } },
          p.thumb
            ? h('img', { className: 'qpd-thumb', src: p.thumb, alt: '标注截图' })
            : h('span', { className: 'qpd-ico', dangerouslySetInnerHTML: { __html: icon('pin') } }),
          h('div', { className: 'qpd-body' },
            h('div', { className: 'qpd-where' }, h('b', null, p.path), ' · ', p.hasImage ? pinWhere(p.where) : p.where, p.quote ? h('span', { className: 'qpd-q' }, `「${clipPin(p.quote, 40)}」`) : null, p.extra ? h('span', { className: 'qpd-q' }, `（${clipPin(p.extra, 40)}）`) : null),
            on
              ? h('textarea', {
                  className: 'qpd-in', rows: 2, value: text, 'aria-label': '标注内容',
                  // The note starts selected: typing replaces it.
                  ref: (n) => { if (n && !n.dataset.armed) { n.dataset.armed = '1'; n.focus(); n.select() } },
                  onClick: (e) => e.stopPropagation(),
                  onChange: (e) => setText(e.target.value),
                  onKeyDown: (e) => {
                    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); setEditing(null) } else if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) { e.preventDefault(); save(p.id) }
                  },
                  onBlur: () => save(p.id)
                })
              : h('div', { className: 'qpd-note' }, p.note || '（没有写说明）')),
          on
            ? null
            : h('span', { className: 'qpd-acts' },
              h('button', { type: 'button', className: 'qpd-btn', title: '修改说明', 'aria-label': '修改说明', onClick: (e) => { e.stopPropagation(); setText(p.note); setEditing(p.id) }, dangerouslySetInnerHTML: { __html: '<svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12.5 4.5l3 3L7 16H4v-3z"/><path d="M11 6l3 3"/></svg>' } }),
              h('button', { type: 'button', className: 'qpd-btn', title: '移除这条标注', 'aria-label': '移除这条标注', onClick: (e) => { e.stopPropagation(); act('remove', { id: p.id }) } }, '×')))
      }
      return h('div', { ref: root, className: 'qpd' + (open ? ' open' : ''), 'data-quire-pins': count },
        h('div', { className: 'qpd-list', hidden: !open }, items.map(card)),
        h('div', { className: 'qpd-foot' },
          h('span', { className: 'qpd-pill' },
            h('button', { type: 'button', className: 'qpd-tog', 'aria-expanded': open, title: open ? '收起标注' : '查看标注', onClick: () => { setOpen(!open); setEditing(null) } },
              h('span', { className: 'qpd-pi', dangerouslySetInnerHTML: { __html: icon('pin') } }),
              `${count} 条标注`,
              h('span', { className: 'qpd-chev', 'aria-hidden': true, dangerouslySetInnerHTML: { __html: '<svg viewBox="0 0 12 12" width="10" height="10" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7.5 6 4.5 9 7.5"/></svg>' } })),
            h('button', { type: 'button', className: 'qpd-x', title: '清空全部标注', 'aria-label': '清空全部标注', onClick: () => act('clear') }, '×')),
          h('span', { className: 'qpd-hint' }, error || '随下一条消息发给 Agent'),
          empty && props.inputActions ? h('button', { type: 'button', className: 'qpd-send', onClick: send }, '直接发送') : null))
    }

    // ------------------------------------------------------------ styles
    const CSS = `
.quire-host{position:relative;height:100%;min-height:420px;display:flex;flex-direction:column}
.quire{--q-bg:var(--dsw-alias-bg-base,#fff);--q-l1:var(--dsw-alias-bg-layer-1,#f6f7f8);--q-l2:var(--dsw-alias-bg-layer-2,#eceef1);--q-b1:var(--dsw-alias-border-l1,#e3e5e8);--q-b2:var(--dsw-alias-border-l2,#cfd3d8);--q-brand:var(--dsw-alias-brand-primary,#1a73e8);--q-t1:var(--dsw-alias-label-primary,#1f2328);--q-t2:var(--dsw-alias-label-secondary,#59636e);--q-t3:var(--dsw-alias-label-tertiary,#8b949e);
  flex:1;min-height:0;display:flex;flex-direction:column;background:var(--q-bg);color:var(--q-t1);font:13px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;outline:none;position:relative}
.quire.q-full{position:fixed;inset:0;z-index:2147483000}
.quire *{box-sizing:border-box}
.q-head{display:flex;align-items:center;gap:6px;padding:2px 6px 0;background:var(--q-l1);min-height:32px}
.q-tabs{flex:1;min-width:0;display:flex;flex-wrap:wrap;align-items:flex-end;gap:1px;align-self:stretch}
.q-tab{border:0;background:transparent;color:var(--q-t2);font:inherit;font-size:13px;padding:5px 8px 6px;cursor:pointer;white-space:nowrap;border-radius:6px 6px 0 0}
.q-tab:hover{color:var(--q-t1);background:color-mix(in srgb,var(--q-t1) 6%,transparent)}
.q-tab.on{color:var(--q-brand);background:var(--q-bg);font-weight:600;box-shadow:0 -1px 0 var(--q-b1),1px 0 0 var(--q-b1),-1px 0 0 var(--q-b1)}
.q-bar{display:flex;align-items:center;padding:5px 6px;border-bottom:1px solid var(--q-b1);background:var(--q-bg);min-height:38px}
.q-tools{flex:1;min-width:0}
.q-panel{display:flex;flex-wrap:wrap;align-items:center;gap:2px}
.q-panel[hidden]{display:none}
.q-hint{font-size:12px;color:var(--q-t3);padding:0 6px}
.q-hint[hidden]{display:none}
.q-side{display:flex;align-items:center;gap:2px;flex-shrink:0}
.q-btn{display:inline-flex;align-items:center;justify-content:center;min-width:26px;height:26px;padding:0 5px;border:1px solid transparent;border-radius:5px;background:transparent;color:var(--q-t1);font:inherit;font-size:13px;cursor:pointer;position:relative}
.q-btn:hover:not(:disabled){background:var(--q-l2)}
.q-btn.on{background:color-mix(in srgb,var(--q-brand) 16%,transparent);color:var(--q-brand);border-color:color-mix(in srgb,var(--q-brand) 30%,transparent)}
.q-btn:disabled{opacity:.38;cursor:default}
.q-btn.q-b{font-weight:700}.q-btn.q-i{font-style:italic;font-family:Georgia,serif}.q-btn.q-u{text-decoration:underline}.q-btn.q-s{text-decoration:line-through}
.q-btn .q-lbl{margin-left:4px;font-size:12px;white-space:nowrap}
.q-btn.q-mini svg{width:14px;height:14px;opacity:.75}
.q-color{gap:2px;padding:0 4px}
.q-color .q-glyph{font-weight:600}
.q-color .q-sw{width:12px;height:4px;border-radius:1px;position:absolute;bottom:3px;left:50%;transform:translateX(-50%)}
.q-color input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;border:0;padding:0}
.q-sel{height:26px;border:1px solid var(--q-b1);border-radius:5px;background:var(--q-bg);color:var(--q-t1);font:inherit;font-size:12px;padding:0 4px;max-width:140px}
.q-sel:disabled{opacity:.45}
.q-sep{width:1px;height:18px;background:var(--q-b2);margin:0 3px}
.q-status{font-size:12px;color:var(--q-t3);padding:0 6px;white-space:nowrap}
.q-status[data-state=saving],.q-status[data-state=pending]{color:var(--q-t2)}
.q-status[data-state=error]{color:#d1242f}
.q-banner{padding:6px 10px;background:#fff8c5;color:#4d3800;border-bottom:1px solid #eedc82;font-size:12px}
.q-main{flex:1;min-height:0;position:relative;display:flex;flex-direction:column}
.q-loading,.q-fail{margin:auto;color:var(--q-t2);padding:24px;text-align:center}
.q-toasts{position:absolute;right:12px;bottom:12px;display:flex;flex-direction:column;gap:6px;z-index:20;max-width:min(380px,80%);pointer-events:none}
.q-toast{background:#24292f;color:#fff;padding:8px 12px;border-radius:7px;font-size:12px;box-shadow:0 4px 14px rgba(0,0,0,.18);pointer-events:auto}
.q-toast.q-warn{background:#9a3412}

.qx{flex:1;min-height:0;display:flex;flex-direction:column;position:relative}
.qx-fbar{display:flex;align-items:center;border-bottom:1px solid var(--q-b1);height:28px;background:var(--q-bg)}
.qx-name{width:84px;height:100%;border:0;border-right:1px solid var(--q-b1);padding:0 6px;font:12px Menlo,Consolas,monospace;background:transparent;color:var(--q-t1);outline:none}
.qx-fxl{padding:0 8px;color:var(--q-t3);font-style:italic;font-family:Georgia,serif}
.qx-fx{flex:1;height:100%;border:0;padding:0 6px;font:13px Menlo,Consolas,"PingFang SC",monospace;background:transparent;color:var(--q-t1);outline:none}
.qx-grid{flex:1;min-height:0;position:relative;overflow:hidden;background:#fff}
.qx-canvas{position:absolute;left:0;top:0;pointer-events:none}
.qx-scroll{position:absolute;inset:0;overflow:auto;cursor:cell}
.qx-spacer{pointer-events:none}
.qx-input{position:absolute;width:1px;height:1px;opacity:0;resize:none;border:0;padding:1px 3px;margin:0;outline:none;overflow:hidden;white-space:pre-wrap;z-index:3;line-height:1.25;color:#000;background:#fff}
.qx-input.on{opacity:1;border:2px solid #217346;box-shadow:0 2px 8px rgba(0,0,0,.18)}
.qx-guide{position:absolute;background:#1a73e8;z-index:4;pointer-events:none}
.qx-charts{position:absolute;right:0;bottom:0;overflow:hidden;pointer-events:none;z-index:2}
.qx-chart{position:absolute;pointer-events:auto;background:#fff;border:1px solid #d0d4d9;box-sizing:border-box;cursor:move;outline:none}
.qx-chart[hidden]{display:none}
.qx-chart-body{position:absolute;inset:0;overflow:hidden}
.qx-chart-body svg{display:block}
.qx-chart.on{border:1.5px solid #1a73e8}
.qx-ch-h{position:absolute;width:8px;height:8px;background:#fff;border:1.5px solid #1a73e8;border-radius:50%;box-sizing:border-box;display:none;margin:-4px 0 0 -4px}
.qx-chart.on .qx-ch-h{display:block}
.qx-ch-h[data-h=nw]{left:0;top:0;cursor:nwse-resize}.qx-ch-h[data-h=n]{left:50%;top:0;cursor:ns-resize}.qx-ch-h[data-h=ne]{left:100%;top:0;cursor:nesw-resize}.qx-ch-h[data-h=e]{left:100%;top:50%;cursor:ew-resize}
.qx-ch-h[data-h=se]{left:100%;top:100%;cursor:nwse-resize}.qx-ch-h[data-h=s]{left:50%;top:100%;cursor:ns-resize}.qx-ch-h[data-h=sw]{left:0;top:100%;cursor:nesw-resize}.qx-ch-h[data-h=w]{left:0;top:50%;cursor:ew-resize}
.qx-guide.v{top:0;bottom:0;width:1px}.qx-guide.h{left:0;right:0;height:1px}
.qx-foot{display:flex;align-items:center;border-top:1px solid var(--q-b1);background:var(--q-l1);height:30px;min-height:30px}
.qx-tabs{flex:1;min-width:0;display:flex;overflow-x:auto;height:100%}
.qx-tab{border:0;border-right:1px solid var(--q-b1);background:transparent;color:var(--q-t2);font:inherit;font-size:12px;padding:0 14px;cursor:pointer;white-space:nowrap}
.qx-tab.on{background:var(--q-bg);color:var(--q-brand);font-weight:600;box-shadow:inset 0 2px 0 var(--q-brand)}
.qx-add{padding:0 8px;display:flex;align-items:center}
.qx-stats{padding:0 12px;font-size:12px;color:var(--q-t2);white-space:nowrap}
.q-menu{position:fixed;z-index:2147483600;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#1f2328);border:1px solid var(--dsw-alias-border-l2,#cfd3d8);border-radius:8px;box-shadow:0 10px 28px rgba(0,0,0,.18);padding:4px;display:flex;flex-direction:column;min-width:168px;max-height:calc(100vh - 8px);overflow:auto;font:13px -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif}
.q-menu>button{display:flex;align-items:center;justify-content:space-between;gap:18px;border:0;background:transparent;text-align:left;padding:5px 10px;border-radius:5px;font:inherit;cursor:pointer;color:inherit;white-space:nowrap}
.q-menu>button:hover:not(:disabled){background:var(--dsw-alias-brand-primary,#1a73e8);color:#fff}
.q-menu>button:disabled{opacity:.4;cursor:default}
.q-mi{display:inline-flex;align-items:center;gap:8px}
.q-space{flex:1}
.q-btn.q-more{min-width:12px;padding:0 1px;margin-left:-3px;font-size:9px;color:var(--q-t3)}
.q-zoom{display:inline-flex;align-items:center;gap:2px;padding:0 8px;font-size:12px;color:var(--q-t2);flex-shrink:0}
.q-zoom-r{width:104px;accent-color:var(--q-brand);margin:0 2px}
.q-zoom button{border:0;background:transparent;color:var(--q-t2);cursor:pointer;font:inherit;padding:0 4px;border-radius:3px;line-height:18px}
.q-zoom button:hover{background:var(--q-l2)}
.q-zoom-p{min-width:44px;text-align:right}
.qx-zoom{display:none}
.qd-stxt,.qp-stxt{overflow:hidden;text-overflow:ellipsis}
.qp-status{height:24px;min-height:24px;display:flex;align-items:center;padding:0 0 0 12px;border-top:1px solid var(--q-b1);background:var(--q-l1);font-size:12px;color:var(--q-t2);white-space:nowrap}
.q-kind{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:3px;color:#fff;font:700 10px/1 -apple-system,sans-serif}
.q-kind-docx{background:#2b579a}.q-kind-xlsx{background:#217346}.q-kind-pptx{background:#c43e1c}
.q-keys-wrap{max-height:60vh;overflow:auto;margin-top:4px}
.q-keys-wrap h4{margin:10px 0 4px;font-size:12px;color:var(--q-t2)}
.q-keys{border-collapse:collapse;width:100%;font-size:12px}
.q-keys td{padding:4px 6px;border-bottom:1px solid var(--q-b1);vertical-align:top}
.q-keys td:first-child{width:44%}
.q-keys kbd{font:11px -apple-system,sans-serif;background:var(--q-l1);border:1px solid var(--q-b2);border-radius:4px;padding:1px 5px;white-space:nowrap}
.q-swatches{display:inline-flex;gap:1px;vertical-align:middle}
.q-swatches i{display:inline-block;width:10px;height:12px;border-radius:1px}
.q-menu kbd{font:12px -apple-system,sans-serif;opacity:.6}
.q-menu-sep{height:1px;background:var(--dsw-alias-border-l1,#e3e5e8);margin:4px 6px}
.q-filter{width:240px;padding:6px}
.q-filter-q{margin:4px 2px;height:26px;border:1px solid var(--dsw-alias-border-l2,#cfd3d8);border-radius:5px;padding:0 8px;font:inherit;background:transparent;color:inherit;outline:none}
.q-filter-list{max-height:220px;overflow:auto;border:1px solid var(--dsw-alias-border-l1,#e3e5e8);border-radius:5px;margin:2px;padding:2px 0}
.q-filter-item{display:flex;align-items:center;gap:6px;padding:3px 8px;cursor:pointer;white-space:nowrap}
.q-filter-item span{flex:1;overflow:hidden;text-overflow:ellipsis}
.q-filter-item em{font-style:normal;opacity:.5;font-size:12px}
.q-filter-item[hidden]{display:none}
.q-filter-all{padding:5px 10px}
.q-filter-btns{display:flex;justify-content:flex-end;gap:6px;padding:6px 2px 2px}
.q-filter-btns button{height:26px;padding:0 14px;border-radius:5px;border:1px solid var(--dsw-alias-border-l2,#cfd3d8);background:transparent;color:inherit;font:inherit;cursor:pointer}
.q-filter-btns .q-filter-ok{background:var(--dsw-alias-brand-primary,#1a73e8);border-color:transparent;color:#fff}
.q-find{display:flex;align-items:center;gap:4px;flex-wrap:wrap;padding:4px 6px;border-bottom:1px solid var(--q-b1);background:var(--q-l1)}
.q-find[hidden]{display:none}
.q-find-in{height:26px;width:160px;border:1px solid var(--q-b2);border-radius:5px;padding:0 8px;background:var(--q-bg);color:var(--q-t1);font:inherit;outline:none}
.q-find-in:focus{border-color:var(--q-brand)}
.q-find-n{font-size:12px;color:var(--q-t3);min-width:44px;text-align:center}
.q-find-r{display:inline-flex;align-items:center;gap:4px}
.q-find-r[hidden]{display:none}
.q-find-b{font-size:12px}
.q-print{position:fixed;width:1px;height:1px;left:-9999px;top:0;border:0;opacity:0}
.qx-zoom{margin:0 6px;flex-shrink:0}
.qx-ac{position:absolute;z-index:30;min-width:220px;max-width:420px;background:var(--q-bg);border:1px solid var(--q-b2);border-radius:7px;box-shadow:0 8px 22px rgba(0,0,0,.16);padding:3px;font-size:12px}
.qx-ac[hidden]{display:none}
.qx-ac-item{display:flex;gap:10px;align-items:baseline;padding:4px 8px;border-radius:4px;cursor:pointer}
.qx-ac-item b{font-family:ui-monospace,Menlo,monospace;font-weight:600;min-width:70px}
.qx-ac-item span{color:var(--q-t2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.qx-ac-item.on{background:color-mix(in srgb,var(--q-brand) 14%,transparent)}
.qx-ac-tip{padding:5px 8px;color:var(--q-t2);font-family:ui-monospace,Menlo,monospace}
.qx-fxl{border:0;background:none;cursor:pointer;font-style:italic;color:var(--q-t2);padding:0 6px;font-family:Georgia,serif}
.qx-fxl:hover{color:var(--q-brand)}
.qx-tip{position:absolute;z-index:25;max-width:260px;background:#fffbe6;border:1px solid #e3c86a;border-radius:5px;padding:5px 8px;font-size:12px;color:#5c4400;box-shadow:0 3px 10px rgba(0,0,0,.12);pointer-events:none;display:flex;flex-direction:column;gap:2px}
.qx-tip[hidden]{display:none}
.q-dlg-back{position:fixed;inset:0;z-index:2147483500;background:rgba(0,0,0,.28);display:flex;align-items:flex-start;justify-content:center;padding-top:12vh}
.q-dlg{background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#1f2328);border-radius:10px;box-shadow:0 18px 50px rgba(0,0,0,.28);width:380px;max-width:calc(100vw - 32px);max-height:76vh;overflow:auto;padding:16px 18px 14px;display:flex;flex-direction:column;gap:10px;font:13px -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif}
.q-dlg-wide{width:560px}
.q-dlg-h{font-size:15px;font-weight:600;margin-bottom:2px}
.q-dlg-row{display:grid;grid-template-columns:76px 1fr;align-items:center;gap:4px 10px}
.q-dlg-row[hidden]{display:none}
.q-dlg-row>span{color:var(--q-t2,#57606a)}
.q-dlg-row input:not([type=checkbox]),.q-dlg-row select,.q-dlg-row textarea{width:100%;box-sizing:border-box;padding:5px 7px;border:1px solid var(--q-b2,#cfd3d8);border-radius:6px;background:transparent;color:inherit;font:inherit}
.q-dlg-row small{grid-column:2;color:var(--q-t3,#8c959f);font-size:11px}
.q-dlg-check input{justify-self:start}
.q-dlg-err{color:#cf222e;font-size:12px}
.q-dlg-err[hidden]{display:none}
.q-dlg-f{display:flex;gap:8px;align-items:center;margin-top:4px}
.q-dlg-f button{padding:5px 14px;border-radius:6px;border:1px solid var(--q-b2,#cfd3d8);background:transparent;color:inherit;cursor:pointer;font:inherit}
.q-dlg-f .q-dlg-ok{background:var(--q-brand,#1a73e8);border-color:var(--q-brand,#1a73e8);color:#fff}
.q-fn{display:flex;flex-direction:column;gap:8px}
.q-fn-top{display:flex;gap:8px}
.q-fn-q{flex:1;padding:6px 8px;border:1px solid var(--q-b2,#cfd3d8);border-radius:6px;background:transparent;color:inherit;font:inherit}
.q-fn-list{height:200px;overflow:auto;border:1px solid var(--q-b2,#cfd3d8);border-radius:6px;padding:3px;font-family:ui-monospace,Menlo,monospace}
.q-fn-it{padding:3px 8px;border-radius:4px;cursor:pointer}
.q-fn-it.on{background:var(--q-brand,#1a73e8);color:#fff}
.q-fn-help{min-height:44px}
.q-fn-help b{font-family:ui-monospace,Menlo,monospace;font-weight:600}
.q-fn-help p{margin:4px 0 0;color:var(--q-t2,#57606a)}
.q-names{display:flex;flex-direction:column;border:1px solid var(--q-b2,#cfd3d8);border-radius:6px;max-height:300px;overflow:auto}
.q-names-r{display:grid;grid-template-columns:1.2fr 1.6fr 1fr auto;gap:8px;align-items:center;padding:6px 8px;border-top:1px solid var(--q-b1,#eaeef2)}
.q-names-r:first-child{border-top:0}
.q-names-h{font-size:12px;color:var(--q-t3,#8c959f);position:sticky;top:0;background:var(--dsw-alias-bg-base,#fff)}
.q-names-r code{font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.q-names-a{display:flex;gap:4px}
.q-names-a button{padding:2px 8px;border:1px solid var(--q-b2,#cfd3d8);border-radius:5px;background:transparent;color:inherit;cursor:pointer;font-size:12px}
.q-names-empty{padding:16px;color:var(--q-t3,#8c959f);text-align:center}
.qx-pop{position:absolute;z-index:26;max-width:240px;background:#fff;border:1px solid #c9ccd1;border-radius:5px;padding:6px 8px;font-size:12px;color:#1f2328;box-shadow:0 4px 14px rgba(0,0,0,.14);pointer-events:none;display:flex;flex-direction:column;gap:4px;white-space:pre-wrap;word-break:break-all}
.qx-pop[hidden]{display:none}
.qx-pop-note{background:#fffbe6;margin:-6px -8px;padding:6px 8px;border-radius:5px}
.qx-pop-link{color:#0563c1;display:flex;flex-direction:column}
.qx-pop-link small{color:#8c959f}
.qx-notes{position:absolute;right:0;bottom:0;overflow:hidden;pointer-events:none;z-index:4}
.qx-note{position:absolute;width:150px;background:#fffbe6;border:1px solid #e3c86a;border-radius:4px;padding:4px 6px;font-size:11px;color:#3d3000;box-shadow:0 2px 6px rgba(0,0,0,.12);white-space:pre-wrap;word-break:break-all}
.qx-pic .qx-chart-body{background:transparent;display:flex}
.qx-pic img{width:100%;height:100%;object-fit:fill;pointer-events:none;user-select:none}
.qx-pic-x{margin:auto;color:#8c959f;font-size:12px}
.q-sort{display:flex;flex-direction:column;gap:6px}
.q-sort-r{display:grid;grid-template-columns:76px 1fr 90px 28px;gap:8px;align-items:center}
.q-sort-r>span{color:var(--q-t2,#57606a)}
.q-sort-r button{border:1px solid var(--q-b2,#cfd3d8);border-radius:5px;background:transparent;color:inherit;cursor:pointer;height:26px}
.q-cf-sw{display:inline-block;font-style:normal;font-size:11px;padding:2px 6px;border:1px solid var(--q-b2,#cfd3d8);border-radius:3px;min-width:44px;height:16px;line-height:16px;text-align:center}

.qd-scroll{flex:1;min-height:0;overflow:auto;background:var(--q-l2);padding:16px 12px 48px}
.qd-zoom{display:flex;justify-content:center}
.qd-sheet{color:#000;font-family:Calibri,Carlito,"Helvetica Neue",Arial,"PingFang SC","Microsoft YaHei",sans-serif;line-height:1.45;flex-shrink:0;position:relative}
.qd-pages{position:absolute;inset:0;z-index:0}
.qd-pg{position:absolute;left:0;right:0;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.18)}
.qd-hfc{position:absolute;left:0;right:0;display:flex;flex-direction:column;justify-content:flex-end;white-space:pre-wrap;font-size:9pt;color:#555;text-align:center;padding-top:24px;padding-bottom:10px;overflow:hidden;cursor:default}
.qd-hfc-f{justify-content:flex-start;padding-top:10px;padding-bottom:24px}
.qd-page{outline:none;caret-color:#000;position:relative;z-index:1;display:flex;flex-direction:column}
.qd-page>.qd-p{margin-top:calc(var(--mt,0pt) + var(--push,0px))}
.qd-page>.qd-t{margin-top:calc(6pt + var(--push,0px))}
.qd-hf{position:absolute;left:0;right:0;z-index:2;display:flex;flex-direction:column;justify-content:flex-end;white-space:pre-wrap;font-size:9pt;color:#555;text-align:center;outline:none;padding-top:24px;padding-bottom:10px;border-bottom:1px dashed transparent;overflow:hidden}
.qd-ftr{justify-content:flex-start;padding-top:10px;padding-bottom:24px}
.qd-status{height:24px;min-height:24px;display:flex;align-items:center;padding:0 12px;border-top:1px solid var(--q-b1);background:var(--q-l1);font-size:12px;color:var(--q-t2);white-space:nowrap;overflow:hidden}
.qd-hf:hover:not(.qd-hf-ro),.qd-hf:focus{border-color:#c7cbd1;background:#fafbfc}
.qd-ftr:hover:not(.qd-hf-ro),.qd-ftr:focus{border-bottom-color:transparent;border-top:1px dashed #c7cbd1}
.qd-hf:empty::before{content:attr(data-ph);color:#c0c4ca}
.qd-hf.qd-hf-ro:empty::before{content:""}
.qd-hf.qd-hf-ro{cursor:default}
::highlight(q-find){background:#ffd54f;color:#000}
.qd-p{white-space:pre-wrap;word-break:break-word;position:relative;min-height:1.45em}
.qd-p[data-list]::before{content:attr(data-mark);position:absolute;margin-left:calc(-4px - var(--mw,20px));width:var(--mw,20px);text-align:right;font-weight:400;font-style:normal;white-space:nowrap}
.qd-page>.qd-p.qd-drop{position:absolute;margin:0;line-height:1;z-index:2;padding:0}
.qd-p.qd-dropnext::before{content:'';float:left;width:var(--dw);height:var(--dh)}
.qd-painting{cursor:copy}
.qd-del{color:#c62828;text-decoration:line-through;cursor:default}
.qd-fn{vertical-align:super;font-size:0.7em;color:#1a5fb4;cursor:pointer;padding:0 1px}
.qd-notes{position:absolute;z-index:1;font-size:9pt;line-height:1.4}
.qd-notes::before{content:'';display:block;width:33%;border-top:1px solid #888;margin-bottom:4px}
.qd-notes-h{font-size:8pt;color:#888;margin:4px 0 2px}
.qd-note{display:flex;gap:4px;align-items:baseline}
.qd-note-n{vertical-align:super;font-size:0.75em;color:#1a5fb4;cursor:pointer;min-width:12px}
.qd-note-t{flex:1;outline:none;white-space:pre-wrap;min-height:1.2em;border-radius:2px}
.qd-note-t:focus{box-shadow:0 0 0 1px #1a73e8}
.qd-img.qd-imgsel{outline:2px solid #1a73e8;outline-offset:1px}
.qd-imgh{position:fixed;width:11px;height:11px;background:#fff;border:2px solid #1a73e8;border-radius:2px;cursor:nwse-resize;z-index:60}
.qd-p[data-list]{counter-increment:none}
.qd-p[data-locked]{background:repeating-linear-gradient(135deg,transparent 0 6px,rgba(127,127,127,.06) 6px 12px);cursor:not-allowed}
.qd-img{max-width:100%;height:auto;vertical-align:bottom}
.qd-pb{display:block;border-top:1px dashed #9aa0a6;margin:10px 0;height:0;position:relative}
.qd-pb::after{content:"分页符";position:absolute;left:50%;top:-8px;transform:translateX(-50%);background:#fff;color:#9aa0a6;font-size:10px;padding:0 6px;font-family:sans-serif}
.qd-body{flex:1;min-height:0;display:flex}
.qd-body>.qd-scroll{min-width:0}
.qd-cm{display:inline}
.qd-cm[data-cm=r]{display:inline-block;width:11px;height:10px;margin:0 1px;vertical-align:super;cursor:pointer;background:no-repeat center/11px 10px url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 11 10'%3E%3Cpath d='M.5.5h10v6.5H4.5L2 9.5V7H.5z' fill='%23ffd54f' stroke='%23b8860b'/%3E%3C/svg%3E")}
::highlight(q-comment){background:rgba(255,213,79,.38)}
::highlight(q-comment-on){background:rgba(255,179,0,.7)}
.qd-p[data-toc]{cursor:pointer;background:none}
.qd-p[data-toc]:hover{color:#1155cc}
.qd-p[data-toc="0"]{cursor:default;font-weight:700;font-size:1.2em}
.qd-p[data-toc="2"]{padding-left:24px}
.qd-p[data-toc="3"]{padding-left:48px}
.qd-p.qd-tocp{display:flex;align-items:baseline}
.qd-tocp>.qd-leader{flex:1;min-width:1.5em;margin:0 .35em;color:transparent;border-bottom:1.5px dotted #777;transform:translateY(-.25em)}
.qd-cpane{width:270px;min-width:220px;max-width:40%;border-left:1px solid var(--q-b1);background:var(--q-l1);overflow-y:auto;display:flex;flex-direction:column;gap:8px;padding:8px;font-size:12px;color:var(--q-t1)}
.qd-cpane[hidden]{display:none}
.qd-chead{display:flex;align-items:center;gap:6px}
.qd-chead b{white-space:nowrap}
.qd-cwho{flex:1;min-width:0;height:24px;border:1px solid var(--q-b2);border-radius:5px;padding:0 6px;font:inherit;background:var(--q-bg);color:var(--q-t1);outline:none}
.qd-card{background:var(--q-bg);border:1px solid var(--q-b1);border-left:3px solid #ffd54f;border-radius:6px;padding:7px 8px;display:flex;flex-direction:column;gap:4px;cursor:pointer}
.qd-card.on{border-color:#f0b400;box-shadow:0 2px 8px rgba(0,0,0,.1)}
.qd-card.done{opacity:.6;border-left-color:var(--q-b2)}
.qd-cmeta{display:flex;gap:6px;align-items:baseline}
.qd-cmeta span{color:var(--q-t3);font-size:11px}
.qd-cmeta em{margin-left:auto;font-style:normal;color:#1a7f37;font-size:11px}
.qd-cquote{color:var(--q-t3);border-left:2px solid var(--q-b2);padding-left:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.qd-ctext{white-space:pre-wrap;word-break:break-word;line-height:1.45}
.qd-creply{border-top:1px solid var(--q-b1);padding-top:5px;margin-top:2px;display:flex;flex-direction:column;gap:3px}
.qd-cacts{display:flex;gap:2px;flex-wrap:wrap}
.qd-cacts button,.qd-cbtns button{border:0;background:transparent;color:var(--q-brand);font:inherit;font-size:11px;padding:2px 5px;border-radius:4px;cursor:pointer}
.qd-cacts button:hover{background:color-mix(in srgb,var(--q-brand) 12%,transparent)}
.qd-ced{display:flex;flex-direction:column;gap:4px}
.qd-cta{width:100%;box-sizing:border-box;border:1px solid var(--q-b2);border-radius:5px;padding:5px 6px;font:inherit;resize:vertical;background:var(--q-bg);color:var(--q-t1);outline:none}
.qd-cta:focus{border-color:var(--q-brand)}
.qd-cbtns{display:flex;justify-content:flex-end;gap:4px}
.qd-cbtns .qd-cok{background:var(--q-brand);color:#fff;padding:3px 12px}
.qd-cnone{color:var(--q-t3);padding:8px 2px;line-height:1.5}
.qd-t{margin:6pt 0}
.qd-t table{border-collapse:collapse;width:100%}
.qd-t td{border:1px dashed #d5d9de;padding:3px 6px;vertical-align:top;white-space:pre-wrap;word-break:break-word;min-width:40px;outline:none}
.qd-t td:focus{box-shadow:inset 0 0 0 2px #1a73e8}
.qd-t td.qd-th{font-weight:700;background:#f2f4f7}
.qd-t td.qd-csel{box-shadow:inset 0 0 0 999px rgba(26,115,232,.16)}
.qd-csel-on .qd-t td{user-select:none;caret-color:transparent}
.qd-t td.qd-rz{cursor:col-resize}
.qd-colguide{position:fixed;width:0;border-left:1px dashed #1a73e8;z-index:50;pointer-events:none}
.qd-t td[contenteditable=false]{color:#555;cursor:not-allowed}

.qp{flex:1;min-height:0;display:flex}
.qp-thumbs{width:172px;min-width:172px;overflow-y:auto;border-right:1px solid var(--q-b1);background:var(--q-l1);padding:8px 6px;display:flex;flex-direction:column;gap:8px}
.qp-thumb{display:flex;gap:6px;align-items:flex-start;border:0;background:transparent;padding:0;cursor:pointer;color:var(--q-t2);font:inherit;font-size:11px}
.qp-num{width:22px;flex:none;white-space:nowrap;font-variant-numeric:tabular-nums;text-align:right;padding-top:2px}
.qp-mini{display:block;overflow:hidden;border:2px solid transparent;border-radius:3px;box-shadow:0 0 0 1px var(--q-b2);position:relative;background:#fff}
.qp-mini .slide{position:absolute;left:0;top:0;pointer-events:none;box-shadow:none}
.qp-thumb.on .qp-mini{border-color:var(--q-brand);box-shadow:none}
.qp-thumb.hidden .qp-mini{opacity:.45}
.qp-thumb.hidden .qp-num{text-decoration:line-through}
.qp-center{flex:1;min-width:0;display:flex;flex-direction:column}
.qp-stage{flex:1;min-height:0;display:flex;background:var(--q-l2);outline:none;overflow:hidden;position:relative;padding:16px}
.qp-stage.qp-zoomed{overflow:auto}
.qp-scaler{position:relative;flex-shrink:0;margin:auto}
.qp .shp.qp-cut{opacity:.45;outline:1px dashed #1a73e8}
.qp td.qp-cell-edit{outline:2px solid #1a73e8;outline-offset:-2px;background:#fff;color:#000;cursor:text;user-select:text;-webkit-user-select:text}
.q-color.q-off{opacity:.38;pointer-events:none}
.qp.qp-show{position:fixed;inset:0;z-index:2147483500;background:#000;display:block;outline:none;cursor:none}
.qp-show>.slide{position:absolute;left:0;top:0;transform-origin:0 0;box-shadow:none}
.qp-show .shp[data-k]:hover{outline:none}
.qp-show-end{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#bbb;font:16px -apple-system,"PingFang SC",sans-serif;cursor:default}
.qp-show-end[hidden]{display:none}
.qp-scaler>.slide{position:absolute;left:0;top:0}
.qp .shp[data-k]{cursor:move}
.qp-tlist{display:flex;flex-direction:column;gap:8px}
/* The host rounds corners as superellipses; slide shapes keep true ellipses. */
.qp .slide,.qp .slide *,.qp-frame *{corner-shape:round}
.qp-mbar{display:flex;align-items:center;justify-content:space-between;gap:6px;padding:2px 2px 8px;border-bottom:1px solid var(--q-b1);font-size:12px;color:var(--q-t1)}
.qp-mbar[hidden]{display:none}
.qp-mbar button{border:1px solid var(--q-b2);background:var(--q-bg);color:var(--q-t1);border-radius:4px;padding:2px 6px;font:inherit;font-size:11px;cursor:pointer}
.qp-thumb.layout{padding-left:22px}
.qp-star{font-size:9px;color:var(--q-t2);align-self:flex-end}
.qp-master .qp-notes{display:none}
.qp-master .qp-stage{box-shadow:inset 0 0 0 2px rgba(232,163,61,.5)}
.qp-center{position:relative}
.qp-box{transform-origin:50% 50%}
.qp-outl,.qp-guides,.qp-badges{position:absolute;inset:0;pointer-events:none;z-index:6}
.qp-ol{position:absolute;border:1px dashed #1a73e8;transform-origin:50% 50%}
.qp-rot{position:absolute;left:calc(50% - 6px);top:-26px;width:12px;height:12px;border-radius:50%;background:#fff;border:1.5px solid #1a73e8;pointer-events:auto;cursor:grab;box-sizing:border-box}
.qp-rot::after{content:'';position:absolute;left:3.75px;top:10.5px;width:1.5px;height:13px;background:#1a73e8}
.qp-box.multi .qp-h,.qp-box.multi .qp-rot,.qp-box.fixed .qp-rot,.qp-box.editing .qp-rot,.qp-box.grp .qp-rot,.qp-box.cropping .qp-rot{display:none}
.qp-box.multi{border-style:dotted}
.qp-box.cropping{border-color:#000;box-shadow:0 0 0 9999px rgba(0,0,0,.4)}
.qp-box.cropping .qp-h{background:#000;border-color:#fff;width:11px;height:11px}
.qp .pic.qp-cropping{overflow:visible;z-index:4}
.qp-marq{position:absolute;border:1px solid #1a73e8;background:rgba(26,115,232,.08);z-index:7;pointer-events:none}
.qp-guide{position:absolute;background:#e8467c}
.qp-guide.v{top:0;bottom:0;width:1px}
.qp-guide.h{left:0;right:0;height:1px}
.qp-badge{position:absolute;min-width:16px;height:16px;padding:0 3px;border-radius:3px;background:#e8a33d;color:#fff;font:600 10px/16px -apple-system,sans-serif;text-align:center;box-sizing:border-box}
.qp-apane{width:240px;flex:none;background:var(--q-bg);border-left:1px solid var(--q-b1);overflow:auto;font-size:12px;color:var(--q-t1)}
.qp-apane[hidden]{display:none}
.qp-ahead{display:flex;align-items:center;gap:6px;padding:8px;border-bottom:1px solid var(--q-b1)}
.qp-ahead button,.qp-arow button{border:0;background:transparent;color:var(--q-t2);cursor:pointer;font:inherit}
.qp-arow{display:flex;align-items:center;gap:6px;padding:5px 8px;cursor:pointer;border-bottom:1px solid var(--q-b1)}
.qp-arow.on{background:var(--q-l1);box-shadow:inset 3px 0 var(--q-brand)}
.qp-arow b{width:18px;text-align:center;color:#e8a33d}
.qp-afx{color:var(--q-t2);width:36px}
.qp-aname{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.qp-anote{padding:10px;color:var(--q-t2);margin:0;line-height:1.5}
.qp-tprev{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden;z-index:9;pointer-events:none}
.qp-tprev>.slide{position:absolute;left:0;top:0}
.qp-fx{display:inline-block;width:18px;font-size:15px;text-align:center}
.qp-fx-shadow{text-shadow:2px 2px 2px rgba(0,0,0,.45)}
.qp-fx-outline{color:#fff;-webkit-text-stroke:1px #4472C4}
.qp-fx-glow{text-shadow:0 0 4px #ffc000,0 0 6px #ffc000}
.qp-fx-gradient{background:linear-gradient(#9DC3E6,#1F4E79);-webkit-background-clip:text;background-clip:text;color:transparent}
.qp-fx-fire{background:linear-gradient(#FFC000,#C00000);-webkit-background-clip:text;background-clip:text;color:transparent}
.qp-fx-neon{color:#fff;text-shadow:0 0 3px #0cf,0 0 6px #0cf}
.qp-fontpre{display:inline-block;min-width:44px}
.qp.qp-show.qp-presenter{display:flex;cursor:default}
.qp-screen{position:absolute;inset:0;overflow:hidden}
.qp-presenter .qp-screen{position:relative;flex:1}
.qp-solo-back{display:none}
.qp-presenter.qp-solo .qp-pv{display:none}
.qp-presenter.qp-solo .qp-solo-back{display:block;position:absolute;right:14px;bottom:14px;z-index:2;opacity:.25;background:#333;color:#eee;border:1px solid #555;border-radius:4px;padding:5px 10px;font:12px -apple-system,"PingFang SC",sans-serif;cursor:pointer}
.qp-presenter.qp-solo .qp-solo-back:hover{opacity:1}
.qp-frame{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden}
.qp-frame>.slide{position:absolute;left:0;top:0;box-shadow:none}
.qp-pv{width:320px;flex:none;background:#1e1e1e;color:#ddd;display:flex;flex-direction:column;gap:8px;padding:12px;font:13px -apple-system,"PingFang SC",sans-serif;overflow:auto;box-sizing:border-box}
.qp-pv-top{display:flex;align-items:center;gap:6px}
.qp-pv-time{font-size:24px;font-variant-numeric:tabular-nums;color:#fff}
.qp-pv button{background:#333;color:#ddd;border:1px solid #444;border-radius:4px;padding:3px 8px;cursor:pointer;font:inherit}
.qp-pv-h{color:#999;font-size:11px;margin-top:4px}
.qp-pv-next{position:relative;height:160px;flex:none;background:#000;border-radius:3px;overflow:hidden;color:#888;display:flex;align-items:center;justify-content:center}
.qp-pv-notes{flex:1;min-height:80px;overflow:auto;white-space:pre-wrap;line-height:1.6;font-size:15px;color:#eee}
.qp-pv-btns{display:flex;flex-wrap:wrap;gap:6px}
.qp .shp[data-k]:hover{outline:1px dashed rgba(26,115,232,.6)}
.qp .txt.qp-editing{cursor:text;outline:none;user-select:text;-webkit-user-select:text;z-index:5}
.qp .txt[data-prompt]:not(.qp-editing):not([data-filled])::before{content:attr(data-prompt);color:#8b8b8b;font-size:var(--qfs,24px);position:absolute;left:0;right:0;top:50%;transform:translateY(-50%);text-align:center;pointer-events:none}
.qp .txt[data-prompt]:not([data-filled]){border:1px dashed rgba(0,0,0,.25)}
.qp-box{position:absolute;border:1.5px solid #1a73e8;pointer-events:none;z-index:6}
.qp-box.editing{border-style:dashed}
.qp-h{position:absolute;width:9px;height:9px;background:#fff;border:1.5px solid #1a73e8;border-radius:2px;pointer-events:auto}
.qp-box.fixed .qp-h,.qp-box.editing .qp-h{display:none}
.qp-nw{left:-5px;top:-5px;cursor:nwse-resize}.qp-box .qp-n{left:calc(50% - 5px);top:-5px;cursor:ns-resize}.qp-ne{right:-5px;top:-5px;cursor:nesw-resize}
.qp-e{right:-5px;top:calc(50% - 5px);cursor:ew-resize}.qp-se{right:-5px;bottom:-5px;cursor:nwse-resize}.qp-s{left:calc(50% - 5px);bottom:-5px;cursor:ns-resize}
.qp-sw{left:-5px;bottom:-5px;cursor:nesw-resize}.qp-w{left:-5px;top:calc(50% - 5px);cursor:ew-resize}
.qp-notes[hidden]{display:none}
.qp-notes{height:84px;min-height:60px;border:0;border-top:1px solid var(--q-b1);resize:vertical;padding:8px 12px;font:13px/1.5 -apple-system,"PingFang SC",sans-serif;background:var(--q-bg);color:var(--q-t1);outline:none}
.qp-missing{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#888}
/* ---- look: each editor wears its Office colour; neutral tints come from
   the host's own text and background, so light and dark themes both work */
.quire{--q-acc:#2b579a;--q-tint1:color-mix(in srgb,var(--q-t1) 3.5%,var(--q-bg));--q-tint2:color-mix(in srgb,var(--q-t1) 7%,var(--q-bg));--q-tint3:color-mix(in srgb,var(--q-t1) 11%,var(--q-bg));--q-acc-soft:color-mix(in srgb,var(--q-acc) 13%,var(--q-bg));--q-acc-line:color-mix(in srgb,var(--q-acc) 45%,var(--q-bg))}
.quire:has(.qx){--q-acc:#217346}
.quire:has(.qp){--q-acc:#c43e1c}
.q-head{background:var(--q-tint1);padding:0 6px;min-height:34px;border-bottom:1px solid var(--q-b1)}
.q-tabs{flex-wrap:nowrap;overflow-x:auto;overflow-y:hidden;scrollbar-width:none;align-items:stretch;gap:2px}
.q-tabs::-webkit-scrollbar{display:none}
.q-tab{position:relative;padding:0 10px;border-radius:0;font-size:13px;color:var(--q-t2);display:inline-flex;align-items:center}
.q-tab:hover{background:transparent;color:var(--q-t1)}
.q-tab::after{content:'';position:absolute;left:10px;right:10px;bottom:0;height:2px;border-radius:2px 2px 0 0;background:transparent;transition:background .12s}
.q-tab:hover::after{background:var(--q-b2)}
.q-tab.on{background:transparent;box-shadow:none;color:var(--q-t1);font-weight:600}
.q-tab.on::after{background:var(--q-acc)}
.q-tab[data-file]{color:var(--q-acc);font-weight:600}
.q-bar{padding:4px 8px;min-height:40px;background:var(--q-bg);box-shadow:0 1px 2px color-mix(in srgb,var(--q-t1) 6%,transparent);position:relative;z-index:1}
.q-panel{gap:2px 3px}
.q-btn{height:28px;min-width:28px;padding:0 6px;border-radius:6px;color:var(--q-t1);transition:background .1s}
.q-btn svg{width:16px;height:16px;flex:none}
.q-btn:hover:not(:disabled){background:var(--q-tint2)}
.q-btn:active:not(:disabled){background:var(--q-tint3)}
.q-btn.on{background:var(--q-acc-soft);color:var(--q-acc);border-color:var(--q-acc-line)}
.q-btn .q-lbl{font-size:12.5px;margin-left:5px}
.q-side .q-btn{min-width:28px}
.q-btn.q-more{min-width:14px;width:14px;padding:0;margin-left:-4px;font-size:10px;color:var(--q-t2)}
.q-sel{height:28px;border-radius:6px;border-color:var(--q-b1);padding:0 3px;font-size:12.5px;transition:border-color .1s}
.q-sel:hover:not(:disabled){border-color:var(--q-b2)}
.q-sel:focus-visible{outline:2px solid var(--q-acc-line);outline-offset:-1px}
.q-sep{height:20px;margin:0 5px;background:var(--q-b1)}
.q-color .q-sw{bottom:4px;width:14px}
.q-status{color:var(--q-t3)}
.q-status[data-state=saved]::before{content:'';display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--q-acc);margin-right:5px;vertical-align:1px;opacity:.75}
.q-menu{border-radius:9px;padding:5px;border-color:var(--dsw-alias-border-l1,#e3e5e8);box-shadow:0 12px 32px rgba(0,0,0,.16),0 2px 6px rgba(0,0,0,.06)}
.q-menu>button{padding:6px 10px;border-radius:6px}
.q-menu>button:hover:not(:disabled){background:color-mix(in srgb,var(--dsw-alias-label-primary,#1f2328) 7%,transparent);color:inherit}
.q-dlg{border-radius:12px}
.q-dlg-f .q-dlg-ok{background:var(--q-acc,#1a73e8);border-color:var(--q-acc,#1a73e8)}
.q-dlg-row input:not([type=checkbox]):focus,.q-dlg-row select:focus,.q-dlg-row textarea:focus{outline:2px solid color-mix(in srgb,var(--q-acc,#1a73e8) 45%,transparent);outline-offset:-1px;border-color:transparent}
/* status bars and work surfaces */
.qx-foot,.qd-status,.qp-status{height:28px;min-height:28px;background:var(--q-tint1);border-top:1px solid var(--q-b1);color:var(--q-t2);font-size:12px}
.q-zoom{color:var(--q-t2)}
.q-zoom-r{accent-color:var(--q-acc);width:96px}
.qx-tab{color:var(--q-t2);padding:0 16px;transition:background .1s}
.qx-tab:hover{background:var(--q-tint2)}
.qx-tab.on{background:var(--q-bg);color:var(--q-acc);box-shadow:inset 0 -2px 0 var(--q-acc)}
.qx-fbar{height:30px;background:var(--q-bg)}
.qx-name{width:92px;color:var(--q-t1)}
.qx-fxl{color:var(--q-t3)}
.qd-scroll{background:var(--q-tint2)}
.qd-pg{box-shadow:0 1px 3px rgba(0,0,0,.12),0 4px 14px rgba(0,0,0,.06)}
.qp-thumbs{background:var(--q-tint1);gap:10px;padding:10px 8px}
.qp-num{color:var(--q-t3)}
.qp-thumb:hover .qp-mini{box-shadow:0 0 0 1px var(--q-t3)}
.qp-thumb.on .qp-num{color:var(--q-acc);font-weight:600}
.qp-thumb.on .qp-mini{border-color:var(--q-acc);box-shadow:0 2px 8px color-mix(in srgb,var(--q-acc) 25%,transparent)}
.qp-stage{background:var(--q-tint2)}
.qp .slide{box-shadow:0 1px 3px rgba(0,0,0,.12),0 6px 20px rgba(0,0,0,.08)}
.qp-mini .slide{box-shadow:none}
.qp-notes{background:var(--q-bg);border-top-color:var(--q-b1)}
.qp-mbar{color:var(--q-acc)}
.qp-show>.slide,.qp-frame>.slide,.qp-mini .slide,.qp-tprev .slide,.qp-pv-next .slide{box-shadow:none}
.q-loading{display:flex;flex-direction:column;align-items:center;gap:12px;font-size:12.5px}
.q-loading::before{content:'';width:22px;height:22px;border-radius:50%;border:2.5px solid var(--q-tint3);border-top-color:var(--q-acc);animation:q-spin .8s linear infinite}
@keyframes q-spin{to{transform:rotate(360deg)}}
/* pins: notes for the Agent, gathered in the chat composer */
.q-btn.q-pin{height:26px;padding:0 9px 0 7px;margin-right:1px;border:1px solid color-mix(in srgb,#e8710a 45%,transparent);border-radius:13px;background:color-mix(in srgb,#e8710a 10%,var(--q-bg));color:#c25e00;font-weight:600}
.q-btn.q-pin .q-lbl{margin-left:4px}
.q-btn.q-pin:hover:not(:disabled){background:color-mix(in srgb,#e8710a 18%,var(--q-bg));border-color:#e8710a}
.q-btn.q-pin.on{background:#e8710a;border-color:#e8710a;color:#fff}
.q-btn.q-pin.arm,.q-btn.q-pin.arm:hover:not(:disabled){background:#e8710a;border-color:#e8710a;color:#fff;box-shadow:0 0 0 3px color-mix(in srgb,#e8710a 28%,transparent)}
.q-pinning .q-main{outline:2px dashed color-mix(in srgb,#e8710a 70%,transparent);outline-offset:-2px}
.q-pinning .q-main,.q-pinning .q-main *{cursor:crosshair!important}
.q-pintip{position:absolute;z-index:31;left:50%;transform:translateX(-50%);padding:5px 12px;border-radius:14px;background:#e8710a;color:#fff;font-size:12px;white-space:nowrap;pointer-events:none;box-shadow:0 4px 14px rgba(0,0,0,.18)}
.q-pinmarq{position:absolute;z-index:29;border:1.5px dashed #e8710a;background:rgba(232,113,10,.10);border-radius:2px;pointer-events:none}
.q-pinmarq[hidden]{display:none}
.q-btn.q-pin.on:hover:not(:disabled){background:#d4650a}
.q-btn.q-pin.on+.q-pin-more{color:#c25e00}
.q-pin-n{min-width:16px;height:16px;padding:0 4px;margin-left:5px;border-radius:8px;background:#fff;color:#c25e00;font:700 10.5px/16px -apple-system,BlinkMacSystemFont,sans-serif;text-align:center}
.q-pin-n[hidden]{display:none}
.q-pinpop{position:absolute;z-index:30;width:340px;max-width:calc(100% - 16px);background:var(--q-bg);border:1px solid var(--q-b1);border-radius:10px;box-shadow:0 12px 32px rgba(0,0,0,.18),0 2px 6px rgba(0,0,0,.06);padding:10px;display:flex;flex-direction:column;gap:8px;font-size:12.5px;color:var(--q-t1)}
.q-pin-h{display:flex;gap:8px;align-items:baseline;min-width:0}
.q-pin-shot{flex:none;align-self:center;max-width:72px;max-height:48px;border:1px solid var(--q-b1);border-radius:4px;background:#fff;object-fit:contain}
.q-pin-h b{flex:none;color:#c25e00;font-weight:600}
.q-pin-h b::before{content:'';display:inline-block;width:7px;height:7px;border-radius:50%;background:#e8710a;margin-right:6px;vertical-align:1px}
.q-pin-q{color:var(--q-t2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
.q-pin-in{resize:vertical;min-height:60px;max-height:200px;font:inherit;line-height:1.5;padding:6px 8px;border:1px solid var(--q-b1);border-radius:7px;background:var(--q-bg);color:var(--q-t1)}
.q-pin-in:focus{outline:2px solid color-mix(in srgb,#e8710a 45%,transparent);outline-offset:-1px;border-color:transparent}
.q-pin-f{display:flex;align-items:center;gap:6px}
.q-pin-k{color:var(--q-t3);font-size:11.5px}
.q-pin-f .q-btn{height:28px;padding:0 12px;border:1px solid var(--q-b1)}
.q-pin-f .q-pin-ok{background:#e8710a;border-color:#e8710a;color:#fff}
.q-pin-f .q-pin-ok:hover:not(:disabled){background:#d4650a}
.q-pin-f .q-pin-ok:disabled{opacity:.45}
::highlight(q-pin){background-color:rgba(232,113,10,.15);text-decoration:underline 2px #e8710a}
::highlight(q-pin-on){background-color:rgba(232,113,10,.3)}
.qp-pins{position:absolute;inset:0;pointer-events:none;z-index:4}
.qp-pin{position:absolute;border:1.5px dashed #e8710a;border-radius:4px;background:rgba(232,113,10,.05)}
.qp-pin::before{content:'';position:absolute;left:-5px;top:-5px;width:9px;height:9px;border-radius:50%;background:#e8710a;box-shadow:0 0 0 2px #fff}
.qp-pin-all{inset:0;border-width:2px}
.qp-thumb.q-pinned .qp-num::after{content:'';display:block;width:6px;height:6px;margin:4px 0 0 auto;border-radius:50%;background:#e8710a}
/* document navigation pane */
.qd-nav{width:210px;flex:none;display:flex;flex-direction:column;min-height:0;border-right:1px solid var(--q-b1);background:var(--q-tint1)}
.qd-nav[hidden]{display:none}
.qd-nav-h{display:flex;align-items:center;gap:6px;height:34px;flex:none;padding:0 6px 0 12px;font-weight:600;font-size:12.5px;color:var(--q-t1)}
.qd-nav-n{color:var(--q-t3);font-weight:400;font-size:11.5px}
.qd-nav-x{margin-left:auto;width:24px;height:24px;border:0;border-radius:6px;background:transparent;color:var(--q-t3);cursor:pointer;font-size:11px}
.qd-nav-x:hover{background:var(--q-tint3);color:var(--q-t1)}
.qd-nav-l{flex:1;min-height:0;overflow-y:auto;padding:0 6px 12px;scrollbar-width:thin}
.qd-nav-i{position:relative;display:flex;align-items:center;gap:2px;height:28px;padding-right:8px;border-radius:6px;cursor:pointer;color:var(--q-t2);font-size:12.5px;white-space:nowrap;user-select:none}
.qd-nav-i[hidden]{display:none}
.qd-nav-i[data-lv="1"]{color:var(--q-t1);font-weight:600}
.qd-nav-i:hover{background:var(--q-tint2);color:var(--q-t1)}
.qd-nav-i.on{background:var(--q-acc-soft);color:var(--q-acc)}
.qd-nav-i.on::before{content:'';position:absolute;left:0;top:7px;bottom:7px;width:3px;border-radius:2px;background:var(--q-acc)}
.qd-nav-t{overflow:hidden;text-overflow:ellipsis;min-width:0}
.qd-nav-tw{width:18px;height:18px;flex:none;display:inline-flex;align-items:center;justify-content:center;border-radius:4px;color:var(--q-t3)}
.qd-nav-tw::before{content:'';border-left:4.5px solid currentColor;border-top:3.5px solid transparent;border-bottom:3.5px solid transparent;transform:rotate(90deg);transition:transform .12s}
.qd-nav-i.shut .qd-nav-tw::before{transform:none}
.qd-nav-tw.none::before{display:none}
.qd-nav-tw:not(.none):hover{background:var(--q-tint3);color:var(--q-t1)}
.qd-nav-e{margin:8px 8px;color:var(--q-t3);font-size:12px;line-height:1.7}
.q-sbtn{display:inline-flex;align-items:center;gap:4px;height:22px;padding:0 7px 0 5px;margin-right:6px;border:0;border-radius:5px;background:transparent;color:var(--q-t2);font:inherit;font-size:12px;cursor:pointer}
.q-sbtn svg{width:14px;height:14px}
.q-sbtn:hover{background:var(--q-tint3);color:var(--q-t1)}
.q-sbtn.on{color:var(--q-acc)}
.qpd{display:flex;flex-direction:column;gap:6px;margin:0 0 8px;font:13px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;color:var(--dsw-alias-label-primary,#1f2328)}
.qpd-list{display:flex;flex-direction:column;gap:6px;max-height:260px;overflow-y:auto}
.qpd-list[hidden]{display:none}
.qpd-card{display:flex;align-items:flex-start;gap:10px;padding:9px 10px;border:1px solid var(--dsw-alias-border-l2,#e3e5e8);border-radius:12px;background:var(--dsw-alias-bg-base,#fff);cursor:pointer}
.qpd-card:hover{border-color:color-mix(in srgb,#e8710a 45%,transparent)}
.qpd-card.on{cursor:default;border-color:#e8710a}
.qpd-ico{flex:none;display:flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:8px;background:color-mix(in srgb,#e8710a 12%,transparent);color:#c25e00}
.qpd-ico svg{width:18px;height:18px}
.qpd-thumb{flex:none;width:56px;height:42px;border:1px solid var(--dsw-alias-border-l2,#e3e5e8);border-radius:6px;background:#fff;object-fit:contain}
.qpd-body{flex:1;min-width:0}
.qpd-where{font-size:12px;color:var(--dsw-alias-label-secondary,#59636e);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.qpd-where b{font-weight:600;color:var(--dsw-alias-label-primary,#1f2328)}
.qpd-q{color:var(--dsw-alias-label-tertiary,#8b949e)}
.qpd-note{margin-top:2px;white-space:pre-wrap;word-break:break-word;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.qpd-in{display:block;width:100%;margin-top:4px;padding:5px 7px;border:1px solid var(--dsw-alias-border-l2,#cfd3d8);border-radius:6px;background:transparent;color:inherit;font:inherit;resize:vertical;outline:none}
.qpd-in:focus{border-color:#e8710a}
.qpd-acts{flex:none;display:flex;gap:2px;opacity:.75}
.qpd-card:hover .qpd-acts{opacity:1}
.qpd-btn{display:flex;align-items:center;justify-content:center;width:24px;height:24px;border:0;border-radius:6px;background:transparent;color:var(--dsw-alias-label-secondary,#59636e);font-family:inherit;font-size:16px;line-height:1;cursor:pointer}
.qpd-btn:hover{background:var(--dsw-alias-interactive-bg-hover,rgba(0,0,0,.06));color:var(--dsw-alias-label-primary,#1f2328)}
.qpd-foot{display:flex;align-items:center;gap:8px;min-height:26px}
.qpd-pill{display:inline-flex;align-items:center;gap:2px;height:26px;padding:0 4px 0 2px;border:1px solid var(--dsw-alias-border-l2,#e3e5e8);border-radius:13px;background:var(--dsw-alias-bg-base,#fff);font-size:12.5px}
.qpd-tog{display:inline-flex;align-items:center;gap:4px;height:22px;padding:0 6px 0 7px;border:0;border-radius:11px;background:transparent;color:inherit;font-family:inherit;font-size:12.5px;line-height:1;cursor:pointer}
.qpd-tog:hover,.qpd.open .qpd-tog{background:var(--dsw-alias-interactive-bg-hover,rgba(0,0,0,.06))}
.qpd-pi{display:flex;color:#c25e00}.qpd-pi svg{width:14px;height:14px}
.qpd-chev{display:flex;color:var(--dsw-alias-label-tertiary,#8b949e);transition:transform .15s}
.qpd.open .qpd-chev{transform:rotate(180deg)}
.qpd-x{display:flex;align-items:center;justify-content:center;width:18px;height:18px;border:0;border-radius:9px;background:transparent;color:var(--dsw-alias-label-tertiary,#8b949e);font-family:inherit;font-size:14px;line-height:1;cursor:pointer}
.qpd-x:hover{background:var(--dsw-alias-interactive-bg-hover,rgba(0,0,0,.06));color:var(--dsw-alias-label-primary,#1f2328)}
.qpd-hint{flex:1;min-width:0;font-size:12px;color:var(--dsw-alias-label-tertiary,#8b949e);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.qpd-send{height:26px;padding:0 11px;border:0;border-radius:13px;background:#e8710a;color:#fff;font-family:inherit;font-weight:600;font-size:12.5px;line-height:1;cursor:pointer}
.qpd-send:hover{background:#d4650a}
`

    let styled = false
    function injectStyles () {
      if (styled || document.getElementById('quire-editor-css')) { styled = true; return }
      document.head.append(el('style', { id: 'quire-editor-css' }, CSS))
      styled = true
    }

    const inject = ['slots', 'documentPreviews']
    function apply (ctx) {
      injectStyles()
      hostCtx = ctx
      ctx.effect(() => ctx.documentPreviews.register({
        id: ID,
        extensions: ['docx', 'xlsx', 'pptx'],
        binaryExtensions: ['docx', 'xlsx', 'pptx'],
        priority: 'extension',
        title: () => 'Quire 编辑',
        loading: 'renderer'
      }), 'quire: document editor')
      ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register({
        name: 'sidebar.right.tab.document',
        key: ID
      }, QuireBody)), 'quire: editor body')
      // Pins waiting for the next message, above the composer.
      ctx.effect(() => ctx.slots.inject('conversation.input.dock', () => ctx.slots.register({
        name: 'conversation.input.dock',
        id: 'quire-pins',
        order: 30,
        inject: (sessionId) => ({ sessionId })
      }, QuirePinDock)), 'quire: pin dock')
    }

    module.exports.apply = apply
    module.exports.inject = inject
    module.exports.name = 'quire'
    return module.exports
  }
})
