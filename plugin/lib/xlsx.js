// Excel workbooks: read sheets as tables, create from rows, edit cells,
// styles and sheets, recalculate formulas, render for the preview pane.
import { Package, NS, REL } from './pkg.js'
import { el, text, doc, child, children, descendants, path as xpath, textContent, remove, clone, insertOrdered, serializeNode, isEl } from './xml.js'
import { coreProps } from './docx.js'
import crypto from 'node:crypto'
import { sheetCharts, addChart, moveChart, deleteChart, editChart, refreshCharts, shiftCharts, sheetPics, addPic, movePic, deletePic } from './chart.js'
import { parseFormula, evaluate, shift, tokenize, colNum, Unsupported, isErr, parseCell, parseRange, cellName, colName, fmtGeneral } from './formula.js'

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const CT = {
  wb: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  ws: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  sst: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
}
const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst']
const WB_ORDER = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']
const STYLE_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']
const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme']
const XF_ORDER = ['alignment', 'protection', 'extLst']

// Strip a namespace prefix: workbooks written by some tools use x:row.
const local = (n) => n.name?.replace(/^\w+:/, '')
const kid = (n, name) => n?.kids?.find((k) => local(k) === name)
const kids = (n, name) => (n?.kids ?? []).filter((k) => local(k) === name)

// ------------------------------------------------------------- template

function newPackage () {
  const p = new Package([])
  p.setXml('[Content_Types].xml', doc(el('Types', { xmlns: NS.ct }, [
    el('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }),
    el('Default', { Extension: 'xml', ContentType: 'application/xml' }),
    el('Override', { PartName: '/xl/workbook.xml', ContentType: CT.wb }),
    el('Override', { PartName: '/xl/styles.xml', ContentType: CT.styles }),
    el('Override', { PartName: '/xl/sharedStrings.xml', ContentType: CT.sst }),
    el('Override', { PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }),
    el('Override', { PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' })
  ])))
  p.setXml('_rels/.rels', doc(el('Relationships', { xmlns: NS.rel }, [
    el('Relationship', { Id: 'rId1', Type: NS.officeDocument, Target: 'xl/workbook.xml' }),
    el('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: 'docProps/core.xml' }),
    el('Relationship', { Id: 'rId3', Type: REL('extended-properties'), Target: 'docProps/app.xml' })
  ])))
  p.setXml('xl/workbook.xml', doc(el('workbook', { xmlns: MAIN, 'xmlns:r': NS.r }, [
    el('bookViews', {}, [el('workbookView', { xWindow: '0', yWindow: '0', windowWidth: '28800', windowHeight: '16000' })]),
    el('sheets'),
    el('calcPr', { calcId: '191029', fullCalcOnLoad: '1' })
  ])))
  p.setXml('xl/_rels/workbook.xml.rels', doc(el('Relationships', { xmlns: NS.rel }, [
    el('Relationship', { Id: 'rId1', Type: REL('styles'), Target: 'styles.xml' }),
    el('Relationship', { Id: 'rId2', Type: REL('sharedStrings'), Target: 'sharedStrings.xml' })
  ])))
  p.setXml('xl/styles.xml', doc(el('styleSheet', { xmlns: MAIN }, [
    el('fonts', { count: '1' }, [el('font', {}, [el('sz', { val: '11' }), el('color', { rgb: 'FF000000' }), el('name', { val: 'Calibri' }), el('family', { val: '2' })])]),
    el('fills', { count: '2' }, [el('fill', {}, [el('patternFill', { patternType: 'none' })]), el('fill', {}, [el('patternFill', { patternType: 'gray125' })])]),
    el('borders', { count: '1' }, [el('border', {}, [el('left'), el('right'), el('top'), el('bottom'), el('diagonal')])]),
    el('cellStyleXfs', { count: '1' }, [el('xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' })]),
    el('cellXfs', { count: '1' }, [el('xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0', xfId: '0' })]),
    el('cellStyles', { count: '1' }, [el('cellStyle', { name: 'Normal', xfId: '0', builtinId: '0' })])
  ])))
  p.setXml('xl/sharedStrings.xml', doc(el('sst', { xmlns: MAIN, count: '0', uniqueCount: '0' })))
  p.setXml('docProps/core.xml', coreProps())
  p.setXml('docProps/app.xml', doc(el('Properties', { xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' }, [el('Application', {}, [text('Quire')])])))
  return p
}

function sheetDoc () {
  return doc(el('worksheet', { xmlns: MAIN, 'xmlns:r': NS.r }, [
    el('dimension', { ref: 'A1' }),
    el('sheetViews', {}, [el('sheetView', { workbookViewId: '0' })]),
    el('sheetFormatPr', { defaultRowHeight: '15' }),
    el('sheetData'),
    el('pageMargins', { left: '0.7', right: '0.7', top: '0.75', bottom: '0.75', header: '0.3', footer: '0.3' })
  ]))
}

// ------------------------------------------------------------- number formats

const BUILTIN_FMT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy-mm-dd', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy-mm-dd h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: 'm-d-yy', 31: 'yyyy"年"m"月"d"日"', 32: 'h"时"mm"分"', 33: 'h"时"mm"分"ss"秒"', 34: '上午/下午h"时"mm"分"', 35: '上午/下午h"时"mm"分"ss"秒"', 36: 'yyyy"年"m"月"',
  50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"', 54: 'm"月"d"日"', 55: '上午/下午h"时"mm"分"', 56: '上午/下午h"时"mm"分"ss"秒"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"'
}
const FMT_ID = { General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '@': 49 }

export function formatValue (v, code) {
  if (v === null || v === undefined) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (isErr(v)) return v.error
  if (typeof v === 'string') return v
  if (!code || code === 'General') return fmtGeneral(+v.toPrecision(11))
  const sections = splitSections(code)
  let sec = sections[0]
  let val = v
  if (v < 0 && sections[1] !== undefined) { sec = sections[1]; val = -v } else if (v === 0 && sections[2] !== undefined) sec = sections[2]
  const bare = sec.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '')
  if (/[yd]|m{1,4}(?![^h]*s)|[hs]/i.test(bare) && !/[#0?]/.test(bare)) return formatDate(val, sec)
  // Literal text and placeholders.
  let pre = ''
  let post = ''
  const cleaned = sec.replace(/\[\$([^\]-]*)[^\]]*\]/g, (_, cur) => `"${cur}"`).replace(/\[[^\]]*\]/g, '').replace(/_./g, ' ').replace(/\*./g, '')
  const m = /^((?:"[^"]*"|\\.|[^#0?.,%E])*)([#0?.,%E+-]*?[#0?](?:[#0?.,%E+-]*[#0?%])?)?(.*)$/s.exec(cleaned)
  if (!m || !m[2]) return cleaned.replace(/"/g, '').replace(/\\/g, '')
  const lit = (s) => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1')
  pre = lit(m[1])
  post = lit(m[3])
  let core = m[2]
  if (post.startsWith('%')) { core += '%'; post = post.slice(1) }
  const pct = (core.match(/%/g) || []).length
  if (pct) val *= Math.pow(100, pct)
  const sci = /E[+-]/i.test(core)
  const dec = (core.split('.')[1] || '').replace(/[^0?#]/g, '').length
  const minDec = (core.split('.')[1] || '').replace(/[^0]/g, '').length
  let s
  if (sci) s = val.toExponential(dec).toUpperCase().replace(/E([+-])(\d)$/, 'E$10$2')
  else {
    s = val.toFixed(dec)
    if (dec > minDec) s = s.replace(new RegExp(`0{1,${dec - minDec}}$`), '').replace(/\.$/, '')
    if (/,/.test(core.split('.')[0])) {
      const [i, f] = s.split('.')
      s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f !== undefined ? '.' + f : '')
    }
    const intZeros = (core.split('.')[0].match(/0/g) || []).length
    if (intZeros === 0 && /^0\./.test(s)) s = s.slice(1)
  }
  return (v < 0 && sections[1] === undefined ? '-' : '') + pre + s + '%'.repeat(pct) + post
}

function splitSections (code) {
  const out = []
  let cur = ''
  let q = false
  for (let i = 0; i < code.length; i++) {
    const c = code[i]
    if (c === '"') q = !q
    if (c === '\\' && !q) { cur += c + code[++i]; continue }
    if (c === ';' && !q) { out.push(cur); cur = ''; continue }
    cur += c
  }
  out.push(cur)
  return out
}

export function serialToDate (n) {
  // Excel's 1900 date system, including its phantom 1900-02-29.
  const ms = Math.round((n - (n >= 60 ? 25569 : 25568)) * 86400000)
  return new Date(ms)
}

function formatDate (v, sec) {
  const d = serialToDate(v)
  const p = (x, n = 2) => String(x).padStart(n, '0')
  const Y = d.getUTCFullYear(); const M = d.getUTCMonth() + 1; const D = d.getUTCDate()
  const h = d.getUTCHours(); const mi = d.getUTCMinutes(); const s = d.getUTCSeconds()
  const bare = sec.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '')
  const hasDate = /[yd]/i.test(bare) || /m/i.test(bare.replace(/h[^m]*m+|m+[^s]*s/gi, ''))
  const hasTime = /[hs]/i.test(bare)
  const parts = []
  if (hasDate) parts.push(`${Y}-${p(M)}-${p(D)}`)
  if (hasTime) parts.push(`${p(h)}:${p(mi)}${/s/i.test(bare) ? ':' + p(s) : ''}`)
  return parts.join(' ') || fmtGeneral(v)
}

// ------------------------------------------------------------- workbook

class Workbook {
  constructor (pkg) {
    this.pkg = pkg
    this.main = pkg.mainPart()
    if (!this.main || !pkg.has(this.main)) throw new Error('not an Excel workbook: no workbook part')
    this.wb = pkg.xml(this.main)
    if (local(this.wb.root) !== 'workbook') throw new Error('not an Excel workbook')
    this.prefix = this.wb.root.name.includes(':') ? this.wb.root.name.split(':')[0] + ':' : ''
    this.sheets = kids(kid(this.wb.root, 'sheets'), 'sheet').map((s) => {
      const rid = s.attrs['r:id'] ?? Object.entries(s.attrs).find(([k]) => k.endsWith(':id'))?.[1]
      const rel = pkg.rel(this.main, rid)
      return { name: s.attrs.name, node: s, rid, part: rel?.part, state: s.attrs.state }
    })
    this.loadStrings()
    this.loadStyles()
    this.cache = new Map()
  }

  // ---------------------------------------------------------- shared strings

  loadStrings () {
    this.sstPart = this.pkg.relsOfType(this.main, 'sharedStrings')[0]?.part
    this.strings = []
    this.stringIndex = new Map()
    if (this.sstPart && this.pkg.has(this.sstPart)) {
      for (const si of kids(this.pkg.xml(this.sstPart).root, 'si')) {
        // Phonetic runs (rPh) are not part of the value.
        const s = si.kids.filter((k) => local(k) === 't' || local(k) === 'r').map((k) => local(k) === 't' ? textContent(k) : textContent(kid(k, 't'))).join('')
        this.stringIndex.set(s, this.stringIndex.has(s) ? this.stringIndex.get(s) : this.strings.length)
        this.strings.push(s)
      }
    }
  }

  stringId (s) {
    if (this.stringIndex.has(s)) return this.stringIndex.get(s)
    if (!this.sstPart || !this.pkg.has(this.sstPart)) {
      this.sstPart = 'xl/sharedStrings.xml'
      this.pkg.setXml(this.sstPart, doc(el('sst', { xmlns: MAIN, count: '0', uniqueCount: '0' })))
      this.pkg.setOverride(this.sstPart, CT.sst)
      this.pkg.addRel(this.main, REL('sharedStrings'), this.sstPart)
    }
    const d = this.pkg.xml(this.sstPart)
    const pre = d.root.name.includes(':') ? d.root.name.split(':')[0] + ':' : ''
    const attrs = /^\s|\s$|\n/.test(s) ? { 'xml:space': 'preserve' } : {}
    d.root.kids.push(el(pre + 'si', {}, [el(pre + 't', attrs, [text(s)])]))
    const id = this.strings.length
    this.strings.push(s)
    this.stringIndex.set(s, id)
    d.root.attrs.uniqueCount = String(this.strings.length)
    d.root.attrs.count = String(Math.max(+(d.root.attrs.count ?? 0) + 1, this.strings.length))
    this.pkg.touch(this.sstPart)
    return id
  }

  // ---------------------------------------------------------- styles

  loadStyles () {
    this.stylesPart = this.pkg.relsOfType(this.main, 'styles')[0]?.part
    if (!this.stylesPart || !this.pkg.has(this.stylesPart)) {
      this.stylesPart = 'xl/styles.xml'
      this.pkg.setXml(this.stylesPart, newPackage().xml('xl/styles.xml'))
      this.pkg.setOverride(this.stylesPart, CT.styles)
      this.pkg.addRel(this.main, REL('styles'), this.stylesPart)
    }
    const r = this.pkg.xml(this.stylesPart).root
    this.st = r
    this.sp = r.name.includes(':') ? r.name.split(':')[0] + ':' : ''
    this.numFmts = new Map(kids(kid(r, 'numFmts'), 'numFmt').map((n) => [+n.attrs.numFmtId, n.attrs.formatCode]))
    this.xfs = kids(kid(r, 'cellXfs'), 'xf')
    this.fonts = kids(kid(r, 'fonts'), 'font')
    this.fills = kids(kid(r, 'fills'), 'fill')
    this.borders = kids(kid(r, 'borders'), 'border')
    this.themeColors = this.loadTheme()
  }

  loadTheme () {
    const part = this.pkg.relsOfType(this.main, 'theme')[0]?.part
    if (!part || !this.pkg.has(part)) return []
    const scheme = descendants(this.pkg.xml(part).root, 'a:clrScheme')[0]
    if (!scheme) return []
    const order = ['a:lt1', 'a:dk1', 'a:lt2', 'a:dk2', 'a:accent1', 'a:accent2', 'a:accent3', 'a:accent4', 'a:accent5', 'a:accent6', 'a:hlink', 'a:folHlink']
    return order.map((n) => {
      const c = child(scheme, n)
      const s = c && (child(c, 'a:srgbClr')?.attrs.val ?? child(c, 'a:sysClr')?.attrs.lastClr)
      return s
    })
  }

  fmtCode (xfIndex) {
    const xf = this.xfs[xfIndex ?? 0]
    const id = +(xf?.attrs.numFmtId ?? 0)
    return this.numFmts.get(id) ?? BUILTIN_FMT[id] ?? 'General'
  }

  color (c) {
    if (!c) return undefined
    if (c.attrs.rgb) return c.attrs.rgb.slice(-6)
    if (c.attrs.theme !== undefined) return this.themeColors[+c.attrs.theme]
    const idx = { 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00', 12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF' }
    return idx[c.attrs.indexed]
  }

  // css describes a cell style for the preview.
  css (xfIndex) {
    const xf = this.xfs[xfIndex ?? 0]
    if (!xf) return {}
    const font = this.fonts[+(xf.attrs.fontId ?? 0)]
    const fill = this.fills[+(xf.attrs.fillId ?? 0)]
    const al = kid(xf, 'alignment')
    const pf = kid(fill, 'patternFill')
    const out = {
      b: font && !!kid(font, 'b') && kid(font, 'b').attrs.val !== '0',
      i: font && !!kid(font, 'i') && kid(font, 'i').attrs.val !== '0',
      u: font && !!kid(font, 'u'),
      color: font && this.color(kid(font, 'color')),
      size: font && kid(font, 'sz')?.attrs.val,
      fill: pf && pf.attrs.patternType && pf.attrs.patternType !== 'none' && pf.attrs.patternType !== 'gray125' ? this.color(kid(pf, 'fgColor')) : undefined,
      align: al?.attrs.horizontal,
      wrap: al?.attrs.wrapText === '1',
      border: !!this.borders[+(xf.attrs.borderId ?? 0)]?.kids?.some((s) => s.attrs?.style)
    }
    if (out.border) {
      // Each side for the editor: [style, color].
      const bd = {}
      const b = this.borders[+(xf.attrs.borderId ?? 0)]
      for (const [k, n] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) {
        const s = kid(b, n)
        if (s?.attrs.style) bd[k] = [s.attrs.style, this.color(kid(s, 'color')) ?? '000000']
      }
      out.bd = bd
    }
    return out
  }

  // deriveSides returns a cellXfs index: base with some border sides set
  // (a style name) or cleared (null); sides not named stay as they are.
  deriveSides (base, sides, color = '000000') {
    const sp = this.sp
    const order = ['left', 'right', 'top', 'bottom', 'diagonal']
    const xf = clone(this.xfs[base ?? 0] ?? this.xfs[0])
    xf.attrs.xfId = '0'
    const old = this.borders[+(xf.attrs.borderId ?? 0)]
    const b = old ? clone(old) : el(sp + 'border')
    for (const n of order) if (!kid(b, n)) b.kids.push(el(sp + n))
    for (const [n, style] of Object.entries(sides)) {
      if (style === undefined) continue
      const at = b.kids.findIndex((k) => local(k) === n)
      b.kids[at] = style ? el(sp + n, { style }, [el(sp + 'color', { rgb: 'FF' + hex(color) })]) : el(sp + n)
    }
    b.kids = b.kids.filter(isEl).sort((x, y) => order.indexOf(local(x)) - order.indexOf(local(y)))
    xf.attrs.borderId = String(this.add('borders', 'border', b, this.borders))
    xf.attrs.applyBorder = '1'
    return this.add('cellXfs', 'xf', xf, this.xfs)
  }

  // moveSheet puts a sheet at a new place (0-based) in the tab order.
  moveSheet (name, to) {
    const s = this.sheet(name).info
    const from = this.sheets.indexOf(s)
    to = Math.max(0, Math.min(this.sheets.length - 1, Math.round(+to)))
    if (from === to) return
    const order = this.sheets.slice()
    order.splice(from, 1)
    order.splice(to, 0, s)
    const remap = new Map(this.sheets.map((x, i) => [i, order.indexOf(x)]))
    const sheetsEl = kid(this.wb.root, 'sheets')
    sheetsEl.kids = [...sheetsEl.kids.filter((k) => local(k) !== 'sheet'), ...order.map((x) => x.node)]
    this.sheets.splice(0, this.sheets.length, ...order)
    // Sheet-scoped names follow their sheet.
    for (const d of descendants(this.wb.root, (n) => local(n) === 'definedName')) {
      if (d.attrs.localSheetId !== undefined) d.attrs.localSheetId = String(remap.get(+d.attrs.localSheetId) ?? d.attrs.localSheetId)
    }
    this.resetView()
  }

  // showSheet hides or shows a sheet; one sheet always stays visible and
  // selected.
  showSheet (name, visible) {
    const sh = this.sheet(name)
    const s = sh.info
    if (visible) delete s.node.attrs.state
    else {
      if (!this.sheets.some((x) => x !== s && (!x.node.attrs.state || x.node.attrs.state === 'visible'))) throw new Error('a workbook must keep at least one visible sheet')
      s.node.attrs.state = 'hidden'
      const sv = descendants(sh.root, (n) => local(n) === 'sheetView')[0]
      if (sv?.attrs.tabSelected) {
        delete sv.attrs.tabSelected
        this.pkg.touch(s.part)
        const first = this.sheets.find((x) => !x.node.attrs.state || x.node.attrs.state === 'visible')
        const fv = descendants(this.sheet(first.name).root, (n) => local(n) === 'sheetView')[0]
        if (fv) { fv.attrs.tabSelected = '1'; this.pkg.touch(first.part) }
      }
    }
    this.resetView()
  }

  resetView () {
    const view = descendants(this.wb.root, (n) => local(n) === 'workbookView')[0]
    if (view) { delete view.attrs.activeTab; delete view.attrs.firstSheet }
    this.pkg.touch(this.main)
  }

  // derive returns a cellXfs index: base with props applied.
  derive (base, props) {
    const sp = this.sp
    const xf = clone(this.xfs[base ?? 0] ?? this.xfs[0])
    delete xf.attrs.xfId
    xf.attrs.xfId = '0'
    const fontKeys = ['bold', 'italic', 'color', 'size', 'underline']
    if (fontKeys.some((k) => props[k] !== undefined)) {
      const f = clone(this.fonts[+(xf.attrs.fontId ?? 0)] ?? this.fonts[0])
      const toggle = (name, onv) => {
        const cur = kid(f, name)
        if (onv && !cur) insertOrdered(f, el(sp + name), FONT_ORDER.map((n) => sp + n))
        if (!onv && cur) remove(f, cur)
      }
      if (props.bold !== undefined) toggle('b', props.bold)
      if (props.italic !== undefined) toggle('i', props.italic)
      if (props.underline !== undefined) toggle('u', props.underline)
      if (props.color !== undefined) {
        const cur = kid(f, 'color')
        if (cur) remove(f, cur)
        insertOrdered(f, el(sp + 'color', { rgb: 'FF' + hex(props.color) }), FONT_ORDER.map((n) => sp + n))
      }
      if (props.size !== undefined) {
        const cur = kid(f, 'sz')
        if (cur) cur.attrs.val = String(props.size)
        else insertOrdered(f, el(sp + 'sz', { val: String(props.size) }), FONT_ORDER.map((n) => sp + n))
      }
      xf.attrs.fontId = String(this.add('fonts', 'font', f, this.fonts))
      xf.attrs.applyFont = '1'
    }
    if (props.fill !== undefined) {
      const f = props.fill
        ? el(sp + 'fill', {}, [el(sp + 'patternFill', { patternType: 'solid' }, [el(sp + 'fgColor', { rgb: 'FF' + hex(props.fill) }), el(sp + 'bgColor', { indexed: '64' })])])
        : el(sp + 'fill', {}, [el(sp + 'patternFill', { patternType: 'none' })])
      xf.attrs.fillId = String(this.add('fills', 'fill', f, this.fills))
      xf.attrs.applyFill = '1'
    }
    if (props.border !== undefined) {
      const side = (n) => props.border ? el(sp + n, { style: 'thin' }, [el(sp + 'color', { rgb: 'FFBFBFBF' })]) : el(sp + n)
      const b = el(sp + 'border', {}, ['left', 'right', 'top', 'bottom'].map(side).concat([el(sp + 'diagonal')]))
      xf.attrs.borderId = String(this.add('borders', 'border', b, this.borders))
      xf.attrs.applyBorder = '1'
    }
    if (props.format !== undefined) {
      xf.attrs.numFmtId = String(this.numFmtId(String(props.format)))
      xf.attrs.applyNumberFormat = '1'
    }
    if (props.locked !== undefined) {
      xf.kids = xf.kids.filter((k) => !(isEl(k) && local(k) === 'protection'))
      if (!props.locked) insertOrdered(xf, el(sp + 'protection', { locked: '0' }), XF_ORDER.map((n) => sp + n))
      xf.attrs.applyProtection = '1'
    }
    if (props.align !== undefined || props.wrap !== undefined || props.valign !== undefined) {
      let al = kid(xf, 'alignment')
      if (!al) { al = el(sp + 'alignment'); insertOrdered(xf, al, XF_ORDER.map((n) => sp + n)) }
      if (props.align !== undefined) { if (props.align) al.attrs.horizontal = props.align; else delete al.attrs.horizontal }
      if (props.valign !== undefined) al.attrs.vertical = props.valign
      if (props.wrap !== undefined) { if (props.wrap) al.attrs.wrapText = '1'; else delete al.attrs.wrapText }
      xf.attrs.applyAlignment = '1'
    }
    return this.add('cellXfs', 'xf', xf, this.xfs)
  }

  // add appends node to a style collection unless an identical one exists.
  add (coll, name, node, list) {
    const key = serializeNode(node).join('')
    const at = list.findIndex((n) => serializeNode(n).join('') === key)
    if (at >= 0) return at
    let c = kid(this.st, coll)
    if (!c) { c = el(this.sp + coll); insertOrdered(this.st, c, STYLE_ORDER.map((n) => this.sp + n)) }
    c.kids.push(node)
    list.push(node)
    c.attrs.count = String(list.length)
    this.pkg.touch(this.stylesPart)
    return list.length - 1
  }

  numFmtId (code) {
    if (FMT_ID[code] !== undefined) return FMT_ID[code]
    for (const [id, c] of this.numFmts) if (c === code) return id
    const id = Math.max(163, ...this.numFmts.keys()) + 1
    let c = kid(this.st, 'numFmts')
    if (!c) { c = el(this.sp + 'numFmts'); insertOrdered(this.st, c, STYLE_ORDER.map((n) => this.sp + n)) }
    c.kids.push(el(this.sp + 'numFmt', { numFmtId: String(id), formatCode: code }))
    c.attrs.count = String(c.kids.length)
    this.numFmts.set(id, code)
    this.pkg.touch(this.stylesPart)
    return id
  }

  // ---------------------------------------------------------- sheets

  sheet (name) {
    const s = name === undefined || name === null || name === ''
      ? this.sheets[0]
      : this.sheets.find((x) => x.name === name) ?? this.sheets.find((x) => x.name.toLowerCase() === String(name).toLowerCase())
    if (!s) throw new Error(`there is no sheet "${name}"; sheets: ${this.sheets.map((x) => x.name).join(', ')}`)
    if (!this.cache.has(s.part)) this.cache.set(s.part, new Sheet(this, s))
    return this.cache.get(s.part)
  }

  addSheet (name) {
    if (!name || /[\\/?*[\]:]/.test(name) || name.length > 31) throw new Error(`"${name}" is not a valid sheet name (1–31 characters, none of \\ / ? * [ ] :)`)
    if (this.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) throw new Error(`a sheet named "${name}" already exists`)
    const part = this.pkg.uniqueName('xl/worksheets/sheet#.xml')
    const d = sheetDoc()
    if (!this.sheets.length) child(child(d.root, 'sheetViews'), 'sheetView').attrs.tabSelected = '1'
    this.pkg.setXml(part, d)
    this.pkg.setOverride(part, CT.ws)
    const rid = this.pkg.addRel(this.main, REL('worksheet'), part)
    const sheetsEl = kid(this.wb.root, 'sheets')
    const ids = kids(sheetsEl, 'sheet').map((s) => +s.attrs.sheetId)
    const node = el(this.prefix + 'sheet', { name, sheetId: String(Math.max(0, ...ids) + 1), 'r:id': rid })
    if (!this.wb.root.attrs['xmlns:r']) this.wb.root.attrs['xmlns:r'] = NS.r
    sheetsEl.kids.push(node)
    this.pkg.touch(this.main)
    const s = { name, node, rid, part }
    this.sheets.push(s)
    return this.sheet(name)
  }

  renameSheet (from, to) {
    const s = this.sheet(from).info
    if (!to || /[\\/?*[\]:]/.test(to) || to.length > 31) throw new Error(`"${to}" is not a valid sheet name`)
    if (this.sheets.some((x) => x !== s && x.name.toLowerCase() === to.toLowerCase())) throw new Error(`a sheet named "${to}" already exists`)
    const old = s.name
    s.node.attrs.name = to
    s.name = to
    this.pkg.touch(this.main)
    // Formulas and defined names that point at the old name follow it.
    // Like Excel: quote only names that are not plain words, or that read
    // as a cell reference.
    const q = (n) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^R\d*C\d*$/i.test(n) ? n : `'${n.replace(/'/g, "''")}'`)
    const lit = (n) => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const re = new RegExp(`(^|[^\\p{L}\\p{N}_'.])(?:'${lit(old.replace(/'/g, "''"))}'|${lit(old)})!`, 'gu')
    for (const x of this.sheets) {
      const sh = this.sheet(x.name)
      let changed = false
      for (const f of descendants(sh.d.root, (n) => local(n) === 'f')) {
        const t = textContent(f)
        const u = t.replace(re, `$1${q(to)}!`)
        if (u !== t) { f.kids = [text(u)]; changed = true }
      }
      if (changed) this.pkg.touch(sh.part)
    }
    for (const dn of descendants(this.wb.root, (n) => local(n) === 'definedName')) {
      const t = textContent(dn)
      const u = t.replace(re, `$1${q(to)}!`)
      if (u !== t) dn.kids = [text(u)]
    }
  }

  deleteSheet (name) {
    const sh = this.sheet(name)
    if (this.sheets.length === 1) throw new Error('a workbook must keep at least one sheet')
    const s = sh.info
    const idx = this.sheets.indexOf(s)
    remove(kid(this.wb.root, 'sheets'), s.node)
    this.pkg.removeRel(this.main, s.rid)
    for (const r of this.pkg.rels(s.part)) if (r.part && /drawing|comments|vmlDrawing|table/.test(r.type) && this.pkg.has(r.part)) { this.pkg.delete(r.part); this.pkg.removeOverride(r.part) }
    this.pkg.delete(s.part)
    this.pkg.delete(this.pkg.relsName(s.part))
    this.pkg.removeOverride(s.part)
    this.sheets.splice(idx, 1)
    this.cache.delete(s.part)
    // Sheet-scoped names shift with the sheet indexes.
    const dns = kid(this.wb.root, 'definedNames')
    if (dns) {
      dns.kids = dns.kids.filter((d) => d.attrs?.localSheetId === undefined || +d.attrs.localSheetId !== idx)
      for (const d of dns.kids) if (d.attrs?.localSheetId !== undefined && +d.attrs.localSheetId > idx) d.attrs.localSheetId = String(+d.attrs.localSheetId - 1)
      if (!dns.kids.length) remove(this.wb.root, dns)
    }
    const view = descendants(this.wb.root, (n) => local(n) === 'workbookView')[0]
    if (view) { delete view.attrs.activeTab; delete view.attrs.firstSheet }
    this.pkg.touch(this.main)
  }

  // ---------------------------------------------------------- defined names

  // names lists the workbook's defined names; sheet is the index of the
  // sheet a local name belongs to.
  names () {
    return kids(kid(this.wb.root, 'definedNames'), 'definedName').map((d) => ({
      name: d.attrs.name,
      ref: textContent(d),
      sheet: d.attrs.localSheetId === undefined ? undefined : +d.attrs.localSheetId,
      hidden: d.attrs.hidden === '1' || d.attrs.hidden === 'true',
      comment: d.attrs.comment
    }))
  }

  // setName adds or replaces a defined name; ref null deletes it.
  setName (name, ref, sheet, comment) {
    if (!/^[A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*$/.test(name ?? '') || /^[A-Za-z]{1,3}\d+$/.test(name) || /^(true|false)$/i.test(name)) throw new Error(`"${name}" 不能作为名称`)
    let dns = kid(this.wb.root, 'definedNames')
    const same = (d) => d.attrs.name.toLowerCase() === name.toLowerCase() && (d.attrs.localSheetId === undefined ? sheet === undefined : +d.attrs.localSheetId === sheet)
    const old = dns?.kids.find((d) => isEl(d) && same(d))
    if (ref === null) {
      if (old) remove(dns, old)
      if (dns && !kids(dns, 'definedName').length) remove(this.wb.root, dns)
    } else {
      const r = String(ref).replace(/^=/, '')
      if (!r) throw new Error('名称需要引用位置')
      if (!dns) { dns = el(this.prefix + 'definedNames'); insertOrdered(this.wb.root, dns, WB_ORDER.map((n) => this.prefix + n)) }
      const attrs = { name, ...(sheet === undefined ? {} : { localSheetId: String(sheet) }), ...(comment ? { comment } : {}) }
      const node = el(this.prefix + 'definedName', attrs, [text(r)])
      if (old) dns.kids[dns.kids.indexOf(old)] = node
      else dns.kids.push(node)
    }
    this.pkg.touch(this.main)
  }

  // ---------------------------------------------------------- recalculation

  // recalc evaluates every formula and stores its cached value. Cells in
  // wrote whose formula cannot be evaluated lose any stale cached value.
  recalc (wrote = new Set()) {
    const sheets = new Map(this.sheets.map((s) => [s.name.toLowerCase(), this.sheet(s.name)]))
    // Every cell is read once per recalculation: values (constants too)
    // are remembered per sheet under a numeric key, and a range many
    // formulas share (a SUMIFS column) is gathered once.
    const memo = new Map([...sheets.values()].map((sh) => [sh, new Map()]))
    const active = new Set()
    const ranges = new Map()
    const named = new Map()
    const sheetOf = (name) => {
      let sh = named.get(name)
      if (sh === undefined) { sh = sheets.get(String(name).toLowerCase()); named.set(name, sh) }
      return sh
    }
    let unsupported = 0
    const valueOf = (sh, r, c) => {
      const m = memo.get(sh)
      const key = r * 16385 + c
      const known = m.get(key)
      if (known !== undefined || m.has(key)) return known
      const cell = sh.cellAt(r, c)
      if (!cell) { m.set(key, null); return null }
      const f = sh.formula(cell, r, c)
      if (f === undefined) { const v = sh.value(cell); m.set(key, v); return v }
      if (active.has(cell)) return { error: '#REF!' } // circular
      active.add(cell)
      let v
      try {
        v = evaluate(parseFormula(f), env(sh, r, c))
        if (typeof v === 'number' && !Number.isFinite(v)) v = { error: '#NUM!' }
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e
        v = sh.value(cell) // keep Excel's cached value
        m.set(key, v)
        active.delete(cell)
        throw e
      }
      active.delete(cell)
      m.set(key, v)
      return v
    }
    const rangeOf = (name, rg) => {
      const sh = sheetOf(name)
      const key = `${sh.info.name}|${rg.r1},${rg.c1},${rg.r2},${rg.c2}`
      let values = ranges.get(key)
      if (!values) {
        values = []
        for (let r = rg.r1; r <= rg.r2; r++) {
          const row = []
          for (let c = rg.c1; c <= rg.c2; c++) row.push(valueOf(sh, r, c))
          values.push(row)
        }
        ranges.set(key, values)
      }
      return values
    }
    const names = this.names()
    const env = (sh, r, c) => ({
      sheet: sh.info.name,
      at: { r, c },
      name: (n) => {
        const u = String(n).toLowerCase()
        const idx = this.sheets.indexOf(sh.info)
        return (names.find((d) => d.name.toLowerCase() === u && d.sheet === idx) ?? names.find((d) => d.name.toLowerCase() === u && d.sheet === undefined))?.ref
      },
      has: (name) => !!sheetOf(name),
      cell: (name, r, c) => valueOf(sheetOf(name), r, c),
      range: rangeOf,
      bounds: (name) => sheetOf(name).bounds()
    })
    for (const sh of sheets.values()) {
      for (const [key, cell] of sh.cells) {
        const [r, c] = key.split(',').map(Number)
        const f = sh.formula(cell, r, c)
        if (f === undefined) continue
        let v
        try { v = valueOf(sh, r, c) } catch (e) {
          if (!(e instanceof Unsupported)) throw e
          unsupported++
          if (wrote.has(`${sh.info.name}!${key}`)) sh.setCached(cell, undefined)
          continue
        }
        sh.setCached(cell, v)
      }
    }
    return unsupported
  }

  finish () {
    // Excel rebuilds the calculation chain; a stale one makes it repair
    // the file.
    const chain = this.pkg.relsOfType(this.main, 'calcChain')[0]
    if (chain) {
      this.pkg.removeRel(this.main, chain.id)
      this.pkg.delete(chain.part)
      this.pkg.removeOverride(chain.part)
    }
    let calc = kid(this.wb.root, 'calcPr')
    if (!calc) { calc = el(this.prefix + 'calcPr', { calcId: '191029' }); insertOrdered(this.wb.root, calc, WB_ORDER.map((n) => this.prefix + n)) }
    calc.attrs.fullCalcOnLoad = '1'
    this.pkg.touch(this.main)
    for (const sh of this.cache.values()) sh.finish()
  }
}

// ------------------------------------------------------------- sheet

class Sheet {
  constructor (wb, info) {
    this.wb = wb
    this.info = info
    this.part = info.part
    if (!info.part || !wb.pkg.has(info.part)) throw new Error(`sheet "${info.name}" is not a worksheet (chart sheets and dialog sheets are not supported)`)
    this.d = wb.pkg.xml(info.part)
    this.root = this.d.root
    this.p = this.root.name.includes(':') ? this.root.name.split(':')[0] + ':' : ''
    this.data = kid(this.root, 'sheetData')
    if (!this.data) { this.data = el(this.p + 'sheetData'); insertOrdered(this.root, this.data, WS_ORDER.map((n) => this.p + n)) }
    this.rows = new Map()
    this.cells = new Map()
    let rn = 0
    for (const row of kids(this.data, 'row')) {
      rn = row.attrs.r ? +row.attrs.r : rn + 1
      row.attrs.r = String(rn)
      this.rows.set(rn, row)
      let cn = 0
      for (const c of kids(row, 'c')) {
        const p = c.attrs.r ? parseCell(c.attrs.r) : { r: rn, c: cn + 1 }
        cn = p.c
        c.attrs.r = cellName(rn, cn)
        this.cells.set(`${rn},${cn}`, c)
      }
    }
    // Shared formula masters, by si.
    this.shared = new Map()
    for (const [key, c] of this.cells) {
      const f = kid(c, 'f')
      if (f && f.attrs.t === 'shared' && f.attrs.ref && textContent(f)) {
        const [r, cc] = key.split(',').map(Number)
        this.shared.set(f.attrs.si, { r, c: cc, text: textContent(f) })
      }
    }
  }

  cellAt (r, c) { return this.cells.get(`${r},${c}`) }

  bounds () {
    let r2 = 0; let c2 = 0
    for (const key of this.cells.keys()) {
      const [r, c] = key.split(',').map(Number)
      if (r > r2) r2 = r
      if (c > c2) c2 = c
    }
    return { r1: 1, c1: 1, r2, c2 }
  }

  // usedRange: the smallest range holding a value, formula or style.
  usedRange () {
    let r1 = Infinity; let c1 = Infinity; let r2 = 0; let c2 = 0
    for (const [key, c] of this.cells) {
      if (!c.kids.length && !c.attrs.s) continue
      const [r, cc] = key.split(',').map(Number)
      r1 = Math.min(r1, r); c1 = Math.min(c1, cc); r2 = Math.max(r2, r); c2 = Math.max(c2, cc)
    }
    return r2 ? { r1, c1, r2, c2 } : undefined
  }

  value (c) {
    if (!c) return null
    const t = c.attrs.t
    const vEl = kid(c, 'v')
    const vs = vEl ? textContent(vEl) : undefined
    switch (t) {
      case 's': return vs === undefined ? null : this.wb.strings[+vs] ?? ''
      case 'str': return vs ?? ''
      case 'inlineStr': { const is = kid(c, 'is'); return is ? is.kids.filter((k) => local(k) === 't' || local(k) === 'r').map((k) => local(k) === 't' ? textContent(k) : textContent(kid(k, 't'))).join('') : '' }
      case 'b': return vs === undefined ? null : vs === '1' || vs === 'true'
      case 'e': return vs === undefined ? null : { error: vs }
      case 'd': return vs ?? null
      default: return vs === undefined || vs === '' ? null : Number(vs)
    }
  }

  // formula returns the cell's formula (without =), translating shared
  // formulas; undefined when there is none.
  formula (c, r, cc) {
    const f = kid(c, 'f')
    if (!f) return undefined
    if (f.attrs.t === 'array' || f.attrs.t === 'dataTable') { const s = textContent(f); return s || undefined }
    const s = textContent(f)
    if (s) return s
    if (f.attrs.t === 'shared') {
      const m = this.shared.get(f.attrs.si)
      if (!m) return undefined
      try { return shift(m.text, r - m.r, cc - m.c) } catch { return undefined }
    }
    return undefined
  }

  row (r) {
    let row = this.rows.get(r)
    if (!row) {
      row = el(this.p + 'row', { r: String(r) })
      const after = [...this.rows.keys()].filter((k) => k > r).sort((a, b) => a - b)[0]
      if (after === undefined) this.data.kids.push(row)
      else this.data.kids.splice(this.data.kids.indexOf(this.rows.get(after)), 0, row)
      this.rows.set(r, row)
    }
    return row
  }

  cell (r, c) {
    let cell = this.cellAt(r, c)
    if (!cell) {
      const row = this.row(r)
      cell = el(this.p + 'c', { r: cellName(r, c) })
      // Inherit the row's or column's style, as Excel does when typing.
      const rs = row.attrs.customFormat === '1' ? row.attrs.s : this.colStyle(c)
      if (rs) cell.attrs.s = rs
      const existing = kids(row, 'c')
      const after = existing.find((x) => parseCell(x.attrs.r).c > c)
      if (after) row.kids.splice(row.kids.indexOf(after), 0, cell)
      else row.kids.push(cell)
      this.cells.set(`${r},${c}`, cell)
    }
    return cell
  }

  colStyle (c) {
    for (const col of kids(kid(this.root, 'cols'), 'col')) if (+col.attrs.min <= c && c <= +col.attrs.max && col.attrs.style) return col.attrs.style
    return undefined
  }

  clearContent (cell) {
    // A shared-formula master hands its formula to nobody; give its
    // dependants their own copies first.
    const f = kid(cell, 'f')
    if (f && f.attrs.t === 'shared' && f.attrs.ref && textContent(f)) this.unshare(f.attrs.si)
    cell.kids = cell.kids.filter((k) => !['f', 'v', 'is'].includes(local(k)))
    delete cell.attrs.t
  }

  unshare (si) {
    for (const [key, c] of this.cells) {
      const f = kid(c, 'f')
      if (f && f.attrs.t === 'shared' && f.attrs.si === si) {
        const [r, cc] = key.split(',').map(Number)
        const s = this.formula(c, r, cc)
        delete f.attrs.t; delete f.attrs.si; delete f.attrs.ref
        if (s !== undefined) f.kids = [text(s)]
      }
    }
    this.shared.delete(si)
  }

  // set writes a value or formula into a cell.
  set (r, c, spec) {
    const cell = this.cell(r, c)
    const s = normalize(spec)
    if ('value' in s || 'formula' in s) {
      this.clearContent(cell)
      const p = this.p
      if (s.formula !== undefined) {
        cell.kids.push(el(p + 'f', {}, [text(xlfn(String(s.formula).replace(/^=/, '')))]))
      } else if (s.value === null || s.value === undefined || s.value === '') {
        // Empty: nothing more.
      } else if (typeof s.value === 'number') {
        if (!Number.isFinite(s.value)) throw new Error(`${cellName(r, c)}: ${s.value} is not a finite number`)
        cell.kids.push(el(p + 'v', {}, [text(String(s.value))]))
      } else if (typeof s.value === 'boolean') {
        cell.attrs.t = 'b'
        cell.kids.push(el(p + 'v', {}, [text(s.value ? '1' : '0')]))
      } else {
        cell.attrs.t = 's'
        cell.kids.push(el(p + 'v', {}, [text(String(this.wb.stringId(String(s.value))))]))
      }
    }
    const style = styleProps(s)
    if (Object.keys(style).length) cell.attrs.s = String(this.wb.derive(+(cell.attrs.s ?? 0), style))
    if (cell.attrs.s === '0') delete cell.attrs.s
    this.wb.pkg.touch(this.part)
    return s.formula !== undefined
  }

  setCached (cell, v) {
    cell.kids = cell.kids.filter((k) => local(k) !== 'v')
    if (cell.attrs.t !== undefined && cell.attrs.t !== 's') delete cell.attrs.t
    if (cell.attrs.t === 's') delete cell.attrs.t
    if (v === undefined || v === null) { this.wb.pkg.touch(this.part); return }
    let vs
    if (isErr(v)) { cell.attrs.t = 'e'; vs = v.error } else if (typeof v === 'boolean') { cell.attrs.t = 'b'; vs = v ? '1' : '0' } else if (typeof v === 'number') { vs = String(v) } else { cell.attrs.t = 'str'; vs = String(v) }
    cell.kids.push(el(this.p + 'v', {}, [text(vs)]))
    this.wb.pkg.touch(this.part)
  }

  clear (rg) {
    for (const [key, cell] of [...this.cells]) {
      const [r, c] = key.split(',').map(Number)
      if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) this.clearContent(cell)
    }
    this.wb.pkg.touch(this.part)
  }

  setWidth (c1, c2, width) {
    let cols = kid(this.root, 'cols')
    if (!cols) { cols = el(this.p + 'cols'); insertOrdered(this.root, cols, WS_ORDER.map((n) => this.p + n)) }
    // Split existing ranges around [c1, c2].
    const next = []
    for (const col of kids(cols, 'col')) {
      const a = +col.attrs.min; const b = +col.attrs.max
      if (b < c1 || a > c2) { next.push(col); continue }
      if (a < c1) next.push({ ...clone(col), attrs: { ...col.attrs, min: String(a), max: String(c1 - 1) } })
      if (b > c2) next.push({ ...clone(col), attrs: { ...col.attrs, min: String(c2 + 1), max: String(b) } })
    }
    const style = this.colStyle(c1)
    next.push(el(this.p + 'col', { min: String(c1), max: String(c2), width: String(width), customWidth: '1', ...(style ? { style } : {}) }))
    next.sort((x, y) => +x.attrs.min - +y.attrs.min)
    cols.kids = next
    this.wb.pkg.touch(this.part)
  }

  // colAttrs calls fn on <col> elements covering exactly columns c1..c2,
  // splitting existing ranges and filling gaps with default-width ones.
  colAttrs (c1, c2, fn) {
    let cols = kid(this.root, 'cols')
    if (!cols) { cols = el(this.p + 'cols'); insertOrdered(this.root, cols, WS_ORDER.map((n) => this.p + n)) }
    const fmt = kid(this.root, 'sheetFormatPr')
    const defW = fmt?.attrs.defaultColWidth ?? (fmt?.attrs.baseColWidth ? String(+fmt.attrs.baseColWidth + 0.71) : '8.43')
    const piece = (col, a, b) => ({ ...clone(col), attrs: { ...col.attrs, min: String(a), max: String(b) } })
    const next = []
    const mine = []
    for (const col of kids(cols, 'col')) {
      const a = +col.attrs.min; const b = +col.attrs.max
      if (b < c1 || a > c2) { next.push(col); continue }
      if (a < c1) next.push(piece(col, a, c1 - 1))
      if (b > c2) next.push(piece(col, c2 + 1, b))
      mine.push(piece(col, Math.max(a, c1), Math.min(b, c2)))
    }
    mine.sort((x, y) => +x.attrs.min - +y.attrs.min)
    let at = c1
    for (const m of [...mine]) {
      if (+m.attrs.min > at) mine.push(el(this.p + 'col', { min: String(at), max: String(+m.attrs.min - 1), width: defW }))
      at = +m.attrs.max + 1
    }
    if (at <= c2) mine.push(el(this.p + 'col', { min: String(at), max: String(c2), width: defW }))
    for (const m of mine) { fn(m); if (!m.attrs.width) m.attrs.width = defW; next.push(m) }
    next.sort((x, y) => +x.attrs.min - +y.attrs.min)
    cols.kids = next
    this.wb.pkg.touch(this.part)
  }

  widths () {
    const out = new Map()
    for (const col of kids(kid(this.root, 'cols'), 'col')) {
      for (let c = +col.attrs.min; c <= Math.min(+col.attrs.max, 200); c++) if (col.attrs.width) out.set(c, +col.attrs.width)
    }
    return out
  }

  freeze (at) {
    const p = parseCell(at)
    if (!p) throw new Error(`freeze: "${at}" is not a cell like A2`)
    let views = kid(this.root, 'sheetViews')
    if (!views) { views = el(this.p + 'sheetViews'); insertOrdered(this.root, views, WS_ORDER.map((n) => this.p + n)) }
    let view = kid(views, 'sheetView')
    if (!view) { view = el(this.p + 'sheetView', { workbookViewId: '0' }); views.kids.push(view) }
    view.kids = view.kids.filter((k) => !['pane', 'selection'].includes(local(k)))
    if (p.r === 1 && p.c === 1) return
    const xs = p.c - 1; const ys = p.r - 1
    const pane = xs && ys ? 'bottomRight' : ys ? 'bottomLeft' : 'topRight'
    const attrs = {}
    if (xs) attrs.xSplit = String(xs)
    if (ys) attrs.ySplit = String(ys)
    Object.assign(attrs, { topLeftCell: cellName(p.r, p.c), activePane: pane, state: 'frozen' })
    view.kids.unshift(el(this.p + 'pane', attrs), el(this.p + 'selection', { pane, activeCell: cellName(p.r, p.c), sqref: cellName(p.r, p.c) }))
    this.wb.pkg.touch(this.part)
  }

  frozen () {
    const pane = descendants(this.root, (n) => local(n) === 'pane')[0]
    return pane?.attrs.state === 'frozen' || pane?.attrs.state === 'frozenSplit' ? pane.attrs.topLeftCell : undefined
  }

  merges () {
    return kids(kid(this.root, 'mergeCells'), 'mergeCell').map((m) => m.attrs.ref)
  }

  finish () {
    const u = this.usedRange()
    let dim = kid(this.root, 'dimension')
    if (!dim) { dim = el(this.p + 'dimension'); insertOrdered(this.root, dim, WS_ORDER.map((n) => this.p + n)) }
    dim.attrs.ref = u ? (u.r1 === u.r2 && u.c1 === u.c2 ? cellName(u.r1, u.c1) : `${cellName(u.r1, u.c1)}:${cellName(u.r2, u.c2)}`) : 'A1'
    // Rows list their span; a stale one is harmless but keep it right.
    for (const row of this.rows.values()) delete row.attrs.spans
    this.wb.pkg.touch(this.part)
  }

  // display returns what a cell shows.
  display (r, c) {
    const cell = this.cellAt(r, c)
    if (!cell) return ''
    return formatValue(this.value(cell), this.wb.fmtCode(cell.attrs.s ? +cell.attrs.s : 0))
  }
}

// Functions newer than Excel 2007 are stored with a _xlfn. prefix, or
// Excel shows #NAME?.
const NEW_FUNCS = 'CONCAT|TEXTJOIN|IFS|SWITCH|MAXIFS|MINIFS|XLOOKUP|XMATCH|IFNA|DAYS|ISOWEEKNUM|STDEV\\.S|STDEV\\.P|VAR\\.S|VAR\\.P|PERCENTILE\\.INC|PERCENTILE\\.EXC|QUARTILE\\.INC|QUARTILE\\.EXC|MODE\\.SNGL|RANK\\.EQ|RANK\\.AVG|NORM\\.DIST|NORM\\.INV|FLOOR\\.MATH|CEILING\\.MATH|AGGREGATE|LET|TEXTBEFORE|TEXTAFTER|TEXTSPLIT|UNIQUE|SORT|SORTBY|FILTER|SEQUENCE'
const NEW_RE = new RegExp(`(^|[^\\w.])(${NEW_FUNCS})\\(`, 'gi')
export function xlfn (f) {
  // Leave string literals alone.
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(NEW_RE, (m, pre, fn) => `${pre}_xlfn.${fn.toUpperCase()}(`))).join('')
}
const unxlfn = (f) => f.replace(/_xlfn\.(_xlws\.)?/gi, '')

function hex (c) {
  const s = String(c).replace(/^#/, '').toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(s)) throw new Error(`"${c}" is not a 6-digit hex colour like 1F4E79`)
  return s
}

// normalize turns a cell given by the Agent into { value?, formula?, …style }.
function normalize (spec) {
  if (spec === null || typeof spec !== 'object' || Array.isArray(spec)) {
    if (typeof spec === 'string' && spec.startsWith('=') && spec.length > 1) return { formula: spec.slice(1) }
    return { value: spec ?? null }
  }
  const s = { ...spec }
  if (typeof s.value === 'string' && s.value.startsWith('=') && s.value.length > 1 && s.formula === undefined) { s.formula = s.value.slice(1); delete s.value }
  if (s.formula !== undefined) delete s.value
  return s
}

function styleProps (s) {
  const out = {}
  for (const k of ['bold', 'italic', 'underline', 'color', 'fill', 'format', 'align', 'valign', 'wrap', 'border', 'size', 'locked']) if (s[k] !== undefined) out[k] = s[k]
  if (out.align && !['left', 'center', 'right', 'justify', 'general'].includes(out.align)) throw new Error(`align must be left, center or right, not "${out.align}"`)
  return out
}

// ------------------------------------------------------------- reading

const LIMIT = 16000
const MAX_ROWS = 200
const MAX_COLS = 40

export function readXlsx (buf, { sheet, range, from } = {}) {
  const wb = new Workbook(Package.fromBuffer(buf))
  const lines = [`Excel workbook: ${wb.sheets.length} sheet${wb.sheets.length === 1 ? '' : 's'}`]
  for (const s of wb.sheets) {
    let desc
    try {
      const sh = wb.sheet(s.name)
      const u = sh.usedRange()
      desc = u ? `${cellName(u.r1, u.c1)}:${cellName(u.r2, u.c2)} (${u.r2 - u.r1 + 1} rows × ${u.c2 - u.c1 + 1} columns)` : 'empty'
      const fz = sh.frozen()
      if (fz) desc += `, frozen at ${fz}`
      const mg = sh.merges()
      if (mg.length) desc += `, merged ${mg.slice(0, 6).join(' ')}${mg.length > 6 ? ' …' : ''}`
    } catch (e) { desc = e.message }
    lines.push(`- ${s.name}: ${desc}${s.state && s.state !== 'visible' ? ` (${s.state})` : ''}`)
  }
  const sh = wb.sheet(sheet)
  const u = sh.usedRange()
  if (!u) { lines.push('', `Sheet "${sh.info.name}" is empty.`); return lines.join('\n') }
  let rg = range ? parseRange(range) : u
  if (!rg) throw new Error(`range "${range}" is not like A1:D20`)
  if (!range) rg = { r1: 1, c1: 1, r2: u.r2, c2: u.c2 }
  const start = Math.max(rg.r1, from ? +from : rg.r1)
  const c2 = Math.min(rg.c2, rg.c1 + MAX_COLS - 1)
  let r2 = Math.min(rg.r2, start + MAX_ROWS - 1)
  const header = '| | ' + Array.from({ length: c2 - rg.c1 + 1 }, (_, k) => colName(rg.c1 + k)).join(' | ') + ' |'
  const body = []
  let size = header.length
  let unsupported = 0
  for (let r = start; r <= r2; r++) {
    const cells = []
    for (let c = rg.c1; c <= c2; c++) {
      const cell = sh.cellAt(r, c)
      if (!cell) { cells.push(''); continue }
      const f = sh.formula(cell, r, c)
      let shown = sh.display(r, c)
      if (f !== undefined) {
        const hasV = !!kid(cell, 'v')
        if (!hasV) unsupported++
        shown = `=${unxlfn(f)} → ${hasV ? shown : '(calculated when opened in Excel)'}`
      }
      cells.push(shown.replace(/\|/g, '\\|').replace(/\n/g, '\\n'))
    }
    const line = `| ${r} | ${cells.join(' | ')} |`
    if (size + line.length > LIMIT && body.length) { r2 = r - 1; break }
    body.push(line)
    size += line.length + 1
  }
  lines.push('', `Sheet "${sh.info.name}", ${cellName(start, rg.c1)}:${cellName(r2, c2)}${c2 < rg.c2 ? ` (columns after ${colName(c2)} not shown: pass range)` : ''}:`, header, '|' + ' --- |'.repeat(c2 - rg.c1 + 2), ...body)
  if (r2 < rg.r2) lines.push('', `[more rows: call office_read with sheet: "${sh.info.name}", from: ${r2 + 1}${range ? `, range: "${range}"` : ''}]`)
  if (unsupported) lines.push('', `${unsupported} formula(s) use functions Quire does not calculate; Excel/WPS calculates them when the file is opened.`)
  return lines.join('\n')
}

// ------------------------------------------------------------- create & edit

function applyRows (sh, at, rows, wrote) {
  const p = parseCell(at ?? 'A1')
  if (!p) throw new Error(`"${at}" is not a cell like A1`)
  if (!Array.isArray(rows)) throw new Error('rows must be an array of arrays')
  rows.forEach((row, i) => {
    if (!Array.isArray(row)) row = [row]
    row.forEach((v, j) => {
      if (v === undefined) return
      if (sh.set(p.r + i, p.c + j, v)) wrote.add(`${sh.info.name}!${p.r + i},${p.c + j}`)
    })
  })
}

function autoWidths (sh) {
  const u = sh.usedRange()
  if (!u) return
  for (let c = u.c1; c <= u.c2; c++) {
    let w = 8
    for (let r = u.r1; r <= Math.min(u.r2, u.r1 + 300); r++) {
      const s = sh.display(r, c)
      let n = 0
      for (const ch of s) n += /[\u1100-\uffff]/.test(ch) ? 2 : 1
      w = Math.max(w, Math.min(60, n + 2))
    }
    sh.setWidth(c, c, w)
  }
}

export function createXlsx (sheets) {
  if (!Array.isArray(sheets) || !sheets.length) throw new Error('sheets must list at least one sheet')
  const pkg = newPackage()
  const wb = new Workbook(pkg)
  const wrote = new Set()
  for (const [k, s] of sheets.entries()) {
    const sh = wb.addSheet(s.name || `Sheet${k + 1}`)
    applyRows(sh, 'A1', s.rows ?? [], wrote)
    if (s.header && s.rows?.length) {
      const width = Math.max(...s.rows.map((r) => (Array.isArray(r) ? r.length : 1)))
      for (let c = 1; c <= width; c++) sh.set(1, c, { bold: true, fill: 'D9E2F3', border: true })
      for (let r = 2; r <= s.rows.length; r++) for (let c = 1; c <= width; c++) if (sh.cellAt(r, c)) sh.set(r, c, { border: true })
      if (!s.freeze) sh.freeze('A2')
    }
    if (s.freeze) sh.freeze(s.freeze)
    if (Array.isArray(s.widths) && s.widths.length) s.widths.forEach((w, i) => { if (w) sh.setWidth(i + 1, i + 1, w) })
    else autoWidths(sh)
  }
  const unsupported = wb.recalc(wrote)
  wb.finish()
  return { buffer: pkg.toBuffer(), summary: `${sheets.length} sheet(s)${unsupported ? `; ${unsupported} formula(s) left for Excel to calculate` : ''}` }
}

export function editXlsx (buf, edits) {
  const pkg = Package.fromBuffer(buf)
  const wb = new Workbook(pkg)
  const wrote = new Set()
  const report = []
  const rangeOf = (e, where) => {
    const rg = parseRange(e.range ?? e.cell ?? '')
    if (!rg) throw new Error(`${where}: range "${e.range ?? e.cell}" is not like A1 or A1:C9`)
    return rg
  }
  for (const [k, e] of edits.entries()) {
    const where = `edits[${k}] (${e?.op})`
    switch (e?.op) {
      case 'set': {
        const sh = wb.sheet(e.sheet)
        const p = parseCell(e.cell ?? '')
        if (!p) throw new Error(`${where}: cell "${e.cell}" is not like B3`)
        const spec = { ...e }
        delete spec.op; delete spec.sheet; delete spec.cell
        if (sh.set(p.r, p.c, spec)) wrote.add(`${sh.info.name}!${p.r},${p.c}`)
        report.push(`set ${sh.info.name}!${e.cell}`)
        break
      }
      case 'write': {
        const sh = wb.sheet(e.sheet)
        applyRows(sh, e.at ?? 'A1', e.rows, wrote)
        report.push(`wrote ${e.rows?.length ?? 0} row(s) at ${sh.info.name}!${e.at ?? 'A1'}`)
        break
      }
      case 'style': {
        const sh = wb.sheet(e.sheet)
        const rg = rangeOf(e, where)
        const props = styleProps(e)
        if (Object.keys(props).length) {
          if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 200000) throw new Error(`${where}: the range is too large to style cell by cell`)
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) sh.set(r, c, props)
        }
        if (e.width !== undefined) sh.setWidth(rg.c1, rg.c2, +e.width)
        report.push(`styled ${sh.info.name}!${e.range ?? e.cell}`)
        break
      }
      case 'clear': {
        const sh = wb.sheet(e.sheet)
        sh.clear(rangeOf(e, where))
        report.push(`cleared ${sh.info.name}!${e.range}`)
        break
      }
      case 'freeze': {
        const sh = wb.sheet(e.sheet)
        sh.freeze(e.cell ?? e.at ?? 'A1')
        report.push(`froze ${sh.info.name} at ${e.cell ?? e.at}`)
        break
      }
      case 'add_sheet': {
        const sh = wb.addSheet(e.name)
        if (e.rows) { applyRows(sh, 'A1', e.rows, wrote); autoWidths(sh) }
        report.push(`added sheet ${e.name}`)
        break
      }
      case 'rename_sheet':
        wb.renameSheet(e.sheet, e.name)
        report.push(`renamed ${e.sheet} to ${e.name}`)
        break
      case 'delete_sheet':
        wb.deleteSheet(e.sheet)
        report.push(`deleted sheet ${e.sheet}`)
        break
      default:
        throw new Error(`${where}: unknown op; use set, write, style, clear, freeze, add_sheet, rename_sheet or delete_sheet`)
    }
  }
  const unsupported = wb.recalc(wrote)
  wb.finish()
  if (unsupported) report.push(`${unsupported} formula(s) left for Excel to calculate`)
  return { buffer: pkg.toBuffer(), summary: report.join('; ') }
}

// ------------------------------------------------------------- rendering

const escH = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export function renderXlsx (buf) {
  const wb = new Workbook(Package.fromBuffer(buf))
  const tabs = []
  const panes = []
  let landscape = false
  wb.sheets.forEach((s, k) => {
    let sh
    try { sh = wb.sheet(s.name) } catch { return }
    const id = `s${k}`
    tabs.push(`<input type="radio" name="tab" id="${id}"${k === 0 ? ' checked' : ''}><label for="${id}">${escH(s.name)}</label>`)
    // Printing keeps to the print area when the sheet has one.
    const page = sheetPage(wb, sh)
    if (page.orientation === 'landscape') landscape = true
    const area = page.area ? parseRange(page.area.replace(/^.*!/, '').replace(/\$/g, '')) : undefined
    const u = area ?? sh.usedRange()
    if (!u) { panes.push(`<div class="pane" data-for="${id}"><p class="empty">（空工作表）</p></div>`); return }
    const r0 = area ? area.r1 : 1; const c0 = area ? area.c1 : 1
    const r2 = Math.min(u.r2, r0 + 499); const c2 = Math.min(u.c2, c0 + 59)
    const widths = sh.widths()
    const covered = new Set()
    const spans = new Map()
    for (const ref of sh.merges()) {
      const m = parseRange(ref)
      if (!m) continue
      spans.set(`${m.r1},${m.c1}`, m)
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`)
    }
    let html = '<table class="grid"><colgroup><col class="rh">'
    for (let c = c0; c <= c2; c++) html += `<col style="width:${Math.round((widths.get(c) ?? 8.43) * 7 + 5)}px">`
    html += '</colgroup><tr><th></th>'
    for (let c = c0; c <= c2; c++) html += `<th>${colName(c)}</th>`
    html += '</tr>'
    for (let r = r0; r <= r2; r++) {
      html += `<tr><th>${r}</th>`
      for (let c = c0; c <= c2; c++) {
        if (covered.has(`${r},${c}`)) continue
        const cell = sh.cellAt(r, c)
        const span = spans.get(`${r},${c}`)
        const attrs = span ? `${span.c2 > span.c1 ? ` colspan="${span.c2 - span.c1 + 1}"` : ''}${span.r2 > span.r1 ? ` rowspan="${span.r2 - span.r1 + 1}"` : ''}` : ''
        if (!cell) { html += `<td${attrs}></td>`; continue }
        const v = sh.value(cell)
        const st = wb.css(cell.attrs.s ? +cell.attrs.s : 0)
        const css = []
        if (st.b) css.push('font-weight:700')
        if (st.i) css.push('font-style:italic')
        if (st.u) css.push('text-decoration:underline')
        if (st.color && st.color !== '000000') css.push(`color:#${st.color}`)
        if (st.fill) css.push(`background:#${st.fill}`)
        if (st.size && +st.size !== 11) css.push(`font-size:${+st.size}pt`)
        const align = st.align && st.align !== 'general' ? st.align : typeof v === 'number' ? 'right' : (typeof v === 'boolean' || isErr(v)) ? 'center' : undefined
        if (align) css.push(`text-align:${align === 'centerContinuous' ? 'center' : align}`)
        if (st.wrap) css.push('white-space:pre-wrap')
        if (st.border) css.push('border-color:#8c8c8c')
        const f = sh.formula(cell, r, c)
        html += `<td${attrs}${css.length ? ` style="${css.join(';')}"` : ''}${f ? ` title="=${escH(unxlfn(f))}"` : ''}>${escH(sh.display(r, c))}</td>`
      }
      html += '</tr>'
    }
    html += '</table>'
    if (u.r2 > r2 || u.c2 > c2) html += `<p class="more">预览只显示前 ${r2} 行、${c2} 列（共 ${u.r2} 行、${u.c2} 列）。</p>`
    panes.push(`<div class="pane" data-for="${id}">${html}</div>`)
  })
  const show = wb.sheets.map((_, k) => `#s${k}:checked~.panes .pane[data-for="s${k}"]`).join(',')
  return {
    kind: 'Excel 表格',
    css: XLSX_CSS + `${show}{display:block}` + (landscape ? '@page{size:landscape}' : ''),
    body: `<div class="book">${tabs.join('')}<div class="panes">${panes.join('')}</div></div>`
  }
}

const XLSX_CSS = `
.book{margin:16px}.book>input{display:none}
.book>label{display:inline-block;padding:5px 14px;margin-right:2px;background:#e7e7e7;border:1px solid #c8c8c8;border-bottom:none;border-radius:4px 4px 0 0;font:13px/1.4 -apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;cursor:pointer;color:#333}
.book>input:checked+label{background:#fff;color:#107C41;font-weight:600}
.panes{background:#fff;border:1px solid #c8c8c8;overflow:auto;max-height:calc(100vh - 120px)}
.pane{display:none}
.grid{border-collapse:collapse;table-layout:fixed;font:11pt/1.35 Calibri,"PingFang SC","Microsoft YaHei",sans-serif;color:#000}
.grid th{background:#f3f3f3;color:#555;font-weight:400;font-size:12px;border:1px solid #d4d4d4;padding:1px 4px;position:sticky;top:0}
.grid tr th:first-child{position:sticky;left:0;text-align:center;min-width:36px}
.grid col.rh{width:40px}
.grid td{border:1px solid #e1e1e1;padding:1px 4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;height:20px}
.empty,.more{color:#888;font:13px sans-serif;padding:8px 12px;margin:0}`

// ------------------------------------------------------------- editing in the UI
//
// xlsxUiModel hands the browser every cell's shown text, its input (the
// formula or raw value) and its style; editXlsxUi takes what the person
// typed and changed, on top of the Agent's edit ops.

const isDateFmt = (code) => {
  if (!code || code === 'General') return false
  const bare = splitSections(code)[0].replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '')
  return /[yd]|m{1,4}(?![^h]*s)|[hs]/i.test(bare) && !/[#0?]/.test(bare)
}
const pxOfWidth = (w) => Math.round(w * 7 + 5)
const MAX_UI_CELLS = 300000

// calcIfStale works out formulas that carry no cached value, as Excel
// does on open: files written by scripts (openpyxl and the like) often
// have none. Only the in-memory copy changes; the file is untouched
// until the next save.
function calcIfStale (wb) {
  for (const info of wb.sheets) {
    let sh
    try { sh = wb.sheet(info.name) } catch { continue }
    for (const c of sh.cells.values()) {
      if (kid(c, 'f') && !textContent(kid(c, 'v') ?? { kids: [] })) { wb.recalc(); return }
    }
  }
}

export function xlsxUiModel (buf) {
  const wb = new Workbook(Package.fromBuffer(buf))
  calcIfStale(wb)
  const styles = {}
  let total = 0
  const sheets = wb.sheets.map((s) => {
    let sh
    try { sh = wb.sheet(s.name) } catch (e) { return { name: s.name, state: s.state, unsupported: e.message } }
    const fmtPr = kid(sh.root, 'sheetFormatPr')
    const defW = fmtPr?.attrs.defaultColWidth ? pxOfWidth(+fmtPr.attrs.defaultColWidth) : fmtPr?.attrs.baseColWidth ? pxOfWidth(+fmtPr.attrs.baseColWidth + 0.71) : 64
    const defH = fmtPr?.attrs.defaultRowHeight ? Math.round(+fmtPr.attrs.defaultRowHeight * 4 / 3) : 20
    const cols = kids(kid(sh.root, 'cols'), 'col').map((c) => ({ min: +c.attrs.min, max: Math.min(+c.attrs.max, 16384), px: c.attrs.width ? pxOfWidth(+c.attrs.width) : defW, hidden: c.attrs.hidden === '1' }))
    const rows = {}
    for (const [r, row] of sh.rows) {
      if (row.attrs.hidden === '1') rows[r] = 0
      else if (row.attrs.ht && row.attrs.customHeight === '1') rows[r] = Math.round(+row.attrs.ht * 4 / 3)
    }
    const cells = {}
    let maxR = 0; let maxC = 0
    let cut = false
    for (const [key, cell] of sh.cells) {
      if (total >= MAX_UI_CELLS) { cut = true; break }
      const [r, c] = key.split(',').map(Number)
      const xf = cell.attrs.s ? +cell.attrs.s : 0
      const f = sh.formula(cell, r, c)
      const v = sh.value(cell)
      if (f === undefined && (v === null || v === '') && !xf) continue
      if (!(xf in styles)) styles[xf] = { ...wb.css(xf), fmt: wb.fmtCode(xf), valign: kid(wb.xfs[xf], 'alignment')?.attrs.vertical, ...(kid(wb.xfs[xf], 'protection')?.attrs.locked === '0' ? { unlocked: true } : {}) }
      const shown = sh.display(r, c)
      let input
      if (f !== undefined) input = '=' + unxlfn(f)
      else if (v === null) input = ''
      else if (typeof v === 'boolean') input = v ? 'TRUE' : 'FALSE'
      else if (isErr(v)) input = v.error
      else if (typeof v === 'number') input = isDateFmt(styles[xf].fmt) ? shown : /%/.test(styles[xf].fmt.replace(/"[^"]*"/g, '')) ? `${+(v * 100).toPrecision(12)}%` : String(v)
      else input = String(v)
      cells[key] = [shown, input, xf, typeof v === 'number' ? 1 : typeof v === 'boolean' || isErr(v) ? 2 : 0]
      total++
      if (r > maxR) maxR = r
      if (c > maxC) maxC = c
    }
    const fz = sh.frozen()
    const fp = fz ? parseCell(fz) : undefined
    const af = kid(sh.root, 'autoFilter')
    const afr = af && parseRange(af.attrs.ref ?? '')
    const autoFilter = afr ? { ...afr, cols: Object.fromEntries(kids(af, 'filterColumn').map((fc) => [afr.c1 + +fc.attrs.colId, filterValues(fc)])) } : undefined
    return {
      autoFilter,
      name: s.name,
      state: s.state,
      cells,
      merges: sh.merges().map(parseRange).filter(Boolean),
      cols,
      rows,
      defW,
      defH,
      frozen: fp ? { r: fp.r - 1, c: fp.c - 1 } : undefined,
      maxR,
      maxC,
      charts: sheetCharts(wb, sh),
      ...cfModel(wb, sh),
      ...(cut ? { cut: true } : {})
    }
  })
  return { kind: 'xlsx', sheets, styles, names: wb.names() }
}

function cfModel (wb, sh) {
  const out = {}
  const cf = sheetCf(wb, sh)
  if (cf.length) { out.cf = cf; out.cfx = cfCells(wb, sh, cf) }
  const dv = sheetDv(wb, sh)
  if (dv.length) out.dv = dv
  const notes = sheetNotes(wb, sh)
  if (notes) out.notes = notes
  const links = sheetLinks(wb, sh)
  if (links) out.links = links
  const lv = sheetOutline(sh)
  if (lv) out.lv = lv
  const pics = sheetPics(wb, sh)
  if (pics.length) out.pics = pics
  if (kid(sh.root, 'sheetProtection')?.attrs.sheet === '1' || kid(sh.root, 'sheetProtection')?.attrs.sheet === 'true') out.protect = true
  out.page = sheetPage(wb, sh)
  return out
}

// parseInput reads what a person typed into a cell the way Excel does.
function parseInput (text, fmt) {
  const s = String(text ?? '')
  if (s === '') return { value: null }
  if (s.startsWith('=') && s.length > 1) return { formula: s.slice(1) }
  if (s.startsWith("'")) return { value: s.slice(1), text: true }
  const t = s.trim()
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return { value: Number(t) }
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { value: Number(t.replace(/,/g, '')), format: fmt === 'General' ? (t.includes('.') ? '#,##0.00' : '#,##0') : undefined }
  let m
  if ((m = /^([+-]?(?:\d+\.?\d*|\.\d+))%$/.exec(t))) return { value: Number(m[1]) / 100, format: fmt === 'General' ? (m[1].includes('.') ? '0.00%' : '0%') : undefined }
  if (/^(true|false)$/i.test(t)) return { value: /^true$/i.test(t) }
  if ((m = /^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0))
    if (!Number.isNaN(ms)) {
      let serial = ms / 86400000 + 25569
      if (serial < 61) serial -= 1
      return { value: serial, format: isDateFmt(fmt) ? undefined : (m[4] ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd') }
    }
  }
  return { value: s }
}

export function editXlsxUi (buf, ops) {
  const pkg = Package.fromBuffer(buf)
  const wb = new Workbook(pkg)
  const wrote = new Set()
  const passthrough = []
  const notes = []
  const flush = () => { if (passthrough.length) editOps(wb, passthrough.splice(0), wrote) }
  for (const [k, e] of ops.entries()) {
    const where = `ops[${k}] (${e?.op})`
    switch (e?.op) {
      case 'input': case 'inputs': {
        flush()
        const sh = wb.sheet(e.sheet)
        const p = parseCell(e.cell ?? e.at ?? '')
        if (!p) throw new Error(`${where}: "${e.cell ?? e.at}" is not a cell`)
        const rows = e.op === 'input' ? [[e.text]] : e.rows
        if (!Array.isArray(rows)) throw new Error(`${where}: rows must be an array`)
        rows.forEach((row, i) => (Array.isArray(row) ? row : [row]).forEach((txt, j) => {
          if (txt === undefined || txt === null) return
          const r = p.r + i; const c = p.c + j
          const cur = sh.cellAt(r, c)
          const fmt = wb.fmtCode(cur?.attrs.s ? +cur.attrs.s : 0)
          const parsed = parseInput(txt, fmt)
          if (parsed.value === null && parsed.formula === undefined) { if (cur) sh.clearContent(cur); wb.pkg.touch(sh.part); return }
          const spec = parsed.formula !== undefined ? { formula: parsed.formula } : { value: parsed.value }
          if (parsed.format) spec.format = parsed.format
          if (sh.set(r, c, spec)) wrote.add(`${sh.info.name}!${r},${c}`)
        }))
        break
      }
      case 'xf': {
        flush()
        const sh = wb.sheet(e.sheet)
        const rg = parseRange(e.range ?? e.cell ?? '')
        if (!rg) throw new Error(`${where}: bad range`)
        const xf = Math.max(0, Math.min(wb.xfs.length - 1, +e.xf || 0))
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          const cell = xf || sh.cellAt(r, c) ? sh.cell(r, c) : undefined
          if (!cell) continue
          if (xf) cell.attrs.s = String(xf); else delete cell.attrs.s
        }
        wb.pkg.touch(sh.part)
        break
      }
      case 'width': {
        flush()
        const sh = wb.sheet(e.sheet)
        sh.setWidth(+e.col1, +(e.col2 ?? e.col1), Math.max(0.5, Math.round((Math.max(8, +e.px) - 5) / 7 * 100) / 100))
        break
      }
      case 'height': {
        flush()
        const sh = wb.sheet(e.sheet)
        for (let r = +e.row1; r <= +(e.row2 ?? e.row1); r++) {
          const row = sh.row(r)
          row.attrs.ht = String(Math.round(Math.max(6, +e.px) * 0.75 * 100) / 100)
          row.attrs.customHeight = '1'
        }
        wb.pkg.touch(sh.part)
        break
      }
      case 'insert_rows': case 'delete_rows': case 'insert_cols': case 'delete_cols': {
        flush()
        const rowsAxis = e.op.endsWith('rows')
        const n = Math.max(1, Math.min(10000, +e.count || 1))
        moveLines(wb, wb.sheet(e.sheet), rowsAxis ? 'r' : 'c', +e.at, e.op.startsWith('insert') ? n : -n, wrote)
        break
      }
      case 'border': {
        flush()
        setBorders(wb, wb.sheet(e.sheet), parseRange(e.range ?? ''), e.kind, e.color, where)
        break
      }
      case 'autofill': {
        flush()
        autofill(wb.sheet(e.sheet), parseRange(e.from ?? ''), parseRange(e.to ?? ''), wrote, where)
        break
      }
      case 'sort': {
        flush()
        sortRange(wb.sheet(e.sheet), parseRange(e.range ?? ''), Array.isArray(e.keys) && e.keys.length ? e.keys.map((k) => ({ col: +k.col, desc: !!k.desc })) : [{ col: +e.col, desc: !!e.desc }], !!e.header, wrote, where)
        break
      }
      case 'filter': {
        flush()
        setFilter(wb.sheet(e.sheet), e, where)
        break
      }
      case 'filter_off': {
        flush()
        const sh = wb.sheet(e.sheet)
        const af = kid(sh.root, 'autoFilter')
        if (af) {
          const rg = parseRange(af.attrs.ref ?? '')
          if (rg) for (let r = rg.r1 + 1; r <= rg.r2; r++) { const row = sh.rows.get(r); if (row) delete row.attrs.hidden }
          remove(sh.root, af)
          wb.pkg.touch(sh.part)
        }
        break
      }
      case 'hide': {
        flush()
        const sh = wb.sheet(e.sheet)
        const a = Math.max(1, +e.from); const b = Math.max(a, +(e.to ?? e.from))
        if (!Number.isFinite(a) || !Number.isFinite(b) || b - a > 100000) throw new Error(`${where}: bad range of ${e.axis === 'c' ? 'columns' : 'rows'}`)
        const hidden = e.hidden !== false
        if (e.axis === 'c') {
          sh.colAttrs(a, b, (col) => {
            if (hidden) col.attrs.hidden = '1'
            else { delete col.attrs.hidden; if (col.attrs.width === '0') delete col.attrs.width }
          })
        } else {
          for (let r = a; r <= b; r++) {
            if (hidden) sh.row(r).attrs.hidden = '1'
            else { const row = sh.rows.get(r); if (row) { delete row.attrs.hidden; if (row.attrs.ht === '0') { delete row.attrs.ht; delete row.attrs.customHeight } } }
          }
          wb.pkg.touch(sh.part)
        }
        break
      }
      case 'add_chart': flush(); addChart(wb, wb.sheet(e.sheet), e, where); break
      case 'chart_move': flush(); moveChart(wb, wb.sheet(e.sheet), e, where); break
      case 'chart_delete': flush(); deleteChart(wb, wb.sheet(e.sheet), e, where); break
      case 'chart_edit': flush(); editChart(wb, wb.sheet(e.sheet), e, where); break
      case 'move_sheet': flush(); wb.moveSheet(e.sheet, e.to); break
      case 'dedupe': { flush(); const n = dedupe(wb.sheet(e.sheet), parseRange(e.range ?? ''), e.cols, !!e.header, wrote, where); notes.push(`删除了 ${n} 个重复行`); break }
      case 'note_set': flush(); setNote(wb, wb.sheet(e.sheet), e, where); break
      case 'link_set': flush(); setLink(wb, wb.sheet(e.sheet), e, where); break
      case 'outline': flush(); setOutline(wb.sheet(e.sheet), e, where); break
      case 'protect': flush(); protectSheet(wb, wb.sheet(e.sheet), e); break
      case 'unprotect': flush(); unprotectSheet(wb, wb.sheet(e.sheet), e); break
      case 'page': flush(); setPage(wb, wb.sheet(e.sheet), e, where); break
      case 'pic_add': flush(); addPic(wb, wb.sheet(e.sheet), e, where); break
      case 'pic_move': flush(); movePic(wb, wb.sheet(e.sheet), e, where); break
      case 'pic_delete': flush(); deletePic(wb, wb.sheet(e.sheet), e, where); break
      case 'cf_add': flush(); addCf(wb, wb.sheet(e.sheet), e, where); break
      case 'cf_clear': flush(); clearCf(wb, wb.sheet(e.sheet), e); break
      case 'cf_delete': flush(); deleteCf(wb, wb.sheet(e.sheet), e, where); break
      case 'dv_set': flush(); setDv(wb, wb.sheet(e.sheet), e, where); break
      case 'dv_clear': flush(); clearDv(wb, wb.sheet(e.sheet), e); break
      case 'name_set': {
        flush()
        const at = e.sheet === undefined || e.sheet === null || e.sheet === '' ? undefined : wb.sheets.findIndex((x) => x.name === e.sheet)
        if (at === -1) throw new Error(`${where}: no sheet "${e.sheet}"`)
        if (e.old && (e.old !== e.name || e.oldSheet !== e.sheet)) {
          const was = e.oldSheet === undefined || e.oldSheet === null || e.oldSheet === '' ? undefined : wb.sheets.findIndex((x) => x.name === e.oldSheet)
          wb.setName(e.old, null, was)
        }
        wb.setName(e.name, e.ref, at, e.comment)
        for (const x of wb.sheets) { try { const o = wb.sheet(x.name); for (const key of o.cells.keys()) wrote.add(`${x.name}!${key}`) } catch {} }
        break
      }
      case 'name_delete': {
        flush()
        const at = e.sheet === undefined || e.sheet === null || e.sheet === '' ? undefined : wb.sheets.findIndex((x) => x.name === e.sheet)
        wb.setName(e.name, null, at)
        break
      }
      case 'sheet_state': flush(); wb.showSheet(e.sheet, !e.hidden); break
      default:
        passthrough.push(e)
    }
  }
  flush()
  const unsupported = wb.recalc(wrote)
  refreshCharts(wb)
  wb.finish()
  return { buffer: pkg.toBuffer(), summary: `${ops.length} change(s)${unsupported ? `; ${unsupported} formula(s) left for Excel to calculate` : ''}${notes.length ? '; ' + notes.join('; ') : ''}` }
}

// ------------------------------------------------------------- conditional formats

const CF_ORDER = ['formula', 'colorScale', 'dataBar', 'iconSet', 'extLst']

// dxfOf reads a differential style (what a conditional format paints).
function dxfOf (wb, i) {
  const d = kids(kid(wb.st, 'dxfs'), 'dxf')[i]
  if (!d) return {}
  const f = kid(d, 'font'); const pf = kid(kid(d, 'fill'), 'patternFill')
  const on = (n) => { const k = kid(f, n); return k ? k.attrs.val !== '0' && k.attrs.val !== 'false' : undefined }
  const out = {}
  if (on('b') !== undefined) out.b = on('b')
  if (on('i') !== undefined) out.i = on('i')
  if (kid(f, 'u')) out.u = true
  if (on('strike')) out.s = true
  const color = wb.color(kid(f, 'color'))
  if (color) out.color = color
  if (pf && pf.attrs.patternType !== 'none') {
    const fill = wb.color(kid(pf, 'bgColor')) ?? wb.color(kid(pf, 'fgColor'))
    if (fill) out.fill = fill
  }
  return out
}

function addDxf (wb, style) {
  const sp = wb.sp
  const font = []
  if (style.bold) font.push(el(sp + 'b'))
  if (style.italic) font.push(el(sp + 'i'))
  if (style.color) font.push(el(sp + 'color', { rgb: 'FF' + hex(style.color) }))
  const kidsOf = []
  if (font.length) kidsOf.push(el(sp + 'font', {}, font))
  if (style.fill) kidsOf.push(el(sp + 'fill', {}, [el(sp + 'patternFill', {}, [el(sp + 'bgColor', { rgb: 'FF' + hex(style.fill) })])]))
  const list = kids(kid(wb.st, 'dxfs'), 'dxf')
  return wb.add('dxfs', 'dxf', el(sp + 'dxf', {}, kidsOf), list)
}

const sqrefOf = (s) => String(s ?? '').split(/\s+/).map(parseRange).filter(Boolean)
const overlaps = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2

// sheetCf describes the sheet's conditional formats for the editor.
function sheetCf (wb, sh) {
  const out = []
  for (const cf of kids(sh.root, 'conditionalFormatting')) {
    for (const r of kids(cf, 'cfRule')) {
      const rule = { sqref: cf.attrs.sqref, type: r.attrs.type, priority: +(r.attrs.priority ?? 0), stop: r.attrs.stopIfTrue === '1' }
      if (r.attrs.operator) rule.op = r.attrs.operator
      if (r.attrs.text !== undefined) rule.text = r.attrs.text
      if (r.attrs.rank) rule.rank = +r.attrs.rank
      if (r.attrs.bottom === '1') rule.bottom = true
      if (r.attrs.percent === '1') rule.percent = true
      if (r.attrs.aboveAverage === '0') rule.below = true
      const f = kids(r, 'formula').map(textContent)
      if (f.length) rule.f = f
      if (r.attrs.dxfId !== undefined) rule.style = dxfOf(wb, +r.attrs.dxfId)
      const scale = kid(r, 'colorScale'); const bar = kid(r, 'dataBar')
      const box = scale ?? bar
      if (box) {
        rule.cfvo = kids(box, 'cfvo').map((v) => ({ type: v.attrs.type, val: v.attrs.val }))
        rule.colors = kids(box, 'color').map((c) => wb.color(c) ?? '638EC6')
      }
      out.push(rule)
    }
  }
  return out.sort((a, b) => a.priority - b.priority)
}

// cfCells works out what each conditional format paints, cell by cell,
// from the cells' calculated values.
function cfCells (wb, sh, rules) {
  const hits = {}
  if (!rules.length) return hits
  const used = sh.usedRange?.() ?? sh.bounds()
  const valueAt = (r, c) => { const cell = sh.cellAt(r, c); return cell ? sh.value(cell) : null }
  const sheets = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s]))
  const names = wb.names()
  const env = (r, c) => ({
    sheet: sh.info.name,
    at: { r, c },
    has: (n) => sheets.has(String(n).toLowerCase()),
    cell: (n, rr, cc) => { const o = wb.sheet(sheets.get(String(n).toLowerCase()).name); const x = o.cellAt(rr, cc); return x ? o.value(x) : null },
    bounds: (n) => wb.sheet(sheets.get(String(n).toLowerCase()).name).bounds(),
    name: (n) => names.find((d) => d.name.toLowerCase() === String(n).toLowerCase())?.ref
  })
  const calc = (src, r0, c0, r, c) => {
    try { return evaluate(parseFormula(shift(src, r - r0, c - c0)), env(r, c)) } catch { return undefined }
  }
  const painted = new Set()
  for (const [k, rule] of rules.entries()) {
    const ranges = sqrefOf(rule.sqref).map((g) => ({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, Math.max(g.r1, used.r2 ?? g.r2)), c2: Math.min(g.c2, Math.max(g.c1, used.c2 ?? g.c2)) }))
    if (!ranges.length) continue
    const top = ranges[0]
    const cells = []
    for (const g of ranges) for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) cells.push([r, c, valueAt(r, c)])
    const nums = cells.map((x) => x[2]).filter((v) => typeof v === 'number')
    const put = (r, c, props) => { const key = `${r},${c}`; hits[key] = { ...props, ...(hits[key] ?? {}) } }
    if (rule.type === 'colorScale' || rule.type === 'dataBar') {
      if (!nums.length) continue
      const sorted = nums.slice().sort((a, b) => a - b)
      const lo = sorted[0]; const hi = sorted[sorted.length - 1]
      const at = (v) => {
        const n = Number(v.val)
        if (v.type === 'min') return lo
        if (v.type === 'max') return hi
        if (v.type === 'percent') return lo + (hi - lo) * n / 100
        if (v.type === 'percentile') { const i = (sorted.length - 1) * n / 100; const a = Math.floor(i); return sorted[a] + (sorted[Math.min(a + 1, sorted.length - 1)] - sorted[a]) * (i - a) }
        if (v.type === 'formula') { const x = calc(v.val ?? '0', top.r1, top.c1, top.r1, top.c1); return typeof x === 'number' ? x : lo }
        return Number.isFinite(n) ? n : lo
      }
      const stops = rule.cfvo.map(at)
      for (const [r, c, v] of cells) {
        if (typeof v !== 'number') continue
        const key = `${r},${c}`
        if (rule.type === 'dataBar') {
          if (hits[key]?.bar) continue
          const a = stops[0]; const b = stops[stops.length - 1]
          const pct = b === a ? 1 : Math.max(0, Math.min(1, (v - a) / (b - a)))
          put(r, c, { bar: [Math.round((0.1 + 0.8 * pct) * 100) / 100, rule.colors[0] ?? '638EC6'] })
        } else {
          if (hits[key]?.scale) continue
          let i = 0
          while (i < stops.length - 2 && v > stops[i + 1]) i++
          const a = stops[i]; const b = stops[i + 1]
          const t = b === a ? 1 : Math.max(0, Math.min(1, (v - a) / (b - a)))
          const mix = (x, y) => [0, 2, 4].map((o) => Math.round(parseInt(x.slice(o, o + 2), 16) * (1 - t) + parseInt(y.slice(o, o + 2), 16) * t).toString(16).padStart(2, '0')).join('').toUpperCase()
          put(r, c, { scale: mix(rule.colors[i] ?? 'FFFFFF', rule.colors[i + 1] ?? rule.colors[i] ?? 'FFFFFF') })
        }
      }
      continue
    }
    let test
    const text = (v) => (v === null || v === undefined ? '' : String(typeof v === 'number' ? fmtGeneral(v) : v)).toLowerCase()
    switch (rule.type) {
      case 'cellIs': test = (r, c, v) => {
        const a = calc(rule.f?.[0] ?? '0', top.r1, top.c1, r, c); const b = calc(rule.f?.[1] ?? '0', top.r1, top.c1, r, c)
        if (v === null || v === undefined || isErr(v) || a === undefined) return false
        const cmp = (x, y) => (typeof x === 'number' && typeof y === 'number' ? x - y : text(x) < text(y) ? -1 : text(x) > text(y) ? 1 : 0)
        const d = cmp(v, a)
        switch (rule.op) {
          case 'greaterThan': return d > 0
          case 'lessThan': return d < 0
          case 'greaterThanOrEqual': return d >= 0
          case 'lessThanOrEqual': return d <= 0
          case 'equal': return d === 0
          case 'notEqual': return d !== 0
          case 'between': return d >= 0 && cmp(v, b) <= 0
          case 'notBetween': return d < 0 || cmp(v, b) > 0
        }
        return false
      }; break
      case 'expression': test = (r, c) => { const x = calc(rule.f?.[0] ?? 'FALSE', top.r1, top.c1, r, c); return x === true || (typeof x === 'number' && x !== 0) }; break
      case 'containsText': test = (r, c, v) => text(v).includes(String(rule.text ?? '').toLowerCase()); break
      case 'notContainsText': test = (r, c, v) => !text(v).includes(String(rule.text ?? '').toLowerCase()); break
      case 'beginsWith': test = (r, c, v) => text(v).startsWith(String(rule.text ?? '').toLowerCase()); break
      case 'endsWith': test = (r, c, v) => text(v).endsWith(String(rule.text ?? '').toLowerCase()); break
      case 'containsBlanks': test = (r, c, v) => text(v).trim() === ''; break
      case 'notContainsBlanks': test = (r, c, v) => text(v).trim() !== ''; break
      case 'containsErrors': test = (r, c, v) => !!isErr(v); break
      case 'notContainsErrors': test = (r, c, v) => !isErr(v); break
      case 'duplicateValues': case 'uniqueValues': {
        const count = new Map()
        for (const [, , v] of cells) if (text(v) !== '') count.set(text(v), (count.get(text(v)) ?? 0) + 1)
        test = (r, c, v) => text(v) !== '' && (rule.type === 'duplicateValues' ? count.get(text(v)) > 1 : count.get(text(v)) === 1)
        break
      }
      case 'top10': {
        const sorted = nums.slice().sort((a, b) => (rule.bottom ? a - b : b - a))
        const n = rule.percent ? Math.max(1, Math.floor(sorted.length * (rule.rank ?? 10) / 100)) : (rule.rank ?? 10)
        const edge = sorted[Math.min(n, sorted.length) - 1]
        test = (r, c, v) => typeof v === 'number' && (rule.bottom ? v <= edge : v >= edge)
        break
      }
      case 'aboveAverage': {
        const avg = nums.reduce((s, x) => s + x, 0) / (nums.length || 1)
        test = (r, c, v) => typeof v === 'number' && (rule.below ? v < avg : v > avg)
        break
      }
      default: continue
    }
    for (const [r, c, v] of cells) {
      const key = `${r},${c}`
      if (painted.has(key + '|stop')) continue
      if (!test(r, c, v)) continue
      put(r, c, rule.style ?? {})
      if (rule.stop) painted.add(key + '|stop')
    }
    void k
  }
  return hits
}

// addCf puts a new rule on top of the sheet's conditional formats.
function addCf (wb, sh, e, where) {
  const ranges = sqrefOf(e.range)
  if (!ranges.length) throw new Error(`${where}: bad range`)
  const rule = e.rule ?? {}
  const p = sh.p
  for (const r of descendants(sh.root, (n) => local(n) === 'cfRule')) r.attrs.priority = String(+(r.attrs.priority ?? 0) + 1)
  const top = cellName(ranges[0].r1, ranges[0].c1)
  const attrs = { type: rule.type, priority: '1' }
  const body = []
  const style = () => { attrs.dxfId = String(addDxf(wb, rule.style ?? { fill: 'FFC7CE', color: '9C0006' })) }
  const q = (s) => '"' + String(s).replace(/"/g, '""') + '"'
  const lit = (v) => { const s = String(v ?? '').replace(/^=/, '').replace(/^([+-]?\d{1,3}(,\d{3})+(\.\d+)?)$/, (n) => n.replace(/,/g, '')); return s === '' ? '0' : /^[+-]?(\d+\.?\d*|\.\d+)$/.test(s) || /^\$?[A-Za-z]{1,3}\$?\d+$/.test(s) || String(v).startsWith('=') ? s : q(s) }
  switch (rule.type) {
    case 'cellIs':
      if (!['greaterThan', 'lessThan', 'between', 'notBetween', 'equal', 'notEqual', 'greaterThanOrEqual', 'lessThanOrEqual'].includes(rule.op)) throw new Error(`${where}: unknown operator ${rule.op}`)
      attrs.operator = rule.op; style()
      body.push(el(p + 'formula', {}, [text(lit(rule.f?.[0]))]))
      if (/etween/.test(rule.op)) body.push(el(p + 'formula', {}, [text(lit(rule.f?.[1]))]))
      break
    case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
      const t = String(rule.text ?? '')
      attrs.operator = rule.type; attrs.text = t; style()
      const f = { containsText: `NOT(ISERROR(SEARCH(${q(t)},${top})))`, notContainsText: `ISERROR(SEARCH(${q(t)},${top}))`, beginsWith: `LEFT(${top},LEN(${q(t)}))=${q(t)}`, endsWith: `RIGHT(${top},LEN(${q(t)}))=${q(t)}` }[rule.type]
      body.push(el(p + 'formula', {}, [text(f)]))
      break
    }
    case 'duplicateValues': case 'uniqueValues': style(); break
    case 'top10':
      attrs.rank = String(Math.max(1, Math.min(1000, +rule.rank || 10)))
      if (rule.bottom) attrs.bottom = '1'
      if (rule.percent) attrs.percent = '1'
      style(); break
    case 'aboveAverage':
      if (rule.below) attrs.aboveAverage = '0'
      style(); break
    case 'expression':
      style(); body.push(el(p + 'formula', {}, [text(String(rule.f?.[0] ?? '').replace(/^=/, ''))])); break
    case 'colorScale': {
      const colors = (rule.colors ?? ['F8696B', 'FFEB84', '63BE7B']).map(hex)
      const cfvo = colors.length === 3 ? [el(p + 'cfvo', { type: 'min' }), el(p + 'cfvo', { type: 'percentile', val: '50' }), el(p + 'cfvo', { type: 'max' })] : [el(p + 'cfvo', { type: 'min' }), el(p + 'cfvo', { type: 'max' })]
      body.push(el(p + 'colorScale', {}, [...cfvo, ...colors.map((c) => el(p + 'color', { rgb: 'FF' + c }))]))
      break
    }
    case 'dataBar':
      body.push(el(p + 'dataBar', {}, [el(p + 'cfvo', { type: 'min' }), el(p + 'cfvo', { type: 'max' }), el(p + 'color', { rgb: 'FF' + hex(rule.color ?? '638EC6') })]))
      break
    default: throw new Error(`${where}: unknown rule type ${rule.type}`)
  }
  const node = el(p + 'conditionalFormatting', { sqref: ranges.map((g) => g.r1 === g.r2 && g.c1 === g.c2 ? cellName(g.r1, g.c1) : `${cellName(g.r1, g.c1)}:${cellName(g.r2, g.c2)}`).join(' ') }, [el(p + 'cfRule', attrs, body)])
  // Conditional formats sit together, in the schema's place.
  const last = kids(sh.root, 'conditionalFormatting').pop()
  if (last) sh.root.kids.splice(sh.root.kids.indexOf(last) + 1, 0, node)
  else insertOrdered(sh.root, node, WS_ORDER.map((n) => p + n))
  wb.pkg.touch(sh.part)
  void CF_ORDER
}

// clearCf drops the conditional formats touching range, or all of them.
function clearCf (wb, sh, e) {
  const rg = e.range ? parseRange(e.range) : undefined
  for (const cf of kids(sh.root, 'conditionalFormatting')) {
    if (!rg || sqrefOf(cf.attrs.sqref).some((g) => overlaps(g, rg))) remove(sh.root, cf)
  }
  wb.pkg.touch(sh.part)
}

// deleteCf drops one rule, found by its place in priority order (as the
// editor lists them) and checked against the range and type it showed.
function deleteCf (wb, sh, e, where) {
  const all = []
  for (const cf of kids(sh.root, 'conditionalFormatting')) for (const r of kids(cf, 'cfRule')) all.push({ cf, r, priority: +(r.attrs.priority ?? 0) })
  all.sort((a, b) => a.priority - b.priority)
  const hit = all[+e.index]
  if (!hit || (e.sqref && hit.cf.attrs.sqref !== e.sqref) || (e.type && hit.r.attrs.type !== e.type)) throw new Error(`${where}: the rule changed meanwhile`)
  remove(hit.cf, hit.r)
  if (!kids(hit.cf, 'cfRule').length) remove(sh.root, hit.cf)
  wb.pkg.touch(sh.part)
}

// ------------------------------------------------------------- data validation

const DV_TYPES = ['list', 'whole', 'decimal', 'textLength', 'date', 'custom']

function sheetDv (wb, sh) {
  const box = kid(sh.root, 'dataValidations')
  return kids(box, 'dataValidation').map((d) => {
    const v = { sqref: d.attrs.sqref, type: d.attrs.type ?? 'any', op: d.attrs.operator ?? 'between', blank: d.attrs.allowBlank === '1' }
    const f1 = textContent(kid(d, 'formula1') ?? { kids: [] }); const f2 = textContent(kid(d, 'formula2') ?? { kids: [] })
    if (f1) v.f1 = f1
    if (f2) v.f2 = f2
    if (d.attrs.error) v.error = d.attrs.error
    if (d.attrs.errorTitle) v.errorTitle = d.attrs.errorTitle
    if (d.attrs.prompt) v.prompt = d.attrs.prompt
    if (d.attrs.promptTitle) v.promptTitle = d.attrs.promptTitle
    if (d.attrs.errorStyle) v.style = d.attrs.errorStyle
    // showDropDown="1" means the arrow is hidden (the name is backwards).
    if (d.attrs.showDropDown === '1') v.noArrow = true
    const val = (f) => {
      if (f === undefined) return undefined
      if (/^"/.test(f)) return f.slice(1, -1).replace(/""/g, '"')
      try { return evaluate(parseFormula(f), { sheet: sh.info.name, has: () => true, cell: (n, r, c) => { const o = wb.sheet(n); const x = o.cellAt(r, c); return x ? o.value(x) : null }, bounds: (n) => wb.sheet(n).bounds(), name: (n) => wb.names().find((x) => x.name.toLowerCase() === n.toLowerCase())?.ref }) } catch { return undefined }
    }
    if (v.type === 'list' && f1) {
      if (/^"/.test(f1)) v.items = f1.slice(1, -1).replace(/""/g, '"').split(',').map((s) => s.trim())
      else {
        try {
          const tree = parseFormula(f1)
          const env = { sheet: sh.info.name, has: (n) => wb.sheets.some((s) => s.name.toLowerCase() === String(n).toLowerCase()), cell: (n, r, c) => { const o = wb.sheet(n); return o.display(r, c) }, bounds: (n) => wb.sheet(n).bounds(), name: (n) => wb.names().find((x) => x.name.toLowerCase() === n.toLowerCase())?.ref }
          const items = []
          const walk = (t) => { if (t.ref) { const b = t.ref.lastIndexOf('!'); const sheet = b >= 0 ? t.ref.slice(0, b).replace(/^'|'$/g, '') : sh.info.name; const rg = parseRange(t.ref.slice(b + 1).replace(/\$/g, '')); if (rg) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) items.push(env.cell(sheet, r, c)) } else if (t.name) { const src = env.name(t.name); if (src) walk(parseFormula(src)) } }
          walk(tree)
          v.items = items.filter((s) => s !== '' && s !== null && s !== undefined).map(String)
        } catch {}
      }
    } else {
      const a = val(f1); const b = val(f2)
      if (typeof a === 'number') v.v1 = a
      if (typeof b === 'number') v.v2 = b
    }
    return v
  })
}

function setDv (wb, sh, e, where) {
  const ranges = sqrefOf(e.range)
  if (!ranges.length) throw new Error(`${where}: bad range`)
  clearDv(wb, sh, e)
  const rule = e.rule
  if (!rule) return
  if (!DV_TYPES.includes(rule.type)) throw new Error(`${where}: unknown validation type ${rule.type}`)
  const p = sh.p
  let box = kid(sh.root, 'dataValidations')
  if (!box) { box = el(p + 'dataValidations', { count: '0' }); insertOrdered(sh.root, box, WS_ORDER.map((n) => p + n)) }
  const attrs = { type: rule.type, allowBlank: rule.blank === false ? '0' : '1', showInputMessage: '1', showErrorMessage: '1' }
  if (rule.type !== 'list' && rule.type !== 'custom') attrs.operator = rule.op ?? 'between'
  if (rule.error) attrs.error = String(rule.error)
  if (rule.prompt) attrs.prompt = String(rule.prompt)
  if (rule.style && rule.style !== 'stop') attrs.errorStyle = rule.style
  attrs.sqref = ranges.map((g) => g.r1 === g.r2 && g.c1 === g.c2 ? cellName(g.r1, g.c1) : `${cellName(g.r1, g.c1)}:${cellName(g.r2, g.c2)}`).join(' ')
  const f = (s) => String(s ?? '').replace(/^=/, '')
  const body = []
  if (rule.type === 'list') {
    const src = String(rule.list ?? '')
    if (!src.trim()) throw new Error(`${where}: list needs items`)
    const ref = src.startsWith('=') ? src.slice(1) : undefined
    if (!ref && src.length > 255) throw new Error(`${where}: the list is longer than 255 characters; put the items in cells and refer to them`)
    body.push(el(p + 'formula1', {}, [text(ref ?? '"' + src.split(/[,，]/).map((s) => s.trim()).filter(Boolean).join(',').replace(/"/g, '""') + '"')]))
  } else {
    body.push(el(p + 'formula1', {}, [text(f(rule.f1))]))
    if (/etween/.test(attrs.operator ?? '')) body.push(el(p + 'formula2', {}, [text(f(rule.f2))]))
  }
  box.kids.push(el(p + 'dataValidation', attrs, body))
  box.attrs.count = String(kids(box, 'dataValidation').length)
  wb.pkg.touch(sh.part)
}

// clearDv takes range out of every validation; validations left empty go.
function clearDv (wb, sh, e) {
  const box = kid(sh.root, 'dataValidations')
  if (!box) return
  const rg = e.range ? parseRange(e.range) : undefined
  for (const d of kids(box, 'dataValidation')) {
    if (!rg) { remove(box, d); continue }
    const parts = []
    for (const g of sqrefOf(d.attrs.sqref)) {
      if (!overlaps(g, rg)) { parts.push(g); continue }
      // Keep the pieces of g outside rg.
      if (g.r1 < rg.r1) parts.push({ ...g, r2: rg.r1 - 1 })
      if (g.r2 > rg.r2) parts.push({ ...g, r1: rg.r2 + 1 })
      const r1 = Math.max(g.r1, rg.r1); const r2 = Math.min(g.r2, rg.r2)
      if (g.c1 < rg.c1) parts.push({ r1, r2, c1: g.c1, c2: rg.c1 - 1 })
      if (g.c2 > rg.c2) parts.push({ r1, r2, c1: rg.c2 + 1, c2: g.c2 })
    }
    if (!parts.length) remove(box, d)
    else d.attrs.sqref = parts.map((g) => g.r1 === g.r2 && g.c1 === g.c2 ? cellName(g.r1, g.c1) : `${cellName(g.r1, g.c1)}:${cellName(g.r2, g.c2)}`).join(' ')
  }
  if (!kids(box, 'dataValidation').length) remove(sh.root, box)
  else box.attrs.count = String(kids(box, 'dataValidation').length)
  wb.pkg.touch(sh.part)
}

// ------------------------------------------------------------- comments

const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function commentsOf (wb, sh, create) {
  const rel = wb.pkg.relsOfType(sh.part, 'comments')[0]
  if (rel?.part && wb.pkg.has(rel.part)) return { part: rel.part, root: wb.pkg.xml(rel.part).root }
  if (!create) return undefined
  const part = wb.pkg.uniqueName('xl/comments#.xml')
  wb.pkg.setXml(part, doc(el('comments', { xmlns: MAIN }, [el('authors'), el('commentList')])))
  wb.pkg.setOverride(part, CT_COMMENTS)
  wb.pkg.addRel(sh.part, REL('comments'), part)
  return { part, root: wb.pkg.xml(part).root }
}

// sheetNotes reads the sheet's (legacy) comments as { "r,c": { a, t } }.
function sheetNotes (wb, sh) {
  const cm = commentsOf(wb, sh, false)
  if (!cm) return undefined
  const authors = kids(kid(cm.root, 'authors'), 'author').map(textContent)
  const out = {}
  for (const c of kids(kid(cm.root, 'commentList'), 'comment')) {
    const p = parseCell(c.attrs.ref ?? '')
    if (!p) continue
    const a = authors[+(c.attrs.authorId ?? 0)] ?? ''
    let t = descendants(kid(c, 'text'), (n) => local(n) === 't').map(textContent).join('')
    if (a && t.startsWith(a + ':')) t = t.slice(a.length + 1).replace(/^\n/, '')
    out[`${p.r},${p.c}`] = { a, t }
  }
  return Object.keys(out).length ? out : undefined
}

// writeVml rebuilds the legacy drawing that shows comment boxes, unless
// it also holds form controls Quire does not draw.
function writeVml (wb, sh, refs) {
  const rel = wb.pkg.relsOfType(sh.part, 'vmlDrawing')[0]
  let part = rel?.part
  if (part && wb.pkg.has(part)) {
    const old = String(wb.pkg.get(part))
    const types = [...old.matchAll(/ObjectType="(\w+)"/g)].map((m) => m[1])
    if (types.some((t) => t !== 'Note')) return
  }
  const leg = kid(sh.root, 'legacyDrawing')
  if (!refs.length) {
    if (rel) { wb.pkg.removeRel(sh.part, rel.id); if (part) { wb.pkg.delete(part) } }
    if (leg) remove(sh.root, leg)
    return
  }
  const shapes = refs.map((ref, i) => {
    const p = parseCell(ref)
    return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144pt;height:79pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${p.c}, 15, ${p.r - 1}, 2, ${p.c + 2}, 15, ${p.r + 3}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${p.r - 1}</x:Row><x:Column>${p.c - 1}</x:Column></x:ClientData></v:shape>`
  }).join('')
  const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`
  if (!part || !wb.pkg.has(part)) {
    part = wb.pkg.uniqueName('xl/drawings/vmlDrawing#.vml')
    wb.pkg.ensureDefault('vml', CT_VML)
  }
  wb.pkg.setBytes(part, Buffer.from(vml, 'utf8'))
  let rid = rel?.id
  if (!rel) rid = wb.pkg.addRel(sh.part, REL('vmlDrawing'), part)
  if (!sh.root.attrs['xmlns:r']) sh.root.attrs['xmlns:r'] = R_NS
  if (leg) leg.attrs['r:id'] = rid
  else insertOrdered(sh.root, el(sh.p + 'legacyDrawing', { 'r:id': rid }), WS_ORDER.map((n) => sh.p + n))
}

function setNote (wb, sh, e, where) {
  const p = parseCell(e.cell ?? '')
  if (!p) throw new Error(`${where}: "${e.cell}" is not a cell`)
  const ref = cellName(p.r, p.c)
  const cm = commentsOf(wb, sh, e.text !== null && e.text !== undefined && e.text !== '')
  if (!cm) return
  const list = kid(cm.root, 'commentList') ?? (() => { const n = el('commentList'); cm.root.kids.push(n); return n })()
  const pre = cm.root.name.includes(':') ? cm.root.name.split(':')[0] + ':' : ''
  list.kids = list.kids.filter((c) => !(isEl(c) && local(c) === 'comment' && parseCell(c.attrs.ref ?? '')?.r === p.r && parseCell(c.attrs.ref ?? '')?.c === p.c))
  if (e.text !== null && e.text !== undefined && e.text !== '') {
    const authors = kid(cm.root, 'authors') ?? (() => { const n = el(pre + 'authors'); cm.root.kids.unshift(n); return n })()
    const name = String(e.author || 'Quire')
    let id = kids(authors, 'author').findIndex((a) => textContent(a) === name)
    if (id < 0) { authors.kids.push(el(pre + 'author', {}, [text(name)])); id = kids(authors, 'author').length - 1 }
    const rpr = (b) => el(pre + 'rPr', {}, [...(b ? [el(pre + 'b')] : []), el(pre + 'sz', { val: '9' }), el(pre + 'color', { indexed: '81' }), el(pre + 'rFont', { val: 'Tahoma' }), el(pre + 'family', { val: '2' })])
    list.kids.push(el(pre + 'comment', { ref, authorId: String(id) }, [el(pre + 'text', {}, [
      el(pre + 'r', {}, [rpr(true), el(pre + 't', {}, [text(name + ':')])]),
      el(pre + 'r', {}, [rpr(false), el(pre + 't', { 'xml:space': 'preserve' }, [text('\n' + String(e.text))])])
    ])]))
  }
  wb.pkg.touch(cm.part)
  const refs = kids(list, 'comment').map((c) => c.attrs.ref)
  if (!refs.length) {
    const rel = wb.pkg.relsOfType(sh.part, 'comments')[0]
    if (rel) wb.pkg.removeRel(sh.part, rel.id)
    wb.pkg.delete(cm.part)
    wb.pkg.removeOverride(cm.part)
  }
  writeVml(wb, sh, refs)
  wb.pkg.touch(sh.part)
}

// ------------------------------------------------------------- hyperlinks

function sheetLinks (wb, sh) {
  const out = {}
  for (const h of kids(kid(sh.root, 'hyperlinks'), 'hyperlink')) {
    const rg = parseRange(h.attrs.ref ?? '')
    if (!rg) continue
    const rid = Object.entries(h.attrs).find(([k]) => k.endsWith(':id'))?.[1]
    const target = rid ? wb.pkg.rels(sh.part).find((r) => r.id === rid)?.target : undefined
    const url = target ?? (h.attrs.location ? '#' + h.attrs.location : undefined)
    if (!url) continue
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 500); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) out[`${r},${c}`] = url
  }
  return Object.keys(out).length ? out : undefined
}

function setLink (wb, sh, e, where) {
  const p = parseCell(e.cell ?? '')
  if (!p) throw new Error(`${where}: "${e.cell}" is not a cell`)
  let box = kid(sh.root, 'hyperlinks')
  for (const h of kids(box, 'hyperlink')) {
    const rg = parseRange(h.attrs.ref ?? '')
    if (!rg || p.r < rg.r1 || p.r > rg.r2 || p.c < rg.c1 || p.c > rg.c2) continue
    const rid = Object.entries(h.attrs).find(([k]) => k.endsWith(':id'))?.[1]
    if (rid) wb.pkg.removeRel(sh.part, rid)
    remove(box, h)
  }
  const url = String(e.url ?? '').trim()
  if (url) {
    if (!/^(https?:|mailto:|file:|#)/i.test(url)) throw new Error(`${where}: 链接要以 http://、https://、mailto: 开头，或用 # 跳到工作表位置`)
    if (!box) { box = el(sh.p + 'hyperlinks'); insertOrdered(sh.root, box, WS_ORDER.map((n) => sh.p + n)) }
    const attrs = { ref: cellName(p.r, p.c) }
    if (url.startsWith('#')) attrs.location = url.slice(1)
    else {
      if (!sh.root.attrs['xmlns:r']) sh.root.attrs['xmlns:r'] = R_NS
      attrs['r:id'] = wb.pkg.addRel(sh.part, REL('hyperlink'), url, true)
    }
    if (e.tip) attrs.tooltip = String(e.tip)
    box.kids.push(el(sh.p + 'hyperlink', attrs))
  }
  if (box && !kids(box, 'hyperlink').length) remove(sh.root, box)
  wb.pkg.touch(sh.part)
}

// ------------------------------------------------------------- duplicates

// dedupe removes rows of a range whose values in cols repeat an earlier
// row's, moving the rest up inside the range; it returns how many went.
function dedupe (sh, rg, cols, header, wrote, where) {
  if (!rg) throw new Error(`${where}: range must be like A1:D20`)
  const b = sh.bounds()
  const r1 = header ? rg.r1 + 1 : rg.r1
  const r2 = Math.min(rg.r2, b.r2)
  const c2 = Math.min(rg.c2, Math.max(b.c2, rg.c1))
  const keyCols = (Array.isArray(cols) && cols.length ? cols : Array.from({ length: c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i)).map(Number).filter((c) => c >= rg.c1 && c <= rg.c2)
  if (r2 <= r1 || !keyCols.length) return 0
  for (const ref of sh.merges()) {
    const m = parseRange(ref)
    if (m && !(m.r2 < r1 || m.r1 > r2 || m.c2 < rg.c1 || m.c1 > c2)) throw new Error(`${where}: 区域里有合并单元格，请先取消合并`)
  }
  for (const si of [...sh.shared.keys()]) sh.unshare(si)
  const seen = new Set()
  const keep = []
  let gone = 0
  for (let r = r1; r <= r2; r++) {
    const key = JSON.stringify(keyCols.map((c) => { const x = sh.cellAt(r, c); const v = x ? sh.value(x) : null; return typeof v === 'string' ? v.toLowerCase() : v }))
    const cells = []
    for (let c = rg.c1; c <= c2; c++) cells.push(dropCell(sh, r, c))
    if (seen.has(key)) { gone++; continue }
    seen.add(key)
    keep.push({ r, cells })
  }
  keep.forEach((row, i) => row.cells.forEach((cell, j) => { if (cell) placeCopy(sh, cell, r1 + i, rg.c1 + j, r1 + i - row.r, 0, wrote) }))
  for (let r = r1 + keep.length; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) wrote.add(`${sh.info.name}!${r},${c}`)
  sh.wb.pkg.touch(sh.part)
  return gone
}

// ------------------------------------------------------------- outline

function setOutline (sh, e, where) {
  const a = Math.max(1, +e.from); const b = Math.max(a, +(e.to ?? e.from))
  if (!Number.isFinite(a) || !Number.isFinite(b) || b - a > 100000) throw new Error(`${where}: bad range`)
  const step = +e.delta < 0 ? -1 : 1
  const fix = (node) => {
    const lv = Math.max(0, Math.min(7, +(node.attrs.outlineLevel ?? 0) + step))
    if (lv) node.attrs.outlineLevel = String(lv); else { delete node.attrs.outlineLevel; delete node.attrs.collapsed }
  }
  if (e.axis === 'c') sh.colAttrs(a, b, fix)
  else for (let r = a; r <= b; r++) fix(sh.row(r))
  // The sheet records its deepest level so Excel draws enough buttons.
  const deepest = e.axis === 'c'
    ? Math.max(0, ...kids(kid(sh.root, 'cols'), 'col').map((c) => +(c.attrs.outlineLevel ?? 0)))
    : Math.max(0, ...[...sh.rows.values()].map((r) => +(r.attrs.outlineLevel ?? 0)))
  let fmt = kid(sh.root, 'sheetFormatPr')
  if (!fmt) { fmt = el(sh.p + 'sheetFormatPr', { defaultRowHeight: '15' }); insertOrdered(sh.root, fmt, WS_ORDER.map((n) => sh.p + n)) }
  const attr = e.axis === 'c' ? 'outlineLevelCol' : 'outlineLevelRow'
  if (deepest) fmt.attrs[attr] = String(deepest); else delete fmt.attrs[attr]
  sh.wb.pkg.touch(sh.part)
}

function sheetOutline (sh) {
  const r = {}
  for (const [n, row] of sh.rows) if (+(row.attrs.outlineLevel ?? 0)) r[n] = +row.attrs.outlineLevel
  const c = {}
  for (const col of kids(kid(sh.root, 'cols'), 'col')) {
    const lv = +(col.attrs.outlineLevel ?? 0)
    if (lv) for (let k = +col.attrs.min; k <= Math.min(+col.attrs.max, +col.attrs.min + 1000); k++) c[k] = lv
  }
  return Object.keys(r).length || Object.keys(c).length ? { r, c } : undefined
}

// ------------------------------------------------------------- protection

// legacyHash is Excel's old 16-bit sheet password hash.
function legacyHash (pw) {
  let h = 0
  const chars = [...String(pw)].map((ch) => ch.charCodeAt(0) & 0x7fff)
  for (let i = chars.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff)
    h ^= chars[i]
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff)
  h ^= chars.length
  h ^= 0xce4b
  return h.toString(16).toUpperCase().padStart(4, '0')
}

function protectSheet (wb, sh, e) {
  let sp = kid(sh.root, 'sheetProtection')
  if (sp) remove(sh.root, sp)
  sp = el(sh.p + 'sheetProtection', { sheet: '1', objects: '1', scenarios: '1', ...(e.password ? { password: legacyHash(e.password) } : {}) })
  insertOrdered(sh.root, sp, WS_ORDER.map((n) => sh.p + n))
  wb.pkg.touch(sh.part)
}

function unprotectSheet (wb, sh, e) {
  const sp = kid(sh.root, 'sheetProtection')
  if (!sp) return
  const pw = String(e.password ?? '')
  let ok = true
  if (sp.attrs.password) ok = legacyHash(pw) === sp.attrs.password.toUpperCase().padStart(4, '0')
  else if (sp.attrs.hashValue) {
    const algo = String(sp.attrs.algorithmName ?? 'SHA-512').toLowerCase().replace('-', '')
    let h = crypto.createHash(algo).update(Buffer.concat([Buffer.from(sp.attrs.saltValue ?? '', 'base64'), Buffer.from(pw, 'utf16le')])).digest()
    const n = +(sp.attrs.spinCount ?? 0)
    for (let i = 0; i < n; i++) { const b = Buffer.alloc(4); b.writeUInt32LE(i); h = crypto.createHash(algo).update(Buffer.concat([h, b])).digest() }
    ok = h.toString('base64') === sp.attrs.hashValue
  }
  if (!ok) throw new Error('密码不正确')
  remove(sh.root, sp)
  wb.pkg.touch(sh.part)
}

// ------------------------------------------------------------- page setup

const PAPERS = { A4: '9', A3: '8', A5: '11', Letter: '1', Legal: '5', B5: '13' }
const MARGINS = { normal: [0.7, 0.7, 0.75, 0.75, 0.3, 0.3], narrow: [0.25, 0.25, 0.75, 0.75, 0.3, 0.3], wide: [1, 1, 1, 1, 0.5, 0.5] }

function sheetPage (wb, sh) {
  const ps = kid(sh.root, 'pageSetup'); const pm = kid(sh.root, 'pageMargins'); const po = kid(sh.root, 'printOptions')
  const fit = kid(kid(sh.root, 'sheetPr'), 'pageSetUpPr')?.attrs.fitToPage === '1'
  const idx = wb.sheets.indexOf(sh.info)
  const dn = (n) => wb.names().find((d) => d.name.toLowerCase() === n && d.sheet === idx)?.ref
  const m = pm ? ['left', 'right', 'top', 'bottom', 'header', 'footer'].map((k) => +(pm.attrs[k] ?? 0)) : null
  const margins = m ? Object.entries(MARGINS).find(([, v]) => v.every((x, i) => Math.abs(x - m[i]) < 0.01))?.[0] ?? 'custom' : 'normal'
  return {
    orientation: ps?.attrs.orientation ?? 'portrait',
    paper: Object.entries(PAPERS).find(([, v]) => v === (ps?.attrs.paperSize ?? '9'))?.[0] ?? 'A4',
    fit: fit ? { w: +(ps?.attrs.fitToWidth ?? 1), h: +(ps?.attrs.fitToHeight ?? 1) } : undefined,
    margins,
    grid: po?.attrs.gridLines === '1',
    area: dn('_xlnm.print_area'),
    titles: dn('_xlnm.print_titles')
  }
}

function setPage (wb, sh, e, where) {
  const p = sh.p
  const order = WS_ORDER.map((n) => p + n)
  if (e.orientation !== undefined || e.paper !== undefined || e.fit !== undefined) {
    let ps = kid(sh.root, 'pageSetup')
    if (!ps) { ps = el(p + 'pageSetup', { paperSize: '9', orientation: 'portrait' }); insertOrdered(sh.root, ps, order) }
    // A page setup that points at printer settings this file does not keep is dropped.
    if (e.orientation !== undefined) ps.attrs.orientation = e.orientation === 'landscape' ? 'landscape' : 'portrait'
    if (e.paper !== undefined) { if (!PAPERS[e.paper]) throw new Error(`${where}: unknown paper ${e.paper}`); ps.attrs.paperSize = PAPERS[e.paper] }
    if (e.fit !== undefined) {
      let pr = kid(sh.root, 'sheetPr')
      if (!pr) { pr = el(p + 'sheetPr'); insertOrdered(sh.root, pr, order) }
      let su = kid(pr, 'pageSetUpPr')
      if (e.fit) {
        if (!su) { su = el(p + 'pageSetUpPr'); insertOrdered(pr, su, ['tabColor', 'outlinePr', 'pageSetUpPr'].map((n) => p + n)) }
        su.attrs.fitToPage = '1'
        ps.attrs.fitToWidth = String(e.fit.w ?? 1)
        ps.attrs.fitToHeight = String(e.fit.h ?? 0)
      } else {
        if (su) { delete su.attrs.fitToPage; if (!Object.keys(su.attrs).length) remove(pr, su) }
        delete ps.attrs.fitToWidth; delete ps.attrs.fitToHeight
        if (!pr.kids.length && !Object.keys(pr.attrs).length) remove(sh.root, pr)
      }
    }
  }
  if (e.margins !== undefined) {
    const m = MARGINS[e.margins]
    if (!m) throw new Error(`${where}: unknown margins ${e.margins}`)
    let pm = kid(sh.root, 'pageMargins')
    if (!pm) { pm = el(p + 'pageMargins'); insertOrdered(sh.root, pm, order) }
    ;['left', 'right', 'top', 'bottom', 'header', 'footer'].forEach((k, i) => { pm.attrs[k] = String(m[i]) })
  }
  if (e.grid !== undefined) {
    let po = kid(sh.root, 'printOptions')
    if (e.grid) { if (!po) { po = el(p + 'printOptions'); insertOrdered(sh.root, po, order) } po.attrs.gridLines = '1' } else if (po) { delete po.attrs.gridLines; if (!Object.keys(po.attrs).length) remove(sh.root, po) }
  }
  const idx = wb.sheets.indexOf(sh.info)
  const q = /^[A-Za-z_][\w.]*$/.test(sh.info.name) ? sh.info.name : `'${sh.info.name.replace(/'/g, "''")}'`
  if (e.area !== undefined) {
    if (e.area) {
      const rg = parseRange(String(e.area).replace(/\$/g, '').replace(/^.*!/, ''))
      if (!rg) throw new Error(`${where}: bad print area`)
      wb.setName('_xlnm.Print_Area', `${q}!$${colName(rg.c1)}$${rg.r1}:$${colName(rg.c2)}$${rg.r2}`, idx)
    } else wb.setName('_xlnm.Print_Area', null, idx)
  }
  if (e.titles !== undefined) {
    if (e.titles) {
      const m = /^\$?(\d+)(?::\$?(\d+))?$/.exec(String(e.titles).replace(/^.*!/, '').trim())
      if (!m) throw new Error(`${where}: 标题行写成 1:1 或 1:2`)
      wb.setName('_xlnm.Print_Titles', `${q}!$${m[1]}:$${m[2] ?? m[1]}`, idx)
    } else wb.setName('_xlnm.Print_Titles', null, idx)
  }
  wb.pkg.touch(sh.part)
}

// filterValues reads the values a filter column keeps ('' stands for
// blanks); null when it filters some other way (conditions, colors).
function filterValues (fc) {
  const f = kid(fc, 'filters')
  if (!f) return null
  const out = kids(f, 'filter').map((x) => x.attrs.val ?? '')
  if (f.attrs.blank === '1' || f.attrs.blank === 'true') out.push('')
  return out
}

const tooLarge = (rg) => (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 200000

// setBorders draws borders the way Excel's border menu does: all lines,
// the outline (thin or thick), one edge of the range, or none.
function setBorders (wb, sh, rg, kind, color, where) {
  if (!rg) throw new Error(`${where}: range must be like A1:C9`)
  if (tooLarge(rg)) throw new Error(`${where}: the range is too large`)
  if (!['all', 'outside', 'thick', 'top', 'bottom', 'left', 'right', 'none'].includes(kind)) throw new Error(`${where}: kind must be all, outside, thick, top, bottom, left, right or none`)
  const st = kind === 'thick' ? 'medium' : 'thin'
  for (let r = rg.r1; r <= rg.r2; r++) {
    for (let c = rg.c1; c <= rg.c2; c++) {
      const edge = { top: r === rg.r1, bottom: r === rg.r2, left: c === rg.c1, right: c === rg.c2 }
      const sides = {}
      if (kind === 'none') { if (!sh.cellAt(r, c)) continue; Object.assign(sides, { top: null, bottom: null, left: null, right: null }) } else if (kind === 'all') Object.assign(sides, { top: st, bottom: st, left: st, right: st })
      else if (kind === 'outside' || kind === 'thick') { for (const n of ['top', 'bottom', 'left', 'right']) if (edge[n]) sides[n] = st } else if (edge[kind]) sides[kind] = st
      if (!Object.keys(sides).length) continue
      const cell = sh.cell(r, c)
      cell.attrs.s = String(wb.deriveSides(+(cell.attrs.s ?? 0), sides, color || '000000'))
      if (cell.attrs.s === '0') delete cell.attrs.s
    }
  }
  wb.pkg.touch(sh.part)
}

// placeCopy writes a copy of a cell somewhere else, moving the relative
// references of its formula by (dr, dc). Shared formulas must have been
// unshared first.
function placeCopy (sh, src, r, c, dr, dc, wrote) {
  const dst = sh.cell(r, c)
  dst.attrs = { ...src.attrs, r: cellName(r, c) }
  dst.kids = src.kids.map(clone)
  const f = kid(dst, 'f')
  if (f) {
    const t = textContent(f)
    if (t) f.kids = [text(shift(t, dr, dc))]
    delete f.attrs.ref
    if (f.attrs.t === 'array') delete f.attrs.t
    dst.kids = dst.kids.filter((k) => local(k) !== 'v')
    wrote.add(`${sh.info.name}!${r},${c}`)
  }
  sh.wb.pkg.touch(sh.part)
  return dst
}

function dropCell (sh, r, c) {
  const cell = sh.cellAt(r, c)
  if (!cell) return undefined
  remove(sh.rows.get(r), cell)
  sh.cells.delete(`${r},${c}`)
  return cell
}

// autofill extends the range from into to, like dragging the fill handle:
// numbers continue their trend, a single date counts up by days, text that
// ends in a number counts up, formulas move their references, and
// everything else repeats. Styles come along.
function autofill (sh, from, to, wrote, where) {
  if (!from || !to) throw new Error(`${where}: from and to must be ranges like A1:B2`)
  const down = to.r2 > from.r2; const up = to.r1 < from.r1; const right = to.c2 > from.c2; const left = to.c1 < from.c1
  const vertical = down || up
  const fits = vertical ? to.c1 === from.c1 && to.c2 === from.c2 && !(right || left) : to.r1 === from.r1 && to.r2 === from.r2
  if ([down, up, right, left].filter(Boolean).length !== 1 || !fits || to.r1 > from.r1 || to.r2 < from.r2 || to.c1 > from.c1 || to.c2 < from.c2) throw new Error(`${where}: fill must extend the range in one direction`)
  if (tooLarge(to)) throw new Error(`${where}: the range is too large`)
  for (const si of [...sh.shared.keys()]) sh.unshare(si)
  const s0 = vertical ? from.r1 : from.c1
  const n = vertical ? from.r2 - from.r1 + 1 : from.c2 - from.c1 + 1
  const lo = vertical ? to.r1 : to.c1; const hi = vertical ? to.r2 : to.c2
  for (let line = vertical ? from.c1 : from.r1; line <= (vertical ? from.c2 : from.r2); line++) {
    const at = (pos) => (vertical ? [pos, line] : [line, pos])
    const src = []
    for (let i = 0; i < n; i++) { const [r, c] = at(s0 + i); src.push({ r, c, cell: sh.cellAt(r, c) }) }
    const val = (s) => (s.cell && sh.formula(s.cell, s.r, s.c) === undefined ? sh.value(s.cell) : undefined)
    const nums = src.map(val)
    let trend
    if (nums.every((v) => typeof v === 'number')) {
      if (n >= 2) {
        const xm = (n - 1) / 2; const ym = nums.reduce((a, b) => a + b, 0) / n
        let num = 0; let den = 0
        nums.forEach((y, x) => { num += (x - xm) * (y - ym); den += (x - xm) ** 2 })
        const b = num / den
        trend = (k) => ym + b * (k - xm)
      } else if (isDateFmt(sh.wb.fmtCode(src[0].cell.attrs.s ? +src[0].cell.attrs.s : 0))) trend = (k) => nums[0] + k
    }
    const word = n === 1 && typeof nums[0] === 'string' ? /^(.*?)(\d+)(\D*)$/.exec(nums[0]) : null
    for (let pos = lo; pos <= hi; pos++) {
      if (pos >= s0 && pos < s0 + n) continue
      const k = pos - s0
      const base = src[((k % n) + n) % n]
      const [r, c] = at(pos)
      if (!base.cell) { const ex = sh.cellAt(r, c); if (ex) { sh.clearContent(ex); delete ex.attrs.s } continue }
      placeCopy(sh, base.cell, r, c, r - base.r, c - base.c, wrote)
      if (trend) sh.set(r, c, { value: +trend(k).toPrecision(15) })
      else if (word) sh.set(r, c, { value: word[1] + String(Math.abs(+word[2] + k)).padStart(word[2].length, '0') + word[3] })
    }
  }
}

// sortRange orders the rows of a range by one column, the way Excel
// does: numbers, then text, then logical values and errors, with blanks
// last either way. Rows move whole within the range; formulas in them
// move their relative references with them.
function sortRange (sh, rg, keys, header, wrote, where) {
  if (!rg) throw new Error(`${where}: range must be like A1:D20`)
  for (const k of keys) if (!(k.col >= rg.c1 && k.col <= rg.c2)) throw new Error(`${where}: the sort column is outside the range`)
  const b = sh.bounds()
  const r1 = header ? rg.r1 + 1 : rg.r1
  const r2 = Math.min(rg.r2, b.r2)
  const c2 = Math.min(rg.c2, Math.max(b.c2, ...keys.map((k) => k.col)))
  if (r2 <= r1) return
  for (const ref of sh.merges()) {
    const m = parseRange(ref)
    if (m && !(m.r2 < r1 || m.r1 > r2 || m.c2 < rg.c1 || m.c1 > c2)) throw new Error(`${where}: 排序区域里有合并单元格，请先取消合并`)
  }
  for (const si of [...sh.shared.keys()]) sh.unshare(si)
  const rows = []
  for (let r = r1; r <= r2; r++) {
    const vals = keys.map((k) => { const x = sh.cellAt(r, k.col); return x ? sh.value(x) : null })
    const cells = []
    for (let c = rg.c1; c <= c2; c++) cells.push(dropCell(sh, r, c))
    rows.push({ r, cells, vals })
  }
  const rank = (v) => (v === null || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3)
  // Blanks go last whatever the order; later keys break ties.
  const cmp = (p, q, desc) => {
    const a = rank(p); const bb = rank(q)
    if (a === 4 || bb === 4) return a === bb ? 0 : a === 4 ? 1 : -1
    let d = a - bb
    if (!d) d = a === 0 ? p - q : a === 1 ? p.localeCompare(q, 'zh-CN', { numeric: true, sensitivity: 'base' }) : a === 2 ? +p - +q : 0
    return desc ? -d : d
  }
  rows.sort((x, y) => { for (let i = 0; i < keys.length; i++) { const d = cmp(x.vals[i], y.vals[i], keys[i].desc); if (d) return d } return x.r - y.r })
  rows.forEach((row, i) => {
    const r = r1 + i
    row.cells.forEach((cell, j) => { if (cell) placeCopy(sh, cell, r, rg.c1 + j, r - row.r, 0, wrote) })
  })
  sh.wb.pkg.touch(sh.part)
}

// setFilter turns on an AutoFilter over a range (its first row holds the
// headings) and sets which values one column shows; rows that fail any
// column's filter are hidden.
function setFilter (sh, e, where) {
  const p = sh.p
  let af = kid(sh.root, 'autoFilter')
  if (e.range) {
    const rg = parseRange(e.range)
    if (!rg) throw new Error(`${where}: range must be like A1:D20`)
    const ref = `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`
    if (!af) { af = el(p + 'autoFilter'); insertOrdered(sh.root, af, WS_ORDER.map((n) => p + n)) }
    if (af.attrs.ref !== ref) {
      const old = parseRange(af.attrs.ref ?? '')
      if (old) for (let r = old.r1 + 1; r <= old.r2; r++) { const row = sh.rows.get(r); if (row) delete row.attrs.hidden }
      af.attrs.ref = ref
      af.kids = []
    }
  }
  if (!af) throw new Error(`${where}: this sheet has no filter`)
  const rg = parseRange(af.attrs.ref ?? '')
  if (!rg) throw new Error(`${where}: the filter range is broken`)
  if (e.col !== undefined && e.col !== null) {
    if (!(+e.col >= rg.c1 && +e.col <= rg.c2)) throw new Error(`${where}: column ${e.col} is outside the filter`)
    const colId = String(+e.col - rg.c1)
    af.kids = af.kids.filter((k) => !(local(k) === 'filterColumn' && k.attrs.colId === colId))
    if (Array.isArray(e.values)) {
      const vals = e.values.map(String)
      const fs = el(p + 'filters', vals.includes('') ? { blank: '1' } : {}, vals.filter((x) => x !== '').map((x) => el(p + 'filter', { val: x })))
      af.kids.push(el(p + 'filterColumn', { colId }, [fs]))
      af.kids.sort((x, y) => +(x.attrs?.colId ?? 1e9) - +(y.attrs?.colId ?? 1e9))
    }
  }
  const cols = kids(af, 'filterColumn').map((fc) => [rg.c1 + +fc.attrs.colId, filterValues(fc)]).filter(([, v]) => v)
  for (let r = rg.r1 + 1; r <= rg.r2; r++) {
    const show = cols.every(([c, vals]) => vals.includes(sh.display(r, c)))
    if (show) { const row = sh.rows.get(r); if (row) delete row.attrs.hidden } else sh.row(r).attrs.hidden = '1'
  }
  sh.wb.pkg.touch(sh.part)
}

// moveLines inserts (count > 0) or deletes (count < 0) rows or columns at
// a 1-based index, moving cells, sizes, merges and every reference that
// points at the sheet: formulas in all sheets, defined names, conditional
// formats, validations, hyperlinks and the filter range.
function moveLines (wb, sh, axis, at, count, wrote) {
  if (!(at >= 1)) throw new Error(`${axis === 'r' ? 'row' : 'column'} index must start at 1`)
  for (const si of [...sh.shared.keys()]) sh.unshare(si)
  const sheetName = sh.info.name
  const del = count < 0 ? -count : 0
  const moveIdx = (x) => (count > 0 ? (x >= at ? x + count : x) : x >= at + del ? x - del : x >= at ? null : x)
  // Cells and rows.
  if (axis === 'r') {
    const rows = [...sh.rows.entries()].sort((a, b) => a[0] - b[0])
    sh.rows = new Map()
    for (const [r, row] of rows) {
      const nr = moveIdx(r)
      if (nr === null) { remove(sh.data, row); continue }
      row.attrs.r = String(nr)
      sh.rows.set(nr, row)
    }
  } else {
    for (const row of sh.rows.values()) {
      for (const c of kids(row, 'c')) {
        const p = parseCell(c.attrs.r)
        const nc = moveIdx(p.c)
        if (nc === null) remove(row, c)
      }
    }
    const cols = kid(sh.root, 'cols')
    if (cols) {
      const next = []
      for (const col of kids(cols, 'col')) {
        let a = +col.attrs.min; let b = +col.attrs.max
        if (count > 0) { if (a >= at) { a += count; b += count } else if (b >= at) b += count } else {
          const lo = at; const hi = at + del - 1
          if (b < lo) { /* before */ } else if (a > hi) { a -= del; b -= del } else {
            const keepA = a < lo ? a : lo; const keepB = b > hi ? b - del : lo - 1
            if (keepB < keepA) continue
            a = keepA; b = keepB
          }
        }
        col.attrs.min = String(a); col.attrs.max = String(b)
        next.push(col)
      }
      cols.kids = next
      if (!next.length) remove(sh.root, cols)
    }
  }
  sh.cells = new Map()
  for (const [r, row] of sh.rows) {
    for (const c of kids(row, 'c')) {
      const p = parseCell(c.attrs.r)
      const nc = axis === 'c' ? moveIdx(p.c) : p.c
      c.attrs.r = cellName(r, nc)
      sh.cells.set(`${r},${nc}`, c)
    }
    if (axis === 'c') row.kids.sort((x, y) => (local(x) === 'c' && local(y) === 'c' ? parseCell(x.attrs.r).c - parseCell(y.attrs.r).c : 0))
    delete row.attrs.spans
  }
  // References.
  const fix = (src, home) => moveRefs(src, sheetName, home, axis, at, count)
  for (const s of wb.sheets) {
    let other
    try { other = wb.sheet(s.name) } catch { continue }
    let changed = false
    for (const f of descendants(other.root, (n) => local(n) === 'f')) {
      const t = textContent(f)
      if (!t) continue
      const u = fix(t, s.name)
      if (u !== t) { f.kids = [text(u)]; changed = true }
      if (f.attrs.ref && other === sh) { const r = moveRange(f.attrs.ref, axis, at, count); if (r) f.attrs.ref = r; else delete f.attrs.ref }
    }
    if (changed) { wb.pkg.touch(other.part); for (const [key, c] of other.cells) if (kid(c, 'f')) wrote.add(`${s.name}!${key}`) }
  }
  for (const dn of descendants(wb.wb.root, (n) => local(n) === 'definedName')) {
    const t = textContent(dn)
    const u = fix(t, undefined)
    if (u !== t) dn.kids = [text(u)]
  }
  wb.pkg.touch(wb.main)
  shiftCharts(wb, sh, (t) => fix(t, undefined), axis, at, count)
  const fixAttr = (parentName, itemName, attr) => {
    const parent = kid(sh.root, parentName)
    if (!parent) return
    for (const item of itemName ? kids(parent, itemName) : [parent]) {
      const parts = String(item.attrs[attr] ?? '').split(/\s+/).filter(Boolean).map((r) => moveRange(r, axis, at, count)).filter(Boolean)
      if (parts.length) item.attrs[attr] = parts.join(' ')
      else if (itemName) remove(parent, item)
      else remove(sh.root, parent)
    }
    if (itemName && !kids(parent, itemName).length) remove(sh.root, parent)
    else if (itemName && parent.attrs.count !== undefined) parent.attrs.count = String(kids(parent, itemName).length)
  }
  fixAttr('mergeCells', 'mergeCell', 'ref')
  for (const cf of kids(sh.root, 'conditionalFormatting')) {
    const parts = String(cf.attrs.sqref ?? '').split(/\s+/).filter(Boolean).map((r) => moveRange(r, axis, at, count)).filter(Boolean)
    if (parts.length) cf.attrs.sqref = parts.join(' '); else remove(sh.root, cf)
  }
  fixAttr('dataValidations', 'dataValidation', 'sqref')
  fixAttr('hyperlinks', 'hyperlink', 'ref')
  fixAttr('autoFilter', null, 'ref')
  // A stale selection points at moved cells; Excel picks a new one.
  for (const view of kids(kid(sh.root, 'sheetViews'), 'sheetView')) view.kids = view.kids.filter((x) => local(x) !== 'selection')
  wb.pkg.touch(sh.part)
}

function moveRange (ref, axis, at, count) {
  const [a, b] = ref.split(':')
  const p = parseCell(a); const q = b === undefined ? p : parseCell(b)
  if (!p || !q) return ref
  const del = count < 0 ? -count : 0
  let lo = axis === 'r' ? p.r : p.c
  let hi = axis === 'r' ? q.r : q.c
  if (count > 0) { if (lo >= at) lo += count; if (hi >= at) hi += count } else {
    const dlo = at; const dhi = at + del - 1
    if (lo > dhi) lo -= del; else if (lo >= dlo) lo = dlo
    if (hi > dhi) hi -= del; else if (hi >= dlo) hi = dlo - 1
    if (hi < lo) return null
  }
  const P = axis === 'r' ? { r: lo, c: p.c } : { r: p.r, c: lo }
  const Q = axis === 'r' ? { r: hi, c: q.c } : { r: q.r, c: hi }
  return b === undefined && lo === hi ? cellName(P.r, P.c) : `${cellName(P.r, P.c)}:${cellName(Q.r, Q.c)}`
}

// moveRefs rewrites the references in one formula that point at sheet.
function moveRefs (src, sheet, home, axis, at, count) {
  let toks
  try { toks = tokenize(src) } catch { return src }
  const del = count < 0 ? -count : 0
  const same = (n) => n.replace(/^'|'$/g, '').replace(/''/g, "'").toLowerCase() === sheet.toLowerCase()
  return toks.map((t) => {
    if (t.kind !== 'ref') return t.text
    const bang = t.text.lastIndexOf('!')
    const sh = bang >= 0 ? t.text.slice(0, bang) : undefined
    if (sh !== undefined ? !same(sh) : home === undefined || home.toLowerCase() !== sheet.toLowerCase()) return t.text
    const prefix = bang >= 0 ? t.text.slice(0, bang + 1) : ''
    const body = bang >= 0 ? t.text.slice(bang + 1) : t.text
    const parts = body.split(':')
    const re = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/
    const ends = parts.map((x) => { const m = re.exec(x); return m && { ca: m[1], col: m[2], ra: m[3], row: m[4] } })
    if (ends.some((x) => !x)) return t.text
    const idx = (e) => (axis === 'r' ? (e.row ? +e.row : undefined) : (e.col ? colNum(e.col) : undefined))
    const set = (e, n) => (axis === 'r' ? { ...e, row: String(n) } : { ...e, col: colName(n) })
    const vals = ends.map(idx)
    if (vals.some((x) => x === undefined)) return t.text // whole rows/columns on the other axis
    let lo = vals[0]; let hi = vals[vals.length - 1]
    if (count > 0) { if (lo >= at) lo += count; if (hi >= at) hi += count } else {
      const dlo = at; const dhi = at + del - 1
      if (ends.length === 1) { if (lo >= dlo && lo <= dhi) return '#REF!'; if (lo > dhi) lo -= del; hi = lo } else {
        if (lo > dhi) lo -= del; else if (lo >= dlo) lo = dlo
        if (hi > dhi) hi -= del; else if (hi >= dlo) hi = dlo - 1
        if (hi < lo) return '#REF!'
      }
    }
    const out = ends.length === 1 ? [set(ends[0], lo)] : [set(ends[0], lo), set(ends[ends.length - 1], hi)]
    return prefix + out.map((e) => `${e.ca}${e.col ?? ''}${e.ra}${e.row ?? ''}`).join(':')
  }).join('')
}

// editOps runs Agent-style edit ops against an open workbook.
function editOps (wb, edits, wrote) {
  const rangeOf = (e, where) => {
    const rg = parseRange(e.range ?? e.cell ?? '')
    if (!rg) throw new Error(`${where}: range "${e.range ?? e.cell}" is not like A1 or A1:C9`)
    return rg
  }
  for (const [k, e] of edits.entries()) {
    const where = `ops[${k}] (${e?.op})`
    switch (e?.op) {
      case 'style': {
        const sh = wb.sheet(e.sheet)
        const rg = rangeOf(e, where)
        const props = styleProps(e)
        if (Object.keys(props).length) {
          if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 200000) throw new Error(`${where}: the range is too large to style cell by cell`)
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) sh.set(r, c, props)
        }
        break
      }
      case 'clear': wb.sheet(e.sheet).clear(rangeOf(e, where)); break
      case 'freeze': wb.sheet(e.sheet).freeze(e.cell ?? 'A1'); break
      case 'add_sheet': wb.addSheet(e.name); break
      case 'rename_sheet': wb.renameSheet(e.sheet, e.name); break
      case 'delete_sheet': wb.deleteSheet(e.sheet); break
      case 'merge': case 'unmerge': {
        const sh = wb.sheet(e.sheet)
        const rg = rangeOf(e, where)
        let mc = kid(sh.root, 'mergeCells')
        const ref = `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`
        if (mc) {
          // Drop merges the range overlaps.
          mc.kids = mc.kids.filter((m) => { const x = parseRange(m.attrs?.ref ?? ''); return !x || x.r2 < rg.r1 || x.r1 > rg.r2 || x.c2 < rg.c1 || x.c1 > rg.c2 })
        }
        if (e.op === 'merge' && (rg.r1 !== rg.r2 || rg.c1 !== rg.c2)) {
          if (!mc) { mc = el(sh.p + 'mergeCells'); insertOrdered(sh.root, mc, WS_ORDER.map((n) => sh.p + n)) }
          mc.kids.push(el(sh.p + 'mergeCell', { ref }))
        }
        if (mc) { if (mc.kids.length) mc.attrs.count = String(mc.kids.length); else remove(sh.root, mc) }
        wb.pkg.touch(sh.part)
        break
      }
      default:
        throw new Error(`${where}: unknown op`)
    }
  }
}
