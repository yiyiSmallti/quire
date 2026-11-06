// Charts on worksheets: read the drawing anchors and chart parts the
// editor shows, make new charts from a range, move, retype and delete
// them, and keep each series' cached values in step with the cells (other
// readers draw from the cache).
import { el, text, doc, descendants, textContent, remove, clone, insertOrdered, isEl } from './xml.js'
import { parseRange, colName } from './formula.js'
import { REL } from './pkg.js'

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const EMU = 9525
const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300']
export const CHART_TYPES = ['col', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']

const local = (n) => n?.name?.replace(/^\w+:/, '')
const kid = (n, name) => n?.kids?.find((k) => local(k) === name)
const kids = (n, name) => (n?.kids ?? []).filter((k) => local(k) === name)
const val = (n) => n?.attrs?.val
const cv = (name, v) => el('c:' + name, { val: String(v) })

// ------------------------------------------------------------- geometry

// geometry gives a sheet's column widths and row heights in pixels, the
// way the editor draws them.
export function geometry (sh) {
  const fmt = kid(sh.root, 'sheetFormatPr')
  const defW = fmt?.attrs.defaultColWidth ? Math.round(+fmt.attrs.defaultColWidth * 7 + 5) : fmt?.attrs.baseColWidth ? Math.round((+fmt.attrs.baseColWidth + 0.71) * 7 + 5) : 64
  const defH = fmt?.attrs.defaultRowHeight ? Math.round(+fmt.attrs.defaultRowHeight * 4 / 3) : 20
  const cols = kids(kid(sh.root, 'cols'), 'col').map((c) => ({ min: +c.attrs.min, max: +c.attrs.max, px: c.attrs.hidden === '1' ? 0 : c.attrs.width ? Math.round(+c.attrs.width * 7 + 5) : defW }))
  return {
    col: (c) => cols.find((x) => c >= x.min && c <= x.max)?.px ?? defW,
    row: (r) => {
      const row = sh.rows.get(r)
      if (row?.attrs.hidden === '1') return 0
      return row?.attrs.ht && row.attrs.customHeight === '1' ? Math.round(+row.attrs.ht * 4 / 3) : defH
    }
  }
}

// walk finds the cell (1-based) and offset a distance from a cell lands in.
function walk (start, offset, size, px) {
  let rem = offset + size
  let i = start
  for (let n = 0; n < 20000 && rem > px(i) && i < 16384; n++) { rem -= px(i); i++ }
  return { i, d: Math.max(0, rem) }
}

function marker (m) {
  return { c: +textContent(kid(m, 'col')) + 1, dx: Math.round(+textContent(kid(m, 'colOff')) / EMU), r: +textContent(kid(m, 'row')) + 1, dy: Math.round(+textContent(kid(m, 'rowOff')) / EMU) }
}

// frameOf reads an anchor as two cell corners in pixels.
function frameOf (anchor, g) {
  const kind = local(anchor)
  if (kind === 'twoCellAnchor') return { from: marker(kid(anchor, 'from')), to: marker(kid(anchor, 'to')) }
  const ext = kid(anchor, 'ext')
  const w = Math.round(+(ext?.attrs.cx ?? 0) / EMU); const h = Math.round(+(ext?.attrs.cy ?? 0) / EMU)
  let from
  if (kind === 'oneCellAnchor') from = marker(kid(anchor, 'from'))
  else {
    const pos = kid(anchor, 'pos')
    const x = walk(1, 0, Math.round(+(pos?.attrs.x ?? 0) / EMU), g.col); const y = walk(1, 0, Math.round(+(pos?.attrs.y ?? 0) / EMU), g.row)
    from = { c: x.i, dx: x.d, r: y.i, dy: y.d }
  }
  const x = walk(from.c, from.dx, w, g.col); const y = walk(from.r, from.dy, h, g.row)
  return { from, to: { c: x.i, dx: x.d, r: y.i, dy: y.d } }
}

const markerXml = (p, name, m) => el(p + name, {}, [el(p + 'col', {}, [text(Math.max(0, m.c - 1))]), el(p + 'colOff', {}, [text(Math.round(Math.max(0, m.dx) * EMU))]), el(p + 'row', {}, [text(Math.max(0, m.r - 1))]), el(p + 'rowOff', {}, [text(Math.round(Math.max(0, m.dy) * EMU))])])

// ------------------------------------------------------------- parts

function drawingOf (pkg, sheetPart, create, sh) {
  const rel = pkg.relsOfType(sheetPart, 'drawing')[0]
  if (rel?.part && pkg.has(rel.part)) return { part: rel.part, root: pkg.xml(rel.part).root }
  if (!create) return undefined
  const part = pkg.uniqueName('xl/drawings/drawing#.xml')
  pkg.setXml(part, doc(el('xdr:wsDr', { 'xmlns:xdr': XDR, 'xmlns:a': A }, [])))
  pkg.setOverride(part, CT_DRAWING)
  const rid = pkg.addRel(sheetPart, REL('drawing'), part)
  if (!sh.root.attrs['xmlns:r']) sh.root.attrs['xmlns:r'] = R
  insertOrdered(sh.root, el(sh.p + 'drawing', { 'r:id': rid }), WS_ORDER.map((n) => sh.p + n))
  pkg.touch(sheetPart)
  return { part, root: pkg.xml(part).root }
}
const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst']

// chartAnchors lists the anchors on a drawing that hold a chart, with the
// chart part each points at.
function chartAnchors (pkg, dr) {
  if (!dr) return []
  const out = []
  for (const a of dr.root.kids) {
    if (!['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(local(a))) continue
    const ref = descendants(a, (n) => local(n) === 'chart' && Object.keys(n.attrs).some((k) => k.endsWith(':id')))[0]
    if (!ref) continue
    const rid = Object.entries(ref.attrs).find(([k]) => k.endsWith(':id'))[1]
    const rel = pkg.rel(dr.part, rid)
    out.push({ anchor: a, rid, part: rel?.part && pkg.has(rel.part) ? rel.part : undefined })
  }
  return out
}

function chartAt (wb, sh, k, where) {
  const dr = drawingOf(wb.pkg, sh.part, false)
  const list = chartAnchors(wb.pkg, dr)
  const hit = list[+k]
  if (!hit) throw new Error(`${where}: sheet "${sh.info.name}" has no chart ${k}; reload`)
  return { dr, ...hit }
}

// ------------------------------------------------------------- reading

// resolve reads a range formula like 'Sheet 1'!$B$2:$B$9 from the cells,
// row by row; undefined when it is not a plain range on a worksheet.
function resolve (wb, f, shown, visOnly) {
  const m = /^(?:'((?:[^']|'')+)'|([^!'"]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(String(f ?? '').trim())
  if (!m) return undefined
  const name = (m[1] ?? m[2]).replace(/''/g, "'")
  let sh
  try { sh = wb.sheet(name) } catch { return undefined }
  const rg = parseRange(`${m[3]}${m[4]}:${m[5] ?? m[3]}${m[6] ?? m[4]}`)
  if (!rg || (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 5000) return undefined
  const out = []
  const hidCols = visOnly ? hiddenCols(sh) : null
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (visOnly && sh.rows.get(r)?.attrs.hidden === '1') continue
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (hidCols?.has(c)) continue
      if (shown) out.push(String(sh.display(r, c) ?? ''))
      else { const v = sh.value(sh.cellAt(r, c)); out.push(typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && Number.isFinite(+v) ? +v : null) }
    }
  }
  return out
}

// hiddenCols lists the columns a sheet hides.
function hiddenCols (sh) {
  const out = new Set()
  for (const c of kids(kid(sh.root, 'cols'), 'col')) if (c.attrs.hidden === '1' || c.attrs.hidden === 'true') for (let i = +c.attrs.min; i <= Math.min(+c.attrs.max, 16384); i++) out.add(i)
  return out
}

function cacheOf (ref) {
  const cache = kid(ref, 'strCache') ?? kid(ref, 'numCache') ?? kid(ref, 'multiLvlStrCache')
  if (!cache) return []
  const n = +(val(kid(cache, 'ptCount')) ?? 0)
  const out = Array.from({ length: n }, () => null)
  for (const pt of descendants(cache, (x) => local(x) === 'pt')) { const i = +pt.attrs.idx; if (i < 1e5) out[i] = textContent(kid(pt, 'v')) }
  return out
}

// dataOf reads a cat/val/tx element: its formula and values (from the
// cells when they resolve, else the cache).
function dataOf (wb, node, shown, visOnly) {
  if (!node) return undefined
  const ref = kid(node, 'strRef') ?? kid(node, 'numRef') ?? kid(node, 'multiLvlStrRef')
  if (ref) {
    const f = textContent(kid(ref, 'f'))
    const live = resolve(wb, f, shown, visOnly)
    const vals = live ?? cacheOf(ref).map((x) => (shown ? (x ?? '') : x === null || x === '' ? null : +x))
    return { f, vals }
  }
  const lit = kid(node, 'strLit') ?? kid(node, 'numLit')
  if (lit) return { vals: cacheOf(lit).map((x) => (shown ? (x ?? '') : x === null ? null : +x)) }
  const v = kid(node, 'v')
  if (v) return { vals: [textContent(v)] }
  return undefined
}

const TYPE_OF = { barChart: 'col', bar3DChart: 'col', lineChart: 'line', line3DChart: 'line', stockChart: 'line', radarChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', bubbleChart: 'scatter' }

function colorOf (ser, theme, i) {
  const pr = kid(ser, 'spPr')
  const fill = kid(pr, 'solidFill') ?? kid(kid(pr, 'ln'), 'solidFill') ?? kid(kid(kid(ser, 'marker'), 'spPr'), 'solidFill')
  const rgb = val(kid(fill, 'srgbClr'))
  if (rgb) return rgb.toUpperCase()
  const scheme = val(kid(fill, 'schemeClr'))
  const m = /^accent(\d)$/.exec(scheme ?? '')
  if (m) return (theme?.[3 + +m[1]] ?? PALETTE[+m[1] - 1]).toUpperCase()
  return PALETTE[i % PALETTE.length]
}

function richText (node) {
  return descendants(node, (n) => local(n) === 'p').map((p) => descendants(p, (n) => local(n) === 't').map(textContent).join('')).join('\n')
}

function parseChart (wb, part) {
  const root = wb.pkg.xml(part).root
  const chart = kid(root, 'chart')
  const plot = kid(chart, 'plotArea')
  const typeEls = (plot?.kids ?? []).filter((k) => TYPE_OF[local(k)])
  const first = typeEls[0]
  if (!first) return { editable: false, type: 'col', series: [] }
  let type = TYPE_OF[local(first)]
  if (type === 'col' && val(kid(first, 'barDir')) === 'bar') type = 'bar'
  const grouping = val(kid(first, 'grouping'))
  const theme = wb.loadTheme?.()
  // Hidden rows and columns (a filter's too) are left out unless the
  // chart says to plot them.
  const visOnly = val(kid(chart, 'plotVisOnly')) !== '0'
  const series = []
  for (const t of typeEls) {
    for (const ser of kids(t, 'ser')) {
      const tx = kid(ser, 'tx')
      const name = tx ? (dataOf(wb, tx, true)?.vals ?? []).join(' ') : ''
      const cat = dataOf(wb, kid(ser, 'cat') ?? kid(ser, 'xVal'), true, visOnly)
      const vals = dataOf(wb, kid(ser, 'val') ?? kid(ser, 'yVal'), false, visOnly)
      series.push({ name: name || `系列${series.length + 1}`, cats: cat?.vals ?? [], vals: vals?.vals ?? [], color: colorOf(ser, theme, series.length) })
    }
  }
  const titleEl = kid(chart, 'title')
  let title
  if (titleEl) {
    const rich = kid(kid(titleEl, 'tx'), 'rich')
    title = rich ? richText(rich) : (dataOf(wb, kid(titleEl, 'tx'), true)?.vals ?? []).join(' ')
    if (!title) title = series.length === 1 ? series[0].name : '图表标题'
  } else if (val(kid(chart, 'autoTitleDeleted')) !== '1' && series.length === 1 && (type === 'pie' || type === 'doughnut')) title = series[0].name
  const legend = kid(chart, 'legend')
  return {
    editable: typeEls.length === 1 && !['bar3DChart', 'line3DChart', 'area3DChart', 'pie3DChart', 'stockChart', 'radarChart', 'bubbleChart', 'ofPieChart'].includes(local(first)),
    type,
    stacked: grouping === 'stacked' ? 'stacked' : grouping === 'percentStacked' ? 'percent' : undefined,
    title,
    legend: legend ? (val(kid(legend, 'legendPos')) ?? 'r') : null,
    ...lookOf(plot, first),
    series
  }
}

// chartAxes finds the axis along the categories (x) and the value axis
// (y): for a scatter chart the bottom and the other value axis.
function chartAxes (plot) {
  const all = (plot?.kids ?? []).filter((k) => ['catAx', 'valAx', 'dateAx'].includes(local(k)))
  let x = all.find((a) => local(a) !== 'valAx')
  if (!x) x = all.find((a) => ['b', 't'].includes(val(kid(a, 'axPos')))) ?? all[0]
  const y = all.find((a) => a !== x && local(a) === 'valAx')
  return { x, y }
}

function dLblsOf (typeEl) {
  return kid(typeEl, 'dLbls') ?? kid(kids(typeEl, 'ser')[0], 'dLbls')
}

// lookOf reads the settings the options dialog changes.
function lookOf (plot, typeEl) {
  const d = dLblsOf(typeEl)
  const on = (name) => val(kid(d, name)) === '1' || val(kid(d, name)) === 'true'
  const labels = d && !kid(d, 'delete') ? (on('showPercent') ? 'percent' : on('showVal') ? 'value' : on('showCatName') ? 'category' : null) : null
  const { x, y } = chartAxes(plot)
  const axTitle = (a) => { const t = kid(a, 'title'); if (!t) return undefined; const rich = kid(kid(t, 'tx'), 'rich'); return rich ? richText(rich) : '' }
  const sc = kid(y, 'scaling')
  const num = (n) => (n && Number.isFinite(+val(n)) ? +val(n) : undefined)
  return { labels, xTitle: axTitle(x), yTitle: axTitle(y), yMin: num(kid(sc, 'min')), yMax: num(kid(sc, 'max')), gridlines: y ? !!kid(y, 'majorGridlines') : undefined }
}

// sheetCharts is what the editor needs to draw a sheet's charts.
export function sheetCharts (wb, sh) {
  const dr = drawingOf(wb.pkg, sh.part, false)
  if (!dr) return []
  const g = geometry(sh)
  return chartAnchors(wb.pkg, dr).map((x, k) => {
    let info
    try { info = x.part ? parseChart(wb, x.part) : { editable: false, type: 'col', series: [] } } catch { info = { editable: false, type: 'col', series: [] } }
    return { k, ...frameOf(x.anchor, g), ...info }
  })
}

// ------------------------------------------------------------- writing

const quote = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
const absRange = (sheet, r1, c1, r2, c2) => `${quote(sheet)}!$${colName(c1)}$${r1}${r2 === r1 && c2 === c1 ? '' : `:$${colName(c2)}$${r2}`}`

function strRef (f, vals) {
  return el('c:strRef', {}, [el('c:f', {}, [text(f)]), el('c:strCache', {}, [cv('ptCount', vals.length), ...vals.map((v, i) => el('c:pt', { idx: String(i) }, [el('c:v', {}, [text(v ?? '')])]))])])
}
function numRef (f, vals) {
  return el('c:numRef', {}, [el('c:f', {}, [text(f)]), el('c:numCache', {}, [el('c:formatCode', {}, [text('General')]), cv('ptCount', vals.length), ...vals.map((v, i) => (v === null || v === undefined ? null : el('c:pt', { idx: String(i) }, [el('c:v', {}, [text(v)])])))])])
}

const fillPr = (color) => el('c:spPr', {}, [el('a:solidFill', {}, [el('a:srgbClr', { val: color })]), el('a:ln', {}, [el('a:noFill')])])
const linePr = (color) => el('c:spPr', {}, [el('a:ln', { w: '28575', cap: 'rnd' }, [el('a:solidFill', {}, [el('a:srgbClr', { val: color })]), el('a:round')])])

// serXml writes one series for a chart type from its parts: tx, cat and
// val (cloned elements or freshly made), in the order the schema wants.
function serXml (type, i, parts, color) {
  const tx = parts.tx ? clone(parts.tx) : null
  const cat = parts.cat ? retag(clone(parts.cat), type === 'scatter' ? 'xVal' : 'cat') : null
  const vals = parts.val ? retag(clone(parts.val), type === 'scatter' ? 'yVal' : 'val') : null
  const head = [cv('idx', i), cv('order', i), tx]
  if (type === 'col' || type === 'bar') return el('c:ser', {}, [...head, fillPr(color), cv('invertIfNegative', 0), cat, vals])
  if (type === 'line') return el('c:ser', {}, [...head, linePr(color), el('c:marker', {}, [cv('symbol', 'none')]), cat, vals, cv('smooth', 0)])
  // Markers only, like Excel's plain scatter chart.
  if (type === 'scatter') return el('c:ser', {}, [...head, el('c:spPr', {}, [el('a:ln', { w: '19050' }, [el('a:noFill')])]), el('c:marker', {}, [cv('symbol', 'circle'), cv('size', 5), el('c:spPr', {}, [el('a:solidFill', {}, [el('a:srgbClr', { val: color })]), el('a:ln', {}, [el('a:noFill')])])]), cat, vals, cv('smooth', 0)])
  if (type === 'area') return el('c:ser', {}, [...head, fillPr(color), cat, vals])
  return el('c:ser', {}, [...head, cat, vals])
}

function retag (node, name) {
  node.name = (node.name.includes(':') ? node.name.split(':')[0] + ':' : '') + name
  return node
}

const gridlinesXml = () => el('c:majorGridlines', {}, [el('c:spPr', {}, [el('a:ln', { w: '9525' }, [el('a:solidFill', {}, [el('a:srgbClr', { val: 'D9D9D9' })])])])])

function axes (type) {
  if (type === 'pie' || type === 'doughnut') return []
  const gridlines = gridlinesXml()
  const common = (id, cross, pos) => [cv('axId', id), el('c:scaling', {}, [cv('orientation', 'minMax')]), cv('delete', 0), cv('axPos', pos)]
  const ticks = (cross) => [cv('majorTickMark', 'out'), cv('minorTickMark', 'none'), cv('tickLblPos', 'nextTo'), cv('crossAx', cross), cv('crosses', 'autoZero')]
  const horiz = type === 'bar'
  if (type === 'scatter') {
    return [el('c:valAx', {}, [...common(111, 222, 'b'), el('c:numFmt', { formatCode: 'General', sourceLinked: '1' }), ...ticks(222), cv('crossBetween', 'midCat')]),
      el('c:valAx', {}, [...common(222, 111, 'l'), gridlines, el('c:numFmt', { formatCode: 'General', sourceLinked: '1' }), ...ticks(111), cv('crossBetween', 'midCat')])]
  }
  return [el('c:catAx', {}, [...common(111, 222, horiz ? 'l' : 'b'), el('c:numFmt', { formatCode: 'General', sourceLinked: '1' }), ...ticks(222), cv('auto', 1), cv('lblAlgn', 'ctr'), cv('lblOffset', 100), cv('noMultiLvlLbl', 0)]),
    el('c:valAx', {}, [...common(222, 111, horiz ? 'b' : 'l'), gridlines, el('c:numFmt', { formatCode: 'General', sourceLinked: '1' }), ...ticks(111), cv('crossBetween', 'between')])]
}

function typeXml (type, sers, stacked) {
  const grouping = stacked === 'percent' ? 'percentStacked' : stacked ? 'stacked' : null
  const axIds = type === 'pie' || type === 'doughnut' ? [] : [cv('axId', 111), cv('axId', 222)]
  if (type === 'col' || type === 'bar') return el('c:barChart', {}, [cv('barDir', type === 'bar' ? 'bar' : 'col'), cv('grouping', grouping ?? 'clustered'), cv('varyColors', 0), ...sers, cv('gapWidth', 150), grouping ? cv('overlap', 100) : null, ...axIds])
  if (type === 'line') return el('c:lineChart', {}, [cv('grouping', grouping ?? 'standard'), cv('varyColors', 0), ...sers, cv('marker', 1), ...axIds])
  if (type === 'area') return el('c:areaChart', {}, [cv('grouping', grouping ?? 'standard'), cv('varyColors', 0), ...sers, ...axIds])
  if (type === 'pie') return el('c:pieChart', {}, [cv('varyColors', 1), ...sers, dLblsXml('percent', true), cv('firstSliceAng', 0)])
  if (type === 'doughnut') return el('c:doughnutChart', {}, [cv('varyColors', 1), ...sers, dLblsXml('percent', true), cv('firstSliceAng', 0), cv('holeSize', 50)])
  return el('c:scatterChart', {}, [cv('scatterStyle', 'lineMarker'), cv('varyColors', 0), ...sers, ...axIds])
}

function titleXml (title, sz = 1400, rot) {
  return el('c:title', {}, [el('c:tx', {}, [el('c:rich', {}, [el('a:bodyPr', rot ? { rot: String(rot), vert: 'horz' } : {}), el('a:lstStyle'), el('a:p', {}, [el('a:pPr', {}, [el('a:defRPr', { sz: String(sz), b: '0' })]), el('a:r', {}, [el('a:rPr', { lang: 'zh-CN', altLang: 'en-US' }), el('a:t', {}, [text(title)])])])])]), cv('overlay', 0)])
}

function chartXml (type, sers, title, stacked) {
  return doc(el('c:chartSpace', { 'xmlns:c': C, 'xmlns:a': A, 'xmlns:r': R }, [
    cv('roundedCorners', 0),
    el('c:chart', {}, [
      title ? titleXml(title) : null,
      cv('autoTitleDeleted', title ? 0 : 1),
      el('c:plotArea', {}, [el('c:layout'), typeXml(type, sers, stacked), ...axes(type)]),
      el('c:legend', {}, [cv('legendPos', type === 'pie' || type === 'doughnut' ? 'r' : 'b'), cv('overlay', 0)]),
      cv('plotVisOnly', 1),
      cv('dispBlanksAs', 'gap')
    ])
  ]))
}

// seriesFrom lays a range out as series the way Excel does: a text first
// row names the series, a text first column labels the categories, and
// series run down columns unless the range is wider than tall.
function seriesFrom (wb, sh, rg, type) {
  const isText = (r, c) => { const v = sh.value(sh.cellAt(r, c)); return typeof v === 'string' && !(v.trim() !== '' && Number.isFinite(+v)) }
  const isBlank = (r, c) => { const v = sh.value(sh.cellAt(r, c)); return v === null || v === '' }
  let head = false
  for (let c = rg.c1; c <= rg.c2; c++) if (isText(rg.r1, c)) head = true
  if (rg.r2 === rg.r1) head = false
  let catCol = rg.c2 > rg.c1
  for (let r = rg.r1 + (head ? 1 : 0); r <= rg.r2 && catCol; r++) if (!isText(r, rg.c1) && !isBlank(r, rg.c1)) catCol = false
  if (rg.c2 > rg.c1 && !catCol && head && isBlank(rg.r1, rg.c1)) catCol = true
  // A scatter chart takes its X values from the first column.
  const xNum = type === 'scatter' && rg.c2 > rg.c1 && !catCol
  if (xNum) catCol = true
  const name = sh.info.name
  const d1 = { r: rg.r1 + (head ? 1 : 0), c: rg.c1 + (catCol ? 1 : 0) }
  const rows = rg.r2 - d1.r + 1; const cols = rg.c2 - d1.c + 1
  if (rows < 1 || cols < 1) throw new Error('the range has no numbers to chart')
  const byCol = rows >= cols
  const out = []
  const n = byCol ? cols : rows
  for (let i = 0; i < Math.min(n, 30); i++) {
    const tx = byCol ? (head ? [rg.r1, d1.c + i] : null) : (catCol ? [d1.r + i, rg.c1] : null)
    const vr = byCol ? [d1.r, d1.c + i, rg.r2, d1.c + i] : [d1.r + i, d1.c, d1.r + i, rg.c2]
    const cr = byCol ? (catCol ? [d1.r, rg.c1, rg.r2, rg.c1] : null) : (head ? [rg.r1, d1.c, rg.r1, rg.c2] : null)
    const parts = {}
    if (tx) { const f = absRange(name, tx[0], tx[1], tx[0], tx[1]); parts.tx = el('c:tx', {}, [strRef(f, resolve(wb, f, true))]) }
    if (cr) { const f = absRange(name, ...cr); parts.cat = el('c:cat', {}, [xNum ? numRef(f, resolve(wb, f, false)) : strRef(f, resolve(wb, f, true))]) }
    const fv = absRange(name, ...vr)
    parts.val = el('c:val', {}, [numRef(fv, resolve(wb, fv, false))])
    out.push(parts)
  }
  return out
}

export function addChart (wb, sh, e, where) {
  const type = CHART_TYPES.includes(e.type) ? e.type : 'col'
  const rg = parseRange(String(e.range ?? ''))
  if (!rg) throw new Error(`${where}: range must be like A1:D9`)
  if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 5000) throw new Error(`${where}: the range is too large for a chart`)
  const parts = seriesFrom(wb, sh, rg, type)
  const sers = parts.map((p, i) => serXml(type, i, p, PALETTE[i % PALETTE.length]))
  const pkg = wb.pkg
  const dr = drawingOf(pkg, sh.part, true, sh)
  // The new anchor says xdr: and a:, whatever prefixes the drawing uses.
  if (dr.root.attrs['xmlns:xdr'] !== XDR) dr.root.attrs['xmlns:xdr'] = XDR
  if (dr.root.attrs['xmlns:a'] !== A) dr.root.attrs['xmlns:a'] = A
  const part = pkg.uniqueName('xl/charts/chart#.xml')
  pkg.setXml(part, chartXml(type, sers, e.title ? String(e.title) : '', e.stacked))
  pkg.setOverride(part, CT_CHART)
  const rid = pkg.addRel(dr.part, REL('chart'), part)
  const g = geometry(sh)
  const at = e.at && e.at.c >= 1 && e.at.r >= 1 ? { c: +e.at.c, r: +e.at.r } : { c: rg.c2 + 2, r: rg.r1 }
  const from = { c: at.c, dx: 0, r: at.r, dy: 0 }
  const w = Math.max(120, Math.min(2000, +e.w || 480)); const h = Math.max(80, Math.min(1500, +e.h || 288))
  const x = walk(from.c, 0, w, g.col); const y = walk(from.r, 0, h, g.row)
  const id = Math.max(1, ...descendants(dr.root, (n) => local(n) === 'cNvPr').map((n) => +n.attrs.id || 0)) + 1
  dr.root.kids.push(el('xdr:twoCellAnchor', { editAs: 'oneCell' }, [
    markerXml('xdr:', 'from', from), markerXml('xdr:', 'to', { c: x.i, dx: x.d, r: y.i, dy: y.d }),
    el('xdr:graphicFrame', { macro: '' }, [
      el('xdr:nvGraphicFramePr', {}, [el('xdr:cNvPr', { id: String(id), name: `图表 ${id - 1}` }), el('xdr:cNvGraphicFramePr')]),
      el('xdr:xfrm', {}, [el('a:off', { x: '0', y: '0' }), el('a:ext', { cx: '0', cy: '0' })]),
      el('a:graphic', {}, [el('a:graphicData', { uri: C }, [el('c:chart', { 'xmlns:c': C, 'xmlns:r': R, 'r:id': rid })])])
    ]),
    el('xdr:clientData')
  ]))
  pkg.touch(dr.part)
}

export function moveChart (wb, sh, e, where) {
  const { dr, anchor } = chartAt(wb, sh, e.chart, where)
  const f = e.from; const t = e.to
  if (!(f?.c >= 1 && f?.r >= 1 && t?.c >= 1 && t?.r >= 1)) throw new Error(`${where}: from and to must be cells with offsets`)
  const p = anchor.name.includes(':') ? anchor.name.split(':')[0] + ':' : ''
  const body = anchor.kids.filter((k) => isEl(k) && !['from', 'to', 'pos', 'ext'].includes(local(k)))
  anchor.name = p + 'twoCellAnchor'
  if (!anchor.attrs.editAs) anchor.attrs.editAs = 'oneCell'
  anchor.kids = [markerXml(p, 'from', f), markerXml(p, 'to', t), ...body]
  wb.pkg.touch(dr.part)
}

export function deleteChart (wb, sh, e, where) {
  const { dr, anchor, rid, part } = chartAt(wb, sh, e.chart, where)
  remove(dr.root, anchor)
  const stillUsed = chartAnchors(wb.pkg, dr).some((x) => x.rid === rid)
  if (!stillUsed) {
    wb.pkg.removeRel(dr.part, rid)
    if (part) {
      for (const r of wb.pkg.rels(part)) if (r.part && wb.pkg.has(r.part) && /chartStyle|chartColorStyle|themeOverride|package/.test(r.type)) { wb.pkg.delete(r.part); wb.pkg.removeOverride(r.part) }
      wb.pkg.delete(part)
      wb.pkg.removeOverride(part)
    }
  }
  wb.pkg.touch(dr.part)
}

// editChart changes a chart's title and, for charts made of one plain
// kind, its type and stacking (series keep their data, take new colors).
export function editChart (wb, sh, e, where) {
  const { part } = chartAt(wb, sh, e.chart, where)
  if (!part) throw new Error(`${where}: the chart part is missing`)
  const d = wb.pkg.xml(part)
  const chart = kid(d.root, 'chart')
  // New elements use the c: and a: prefixes; declare them on the root
  // when the file used others.
  if (d.root.attrs['xmlns:c'] !== C) d.root.attrs['xmlns:c'] = C
  if (d.root.attrs['xmlns:a'] !== A) d.root.attrs['xmlns:a'] = A
  if (e.title !== undefined) {
    chart.kids = chart.kids.filter((k) => !['title', 'autoTitleDeleted'].includes(local(k)))
    const t = String(e.title ?? '').trim()
    chart.kids.unshift(...(t ? [titleXml(t), cv('autoTitleDeleted', 0)] : [cv('autoTitleDeleted', 1)]))
  }
  if (e.type !== undefined || e.stacked !== undefined) {
    const info = parseChart(wb, part)
    if (!info.editable) throw new Error(`${where}: this chart combines kinds Quire cannot change; edit it in Excel`)
    const type = CHART_TYPES.includes(e.type) ? e.type : info.type
    const plot = kid(chart, 'plotArea')
    const old = plot.kids.find((k) => TYPE_OF[local(k)])
    const sers = kids(old, 'ser').map((s, i) => serXml(type, i, { tx: kid(s, 'tx'), cat: kid(s, 'cat') ?? kid(s, 'xVal'), val: kid(s, 'val') ?? kid(s, 'yVal') }, info.series[i]?.color ?? PALETTE[i % PALETTE.length]))
    // The new axes keep the old titles, bounds and gridlines.
    const was = lookOf(plot, old)
    const toPie = type === 'pie' || type === 'doughnut'
    const carry = { ...(was.labels ? { labels: was.labels } : {}), ...(toPie || was.gridlines === undefined ? {} : { x_title: was.xTitle, y_title: was.yTitle, y_min: was.yMin ?? null, y_max: was.yMax ?? null, gridlines: was.gridlines }) }
    for (const [k, v] of Object.entries(carry)) if (e[k] === undefined && v !== undefined) e = { ...e, [k]: v }
    const keep = plot.kids.filter((k) => isEl(k) && ['layout'].includes(local(k)))
    const tail = plot.kids.filter((k) => isEl(k) && ['dTable', 'spPr', 'extLst'].includes(local(k)))
    plot.kids = [...keep, typeXml(type, sers, e.stacked !== undefined ? e.stacked : info.stacked), ...axes(type), ...tail]
    const legend = kid(chart, 'legend')
    const pos = legend && kid(legend, 'legendPos')
    if (pos && (type === 'pie' || type === 'doughnut')) pos.attrs.val = 'r'
  }
  applyLook(chart, e, where)
  wb.pkg.touch(part)
}

const LEGEND_POS = ['r', 'b', 't', 'l', 'tr']

// dLblsXml shows one thing on every point: its value, its share of the
// pie (percent; a value elsewhere) or its category.
function dLblsXml (mode, pie) {
  const show = (k) => cv(k, (mode === 'value' && k === 'showVal') || (mode === 'category' && k === 'showCatName') || (mode === 'percent' && k === (pie ? 'showPercent' : 'showVal')) ? 1 : 0)
  return el('c:dLbls', {}, [show('showLegendKey'), show('showVal'), show('showCatName'), show('showSerName'), show('showPercent'), show('showBubbleSize'), pie ? cv('showLeaderLines', 1) : null])
}
const AXIS_AFTER = ['axId', 'scaling', 'delete', 'axPos', 'majorGridlines', 'minorGridlines', 'title']

// putAxisKid puts node in an axis where the schema wants it: after the
// children named in AXIS_AFTER up to its own name, before the rest.
function putAxisKid (ax, node) {
  const name = local(node)
  ax.kids = ax.kids.filter((k) => local(k) !== name)
  const upto = AXIS_AFTER.slice(0, AXIS_AFTER.indexOf(name))
  let i = 0
  ax.kids.forEach((k, j) => { if (isEl(k) && upto.includes(local(k))) i = j + 1 })
  ax.kids.splice(i, 0, node)
}

// applyLook sets data labels, the legend, axis titles, the value axis
// bounds and its gridlines; a key left out stays as it is.
function applyLook (chart, e, where) {
  const plot = kid(chart, 'plotArea')
  const typeEl = plot?.kids.find((k) => TYPE_OF[local(k)])
  if (e.labels !== undefined && typeEl) {
    const mode = e.labels === true ? 'value' : e.labels || null
    if (mode && !['value', 'percent', 'category', 'none'].includes(mode)) throw new Error(`${where}: labels must be value, percent, category or none`)
    for (const s of kids(typeEl, 'ser')) s.kids = s.kids.filter((k) => local(k) !== 'dLbls')
    typeEl.kids = typeEl.kids.filter((k) => local(k) !== 'dLbls')
    if (mode && mode !== 'none') {
      const d = dLblsXml(mode, /^(pie|doughnut|ofPie|pie3D)Chart$/.test(local(typeEl)))
      const sers = kids(typeEl, 'ser')
      const at = sers.length ? typeEl.kids.indexOf(sers[sers.length - 1]) + 1 : typeEl.kids.findIndex((k) => isEl(k) && !['barDir', 'grouping', 'varyColors', 'scatterStyle', 'radarStyle', 'wireframe'].includes(local(k)))
      typeEl.kids.splice(at < 0 ? typeEl.kids.length : at, 0, d)
    }
  }
  if (e.legend !== undefined) {
    const pos = e.legend === false || e.legend === 'none' ? null : String(e.legend)
    if (pos && !LEGEND_POS.includes(pos)) throw new Error(`${where}: legend must be r, b, t, l, tr or none`)
    let legend = kid(chart, 'legend')
    if (!pos) chart.kids = chart.kids.filter((k) => local(k) !== 'legend')
    else if (legend) { const lp = kid(legend, 'legendPos'); if (lp) lp.attrs.val = pos; else legend.kids.unshift(cv('legendPos', pos)) } else {
      legend = el('c:legend', {}, [cv('legendPos', pos), cv('overlay', 0)])
      chart.kids.splice(chart.kids.indexOf(plot) + 1, 0, legend)
    }
  }
  const { x, y } = chartAxes(plot)
  const axTitle = (ax, t, rot) => {
    if (t === undefined || !ax) return
    const s = String(t ?? '').trim()
    if (s) putAxisKid(ax, titleXml(s, 1000, rot))
    else ax.kids = ax.kids.filter((k) => local(k) !== 'title')
  }
  if ((e.x_title !== undefined && !x) || ((e.y_title !== undefined || e.y_min !== undefined || e.y_max !== undefined || e.gridlines !== undefined) && !y)) throw new Error(`${where}: a pie or doughnut chart has no axes`)
  const horiz = typeEl && val(kid(typeEl, 'barDir')) === 'bar'
  axTitle(x, e.x_title, horiz ? -5400000 : 0)
  axTitle(y, e.y_title, horiz ? 0 : -5400000)
  for (const [key, name] of [['y_min', 'min'], ['y_max', 'max']]) {
    if (e[key] === undefined) continue
    let sc = kid(y, 'scaling')
    if (!sc) { sc = el('c:scaling', {}, [cv('orientation', 'minMax')]); putAxisKid(y, sc) }
    sc.kids = sc.kids.filter((k) => local(k) !== name)
    const v = e[key] === null || e[key] === '' ? undefined : +e[key]
    if (v !== undefined && !Number.isFinite(v)) throw new Error(`${where}: ${key} must be a number`)
    if (v === undefined) continue
    // CT_Scaling: logBase?, orientation?, max?, min?
    const order = ['logBase', 'orientation', 'max', 'min']
    let i = 0
    sc.kids.forEach((k, j) => { if (isEl(k) && order.indexOf(local(k)) < order.indexOf(name)) i = j + 1 })
    sc.kids.splice(i, 0, cv(name, v))
  }
  if (y) {
    const sc = kid(y, 'scaling')
    const lo = +val(kid(sc, 'min')); const hi = +val(kid(sc, 'max'))
    if (Number.isFinite(lo) && Number.isFinite(hi) && kid(sc, 'min') && kid(sc, 'max') && lo >= hi) throw new Error(`${where}: y_min must be below y_max`)
  }
  if (e.gridlines !== undefined) {
    if (e.gridlines) { if (!kid(y, 'majorGridlines')) putAxisKid(y, gridlinesXml()) } else y.kids = y.kids.filter((k) => local(k) !== 'majorGridlines')
  }
}

// refreshCharts rewrites every series cache from the cells it refers to.
export function refreshCharts (wb) {
  for (const s of wb.sheets) {
    let sh
    try { sh = wb.sheet(s.name) } catch { continue }
    for (const x of chartAnchors(wb.pkg, drawingOf(wb.pkg, sh.part, false))) {
      if (!x.part) continue
      const d = wb.pkg.xml(x.part)
      let changed = false
      for (const ref of descendants(d.root, (n) => ['strRef', 'numRef'].includes(local(n)))) {
        const f = textContent(kid(ref, 'f'))
        const shown = local(ref) === 'strRef'
        const vals = resolve(wb, f, shown)
        if (!vals) continue
        const p = ref.name.includes(':') ? ref.name.split(':')[0] + ':' : ''
        const old = kid(ref, shown ? 'strCache' : 'numCache')
        const fmt = !shown && kid(old, 'formatCode')
        const pts = vals.map((v, i) => (v === null || v === undefined ? null : el(p + 'pt', { idx: String(i) }, [el(p + 'v', {}, [text(v)])])))
        const cache = el(p + (shown ? 'strCache' : 'numCache'), {}, [fmt ? clone(fmt) : null, el(p + 'ptCount', { val: String(vals.length) }), ...pts])
        ref.kids = [kid(ref, 'f'), cache, ...ref.kids.filter((k) => isEl(k) && local(k) === 'extLst')]
        changed = true
      }
      if (changed) wb.pkg.touch(x.part)
    }
  }
}

// shiftCharts follows inserted or deleted rows and columns: series
// formulas through fix, and anchors on the sheet that moved.
export function shiftCharts (wb, sh, fix, axis, at, count) {
  for (const s of wb.sheets) {
    let other
    try { other = wb.sheet(s.name) } catch { continue }
    const dr = drawingOf(wb.pkg, other.part, false)
    for (const x of [...chartAnchors(wb.pkg, dr), ...picAnchors(wb.pkg, dr)]) {
      if (x.part && !x.pic) {
        const d = wb.pkg.xml(x.part)
        let changed = false
        for (const f of descendants(d.root, (n) => local(n) === 'f')) {
          const t = textContent(f)
          const u = fix(t)
          if (u !== t) { f.kids = [text(u)]; changed = true }
        }
        if (changed) wb.pkg.touch(x.part)
      }
      if (other !== sh || local(x.anchor) === 'absoluteAnchor') continue
      const del = count < 0 ? -count : 0
      for (const m of [kid(x.anchor, 'from'), kid(x.anchor, 'to')].filter(Boolean)) {
        const node = kid(m, axis === 'r' ? 'row' : 'col')
        const i = +textContent(node) + 1
        let j = i
        if (count > 0) { if (i >= at) j = i + count } else if (i >= at + del) j = i - del
        else if (i >= at) j = at
        if (j !== i) { node.kids = [text(j - 1)]; wb.pkg.touch(dr.part) }
      }
    }
  }
}

// ------------------------------------------------------------- pictures

const PIC_CT = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif' }

// picAnchors lists the anchors on a drawing that hold a picture, with the
// media part each shows.
function picAnchors (pkg, dr) {
  if (!dr) return []
  const out = []
  for (const a of dr.root.kids) {
    if (!['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(local(a))) continue
    const pic = kid(a, 'pic')
    if (!pic) continue
    const blip = descendants(pic, (n) => local(n) === 'blip')[0]
    const rid = blip && Object.entries(blip.attrs).find(([k]) => k.endsWith(':embed'))?.[1]
    const rel = rid && pkg.rel(dr.part, rid)
    out.push({ anchor: a, rid, part: rel?.part && pkg.has(rel.part) ? rel.part : undefined, pic: true, name: descendants(pic, (n) => local(n) === 'cNvPr')[0]?.attrs.name })
  }
  return out
}

// sheetPics gives the editor each picture's place and its image.
export function sheetPics (wb, sh) {
  const dr = drawingOf(wb.pkg, sh.part, false)
  if (!dr) return []
  const g = geometry(sh)
  return picAnchors(wb.pkg, dr).map((x, k) => {
    let src
    if (x.part) {
      const ext = x.part.split('.').pop().toLowerCase().replace('jpg', 'jpeg')
      const data = wb.pkg.get(x.part)
      if (data && data.length <= 8 * 1024 * 1024 && PIC_CT[ext]) src = `data:${PIC_CT[ext]};base64,${Buffer.from(data).toString('base64')}`
    }
    return { k, ...frameOf(x.anchor, g), name: x.name, src }
  })
}

function picAt (wb, sh, k, where) {
  const dr = drawingOf(wb.pkg, sh.part, false)
  const hit = picAnchors(wb.pkg, dr)[+k]
  if (!hit) throw new Error(`${where}: sheet "${sh.info.name}" has no picture ${k}; reload`)
  return { dr, ...hit }
}

export function addPic (wb, sh, e, where) {
  const m = /^data:image\/(png|jpe?g|gif);base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(e.src ?? ''))
  if (!m) throw new Error(`${where}: only PNG, JPEG and GIF pictures`)
  const data = Buffer.from(m[2], 'base64')
  if (data.length > 20 * 1024 * 1024) throw new Error(`${where}: 图片超过 20 MB`)
  const pkg = wb.pkg
  const dr = drawingOf(pkg, sh.part, true, sh)
  if (dr.root.attrs['xmlns:xdr'] !== XDR) dr.root.attrs['xmlns:xdr'] = XDR
  if (dr.root.attrs['xmlns:a'] !== A) dr.root.attrs['xmlns:a'] = A
  if (!dr.root.attrs['xmlns:r']) dr.root.attrs['xmlns:r'] = R
  const media = pkg.addMedia('xl/media', m[1].toLowerCase().replace('jpg', 'jpeg'), data)
  const rid = pkg.addRel(dr.part, REL('image'), media)
  const g = geometry(sh)
  const at = e.at && e.at.c >= 1 && e.at.r >= 1 ? { c: +e.at.c, r: +e.at.r } : { c: 1, r: 1 }
  const from = { c: at.c, dx: 0, r: at.r, dy: 0 }
  const w = Math.max(16, Math.min(4000, Math.round(+e.w || 320))); const h = Math.max(16, Math.min(4000, Math.round(+e.h || 240)))
  const x = walk(from.c, 0, w, g.col); const y = walk(from.r, 0, h, g.row)
  const id = Math.max(1, ...descendants(dr.root, (n) => local(n) === 'cNvPr').map((n) => +n.attrs.id || 0)) + 1
  dr.root.kids.push(el('xdr:twoCellAnchor', { editAs: 'oneCell' }, [
    markerXml('xdr:', 'from', from), markerXml('xdr:', 'to', { c: x.i, dx: x.d, r: y.i, dy: y.d }),
    el('xdr:pic', {}, [
      el('xdr:nvPicPr', {}, [el('xdr:cNvPr', { id: String(id), name: String(e.name || `图片 ${id - 1}`) }), el('xdr:cNvPicPr', {}, [el('a:picLocks', { noChangeAspect: '1' })])]),
      el('xdr:blipFill', {}, [el('a:blip', { 'xmlns:r': R, 'r:embed': rid }), el('a:stretch', {}, [el('a:fillRect')])]),
      el('xdr:spPr', {}, [el('a:xfrm', {}, [el('a:off', { x: '0', y: '0' }), el('a:ext', { cx: String(w * EMU), cy: String(h * EMU) })]), el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')])])
    ]),
    el('xdr:clientData')
  ]))
  pkg.touch(dr.part)
}

export function movePic (wb, sh, e, where) {
  const { dr, anchor } = picAt(wb, sh, e.pic, where)
  const f = e.from; const t = e.to
  if (!(f?.c >= 1 && f?.r >= 1 && t?.c >= 1 && t?.r >= 1)) throw new Error(`${where}: from and to must be cells with offsets`)
  const p = anchor.name.includes(':') ? anchor.name.split(':')[0] + ':' : ''
  const body = anchor.kids.filter((k) => isEl(k) && !['from', 'to', 'pos', 'ext'].includes(local(k)))
  anchor.name = p + 'twoCellAnchor'
  if (!anchor.attrs.editAs) anchor.attrs.editAs = 'oneCell'
  anchor.kids = [markerXml(p, 'from', f), markerXml(p, 'to', t), ...body]
  wb.pkg.touch(dr.part)
}

export function deletePic (wb, sh, e, where) {
  const { dr, anchor, rid, part } = picAt(wb, sh, e.pic, where)
  remove(dr.root, anchor)
  if (rid && !picAnchors(wb.pkg, dr).some((x) => x.rid === rid)) {
    wb.pkg.removeRel(dr.part, rid)
    // Other drawings may show the same image.
    const used = part && [...wb.pkg.names?.() ?? []].some((n) => n.endsWith('.rels') && n !== wb.pkg.relsName(dr.part) && wb.pkg.rels(n.replace(/_rels\/(.*)\.rels$/, '$1')).some((r) => r.part === part))
    if (part && !used) wb.pkg.delete(part)
  }
  wb.pkg.touch(dr.part)
}
