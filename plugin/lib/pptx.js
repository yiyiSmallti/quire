// PowerPoint presentations: read slides and shapes, create decks, edit
// slides, render for the preview pane.
import path from 'node:path/posix'
import { Package, NS, REL, dataUri } from './pkg.js'
import { el, text, doc, child, children, descendants, path as xpath, textContent, remove, clone, isEl, insertOrdered } from './xml.js'
import { inline } from './md.js'
import { coreProps, replaceInParts } from './docx.js'
import { imageInfo } from './image.js'

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const NSA = { 'xmlns:a': A, 'xmlns:r': NS.r, 'xmlns:p': P }
const CT = {
  pres: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  notes: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
  presProps: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
  tableStyles: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'
}
const SW = 12192000
const SH = 6858000

// ------------------------------------------------------------- template

const xfrm = (x, y, cx, cy) => el('a:xfrm', {}, [el('a:off', { x: String(Math.round(x)), y: String(Math.round(y)) }), el('a:ext', { cx: String(Math.round(cx)), cy: String(Math.round(cy)) })])
const grpRoot = () => [
  el('p:nvGrpSpPr', {}, [el('p:cNvPr', { id: '1', name: '' }), el('p:cNvGrpSpPr'), el('p:nvPr')]),
  el('p:grpSpPr', {}, [el('a:xfrm', {}, [el('a:off', { x: '0', y: '0' }), el('a:ext', { cx: '0', cy: '0' }), el('a:chOff', { x: '0', y: '0' }), el('a:chExt', { cx: '0', cy: '0' })])])
]
const solid = (hexv) => el('a:solidFill', {}, [el('a:srgbClr', { val: hexv })])
const scheme = (val) => el('a:solidFill', {}, [el('a:schemeClr', { val })])

function phShape (id, name, ph, box, { anchor, lst, prompt } = {}) {
  return el('p:sp', {}, [
    el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id), name }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1' })]), el('p:nvPr', {}, [el('p:ph', ph)])]),
    el('p:spPr', {}, box ? [xfrm(...box)] : []),
    el('p:txBody', {}, [
      el('a:bodyPr', anchor ? { anchor } : {}, anchor ? [el('a:normAutofit')] : []),
      el('a:lstStyle', {}, lst ?? []),
      el('a:p', {}, prompt ? [el('a:r', {}, [el('a:rPr', { lang: 'en-US' }), el('a:t', {}, [text(prompt)])])] : [el('a:endParaRPr', { lang: 'en-US' })])
    ])
  ])
}

const lvl = (n, attrs, kids) => el(`a:lvl${n}pPr`, attrs, kids)
const defR = (sz, extra = []) => el('a:defRPr', { sz: String(sz) }, extra)

function themeDoc (t) {
  const accent = t.accent
  const font = t.font
  const c = (name, val) => el('a:' + name, {}, [el('a:srgbClr', { val })])
  const tri = (make) => [make(), make(), make()]
  return doc(el('a:theme', { 'xmlns:a': A, name: 'Quire' }, [
    el('a:themeElements', {}, [
      el('a:clrScheme', { name: 'Quire' }, [
        el('a:dk1', {}, [el('a:srgbClr', { val: t.text })]), el('a:lt1', {}, [el('a:srgbClr', { val: 'FFFFFF' })]),
        c('dk2', accent), c('lt2', 'E7E6E6'),
        c('accent1', accent), c('accent2', 'ED7D31'), c('accent3', 'A5A5A5'), c('accent4', 'FFC000'), c('accent5', '5B9BD5'), c('accent6', '70AD47'),
        c('hlink', '0563C1'), c('folHlink', '954F72')
      ]),
      el('a:fontScheme', { name: 'Quire' }, ['a:majorFont', 'a:minorFont'].map((n) => el(n, {}, [
        el('a:latin', { typeface: font.latin }), el('a:ea', { typeface: font.ea }), el('a:cs', { typeface: '' })
      ]))),
      el('a:fmtScheme', { name: 'Quire' }, [
        el('a:fillStyleLst', {}, tri(() => scheme('phClr'))),
        el('a:lnStyleLst', {}, tri(() => el('a:ln', { w: '6350' }, [scheme('phClr')]))),
        el('a:effectStyleLst', {}, tri(() => el('a:effectStyle', {}, [el('a:effectLst')]))),
        el('a:bgFillStyleLst', {}, tri(() => scheme('phClr')))
      ])
    ]),
    el('a:objectDefaults'),
    el('a:extraClrSchemeLst')
  ]))
}

const CLRMAP = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink' }

function masterDoc (bg, layouts) {
  const bullet = (n, marL, sz) => lvl(n, { marL: String(marL), indent: '-228600', algn: 'l', defTabSz: '914400', rtl: '0', eaLnBrk: '1', latinLnBrk: '0', hangingPunct: '1' }, [
    el('a:lnSpc', {}, [el('a:spcPct', { val: '100000' })]), el('a:spcBef', {}, [el('a:spcPts', { val: '1000' })]),
    el('a:buFont', { typeface: 'Arial', panose: '020B0604020202020204', pitchFamily: '34', charset: '0' }), el('a:buChar', { char: n === 1 ? '•' : n === 2 ? '–' : '•' }),
    defR(sz, [scheme('tx1'), el('a:latin', { typeface: '+mn-lt' }), el('a:ea', { typeface: '+mn-ea' }), el('a:cs', { typeface: '+mn-cs' })])
  ])
  return doc(el('p:sldMaster', NSA, [
    el('p:cSld', {}, [
      el('p:bg', {}, [el('p:bgPr', {}, [solid(bg), el('a:effectLst')])]),
      el('p:spTree', {}, [...grpRoot(),
        phShape(2, 'Title Placeholder 1', { type: 'title' }, [838200, 365125, 10515600, 1325563], { anchor: 'ctr', prompt: 'Click to edit Master title style' }),
        phShape(3, 'Text Placeholder 2', { type: 'body', idx: '1' }, [838200, 1825625, 10515600, 4351338], { anchor: 't', prompt: 'Click to edit Master text styles' })
      ])
    ]),
    el('p:clrMap', { ...CLRMAP }),
    el('p:sldLayoutIdLst', {}, layouts.map((rid, i) => el('p:sldLayoutId', { id: String(2147483649 + i), 'r:id': rid }))),
    el('p:txStyles', {}, [
      el('p:titleStyle', {}, [lvl(1, { algn: 'l', defTabSz: '914400', rtl: '0', eaLnBrk: '1', latinLnBrk: '0', hangingPunct: '1' }, [
        el('a:lnSpc', {}, [el('a:spcPct', { val: '90000' })]), el('a:spcBef', {}, [el('a:spcPct', { val: '0' })]), el('a:buNone'),
        defR(4000, [el('a:b'), scheme('tx2'), el('a:latin', { typeface: '+mj-lt' }), el('a:ea', { typeface: '+mj-ea' }), el('a:cs', { typeface: '+mj-cs' })].filter((n) => n.name !== 'a:b'))
      ])]),
      el('p:bodyStyle', {}, [bullet(1, 228600, 2400), bullet(2, 685800, 2000), bullet(3, 1143000, 1800), bullet(4, 1600200, 1600), bullet(5, 2057400, 1600)]),
      el('p:otherStyle', {}, [el('a:defPPr', {}, [el('a:defRPr', { lang: 'en-US' })]), lvl(1, { marL: '0', algn: 'l', defTabSz: '914400' }, [defR(1800, [scheme('tx1'), el('a:latin', { typeface: '+mn-lt' }), el('a:ea', { typeface: '+mn-ea' })])])])
    ])
  ]))
}

// Layout definitions: [file, type, name, placeholders].
const LAYOUTS = [
  ['title', 'Title Slide', () => [
    phShape(2, 'Title 1', { type: 'ctrTitle' }, [1524000, 1122363, 9144000, 2387600], { anchor: 'b', lst: [lvl(1, { algn: 'ctr' }, [defR(5400)])] }),
    phShape(3, 'Subtitle 2', { type: 'subTitle', idx: '1' }, [1524000, 3602038, 9144000, 1655762], { lst: [lvl(1, { marL: '0', indent: '0', algn: 'ctr' }, [el('a:buNone'), defR(2400, [scheme('tx1')])])] })
  ]],
  ['obj', 'Title and Content', () => [
    phShape(2, 'Title 1', { type: 'title' }),
    phShape(3, 'Content Placeholder 2', { idx: '1' })
  ]],
  ['secHead', 'Section Header', () => [
    phShape(2, 'Title 1', { type: 'title' }, [831850, 1709738, 10515600, 2852737], { anchor: 'b', lst: [lvl(1, {}, [defR(5400)])] }),
    phShape(3, 'Text Placeholder 2', { type: 'body', idx: '1' }, [831850, 4589463, 10515600, 1500187], { lst: [lvl(1, { marL: '0', indent: '0' }, [el('a:buNone'), defR(2400, [el('a:solidFill', {}, [el('a:schemeClr', { val: 'tx1' }, [el('a:tint', { val: '75000' })])])])])] })
  ]],
  ['twoObj', 'Two Content', () => [
    phShape(2, 'Title 1', { type: 'title' }),
    phShape(3, 'Content Placeholder 2', { sz: 'half', idx: '1' }, [838200, 1825625, 5181600, 4351338]),
    phShape(4, 'Content Placeholder 3', { sz: 'half', idx: '2' }, [6172200, 1825625, 5181600, 4351338])
  ]],
  ['titleOnly', 'Title Only', () => [phShape(2, 'Title 1', { type: 'title' })]],
  ['blank', 'Blank', () => []]
]

function notesMasterDoc () {
  return doc(el('p:notesMaster', NSA, [
    el('p:cSld', {}, [el('p:bg', {}, [el('p:bgRef', { idx: '1001' }, [el('a:schemeClr', { val: 'bg1' })])]), el('p:spTree', {}, [...grpRoot(),
      el('p:sp', {}, [
        el('p:nvSpPr', {}, [el('p:cNvPr', { id: '2', name: 'Slide Image Placeholder 1' }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1', noRot: '1', noChangeAspect: '1' })]), el('p:nvPr', {}, [el('p:ph', { type: 'sldImg', idx: '2' })])]),
        el('p:spPr', {}, [xfrm(685800, 1143000, 5486400, 3086100), el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]), el('a:noFill'), el('a:ln', { w: '12700' }, [el('a:solidFill', {}, [el('a:prstClr', { val: 'black' })])])])
      ]),
      phShape(3, 'Notes Placeholder 2', { type: 'body', sz: 'quarter', idx: '3' }, [685800, 4400550, 5486400, 3600450], { prompt: 'Click to edit Master text styles' })
    ])]),
    el('p:clrMap', { ...CLRMAP }),
    el('p:notesStyle', {}, [lvl(1, { marL: '0', algn: 'l', defTabSz: '914400' }, [defR(1200, [scheme('tx1'), el('a:latin', { typeface: '+mn-lt' }), el('a:ea', { typeface: '+mn-ea' })])])])
  ]))
}

function newPackage (theme) {
  const p = new Package([])
  p.setXml('[Content_Types].xml', doc(el('Types', { xmlns: NS.ct }, [
    el('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }),
    el('Default', { Extension: 'xml', ContentType: 'application/xml' }),
    el('Override', { PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }),
    el('Override', { PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' })
  ])))
  p.setXml('_rels/.rels', doc(el('Relationships', { xmlns: NS.rel }, [
    el('Relationship', { Id: 'rId1', Type: NS.officeDocument, Target: 'ppt/presentation.xml' }),
    el('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: 'docProps/core.xml' }),
    el('Relationship', { Id: 'rId3', Type: REL('extended-properties'), Target: 'docProps/app.xml' })
  ])))
  p.setXml('docProps/core.xml', coreProps())
  p.setXml('docProps/app.xml', doc(el('Properties', { xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' }, [el('Application', {}, [text('Quire')])])))
  const pres = 'ppt/presentation.xml'
  p.setXml(pres, doc(el('p:presentation', { ...NSA, saveSubsetFonts: '1' }, [
    el('p:sldMasterIdLst'), el('p:notesMasterIdLst'), el('p:sldIdLst'),
    el('p:sldSz', { cx: String(SW), cy: String(SH) }), el('p:notesSz', { cx: '6858000', cy: '9144000' })
  ])))
  p.setOverride(pres, CT.pres)
  p.setXml('ppt/theme/theme1.xml', themeDoc(theme))
  p.setOverride('ppt/theme/theme1.xml', CT.theme)
  // Layouts first: the master lists them.
  const master = 'ppt/slideMasters/slideMaster1.xml'
  const layoutRids = []
  LAYOUTS.forEach(([type, name, phs], i) => {
    const part = `ppt/slideLayouts/slideLayout${i + 1}.xml`
    p.setXml(part, doc(el('p:sldLayout', { ...NSA, type, preserve: '1' }, [
      el('p:cSld', { name }, [el('p:spTree', {}, [...grpRoot(), ...phs()])]),
      el('p:clrMapOvr', {}, [el('a:masterClrMapping')])
    ])))
    p.setOverride(part, CT.layout)
    p.addRel(part, REL('slideMaster'), master)
    layoutRids.push(p.addRel(master, REL('slideLayout'), part))
  })
  p.setXml(master, masterDoc(theme.background, layoutRids))
  p.setOverride(master, CT.master)
  p.addRel(master, REL('theme'), 'ppt/theme/theme1.xml')
  const pd = p.xml(pres)
  child(pd.root, 'p:sldMasterIdLst').kids.push(el('p:sldMasterId', { id: '2147483648', 'r:id': p.addRel(pres, REL('slideMaster'), master) }))
  p.setXml('ppt/theme/theme2.xml', themeDoc(theme))
  p.setOverride('ppt/theme/theme2.xml', CT.theme)
  p.setXml('ppt/notesMasters/notesMaster1.xml', notesMasterDoc())
  p.setOverride('ppt/notesMasters/notesMaster1.xml', CT.notesMaster)
  p.addRel('ppt/notesMasters/notesMaster1.xml', REL('theme'), 'ppt/theme/theme2.xml')
  child(pd.root, 'p:notesMasterIdLst').kids.push(el('p:notesMasterId', { 'r:id': p.addRel(pres, REL('notesMaster'), 'ppt/notesMasters/notesMaster1.xml') }))
  p.addRel(pres, REL('theme'), 'ppt/theme/theme1.xml')
  p.setXml('ppt/presProps.xml', doc(el('p:presentationPr', NSA)))
  p.setOverride('ppt/presProps.xml', CT.presProps)
  p.addRel(pres, REL('presProps'), 'ppt/presProps.xml')
  p.setXml('ppt/tableStyles.xml', doc(el('a:tblStyleLst', { 'xmlns:a': A, def: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' })))
  p.setOverride('ppt/tableStyles.xml', CT.tableStyles)
  p.addRel(pres, REL('tableStyles'), 'ppt/tableStyles.xml')
  return p
}

// ------------------------------------------------------------- deck model

class Deck {
  constructor (pkg) {
    this.pkg = pkg
    this.main = pkg.mainPart()
    if (!this.main || !pkg.has(this.main)) throw new Error('not a PowerPoint presentation: no presentation part')
    this.pres = pkg.xml(this.main)
    if (!/presentation$/.test(this.pres.root.name)) throw new Error('not a PowerPoint presentation')
    const sz = child(this.pres.root, 'p:sldSz')
    this.cx = +(sz?.attrs.cx ?? SW)
    this.cy = +(sz?.attrs.cy ?? SH)
    this.idList = child(this.pres.root, 'p:sldIdLst')
    if (!this.idList) {
      this.idList = el('p:sldIdLst')
      const at = this.pres.root.kids.findIndex((k) => k.name === 'p:sldSz')
      this.pres.root.kids.splice(at < 0 ? this.pres.root.kids.length : at, 0, this.idList)
    }
  }

  slides () {
    return children(this.idList, 'p:sldId').map((n, i) => {
      const rel = this.pkg.rel(this.main, n.attrs['r:id'])
      return { n: i + 1, node: n, id: n.attrs.id, rid: n.attrs['r:id'], part: rel?.part }
    })
  }

  slide (n, where) {
    const all = this.slides()
    const s = all[(+n) - 1]
    if (!s) throw new Error(`${where}: there is no slide ${n}; the deck has ${all.length} slides (numbering starts at 1)`)
    return s
  }

  layoutOf (slidePart) { return this.pkg.relsOfType(slidePart, 'slideLayout')[0]?.part }
  masterOf (layoutPart) { return this.pkg.relsOfType(layoutPart, 'slideMaster')[0]?.part }
  notesOf (slidePart) { return this.pkg.relsOfType(slidePart, 'notesSlide')[0]?.part }

  layouts () {
    const out = []
    const masters = children(child(this.pres.root, 'p:sldMasterIdLst'), 'p:sldMasterId').map((m) => this.pkg.rel(this.main, m.attrs['r:id'])?.part).filter(Boolean)
    for (const m of masters) {
      for (const lid of children(child(this.pkg.xml(m).root, 'p:sldLayoutIdLst'), 'p:sldLayoutId')) {
        const part = this.pkg.rel(m, lid.attrs['r:id'])?.part
        if (!part || !this.pkg.has(part)) continue
        const d = this.pkg.xml(part)
        out.push({ part, type: d.root.attrs.type ?? 'cust', name: child(d.root, 'p:cSld')?.attrs.name ?? '' })
      }
    }
    return out
  }

  // pickLayout finds the deck's own layout for a kind of slide.
  pickLayout (kind) {
    const want = {
      title: ['title', 'ctrTitle'],
      section: ['secHead', 'title'],
      content: ['obj', 'tx', 'txAndObj', 'objTx', 'twoObj'],
      two_column: ['twoObj', 'twoTxTwoObj', 'twoColTx', 'obj', 'tx'],
      blank: ['blank', 'titleOnly', 'obj'],
      title_only: ['titleOnly', 'obj', 'blank']
    }[kind] ?? ['obj']
    const all = this.layouts()
    if (!all.length) throw new Error('the presentation has no slide layouts')
    for (const t of want) {
      const hit = all.find((l) => l.type === t)
      if (hit) return hit
    }
    // Custom layouts: go by placeholders.
    const score = (l) => {
      const phs = placeholders(this.pkg.xml(l.part).root)
      const content = phs.filter((p) => isContent(p.type)).length
      if (kind === 'blank') return -phs.length
      if (kind === 'two_column') return content >= 2 ? 10 : content
      if (kind === 'title' || kind === 'section') return phs.some((p) => p.type === 'ctrTitle' || p.type === 'subTitle') ? 10 : phs.some((p) => p.type === 'title') ? 1 : 0
      return content === 1 ? 10 : content
    }
    return [...all].sort((a, b) => score(b) - score(a))[0]
  }

  nextShapeId (spTree) {
    return Math.max(1, ...descendants(spTree, 'p:cNvPr').map((n) => +n.attrs.id || 0)) + 1
  }
}

const isContent = (type) => !type || type === 'obj' || type === 'body' || type === 'tx'
const isTitle = (type) => type === 'title' || type === 'ctrTitle'

function placeholders (root) {
  const tree = xpath(root, 'p:cSld', 'p:spTree')
  return descendants(tree, 'p:sp').map((sp) => {
    const ph = xpath(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph')
    return ph ? { sp, ph, type: ph.attrs.type, idx: ph.attrs.idx ?? '0' } : undefined
  }).filter(Boolean)
}

// shapes lists a slide's shapes in tree order, descending into groups.
function shapes (spTree) {
  const out = []
  const walk = (node) => {
    for (const k of node.kids) {
      if (!isEl(k)) continue
      if (k.name === 'p:sp' || k.name === 'p:pic' || k.name === 'p:graphicFrame') out.push(k)
      else if (k.name === 'p:grpSp') walk(k)
      else if (k.name === 'mc:AlternateContent') {
        const pick = child(k, 'mc:Choice') ?? child(k, 'mc:Fallback')
        if (pick) walk(pick)
      }
    }
  }
  walk(spTree)
  return out
}

const nv = (shape) => child(shape, 'p:nvSpPr') ?? child(shape, 'p:nvPicPr') ?? child(shape, 'p:nvGraphicFramePr')
const phOf = (shape) => xpath(nv(shape), 'p:nvPr', 'p:ph')

function paraText (p) {
  let s = ''
  for (const k of p.kids) {
    if (k.name === 'a:r' || k.name === 'a:fld') s += textContent(child(k, 'a:t'))
    else if (k.name === 'a:br') s += '\n'
  }
  return s
}

// ------------------------------------------------------------- reading

const LIMIT = 16000

function parseSlides (spec, total) {
  if (spec === undefined || spec === null || spec === '') return undefined
  const set = new Set()
  for (const part of String(spec).split(',')) {
    const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(part)
    if (!m) throw new Error(`slides "${spec}" should look like 3, 2-5 or 1,4`)
    for (let i = +m[1]; i <= +(m[2] ?? m[1]); i++) if (i >= 1 && i <= total) set.add(i)
  }
  return set
}

export function readPptx (buf, { slides: spec, from } = {}) {
  const deck = new Deck(Package.fromBuffer(buf))
  const all = deck.slides()
  const pick = parseSlides(spec, all.length)
  const out = [`PowerPoint presentation: ${all.length} slides, ${(deck.cx / 360000).toFixed(1)}×${(deck.cy / 360000).toFixed(1)} cm. Refs [sN.k] name shapes for pptx_edit set_text.`]
  const layouts = deck.layouts()
  out.push('Layouts: ' + layouts.map((l) => `${l.name || l.type}`).join(', '))
  let size = out.join('\n').length
  let next
  for (const s of all) {
    if (pick && !pick.has(s.n)) continue
    if (from && s.n < +from) continue
    const lines = [slideText(deck, s)]
    const block = lines.join('\n')
    if (size + block.length > LIMIT && out.length > 2) { next = s.n; break }
    out.push('', block)
    size += block.length + 2
  }
  if (next) out.push('', `[more slides: call office_read with from: ${next}]`)
  return out.join('\n')
}

function slideText (deck, s) {
  const pkg = deck.pkg
  if (!s.part || !pkg.has(s.part)) return `## Slide ${s.n} (missing part)`
  const d = pkg.xml(s.part)
  const layout = deck.layoutOf(s.part)
  const lname = layout && pkg.has(layout) ? child(pkg.xml(layout).root, 'p:cSld')?.attrs.name : ''
  const lines = [`## Slide ${s.n}${lname ? ' · ' + lname : ''}${d.root.attrs.show === '0' ? ' (hidden)' : ''}`]
  const tree = xpath(d.root, 'p:cSld', 'p:spTree')
  shapes(tree).forEach((sh, i) => {
    const ref = `[s${s.n}.${i + 1}]`
    const ph = phOf(sh)
    const name = xpath(nv(sh), 'p:cNvPr')?.attrs.name ?? ''
    const label = ph ? (ph.attrs.type ?? 'content') : name
    if (sh.name === 'p:pic') { lines.push(`${ref} picture "${xpath(nv(sh), 'p:cNvPr')?.attrs.descr || name}"`); return }
    if (sh.name === 'p:graphicFrame') {
      const tbl = descendants(sh, 'a:tbl')[0]
      if (tbl) {
        const rows = children(tbl, 'a:tr').map((tr) => children(tr, 'a:tc').map((tc) => children(child(tc, 'a:txBody'), 'a:p').map(paraText).join(' ').replace(/\|/g, '\\|')))
        lines.push(`${ref} table ${rows.length}×${rows[0]?.length ?? 0} "${name}"`)
        rows.forEach((r, k) => { lines.push('| ' + r.join(' | ') + ' |'); if (k === 0) lines.push('|' + ' --- |'.repeat(r.length)) })
        return
      }
      const uri = descendants(sh, 'a:graphicData')[0]?.attrs.uri ?? ''
      lines.push(`${ref} ${/chart/.test(uri) ? 'chart' : /diagram/.test(uri) ? 'SmartArt' : 'object'} "${name}"`)
      return
    }
    const body = child(sh, 'p:txBody')
    const paras = children(body, 'a:p')
    const texts = paras.map((p) => ({ lvl: +(child(p, 'a:pPr')?.attrs.lvl ?? 0), t: paraText(p) }))
    if (!texts.some((x) => x.t.trim())) {
      lines.push(`${ref} ${label}: (empty)`)
      return
    }
    if (texts.length === 1 || isTitle(ph?.attrs.type)) lines.push(`${ref} ${label}: ${texts.map((x) => x.t).join(' / ')}`)
    else {
      lines.push(`${ref} ${label}:`)
      for (const x of texts) if (x.t.trim()) lines.push(`${'  '.repeat(x.lvl + 1)}- ${x.t}`)
    }
  })
  const notes = deck.notesOf(s.part)
  if (notes && pkg.has(notes)) {
    const body = placeholders(pkg.xml(notes).root).find((p) => p.type === 'body')
    const t = body ? children(child(body.sp, 'p:txBody'), 'a:p').map(paraText).join('\n').trim() : ''
    if (t) lines.push(`notes: ${t.replace(/\n/g, ' / ')}`)
  }
  return lines.join('\n')
}

// ------------------------------------------------------------- building slides

function runsXml (str, rpr = {}) {
  return inline(String(str)).map((r) => {
    if (r.br) return el('a:br', {}, [el('a:rPr', { lang: 'zh-CN', altLang: 'en-US' })])
    const attrs = { lang: 'zh-CN', altLang: 'en-US', ...rpr }
    if (r.b) attrs.b = '1'
    if (r.i) attrs.i = '1'
    if (r.s) attrs.strike = 'sngStrike'
    return el('a:r', {}, [el('a:rPr', attrs), el('a:t', {}, [text(r.text ?? r.alt ?? '')])])
  })
}

function paragraphs (items, base = {}) {
  return items.map((it) => {
    const o = typeof it === 'object' && it !== null ? it : { text: it }
    const lv = Math.max(0, Math.min(8, +(o.level ?? 0)))
    return el('a:p', {}, [lv ? el('a:pPr', { lvl: String(lv) }) : null, ...runsXml(o.text ?? '', base)])
  })
}

// textBody makes a placeholder's text body from paragraphs.
const textBody = (paras) => el('p:txBody', {}, [el('a:bodyPr'), el('a:lstStyle'), ...(paras.length ? paras : [el('a:p', {}, [el('a:endParaRPr', { lang: 'zh-CN', altLang: 'en-US' })])])])

function geometry (deck, sp, layoutPart) {
  const own = xfrmOf(sp)
  if (own) return own
  const ph = phOf(sp)
  if (!ph || !layoutPart) return undefined
  const find = (root) => {
    const phs = placeholders(root)
    const type = ph.attrs.type
    const idx = ph.attrs.idx ?? '0'
    return phs.find((p) => p.idx === idx && idx !== '0') ??
      phs.find((p) => (p.type ?? 'obj') === (type ?? 'obj')) ??
      phs.find((p) => (isTitle(type) && isTitle(p.type)) || (isContent(type) && isContent(p.type)) || (type === 'subTitle' && isContent(p.type)))
  }
  const lp = find(deck.pkg.xml(layoutPart).root)
  if (lp && xfrmOf(lp.sp)) return xfrmOf(lp.sp)
  const master = deck.masterOf(layoutPart)
  if (master) {
    const phs = placeholders(deck.pkg.xml(master).root)
    const type = ph.attrs.type
    const mp = phs.find((p) => isTitle(type) ? p.type === 'title' : (type === 'dt' || type === 'ftr' || type === 'sldNum') ? p.type === type : p.type === 'body')
    if (mp && xfrmOf(mp.sp)) return xfrmOf(mp.sp)
  }
  return undefined
}

function xfrmOf (shape) {
  const x = shape.name === 'p:graphicFrame' ? child(shape, 'p:xfrm') : xpath(shape, 'p:spPr', 'a:xfrm') ?? xpath(shape, 'p:grpSpPr', 'a:xfrm')
  const off = child(x, 'a:off'); const ext = child(x, 'a:ext')
  if (!off || !ext) return undefined
  return { x: +off.attrs.x, y: +off.attrs.y, cx: +ext.attrs.cx, cy: +ext.attrs.cy, rot: +(x.attrs.rot ?? 0) / 60000, flipH: x.attrs.flipH === '1' }
}

function tableFrame (id, box, spec, accent) {
  const rows = spec.rows ?? []
  if (!rows.length) throw new Error('table.rows is empty')
  const cols = Math.max(...rows.map((r) => r.length))
  const header = spec.header !== false
  const rowH = Math.min(Math.floor(box.cy / rows.length), 457200)
  const sz = rows.length > 10 ? 1100 : rows.length > 6 ? 1200 : 1400
  const colW = Math.floor(box.cx / cols)
  const line = (n) => el(n, { w: '9525' }, [solid('BFBFBF')])
  const trs = rows.map((r, i) => el('a:tr', { h: String(rowH) }, Array.from({ length: cols }, (_, c) => {
    const head = header && i === 0
    const rpr = { sz: String(sz) }
    const runs = runsXml(r[c] ?? '', head ? { ...rpr, b: '1' } : rpr)
    if (head) for (const run of runs) { const pr = child(run, 'a:rPr'); if (pr) pr.kids.push(solid('FFFFFF')) }
    return el('a:tc', {}, [
      el('a:txBody', {}, [el('a:bodyPr'), el('a:lstStyle'), el('a:p', {}, runs.length ? runs : [el('a:endParaRPr', { lang: 'zh-CN', altLang: 'en-US', sz: String(sz) })])]),
      el('a:tcPr', { anchor: 'ctr' }, [line('a:lnL'), line('a:lnR'), line('a:lnT'), line('a:lnB'), solid(head ? accent : i % 2 ? 'FFFFFF' : 'F2F2F2')])
    ])
  })))
  return el('p:graphicFrame', {}, [
    el('p:nvGraphicFramePr', {}, [el('p:cNvPr', { id: String(id), name: `Table ${id}` }), el('p:cNvGraphicFramePr', {}, [el('a:graphicFrameLocks', { noGrp: '1' })]), el('p:nvPr')]),
    el('p:xfrm', {}, [el('a:off', { x: String(box.x), y: String(box.y) }), el('a:ext', { cx: String(colW * cols), cy: String(rowH * rows.length) })]),
    el('a:graphic', {}, [el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/table' }, [
      el('a:tbl', {}, [el('a:tblPr', { firstRow: header ? '1' : '0', bandRow: '1' }), el('a:tblGrid', {}, Array.from({ length: cols }, () => el('a:gridCol', { w: String(colW) }))), ...trs])
    ])])
  ])
}

async function pictureShape (deck, slidePart, id, box, src, io) {
  const data = await io.loadImage(src)
  const info = imageInfo(data)
  if (!info) throw new Error(`"${src}" is not a PNG, JPEG or GIF image`)
  const media = deck.pkg.addMedia('ppt/media', info.ext, data)
  const rid = deck.pkg.addRel(slidePart, NS.image, media)
  const ratio = info.width / info.height
  let cx = box.cx; let cy = Math.round(cx / ratio)
  if (cy > box.cy) { cy = box.cy; cx = Math.round(cy * ratio) }
  const x = box.x + Math.round((box.cx - cx) / 2); const y = box.y + Math.round((box.cy - cy) / 2)
  return el('p:pic', {}, [
    el('p:nvPicPr', {}, [el('p:cNvPr', { id: String(id), name: `Picture ${id}`, descr: src.startsWith('data:') ? '' : path.basename(src) }), el('p:cNvPicPr', {}, [el('a:picLocks', { noChangeAspect: '1' })]), el('p:nvPr')]),
    el('p:blipFill', {}, [el('a:blip', { 'r:embed': rid }), el('a:stretch', {}, [el('a:fillRect')])]),
    el('p:spPr', {}, [xfrm(x, y, cx, cy), el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')])])
  ])
}

function inferKind (s) {
  if (s.layout) {
    const k = String(s.layout).toLowerCase().replace(/[\s-]/g, '_')
    if (!['title', 'section', 'content', 'two_column', 'blank', 'title_only'].includes(k)) throw new Error(`layout "${s.layout}" is not one of title, section, content, two_column, blank, title_only`)
    return k
  }
  if (s.left || s.right) return 'two_column'
  if (s.bullets || s.table || s.image || s.text) return 'content'
  if (s.subtitle !== undefined) return 'title'
  if (s.title) return 'title_only'
  return 'blank'
}

function themeAccent (deck) {
  const theme = deck.pkg.relsOfType(deck.main, 'theme')[0]?.part
  const a1 = theme && descendants(deck.pkg.xml(theme).root, 'a:accent1')[0]
  return child(a1, 'a:srgbClr')?.attrs.val ?? '1F4E79'
}

// buildSlide writes a new slide part from a spec and returns its name.
async function buildSlide (deck, spec, io) {
  const kind = inferKind(spec)
  const layout = deck.pickLayout(kind)
  const pkg = deck.pkg
  const part = pkg.uniqueName('ppt/slides/slide#.xml')
  pkg.setXml(part, doc(el('p:sld', NSA, [el('p:cSld', {}, [el('p:spTree', {}, grpRoot())]), el('p:clrMapOvr', {}, [el('a:masterClrMapping')])])))
  pkg.setOverride(part, CT.slide)
  pkg.addRel(part, REL('slideLayout'), layout.part)
  const tree = xpath(pkg.xml(part).root, 'p:cSld', 'p:spTree')
  const phs = placeholders(pkg.xml(layout.part).root).filter((p) => !['dt', 'ftr', 'sldNum', 'hdr'].includes(p.type))
  let id = 2
  const used = new Set()
  const take = (test) => { const p = phs.find((x) => !used.has(x) && test(x)); if (p) used.add(p); return p }
  const placeSp = (p, paras) => {
    const ph = {}
    for (const k of ['type', 'orient', 'sz', 'idx']) if (p.ph.attrs[k] !== undefined) ph[k] = p.ph.attrs[k]
    const name = (child(xpath(p.sp, 'p:nvSpPr'), 'p:cNvPr')?.attrs.name ?? 'Placeholder').replace(/\s*\d+$/, '') + ' ' + id
    tree.kids.push(el('p:sp', {}, [
      el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id++), name }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1' })]), el('p:nvPr', {}, [el('p:ph', ph)])]),
      el('p:spPr'),
      textBody(paras)
    ]))
  }
  const titlePh = take((p) => isTitle(p.type))
  if (spec.title !== undefined && spec.title !== null) {
    if (titlePh) placeSp(titlePh, paragraphs([spec.title]))
    else tree.kids.push(textBox(id++, { x: deck.cx * 0.06, y: deck.cy * 0.05, cx: deck.cx * 0.88, cy: deck.cy * 0.16 }, paragraphs([spec.title], { sz: '3600', b: '1' })))
  }
  if (spec.subtitle !== undefined && spec.subtitle !== null) {
    const p = take((x) => x.type === 'subTitle') ?? take((x) => isContent(x.type))
    if (p) placeSp(p, paragraphs(String(spec.subtitle).split('\n')))
    else tree.kids.push(textBox(id++, { x: deck.cx * 0.06, y: deck.cy * 0.6, cx: deck.cx * 0.88, cy: deck.cy * 0.2 }, paragraphs(String(spec.subtitle).split('\n'), { sz: '2000' })))
  }
  const contentArea = (p) => {
    const g = p && geometry(deck, p.sp, layout.part)
    return g ?? { x: Math.round(deck.cx * 0.07), y: Math.round(deck.cy * 0.27), cx: Math.round(deck.cx * 0.86), cy: Math.round(deck.cy * 0.63) }
  }
  const accent = themeAccent(deck)
  const fill = async (items, p) => {
    if (!items) return
    if (Array.isArray(items)) {
      if (p) placeSp(p, paragraphs(items))
      else tree.kids.push(textBox(id++, contentArea(null), paragraphs(items, { sz: '2000' })))
      return
    }
    if (typeof items === 'string') return fill(items.split('\n').map((l) => { const m = /^(\s*)[-*•]?\s*(.*)$/.exec(l); return { text: m[2], level: Math.floor(m[1].length / 2) } }), p)
    if (items.table) { tree.kids.push(tableFrame(id++, contentArea(p), items.table, accent)); return }
    if (items.image) { tree.kids.push(await pictureShape(deck, part, id++, contentArea(p), items.image.path ?? items.image, io)); return }
    if (items.bullets) return fill(items.bullets, p)
    if (items.text) return fill(items.text, p)
  }
  if (kind === 'two_column') {
    const l = take((x) => isContent(x.type))
    const r = take((x) => isContent(x.type))
    if (l && r) {
      await fill(spec.left, l)
      await fill(spec.right, r)
    } else {
      // One content area or none: split it into two text boxes.
      const g = contentArea(l)
      const half = { ...g, cx: Math.round(g.cx * 0.48) }
      if (spec.left) tree.kids.push(textBox(id++, half, paragraphs(asItems(spec.left), { sz: '2000' })))
      if (spec.right) tree.kids.push(textBox(id++, { ...half, x: g.x + g.cx - half.cx }, paragraphs(asItems(spec.right), { sz: '2000' })))
    }
  } else {
    const body = take((x) => isContent(x.type))
    const items = spec.bullets ?? spec.text
    if (items) await fill(items, body)
    if (spec.table) await fill({ table: spec.table }, items ? null : body)
    if (spec.image) await fill({ image: spec.image }, items || spec.table ? null : body)
  }
  if (spec.notes) setNotes(deck, part, String(spec.notes))
  return part
}

const asItems = (x) => Array.isArray(x) ? x : typeof x === 'string' ? x.split('\n') : x?.bullets ?? []

function textBox (id, box, paras) {
  return el('p:sp', {}, [
    el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id), name: `TextBox ${id}` }), el('p:cNvSpPr', { txBox: '1' }), el('p:nvPr')]),
    el('p:spPr', {}, [xfrm(box.x, box.y, box.cx, box.cy), el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]), el('a:noFill')]),
    el('p:txBody', {}, [el('a:bodyPr', { wrap: 'square', rtlCol: '0' }, [el('a:normAutofit')]), el('a:lstStyle'), ...paras])
  ])
}

function insertSlide (deck, part, at) {
  const rid = deck.pkg.addRel(deck.main, REL('slide'), part)
  const ids = [...descendants(deck.pres.root, 'p:sldId').map((n) => +n.attrs.id), 255]
  const node = el('p:sldId', { id: String(Math.max(...ids) + 1), 'r:id': rid })
  const list = children(deck.idList, 'p:sldId')
  const pos = at === undefined || at === null ? list.length : Math.max(0, Math.min(list.length, +at - 1))
  if (pos >= list.length) deck.idList.kids.push(node)
  else deck.idList.kids.splice(deck.idList.kids.indexOf(list[pos]), 0, node)
  deck.pkg.touch(deck.main)
}

function ensureNotesMaster (deck) {
  const pkg = deck.pkg
  const existing = pkg.relsOfType(deck.main, 'notesMaster')[0]?.part
  if (existing && pkg.has(existing)) return existing
  const part = pkg.uniqueName('ppt/notesMasters/notesMaster#.xml')
  pkg.setXml(part, notesMasterDoc())
  pkg.setOverride(part, CT.notesMaster)
  const theme = pkg.uniqueName('ppt/theme/theme#.xml')
  pkg.setXml(theme, themeDoc({ accent: '1F4E79', text: '1F2328', font: { latin: 'Calibri', ea: 'Microsoft YaHei' } }))
  pkg.setOverride(theme, CT.theme)
  pkg.addRel(part, REL('theme'), theme)
  const rid = pkg.addRel(deck.main, REL('notesMaster'), part)
  let lst = child(deck.pres.root, 'p:notesMasterIdLst')
  if (!lst) {
    lst = el('p:notesMasterIdLst')
    const at = deck.pres.root.kids.findIndex((k) => k.name === 'p:handoutMasterIdLst' || k.name === 'p:sldIdLst' || k.name === 'p:sldSz')
    deck.pres.root.kids.splice(at < 0 ? deck.pres.root.kids.length : at, 0, lst)
  }
  lst.kids = [el('p:notesMasterId', { 'r:id': rid })]
  pkg.touch(deck.main)
  return part
}

function setNotes (deck, slidePart, str) {
  const pkg = deck.pkg
  let notes = deck.notesOf(slidePart)
  if (!notes || !pkg.has(notes)) {
    const master = ensureNotesMaster(deck)
    notes = pkg.uniqueName('ppt/notesSlides/notesSlide#.xml')
    pkg.setXml(notes, doc(el('p:notes', NSA, [
      el('p:cSld', {}, [el('p:spTree', {}, [...grpRoot(),
        el('p:sp', {}, [el('p:nvSpPr', {}, [el('p:cNvPr', { id: '2', name: 'Slide Image Placeholder 1' }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1', noRot: '1', noChangeAspect: '1' })]), el('p:nvPr', {}, [el('p:ph', { type: 'sldImg' })])]), el('p:spPr')]),
        el('p:sp', {}, [el('p:nvSpPr', {}, [el('p:cNvPr', { id: '3', name: 'Notes Placeholder 2' }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1' })]), el('p:nvPr', {}, [el('p:ph', { type: 'body', idx: '1' })])]), el('p:spPr'), textBody([])])
      ])]),
      el('p:clrMapOvr', {}, [el('a:masterClrMapping')])
    ])))
    pkg.setOverride(notes, CT.notes)
    pkg.addRel(notes, REL('notesMaster'), master)
    pkg.addRel(notes, REL('slide'), slidePart)
    pkg.addRel(slidePart, REL('notesSlide'), notes)
  }
  const body = placeholders(pkg.xml(notes).root).find((p) => p.type === 'body')
  if (!body) throw new Error('the notes page has no notes placeholder')
  let tb = child(body.sp, 'p:txBody')
  if (!tb) { tb = textBody([]); body.sp.kids.push(tb) }
  setParagraphs(tb, str)
  pkg.touch(notes)
}

// setParagraphs replaces a text body's paragraphs, keeping the formatting
// of the first paragraph at each level.
function setParagraphs (tb, str) {
  const old = children(tb, 'a:p')
  const proto = new Map()
  for (const p of old) {
    const l = +(child(p, 'a:pPr')?.attrs.lvl ?? 0)
    if (!proto.has(l)) proto.set(l, { pPr: child(p, 'a:pPr'), rPr: descendants(p, 'a:rPr')[0] ?? child(p, 'a:endParaRPr') })
  }
  const first = proto.get(0) ?? [...proto.values()][0] ?? {}
  const lines = String(str).split('\n')
  const paras = lines.map((line) => {
    const m = /^(\s*)(?:[-•*]\s+)?(.*)$/.exec(line)
    const lv = Math.min(8, Math.floor(m[1].replace(/\t/g, '  ').length / 2))
    const pr = proto.get(lv) ?? first
    let pPr = pr.pPr ? clone(pr.pPr) : (lv ? el('a:pPr') : null)
    if (pPr) { if (lv) pPr.attrs.lvl = String(lv); else delete pPr.attrs.lvl }
    if (pPr && !Object.keys(pPr.attrs).length && !pPr.kids.length) pPr = null
    const rattrs = pr.rPr ? { ...pr.rPr.attrs } : { lang: 'zh-CN', altLang: 'en-US' }
    const rkids = pr.rPr ? pr.rPr.kids.map(clone) : []
    delete rattrs.dirty
    const runs = m[2] ? inline(m[2]).map((r) => {
      if (r.br) return el('a:br')
      const a = { ...rattrs }
      if (r.b) a.b = '1'
      if (r.i) a.i = '1'
      return el('a:r', {}, [el('a:rPr', a, rkids.map(clone)), el('a:t', {}, [text(r.text ?? '')])])
    }) : [el('a:endParaRPr', rattrs)]
    return el('a:p', {}, [pPr, ...runs])
  })
  const at = tb.kids.findIndex((k) => k.name === 'a:p')
  tb.kids = tb.kids.filter((k) => k.name !== 'a:p')
  tb.kids.splice(at < 0 ? tb.kids.length : at, 0, ...paras)
}

// clonePart copies a part and its relationships for a duplicated slide;
// layouts, masters, media and links stay shared, notes are dropped.
function clonePart (pkg, from, to) {
  pkg.setBytes(to, pkg.get(from))
  const ct = pkg.contentType(from)
  if (ct && !pkg.contentType(to)) pkg.setOverride(to, ct)
  else if (ct) pkg.setOverride(to, ct)
  for (const r of pkg.rels(from)) {
    if (r.external) { pkg.addRelWithId(to, r.id, r.type, r.target, true); continue }
    if (/\/notesSlide$/.test(r.type)) continue
    if (/\/(slideLayout|slideMaster|image|media|video|audio|hyperlink|notesMaster|theme)$/.test(r.type) || /\/(media|image)/.test(r.type) || !r.part || !pkg.has(r.part)) {
      pkg.addRelWithId(to, r.id, r.type, r.target)
      continue
    }
    const ext = path.extname(r.part)
    const base = r.part.slice(0, -ext.length).replace(/\d+$/, '')
    const copy = pkg.uniqueName(base + '#' + ext)
    clonePart(pkg, r.part, copy)
    pkg.addRelWithId(to, r.id, r.type, path.relative(path.dirname(to), copy))
  }
}

Package.prototype.addRelWithId = function (part, id, type, target, external = false) {
  const { name, d } = this.relsDoc(part, true)
  const attrs = { Id: id, Type: type, Target: target }
  if (external) attrs.TargetMode = 'External'
  d.root.kids.push(el('Relationship', attrs))
  this.touch(name)
}

function deleteSlide (deck, s) {
  const pkg = deck.pkg
  remove(deck.idList, s.node)
  pkg.removeRel(deck.main, s.rid)
  // Sections list slide ids; a dangling id makes PowerPoint repair the file.
  for (const n of descendants(deck.pres.root, (x) => /:sldId$/.test(x.name) && x.name !== 'p:sldId' && x.attrs.id === s.id)) {
    const parent = descendants(deck.pres.root, (x) => x.kids?.includes(n))[0]
    if (parent) remove(parent, n)
  }
  const notes = deck.notesOf(s.part)
  for (const part of [notes, s.part]) {
    if (!part || !pkg.has(part)) continue
    pkg.delete(part)
    pkg.delete(pkg.relsName(part))
    pkg.removeOverride(part)
  }
  pkg.touch(deck.main)
}

// ------------------------------------------------------------- create & edit

export async function createPptx (slides, theme = {}, io) {
  if (!Array.isArray(slides) || !slides.length) throw new Error('slides must list at least one slide')
  const t = {
    accent: theme.accent ? hex(theme.accent) : '1F4E79',
    background: theme.background ? hex(theme.background) : 'FFFFFF',
    text: theme.text ? hex(theme.text) : '1F2328',
    font: { latin: theme.font || 'Calibri', ea: theme.font || 'Microsoft YaHei' }
  }
  const pkg = newPackage(t)
  const deck = new Deck(pkg)
  for (const s of slides) insertSlide(deck, await buildSlide(deck, s, io))
  return { buffer: pkg.toBuffer(), summary: `${slides.length} slide(s)` }
}

function hex (c) {
  const s = String(c).replace(/^#/, '').toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(s)) throw new Error(`"${c}" is not a 6-digit hex colour like 1F4E79`)
  return s
}

export async function editPptx (buf, edits, io) {
  const pkg = Package.fromBuffer(buf)
  const deck = new Deck(pkg)
  // Refs and slide numbers name the deck as it was before this call.
  const before = deck.slides()
  const byNumber = (n, where) => {
    const s = before[(+n) - 1]
    if (!s) throw new Error(`${where}: there is no slide ${n}; the deck has ${before.length} slides (numbering starts at 1)`)
    if (!children(deck.idList, 'p:sldId').includes(s.node)) throw new Error(`${where}: slide ${n} was deleted by an earlier edit in this call`)
    return s
  }
  const report = []
  for (const [k, e] of edits.entries()) {
    const where = `edits[${k}] (${e?.op})`
    switch (e?.op) {
      case 'replace': {
        if (!e.find) throw new Error(`${where}: find is empty`)
        const pick = parseSlides(e.slides, before.length)
        const parts = deck.slides().filter((s) => s.part && pkg.has(s.part) && (!pick || pick.has(before.findIndex((b) => b.node === s.node) + 1))).map((s) => s.part)
        const n = replaceInParts(pkg, parts, e.find, e.replace ?? '', !!e.all, where, (root) => descendants(root, 'a:p').map((p) => descendants(p, 'a:t')))
        report.push(`replaced ${n}×`)
        break
      }
      case 'set_text': {
        const m = /^\[?s(\d+)\.(\d+)\]?$/.exec(String(e.shape ?? ''))
        if (!m) throw new Error(`${where}: shape should be a ref like s3.2 from office_read`)
        const s = byNumber(m[1], where)
        const list = shapes(xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree'))
        const sh = list[+m[2] - 1]
        if (!sh) throw new Error(`${where}: slide ${m[1]} has ${list.length} shapes; call office_read for the current refs`)
        const tb = child(sh, 'p:txBody')
        if (!tb) throw new Error(`${where}: s${m[1]}.${m[2]} is a ${sh.name === 'p:pic' ? 'picture' : 'table or object'} without its own text; use replace for text in tables`)
        setParagraphs(tb, e.text ?? '')
        pkg.touch(s.part)
        report.push(`set text of s${m[1]}.${m[2]}`)
        break
      }
      case 'add_slide': {
        const part = await buildSlide(deck, e.slide ?? {}, io)
        insertSlide(deck, part, e.at)
        report.push(`added a slide${e.at ? ` at ${e.at}` : ''}`)
        break
      }
      case 'duplicate_slide': {
        const s = byNumber(e.slide, where)
        const part = pkg.uniqueName('ppt/slides/slide#.xml')
        clonePart(pkg, s.part, part)
        const pos = e.at ?? (deck.slides().findIndex((x) => x.node === s.node) + 2)
        insertSlide(deck, part, pos)
        report.push(`duplicated slide ${e.slide}`)
        break
      }
      case 'delete_slide': {
        const s = byNumber(e.slide, where)
        if (deck.slides().length === 1) throw new Error(`${where}: a presentation must keep at least one slide`)
        deleteSlide(deck, s)
        report.push(`deleted slide ${e.slide}`)
        break
      }
      case 'move_slide': {
        const s = byNumber(e.slide, where)
        const list = children(deck.idList, 'p:sldId')
        const to = Math.max(1, Math.min(list.length, +e.to))
        remove(deck.idList, s.node)
        const rest = children(deck.idList, 'p:sldId')
        if (to - 1 >= rest.length) deck.idList.kids.push(s.node)
        else deck.idList.kids.splice(deck.idList.kids.indexOf(rest[to - 1]), 0, s.node)
        pkg.touch(deck.main)
        report.push(`moved slide ${e.slide} to ${to}`)
        break
      }
      case 'notes': {
        const s = byNumber(e.slide, where)
        setNotes(deck, s.part, String(e.text ?? ''))
        report.push(`set notes of slide ${e.slide}`)
        break
      }
      default:
        throw new Error(`${where}: unknown op; use replace, set_text, add_slide, duplicate_slide, delete_slide, move_slide or notes`)
    }
  }
  return { buffer: pkg.toBuffer(), summary: report.join('; ') }
}

// ------------------------------------------------------------- rendering

const escH = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

// Colours: theme scheme, the master's colour map, and the usual modifiers.
class Colors {
  constructor (deck, master) {
    this.map = { ...CLRMAP }
    const cm = master && child(deck.pkg.xml(master).root, 'p:clrMap')
    if (cm) Object.assign(this.map, cm.attrs)
    this.scheme = {}
    const theme = master && deck.pkg.relsOfType(master, 'theme')[0]?.part
    const cs = theme && deck.pkg.has(theme) && descendants(deck.pkg.xml(theme).root, 'a:clrScheme')[0]
    if (cs) for (const k of cs.kids) if (isEl(k)) this.scheme[k.name.slice(2)] = child(k, 'a:srgbClr')?.attrs.val ?? child(k, 'a:sysClr')?.attrs.lastClr
    const fonts = theme && deck.pkg.has(theme) && descendants(deck.pkg.xml(theme).root, 'a:fontScheme')[0]
    this.major = xpath(fonts, 'a:majorFont', 'a:latin')?.attrs.typeface
    this.minor = xpath(fonts, 'a:minorFont', 'a:latin')?.attrs.typeface
    this.majorEa = xpath(fonts, 'a:majorFont', 'a:ea')?.attrs.typeface
    this.minorEa = xpath(fonts, 'a:minorFont', 'a:ea')?.attrs.typeface
  }

  // of resolves a colour element's child (srgbClr, schemeClr, …).
  of (fillEl, ph) {
    const c = fillEl?.kids?.find((k) => isEl(k))
    if (!c) return undefined
    let base
    switch (c.name) {
      case 'a:srgbClr': base = c.attrs.val; break
      case 'a:sysClr': base = c.attrs.lastClr ?? (c.attrs.val === 'window' ? 'FFFFFF' : '000000'); break
      case 'a:prstClr': base = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '00FF00', gray: '808080' }[c.attrs.val] ?? '000000'; break
      case 'a:schemeClr': {
        if (c.attrs.val === 'phClr') { base = ph; break }
        const key = this.map[c.attrs.val] ?? c.attrs.val
        base = this.scheme[key]
        break
      }
      case 'a:scrgbClr': base = [c.attrs.r, c.attrs.g, c.attrs.b].map((x) => Math.round(+x / 100000 * 255).toString(16).padStart(2, '0')).join(''); break
    }
    if (!base) return undefined
    return modify(base, c)
  }
}

function modify (hexv, c) {
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hexv.slice(i, i + 2), 16) / 255)
  const mods = Object.fromEntries(c.kids.filter(isEl).map((k) => [k.name.slice(2), +k.attrs.val / 100000]))
  if (mods.lumMod !== undefined || mods.lumOff !== undefined) {
    let [h, s, l] = rgbToHsl(r, g, b)
    l = Math.min(1, Math.max(0, l * (mods.lumMod ?? 1) + (mods.lumOff ?? 0)))
    ;[r, g, b] = hslToRgb(h, s, l)
  }
  if (mods.tint !== undefined) [r, g, b] = [r, g, b].map((x) => x + (1 - x) * (1 - mods.tint))
  if (mods.shade !== undefined) [r, g, b] = [r, g, b].map((x) => x * mods.shade)
  const out = [r, g, b].map((x) => Math.round(Math.min(1, Math.max(0, x)) * 255).toString(16).padStart(2, '0')).join('')
  return mods.alpha !== undefined && mods.alpha < 1 ? out + Math.round(mods.alpha * 255).toString(16).padStart(2, '0') : out
}

function rgbToHsl (r, g, b) {
  const max = Math.max(r, g, b); const min = Math.min(r, g, b)
  let h = 0; let s = 0; const l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  return [h, s, l]
}

function hslToRgb (h, s, l) {
  if (!s) return [l, l, l]
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)]
}

const css = (hexv) => hexv && (hexv.length === 8 ? `#${hexv}` : `#${hexv}`)

export function renderPptx (buf) {
  const pkg = Package.fromBuffer(buf)
  const deck = new Deck(pkg)
  const W = 960
  const k = W / deck.cx
  const H = Math.round(deck.cy * k)
  const ptPx = W / (deck.cx / 12700)
  const out = []
  for (const s of deck.slides()) {
    if (!s.part || !pkg.has(s.part)) continue
    const layout = deck.layoutOf(s.part)
    const master = layout && deck.masterOf(layout)
    const colors = new Colors(deck, master)
    const sd = pkg.xml(s.part)
    const ld = layout && pkg.has(layout) ? pkg.xml(layout) : undefined
    const md = master && pkg.has(master) ? pkg.xml(master) : undefined
    const ctx = { deck, pkg, colors, k, ptPx, layout, master, ld, md }
    const bg = background(ctx, [[sd, s.part], [ld, layout], [md, master]])
    let html = ''
    const showMaster = (d) => d && d.root.attrs.showMasterSp !== '0'
    if (md && showMaster(sd) && showMaster(ld)) html += renderTree(ctx, xpath(md.root, 'p:cSld', 'p:spTree'), master, { decorOnly: true })
    if (ld && showMaster(sd)) html += renderTree(ctx, xpath(ld.root, 'p:cSld', 'p:spTree'), layout, { decorOnly: true })
    html += renderTree(ctx, xpath(sd.root, 'p:cSld', 'p:spTree'), s.part, {})
    let notes = ''
    const np = deck.notesOf(s.part)
    if (np && pkg.has(np)) {
      const body = placeholders(pkg.xml(np).root).find((p) => p.type === 'body')
      const t = body ? children(child(body.sp, 'p:txBody'), 'a:p').map(paraText).join('\n').trim() : ''
      if (t) notes = `<div class="notes"><b>备注</b> ${escH(t).replace(/\n/g, '<br>')}</div>`
    }
    out.push(`<section class="slide-wrap"><div class="num">${s.n}${sd.root.attrs.show === '0' ? ' · 已隐藏' : ''}</div><div class="slide" style="width:${W}px;height:${H}px;${bg}">${html}</div>${notes}</section>`)
  }
  return {
    kind: 'PowerPoint 演示文稿',
    css: PPTX_CSS,
    body: `<div class="deck">${out.join('\n')}</div>`
  }
}

function background (ctx, chain) {
  for (const [d, part] of chain) {
    const bg = d && xpath(d.root, 'p:cSld', 'p:bg')
    if (!bg) continue
    const pr = child(bg, 'p:bgPr')
    if (pr) {
      const sf = child(pr, 'a:solidFill')
      if (sf) return `background:${css(ctx.colors.of(sf))};`
      const gf = child(pr, 'a:gradFill')
      if (gf) { const g = gradCss(ctx, gf); if (g) return `background:${g};` }
      const bf = child(pr, 'a:blipFill')
      const rid = bf && child(bf, 'a:blip')?.attrs['r:embed']
      const img = rid && ctx.pkg.rel(part, rid)?.part
      const uri = img && dataUri(ctx.pkg, img)
      if (uri) return `background:url(${uri}) center/cover;`
    }
    const ref = child(bg, 'p:bgRef')
    if (ref) { const c = ctx.colors.of(ref); if (c) return `background:${css(c)};` }
  }
  return 'background:#fff;'
}

// renderTree draws shapes. decorOnly skips placeholders (master and layout
// placeholders only lend position and style to the slide's own).
function renderTree (ctx, tree, part, { decorOnly, group, outer }) {
  if (!tree) return ''
  let html = ''
  const map = group ?? ((g) => g)
  for (const node of tree.kids) {
    if (!isEl(node)) continue
    if (node.name === 'mc:AlternateContent') {
      const pick = child(node, 'mc:Choice') ?? child(node, 'mc:Fallback')
      if (pick) html += renderTree(ctx, pick, part, { decorOnly, group, outer })
      continue
    }
    if (node.name === 'p:grpSp') {
      const x = xpath(node, 'p:grpSpPr', 'a:xfrm')
      const off = child(x, 'a:off'); const ext = child(x, 'a:ext'); const choff = child(x, 'a:chOff'); const chext = child(x, 'a:chExt')
      const g = off && ext && choff && chext && +chext.attrs.cx && +chext.attrs.cy
        ? (b) => map({ ...b, x: +off.attrs.x + (b.x - +choff.attrs.x) * (+ext.attrs.cx / +chext.attrs.cx), y: +off.attrs.y + (b.y - +choff.attrs.y) * (+ext.attrs.cy / +chext.attrs.cy), cx: b.cx * (+ext.attrs.cx / +chext.attrs.cx), cy: b.cy * (+ext.attrs.cy / +chext.attrs.cy) })
        : map
      // The editor moves a top-level group as one; its shapes carry its number.
      const o = outer ?? (ctx.ui?.groups?.has(node) && !decorOnly ? { i: ctx.ui.groups.get(node) + 1, box: map(xfrmOf(node) ?? { x: 0, y: 0, cx: 0, cy: 0 }), rot: xfrmOf(node)?.rot ?? 0, id: xpath(node, 'p:nvGrpSpPr', 'p:cNvPr')?.attrs.id } : undefined)
      html += renderTree(ctx, node, part, { decorOnly, group: g, outer: o })
      continue
    }
    if (!['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp'].includes(node.name)) continue
    const ph = phOf(node)
    if (decorOnly && ph) continue
    let box = xfrmOf(node) ?? (ph && ctx.layout ? geometry(ctx.deck, node, ctx.layout) : undefined)
    if (!box) continue
    box = map(box)
    // The editor addresses a slide's own shapes by their place in shapes().
    const k = !decorOnly && ctx.ui ? ctx.ui.index.get(node) : undefined
    const kind = node.name === 'p:pic' ? 'pic' : node.name === 'p:graphicFrame' ? (descendants(node, 'a:tbl')[0] ? 'tbl' : 'obj') : child(node, 'p:txBody') ? 'txt' : 'shape'
    const cid = xpath(nv(node), 'p:cNvPr')?.attrs.id
    const ua = k !== undefined ? ` data-k="${k + 1}" data-t="${kind}"${cid ? ` data-id="${cid}"` : ''}${ph ? ' data-ph="1"' : ''}${outer?.id ? ` data-gid="${outer.id}"` : ''}${outer ? ` data-g="${outer.i}" data-gbox="${[outer.box.x, outer.box.y, outer.box.cx, outer.box.cy].map(Math.round).join(',')},${+(outer.rot || 0).toFixed(2)}"` : group ? ' data-g="0"' : ''}${box.rot ? ` data-rot="${+box.rot.toFixed(2)}"` : ''} data-box="${[box.x, box.y, box.cx, box.cy].map(Math.round).join(',')}"` : ''
    const style = `left:${box.x * ctx.k}px;top:${box.y * ctx.k}px;width:${box.cx * ctx.k}px;height:${box.cy * ctx.k}px;${box.rot ? `transform:rotate(${box.rot}deg)${box.flipH ? ' scaleX(-1)' : ''};` : box.flipH ? 'transform:scaleX(-1);' : ''}`
    if (node.name === 'p:pic') {
      const rid = descendants(node, 'a:blip')[0]?.attrs['r:embed']
      const img = rid && ctx.pkg.rel(part, rid)?.part
      const uri = img && dataUri(ctx.pkg, img)
      const sr = xpath(node, 'p:blipFill', 'a:srcRect')
      const [l, t, r, b] = ['l', 't', 'r', 'b'].map((x) => +(sr?.attrs[x] ?? 0) / 100000)
      const w = Math.max(0.01, 1 - l - r); const h = Math.max(0.01, 1 - t - b)
      html += uri ? `<div class="shp pic"${ua}${ua ? ` data-crop="${[l, t, r, b].join(',')}"` : ''} style="${style}${picFx(ctx, node)}"><img src="${uri}" alt="" style="left:${(-l / w * 100).toFixed(3)}%;top:${(-t / h * 100).toFixed(3)}%;width:${(100 / w).toFixed(3)}%;height:${(100 / h).toFixed(3)}%"></div>` : `<div class="shp ph-missing"${ua} style="${style}">图片</div>`
      continue
    }
    if (node.name === 'p:graphicFrame') {
      const tbl = descendants(node, 'a:tbl')[0]
      if (tbl) { html += `<div class="shp"${ua} style="${style}">${renderTable(ctx, tbl, box)}</div>`; continue }
      const uri = descendants(node, 'a:graphicData')[0]?.attrs.uri ?? ''
      html += `<div class="shp ph-missing"${ua} style="${style}">${/chart/.test(uri) ? '图表' : /diagram/.test(uri) ? 'SmartArt' : '对象'}</div>`
      continue
    }
    html += renderShape(ctx, node, part, style, ph, ua)
  }
  return html
}

const CLIPS = {
  triangle: '50% 0,100% 100%,0 100%',
  diamond: '50% 0,100% 50%,50% 100%,0 50%',
  rightArrow: '0 25%,62% 25%,62% 0,100% 50%,62% 100%,62% 75%,0 75%',
  leftArrow: '100% 25%,38% 25%,38% 0,0 50%,38% 100%,38% 75%,100% 75%',
  pentagon: '0 0,80% 0,100% 50%,80% 100%,0 100%',
  star5: '50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%'
}

function shapeFill (ctx, sp) {
  const spPr = child(sp, 'p:spPr')
  const parts = []
  if (spPr) {
    if (child(spPr, 'a:noFill')) { /* transparent */ } else if (child(spPr, 'a:solidFill')) parts.push(`background:${css(ctx.colors.of(child(spPr, 'a:solidFill')))}`)
    else if (child(spPr, 'a:gradFill')) {
      const g = gradCss(ctx, child(spPr, 'a:gradFill'))
      if (g) parts.push(`background:${g}`)
    } else {
      const ref = xpath(sp, 'p:style', 'a:fillRef')
      if (ref && +ref.attrs.idx > 0) { const c = ctx.colors.of(ref); if (c) parts.push(`background:${css(c)}`) }
    }
    const ln = child(spPr, 'a:ln')
    if (ln && !child(ln, 'a:noFill')) {
      const c = child(ln, 'a:solidFill') ? ctx.colors.of(child(ln, 'a:solidFill')) : undefined
      const w = Math.max(1, Math.round(+(ln.attrs.w ?? 12700) * ctx.k))
      if (c) parts.push(`border:${w}px solid ${css(c)}`)
    } else if (!ln) {
      const ref = xpath(sp, 'p:style', 'a:lnRef')
      if (ref && +ref.attrs.idx > 0) { const c = ctx.colors.of(ref); if (c) parts.push(`border:1px solid ${css(c)}`) }
    }
    const geom = child(spPr, 'a:prstGeom')?.attrs.prst
    if (geom === 'ellipse') parts.push('border-radius:50%')
    if (geom === 'roundRect') parts.push('border-radius:8px')
    if (CLIPS[geom]) parts.push(`clip-path:polygon(${CLIPS[geom]})`)
    const sh = xpath(spPr, 'a:effectLst', 'a:outerShdw')
    if (sh && !CLIPS[geom]) parts.push(`box-shadow:${shadowCss(ctx, sh)}`)
  }
  return parts.join(';')
}

function picFx (ctx, pic) {
  const spPr = child(pic, 'p:spPr')
  let s = ''
  const ln = child(spPr, 'a:ln')
  const c = ln && !child(ln, 'a:noFill') && ctx.colors.of(child(ln, 'a:solidFill'))
  if (c) { const w = Math.max(1, Math.round(+(ln.attrs.w ?? 12700) * ctx.k)); s += `outline:${w}px solid ${css(c)};outline-offset:-${w}px;` }
  const sh = xpath(spPr, 'a:effectLst', 'a:outerShdw')
  if (sh) s += `box-shadow:${shadowCss(ctx, sh)};`
  const geom = child(spPr, 'a:prstGeom')?.attrs.prst
  if (geom === 'ellipse') s += 'border-radius:50%;'
  if (geom === 'roundRect') s += 'border-radius:8px;'
  return s
}

// styleChain lists where a placeholder's text formatting comes from, most
// specific first: its own list style, the layout's and master's
// placeholders, then the master's text styles.
function styleChain (ctx, sp, ph) {
  const chain = [child(child(sp, 'p:txBody'), 'a:lstStyle')]
  const type = ph?.attrs.type
  if (ph) {
    const find = (d) => {
      if (!d) return undefined
      const phs = placeholders(d.root)
      const idx = ph.attrs.idx ?? '0'
      return phs.find((p) => idx !== '0' && p.idx === idx) ?? phs.find((p) => (p.type ?? 'obj') === (type ?? 'obj')) ?? phs.find((p) => (isTitle(type) && isTitle(p.type)) || (!isTitle(type) && isContent(p.type)))
    }
    const lp = find(ctx.ld)
    if (lp) chain.push(child(child(lp.sp, 'p:txBody'), 'a:lstStyle'))
    const mp = ctx.md && placeholders(ctx.md.root).find((p) => isTitle(type) ? p.type === 'title' : p.type === 'body')
    if (mp) chain.push(child(child(mp.sp, 'p:txBody'), 'a:lstStyle'))
    const ts = ctx.md && child(ctx.md.root, 'p:txStyles')
    chain.push(child(ts, isTitle(type) ? 'p:titleStyle' : ['body', 'obj', 'subTitle', undefined].includes(type) ? 'p:bodyStyle' : 'p:otherStyle'))
    // Body placeholders: anchor and insets from layout/master too.
    chain.bodyPr = [child(child(sp, 'p:txBody'), 'a:bodyPr'), lp && child(child(lp.sp, 'p:txBody'), 'a:bodyPr'), mp && child(child(mp.sp, 'p:txBody'), 'a:bodyPr')].filter(Boolean)
  } else {
    const other = ctx.md && xpath(ctx.md.root, 'p:txStyles', 'p:otherStyle')
    chain.push(other)
    chain.bodyPr = [child(child(sp, 'p:txBody'), 'a:bodyPr')]
  }
  const out = chain.filter((x) => x !== undefined)
  out.bodyPr = chain.bodyPr
  return out
}

function lvlProp (chain, lv, pick) {
  for (const src of chain) {
    if (!src) continue
    const l = child(src, `a:lvl${lv + 1}pPr`)
    if (l) { const v = pick(l); if (v !== undefined) return v }
  }
  return undefined
}

function renderShape (ctx, sp, part, style, ph, ua = '') {
  const fill = shapeFill(ctx, sp)
  const tb = child(sp, 'p:txBody')
  const paras = children(tb, 'a:p')
  const empty = !paras.some((p) => paraText(p).trim())
  if (empty && !fill && !(ua && (tb || ph))) return ''
  const chain = styleChain(ctx, sp, ph)
  const bodyPr = chain.bodyPr ?? []
  const bp = (name) => { for (const b of bodyPr) if (b.attrs[name] !== undefined) return b.attrs[name]; return undefined }
  const anchor = bp('anchor') ?? (isTitle(ph?.attrs.type) ? 'ctr' : 't')
  const scale = (() => { for (const b of bodyPr) { const a = child(b, 'a:normAutofit'); if (a) return +(a.attrs.fontScale ?? 100000) / 100000 } return 1 })()
  const lnRed = (() => { for (const b of bodyPr) { const a = child(b, 'a:normAutofit'); if (a) return 1 - +(a.attrs.lnSpcReduction ?? 0) / 100000 } return 1 })()
  const fit = (() => { for (const b of bodyPr) { if (child(b, 'a:spAutoFit')) return 'shape'; if (child(b, 'a:normAutofit')) return 'shrink'; if (child(b, 'a:noAutofit')) return 'none' } return '' })()
  const ins = (n, d) => (+(bp(n) ?? d)) * ctx.k
  const vert = bp('vert')
  const isTitleType = isTitle(ph?.attrs.type) || ph?.attrs.type === 'subTitle'
  let inner = ''
  const counters = new Map()
  for (const p of paras) {
    const ppr = child(p, 'a:pPr')
    const lv = +(ppr?.attrs.lvl ?? 0)
    const prop = (pick) => (ppr && pick(ppr)) ?? lvlProp(chain, lv, pick)
    const algn = prop((l) => l.attrs.algn) ?? 'l'
    const marL = +(prop((l) => l.attrs.marL) ?? 0) * ctx.k
    const indent = +(prop((l) => l.attrs.indent) ?? 0) * ctx.k
    const defSz = prop((l) => child(l, 'a:defRPr')?.attrs.sz)
    const defB = prop((l) => child(l, 'a:defRPr')?.attrs.b)
    const defColor = prop((l) => { const f = child(child(l, 'a:defRPr'), 'a:solidFill'); return f ? ctx.colors.of(f) : undefined })
    const buNone = prop((l) => (child(l, 'a:buNone') ? true : (child(l, 'a:buChar') || child(l, 'a:buAutoNum')) ? false : undefined))
    const buChar = prop((l) => child(l, 'a:buChar')?.attrs.char)
    const buAuto = prop((l) => child(l, 'a:buAutoNum')?.attrs.type)
    const spcBef = prop((l) => child(child(l, 'a:spcBef'), 'a:spcPts')?.attrs.val)
    const spcAft = prop((l) => child(child(l, 'a:spcAft'), 'a:spcPts')?.attrs.val)
    const buStart = prop((l) => child(l, 'a:buAutoNum')?.attrs.startAt)
    const lnPct = prop((l) => child(child(l, 'a:lnSpc'), 'a:spcPct')?.attrs.val)
    const lnPts = prop((l) => child(child(l, 'a:lnSpc'), 'a:spcPts')?.attrs.val)
    let runs = ''
    let firstSz
    for (const r of p.kids) {
      if (r.name === 'a:br') { runs += '<br>'; continue }
      if (r.name !== 'a:r' && r.name !== 'a:fld') continue
      const rpr = child(r, 'a:rPr')
      const sz = +(rpr?.attrs.sz ?? defSz ?? (isTitle(ph?.attrs.type) ? 4400 : 1800)) / 100 * scale
      firstSz ??= sz
      const c = (rpr && child(rpr, 'a:solidFill') && ctx.colors.of(child(rpr, 'a:solidFill'))) ?? defColor ?? ctx.colors.of(el('x', {}, [el('a:schemeClr', { val: 'tx1' })]))
      const b = rpr?.attrs.b ?? defB
      const font = child(rpr, 'a:latin')?.attrs.typeface
      const fam = font?.startsWith('+mj') ? faceCss(ctx, font) + ';' : font && !font.startsWith('+') ? `font-family:'${font}',sans-serif;` : ''
      const da = ua ? runData(rpr) : ''
      runs += `<span${da} style="font-size:${(sz * ctx.ptPx).toFixed(1)}px;${c ? `color:${css(c)};` : ''}${b === '1' ? 'font-weight:700;' : ''}${rpr?.attrs.i === '1' ? 'font-style:italic;' : ''}${rpr?.attrs.u && rpr.attrs.u !== 'none' ? 'text-decoration:underline;' : ''}${fam}${runFx(ctx, rpr)}">${escH(textContent(child(r, 'a:t')))}</span>`
    }
    const size = (firstSz ?? +(defSz ?? 1800) / 100 * scale) * ctx.ptPx
    let bullet = ''
    const hasText = paraText(p).trim()
    for (const key of [...counters.keys()]) if (key > lv) counters.delete(key)
    if (hasText && !isTitleType && buNone !== true && (buChar || buAuto)) {
      if (buAuto) {
        const n = (counters.get(lv) ?? (+(buStart ?? 1) - 1)) + 1
        counters.set(lv, n)
        bullet = (AUTONUM[buAuto] ?? AUTONUM.arabicPeriod)(n)
      } else bullet = buChar
    } else if (hasText) counters.delete(lv)
    const ownBu = ua && ppr ? (child(ppr, 'a:buNone') ? 'none' : child(ppr, 'a:buChar') ? 'c:' + child(ppr, 'a:buChar').attrs.char : child(ppr, 'a:buAutoNum') ? 'a:' + child(ppr, 'a:buAutoNum').attrs.type : undefined) : undefined
    const ownLs = ua && child(child(ppr, 'a:lnSpc'), 'a:spcPct')?.attrs.val
    const lh = lnPct ? `line-height:${(1.2 * +lnPct / 100000 * lnRed).toFixed(3)};` : lnPts ? `line-height:${(+lnPts / 100 * ctx.ptPx).toFixed(1)}px;` : lnRed < 1 ? `line-height:${(1.2 * lnRed).toFixed(3)};` : ''
    const align = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] ?? 'left'
    inner += `<p${ua && ppr?.attrs.algn ? ` data-algn="${ppr.attrs.algn}"` : ''}${ua ? ` data-lv="${lv}"` : ''}${ownBu ? ` data-bu="${escH(ownBu)}"` : ''}${ownLs ? ` data-ls="${+ownLs / 1000}"` : ''} style="text-align:${align};padding-left:${Math.max(0, marL)}px;text-indent:${bullet ? indent : 0}px;margin:${spcBef ? (+spcBef / 100 * ctx.ptPx).toFixed(1) : 0}px 0 ${spcAft ? (+spcAft / 100 * ctx.ptPx).toFixed(1) : 0}px;${lh}font-size:${size.toFixed(1)}px">${bullet ? `<span class="bu">${escH(bullet)}</span>` : ''}${runs || '&#8203;'}</p>`
  }
  const just = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[anchor] ?? 'flex-start'
  // Only placeholders and text boxes invite typing; a plain shape stays blank.
  const prompts = ph || xpath(sp, 'p:nvSpPr', 'p:cNvSpPr')?.attrs.txBox === '1'
  const prompt = ua && empty && prompts ? ` data-prompt="${escH(isTitle(ph?.attrs.type) ? '单击此处添加标题' : ph?.attrs.type === 'subTitle' ? '单击此处添加副标题' : '单击此处添加文本')}" data-fs="${(+(lvlProp(chain, 0, (l) => child(l, 'a:defRPr')?.attrs.sz) ?? (isTitle(ph?.attrs.type) ? 4400 : 1800)) / 100 * scale * ctx.ptPx).toFixed(1)}"` : ''
  const body = ua ? ` data-ins="${['lIns', 'tIns', 'rIns', 'bIns'].map((n, i) => Math.round(+(bp(n) ?? (i % 2 ? 45720 : 91440)))).join(',')}" data-fit="${fit}" data-anchor="${anchor}"${vert && vert !== 'horz' ? ` data-vert="${vert}"` : ''}${bp('wrap') === 'none' ? ' data-nowrap="1"' : ''}${scale < 1 ? ` data-scale="${scale}"` : ''}` : ''
  return `<div class="shp txt"${ua}${body}${prompt} style="${style}${fill ? fill + ';' : ''}${fit === 'shape' ? 'overflow:visible;' : ''}${bp('wrap') === 'none' ? 'white-space:pre;' : ''}justify-content:${just};padding:${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px;${vert && vert !== 'horz' ? 'writing-mode:vertical-rl;' : ''}${faceCss(ctx, lvlProp(chain, 0, (l) => child(child(l, 'a:defRPr'), 'a:latin')?.attrs.typeface), lvlProp(chain, 0, (l) => child(child(l, 'a:defRPr'), 'a:ea')?.attrs.typeface))}">${inner}</div>`
}

function renderTable (ctx, tbl, box) {
  const rows = children(tbl, 'a:tr')
  const tp = child(tbl, 'a:tblPr')
  const first = tp?.attrs.firstRow === '1'
  const band = tp?.attrs.bandRow === '1'
  const accent = ctx.colors.scheme.accent1 ?? '4472C4'
  let html = '<table class="tbl">'
  rows.forEach((tr, i) => {
    html += `<tr style="height:${+tr.attrs.h * ctx.k}px">`
    for (const [j, tc] of children(tr, 'a:tc').entries()) {
      if (tc.attrs.hMerge === '1' || tc.attrs.vMerge === '1') continue
      const tcPr = child(tc, 'a:tcPr')
      let bg = tcPr && child(tcPr, 'a:solidFill') ? ctx.colors.of(child(tcPr, 'a:solidFill')) : undefined
      let color
      if (!bg && !(tcPr && child(tcPr, 'a:noFill'))) {
        if (first && i === 0) { bg = accent; color = 'FFFFFF' } else if (band) bg = (first ? i : i + 1) % 2 ? 'CFD5EA' : 'E9EBF5'
      }
      const paras = children(child(tc, 'a:txBody'), 'a:p').map((p) => {
        const runs = children(p, 'a:r').map((r) => {
          const rpr = child(r, 'a:rPr')
          const sz = +(rpr?.attrs.sz ?? 1800) / 100
          const c = (rpr && child(rpr, 'a:solidFill') && ctx.colors.of(child(rpr, 'a:solidFill'))) ?? color
          return `<span style="font-size:${(sz * ctx.ptPx).toFixed(1)}px;${c ? `color:${css(c)};` : ''}${rpr?.attrs.b === '1' || (first && i === 0) ? 'font-weight:700;' : ''}">${escH(textContent(child(r, 'a:t')))}</span>`
        }).join('')
        return `<div>${runs || '&#8203;'}</div>`
      }).join('')
      const span = tc.attrs.gridSpan ? ` colspan="${tc.attrs.gridSpan}"` : ''
      const rspan = tc.attrs.rowSpan ? ` rowspan="${tc.attrs.rowSpan}"` : ''
      html += `<td data-r="${i + 1}" data-c="${j + 1}"${span}${rspan} style="${bg ? `background:${css(bg)};` : ''}">${paras}</td>`
    }
    html += '</tr>'
  })
  return html + '</table>'
}

const PPTX_CSS = `
.deck{padding:16px 0 40px}
.slide-wrap{margin:0 auto 22px;width:960px;max-width:calc(100% - 32px)}
.slide-wrap .num{font:12px -apple-system,"Segoe UI","PingFang SC",sans-serif;color:#777;margin:0 0 4px 2px}
.slide{position:relative;overflow:hidden;box-shadow:0 1px 5px rgba(0,0,0,.22);font-family:Calibri,"PingFang SC","Microsoft YaHei",sans-serif;color:#000}
.shp{position:absolute;box-sizing:border-box}
.txt{display:flex;flex-direction:column;overflow:hidden;line-height:1.2;word-break:break-word}
.pic{overflow:hidden}
.pic img{position:absolute;max-width:none;pointer-events:none}
.txt p{margin:0}
.bu{display:inline-block;min-width:.9em;padding-right:.35em;text-indent:0}
.ph-missing{border:1px dashed #aaa;color:#888;display:flex;align-items:center;justify-content:center;font:13px sans-serif;background:rgba(0,0,0,.03)}
.tbl{border-collapse:collapse;width:100%;height:100%;table-layout:fixed}
.tbl td{border:1px solid #fff;padding:2px 6px;vertical-align:middle;line-height:1.15;overflow:hidden}
.notes{font:13px/1.5 -apple-system,"Segoe UI","PingFang SC",sans-serif;color:#444;background:#fff;border:1px solid #e3e3e3;border-top:none;padding:6px 10px}`

// ------------------------------------------------------------- editing in the UI
//
// pptxUiModel renders each slide with its own shapes marked for the
// editor (data-k is the shape's place in shapes(), data-box its frame in
// EMU) and the run properties each piece of text sets itself.
// editPptxUi applies the person's edits one after another.

// faceCss names the font a text uses: the theme's heading fonts for
// +mj references (titles), its body fonts for +mn or none, else the font set.
function faceCss (ctx, latin, ea) {
  const c = ctx.colors
  const major = String(latin ?? '').startsWith('+mj')
  const l = latin && !latin.startsWith('+') ? latin : major ? c.major : c.minor
  const e = ea && !ea.startsWith('+') ? ea : (major || String(ea ?? '').startsWith('+mj')) ? c.majorEa : c.minorEa
  const faces = [l, e, 'PingFang SC', 'Microsoft YaHei'].filter(Boolean)
  return l || e ? `font-family:${[...new Set(faces)].map((f) => `'${f}'`).join(',')},sans-serif` : ''
}

function runData (rpr) {
  if (!rpr) return ''
  const a = rpr.attrs
  const fill = child(rpr, 'a:solidFill')
  const rgb = child(fill, 'a:srgbClr')?.attrs.val
  const face = child(rpr, 'a:ea')?.attrs.typeface ?? child(rpr, 'a:latin')?.attrs.typeface
  const font = face && !face.startsWith('+') ? ` data-font="${escH(face)}"` : ''
  return `${font}${a.b !== undefined ? ` data-b="${a.b}"` : ''}${a.i !== undefined ? ` data-i="${a.i}"` : ''}${a.u !== undefined ? ` data-u="${a.u === 'none' ? 0 : 1}"` : ''}${a.sz !== undefined ? ` data-sz="${a.sz}"` : ''}${rgb ? ` data-c="${rgb}"` : ''}`
}

export function pptxUiModel (buf) {
  const pkg = Package.fromBuffer(buf)
  const deck = new Deck(pkg)
  const W = 960
  const k = W / deck.cx
  const H = Math.round(deck.cy * k)
  const ptPx = W / (deck.cx / 12700)
  const slides = []
  for (const s of deck.slides()) {
    if (!s.part || !pkg.has(s.part)) { slides.push({ n: s.n, html: '', missing: true }); continue }
    const layout = deck.layoutOf(s.part)
    const master = layout && deck.masterOf(layout)
    const colors = new Colors(deck, master)
    const sd = pkg.xml(s.part)
    const ld = layout && pkg.has(layout) ? pkg.xml(layout) : undefined
    const md = master && pkg.has(master) ? pkg.xml(master) : undefined
    const tree = xpath(sd.root, 'p:cSld', 'p:spTree')
    const list = shapes(tree)
    const ctx = { deck, pkg, colors, k, ptPx, layout, master, ld, md, ui: { index: new Map(list.map((n, i) => [n, i])), groups: new Map(descendants(tree, 'p:grpSp').map((g, i) => [g, i])) } }
    const bg = background(ctx, [[sd, s.part], [ld, layout], [md, master]])
    let html = ''
    const showMaster = (d) => d && d.root.attrs.showMasterSp !== '0'
    if (md && showMaster(sd) && showMaster(ld)) html += renderTree({ ...ctx, ui: undefined }, xpath(md.root, 'p:cSld', 'p:spTree'), master, { decorOnly: true })
    if (ld && showMaster(sd)) html += renderTree({ ...ctx, ui: undefined }, xpath(ld.root, 'p:cSld', 'p:spTree'), layout, { decorOnly: true })
    html += renderTree(ctx, tree, s.part, {})
    const np = deck.notesOf(s.part)
    let notes = ''
    if (np && pkg.has(np)) {
      const body = placeholders(pkg.xml(np).root).find((p) => p.type === 'body')
      notes = body ? children(child(body.sp, 'p:txBody'), 'a:p').map(paraText).join('\n') : ''
    }
    const anims = animationsOf(sd.root)
    slides.push({ n: s.n, html, bg, notes, hidden: sd.root.attrs.show === '0', layout: ld ? child(ld.root, 'p:cSld')?.attrs.name ?? '' : '', layoutPart: layout, transition: transitionOf(sd.root), anims: anims.list, ...(anims.locked ? { animLocked: true } : {}), ...(child(child(sd.root, 'p:cSld'), 'p:bg') ? { ownBg: true } : {}) })
  }
  return {
    kind: 'pptx',
    width: W,
    height: H,
    emu: k,
    css: PPTX_CSS,
    slides,
    layouts: deck.layouts().map((l) => ({ part: l.part, name: l.name || l.type, type: l.type })),
    masters: themeParts(deck).masters.filter((m) => pkg.has(m)).flatMap((m) => [partView(deck, m, k, ptPx), ...children(child(pkg.xml(m).root, 'p:sldLayoutIdLst'), 'p:sldLayoutId').map((lid) => pkg.rel(m, lid.attrs['r:id'])?.part).filter((x) => x && pkg.has(x)).map((x) => ({ ...partView(deck, x, k, ptPx), master: m }))]),
    theme: themeInfo(deck)
  }
}

const RPR_KIDS = ['a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst']

function enforceA (rpr, want) {
  const a = rpr.attrs
  const setFlag = (key) => {
    const cur = a[key] === undefined ? undefined : a[key] === '1' || a[key] === 'true'
    if (cur === want[key]) return
    if (want[key] === undefined) delete a[key]; else a[key] = want[key] ? '1' : '0'
  }
  setFlag('b'); setFlag('i')
  const curU = a.u === undefined ? undefined : a.u !== 'none'
  if (curU !== want.u) { if (want.u === undefined) delete a.u; else a.u = want.u ? 'sng' : 'none' }
  const curSz = a.sz === undefined ? undefined : +a.sz
  if (curSz !== want.sz) { if (want.sz === undefined) delete a.sz; else a.sz = String(Math.round(want.sz)) }
  const curC = child(child(rpr, 'a:solidFill'), 'a:srgbClr')?.attrs.val?.toUpperCase()
  if (curC !== want.color && (want.color || curC)) {
    rpr.kids = rpr.kids.filter((x) => !['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'].includes(x.name))
    if (want.color) insertOrdered(rpr, el('a:solidFill', {}, [el('a:srgbClr', { val: hex(want.color) })]), RPR_KIDS)
  }
  if (want.font) {
    rpr.kids = rpr.kids.filter((x) => x.name !== 'a:latin' && x.name !== 'a:ea')
    insertOrdered(rpr, el('a:latin', { typeface: want.font }), RPR_KIDS)
    insertOrdered(rpr, el('a:ea', { typeface: want.font }), RPR_KIDS)
  }
  return rpr
}

function writeShapeText (sp, paras) {
  let tb = child(sp, 'p:txBody')
  if (!tb) { tb = el('p:txBody', {}, [el('a:bodyPr'), el('a:lstStyle')]); sp.kids.push(tb) }
  const oldPs = children(tb, 'a:p')
  const owners = []
  let old = ''
  oldPs.forEach((p, i) => {
    if (i) { owners.push(null); old += '\n' }
    for (const r of p.kids) {
      if (r.name === 'a:r' || r.name === 'a:fld') { const s = textContent(child(r, 'a:t')); for (let j = 0; j < s.length; j++) owners.push(r); old += s } else if (r.name === 'a:br') { owners.push(null); old += '\v' }
    }
  })
  const neu = paras.map((p) => (p.runs ?? []).map((r) => String(r.text ?? '').replace(/\n/g, '\v')).join('')).join('\n')
  let pre = 0
  while (pre < old.length && pre < neu.length && old[pre] === neu[pre]) pre++
  let suf = 0
  while (suf < old.length - pre && suf < neu.length - pre && old[old.length - 1 - suf] === neu[neu.length - 1 - suf]) suf++
  const nearest = (j) => { for (let d = 0; d < owners.length; d++) { if (owners[j - d]) return owners[j - d]; if (owners[j + d]) return owners[j + d] } return undefined }
  const ownerAt = (j) => {
    if (!old.length) return undefined
    const o = j < pre ? j : j >= neu.length - suf ? j - neu.length + old.length : Math.max(0, pre - 1)
    return owners[o] ?? nearest(Math.min(o, owners.length - 1))
  }
  let pos = 0
  const newPs = paras.map((p, i) => {
    const src = oldPs[i] ?? oldPs[oldPs.length - 1]
    const kids = []
    const ppr = src && child(src, 'a:pPr')
    const pp = ppr ? clone(ppr) : el('a:pPr')
    if (p.algn !== undefined) { if (p.algn) pp.attrs.algn = p.algn; else delete pp.attrs.algn }
    if (p.lvl !== undefined) { if (+p.lvl) pp.attrs.lvl = String(p.lvl); else delete pp.attrs.lvl }
    if (p.bu !== undefined) setBullet(pp, p.bu, !phOf(sp))
    if (p.ls !== undefined) { pp.kids = pp.kids.filter((x) => x.name !== 'a:lnSpc'); if (+p.ls) insertOrdered(pp, el('a:lnSpc', {}, [el('a:spcPct', { val: String(Math.round(+p.ls * 1000)) })]), PPR_KIDS) }
    if (Object.keys(pp.attrs).length || pp.kids.length) kids.push(pp)
    const end = src && child(src, 'a:endParaRPr')
    for (const r of p.runs ?? []) {
      const parts = String(r.text ?? '').split('\n')
      parts.forEach((t, j) => {
        if (j) { kids.push(el('a:br', {}, [end ? Object.assign(clone(end), { name: 'a:rPr' }) : null])); pos++ }
        if (!t) return
        const o = ownerAt(pos)
        const base = o && child(o, 'a:rPr') ? clone(child(o, 'a:rPr')) : end ? Object.assign(clone(end), { name: 'a:rPr' }) : el('a:rPr', { lang: 'zh-CN', altLang: 'en-US' })
        kids.push(el('a:r', {}, [enforceA(base, r), el('a:t', {}, [text(t)])]))
        pos += t.length
      })
    }
    pos++ // the paragraph break
    if (end) kids.push(clone(end))
    return el('a:p', {}, kids)
  })
  if (!newPs.length) newPs.push(el('a:p'))
  tb.kids = [...tb.kids.filter((x) => x.name !== 'a:p'), ...newPs]
}

// boxOf reads a box in EMU from e, defaulting to fractions of the slide.
function boxOf (deck, e, fx, fy, fw, fh) {
  const n = (v, d) => (Number.isFinite(+v) && v !== null && v !== undefined && v !== '' ? Math.round(+v) : Math.round(d))
  return { x: n(e.x, deck.cx * fx), y: n(e.y, deck.cy * fy), cx: Math.max(1, n(e.cx, deck.cx * fw)), cy: Math.max(1, n(e.cy, deck.cy * fh)) }
}

const darker = (h) => h.match(/../g).map((x) => Math.round(parseInt(x, 16) * 0.75).toString(16).padStart(2, '0')).join('').toUpperCase()

// writeCellText puts plain lines in a table cell, keeping the formatting
// of its first run.
function writeCellText (tc, value) {
  let tb = child(tc, 'a:txBody')
  if (!tb) { tb = el('a:txBody', {}, [el('a:bodyPr'), el('a:lstStyle')]); tc.kids.unshift(tb) }
  const first = children(tb, 'a:p')[0]
  const r0 = first && descendants(first, 'a:r')[0]
  const rpr = r0 && child(r0, 'a:rPr')
  const ppr = first && child(first, 'a:pPr')
  const end = first && child(first, 'a:endParaRPr')
  const paras = value.split('\n').map((line) => el('a:p', {}, [
    ppr ? clone(ppr) : null,
    line ? el('a:r', {}, [rpr ? clone(rpr) : end ? Object.assign(clone(end), { name: 'a:rPr' }) : el('a:rPr', { lang: 'zh-CN', altLang: 'en-US' }), el('a:t', {}, [text(line)])]) : null,
    end ? clone(end) : null
  ]))
  tb.kids = [...tb.kids.filter((x) => x.name !== 'a:p'), ...paras]
}

// ------------------------------------------------------------- effects in CSS

const alphaNum = (n) => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26) } return s }
const romanNum = (n) => { let s = ''; for (const [v, r] of [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]) while (n >= v) { s += r; n -= v } return s }
const chsNum = (n) => { const d = '零一二三四五六七八九'; if (n < 10) return d[n]; if (n < 20) return '十' + (n % 10 ? d[n % 10] : ''); if (n < 100) return d[Math.floor(n / 10)] + '十' + (n % 10 ? d[n % 10] : ''); return String(n) }
// AUTONUM formats the automatic numbering schemes the editor offers.
const AUTONUM = {
  arabicPeriod: (n) => n + '.',
  arabicParenR: (n) => n + ')',
  arabicParenBoth: (n) => `(${n})`,
  arabicPlain: (n) => String(n),
  alphaLcPeriod: (n) => alphaNum(n) + '.',
  alphaUcPeriod: (n) => alphaNum(n).toUpperCase() + '.',
  alphaLcParenR: (n) => alphaNum(n) + ')',
  alphaUcParenR: (n) => alphaNum(n).toUpperCase() + ')',
  romanLcPeriod: (n) => romanNum(n).toLowerCase() + '.',
  romanUcPeriod: (n) => romanNum(n) + '.',
  ea1ChsPeriod: (n) => chsNum(n) + '．',
  ea1ChsPlain: (n) => chsNum(n),
  circleNumDbPlain: (n) => n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : `(${n})`
}

function gradCss (ctx, gf) {
  const stops = descendants(gf, 'a:gs').map((g) => [+g.attrs.pos, ctx.colors.of(g)]).filter((s) => s[1]).sort((a, b) => a[0] - b[0]).map(([p, c]) => `${css(c)} ${p / 1000}%`)
  if (!stops.length) return ''
  if (child(gf, 'a:path')) return `radial-gradient(circle,${stops.join(',')})`
  const ang = +(child(gf, 'a:lin')?.attrs.ang ?? 5400000) / 60000
  return `linear-gradient(${(ang + 90) % 360}deg,${stops.join(',')})`
}

function shadowCss (ctx, sh) {
  const dist = +(sh.attrs.dist ?? 0) * ctx.k
  const dir = +(sh.attrs.dir ?? 0) / 60000 * Math.PI / 180
  const blur = +(sh.attrs.blurRad ?? 0) * ctx.k
  const c = ctx.colors.of(sh) ?? '00000066'
  return `${(dist * Math.cos(dir)).toFixed(1)}px ${(dist * Math.sin(dir)).toFixed(1)}px ${blur.toFixed(1)}px ${css(c)}`
}

// runFx turns WordArt-style run effects (outline, gradient fill, shadow,
// glow) into CSS.
function runFx (ctx, rpr) {
  if (!rpr) return ''
  let s = ''
  const ln = child(rpr, 'a:ln')
  if (ln && !child(ln, 'a:noFill')) {
    const c = ctx.colors.of(child(ln, 'a:solidFill'))
    if (c) s += `-webkit-text-stroke:${Math.max(0.5, +(ln.attrs.w ?? 12700) * ctx.k).toFixed(1)}px ${css(c)};paint-order:stroke fill;`
  }
  const gf = child(rpr, 'a:gradFill')
  if (gf) { const g = gradCss(ctx, gf); if (g) s += `background:${g};-webkit-background-clip:text;background-clip:text;color:transparent;` }
  const fx = child(rpr, 'a:effectLst')
  const shadows = []
  const sh = child(fx, 'a:outerShdw')
  if (sh) shadows.push(shadowCss(ctx, sh))
  const gl = child(fx, 'a:glow')
  if (gl) { const c = ctx.colors.of(gl); if (c) shadows.push(`0 0 ${Math.max(2, +(gl.attrs.rad ?? 0) * ctx.k).toFixed(1)}px ${css(c)}`, `0 0 ${Math.max(1, +(gl.attrs.rad ?? 0) * ctx.k / 2).toFixed(1)}px ${css(c)}`) }
  // A text shadow would cover gradient-clipped text; drop-shadow sits behind it.
  if (shadows.length) s += gf ? `filter:${shadows.map((x) => `drop-shadow(${x})`).join(' ')};` : `text-shadow:${shadows.join(',')};`
  return s
}

// ------------------------------------------------------------- reading motion

const TRANSITIONS = ['fade', 'push', 'wipe', 'split', 'cover', 'pull', 'zoom', 'dissolve', 'cut', 'randomBar']

function transitionOf (root) {
  let t = child(root, 'p:transition')
  if (!t) {
    const alt = root.kids.find((k) => k.name === 'mc:AlternateContent' && descendants(k, 'p:transition').length)
    if (alt) t = descendants(child(alt, 'mc:Fallback') ?? alt, 'p:transition')[0] ?? descendants(alt, 'p:transition')[0]
  }
  if (!t) return undefined
  const kind = t.kids.find(isEl)
  const name = kind ? kind.name.replace(/^\w+:/, '') : ''
  const out = { type: TRANSITIONS.includes(name) ? name : name ? 'other' : 'none', spd: t.attrs.spd ?? 'fast' }
  if (kind?.attrs.dir) out.dir = kind.attrs.dir
  if (kind?.attrs.orient) out.orient = kind.attrs.orient
  if (t.attrs.advTm !== undefined) out.adv = +t.attrs.advTm / 1000
  if (t.attrs.advClick === '0') out.noClick = true
  return out
}

const ANIM_PRESETS = { 1: 'appear', 2: 'fly', 10: 'fade', 22: 'wipe', 23: 'zoom', 53: 'zoom' }
const ANIM_IDS = { appear: 1, fly: 2, fade: 10, wipe: 22, zoom: 23 }
const DIR_SUB = { t: 1, r: 2, b: 4, l: 8 }
const SUB_DIR = { 1: 't', 2: 'r', 4: 'b', 8: 'l' }

// animationsOf lists the entrance effects of a slide's main sequence in
// order. locked says the slide holds effects the editor cannot rewrite.
function animationsOf (root) {
  const timing = child(root, 'p:timing')
  if (!timing) return { list: [] }
  const list = []
  let locked = false
  for (const tn of descendants(timing, 'p:cTn')) {
    const cls = tn.attrs.presetClass
    if (!cls) continue
    if (cls !== 'entr' || !ANIM_PRESETS[tn.attrs.presetID]) { locked = true; continue }
    const spid = descendants(tn, 'p:spTgt')[0]?.attrs.spid
    if (!spid) { locked = true; continue }
    const timed = descendants(tn, 'p:cTn').find((c) => c !== tn && +c.attrs.dur > 1)
    list.push({
      spid,
      effect: ANIM_PRESETS[tn.attrs.presetID],
      start: tn.attrs.nodeType === 'withEffect' ? 'with' : tn.attrs.nodeType === 'afterEffect' ? 'after' : 'click',
      dur: timed ? +timed.attrs.dur : 500,
      ...(SUB_DIR[tn.attrs.presetSubtype] ? { dir: SUB_DIR[tn.attrs.presetSubtype] } : {})
    })
  }
  if (descendants(timing, 'p:seq').some((s) => xpath(s, 'p:cTn')?.attrs.nodeType === 'interactiveSeq')) locked = true
  return locked ? { list, locked } : { list }
}

// ------------------------------------------------------------- writing motion

function setTransition (root, t) {
  root.kids = root.kids.filter((k) => !(k.name === 'p:transition' || (k.name === 'mc:AlternateContent' && descendants(k, 'p:transition').length)))
  const type = t?.type ?? 'none'
  const adv = t?.adv === undefined || t?.adv === null || t?.adv === '' ? undefined : Math.max(0, +t.adv)
  if (type === 'none' && adv === undefined) return
  const attrs = { spd: ['slow', 'med', 'fast'].includes(t.spd) ? t.spd : 'med' }
  if (adv !== undefined && Number.isFinite(adv)) attrs.advTm = String(Math.round(adv * 1000))
  if (t.noClick) attrs.advClick = '0'
  const dir = ['l', 'r', 'u', 'd'].includes(t.dir) ? t.dir : undefined
  let kid = null
  if (type === 'fade') kid = el('p:fade')
  else if (type === 'push' || type === 'wipe' || type === 'cover' || type === 'pull') kid = el('p:' + type, dir ? { dir } : {})
  else if (type === 'split') kid = el('p:split', { orient: t.orient === 'vert' ? 'vert' : 'horz', dir: t.dir === 'in' ? 'in' : 'out' })
  else if (type === 'zoom') kid = el('p:zoom')
  else if (type === 'dissolve') kid = el('p:dissolve')
  else if (type !== 'none') throw new Error(`transition type "${type}" is not one of none, fade, push, wipe, split, cover, pull, zoom, dissolve`)
  const node = el('p:transition', attrs, kid ? [kid] : [])
  const at = root.kids.findIndex((k) => k.name === 'p:timing' || k.name === 'p:extLst')
  root.kids.splice(at < 0 ? root.kids.length : at, 0, node)
}

function setAnimations (root, list, spids, where) {
  const cur = animationsOf(root)
  if (cur.locked) throw new Error(`${where}: 这张幻灯片含有强调、退出、路径或触发器动画，请在 PowerPoint 中编辑`)
  root.kids = root.kids.filter((k) => k.name !== 'p:timing')
  if (!list.length) return
  let id = 1
  const cTn = (attrs, kids) => el('p:cTn', { id: String(id++), ...attrs }, kids)
  const cond = (delay) => el('p:stCondLst', {}, [el('p:cond', { delay: String(delay) })])
  const tgt = (spid) => el('p:tgtEl', {}, [el('p:spTgt', { spid })])
  const set = (spid) => el('p:set', {}, [el('p:cBhvr', {}, [cTn({ dur: '1', fill: 'hold' }, [cond(0)]), tgt(spid), el('p:attrNameLst', {}, [el('p:attrName', {}, [text('style.visibility')])])]), el('p:to', {}, [el('p:strVal', { val: 'visible' })])])
  const anim = (spid, dur, attr, from, to) => el('p:anim', { calcmode: 'lin', valueType: 'num' }, [
    el('p:cBhvr', { additive: 'base' }, [cTn({ dur: String(dur), fill: 'hold' }), tgt(spid), el('p:attrNameLst', {}, [el('p:attrName', {}, [text(attr)])])]),
    el('p:tavLst', {}, [el('p:tav', { tm: '0' }, [el('p:val', {}, [el('p:strVal', { val: from })])]), el('p:tav', { tm: '100000' }, [el('p:val', {}, [el('p:strVal', { val: to })])])])
  ])
  const effect = (spid, filter, dur) => el('p:animEffect', { transition: 'in', filter }, [el('p:cBhvr', {}, [cTn({ dur: String(dur) }), tgt(spid)])])
  const behaviours = (a) => {
    const d = Math.max(100, Math.min(20000, Math.round(+a.dur || 500)))
    const out = [set(a.spid)]
    if (a.effect === 'fade') out.push(effect(a.spid, 'fade', d))
    if (a.effect === 'wipe') out.push(effect(a.spid, `wipe(${{ b: 'up', t: 'down', l: 'right', r: 'left' }[a.dir ?? 'b']})`, d))
    if (a.effect === 'fly') {
      const dir = a.dir ?? 'b'
      out.push(anim(a.spid, d, 'ppt_x', dir === 'l' ? '0-#ppt_w/2' : dir === 'r' ? '1+#ppt_w/2' : '#ppt_x', '#ppt_x'))
      out.push(anim(a.spid, d, 'ppt_y', dir === 't' ? '0-#ppt_h/2' : dir === 'b' ? '1+#ppt_h/2' : '#ppt_y', '#ppt_y'))
    }
    if (a.effect === 'zoom') { out.push(anim(a.spid, d, 'ppt_w', '0', '#ppt_w')); out.push(anim(a.spid, d, 'ppt_h', '0', '#ppt_h')) }
    return { d, kids: out }
  }
  // Click groups hold steps; a step holds effects that start together.
  const groups = []
  for (const a of list) {
    if (!ANIM_IDS[a.effect]) throw new Error(`${where}: effect "${a.effect}" is not one of appear, fade, fly, wipe, zoom`)
    if (!spids.has(String(a.spid))) throw new Error(`${where}: the slide has no shape with id ${a.spid}`)
    if (a.start === 'click' || !groups.length) groups.push({ steps: [{ at: 0, items: [] }], end: 0 })
    const g = groups[groups.length - 1]
    if (a.start === 'after') g.steps.push({ at: g.end, items: [] })
    const step = g.steps[g.steps.length - 1]
    step.items.push(a)
    g.end = Math.max(g.end, step.at + Math.max(100, Math.round(+a.dur || 500)))
  }
  const clicks = groups.map((g, gi) => {
    const first = list.indexOf(g.steps[0].items[0])
    const auto = gi === 0 && list[first].start !== 'click'
    return el('p:par', {}, [cTn({ fill: 'hold' }, [cond(auto ? 0 : 'indefinite'), el('p:childTnLst', {}, g.steps.map((st) => el('p:par', {}, [cTn({ fill: 'hold' }, [cond(st.at), el('p:childTnLst', {}, st.items.map((a) => {
      const sub = a.effect === 'fly' || a.effect === 'wipe' ? DIR_SUB[a.dir ?? 'b'] : a.effect === 'zoom' ? 16 : 0
      const node = el('p:par', {}, [])
      const c = cTn({ presetID: String(ANIM_IDS[a.effect]), presetClass: 'entr', presetSubtype: String(sub), fill: 'hold', grpId: '0', nodeType: a.start === 'with' ? 'withEffect' : a.start === 'after' ? 'afterEffect' : 'clickEffect' }, [cond(0)])
      c.kids.push(el('p:childTnLst', {}, behaviours(a).kids))
      node.kids.push(c)
      return node
    }))])])))])])
  })
  const seqTn = cTn({ dur: 'indefinite', nodeType: 'mainSeq' }, [el('p:childTnLst', {}, clicks)])
  const root1 = el('p:tnLst', {}, [el('p:par', {}, [el('p:cTn', { id: '1', dur: 'indefinite', restart: 'never', nodeType: 'tmRoot' }, [el('p:childTnLst', {}, [
    el('p:seq', { concurrent: '1', nextAc: 'seek' }, [seqTn,
      el('p:prevCondLst', {}, [el('p:cond', { evt: 'onPrev', delay: '0' }, [el('p:tgtEl', {}, [el('p:sldTgt')])])]),
      el('p:nextCondLst', {}, [el('p:cond', { evt: 'onNext', delay: '0' }, [el('p:tgtEl', {}, [el('p:sldTgt')])])])])
  ])])])])
  // The root keeps id 1; renumber the rest in document order after it.
  let n = 2
  for (const c of descendants(root1, 'p:cTn')) if (c.attrs.nodeType !== 'tmRoot') c.attrs.id = String(n++)
  const bld = el('p:bldLst', {}, [...new Set(list.map((a) => String(a.spid)))].map((spid) => el('p:bldP', { spid, grpId: '0', animBg: '1' })))
  const node = el('p:timing', {}, [root1, bld])
  const at = root.kids.findIndex((k) => k.name === 'p:extLst')
  root.kids.splice(at < 0 ? root.kids.length : at, 0, node)
}

// ------------------------------------------------------------- structure

function hostOf (deck, e, where) {
  if (e.part) {
    const ct = deck.pkg.contentType(e.part)
    if (!deck.pkg.has(e.part) || (ct !== CT.master && ct !== CT.layout)) throw new Error(`${where}: ${e.part} is not a slide master or layout`)
    return { part: e.part }
  }
  return deck.slide(e.slide, where)
}

function shapeOf (deck, e, where) {
  const s = hostOf(deck, e, where)
  const tree = xpath(deck.pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
  const node = e.group ? descendants(tree, 'p:grpSp')[+e.group - 1] : shapes(tree)[+e.shape - 1]
  if (!node) throw new Error(`${where}: ${e.group ? `there is no group ${e.group}` : `there is no shape ${e.shape}`} here; reload the presentation`)
  return { s, tree, node }
}

const spPrOf = (node) => {
  const name = node.name === 'p:grpSp' ? 'p:grpSpPr' : 'p:spPr'
  let pr = child(node, name)
  if (!pr) { pr = el(name); node.kids.splice(1, 0, pr) }
  return pr
}

function setBox (node, b) {
  let x
  if (node.name === 'p:graphicFrame') {
    x = child(node, 'p:xfrm')
    if (!x) { x = el('p:xfrm'); node.kids.splice(1, 0, x) }
  } else {
    const pr = spPrOf(node)
    x = child(pr, 'a:xfrm')
    if (!x) { x = el('a:xfrm'); pr.kids.unshift(x) }
  }
  const off = child(x, 'a:off') ?? (x.kids.unshift(el('a:off')), child(x, 'a:off'))
  const ext = child(x, 'a:ext') ?? (x.kids.splice(1, 0, el('a:ext')), child(x, 'a:ext'))
  if (node.name === 'p:grpSp' && !child(x, 'a:chOff')) {
    x.kids.push(el('a:chOff', { x: off.attrs.x ?? String(Math.round(b.x)), y: off.attrs.y ?? String(Math.round(b.y)) }), el('a:chExt', { cx: ext.attrs.cx ?? String(Math.round(b.cx)), cy: ext.attrs.cy ?? String(Math.round(b.cy)) }))
  }
  off.attrs.x = String(Math.round(b.x)); off.attrs.y = String(Math.round(b.y))
  ext.attrs.cx = String(Math.max(1, Math.round(b.cx))); ext.attrs.cy = String(Math.max(1, Math.round(b.cy)))
  if (b.rot !== undefined && node.name !== 'p:graphicFrame') {
    const r = ((Math.round(+b.rot * 60000) % 21600000) + 21600000) % 21600000
    if (r) x.attrs.rot = String(r); else delete x.attrs.rot
  }
  return x
}

function groupShapes (deck, tree, nodes, where) {
  if (nodes.length < 2) throw new Error(`${where}: select at least two shapes to group`)
  for (const n of nodes) {
    if (!tree.kids.includes(n)) throw new Error(`${where}: only shapes directly on the slide can be grouped`)
    if (phOf(n)) throw new Error(`${where}: 占位符不能组合`)
  }
  const boxes = nodes.map((n) => xfrmOf(n))
  if (boxes.some((b) => !b)) throw new Error(`${where}: a shape has no position of its own`)
  const x1 = Math.min(...boxes.map((b) => b.x)); const y1 = Math.min(...boxes.map((b) => b.y))
  const x2 = Math.max(...boxes.map((b) => b.x + b.cx)); const y2 = Math.max(...boxes.map((b) => b.y + b.cy))
  const id = deck.nextShapeId(tree)
  const ordered = tree.kids.filter((k) => nodes.includes(k))
  const at = tree.kids.indexOf(ordered[ordered.length - 1])
  const g = el('p:grpSp', {}, [
    el('p:nvGrpSpPr', {}, [el('p:cNvPr', { id: String(id), name: `组合 ${id}` }), el('p:cNvGrpSpPr'), el('p:nvPr')]),
    el('p:grpSpPr', {}, [el('a:xfrm', {}, [el('a:off', { x: String(x1), y: String(y1) }), el('a:ext', { cx: String(x2 - x1), cy: String(y2 - y1) }), el('a:chOff', { x: String(x1), y: String(y1) }), el('a:chExt', { cx: String(x2 - x1), cy: String(y2 - y1) })])]),
    ...ordered
  ])
  tree.kids.splice(at, 1, g)
  tree.kids = tree.kids.filter((k) => k === g || !ordered.includes(k))
  return g
}

function ungroupShape (tree, g, where) {
  const parent = [tree, ...descendants(tree, 'p:grpSp')].find((p) => p.kids.includes(g))
  if (!parent) throw new Error(`${where}: group not found`)
  const x = xpath(g, 'p:grpSpPr', 'a:xfrm')
  const off = child(x, 'a:off'); const ext = child(x, 'a:ext'); const choff = child(x, 'a:chOff'); const chext = child(x, 'a:chExt')
  const sx = chext && +chext.attrs.cx ? +ext.attrs.cx / +chext.attrs.cx : 1
  const sy = chext && +chext.attrs.cy ? +ext.attrs.cy / +chext.attrs.cy : 1
  const rot = +(x?.attrs.rot ?? 0) / 60000
  const gc = off && ext ? { x: +off.attrs.x + +ext.attrs.cx / 2, y: +off.attrs.y + +ext.attrs.cy / 2 } : { x: 0, y: 0 }
  const kids = g.kids.filter((k) => isEl(k) && !['p:nvGrpSpPr', 'p:grpSpPr', 'p:extLst'].includes(k.name))
  for (const k of kids) {
    const b = xfrmOf(k)
    if (!b || !off) continue
    const nb = { x: +off.attrs.x + (b.x - +(choff?.attrs.x ?? off.attrs.x)) * sx, y: +off.attrs.y + (b.y - +(choff?.attrs.y ?? off.attrs.y)) * sy, cx: b.cx * sx, cy: b.cy * sy }
    if (rot) {
      const cx = nb.x + nb.cx / 2 - gc.x; const cy = nb.y + nb.cy / 2 - gc.y
      const a = rot * Math.PI / 180
      nb.x = gc.x + cx * Math.cos(a) - cy * Math.sin(a) - nb.cx / 2
      nb.y = gc.y + cx * Math.sin(a) + cy * Math.cos(a) - nb.cy / 2
      nb.rot = (b.rot ?? 0) + rot
    }
    setBox(k, nb)
  }
  parent.kids.splice(parent.kids.indexOf(g), 1, ...kids)
  return kids
}

function changeLayout (deck, s, layout, where) {
  const pkg = deck.pkg
  const oldLayout = deck.layoutOf(s.part)
  const tree = xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
  const KEEP = ['dt', 'ftr', 'sldNum', 'hdr']
  const fresh = placeholders(pkg.xml(layout.part).root).filter((p) => !KEEP.includes(p.type))
  const kind = (t) => isTitle(t) ? 'title' : t === 'subTitle' ? 'sub' : isContent(t) ? 'body' : t
  const used = new Set()
  const pick = (want) => fresh.find((p) => !used.has(p) && kind(p.type) === want)
  for (const node of tree.kids.filter((k) => isEl(k) && ['p:sp', 'p:pic', 'p:graphicFrame'].includes(k.name))) {
    const ph = phOf(node)
    if (!ph || KEEP.includes(ph.attrs.type)) continue
    const want = kind(ph.attrs.type)
    const m = pick(want) ?? (want === 'sub' ? pick('body') : want === 'body' ? pick('sub') : undefined)
    if (m) {
      used.add(m)
      for (const key of ['type', 'idx', 'orient', 'sz']) { if (m.ph.attrs[key] !== undefined) ph.attrs[key] = m.ph.attrs[key]; else delete ph.attrs[key] }
      const pr = child(node, 'p:spPr')
      if (pr) pr.kids = pr.kids.filter((k) => k.name !== 'a:xfrm')
      continue
    }
    const tb = child(node, 'p:txBody')
    const hasText = tb && children(tb, 'a:p').some((p) => paraText(p).trim())
    if (hasText || node.name !== 'p:sp') {
      const box = geometry(deck, node, oldLayout)
      if (box && !xfrmOf(node)) setBox(node, box)
    } else remove(tree, node)
  }
  let id = deck.nextShapeId(tree)
  for (const p of fresh) {
    if (used.has(p)) continue
    const ph = {}
    for (const key of ['type', 'orient', 'sz', 'idx']) if (p.ph.attrs[key] !== undefined) ph[key] = p.ph.attrs[key]
    const name = (xpath(p.sp, 'p:nvSpPr', 'p:cNvPr')?.attrs.name ?? 'Placeholder').replace(/\s*\d+$/, '') + ' ' + id
    tree.kids.push(el('p:sp', {}, [
      el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id++), name }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1' })]), el('p:nvPr', {}, [el('p:ph', ph)])]),
      el('p:spPr'),
      el('p:txBody', {}, [el('a:bodyPr'), el('a:lstStyle'), el('a:p', {}, [el('a:endParaRPr', { lang: 'zh-CN', altLang: 'en-US' })])])
    ]))
  }
  for (const r of pkg.relsOfType(s.part, 'slideLayout')) pkg.removeRel(s.part, r.id)
  pkg.addRel(s.part, REL('slideLayout'), layout.part)
  pkg.touch(s.part)
}

function themeParts (deck) {
  const masters = children(child(deck.pres.root, 'p:sldMasterIdLst'), 'p:sldMasterId').map((m) => deck.pkg.rel(deck.main, m.attrs['r:id'])?.part).filter(Boolean)
  return { masters, themes: [...new Set(masters.map((m) => deck.pkg.relsOfType(m, 'theme')[0]?.part).filter((t) => t && deck.pkg.has(t)))] }
}

const SCHEME_KEYS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

function setTheme (deck, e, where) {
  const { themes } = themeParts(deck)
  if (!themes.length) throw new Error(`${where}: the presentation has no theme`)
  for (const t of themes) {
    const root = deck.pkg.xml(t).root
    if (e.colors) {
      const cs = descendants(root, 'a:clrScheme')[0]
      if (!cs) throw new Error(`${where}: the theme has no colour scheme`)
      if (e.name) cs.attrs.name = String(e.name)
      for (const key of SCHEME_KEYS) {
        const v = e.colors[key]
        if (!v) continue
        let node = child(cs, 'a:' + key)
        if (!node) { node = el('a:' + key); cs.kids.push(node) }
        node.kids = [el('a:srgbClr', { val: hex(v) })]
      }
    }
    if (e.fonts) {
      const fs = descendants(root, 'a:fontScheme')[0]
      if (!fs) throw new Error(`${where}: the theme has no font scheme`)
      for (const [kind, f] of [['a:majorFont', e.fonts.major], ['a:minorFont', e.fonts.minor]]) {
        if (!f) continue
        const node = child(fs, kind)
        if (!node) continue
        for (const [tag, face] of [['a:latin', f.latin], ['a:ea', f.ea]]) {
          if (!face) continue
          let n = child(node, tag)
          if (!n) { n = el(tag); node.kids.splice(tag === 'a:latin' ? 0 : 1, 0, n) }
          n.attrs = { typeface: String(face) }
        }
      }
    }
    deck.pkg.touch(t)
  }
}

async function bgNode (deck, part, e, io) {
  if (e.image) {
    const data = await io.loadImage(String(e.image))
    const info = imageInfo(data)
    if (!info) throw new Error('the background picture is not a PNG, JPEG or GIF image')
    const media = deck.pkg.addMedia('ppt/media', info.ext, data)
    const rid = deck.pkg.addRel(part, NS.image, media)
    return el('p:bg', {}, [el('p:bgPr', {}, [el('a:blipFill', { dpi: '0', rotWithShape: '1' }, [el('a:blip', { 'r:embed': rid }), el('a:srcRect'), el('a:stretch', {}, [el('a:fillRect')])]), el('a:effectLst')])])
  }
  if (Array.isArray(e.gradient)) {
    const [c1, c2, ang] = e.gradient
    return el('p:bg', {}, [el('p:bgPr', {}, [el('a:gradFill', { rotWithShape: '1' }, [el('a:gsLst', {}, [el('a:gs', { pos: '0' }, [el('a:srgbClr', { val: hex(c1) })]), el('a:gs', { pos: '100000' }, [el('a:srgbClr', { val: hex(c2) })])]), el('a:lin', { ang: String(Math.round((+ang || 90) * 60000)), scaled: '0' })]), el('a:effectLst')])])
  }
  if (e.fill) return el('p:bg', {}, [el('p:bgPr', {}, [solid(hex(e.fill)), el('a:effectLst')])])
  return undefined
}

function putBg (deck, part, node) {
  const root = deck.pkg.xml(part).root
  const cSld = child(root, 'p:cSld')
  const old = child(cSld, 'p:bg')
  // Pictures the old background used stay in the package only if
  // something else refers to them; the relationship is dropped.
  const rid = old && descendants(old, 'a:blip')[0]?.attrs['r:embed']
  if (old) remove(cSld, old)
  if (rid) deck.pkg.removeRel(part, rid)
  if (node) cSld.kids.unshift(node)
  deck.pkg.touch(part)
}

const BODY_KIDS = ['a:prstTxWarp', 'a:noAutofit', 'a:normAutofit', 'a:spAutoFit', 'a:scene3d', 'a:sp3d', 'a:flatTx', 'a:extLst']

function setBody (sp, e) {
  let tb = child(sp, 'p:txBody')
  if (!tb) { tb = el('p:txBody', {}, [el('a:bodyPr'), el('a:lstStyle'), el('a:p')]); sp.kids.push(tb) }
  let bp = child(tb, 'a:bodyPr')
  if (!bp) { bp = el('a:bodyPr'); tb.kids.unshift(bp) }
  if (Array.isArray(e.ins)) ['lIns', 'tIns', 'rIns', 'bIns'].forEach((k, i) => { const v = e.ins[i]; if (v !== undefined && v !== null && Number.isFinite(+v)) bp.attrs[k] = String(Math.max(0, Math.round(+v))) })
  if (e.anchor !== undefined) { if (['t', 'ctr', 'b'].includes(e.anchor)) bp.attrs.anchor = e.anchor; else delete bp.attrs.anchor }
  if (e.vert !== undefined) { if (['eaVert', 'vert', 'vert270', 'wordArtVertRtl'].includes(e.vert)) bp.attrs.vert = e.vert; else delete bp.attrs.vert }
  if (e.wrap !== undefined) bp.attrs.wrap = e.wrap ? 'square' : 'none'
  if (e.fit !== undefined) {
    bp.kids = bp.kids.filter((k) => !['a:noAutofit', 'a:normAutofit', 'a:spAutoFit'].includes(k.name))
    if (e.fit === 'shape') insertOrdered(bp, el('a:spAutoFit'), BODY_KIDS)
    else if (e.fit === 'shrink') {
      const s = Math.max(0.25, Math.min(1, +e.scale || 1))
      insertOrdered(bp, el('a:normAutofit', s < 1 ? { fontScale: String(Math.round(s * 100000)), lnSpcReduction: s < 0.9 ? '10000' : '0' } : {}), BODY_KIDS)
    } else if (e.fit === 'none') insertOrdered(bp, el('a:noAutofit'), BODY_KIDS)
  } else if (e.scale !== undefined) {
    const na = child(bp, 'a:normAutofit')
    const s = Math.max(0.25, Math.min(1, +e.scale || 1))
    if (na) { if (s < 1) { na.attrs.fontScale = String(Math.round(s * 100000)); na.attrs.lnSpcReduction = s < 0.9 ? '10000' : '0' } else na.attrs = {} }
  }
}

// TEXT_FX are the WordArt styles: what each puts on every run.
const TEXT_FX = {
  shadow: () => ({ effect: el('a:effectLst', {}, [el('a:outerShdw', { blurRad: '38100', dist: '38100', dir: '2700000', algn: 'tl', rotWithShape: '0' }, [el('a:srgbClr', { val: '000000' }, [el('a:alpha', { val: '43000' })])])]) }),
  outline: (acc) => ({ ln: el('a:ln', { w: '15875' }, [solid(acc)]), fill: solid('FFFFFF') }),
  glow: (acc) => ({ effect: el('a:effectLst', {}, [el('a:glow', { rad: '101600' }, [el('a:srgbClr', { val: acc }, [el('a:alpha', { val: '60000' })])])]) }),
  gradient: (acc) => ({ fill: el('a:gradFill', {}, [el('a:gsLst', {}, [el('a:gs', { pos: '0' }, [el('a:srgbClr', { val: acc })]), el('a:gs', { pos: '100000' }, [el('a:srgbClr', { val: darker(acc) })])]), el('a:lin', { ang: '5400000', scaled: '0' })]) }),
  fire: () => ({ fill: el('a:gradFill', {}, [el('a:gsLst', {}, [el('a:gs', { pos: '0' }, [el('a:srgbClr', { val: 'FFC000' })]), el('a:gs', { pos: '100000' }, [el('a:srgbClr', { val: 'C00000' })])]), el('a:lin', { ang: '5400000', scaled: '0' })]), effect: el('a:effectLst', {}, [el('a:outerShdw', { blurRad: '50800', dist: '38100', dir: '2700000', algn: 'tl', rotWithShape: '0' }, [el('a:srgbClr', { val: '000000' }, [el('a:alpha', { val: '40000' })])])]) }),
  neon: (acc) => ({ fill: solid('FFFFFF'), effect: el('a:effectLst', {}, [el('a:glow', { rad: '139700' }, [el('a:srgbClr', { val: acc }, [el('a:alpha', { val: '75000' })])])]) })
}

function setTextFx (sp, preset, accent) {
  if (preset !== 'none' && !TEXT_FX[preset]) throw new Error(`text effect "${preset}" is not one of none, ${Object.keys(TEXT_FX).join(', ')}`)
  const tb = child(sp, 'p:txBody')
  if (!tb) return
  const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']
  for (const p of children(tb, 'a:p')) {
    for (const r of p.kids.filter((k) => k.name === 'a:r' || k.name === 'a:endParaRPr' || k.name === 'a:fld')) {
      let rpr = r.name === 'a:endParaRPr' ? r : child(r, 'a:rPr')
      if (!rpr) { rpr = el('a:rPr', { lang: 'zh-CN', altLang: 'en-US' }); r.kids.unshift(rpr) }
      const fx = preset === 'none' ? {} : TEXT_FX[preset](accent)
      rpr.kids = rpr.kids.filter((k) => k.name !== 'a:ln' && k.name !== 'a:effectLst' && k.name !== 'a:gradFill')
      if (fx.fill) { rpr.kids = rpr.kids.filter((k) => !FILLS.includes(k.name)); insertOrdered(rpr, clone(fx.fill), RPR_KIDS) }
      if (fx.ln) insertOrdered(rpr, clone(fx.ln), RPR_KIDS)
      if (fx.effect) insertOrdered(rpr, clone(fx.effect), RPR_KIDS)
    }
  }
}

function setCrop (node, e, where) {
  if (node.name !== 'p:pic') throw new Error(`${where}: only pictures can be cropped`)
  const bf = child(node, 'p:blipFill')
  if (!bf) throw new Error(`${where}: the picture has no image`)
  let sr = child(bf, 'a:srcRect')
  if (!sr) { sr = el('a:srcRect'); const at = bf.kids.findIndex((k) => k.name === 'a:blip'); bf.kids.splice(at + 1, 0, sr) }
  const v = ['l', 't', 'r', 'b'].map((k) => Math.round((+e[k] || 0) * 100000))
  if (v[0] + v[2] >= 99000 || v[1] + v[3] >= 99000) throw new Error(`${where}: the crop leaves nothing of the picture`)
  sr.attrs = {}
  ;['l', 't', 'r', 'b'].forEach((k, i) => { if (v[i]) sr.attrs[k] = String(v[i]) })
  if (['x', 'y', 'cx', 'cy'].every((k) => Number.isFinite(+e[k]))) setBox(node, { x: +e.x, y: +e.y, cx: +e.cx, cy: +e.cy })
}

// partView renders a master or layout for the master view, its own
// shapes marked for the editor like a slide's.
function partView (deck, part, k, ptPx) {
  const pkg = deck.pkg
  const isMaster = pkg.contentType(part) === CT.master
  const master = isMaster ? part : deck.masterOf(part)
  const d = pkg.xml(part)
  const md = master && pkg.has(master) ? pkg.xml(master) : undefined
  const tree = xpath(d.root, 'p:cSld', 'p:spTree')
  const list = shapes(tree)
  const ctx = { deck, pkg, colors: new Colors(deck, master), k, ptPx, layout: isMaster ? undefined : part, master, ld: isMaster ? undefined : d, md, ui: { index: new Map(list.map((n, i) => [n, i])), groups: new Map(descendants(tree, 'p:grpSp').map((g, i) => [g, i])) } }
  const bg = background(ctx, isMaster ? [[md, master]] : [[d, part], [md, master]])
  let html = ''
  if (!isMaster && md && d.root.attrs.showMasterSp !== '0') html += renderTree({ ...ctx, ui: undefined }, xpath(md.root, 'p:cSld', 'p:spTree'), master, { decorOnly: true })
  html += renderTree(ctx, tree, part, {})
  return { part, kind: isMaster ? 'master' : 'layout', name: child(d.root, 'p:cSld')?.attrs.name ?? (isMaster ? '母版' : ''), html, bg }
}

function themeInfo (deck) {
  const { masters, themes } = themeParts(deck)
  const c = new Colors(deck, masters[0])
  const fs = themes[0] && descendants(deck.pkg.xml(themes[0]).root, 'a:clrScheme')[0]
  return { name: fs?.attrs.name ?? '', colors: Object.fromEntries(SCHEME_KEYS.map((k) => [k, c.scheme[k]]).filter((x) => x[1])), major: c.major, minor: c.minor, majorEa: c.majorEa, minorEa: c.minorEa }
}

const PPR_KIDS = ['a:lnSpc', 'a:spcBef', 'a:spcAft', 'a:buClrTx', 'a:buClr', 'a:buSzTx', 'a:buSzPct', 'a:buSzPts', 'a:buFontTx', 'a:buFont', 'a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip', 'a:tabLst', 'a:defRPr', 'a:extLst']
const BU_KIDS = ['a:buFontTx', 'a:buFont', 'a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip']
const SPPR_KIDS = ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']

// setBullet gives a paragraph its own bullet: 'none', 'c:<char>' or
// 'a:<numbering scheme>'; nothing goes back to the inherited one. Text
// boxes get a hanging indent so the text clears the bullet.
function setBullet (pp, bu, indent) {
  pp.kids = pp.kids.filter((x) => !BU_KIDS.includes(x.name))
  const hang = pp.attrs.indent === '-342900'
  if (!bu || bu === 'none') {
    if (bu === 'none') insertOrdered(pp, el('a:buNone'), PPR_KIDS)
    if (hang) { delete pp.attrs.marL; delete pp.attrs.indent }
    return
  }
  if (bu.startsWith('c:')) {
    insertOrdered(pp, el('a:buFont', { typeface: 'Arial', panose: '020B0604020202020204', pitchFamily: '34', charset: '0' }), PPR_KIDS)
    insertOrdered(pp, el('a:buChar', { char: bu.slice(2) || '•' }), PPR_KIDS)
  } else insertOrdered(pp, el('a:buAutoNum', { type: AUTONUM[bu.slice(2)] ? bu.slice(2) : 'arabicPeriod' }), PPR_KIDS)
  if (indent && pp.attrs.marL === undefined) { pp.attrs.marL = String(342900 + (+(pp.attrs.lvl ?? 0)) * 457200); pp.attrs.indent = '-342900' }
}

// dropAnimsOf removes the effects that target shapes being deleted, so
// the timing never points at a shape that is gone.
function dropAnimsOf (root, ids) {
  const timing = child(root, 'p:timing')
  if (!timing || !ids.size) return
  const hit = (n) => descendants(n, 'p:spTgt').some((t) => ids.has(t.attrs.spid))
  for (const tn of descendants(timing, 'p:cTn').filter((c) => c.attrs.presetClass && hit(c))) {
    const par = descendants(timing, (n) => n.kids?.includes(tn))[0]
    const holder = par && descendants(timing, (n) => n.kids?.includes(par))[0]
    if (holder) remove(holder, par)
  }
  // Steps and clicks left without effects go too.
  for (let again = true; again;) {
    again = false
    for (const par of descendants(timing, 'p:par')) {
      const tn = child(par, 'p:cTn')
      if (!tn || tn.attrs.presetClass || tn.attrs.nodeType === 'tmRoot') continue
      const kids = child(tn, 'p:childTnLst')
      if (kids && !kids.kids.some(isEl)) { const holder = descendants(timing, (n) => n.kids?.includes(par))[0]; if (holder) { remove(holder, par); again = true } }
    }
  }
  const bld = child(timing, 'p:bldLst')
  if (bld) { bld.kids = bld.kids.filter((k) => !ids.has(k.attrs?.spid)); if (!bld.kids.some(isEl)) remove(timing, bld) }
  if (!descendants(timing, 'p:cTn').some((c) => c.attrs.presetClass)) remove(root, timing)
}


export async function editPptxUi (buf, ops, io) {
  let cur = buf
  for (const [k, e] of ops.entries()) {
    const where = `ops[${k}] (${e?.op})`
    if (['add_slide', 'duplicate_slide', 'delete_slide', 'move_slide', 'notes', 'replace', 'set_text'].includes(e?.op)) {
      cur = (await editPptx(cur, [e], io)).buffer
      continue
    }
    const pkg = Package.fromBuffer(cur)
    const deck = new Deck(pkg)
    switch (e?.op) {
      case 'shape_text': {
        const { s, node } = shapeOf(deck, e, where)
        if (node.name !== 'p:sp') throw new Error(`${where}: shape ${e.shape} holds no text of its own`)
        writeShapeText(node, Array.isArray(e.paras) ? e.paras : [])
        pkg.touch(s.part)
        break
      }
      case 'xfrm': {
        const { s, node } = shapeOf(deck, e, where)
        const box = ['x', 'y', 'cx', 'cy'].map((n) => Math.round(+e[n]))
        if (box.some((x) => !Number.isFinite(x)) || box[2] <= 0 || box[3] <= 0) throw new Error(`${where}: x, y, cx and cy must be numbers (EMU)`)
        setBox(node, { x: box[0], y: box[1], cx: box[2], cy: box[3], rot: e.rot === undefined || e.rot === null ? undefined : +e.rot })
        pkg.touch(s.part)
        break
      }
      case 'add_textbox': {
        const s = hostOf(deck, e, where)
        const tree = xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
        const id = deck.nextShapeId(tree)
        const box = { x: Math.round(+e.x || deck.cx * 0.3), y: Math.round(+e.y || deck.cy * 0.4), cx: Math.round(+e.cx || deck.cx * 0.4), cy: Math.round(+e.cy || 600000) }
        tree.kids.push(el('p:sp', {}, [
          el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id), name: `文本框 ${id}` }), el('p:cNvSpPr', { txBox: '1' }), el('p:nvPr')]),
          el('p:spPr', {}, [el('a:xfrm', {}, [el('a:off', { x: String(box.x), y: String(box.y) }), el('a:ext', { cx: String(box.cx), cy: String(box.cy) })]), el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]), el('a:noFill')]),
          el('p:txBody', {}, [el('a:bodyPr', { wrap: 'square', rtlCol: '0' }, [el('a:spAutoFit')]), el('a:lstStyle'),
            el('a:p', {}, [e.text ? el('a:r', {}, [el('a:rPr', { lang: 'zh-CN', altLang: 'en-US', sz: String(e.sz ?? 1800) }), el('a:t', {}, [text(String(e.text))])]) : null, el('a:endParaRPr', { lang: 'zh-CN', altLang: 'en-US', sz: String(e.sz ?? 1800) })])])
        ]))
        pkg.touch(s.part)
        break
      }
      case 'delete_shape': {
        const { s, tree, node } = shapeOf(deck, e, where)
        dropAnimsOf(pkg.xml(s.part).root, new Set(descendants({ kids: [node] }, 'p:cNvPr').map((n) => n.attrs.id)))
        const parent = node === undefined ? undefined : [tree, ...descendants(tree, 'p:grpSp')].find((g) => g.kids.includes(node)) ?? descendants(tree, (n) => n.kids?.includes(node))[0]
        if (parent) remove(parent, node)
        pkg.touch(s.part)
        break
      }
      case 'new_slide': {
        const layout = deck.layouts().find((l) => l.part === e.layout) ?? deck.pickLayout('content')
        const part = pkg.uniqueName('ppt/slides/slide#.xml')
        pkg.setXml(part, doc(el('p:sld', NSA, [el('p:cSld', {}, [el('p:spTree', {}, grpRoot())]), el('p:clrMapOvr', {}, [el('a:masterClrMapping')])])))
        pkg.setOverride(part, CT.slide)
        pkg.addRel(part, REL('slideLayout'), layout.part)
        const tree = xpath(pkg.xml(part).root, 'p:cSld', 'p:spTree')
        let id = 2
        for (const p of placeholders(pkg.xml(layout.part).root)) {
          if (['dt', 'ftr', 'sldNum', 'hdr'].includes(p.type)) continue
          const ph = {}
          for (const key of ['type', 'orient', 'sz', 'idx']) if (p.ph.attrs[key] !== undefined) ph[key] = p.ph.attrs[key]
          const name = (child(xpath(p.sp, 'p:nvSpPr'), 'p:cNvPr')?.attrs.name ?? 'Placeholder').replace(/\s*\d+$/, '') + ' ' + id
          tree.kids.push(el('p:sp', {}, [
            el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id++), name }), el('p:cNvSpPr', {}, [el('a:spLocks', { noGrp: '1' })]), el('p:nvPr', {}, [el('p:ph', ph)])]),
            el('p:spPr'),
            el('p:txBody', {}, [el('a:bodyPr'), el('a:lstStyle'), el('a:p', {}, [el('a:endParaRPr', { lang: 'zh-CN', altLang: 'en-US' })])])
          ]))
        }
        insertSlide(deck, part, e.at)
        break
      }
      case 'hide_slide': {
        const s = deck.slide(e.slide, where)
        const root = pkg.xml(s.part).root
        if (e.hidden) root.attrs.show = '0'; else delete root.attrs.show
        pkg.touch(s.part)
        break
      }
      case 'add_shape': {
        const s = hostOf(deck, e, where)
        const tree = xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
        const prst = ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'rightArrow', 'leftArrow', 'pentagon', 'star5'].includes(e.prst) ? e.prst : 'rect'
        const id = deck.nextShapeId(tree)
        const box = boxOf(deck, e, 0.35, 0.35, 0.3, 0.3)
        // Without a colour of its own a new shape follows the theme, as in
        // PowerPoint: accent 1 with a darker outline.
        const fill = e.fill ? solid(hex(e.fill)) : el('a:solidFill', {}, [el('a:schemeClr', { val: 'accent1' })])
        const line = e.line ? solid(hex(e.line)) : e.fill ? solid(darker(hex(e.fill))) : el('a:solidFill', {}, [el('a:schemeClr', { val: 'accent1' }, [el('a:shade', { val: '50000' })])])
        tree.kids.push(el('p:sp', {}, [
          el('p:nvSpPr', {}, [el('p:cNvPr', { id: String(id), name: `形状 ${id}` }), el('p:cNvSpPr'), el('p:nvPr')]),
          el('p:spPr', {}, [xfrm(box.x, box.y, box.cx, box.cy), el('a:prstGeom', { prst }, [el('a:avLst')]), fill, el('a:ln', { w: '12700' }, [line])]),
          el('p:txBody', {}, [el('a:bodyPr', { rtlCol: '0', anchor: 'ctr' }), el('a:lstStyle'), el('a:p', {}, [el('a:pPr', { algn: 'ctr' }), el('a:endParaRPr', { lang: 'zh-CN', altLang: 'en-US', sz: '1800' }, [solid('FFFFFF')])])])
        ]))
        pkg.touch(s.part)
        break
      }
      case 'add_image': {
        const s = hostOf(deck, e, where)
        const tree = xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
        tree.kids.push(await pictureShape(deck, s.part, deck.nextShapeId(tree), boxOf(deck, e, 0.2, 0.2, 0.6, 0.6), String(e.src ?? ''), io))
        pkg.touch(s.part)
        break
      }
      case 'add_table': {
        const s = hostOf(deck, e, where)
        const tree = xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
        const rows = Math.max(1, Math.min(30, +e.rows || 3)); const cols = Math.max(1, Math.min(15, +e.cols || 3))
        const grid = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''))
        tree.kids.push(tableFrame(deck.nextShapeId(tree), boxOf(deck, e, 0.1, 0.25, 0.8, 0.5), { rows: grid, header: true }, themeAccent(deck)))
        pkg.touch(s.part)
        break
      }
      case 'table_cell': {
        const { s, node } = shapeOf(deck, e, where)
        const tbl = descendants(node, 'a:tbl')[0]
        const tc = tbl && children(children(tbl, 'a:tr')[e.row - 1] ?? el('x'), 'a:tc')[e.col - 1]
        if (!tc) throw new Error(`${where}: shape ${e.shape} has no table cell at row ${e.row}, column ${e.col}`)
        writeCellText(tc, String(e.text ?? ''))
        pkg.touch(s.part)
        break
      }
      case 'shape_style': {
        const { s, node } = shapeOf(deck, e, where)
        if (!['p:sp', 'p:pic', 'p:cxnSp'].includes(node.name)) throw new Error(`${where}: only shapes, text boxes and pictures take a fill, outline or shadow`)
        if (node.name !== 'p:sp' && (e.fill !== undefined || e.gradient)) throw new Error(`${where}: pictures take an outline or shadow, not a fill`)
        const pr = spPrOf(node)
        const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']
        if (e.fill !== undefined || e.gradient) {
          pr.kids = pr.kids.filter((x) => !FILLS.includes(x.name))
          if (Array.isArray(e.gradient)) {
            const [c1, c2, ang] = e.gradient
            insertOrdered(pr, el('a:gradFill', { rotWithShape: '1' }, [el('a:gsLst', {}, [el('a:gs', { pos: '0' }, [el('a:srgbClr', { val: hex(c1) })]), el('a:gs', { pos: '100000' }, [el('a:srgbClr', { val: hex(c2) })])]), el('a:lin', { ang: String(Math.round((+ang || 90) * 60000)), scaled: '0' })]), SPPR_KIDS)
          } else insertOrdered(pr, e.fill ? solid(hex(e.fill)) : el('a:noFill'), SPPR_KIDS)
        }
        if (e.line !== undefined) {
          let ln = child(pr, 'a:ln')
          if (!ln) { ln = el('a:ln', { w: '12700' }); insertOrdered(pr, ln, SPPR_KIDS) }
          ln.kids = ln.kids.filter((x) => !FILLS.includes(x.name))
          ln.kids.unshift(e.line ? solid(hex(e.line)) : el('a:noFill'))
        }
        if (e.lineW !== undefined && +e.lineW > 0) {
          let ln = child(pr, 'a:ln')
          if (!ln) { ln = el('a:ln', {}, [solid('000000')]); insertOrdered(pr, ln, SPPR_KIDS) }
          ln.attrs.w = String(Math.round(+e.lineW * 12700))
        }
        if (e.shadow !== undefined) {
          pr.kids = pr.kids.filter((x) => x.name !== 'a:effectLst' && x.name !== 'a:effectDag')
          if (e.shadow) insertOrdered(pr, el('a:effectLst', {}, [el('a:outerShdw', { blurRad: '63500', dist: '38100', dir: '2700000', algn: 'tl', rotWithShape: '0' }, [el('a:srgbClr', { val: '000000' }, [el('a:alpha', { val: '40000' })])])]), SPPR_KIDS)
        }
        pkg.touch(s.part)
        break
      }
      case 'group': {
        const s = hostOf(deck, e, where)
        const tree = xpath(pkg.xml(s.part).root, 'p:cSld', 'p:spTree')
        const all = shapes(tree); const groups = descendants(tree, 'p:grpSp')
        const nodes = (Array.isArray(e.items) ? e.items : []).map((it) => it.group ? groups[+it.group - 1] : all[+it.shape - 1])
        if (nodes.some((n) => !n)) throw new Error(`${where}: a shape to group is missing; reload the presentation`)
        groupShapes(deck, tree, nodes, where)
        pkg.touch(s.part)
        break
      }
      case 'ungroup': {
        const { s, tree, node } = shapeOf(deck, e, where)
        if (node.name !== 'p:grpSp') throw new Error(`${where}: that is not a group`)
        ungroupShape(tree, node, where)
        pkg.touch(s.part)
        break
      }
      case 'crop': {
        const { s, node } = shapeOf(deck, e, where)
        setCrop(node, e, where)
        pkg.touch(s.part)
        break
      }
      case 'body': {
        const { s, node } = shapeOf(deck, e, where)
        if (node.name !== 'p:sp') throw new Error(`${where}: only shapes and text boxes hold text`)
        setBody(node, e)
        pkg.touch(s.part)
        break
      }
      case 'text_effect': {
        const { s, node } = shapeOf(deck, e, where)
        if (node.name !== 'p:sp') throw new Error(`${where}: only shapes and text boxes hold text`)
        setTextFx(node, String(e.preset ?? 'none'), hex(e.color || themeInfo(deck).colors.accent1 || themeAccent(deck)))
        pkg.touch(s.part)
        break
      }
      case 'slide_layout': {
        const s = deck.slide(e.slide, where)
        const layout = deck.layouts().find((l) => l.part === e.layout)
        if (!layout) throw new Error(`${where}: there is no layout ${e.layout}`)
        changeLayout(deck, s, layout, where)
        break
      }
      case 'theme':
        setTheme(deck, e, where)
        break
      case 'background': {
        if (e.all) {
          for (const m of themeParts(deck).masters) putBg(deck, m, e.reset ? el('p:bg', {}, [el('p:bgRef', { idx: '1001' }, [el('a:schemeClr', { val: 'bg1' })])]) : await bgNode(deck, m, e, io))
          for (const sl of deck.slides()) if (sl.part && pkg.has(sl.part) && child(child(pkg.xml(sl.part).root, 'p:cSld'), 'p:bg')) putBg(deck, sl.part, undefined)
        } else {
          const s = hostOf(deck, e, where)
          const isMaster = pkg.contentType(s.part) === CT.master
          const node = e.reset ? (isMaster ? el('p:bg', {}, [el('p:bgRef', { idx: '1001' }, [el('a:schemeClr', { val: 'bg1' })])]) : undefined) : await bgNode(deck, s.part, e, io)
          if (!e.reset && !node) throw new Error(`${where}: give fill, gradient or image`)
          putBg(deck, s.part, node)
        }
        break
      }
      case 'transition': {
        const list = e.all ? deck.slides() : [deck.slide(e.slide, where)]
        for (const sl of list) {
          if (!sl.part || !pkg.has(sl.part)) continue
          setTransition(pkg.xml(sl.part).root, e)
          pkg.touch(sl.part)
        }
        break
      }
      case 'animations': {
        const s = deck.slide(e.slide, where)
        const root = pkg.xml(s.part).root
        const spids = new Set(descendants(xpath(root, 'p:cSld', 'p:spTree'), 'p:cNvPr').map((n) => n.attrs.id))
        setAnimations(root, Array.isArray(e.list) ? e.list : [], spids, where)
        pkg.touch(s.part)
        break
      }
      case 'order': {
        const { s, tree, node } = shapeOf(deck, e, where)
        const parent = [tree, ...descendants(tree, 'p:grpSp')].find((g) => g.kids.includes(node))
        const sibs = parent.kids.filter((x) => isEl(x) && !['p:nvGrpSpPr', 'p:grpSpPr', 'p:extLst'].includes(x.name))
        const i = sibs.indexOf(node)
        const j = { front: sibs.length - 1, back: 0, forward: Math.min(sibs.length - 1, i + 1), backward: Math.max(0, i - 1) }[e.to]
        if (j === undefined) throw new Error(`${where}: to must be front, back, forward or backward`)
        if (i !== j) {
          remove(parent, node)
          const rest = sibs.filter((x) => x !== node)
          const anchor = rest[j]
          if (anchor) parent.kids.splice(parent.kids.indexOf(anchor), 0, node)
          else { const ext = child(parent, 'p:extLst'); parent.kids.splice(ext ? parent.kids.indexOf(ext) : parent.kids.length, 0, node) }
          pkg.touch(s.part)
        }
        break
      }
      case 'copy_shape': {
        const from = shapeOf(deck, { slide: e.from, part: e.fromPart, shape: e.shape, group: e.group }, where)
        const to = hostOf(deck, e.slide !== undefined || e.part !== undefined ? { slide: e.slide, part: e.part } : { slide: e.from, part: e.fromPart }, where)
        const tree = xpath(pkg.xml(to.part).root, 'p:cSld', 'p:spTree')
        const node = clone(from.node)
        let id = deck.nextShapeId(tree)
        for (const c of descendants(node, (n) => n.name === 'p:cNvPr')) c.attrs.id = String(id++)
        // Pictures and links copied to another slide need relationships there.
        if (to.part !== from.s.part) {
          for (const n of descendants(node, (x) => x.attrs && (x.attrs['r:embed'] || x.attrs['r:id'] || x.attrs['r:link']))) {
            for (const key of ['r:embed', 'r:id', 'r:link']) {
              const rel = n.attrs[key] && pkg.rel(from.s.part, n.attrs[key])
              if (rel) n.attrs[key] = rel.external ? pkg.addRel(to.part, rel.type, rel.target, true) : pkg.addRel(to.part, rel.type, rel.part)
            }
          }
        }
        // A copied placeholder becomes an ordinary shape where the original sat.
        const ph = descendants(node, 'p:ph')[0]
        if (ph) for (const g of descendants(node, (x) => x.kids?.includes(ph))) remove(g, ph)
        if (e.box) {
          const b = e.box
          const dx = +e.dx || 0; const dy = +e.dy || 0
          if (node.name === 'p:graphicFrame') {
            const x = child(node, 'p:xfrm')
            if (x) { child(x, 'a:off').attrs.x = String(Math.round(+b.x + dx)); child(x, 'a:off').attrs.y = String(Math.round(+b.y + dy)) }
          } else {
            const prName = node.name === 'p:grpSp' ? 'p:grpSpPr' : 'p:spPr'
            let pr = child(node, prName)
            if (!pr) { pr = el(prName); node.kids.splice(1, 0, pr) }
            const old = child(pr, 'a:xfrm')
            const nx = xfrm(+b.x + dx, +b.y + dy, +b.cx, +b.cy)
            if (old) { child(old, 'a:off').attrs.x = child(nx, 'a:off').attrs.x; child(old, 'a:off').attrs.y = child(nx, 'a:off').attrs.y; if (!child(old, 'a:ext')) old.kids.push(child(nx, 'a:ext')) } else pr.kids.unshift(nx)
          }
        }
        tree.kids.push(node)
        pkg.touch(to.part)
        break
      }
      default:
        throw new Error(`${where}: unknown op`)
    }
    cur = pkg.toBuffer()
  }
  return { buffer: cur, summary: `${ops.length} change(s)` }
}
