// Word documents: read as numbered blocks, create from Markdown, edit by
// block, render for the preview pane.
import { Package, NS, REL, dataUri } from './pkg.js'
import { el, text, doc, child, children, descendants, path as xpath, textContent, remove, clone, insertOrdered, isEl } from './xml.js'
import { parseMarkdown } from './md.js'
import { imageInfo, EMU_PER_PX, EMU_PER_CM } from './image.js'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const CT_DOC = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
const CT_NUMBERING = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
const CT_SETTINGS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const DRAW_NS = {
  'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
}
const PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange']

const on = (n) => !!n && !['0', 'false', 'off', 'none'].includes(n.attrs['w:val'])

// ------------------------------------------------------------- template

function newPackage () {
  const p = new Package([])
  p.setXml('[Content_Types].xml', doc(el('Types', { xmlns: NS.ct }, [
    el('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }),
    el('Default', { Extension: 'xml', ContentType: 'application/xml' }),
    el('Override', { PartName: '/word/document.xml', ContentType: CT_DOC }),
    el('Override', { PartName: '/word/styles.xml', ContentType: CT_STYLES }),
    el('Override', { PartName: '/word/settings.xml', ContentType: CT_SETTINGS }),
    el('Override', { PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }),
    el('Override', { PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' })
  ])))
  p.setXml('_rels/.rels', doc(el('Relationships', { xmlns: NS.rel }, [
    el('Relationship', { Id: 'rId1', Type: NS.officeDocument, Target: 'word/document.xml' }),
    el('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: 'docProps/core.xml' }),
    el('Relationship', { Id: 'rId3', Type: REL('extended-properties'), Target: 'docProps/app.xml' })
  ])))
  p.setXml('word/document.xml', doc(el('w:document', { 'xmlns:w': W, 'xmlns:r': NS.r, ...DRAW_NS }, [
    el('w:body', {}, [el('w:sectPr', {}, [
      el('w:pgSz', { 'w:w': '11906', 'w:h': '16838' }),
      el('w:pgMar', { 'w:top': '1440', 'w:right': '1440', 'w:bottom': '1440', 'w:left': '1440', 'w:header': '708', 'w:footer': '708', 'w:gutter': '0' })
    ])])
  ])))
  p.setXml('word/_rels/document.xml.rels', doc(el('Relationships', { xmlns: NS.rel }, [
    el('Relationship', { Id: 'rId1', Type: REL('styles'), Target: 'styles.xml' }),
    el('Relationship', { Id: 'rId2', Type: REL('settings'), Target: 'settings.xml' })
  ])))
  p.setXml('word/styles.xml', doc(el('w:styles', { 'xmlns:w': W }, [
    el('w:docDefaults', {}, [
      el('w:rPrDefault', {}, [el('w:rPr', {}, [
        el('w:rFonts', { 'w:ascii': 'Calibri', 'w:hAnsi': 'Calibri', 'w:eastAsia': 'Microsoft YaHei', 'w:cs': 'Calibri' }),
        el('w:sz', { 'w:val': '22' }), el('w:szCs', { 'w:val': '22' }),
        el('w:lang', { 'w:val': 'en-US', 'w:eastAsia': 'zh-CN', 'w:bidi': 'ar-SA' })
      ])]),
      el('w:pPrDefault', {}, [el('w:pPr', {}, [el('w:spacing', { 'w:after': '120', 'w:line': '288', 'w:lineRule': 'auto' })])])
    ]),
    el('w:style', { 'w:type': 'paragraph', 'w:default': '1', 'w:styleId': 'Normal' }, [el('w:name', { 'w:val': 'Normal' }), el('w:qFormat')]),
    el('w:style', { 'w:type': 'character', 'w:default': '1', 'w:styleId': 'DefaultParagraphFont' }, [el('w:name', { 'w:val': 'Default Paragraph Font' }), el('w:uiPriority', { 'w:val': '1' }), el('w:semiHidden'), el('w:unhideWhenUsed')]),
    el('w:style', { 'w:type': 'table', 'w:default': '1', 'w:styleId': 'TableNormal' }, [el('w:name', { 'w:val': 'Normal Table' }), el('w:semiHidden'), el('w:unhideWhenUsed'),
      el('w:tblPr', {}, [el('w:tblInd', { 'w:w': '0', 'w:type': 'dxa' }), el('w:tblCellMar', {}, [
        el('w:top', { 'w:w': '0', 'w:type': 'dxa' }), el('w:left', { 'w:w': '108', 'w:type': 'dxa' }), el('w:bottom', { 'w:w': '0', 'w:type': 'dxa' }), el('w:right', { 'w:w': '108', 'w:type': 'dxa' })])])])
  ])))
  p.setXml('word/settings.xml', doc(el('w:settings', { 'xmlns:w': W }, [
    el('w:defaultTabStop', { 'w:val': '720' }),
    el('w:characterSpacingControl', { 'w:val': 'doNotCompress' }),
    el('w:compat', {}, [el('w:compatSetting', { 'w:name': 'compatibilityMode', 'w:uri': 'http://schemas.microsoft.com/office/word', 'w:val': '15' })])
  ])))
  p.setXml('docProps/core.xml', coreProps())
  p.setXml('docProps/app.xml', doc(el('Properties', { xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' }, [el('Application', {}, [text('Quire')])])))
  return p
}

export function coreProps () {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  return doc(el('cp:coreProperties', {
    'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
    'xmlns:dcterms': 'http://purl.org/dc/terms/',
    'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance'
  }, [
    el('dc:creator', {}, [text('Quire')]),
    el('dcterms:created', { 'xsi:type': 'dcterms:W3CDTF' }, [text(now)]),
    el('dcterms:modified', { 'xsi:type': 'dcterms:W3CDTF' }, [text(now)])
  ]))
}

// ------------------------------------------------------------- styles

const rPr = (...kids) => el('w:rPr', {}, kids)
const pPr = (...kids) => el('w:pPr', {}, kids)
const v = (name, val) => el(name, { 'w:val': String(val) })

// The styles Markdown maps to, by Word's internal (English, lower-case)
// style name; ids are looked up by name because localized Word writes
// other ids (Chinese Word: heading 1 is "1").
function styleDefs (normal) {
  const heading = (n, sz, extra = []) => ({
    name: `heading ${n}`,
    id: `Heading${n}`,
    node: (id) => el('w:style', { 'w:type': 'paragraph', 'w:styleId': id }, [
      v('w:name', `heading ${n}`), v('w:basedOn', normal), v('w:next', normal), v('w:uiPriority', 9), el('w:qFormat'),
      pPr(el('w:keepNext'), el('w:keepLines'), el('w:spacing', { 'w:before': n === 1 ? '360' : '240', 'w:after': '120' }), v('w:outlineLvl', n - 1)),
      rPr(el('w:b'), el('w:bCs'), v('w:color', n <= 2 ? '1F3864' : '2F5496'), v('w:sz', sz), v('w:szCs', sz), ...extra)
    ])
  })
  return {
    h1: heading(1, 32), h2: heading(2, 28), h3: heading(3, 24), h4: heading(4, 22, [el('w:i')]), h5: heading(5, 22), h6: heading(6, 20),
    title: { name: 'title', id: 'Title', node: (id) => el('w:style', { 'w:type': 'paragraph', 'w:styleId': id }, [v('w:name', 'Title'), v('w:basedOn', normal), v('w:next', normal), el('w:qFormat'), pPr(el('w:spacing', { 'w:after': '240' })), rPr(v('w:sz', 52), v('w:szCs', 52))]) },
    quote: { name: 'quote', id: 'Quote', node: (id) => el('w:style', { 'w:type': 'paragraph', 'w:styleId': id }, [v('w:name', 'Quote'), v('w:basedOn', normal), v('w:next', normal), el('w:qFormat'), pPr(el('w:pBdr', {}, [el('w:left', { 'w:val': 'single', 'w:sz': '18', 'w:space': '8', 'w:color': 'BFBFBF' })]), el('w:ind', { 'w:left': '567' })), rPr(el('w:i'), v('w:color', '595959'))]) },
    code: { name: 'quire code', id: 'QuireCode', node: (id) => el('w:style', { 'w:type': 'paragraph', 'w:customStyle': '1', 'w:styleId': id }, [v('w:name', 'Quire Code'), v('w:basedOn', normal), pPr(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'F2F2F2' }), el('w:spacing', { 'w:after': '120', 'w:line': '240', 'w:lineRule': 'auto' })), rPr(el('w:rFonts', { 'w:ascii': 'Consolas', 'w:hAnsi': 'Consolas', 'w:cs': 'Consolas' }), v('w:sz', 19), v('w:szCs', 19))]) },
    list: { name: 'list paragraph', id: 'ListParagraph', node: (id) => el('w:style', { 'w:type': 'paragraph', 'w:styleId': id }, [v('w:name', 'List Paragraph'), v('w:basedOn', normal), v('w:uiPriority', 34), el('w:qFormat'), pPr(el('w:spacing', { 'w:after': '60' }), el('w:ind', { 'w:left': '720' }), el('w:contextualSpacing'))]) },
    link: { name: 'hyperlink', id: 'Hyperlink', node: (id) => el('w:style', { 'w:type': 'character', 'w:styleId': id }, [v('w:name', 'Hyperlink'), v('w:uiPriority', 99), el('w:unhideWhenUsed'), rPr(v('w:color', '0563C1'), v('w:u', 'single'))]) },
    table: {
      name: 'quire table',
      id: 'QuireTable',
      node: (id) => {
        const b = (side) => el(side, { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'BFBFBF' })
        return el('w:style', { 'w:type': 'table', 'w:customStyle': '1', 'w:styleId': id }, [v('w:name', 'Quire Table'), v('w:basedOn', 'TableNormal'),
          pPr(el('w:spacing', { 'w:before': '40', 'w:after': '40' })),
          el('w:tblPr', {}, [el('w:tblBorders', {}, ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV'].map(b)),
            el('w:tblCellMar', {}, [el('w:left', { 'w:w': '108', 'w:type': 'dxa' }), el('w:right', { 'w:w': '108', 'w:type': 'dxa' })])]),
          el('w:tblStylePr', { 'w:type': 'firstRow' }, [rPr(el('w:b'), el('w:bCs')), el('w:tcPr', {}, [el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'D9E2F3' })])])])
      }
    }
  }
}

// ensureStyles returns logical name → style id, adding what the document
// lacks.
function ensureStyles (pkg, main) {
  if (pkg._quireStyles) return pkg._quireStyles
  let name = stylesPart(pkg, main)
  if (!name) {
    name = 'word/styles.xml'
    const t = newPackage()
    pkg.setXml(name, t.xml('word/styles.xml'))
    pkg.setOverride(name, CT_STYLES)
    pkg.addRel(main, REL('styles'), name)
  }
  const d = pkg.xml(name)
  const byName = new Map()
  const ids = new Set()
  let normal = 'Normal'
  for (const s of children(d.root, 'w:style')) {
    const nm = child(s, 'w:name')?.attrs['w:val']?.toLowerCase()
    ids.add(s.attrs['w:styleId'])
    if (nm) byName.set(`${s.attrs['w:type']}:${nm}`, s.attrs['w:styleId'])
    if (s.attrs['w:type'] === 'paragraph' && on({ attrs: { 'w:val': s.attrs['w:default'] } }) && s.attrs['w:default']) normal = s.attrs['w:styleId']
  }
  if (!ids.has('TableNormal') && !byName.has('table:normal table')) {
    d.root.kids.push(newPackage().xml('word/styles.xml').root.kids.find((k) => k.attrs?.['w:styleId'] === 'TableNormal'))
  }
  const out = { normal }
  const defs = styleDefs(normal)
  for (const [key, def] of Object.entries(defs)) {
    const type = key === 'link' ? 'character' : key === 'table' ? 'table' : 'paragraph'
    let id = byName.get(`${type}:${def.name}`)
    if (!id) {
      id = def.id
      let n = 2
      while (ids.has(id)) id = def.id + n++
      ids.add(id)
      d.root.kids.push(def.node(id))
      pkg.touch(name)
    }
    out[key] = id
  }
  pkg._quireStyles = out
  return out
}

function stylesPart (pkg, main) {
  return pkg.relsOfType(main, 'styles')[0]?.part
}

// ------------------------------------------------------------- numbering

const BULLETS = ['•', '◦', '▪', '•', '◦', '▪', '•', '◦', '▪']
// List styles the editor offers: [numFmt, lvlText] for the first levels.
export const LIST_STYLES = {
  disc: { bullet: true, levels: [['bullet', '•']] },
  circle: { bullet: true, levels: [['bullet', '○']] },
  square: { bullet: true, levels: [['bullet', '■']] },
  diamond: { bullet: true, levels: [['bullet', '◆']] },
  check: { bullet: true, levels: [['bullet', '✓']] },
  arrow: { bullet: true, levels: [['bullet', '➢']] },
  decimal: { levels: [['decimal', '%1.']] },
  paren: { levels: [['decimal', '%1)']] },
  fullparen: { levels: [['decimal', '(%1)']] },
  cn: { levels: [['chineseCounting', '%1、']] },
  cnparen: { levels: [['chineseCounting', '(%1)']] },
  lowerLetter: { levels: [['lowerLetter', '%1.']] },
  upperLetter: { levels: [['upperLetter', '%1.']] },
  lowerRoman: { levels: [['lowerRoman', '%1.']] },
  upperRoman: { levels: [['upperRoman', '%1.']] },
  circled: { levels: [['decimalEnclosedCircle', '%1']] },
  multi: { multi: true, levels: Array.from({ length: 9 }, (_, i) => ['decimal', Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + (i ? '' : '.')]) },
  outline: { multi: true, levels: [['chineseCounting', '%1、'], ['chineseCounting', '(%2)'], ['decimal', '%3.'], ['decimal', '(%4)'], ['decimalEnclosedCircle', '%5'], ['lowerLetter', '%6.'], ['lowerLetter', '(%7)'], ['lowerRoman', '%8.'], ['lowerRoman', '(%9)']] }
}
const ORDERED = ['decimal', 'lowerLetter', 'lowerRoman', 'decimal', 'lowerLetter', 'lowerRoman', 'decimal', 'lowerLetter', 'lowerRoman']

function numberingDoc (pkg, main, create) {
  let name = pkg.relsOfType(main, 'numbering')[0]?.part
  if (!name && create) {
    name = 'word/numbering.xml'
    pkg.setXml(name, doc(el('w:numbering', { 'xmlns:w': W })))
    pkg.setOverride(name, CT_NUMBERING)
    pkg.addRel(main, REL('numbering'), name)
  }
  return name ? { name, d: pkg.xml(name) } : undefined
}

// numIds keeps one bullet list definition and one numbered definition per
// Markdown list.
class Numbering {
  constructor (pkg, main) { this.pkg = pkg; this.main = main; this.ordered = new Map() }

  get doc () {
    if (!this._doc) this._doc = numberingDoc(this.pkg, this.main, true)
    return this._doc
  }

  abstract (kind) {
    const { d, name } = this.doc
    const used = children(d.root, 'w:abstractNum').map((a) => +a.attrs['w:abstractNumId'])
    const id = String(Math.max(-1, ...used) + 1)
    const lvls = []
    for (let i = 0; i < 9; i++) {
      lvls.push(el('w:lvl', { 'w:ilvl': String(i) }, [
        v('w:start', 1),
        v('w:numFmt', kind === 'bullet' ? 'bullet' : ORDERED[i]),
        v('w:lvlText', kind === 'bullet' ? BULLETS[i] : `%${i + 1}.`),
        v('w:lvlJc', 'left'),
        pPr(el('w:ind', { 'w:left': String(720 + 360 * i), 'w:hanging': '360' })),
        ...(kind === 'bullet' ? [rPr(el('w:rFonts', { 'w:ascii': 'Arial', 'w:hAnsi': 'Arial', 'w:hint': 'default' }))] : [])
      ]))
    }
    const node = el('w:abstractNum', { 'w:abstractNumId': id }, [v('w:multiLevelType', 'hybridMultilevel'), ...lvls])
    const at = d.root.kids.findIndex((k) => k.name === 'w:num' || k.name === 'w:numIdMacAtCleanup')
    if (at < 0) d.root.kids.push(node)
    else d.root.kids.splice(at, 0, node)
    this.pkg.touch(name)
    return id
  }

  num (abstractId, restart) {
    const { d, name } = this.doc
    const used = children(d.root, 'w:num').map((a) => +a.attrs['w:numId'])
    const id = String(Math.max(0, ...used) + 1)
    const kids = [v('w:abstractNumId', abstractId)]
    if (restart) kids.push(el('w:lvlOverride', { 'w:ilvl': '0' }, [v('w:startOverride', 1)]))
    const node = el('w:num', { 'w:numId': id }, kids)
    const at = d.root.kids.findIndex((k) => k.name === 'w:numIdMacAtCleanup')
    if (at < 0) d.root.kids.push(node)
    else d.root.kids.splice(at, 0, node)
    this.pkg.touch(name)
    return id
  }

  bullet () {
    if (!this.bulletId) this.bulletId = this.num(this.abstract('bullet'), false)
    return this.bulletId
  }

  // styled makes a list in one of LIST_STYLES; each call starts a list.
  styled (key) {
    const st = LIST_STYLES[key]
    if (!st) throw new Error(`unknown list style ${key}`)
    const { d, name } = this.doc
    const used = children(d.root, 'w:abstractNum').map((a) => +a.attrs['w:abstractNumId'])
    const id = String(Math.max(-1, ...used) + 1)
    const lvls = []
    for (let i = 0; i < 9; i++) {
      const [fmt, txt] = st.levels[i] ?? (st.bullet ? ['bullet', BULLETS[i]] : [ORDERED[i], `%${i + 1}.`])
      lvls.push(el('w:lvl', { 'w:ilvl': String(i) }, [
        v('w:start', 1),
        v('w:numFmt', fmt),
        v('w:lvlText', txt),
        v('w:lvlJc', 'left'),
        pPr(el('w:ind', { 'w:left': String(st.multi ? 425 + 567 * i : 720 + 360 * i), 'w:hanging': String(st.multi ? 425 + 142 * Math.min(i, 3) : 360) }))
      ]))
    }
    const node = el('w:abstractNum', { 'w:abstractNumId': id }, [v('w:multiLevelType', st.multi ? 'multilevel' : 'hybridMultilevel'), v('w:name', 'quire:' + key), ...lvls])
    const at = d.root.kids.findIndex((k) => k.name === 'w:num' || k.name === 'w:numIdMacAtCleanup')
    if (at < 0) d.root.kids.push(node)
    else d.root.kids.splice(at, 0, node)
    this.pkg.touch(name)
    return this.num(id, false)
  }

  numbered (list) {
    if (!this.ordered.has(list)) {
      if (!this.orderedAbstract) this.orderedAbstract = this.abstract('ordered')
      this.ordered.set(list, this.num(this.orderedAbstract, true))
    }
    return this.ordered.get(list)
  }
}

// ------------------------------------------------------------- Markdown → WordprocessingML

// Writer turns Markdown blocks into body elements of one package.
class Writer {
  constructor (pkg, io) {
    this.pkg = pkg
    this.io = io
    this.main = pkg.mainPart()
    this.styles = ensureStyles(pkg, this.main)
    this.numbering = new Numbering(pkg, this.main)
    const d = pkg.xml(this.main)
    for (const [k, val] of Object.entries(DRAW_NS)) {
      if (!d.root.attrs[k]) { d.root.attrs[k] = val; pkg.touch(this.main) }
    }
    if (!d.root.attrs['xmlns:r']) d.root.attrs['xmlns:r'] = NS.r
    this.docPr = Math.max(0, ...descendants(d.root, 'wp:docPr').map((n) => +n.attrs.id || 0))
    this.listOffset = Date.now() % 100000 // keeps Markdown list ids from colliding across edits
  }

  async blocks (markdown) {
    const out = []
    for (const b of parseMarkdown(markdown)) out.push(...await this.block(b))
    return out
  }

  async block (b) {
    const s = this.styles
    switch (b.type) {
      case 'heading':
        return [this.para(s['h' + b.level], await this.runs(b.runs))]
      case 'paragraph':
        return [this.para(null, await this.runs(b.runs))]
      case 'quote':
        return [this.para(s.quote, await this.runs(b.runs))]
      case 'item': {
        const numId = b.ordered ? this.numbering.numbered(this.listOffset + b.list) : this.numbering.bullet()
        const p = this.para(s.list, await this.runs(b.runs))
        insertOrdered(child(p, 'w:pPr'), el('w:numPr', {}, [v('w:ilvl', b.level), v('w:numId', numId)]), PPR_ORDER)
        return [p]
      }
      case 'code': {
        const runs = []
        b.text.split('\n').forEach((line, i) => {
          if (i) runs.push(el('w:r', {}, [el('w:br')]))
          if (line) runs.push(el('w:r', {}, [t(line)]))
        })
        return [this.para(s.code, runs)]
      }
      case 'pagebreak':
        return [el('w:p', {}, [el('w:r', {}, [el('w:br', { 'w:type': 'page' })])])]
      case 'rule':
        return [el('w:p', {}, [pPr(el('w:pBdr', {}, [el('w:bottom', { 'w:val': 'single', 'w:sz': '6', 'w:space': '1', 'w:color': 'BFBFBF' })]))])]
      case 'image':
        return [this.para(null, [await this.picture(b.src, b.alt)], 'center')]
      case 'table':
        return [await this.table(b)]
    }
    return []
  }

  para (style, runs, jc) {
    const ppr = el('w:pPr')
    if (style) insertOrdered(ppr, v('w:pStyle', style), PPR_ORDER)
    if (jc) insertOrdered(ppr, v('w:jc', jc), PPR_ORDER)
    return el('w:p', {}, [ppr.kids.length ? ppr : null, ...runs])
  }

  async runs (runs, base = []) {
    const out = []
    for (const r of runs) {
      if (r.br) { out.push(el('w:r', {}, [el('w:br')])); continue }
      if (r.image) { out.push(await this.picture(r.image, r.alt)); continue }
      const props = [...base.map(clone)]
      if (r.link) props.push(v('w:rStyle', this.styles.link))
      if (r.code) props.push(el('w:rFonts', { 'w:ascii': 'Consolas', 'w:hAnsi': 'Consolas', 'w:cs': 'Consolas' }))
      if (r.b) props.push(el('w:b'), el('w:bCs'))
      if (r.i) props.push(el('w:i'), el('w:iCs'))
      if (r.s) props.push(el('w:strike'))
      if (r.code) props.push(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'F2F2F2' }))
      const run = el('w:r', {}, [props.length ? el('w:rPr', {}, sortRPr(props)) : null, ...textNodes(r.text)])
      if (r.link) {
        if (r.link.startsWith('#')) out.push(el('w:hyperlink', { 'w:anchor': r.link.slice(1) }, [run]))
        else out.push(el('w:hyperlink', { 'r:id': this.pkg.addRel(this.main, NS.hyperlink, r.link, true), 'w:history': '1' }, [run]))
      } else out.push(run)
    }
    return out
  }

  async picture (src, alt) {
    const data = await this.io.loadImage(src)
    const info = imageInfo(data)
    if (!info) throw new Error(`"${src}" is not a PNG, JPEG or GIF image`)
    const media = this.pkg.addMedia('word/media', info.ext, data)
    const rid = this.pkg.addRel(this.main, NS.image, media)
    let cx = info.width * EMU_PER_PX
    let cy = info.height * EMU_PER_PX
    const max = 16 * EMU_PER_CM
    if (cx > max) { cy = Math.round(cy * max / cx); cx = max }
    const id = String(++this.docPr)
    return el('w:r', {}, [el('w:drawing', {}, [el('wp:inline', { distT: '0', distB: '0', distL: '0', distR: '0' }, [
      el('wp:extent', { cx: String(cx), cy: String(cy) }),
      el('wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' }),
      el('wp:docPr', { id, name: `Picture ${id}`, descr: alt || '' }),
      el('wp:cNvGraphicFramePr', {}, [el('a:graphicFrameLocks', { noChangeAspect: '1' })]),
      el('a:graphic', {}, [el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture' }, [
        el('pic:pic', {}, [
          el('pic:nvPicPr', {}, [el('pic:cNvPr', { id: '0', name: media.split('/').pop() }), el('pic:cNvPicPr')]),
          el('pic:blipFill', {}, [el('a:blip', { 'r:embed': rid }), el('a:stretch', {}, [el('a:fillRect')])]),
          el('pic:spPr', {}, [el('a:xfrm', {}, [el('a:off', { x: '0', y: '0' }), el('a:ext', { cx: String(cx), cy: String(cy) })]),
            el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')])])
        ])
      ])])
    ])])])
  }

  async table (b) {
    const cols = Math.max(...b.rows.map((r) => r.length))
    const width = Math.floor(9026 / cols)
    const rows = []
    for (let r = 0; r < b.rows.length; r++) {
      const cells = []
      for (let c = 0; c < cols; c++) {
        // The header row is bold and shaded directly too, for renderers
        // that ignore table styles.
        const runs = await this.runs(b.rows[r][c] ?? [], r === 0 ? [el('w:b'), el('w:bCs')] : [])
        const p = this.para(null, runs, b.align?.[c] && b.align[c] !== 'left' ? b.align[c] : undefined)
        const tcPr = el('w:tcPr', {}, [el('w:tcW', { 'w:w': String(width), 'w:type': 'dxa' }), r === 0 ? el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'D9E2F3' }) : null])
        cells.push(el('w:tc', {}, [tcPr, p]))
      }
      rows.push(el('w:tr', {}, [r === 0 ? el('w:trPr', {}, [el('w:tblHeader')]) : null, ...cells]))
    }
    return el('w:tbl', {}, [
      el('w:tblPr', {}, [v('w:tblStyle', this.styles.table), el('w:tblW', { 'w:w': '5000', 'w:type': 'pct' }),
        el('w:tblBorders', {}, ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV'].map((side) => el(side, { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': 'BFBFBF' }))),
        el('w:tblLook', { 'w:val': '04A0', 'w:firstRow': '1', 'w:lastRow': '0', 'w:firstColumn': '0', 'w:lastColumn': '0', 'w:noHBand': '0', 'w:noVBand': '1' })]),
      el('w:tblGrid', {}, Array.from({ length: cols }, () => el('w:gridCol', { 'w:w': String(width) }))),
      ...rows
    ])
  }
}

const RPR_ORDER = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath']
function sortRPr (props) {
  const seen = new Set()
  return props.filter((p) => !seen.has(p.name) && seen.add(p.name)).sort((a, b) => RPR_ORDER.indexOf(a.name) - RPR_ORDER.indexOf(b.name))
}

function t (s) {
  const attrs = /^\s|\s$/.test(s) ? { 'xml:space': 'preserve' } : {}
  return el('w:t', attrs, [text(s)])
}

function textNodes (s) {
  const out = []
  const parts = s.split(/(\t|\n)/)
  for (const p of parts) {
    if (p === '\t') out.push(el('w:tab'))
    else if (p === '\n') out.push(el('w:br'))
    else if (p) out.push(t(p))
  }
  return out
}

// ------------------------------------------------------------- reading

// Model: the document's blocks with refs, plus what reading and rendering
// need to know about styles and numbering.
class Model {
  constructor (pkg) {
    this.pkg = pkg
    this.main = pkg.mainPart()
    if (!this.main || !pkg.has(this.main)) throw new Error('not a Word document: no main document part')
    this.d = pkg.xml(this.main)
    this.body = child(this.d.root, 'w:body')
    if (!this.body) throw new Error('not a Word document: the main part has no body')
    this.styleInfo = new Map()
    const sp = stylesPart(pkg, this.main)
    if (sp && pkg.has(sp)) {
      for (const s of children(pkg.xml(sp).root, 'w:style')) {
        this.styleInfo.set(s.attrs['w:styleId'], {
          name: child(s, 'w:name')?.attrs['w:val'] ?? s.attrs['w:styleId'],
          type: s.attrs['w:type'],
          node: s,
          basedOn: child(s, 'w:basedOn')?.attrs['w:val'],
          default: !!s.attrs['w:default'] && on({ attrs: { 'w:val': s.attrs['w:default'] } })
        })
      }
    }
    this.numFmt = new Map()
    this.numLvls = new Map()
    const num = numberingDoc(pkg, this.main, false)
    if (num) {
      const abs = new Map(children(num.d.root, 'w:abstractNum').map((a) => [a.attrs['w:abstractNumId'], a]))
      for (const n of children(num.d.root, 'w:num')) {
        const a = abs.get(child(n, 'w:abstractNumId')?.attrs['w:val'])
        if (!a) continue
        const lv = children(a, 'w:lvl').map((l) => child(l, 'w:numFmt')?.attrs['w:val'] ?? 'decimal')
        this.numFmt.set(n.attrs['w:numId'], lv)
        const starts = new Map(children(n, 'w:lvlOverride').map((o) => [+o.attrs['w:ilvl'], child(o, 'w:startOverride')?.attrs['w:val']]))
        this.numLvls.set(n.attrs['w:numId'], children(a, 'w:lvl').map((l, i) => ({
          fmt: child(l, 'w:numFmt')?.attrs['w:val'] ?? 'decimal',
          text: child(l, 'w:lvlText')?.attrs['w:val'] ?? '',
          start: +(starts.get(i) ?? child(l, 'w:start')?.attrs['w:val'] ?? 1)
        })))
      }
    }
    this.blocks = []
    this.refs = new Map()
    let np = 0
    let nt = 0
    const walk = (owner) => {
      for (const k of owner.kids) {
        if (k.name === 'w:p') { const b = { ref: 'p' + ++np, kind: 'p', node: k, owner }; this.blocks.push(b); this.refs.set(b.ref, b) } else if (k.name === 'w:tbl') { const b = { ref: 't' + ++nt, kind: 't', node: k, owner }; this.blocks.push(b); this.refs.set(b.ref, b) } else if (k.name === 'w:sdt') { const c = child(k, 'w:sdtContent'); if (c) walk(c) } else if (k.name === 'w:customXml' || k.name === 'w:ins' || k.name === 'w:moveTo') walk(k)
      }
    }
    walk(this.body)
  }

  style (p) {
    const id = xpath(p, 'w:pPr', 'w:pStyle')?.attrs['w:val']
    return id ? this.styleInfo.get(id) : [...this.styleInfo.values()].find((s) => s.type === 'paragraph' && s.default)
  }

  headingLevel (p) {
    const lvl = xpath(p, 'w:pPr', 'w:outlineLvl')?.attrs['w:val']
    const st = this.style(p)
    const name = st?.name?.toLowerCase() ?? ''
    let m
    if ((m = /^heading\s*(\d)$/.exec(name)) || (m = /^标题\s*(\d)$/.exec(name))) return +m[1]
    if (name === 'title' || name === '标题') return 1
    if (lvl !== undefined && +lvl < 9) return +lvl + 1
    const sl = st && xpath(st.node, 'w:pPr', 'w:outlineLvl')?.attrs['w:val']
    if (sl !== undefined && +sl < 9) return +sl + 1
    return 0
  }

  // list returns { level, ordered } for a list paragraph.
  list (p) {
    let np = xpath(p, 'w:pPr', 'w:numPr')
    if (!np) {
      let st = this.style(p)
      for (let k = 0; st && k < 5 && !np; k++) {
        np = xpath(st.node, 'w:pPr', 'w:numPr')
        st = st.basedOn ? this.styleInfo.get(st.basedOn) : undefined
      }
    }
    if (!np) return undefined
    const id = child(np, 'w:numId')?.attrs['w:val']
    if (!id || id === '0') return undefined
    const level = +(child(np, 'w:ilvl')?.attrs['w:val'] ?? 0)
    const fmt = this.numFmt.get(id)?.[level] ?? 'bullet'
    return { level, ordered: fmt !== 'bullet' && fmt !== 'none', numId: id }
  }

  // segments flattens a paragraph into styled pieces of text.
  segments (p) {
    const out = []
    const hyperlinkUrl = (h) => {
      const id = h.attrs['r:id']
      if (id) return this.pkg.rel(this.main, id)?.target
      return h.attrs['w:anchor'] ? '#' + h.attrs['w:anchor'] : undefined
    }
    const visit = (node, link) => {
      for (const k of node.kids) {
        if (!isEl(k)) continue
        switch (k.name) {
          case 'w:r': run(k, link); break
          case 'w:hyperlink': visit(k, hyperlinkUrl(k)); break
          case 'w:ins': case 'w:smartTag': case 'w:fldSimple': case 'w:customXml': case 'w:moveTo': case 'w:dir': case 'w:bdo': visit(k, link); break
          case 'w:sdt': { const c = child(k, 'w:sdtContent'); if (c) visit(c, link); break }
        }
      }
    }
    const run = (r, link) => {
      const pr = child(r, 'w:rPr')
      const style = {
        b: on(child(pr, 'w:b')),
        i: on(child(pr, 'w:i')),
        s: on(child(pr, 'w:strike')) || on(child(pr, 'w:dstrike')),
        u: on(child(pr, 'w:u')),
        color: child(pr, 'w:color')?.attrs['w:val'],
        size: child(pr, 'w:sz')?.attrs['w:val'],
        highlight: child(pr, 'w:highlight')?.attrs['w:val'] ?? child(pr, 'w:shd')?.attrs['w:fill'],
        vert: child(pr, 'w:vertAlign')?.attrs['w:val'],
        link
      }
      for (const k of r.kids) {
        switch (k.name) {
          case 'w:t': out.push({ ...style, text: textContent(k) }); break
          case 'w:tab': case 'w:ptab': out.push({ ...style, text: '\t' }); break
          case 'w:br': case 'w:cr':
            if (k.attrs['w:type'] === 'page') out.push({ pagebreak: true })
            else out.push({ ...style, text: '\n' })
            break
          case 'w:noBreakHyphen': out.push({ ...style, text: '-' }); break
          case 'w:sym': out.push({ ...style, text: k.attrs['w:char'] ? String.fromCharCode(parseInt(k.attrs['w:char'], 16) & 0xfff) : '' }); break
          case 'w:drawing': case 'w:pict': case 'w:object': {
            const blip = descendants(k, (n) => n.name === 'a:blip' || n.name === 'v:imagedata')[0]
            const rid = blip?.attrs['r:embed'] ?? blip?.attrs['r:id']
            const ext = descendants(k, 'wp:extent')[0]
            const pr = descendants(k, 'wp:docPr')[0]
            out.push({
              image: rid ? this.pkg.rel(this.main, rid)?.part : undefined,
              cx: +(ext?.attrs.cx ?? 0),
              cy: +(ext?.attrs.cy ?? 0),
              alt: pr?.attrs.descr || pr?.attrs.name || '',
              textbox: descendants(k, 'w:txbxContent').map((c) => children(c, 'w:p').map((pp) => this.plain(pp)).join(' ')).join(' ')
            })
            break
          }
          case 'w:footnoteReference': out.push({ ...style, text: `[^${k.attrs['w:id']}]` }); break
        }
      }
    }
    visit(p)
    return out
  }

  plain (p) {
    return this.segments(p).map((s) => s.text ?? '').join('')
  }

  markdown (p) {
    let s = ''
    for (const g of this.segments(p)) {
      if (g.pagebreak) { s += ' [page break] '; continue }
      if (g.image !== undefined || g.alt !== undefined) {
        const cm = (e) => (e / EMU_PER_CM).toFixed(1)
        s += `[image: ${g.alt || 'picture'}${g.cx ? ` ${cm(g.cx)}×${cm(g.cy)} cm` : ''}]`
        if (g.textbox) s += ` [text box: ${g.textbox}]`
        continue
      }
      let x = g.text
      if (!x.trim()) { s += x; continue }
      const lead = x.match(/^\s*/)[0]
      const trail = x.match(/\s*$/)[0]
      x = x.trim()
      if (g.s) x = `~~${x}~~`
      if (g.i) x = `*${x}*`
      if (g.b) x = `**${x}**`
      if (g.link) x = `[${x}](${g.link})`
      s += lead + x + trail
    }
    return s.replace(/\*\*\*\*/g, '').replace(/~~~~/g, '')
  }

  blockText (b) {
    if (b.kind === 'p') {
      const p = b.node
      const h = this.headingLevel(p)
      const md = this.markdown(p)
      if (h) return `${'#'.repeat(h)} ${md}`
      const li = this.list(p)
      if (li) return `${'  '.repeat(li.level)}${li.ordered ? '1.' : '-'} ${md}`
      const st = this.style(p)
      const tag = st && !st.default && !/^normal$/i.test(st.name) ? `{${st.name}} ` : ''
      return tag + md
    }
    const rows = this.tableRows(b.node)
    const cols = Math.max(0, ...rows.map((r) => r.length))
    const lines = [`table ${rows.length} rows × ${cols} columns`]
    rows.forEach((r, i) => {
      lines.push('| ' + r.map((c) => c.text.replace(/\|/g, '\\|').replace(/\n/g, '<br>')).join(' | ') + ' |')
      if (i === 0) lines.push('|' + ' --- |'.repeat(r.length))
    })
    return lines.join('\n')
  }

  // tableRows: rows of { text, node, span, vmerge }.
  tableRows (tbl) {
    return children(tbl, 'w:tr').map((tr) => cellsOf(tr).map((tc) => ({
      node: tc,
      text: children(tc, 'w:p').map((p) => this.plain(p)).join('\n') + descendants(tc, 'w:tbl').map(() => ' [nested table]').join(''),
      span: +(xpath(tc, 'w:tcPr', 'w:gridSpan')?.attrs['w:val'] ?? 1),
      vmerge: xpath(tc, 'w:tcPr', 'w:vMerge') ? (xpath(tc, 'w:tcPr', 'w:vMerge').attrs['w:val'] === 'restart' ? 'restart' : 'continue') : undefined
    })))
  }
}

// cellsOf finds a row's cells, also inside content controls.
function cellsOf (tr) {
  const out = []
  for (const k of tr.kids) {
    if (k.name === 'w:tc') out.push(k)
    else if (k.name === 'w:sdt' || k.name === 'w:customXml') { const c = child(k, 'w:sdtContent') ?? k; out.push(...cellsOf(c)) }
  }
  return out
}

const LIMIT = 16000

export function readDocx (buf, { from = 1 } = {}) {
  const pkg = Package.fromBuffer(buf)
  const m = new Model(pkg)
  const np = m.blocks.filter((b) => b.kind === 'p').length
  const nt = m.blocks.length - np
  const images = descendants(m.body, (n) => n.name === 'w:drawing' || n.name === 'w:pict').length
  const headers = pkg.relsOfType(m.main, 'header').length
  const footers = pkg.relsOfType(m.main, 'footer').length
  const comments = pkg.relsOfType(m.main, 'comments')[0]?.part
  const nComments = comments && pkg.has(comments) ? children(pkg.xml(comments).root, 'w:comment').length : 0
  const tracked = descendants(m.body, (n) => n.name === 'w:ins' || n.name === 'w:del').length
  let head = `Word document: ${np} paragraphs, ${nt} tables, ${images} images`
  const extras = []
  if (headers || footers) extras.push(`${headers} headers, ${footers} footers (edit their text with docx_edit replace)`)
  if (nComments) extras.push(`${nComments} comments`)
  if (tracked) extras.push(`${tracked} tracked changes (shown as accepted)`)
  if (extras.length) head += '; ' + extras.join('; ')
  head += '.\nRefs [p…] and [t…] name blocks for docx_edit.\n'
  const lines = []
  let size = 0
  let next
  const start = Math.max(1, from)
  for (let i = start - 1; i < m.blocks.length; i++) {
    const b = m.blocks[i]
    const line = `[${b.ref}] ${m.blockText(b)}`
    if (size + line.length > LIMIT && lines.length) { next = i + 1; break }
    lines.push(line)
    size += line.length + 1
  }
  let out = head
  if (start > 1) out += `(starting at block ${start} of ${m.blocks.length})\n`
  out += '\n' + (lines.length ? lines.join('\n') : '(the document has no text)')
  if (next) out += `\n\n[the document goes on: call office_read with from: ${next} for the next part]`
  return out
}

// ------------------------------------------------------------- create & edit

export async function createDocx (markdown, io) {
  const pkg = newPackage()
  const w = new Writer(pkg, io)
  const nodes = await w.blocks(markdown)
  const body = child(pkg.xml(w.main).root, 'w:body')
  body.kids.splice(body.kids.length - 1, 0, ...nodes)
  pkg.touch(w.main)
  const m = new Model(pkg)
  return { buffer: pkg.toBuffer(), summary: `${m.blocks.length} blocks` }
}

export async function editDocx (buf, edits, io) {
  const pkg = Package.fromBuffer(buf)
  const m = new Model(pkg)
  let writer
  const W_ = () => (writer ??= new Writer(pkg, io))
  const gone = new Set()
  const tails = new Map()
  const block = (ref, op) => {
    const b = m.refs.get(String(ref ?? '').replace(/^\[|\]$/g, ''))
    if (!b) throw new Error(`edit ${op}: there is no block ${ref}; call office_read for the current refs`)
    if (gone.has(b.ref)) throw new Error(`edit ${op}: block ${ref} was already removed by an earlier edit in this call`)
    return b
  }
  const report = []
  for (const [k, e] of edits.entries()) {
    const op = e?.op
    const where = `edits[${k}] (${op})`
    switch (op) {
      case 'replace': {
        if (!e.find) throw new Error(`${where}: find is empty`)
        const n = replaceInParts(pkg, docxTextParts(pkg, m.main), e.find, e.replace ?? '', !!e.all, where, wordRuns)
        report.push(`replaced ${n}×`)
        break
      }
      case 'insert': {
        const ref = e.after ?? e.before
        if (!ref) throw new Error(`${where}: give after or before (a block ref)`)
        const b = block(ref, op)
        const nodes = await W_().blocks(e.markdown ?? '')
        if (e.after !== undefined) {
          const anchor = tails.get(b.ref) ?? b.node
          const i = b.owner.kids.indexOf(anchor)
          b.owner.kids.splice(i + 1, 0, ...nodes)
          if (nodes.length) tails.set(b.ref, nodes[nodes.length - 1])
        } else {
          b.owner.kids.splice(b.owner.kids.indexOf(b.node), 0, ...nodes)
        }
        report.push(`inserted ${nodes.length} block(s) ${e.after !== undefined ? 'after' : 'before'} ${b.ref}`)
        break
      }
      case 'append': {
        const nodes = await W_().blocks(e.markdown ?? '')
        const body = m.body
        const last = body.kids[body.kids.length - 1]
        const at = last?.name === 'w:sectPr' ? body.kids.length - 1 : body.kids.length
        body.kids.splice(at, 0, ...nodes)
        report.push(`appended ${nodes.length} block(s)`)
        break
      }
      case 'set': {
        const b = block(e.block, op)
        const nodes = await W_().blocks(e.markdown ?? '')
        const sect = xpath(b.node, 'w:pPr', 'w:sectPr')
        if (sect) {
          // The paragraph ends a section: its section properties move to
          // the last new paragraph.
          let last = [...nodes].reverse().find((n) => n.name === 'w:p')
          if (!last) { last = el('w:p'); nodes.push(last) }
          let ppr = child(last, 'w:pPr')
          if (!ppr) { ppr = el('w:pPr'); last.kids.unshift(ppr) }
          insertOrdered(ppr, sect, PPR_ORDER)
        }
        b.owner.kids.splice(b.owner.kids.indexOf(b.node), 1, ...nodes)
        gone.add(b.ref)
        report.push(`rewrote ${b.ref} as ${nodes.length} block(s)`)
        break
      }
      case 'delete': {
        const refs = Array.isArray(e.blocks) ? e.blocks : [e.blocks ?? e.block]
        for (const r of refs) {
          const b = block(r, op)
          const sect = xpath(b.node, 'w:pPr', 'w:sectPr')
          if (sect) b.node.kids = [child(b.node, 'w:pPr')]
          else remove(b.owner, b.node)
          gone.add(b.ref)
        }
        // A table cell or the body must keep a paragraph after a table.
        report.push(`deleted ${refs.length} block(s)`)
        break
      }
      case 'cell': {
        const b = block(e.table, op)
        if (b.kind !== 't') throw new Error(`${where}: ${b.ref} is not a table`)
        const rows = children(b.node, 'w:tr')
        const r = rows[(e.row ?? 0) - 1]
        if (!r) throw new Error(`${where}: table ${b.ref} has ${rows.length} rows (row starts at 1)`)
        const cells = cellsOf(r)
        const tc = cells[(e.col ?? 0) - 1]
        if (!tc) throw new Error(`${where}: row ${e.row} of ${b.ref} has ${cells.length} cells (col starts at 1)`)
        setCellText(tc, String(e.text ?? ''))
        report.push(`set ${b.ref} row ${e.row} col ${e.col}`)
        break
      }
      default:
        throw new Error(`${where}: unknown op; use replace, insert, append, set, delete or cell`)
    }
  }
  // Word wants a paragraph right before sectPr when the body ends with a table.
  const kids = m.body.kids.filter(isEl)
  if (kids.length >= 2 && kids[kids.length - 1].name === 'w:sectPr' && kids[kids.length - 2].name === 'w:tbl') {
    m.body.kids.splice(m.body.kids.indexOf(kids[kids.length - 1]), 0, el('w:p'))
  }
  for (const tc of descendants(m.body, 'w:tc')) {
    if (!children(tc, 'w:p').length && !children(tc, 'w:tbl').length) tc.kids.push(el('w:p'))
  }
  pkg.touch(m.main)
  return { buffer: pkg.toBuffer(), summary: report.join('; ') }
}

function setCellText (tc, s) {
  const ps = children(tc, 'w:p')
  const first = ps[0]
  const ppr = first && child(first, 'w:pPr')
  const rpr = first && (descendants(first, 'w:r').map((r) => child(r, 'w:rPr')).find(Boolean) ?? xpath(first, 'w:pPr', 'w:rPr'))
  tc.kids = tc.kids.filter((k) => k.name !== 'w:p' && k.name !== 'w:tbl' && k.name !== 'w:sdt')
  for (const line of s.split('\n')) {
    tc.kids.push(el('w:p', {}, [ppr ? clone(ppr) : null, line ? el('w:r', {}, [rpr ? clone(rpr) : null, ...textNodes(line)]) : null]))
  }
}

function docxTextParts (pkg, main) {
  const parts = [main]
  for (const r of pkg.rels(main)) {
    if (/\/(header|footer|footnotes|endnotes)$/.test(r.type) && r.part && pkg.has(r.part)) parts.push(r.part)
  }
  return parts
}

// wordRuns lists, per paragraph, the w:t elements of live runs.
function wordRuns (root) {
  return descendants(root, 'w:p').map((p) => {
    const ts = []
    const visit = (n) => {
      for (const k of n.kids) {
        if (!isEl(k)) continue
        if (k.name === 'w:del' || k.name === 'w:moveFrom' || k.name === 'w:p') continue
        if (k.name === 'w:t') ts.push(k)
        else if (k.name !== 'w:instrText') visit(k)
      }
    }
    visit(p)
    return ts
  })
}

// replaceInParts replaces literal text across the text runs of a
// paragraph; the replacement takes the formatting of the run where the
// match starts.
export function replaceInParts (pkg, parts, find, repl, all, where, runsOf) {
  const plan = []
  for (const part of parts) {
    const d = pkg.xml(part)
    for (const ts of runsOf(d.root)) {
      const segs = ts.map((t) => ({ t, s: textContent(t) }))
      const joined = segs.map((x) => x.s).join('')
      let i = joined.indexOf(find)
      while (i >= 0) {
        plan.push({ part, segs, at: i })
        i = joined.indexOf(find, i + find.length)
      }
    }
  }
  if (!plan.length) throw new Error(`${where}: "${find}" was not found`)
  if (plan.length > 1 && !all) throw new Error(`${where}: "${find}" occurs ${plan.length} times; make find longer to pick one, or set all: true`)
  // Apply from the end of each paragraph so earlier offsets stay valid.
  for (const p of [...plan].reverse()) {
    let pos = 0
    let left = find.length
    let placed = false
    for (const seg of p.segs) {
      const len = seg.s.length
      const segStart = pos
      pos += len
      if (left <= 0) break
      if (p.at >= segStart + len) continue
      const from = Math.max(0, p.at - segStart)
      const take = Math.min(len - from, left)
      const s = seg.s.slice(0, from) + (placed ? '' : repl) + seg.s.slice(from + take)
      placed = true
      left -= take
      seg.s = s
      seg.t.kids = s ? [text(s)] : []
      if (/^\s|\s$/.test(s)) seg.t.attrs['xml:space'] = 'preserve'
    }
    pkg.touch(p.part)
  }
  return plan.length
}

// ------------------------------------------------------------- rendering

const escH = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export function renderDocx (buf) {
  const pkg = Package.fromBuffer(buf)
  const m = new Model(pkg)
  const counters = new Map()
  const para = (p) => {
    const segs = m.segments(p)
    let html = ''
    for (const g of segs) {
      if (g.pagebreak) { html += '<span class="pb"></span>'; continue }
      if (g.alt !== undefined) {
        const src = g.image && dataUri(pkg, g.image)
        const w = g.cx ? Math.round(g.cx / EMU_PER_PX) : 0
        html += src ? `<img src="${src}" alt="${escH(g.alt)}"${w ? ` style="width:${w}px;max-width:100%"` : ''}>` : `<span class="missing">[${escH(g.alt || 'image')}]</span>`
        if (g.textbox) html += `<span class="tb">${escH(g.textbox)}</span>`
        continue
      }
      let x = escH(g.text).replace(/\n/g, '<br>').replace(/\t/g, '<span class="tab"></span>')
      const css = []
      if (g.color && g.color !== 'auto') css.push(`color:#${g.color}`)
      if (g.size) css.push(`font-size:${g.size / 2}pt`)
      if (g.highlight && /^[0-9A-Fa-f]{6}$/.test(g.highlight)) css.push(`background:#${g.highlight}`)
      else if (g.highlight && g.highlight !== 'none' && g.highlight !== 'auto') css.push(`background:${g.highlight}`)
      if (css.length) x = `<span style="${css.join(';')}">${x}</span>`
      if (g.b) x = `<b>${x}</b>`
      if (g.i) x = `<i>${x}</i>`
      if (g.u) x = `<u>${x}</u>`
      if (g.s) x = `<s>${x}</s>`
      if (g.vert === 'superscript') x = `<sup>${x}</sup>`
      if (g.vert === 'subscript') x = `<sub>${x}</sub>`
      if (g.link) x = `<a>${x}</a>`
      html += x
    }
    const jc = xpath(p, 'w:pPr', 'w:jc')?.attrs['w:val']
    const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc]
    const style = align ? ` style="text-align:${align}"` : ''
    const h = m.headingLevel(p)
    if (h) return `<h${Math.min(h, 6)}${style}>${html || '&nbsp;'}</h${Math.min(h, 6)}>`
    const li = m.list(p)
    if (li) {
      const key = xpath(p, 'w:pPr', 'w:numPr', 'w:numId')?.attrs['w:val'] + ':' + li.level
      const n = (counters.get(key) ?? 0) + 1
      counters.set(key, n)
      const mark = li.ordered ? `${n}.` : ['•', '◦', '▪'][li.level % 3]
      return `<p class="li" style="margin-left:${1.5 + li.level * 1.5}em${align ? ';text-align:' + align : ''}"><span class="mk">${mark}</span>${html}</p>`
    }
    const name = m.style(p)?.name?.toLowerCase() ?? ''
    const cls = name === 'quote' || name === 'intense quote' ? ' class="q"' : name.includes('code') ? ' class="code"' : ''
    return `<p${cls}${style}>${html || '&nbsp;'}</p>`
  }
  const table = (tbl) => {
    const rows = m.tableRows(tbl)
    // Row spans from vertical merges.
    const grid = rows.map((r) => {
      let col = 0
      return r.map((c) => { const x = { ...c, col }; col += c.span; return x })
    })
    const look = xpath(tbl, 'w:tblPr', 'w:tblLook')
    const header = look && (look.attrs['w:firstRow'] === '1' || (parseInt(look.attrs['w:val'] ?? '0', 16) & 0x20))
    let html = '<table>'
    grid.forEach((r, ri) => {
      html += ri === 0 && header ? '<tr class="hd">' : '<tr>'
      for (const c of r) {
        if (c.vmerge === 'continue') continue
        let rs = 1
        if (c.vmerge === 'restart') {
          for (let k = ri + 1; k < grid.length; k++) {
            const below = grid[k].find((x) => x.col === c.col)
            if (below?.vmerge === 'continue') rs++
            else break
          }
        }
        const fill = xpath(c.node, 'w:tcPr', 'w:shd')?.attrs['w:fill']
        const inner = c.node.kids.map((k) => k.name === 'w:p' ? para(k) : k.name === 'w:tbl' ? table(k) : '').join('')
        html += `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${rs > 1 ? ` rowspan="${rs}"` : ''}${fill && /^[0-9A-F]{6}$/i.test(fill) ? ` style="background:#${fill}"` : ''}>${inner}</td>`
      }
      html += '</tr>'
    })
    return html + '</table>'
  }
  const body = m.blocks.map((b) => b.kind === 'p' ? para(b.node) : table(b.node)).join('\n')
  return {
    kind: 'Word 文档',
    css: DOCX_CSS,
    body: `<div class="page">${body.replace(/<span class="pb"><\/span>/g, '</div><div class="page">')}</div>`
  }
}

const DOCX_CSS = `
.page{background:#fff;width:794px;max-width:calc(100% - 32px);margin:24px auto;padding:72px 84px;box-sizing:border-box;box-shadow:0 1px 4px rgba(0,0,0,.18);font:11pt/1.5 Calibri,"Microsoft YaHei","PingFang SC",sans-serif;color:#000}
.page h1,.page h2,.page h3,.page h4,.page h5,.page h6{color:#1F3864;margin:1em 0 .4em;line-height:1.25}
.page h1{font-size:16pt}.page h2{font-size:14pt}.page h3{font-size:12pt}.page h4,.page h5,.page h6{font-size:11pt}
.page p{margin:0 0 6pt}
.page p.li{text-indent:-1.2em;padding-left:1.2em;margin-bottom:3pt}.page .mk{display:inline-block;width:1.2em}
.page p.q{border-left:3px solid #bfbfbf;padding-left:10px;color:#595959;font-style:italic}
.page p.code{font-family:Consolas,monospace;background:#f2f2f2;padding:6px 8px;white-space:pre-wrap}
.page table{border-collapse:collapse;width:100%;margin:6pt 0}.page td{border:1px solid #bfbfbf;padding:3px 7px;vertical-align:top}.page td p{margin:0}
.page tr.hd td{background:#D9E2F3;font-weight:600}
.page img{display:inline-block;vertical-align:middle}
.page .tab{display:inline-block;width:3em}
.page .missing{color:#999}.page .tb{border:1px dashed #bbb;padding:2px 4px;margin-left:4px}
.page a{color:#0563C1;text-decoration:underline}`

// ------------------------------------------------------------- editing in the UI
//
// The browser edits a document as blocks of runs. docxUiModel describes
// the document that way; editDocxUi applies what the person changed. A
// changed paragraph keeps its paragraph properties, bookmarks and
// pictures; each new run starts from the formatting of the text it
// replaces, then takes the run properties the person set. Paragraphs with
// fields, comments, tracked changes, content controls or equations are
// shown but not editable, so nothing they hold is lost.

// Comment anchors may sit in paragraphs the editor changes (it carries
// them as marker runs); table cells and headers still lock on them.
const CM_NAMES = new Set(['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference'])
const LOCKED = new Set(['w:fldChar', 'w:fldSimple', 'w:instrText', 'w:sdt', 'm:oMath', 'm:oMathPara', 'w:moveFrom', 'w:moveTo', 'w:object', 'w:txbxContent', 'w:smartTag', 'w:customXml', 'w:ruby', 'w:delInstrText'])
const LOCKED_CELL = new Set([...LOCKED, ...CM_NAMES, 'w:ins', 'w:del', 'w:footnoteReference', 'w:endnoteReference'])
const tri = (n) => (n ? on(n) : undefined)
const HL = { yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF' }

// runProps reads a run's own formatting, tri-state: true, false (turned
// off against the style) or undefined (inherited).
function runProps (pr) {
  const u = child(pr, 'w:u')
  const hl = child(pr, 'w:highlight')?.attrs['w:val']
  const shd = child(pr, 'w:shd')?.attrs['w:fill']
  const color = child(pr, 'w:color')?.attrs['w:val']
  const rf = child(pr, 'w:rFonts')
  return {
    font: rf ? (rf.attrs['w:eastAsia'] ?? rf.attrs['w:ascii'] ?? rf.attrs['w:hAnsi']) : undefined,
    b: tri(child(pr, 'w:b')),
    i: tri(child(pr, 'w:i')),
    u: u ? u.attrs['w:val'] !== 'none' : undefined,
    s: tri(child(pr, 'w:strike')),
    color: color && /^[0-9A-Fa-f]{6}$/.test(color) ? color.toUpperCase() : undefined,
    size: child(pr, 'w:sz') ? +child(pr, 'w:sz').attrs['w:val'] : undefined,
    highlight: hl && hl !== 'none' ? (HL[hl] ?? undefined) : shd && /^[0-9A-Fa-f]{6}$/.test(shd) ? shd.toUpperCase() : undefined,
    vert: child(pr, 'w:vertAlign')?.attrs['w:val'] === 'superscript' ? 'sup' : child(pr, 'w:vertAlign')?.attrs['w:val'] === 'subscript' ? 'sub' : undefined
  }
}

// enforce makes rPr carry exactly the run properties in want, touching
// only those that differ.
function enforce (rpr, want) {
  const have = runProps(rpr)
  const put = (...nodes) => { for (const n of nodes) insertOrdered(rpr, n, RPR_ORDER) }
  const drop = (...names) => { rpr.kids = rpr.kids.filter((k) => !names.includes(k.name)) }
  const flag = (key, name, cs) => {
    if (have[key] === want[key]) return
    drop(name, ...(cs ? [cs] : []))
    if (want[key] === true) put(el(name), ...(cs ? [el(cs)] : []))
    if (want[key] === false) put(v(name, 0), ...(cs ? [v(cs, 0)] : []))
  }
  flag('b', 'w:b', 'w:bCs')
  flag('i', 'w:i', 'w:iCs')
  flag('s', 'w:strike')
  if (have.u !== want.u) { drop('w:u'); if (want.u !== undefined) put(v('w:u', want.u ? 'single' : 'none')) }
  if (have.color !== want.color) { drop('w:color'); if (want.color) put(v('w:color', want.color)) }
  if (have.size !== want.size) { drop('w:sz', 'w:szCs'); if (want.size) put(v('w:sz', want.size), v('w:szCs', want.size)) }
  if (have.highlight !== want.highlight) {
    drop('w:highlight', 'w:shd')
    if (want.highlight) {
      const named = Object.entries(HL).find(([, h]) => h === want.highlight)?.[0]
      put(named ? v('w:highlight', named) : el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': want.highlight }))
    }
  }
  if (have.vert !== want.vert) { drop('w:vertAlign'); if (want.vert) put(v('w:vertAlign', want.vert === 'sup' ? 'superscript' : 'subscript')) }
  if (have.font !== want.font && want.font !== undefined) {
    drop('w:rFonts')
    if (want.font) put(el('w:rFonts', { 'w:ascii': want.font, 'w:hAnsi': want.font, 'w:eastAsia': want.font, 'w:cs': want.font }))
  }
  return rpr
}

// charMap lists a paragraph's live runs with the text each holds, its
// pictures and bookmarks.
function charMap (p) {
  const owners = []
  let textAll = ''
  const drawings = []
  const marks = { start: [], end: [] }
  const cm = new Map()
  const insOf = new Map()
  const insEls = new Map()
  const dels = new Map()
  const delAt = new Map()
  const visit = (node, link, ins) => {
    for (const k of node.kids) {
      if (!isEl(k)) continue
      if (k.name === 'w:commentRangeStart') cm.set('cs:' + k.attrs['w:id'], k)
      else if (k.name === 'w:commentRangeEnd') cm.set('ce:' + k.attrs['w:id'], k)
      if (k.name === 'w:ins') { insEls.set(k.attrs['w:id'], k); visit(k, link, k); continue }
      if (k.name === 'w:del') { dels.set(k.attrs['w:id'], k); delAt.set(k.attrs['w:id'], textAll.length); continue }
      if (k.name === 'w:r') {
        if (ins) insOf.set(k, ins)
        const ref = child(k, 'w:commentReference')
        if (ref) cm.set('cr:' + ref.attrs['w:id'], k)
        const fn = child(k, 'w:footnoteReference') ?? child(k, 'w:endnoteReference')
        if (fn) cm.set((fn.name === 'w:footnoteReference' ? 'fn:' : 'en:') + fn.attrs['w:id'], k)
        for (const x of k.kids) {
          let s = ''
          if (x.name === 'w:t') s = textContent(x)
          else if (x.name === 'w:tab' || x.name === 'w:ptab') s = '\t'
          else if ((x.name === 'w:br' || x.name === 'w:cr') && x.attrs['w:type'] !== 'page') s = '\n'
          else if (x.name === 'w:noBreakHyphen') s = '-'
          else if (x.name === 'w:drawing' || x.name === 'w:pict') drawings.push(k)
          for (let i = 0; i < s.length; i++) owners.push(k)
          textAll += s
        }
      } else if (k.name === 'w:hyperlink') visit(k, k, ins)
      else if (k.name === 'w:bookmarkStart') marks.start.push(k)
      else if (k.name === 'w:bookmarkEnd') marks.end.push(k)
      else if (k.name === 'w:dir' || k.name === 'w:bdo') visit(k, link, ins)
    }
  }
  visit(p)
  return { text: textAll, owners, drawings, marks, cm, insOf, insEls, dels, delAt }
}

function uiRuns (m, p) {
  const out = []
  let pic = 0
  const visit = (node, link, rev) => {
    for (const k of node.kids) {
      if (!isEl(k)) continue
      if (k.name === 'w:hyperlink') {
        const id = k.attrs['r:id']
        visit(k, id ? m.pkg.rel(m.main, id)?.target : k.attrs['w:anchor'] ? '#' + k.attrs['w:anchor'] : undefined, rev)
      } else if (k.name === 'w:dir' || k.name === 'w:bdo') visit(k, link, rev)
      else if (k.name === 'w:ins') visit(k, link, 'ins:' + k.attrs['w:id'])
      else if (k.name === 'w:del') out.push({ t: 'del', id: k.attrs['w:id'], text: descendants(k, (x) => x.name === 'w:delText' || x.name === 'w:t').map(textContent).join('') })
      else if (k.name === 'w:commentRangeStart') out.push({ t: 'cs', id: k.attrs['w:id'] })
      else if (k.name === 'w:commentRangeEnd') out.push({ t: 'ce', id: k.attrs['w:id'] })
      else if (k.name === 'w:r' && child(k, 'w:commentReference')) out.push({ t: 'cr', id: child(k, 'w:commentReference').attrs['w:id'] })
      else if (k.name === 'w:r' && child(k, 'w:footnoteReference')) out.push({ t: 'fn', id: child(k, 'w:footnoteReference').attrs['w:id'] })
      else if (k.name === 'w:r' && child(k, 'w:endnoteReference')) out.push({ t: 'en', id: child(k, 'w:endnoteReference').attrs['w:id'] })
      else if (k.name === 'w:r') {
        const props = runProps(child(k, 'w:rPr'))
        let s = ''
        const flush = () => { if (s) { out.push({ t: 'text', text: s, ...props, ...(link ? { link } : {}), ...(rev ? { rev } : {}) }); s = '' } }
        for (const x of k.kids) {
          if (x.name === 'w:t') s += textContent(x)
          else if (x.name === 'w:tab' || x.name === 'w:ptab') s += '\t'
          else if (x.name === 'w:br' || x.name === 'w:cr') {
            if (x.attrs['w:type'] === 'page') { flush(); out.push({ t: 'pb' }) } else s += '\n'
          } else if (x.name === 'w:noBreakHyphen') s += '-'
          else if (x.name === 'w:drawing' || x.name === 'w:pict') {
            flush()
            const blip = descendants(x, (n) => n.name === 'a:blip' || n.name === 'v:imagedata')[0]
            const rid = blip?.attrs['r:embed'] ?? blip?.attrs['r:id']
            const part = rid ? m.pkg.rel(m.main, rid)?.part : undefined
            const ext = descendants(x, 'wp:extent')[0]
            out.push({ t: 'img', n: pic++, src: part && m.pkg.has(part) ? dataUri(m.pkg, part) : undefined, w: ext ? Math.round(+ext.attrs.cx / EMU_PER_PX) : 0, h: ext ? Math.round(+ext.attrs.cy / EMU_PER_PX) : 0, ...wrapOf(x) })
          }
        }
        flush()
      }
    }
  }
  visit(p)
  return out
}

// styleLook resolves what a paragraph style looks like, for the editor.
function styleLook (m, id) {
  const out = {}
  let st = id ? m.styleInfo.get(id) : [...m.styleInfo.values()].find((s) => s.type === 'paragraph' && s.default)
  for (let k = 0; st && k < 8; k++) {
    const pr = xpath(st.node, 'w:rPr')
    const p = runProps(pr)
    for (const key of ['b', 'i', 'color', 'size']) if (out[key] === undefined && p[key] !== undefined) out[key] = p[key]
    const jc = xpath(st.node, 'w:pPr', 'w:jc')?.attrs['w:val']
    if (out.align === undefined && jc) out.align = jc
    const sp = xpath(st.node, 'w:pPr', 'w:spacing')
    if (out.before === undefined && sp?.attrs['w:before']) out.before = +sp.attrs['w:before']
    if (out.after === undefined && sp?.attrs['w:after']) out.after = +sp.attrs['w:after']
    const ind = xpath(st.node, 'w:pPr', 'w:ind')
    if (out.indent === undefined && ind?.attrs['w:left']) out.indent = +ind.attrs['w:left']
    const shd = xpath(st.node, 'w:pPr', 'w:shd')?.attrs['w:fill']
    if (out.shade === undefined && shd && /^[0-9A-F]{6}$/i.test(shd)) out.shade = shd
    st = st.basedOn ? m.styleInfo.get(st.basedOn) : undefined
  }
  return out
}

// spacingOf and indentOf read a paragraph's own spacing and indents, in
// twips (line: 240 is single spacing).
function spacingOf (p) {
  const s = xpath(p, 'w:pPr', 'w:spacing')
  if (!s) return undefined
  const out = {}
  if (s.attrs['w:line'] && (!s.attrs['w:lineRule'] || s.attrs['w:lineRule'] === 'auto')) out.line = +s.attrs['w:line']
  if (s.attrs['w:before'] !== undefined) out.before = +s.attrs['w:before']
  if (s.attrs['w:after'] !== undefined) out.after = +s.attrs['w:after']
  return Object.keys(out).length ? out : undefined
}
function indentOf (p) {
  const s = xpath(p, 'w:pPr', 'w:ind')
  if (!s) return undefined
  const out = {}
  const left = s.attrs['w:left'] ?? s.attrs['w:start']
  if (left !== undefined) out.left = +left
  if (s.attrs['w:firstLine'] !== undefined) out.firstLine = +s.attrs['w:firstLine']
  if (s.attrs['w:hanging'] !== undefined) out.firstLine = -s.attrs['w:hanging']
  return Object.keys(out).length ? out : undefined
}

export function docxUiModel (buf) {
  const pkg = Package.fromBuffer(buf)
  const m = new Model(pkg)
  const tocOwners = new Set(tocBlocks(m).map((x) => child(x, 'w:sdtContent')))
  const counters = new Map()
  const numbering = {}
  const looks = {}
  const look = (id) => { const k = id ?? ''; if (!(k in looks)) looks[k] = styleLook(m, id); return k }
  const para = (b) => {
    const p = b.node
    const style = xpath(p, 'w:pPr', 'w:pStyle')?.attrs['w:val'] ?? null
    const li = m.list(p)
    let list
    if (li) {
      const key = xpath(p, 'w:pPr', 'w:numPr', 'w:numId')?.attrs['w:val'] + ':' + li.level
      const n = (counters.get(key) ?? 0) + 1
      counters.set(key, n)
      list = { level: li.level, ordered: li.ordered, mark: li.ordered ? `${n}.` : ['•', '◦', '▪'][li.level % 3], ...(m.numLvls.has(li.numId) ? { num: li.numId } : {}) }
      if (m.numLvls.has(li.numId)) numbering[li.numId] ??= m.numLvls.get(li.numId)
    }
    const fp = xpath(p, 'w:pPr', 'w:framePr')
    const dropCap = fp && ['drop', 'margin'].includes(fp.attrs['w:dropCap']) ? +(fp.attrs['w:lines'] ?? 3) : undefined
    const shd = xpath(p, 'w:pPr', 'w:shd')?.attrs['w:fill']
    const locked = descendants(p, (n) => LOCKED.has(n.name)).length > 0
    const toc = tocOwners.has(b.owner) ? (tocLevel(m, p) || 0) : undefined
    return {
      ref: b.ref,
      kind: 'p',
      ...(toc !== undefined ? { toc } : {}),
      style,
      look: look(style),
      heading: m.headingLevel(p) || undefined,
      list,
      align: xpath(p, 'w:pPr', 'w:jc')?.attrs['w:val'],
      spacing: spacingOf(p),
      indent: indentOf(p),
      runs: uiRuns(m, p),
      ...(paraBorder(p) ? { border: paraBorder(p) } : {}),
      ...(shd && /^[0-9A-F]{6}$/i.test(shd) ? { shade: shd.toUpperCase() } : {}),
      ...(dropCap ? { drop: dropCap } : {}),
      ...(locked || toc !== undefined ? { locked: true } : {})
    }
  }
  const table = (b) => {
    const grid = tableGrid(b.node)
    const widths = gridWidths(b.node)
    return {
      ref: b.ref,
      kind: 't',
      rows: m.tableRows(b.node).map((r, ri) => r.map((c, ci) => {
        const p0 = descendants(c.node, 'w:p')[0]
        const r0 = p0 && descendants(p0, 'w:r').find((x) => child(x, 'w:t'))
        const rp = runProps(child(r0, 'w:rPr') ?? xpath(p0, 'w:pPr', 'w:rPr'))
        const valign = xpath(c.node, 'w:tcPr', 'w:vAlign')?.attrs['w:val']
        return {
          text: c.text,
          span: c.span,
          vmerge: c.vmerge,
          g: grid[ri]?.[ci]?.g ?? ci,
          fill: xpath(c.node, 'w:tcPr', 'w:shd')?.attrs['w:fill'],
          ...(valign && valign !== 'top' ? { valign } : {}),
          ...(p0 && xpath(p0, 'w:pPr', 'w:jc') ? { align: xpath(p0, 'w:pPr', 'w:jc').attrs['w:val'] } : {}),
          ...(rp.b !== undefined ? { b: rp.b } : {}),
          ...(rp.color ? { color: rp.color } : {}),
          ...(bordersOf(xpath(c.node, 'w:tcPr', 'w:tcBorders')) ? { bd: bordersOf(xpath(c.node, 'w:tcPr', 'w:tcBorders')) } : {}),
          ...(descendants(c.node, (n) => LOCKED_CELL.has(n.name) || n.name === 'w:tbl' || n.name === 'w:drawing').length ? { locked: true } : {})
        }
      })),
      widths: widths.every((w) => w > 0) ? widths.map((w) => Math.round(w / 15)) : undefined,
      borders: { ...tableStyleBorders(m, b.node), ...bordersOf(xpath(b.node, 'w:tblPr', 'w:tblBorders')) },
      header: (() => { const l = xpath(b.node, 'w:tblPr', 'w:tblLook'); return !!(l && (l.attrs['w:firstRow'] === '1' || (parseInt(l.attrs['w:val'] ?? '0', 16) & 0x20))) })()
    }
  }
  const blocks = m.blocks.map((b) => (b.kind === 'p' ? para(b) : table(b)))
  const sect = child(m.body, 'w:sectPr')
  const pg = child(sect, 'w:pgSz')
  const mar = child(sect, 'w:pgMar')
  const styles = [...m.styleInfo.entries()]
    .filter(([, s]) => s.type === 'paragraph' && !child(s.node, 'w:semiHidden') && !child(s.node, 'w:hidden'))
    .map(([id, s]) => ({ id, name: s.name, look: look(id) }))
  look(null)
  return {
    kind: 'docx',
    blocks,
    styles,
    looks,
    page: {
      width: pg ? Math.round(+pg.attrs['w:w'] / 15) : 794,
      height: pg ? Math.round(+pg.attrs['w:h'] / 15) : 1123,
      left: mar ? Math.round(+mar.attrs['w:left'] / 15) : 96,
      right: mar ? Math.round(+mar.attrs['w:right'] / 15) : 96,
      top: mar ? Math.round(+mar.attrs['w:top'] / 15) : 96,
      bottom: mar ? Math.round(+mar.attrs['w:bottom'] / 15) : 96,
      size: pg ? Object.entries(PAGE_SIZES).find(([, d]) => (d[0] === +pg.attrs['w:w'] && d[1] === +pg.attrs['w:h']) || (d[1] === +pg.attrs['w:w'] && d[0] === +pg.attrs['w:h']))?.[0] : 'Letter',
      orient: pg && +pg.attrs['w:w'] > +pg.attrs['w:h'] ? 'landscape' : 'portrait',
      margin: mar ? Object.entries(MARGINS).find(([, d]) => d[0] === +mar.attrs['w:top'] && d[1] === +mar.attrs['w:left'])?.[0] : undefined,
      cols: Math.max(1, +(child(sect, 'w:cols')?.attrs['w:num'] ?? 1)),
      colSpace: Math.round(+(child(sect, 'w:cols')?.attrs['w:space'] ?? 425) / 15)
    },
    header: hfInfo(m, 'header'),
    footer: hfInfo(m, 'footer'),
    comments: readComments(m),
    numbering,
    notes: readNotes(m),
    track: trackOn(m),
    revs: readRevisions(m),
    toc: tocBlocks(m).length > 0 || descendants(m.body, (n) => n.name === 'w:instrText' && /^\s*TOC\b/.test(textContent(n))).length > 0
  }
}

export async function editDocxUi (buf, ops, io) {
  const pkg = Package.fromBuffer(buf)
  const m = new Model(pkg)
  let writer
  const W_ = () => (writer ??= new Writer(pkg, io))
  const gone = new Set()
  const tails = new Map()
  const block = (ref, where) => {
    const b = m.refs.get(String(ref ?? ''))
    if (!b) throw new Error(`${where}: there is no block ${ref}; reload the document`)
    if (gone.has(b.ref)) throw new Error(`${where}: block ${ref} was already removed`)
    return b
  }
  const index = m.blocks.reduce((acc, b, i) => acc.set(b.ref, i), new Map())
  let tracking = trackOn(m)
  const noted = { fn: new Set(descendants(m.body, 'w:footnoteReference').map((x) => x.attrs['w:id'])), en: new Set(descendants(m.body, 'w:endnoteReference').map((x) => x.attrs['w:id'])) }
  const freshNotes = new Map()
  // Comments the body referenced before; one whose reference the person
  // deleted goes with it.
  const anchored = new Set(descendants(m.body, 'w:commentReference').map((x) => x.attrs['w:id']))
  const fresh = new Map()
  // Pictures are found by the block they came from, before any change.
  const maps = new Map(m.blocks.filter((b) => b.kind === 'p').map((b) => [b.ref, charMap(b.node)]))
  m.drawingsOf = (ref) => maps.get(String(ref))?.drawings
  let n = 0
  for (const [k, e] of ops.entries()) {
    const where = `ops[${k}] (${e?.op})`
    switch (e?.op) {
      case 'para': {
        const b = block(e.block, where)
        if (b.kind !== 'p') throw new Error(`${where}: ${b.ref} is not a paragraph`)
        if (descendants(b.node, (x) => LOCKED.has(x.name)).length) throw new Error(`${where}: ${b.ref} holds fields, content controls or tracked changes and cannot be edited here`)
        writePara(m, W_, b.node, e.runs ?? [], e.props, prevList(m, index.get(b.ref)), undefined, { track: tracking, author: e.author })
        break
      }
      case 'insert': {
        const ref = e.after ?? e.before
        let owner = m.body
        let at
        let like = e.like ? block(e.like, where) : undefined
        if (ref !== undefined && ref !== null) {
          const b = block(ref, where)
          owner = b.owner
          if (e.after !== undefined) { const anchor = tails.get(b.ref) ?? b.node; at = owner.kids.indexOf(anchor) + 1 } else at = owner.kids.indexOf(b.node)
          like ??= b.kind === 'p' ? b : undefined
        } else {
          const last = m.body.kids[m.body.kids.length - 1]
          at = last?.name === 'w:sectPr' ? m.body.kids.length - 1 : m.body.kids.length
        }
        const p = el('w:p')
        if (like?.kind === 'p') {
          const ppr = child(like.node, 'w:pPr')
          if (ppr) { const c = clone(ppr); c.kids = c.kids.filter((x) => x.name !== 'w:sectPr' && x.name !== 'w:pPrChange'); p.kids.push(c) }
        }
        writePara(m, W_, p, e.runs ?? [], e.props, undefined, like?.node, { track: tracking, author: e.author, fresh: true })
        owner.kids.splice(at, 0, p)
        if (e.after !== undefined && ref) tails.set(String(ref), p)
        break
      }
      case 'delete': {
        for (const r of Array.isArray(e.blocks) ? e.blocks : [e.block]) {
          const b = block(r, where)
          if (tracking && b.kind === 'p') { trackDelete(m, b.node, { author: e.author }); gone.add(b.ref); continue }
          const sect = xpath(b.node, 'w:pPr', 'w:sectPr')
          if (sect) b.node.kids = [child(b.node, 'w:pPr')]
          else remove(b.owner, b.node)
          gone.add(b.ref)
        }
        break
      }
      case 'cell': {
        const b = block(e.table, where)
        if (b.kind !== 't') throw new Error(`${where}: ${b.ref} is not a table`)
        const tc = cellsOf(children(b.node, 'w:tr')[e.row - 1] ?? el('x'))[e.col - 1]
        if (!tc) throw new Error(`${where}: table ${b.ref} has no cell at row ${e.row}, column ${e.col}`)
        setCellText(tc, String(e.text ?? ''))
        break
      }
      case 'row': {
        const b = block(e.table, where)
        const rows = children(b.node, 'w:tr')
        const tr = rows[e.row - 1]
        if (!tr) throw new Error(`${where}: table ${b.ref} has no row ${e.row}`)
        if (e.action === 'delete') {
          if (rows.length === 1) { remove(b.owner, b.node); gone.add(b.ref) } else remove(b.node, tr)
        } else {
          const copy = clone(tr)
          copy.kids = copy.kids.filter((x) => x.name !== 'w:trPr' || !child(x, 'w:tblHeader'))
          for (const tc of cellsOf(copy)) { const vm = xpath(tc, 'w:tcPr', 'w:vMerge'); if (vm) remove(child(tc, 'w:tcPr'), vm); setCellText(tc, '') }
          b.node.kids.splice(b.node.kids.indexOf(tr) + (e.action === 'above' ? 0 : 1), 0, copy)
        }
        break
      }
      case 'table': {
        const rows = Math.max(1, Math.min(50, +e.rows || 2))
        const cols = Math.max(1, Math.min(20, +e.cols || 2))
        const node = await W_().table({ rows: Array.from({ length: rows }, () => Array.from({ length: cols }, () => [])) })
        const b = e.after ? block(e.after, where) : undefined
        const owner = b?.owner ?? m.body
        let at
        if (b) { const anchor = tails.get(b.ref) ?? b.node; at = owner.kids.indexOf(anchor) + 1 } else { const last = owner.kids[owner.kids.length - 1]; at = last?.name === 'w:sectPr' ? owner.kids.length - 1 : owner.kids.length }
        owner.kids.splice(at, 0, node)
        if (b) tails.set(b.ref, node)
        break
      }
      case 'image': {
        // A picture of its own, centred, after (or before) a block.
        const ref = e.after ?? e.before
        const b = ref ? block(ref, where) : undefined
        const owner = b?.owner ?? m.body
        let at
        if (b && e.after !== undefined) { const anchor = tails.get(b.ref) ?? b.node; at = owner.kids.indexOf(anchor) + 1 } else if (b) at = owner.kids.indexOf(b.node)
        else { const last = owner.kids[owner.kids.length - 1]; at = last?.name === 'w:sectPr' ? owner.kids.length - 1 : owner.kids.length }
        const run = await W_().picture(String(e.src ?? ''), e.alt ?? '')
        const p = el('w:p', {}, [el('w:pPr', {}, [v('w:jc', 'center')]), run])
        owner.kids.splice(at, 0, p)
        if (b && e.after !== undefined) tails.set(b.ref, p)
        break
      }
      case 'col': {
        const b = block(e.table, where)
        if (b.kind !== 't') throw new Error(`${where}: ${b.ref} is not a table`)
        tableColumn(b, +e.col, e.action, where, () => gone.add(b.ref))
        break
      }
      case 'merge': case 'unmerge': case 'split': case 'twidth': case 'tcell': case 'tstyle': {
        const b = block(e.table, where)
        if (b.kind !== 't') throw new Error(`${where}: ${b.ref} is not a table`)
        if (e.op === 'merge') mergeCells(b.node, +e.r1, +e.r2, +e.g1, +e.g2, where)
        else if (e.op === 'unmerge') unmergeCell(b.node, +e.row, +e.g, where)
        else if (e.op === 'split') splitCell(b.node, +e.row, +e.g, +e.cols || 1, +e.rows || 1, where)
        else if (e.op === 'twidth') setWidths(b.node, e.widths, where)
        else if (e.op === 'tcell') formatCells(b.node, e, where)
        else styleTable(b.node, String(e.style), where)
        break
      }
      case 'liststyle': {
        const key = String(e.style ?? '')
        if (!LIST_STYLES[key]) throw new Error(`${where}: 未知的列表样式 ${key}`)
        const numId = W_().numbering.styled(key)
        for (const r of e.blocks ?? []) {
          const b = block(r, where)
          if (b.kind !== 'p') continue
          let ppr = child(b.node, 'w:pPr')
          if (!ppr) { ppr = el('w:pPr'); b.node.kids.unshift(ppr) }
          const level = m.list(b.node)?.level ?? 0
          setPr(ppr, 'w:numPr', el('w:numPr', {}, [v('w:ilvl', level), v('w:numId', numId)]), PPR_ORDER)
        }
        break
      }
      case 'dropcap': {
        const b = block(e.block, where)
        if (b.kind !== 'p') throw new Error(`${where}: ${b.ref} is not a paragraph`)
        setDropCap(m, b, +e.lines || 0, where)
        break
      }
      case 'track': {
        setTrack(m, !!e.on)
        tracking = !!e.on
        break
      }
      case 'rev': {
        const scope = e.block ? block(e.block, where).node : m.body
        resolveRevisions(m, scope, e.action === 'reject' ? 'reject' : 'accept', e.id !== undefined ? String(e.id) : undefined)
        break
      }
      case 'note_new': {
        const kind = e.kind === 'en' ? 'en' : 'fn'
        freshNotes.set(String(e.id), { kind, id: addNote(m, kind, String(e.text ?? '')) })
        break
      }
      case 'note_edit': {
        const kind = e.kind === 'en' ? 'en' : 'fn'
        editNote(m, kind, String(freshNotes.get(String(e.id))?.id ?? e.id), String(e.text ?? ''), where)
        break
      }
      case 'hf': {
        setHeaderFooter(m, e.kind === 'footer' ? 'footer' : 'header', e.pageNumber ? null : String(e.text ?? ''), where)
        break
      }
      case 'page': {
        setPage(m, e, where)
        break
      }
      case 'comment_new': {
        if (!e.id || !String(e.text ?? '').trim()) throw new Error(`${where}: id and text are required`)
        fresh.set(String(e.id), addComment(m, e).id)
        break
      }
      case 'comment_reply': replyComment(m, e, where); break
      case 'comment_edit': editComment(m, e, where); break
      case 'comment_done': doneComment(m, e, where); break
      case 'comment_delete': deleteComments(m, [String(e.id)], where); break
      case 'toc': {
        const b = e.after ?? e.before ? block(e.after ?? e.before, where) : undefined
        const sdt = buildToc(m, +e.levels || 3)
        const owner = b?.owner ?? m.body
        let at
        if (b) at = owner.kids.indexOf(b.node) + (e.after !== undefined ? 1 : 0)
        else at = 0
        owner.kids.splice(at, 0, sdt)
        break
      }
      case 'toc_update': updateToc(m, +e.levels || 3, where); break
      case 'toc_delete': {
        const all = tocBlocks(m)
        if (!all.length) throw new Error(`${where}: the document has no table of contents made here`)
        for (const x of all) remove(parentOf(m.body, x), x)
        break
      }
      default:
        throw new Error(`${where}: unknown op`)
    }
    n++
  }
  settleComments(m, anchored, fresh)
  settleNotes(m, noted, freshNotes)
  const kids = m.body.kids.filter(isEl)
  if (kids.length >= 2 && kids[kids.length - 1].name === 'w:sectPr' && kids[kids.length - 2].name === 'w:tbl') m.body.kids.splice(m.body.kids.indexOf(kids[kids.length - 1]), 0, el('w:p'))
  if (!m.body.kids.some((x) => x.name === 'w:p' || x.name === 'w:tbl')) m.body.kids.unshift(el('w:p'))
  for (const tc of descendants(m.body, 'w:tc')) if (!children(tc, 'w:p').length && !children(tc, 'w:tbl').length) tc.kids.push(el('w:p'))
  pkg.touch(m.main)
  return { buffer: pkg.toBuffer(), summary: `${n} change(s)` }
}

// tableColumn inserts a column left or right of col (1-based), or deletes
// it, in a table without merged cells.
function tableColumn (b, col, action, where, dropped) {
  const rows = children(b.node, 'w:tr')
  if (descendants(b.node, (x) => x.name === 'w:gridSpan' || x.name === 'w:vMerge' || x.name === 'w:hMerge').length) throw new Error(`${where}: 表格里有合并单元格，请在 Word/WPS 中增删列`)
  const grid = child(b.node, 'w:tblGrid')
  const gcols = children(grid, 'w:gridCol')
  for (const tr of rows) {
    const cells = cellsOf(tr)
    const tc = cells[col - 1]
    if (!tc) throw new Error(`${where}: table ${b.ref} has no column ${col}`)
    const owner = tr.kids.includes(tc) ? tr : descendants(tr, (x) => x.kids?.includes(tc))[0]
    if (action === 'delete') remove(owner, tc)
    else {
      const copy = clone(tc)
      setCellText(copy, '')
      owner.kids.splice(owner.kids.indexOf(tc) + (action === 'left' ? 0 : 1), 0, copy)
    }
  }
  if (action === 'delete') {
    if (cellsOf(rows[0]).length === 0) { remove(b.owner, b.node); dropped(); return }
    if (gcols[col - 1]) remove(grid, gcols[col - 1])
  } else if (gcols[col - 1]) grid.kids.splice(grid.kids.indexOf(gcols[col - 1]) + (action === 'left' ? 0 : 1), 0, clone(gcols[col - 1]))
}


// ------------------------------------------------------------- comments

const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
const CT_COMMENTS_EX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml'
const REL_COMMENTS_EX = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

function parentOf (root, node) {
  if (!root?.kids) return undefined
  for (const k of root.kids) {
    if (k === node) return root
    if (isEl(k)) { const p = parentOf(k, node); if (p) return p }
  }
  return undefined
}

// ignorable declares a Word extension namespace on a part's root so
// readers that do not know it skip it.
function ignorable (root, prefix, ns) {
  if (!root.attrs['xmlns:' + prefix]) root.attrs['xmlns:' + prefix] = ns
  if (!root.attrs['xmlns:mc']) root.attrs['xmlns:mc'] = MC
  const list = (root.attrs['mc:Ignorable'] ?? '').split(/\s+/).filter(Boolean)
  if (!list.includes(prefix)) root.attrs['mc:Ignorable'] = [...list, prefix].join(' ')
}

function commentsPart (m, create) {
  let part = m.pkg.relsOfType(m.main, 'comments')[0]?.part
  if (part && m.pkg.has(part)) { const root = m.pkg.xml(part).root; if (create) ignorable(root, 'w14', W14); return { part, root } }
  if (!create) return undefined
  part = m.pkg.has('word/comments.xml') ? m.pkg.uniqueName('word/comments#.xml') : 'word/comments.xml'
  m.pkg.setXml(part, doc(el('w:comments', { 'xmlns:w': W, 'xmlns:w14': W14, 'xmlns:mc': MC, 'mc:Ignorable': 'w14' }, [])))
  m.pkg.setOverride(part, CT_COMMENTS)
  m.pkg.addRel(m.main, REL('comments'), part)
  return { part, root: m.pkg.xml(part).root }
}

function commentsExPart (m, create) {
  let part = m.pkg.relsOfType(m.main, 'commentsExtended')[0]?.part
  if (part && m.pkg.has(part)) return { part, root: m.pkg.xml(part).root }
  if (!create) return undefined
  part = m.pkg.has('word/commentsExtended.xml') ? m.pkg.uniqueName('word/commentsExtended#.xml') : 'word/commentsExtended.xml'
  m.pkg.setXml(part, doc(el('w15:commentsEx', { 'xmlns:w15': W15, 'xmlns:mc': MC, 'mc:Ignorable': 'w15' }, [])))
  m.pkg.setOverride(part, CT_COMMENTS_EX)
  m.pkg.addRel(m.main, REL_COMMENTS_EX, part)
  return { part, root: m.pkg.xml(part).root }
}

const newParaId = () => (Math.floor(Math.random() * 0x7ffffffe) + 1).toString(16).toUpperCase().padStart(8, '0')

// paraIdOf is the id of a comment's last paragraph, which threads replies
// and keeps the resolved flag; it is added when missing.
function paraIdOf (m, c, make) {
  const ps = children(c, 'w:p')
  const last = ps[ps.length - 1]
  if (!last) return undefined
  if (!last.attrs['w14:paraId'] && make) {
    ignorable(commentsPart(m, true).root, 'w14', W14)
    last.attrs['w14:paraId'] = newParaId()
    last.attrs['w14:textId'] = '77777777'
  }
  return last.attrs['w14:paraId']
}

function commentNode (m, id, where) {
  const cp = commentsPart(m)
  const c = cp && children(cp.root, 'w:comment').find((x) => x.attrs['w:id'] === String(id))
  if (!c) throw new Error(`${where}: there is no comment ${id}; reload the document`)
  return c
}

function commentParas (lines) {
  return lines.map((line, i) => el('w:p', { 'w14:paraId': newParaId(), 'w14:textId': '77777777' }, [
    i === 0 ? el('w:r', {}, [el('w:annotationRef')]) : null,
    line ? el('w:r', {}, textNodes(line)) : null
  ]))
}

function nextCommentId (m) {
  const cp = commentsPart(m)
  const ids = [...children(cp?.root, 'w:comment').map((c) => +c.attrs['w:id']), ...descendants(m.body, (n) => CM_NAMES.has(n.name)).map((n) => +n.attrs['w:id'])].filter(Number.isFinite)
  return String(Math.max(-1, ...ids) + 1)
}

// wordDate is the time the way Word stamps comments: local wall-clock
// time, though marked Z, so Word and WPS show the hour the person saw.
function wordDate (d = new Date()) {
  const p = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}Z`
}

function addComment (m, e) {
  const id = nextCommentId(m)
  const author = String(e.author ?? '').trim() || '用户'
  const initials = String(e.initials ?? '').trim() || [...author][0]
  const node = el('w:comment', { 'w:id': id, 'w:author': author, 'w:date': wordDate(), 'w:initials': initials }, commentParas(String(e.text).replace(/\r\n?/g, '\n').split('\n')))
  const cp = commentsPart(m, true)
  cp.root.kids.push(node)
  m.pkg.touch(cp.part)
  return { id, node }
}

function setEx (m, pid, attrs) {
  const ex = commentsExPart(m, true)
  let e = children(ex.root, 'w15:commentEx').find((x) => x.attrs['w15:paraId'] === pid)
  if (!e) { e = el('w15:commentEx', { 'w15:paraId': pid, 'w15:done': '0' }); ex.root.kids.push(e) }
  Object.assign(e.attrs, attrs)
  m.pkg.touch(ex.part)
}

function markersOf (m, id) {
  const out = {}
  for (const n of descendants(m.body, (x) => CM_NAMES.has(x.name) && x.attrs['w:id'] === String(id))) {
    if (n.name === 'w:commentRangeStart') out.s = n
    else if (n.name === 'w:commentRangeEnd') out.e = n
    else out.r = parentOf(m.body, n)
  }
  return out
}

function replyComment (m, e, where) {
  const parent = commentNode(m, e.parent, where)
  if (!String(e.text ?? '').trim()) throw new Error(`${where}: text is required`)
  const { id, node } = addComment(m, e)
  setEx(m, paraIdOf(m, node, true), { 'w15:paraIdParent': paraIdOf(m, parent, true) })
  // The reply shares its parent's place in the text.
  const mk = markersOf(m, e.parent)
  const after = (anchor, made) => { const p = parentOf(m.body, anchor); p.kids.splice(p.kids.indexOf(anchor) + 1, 0, made) }
  if (mk.s) after(mk.s, el('w:commentRangeStart', { 'w:id': id }))
  if (mk.e) after(mk.e, el('w:commentRangeEnd', { 'w:id': id }))
  if (mk.r) after(mk.r, el('w:r', {}, [el('w:commentReference', { 'w:id': id })]))
}

function editComment (m, e, where) {
  const c = commentNode(m, e.id, where)
  if (!String(e.text ?? '').trim()) throw new Error(`${where}: text is required`)
  const pid = paraIdOf(m, c, false)
  const paras = commentParas(String(e.text).replace(/\r\n?/g, '\n').split('\n'))
  if (pid) paras[paras.length - 1].attrs['w14:paraId'] = pid
  ignorable(commentsPart(m, true).root, 'w14', W14)
  c.kids = paras
  m.pkg.touch(commentsPart(m).part)
}

function doneComment (m, e, where) {
  const c = commentNode(m, e.id, where)
  setEx(m, paraIdOf(m, c, true), { 'w15:done': e.done === false ? '0' : '1' })
  m.pkg.touch(commentsPart(m).part)
}

// deleteComments removes comments with their replies, anchors and thread
// entries.
function deleteComments (m, ids, where) {
  const cp = commentsPart(m)
  if (!cp) { if (where) throw new Error(`${where}: the document has no comments`); return }
  const ex = commentsExPart(m)
  const all = children(cp.root, 'w:comment')
  const pidOf = new Map(all.map((c) => [c.attrs['w:id'], paraIdOf(m, c, false)]))
  const parentPid = new Map(children(ex?.root, 'w15:commentEx').filter((x) => x.attrs['w15:paraIdParent']).map((x) => [x.attrs['w15:paraId'], x.attrs['w15:paraIdParent']]))
  const doomed = new Set(ids.map(String))
  if (where && !all.some((c) => doomed.has(c.attrs['w:id']))) throw new Error(`${where}: there is no comment ${ids[0]}; reload the document`)
  for (let grew = true; grew;) {
    grew = false
    for (const c of all) {
      const id = c.attrs['w:id']
      if (doomed.has(id)) continue
      const pp = parentPid.get(pidOf.get(id))
      if (pp && [...doomed].some((d) => pidOf.get(d) === pp)) { doomed.add(id); grew = true }
    }
  }
  cp.root.kids = cp.root.kids.filter((c) => !(c.name === 'w:comment' && doomed.has(c.attrs['w:id'])))
  m.pkg.touch(cp.part)
  if (ex) {
    const pids = new Set([...doomed].map((d) => pidOf.get(d)).filter(Boolean))
    ex.root.kids = ex.root.kids.filter((x) => !(x.name === 'w15:commentEx' && pids.has(x.attrs['w15:paraId'])))
    m.pkg.touch(ex.part)
  }
  for (const n of descendants(m.body, (x) => CM_NAMES.has(x.name) && doomed.has(x.attrs['w:id']))) {
    const target = n.name === 'w:commentReference' ? parentOf(m.body, n) : n
    const p = parentOf(m.body, target)
    if (p) remove(p, target)
  }
}

// settleComments gives new comments their numbers and keeps every comment
// whole: a comment whose reference was deleted goes, a lone start or end
// gets its partner, and anchors of comments that do not exist are dropped.
function settleComments (m, anchored, fresh) {
  const markers = descendants(m.body, (n) => CM_NAMES.has(n.name))
  for (const n of markers) { const real = fresh.get(n.attrs['w:id']); if (real) n.attrs['w:id'] = real }
  const cp = commentsPart(m)
  const known = new Set(children(cp?.root, 'w:comment').map((c) => c.attrs['w:id']))
  const groups = new Map()
  for (const n of markers) {
    const g = groups.get(n.attrs['w:id']) ?? {}
    if (n.name === 'w:commentRangeStart') g.s = n; else if (n.name === 'w:commentRangeEnd') g.e = n; else g.r = n
    groups.set(n.attrs['w:id'], g)
  }
  const drop = (n) => { const t = n.name === 'w:commentReference' ? parentOf(m.body, n) : n; const p = t && parentOf(m.body, t); if (p) remove(p, t) }
  for (const [id, g] of groups) if (!known.has(id)) for (const n of [g.s, g.e, g.r]) if (n) drop(n)
  const lost = []
  for (const id of known) {
    const g = groups.get(id)
    const isNew = [...fresh.values()].includes(id)
    if (!g?.r) { if (anchored.has(id) || isNew) lost.push(id); continue }
    if (g.s && !g.e) { const p = parentOf(m.body, g.s); p.kids.splice(p.kids.indexOf(g.s) + 1, 0, el('w:commentRangeEnd', { 'w:id': id })) }
    if (g.e && !g.s) { const p = parentOf(m.body, g.e); p.kids.splice(p.kids.indexOf(g.e), 0, el('w:commentRangeStart', { 'w:id': id })) }
  }
  if (lost.length) deleteComments(m, lost)
}

// readComments lists comments with their thread, state and the text they
// are anchored to.
function readComments (m) {
  const cp = commentsPart(m)
  if (!cp) return []
  const ex = commentsExPart(m)
  const exBy = new Map(children(ex?.root, 'w15:commentEx').map((x) => [x.attrs['w15:paraId'], x]))
  const quotes = new Map()
  const open = new Set()
  for (const n of descendants(m.body, (x) => x.name === 'w:commentRangeStart' || x.name === 'w:commentRangeEnd' || x.name === 'w:t' || x.name === 'w:tab')) {
    if (n.name === 'w:commentRangeStart') { open.add(n.attrs['w:id']); quotes.set(n.attrs['w:id'], quotes.get(n.attrs['w:id']) ?? '') } else if (n.name === 'w:commentRangeEnd') open.delete(n.attrs['w:id'])
    else { const s = n.name === 'w:tab' ? ' ' : textContent(n); for (const id of open) quotes.set(id, quotes.get(id) + s) }
  }
  const byPid = new Map()
  const list = children(cp.root, 'w:comment').map((c) => {
    const pid = paraIdOf(m, c, false)
    const x = pid && exBy.get(pid)
    const item = {
      id: c.attrs['w:id'],
      author: c.attrs['w:author'] ?? '',
      date: c.attrs['w:date'] ?? '',
      text: children(c, 'w:p').map((p) => charMap(p).text).join('\n'),
      quote: (quotes.get(c.attrs['w:id']) ?? '').slice(0, 120),
      done: x?.attrs['w15:done'] === '1',
      parentPid: x?.attrs['w15:paraIdParent']
    }
    if (pid) byPid.set(pid, item)
    return item
  })
  for (const it of list) { if (it.parentPid) it.parent = byPid.get(it.parentPid)?.id; delete it.parentPid }
  return list
}

// ------------------------------------------------------- table of contents

const TOC_GALLERY = 'Table of Contents'

function tocBlocks (m) {
  return descendants(m.body, (n) => n.name === 'w:sdt' && xpath(n, 'w:sdtPr', 'w:docPartObj', 'w:docPartGallery')?.attrs['w:val'] === TOC_GALLERY)
}

function tocLevel (m, p) {
  const st = m.style(p)
  const r = /^toc\s*(\d)$/i.exec(st?.name ?? '')
  return r ? +r[1] : 0
}

// ensureTocStyles adds Word's "toc 1…3" and "TOC Heading" paragraph
// styles when the document lacks them, and returns their ids.
function ensureTocStyles (m) {
  const base = ensureStyles(m.pkg, m.main)
  const part = stylesPart(m.pkg, m.main)
  const d = m.pkg.xml(part)
  const byName = new Map(children(d.root, 'w:style').map((s) => [child(s, 'w:name')?.attrs['w:val']?.toLowerCase(), s.attrs['w:styleId']]))
  const ids = new Set(children(d.root, 'w:style').map((s) => s.attrs['w:styleId']))
  const out = {}
  const want = [['toc heading', 'TOCHeading', () => [v('w:basedOn', base.h1), v('w:next', base.normal), v('w:uiPriority', 39), el('w:unhideWhenUsed'), el('w:qFormat'), pPr(el('w:outlineLvl', { 'w:val': '9' }))]],
    ...[1, 2, 3].map((n) => [`toc ${n}`, `TOC${n}`, () => [v('w:basedOn', base.normal), v('w:next', base.normal), el('w:autoRedefine'), v('w:uiPriority', 39), el('w:unhideWhenUsed'), pPr(el('w:spacing', { 'w:after': '100' }), el('w:ind', { 'w:left': String((n - 1) * 420) }))]])]
  for (const [name, id0, body] of want) {
    let id = byName.get(name)
    if (!id) {
      id = id0
      for (let k = 2; ids.has(id); k++) id = id0 + k
      ids.add(id)
      d.root.kids.push(el('w:style', { 'w:type': 'paragraph', 'w:styleId': id }, [v('w:name', name === 'toc heading' ? 'TOC Heading' : name), ...body()]))
      m.pkg.touch(part)
    }
    out[name] = id
  }
  return out
}

// tocEntries lists the headings to level, giving each a _Toc bookmark the
// entries link to.
function tocEntries (m, levels, skip) {
  let bm = Math.max(0, ...descendants(m.body, 'w:bookmarkStart').map((b) => +b.attrs['w:id'] || 0))
  const out = []
  for (const b of m.blocks) {
    if (b.kind !== 'p' || skip.has(b.owner)) continue
    const lvl = m.headingLevel(b.node)
    const label = charMap(b.node).text.replace(/\s+/g, ' ').trim()
    if (!lvl || lvl > levels || !label) continue
    let mark = children(b.node, 'w:bookmarkStart').find((x) => /^_Toc/.test(x.attrs['w:name'] ?? ''))
    if (!mark) {
      const id = String(++bm)
      mark = el('w:bookmarkStart', { 'w:id': id, 'w:name': '_Toc' + String(Math.floor(Math.random() * 9e8) + 1e8) })
      const ppr = child(b.node, 'w:pPr')
      b.node.kids.splice(ppr ? b.node.kids.indexOf(ppr) + 1 : 0, 0, mark)
      b.node.kids.push(el('w:bookmarkEnd', { 'w:id': id }))
    }
    out.push({ lvl, label, anchor: mark.attrs['w:name'] })
  }
  return out
}

// tocContent is the field Word fills in: entries link to their headings;
// page numbers come when Word or WPS updates the field (it is marked
// dirty, so Word offers to on opening).
function tocContent (m, levels, skip) {
  const st = ensureTocStyles(m)
  const entries = tocEntries(m, levels, skip)
  const begin = [el('w:r', {}, [el('w:fldChar', { 'w:fldCharType': 'begin', 'w:dirty': 'true' })]), el('w:r', {}, [el('w:instrText', { 'xml:space': 'preserve' }, [text(` TOC \\o "1-${levels}" \\h \\z \\u `)])]), el('w:r', {}, [el('w:fldChar', { 'w:fldCharType': 'separate' })])]
  const end = el('w:r', {}, [el('w:fldChar', { 'w:fldCharType': 'end' })])
  const title = el('w:p', {}, [pPr(v('w:pStyle', st['toc heading'])), el('w:r', {}, textNodes('目录'))])
  const paras = entries.length
    ? entries.map((x, i) => el('w:p', {}, [pPr(v('w:pStyle', st[`toc ${Math.min(3, x.lvl)}`])), ...(i === 0 ? begin : []), el('w:hyperlink', { 'w:anchor': x.anchor, 'w:history': '1' }, [el('w:r', {}, textNodes(x.label))]), ...(i === entries.length - 1 ? [end] : [])]))
    : [el('w:p', {}, [...begin, el('w:r', {}, textNodes('未找到目录项：先给标题设上“标题 1/2/3”样式。')), end])]
  return [title, ...paras]
}

function buildToc (m, levels) {
  if (tocBlocks(m).length) throw new Error('the document already has a table of contents; update it instead')
  levels = Math.max(1, Math.min(9, levels))
  return el('w:sdt', {}, [el('w:sdtPr', {}, [el('w:docPartObj', {}, [v('w:docPartGallery', TOC_GALLERY), el('w:docPartUnique')])]), el('w:sdtContent', {}, tocContent(m, levels, new Set()))])
}

function updateToc (m, levels, where) {
  const all = tocBlocks(m)
  if (all.length) {
    const sc = child(all[0], 'w:sdtContent')
    const instr = descendants(sc, 'w:instrText').map(textContent).join('')
    const lv = /\\o\s+"1-(\d)"/.exec(instr)
    sc.kids = tocContent(m, lv ? +lv[1] : levels, new Set([sc]))
    return
  }
  // A TOC field outside a content control: replace its paragraphs.
  const ps = m.blocks.filter((b) => b.kind === 'p').map((b) => b.node)
  const first = ps.findIndex((p) => descendants(p, (n) => n.name === 'w:instrText' && /^\s*TOC\b/.test(textContent(n))).length)
  if (first < 0) throw new Error(`${where}: the document has no table of contents`)
  let depth = 0; let last = first
  for (let i = first; i < ps.length; i++) {
    for (const f of descendants(ps[i], 'w:fldChar')) { if (f.attrs['w:fldCharType'] === 'begin') depth++; else if (f.attrs['w:fldCharType'] === 'end') depth-- }
    last = i
    if (depth <= 0) break
  }
  const owner = parentOf(m.body, ps[first])
  const at = owner.kids.indexOf(ps[first])
  const skip = new Set()
  for (let i = first; i <= last; i++) remove(parentOf(m.body, ps[i]), ps[i])
  const made = tocContent(m, levels, skip).slice(1)
  owner.kids.splice(at, 0, ...made)
}

const CT_HF = { header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml', footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' }
const HF_LOCKED = (n) => LOCKED_CELL.has(n.name) || n.name === 'w:drawing' || n.name === 'w:pict' || n.name === 'w:tbl'

function bodySect (m, create) {
  let sect = child(m.body, 'w:sectPr')
  if (!sect && create) { sect = el('w:sectPr'); m.body.kids.push(sect) }
  return sect
}

// headerFooterOf finds the default header or footer of the last section.
function headerFooterOf (m, kind) {
  const ref = children(bodySect(m) ?? el('x'), `w:${kind}Reference`).find((r) => (r.attrs['w:type'] ?? 'default') === 'default')
  const part = ref && m.pkg.rel(m.main, ref.attrs['r:id'])?.part
  return part && m.pkg.has(part) ? { part, root: m.pkg.xml(part).root } : undefined
}

function hfInfo (m, kind) {
  const hf = headerFooterOf(m, kind)
  if (!hf) return { exists: false, text: '', editable: true }
  const paras = children(hf.root, 'w:p')
  const page = descendants(hf.root, (n) => (n.name === 'w:fldSimple' && /\bPAGE\b/.test(n.attrs['w:instr'] ?? '')) || (n.name === 'w:instrText' && /\bPAGE\b/.test(textContent(n)))).length > 0
  const editable = !descendants(hf.root, HF_LOCKED).length && hf.root.kids.filter(isEl).every((k) => k.name === 'w:p')
  return { exists: true, text: paras.map((p) => charMap(p).text).join('\n'), editable, page }
}

// setHeaderFooter writes the default header or footer as plain lines
// (text null: a centred page number), creating the part when there is
// none. Existing paragraphs lend their formatting to the new lines.
function setHeaderFooter (m, kind, textValue, where) {
  let hf = headerFooterOf(m, kind)
  if (!hf) {
    const part = m.pkg.uniqueName(`word/${kind}#.xml`)
    m.pkg.setXml(part, doc(el(kind === 'header' ? 'w:hdr' : 'w:ftr', { 'xmlns:w': W, 'xmlns:r': NS.r, ...DRAW_NS }, [])))
    m.pkg.setOverride(part, CT_HF[kind])
    const rid = m.pkg.addRel(m.main, REL(kind), part)
    if (!m.d.root.attrs['xmlns:r']) m.d.root.attrs['xmlns:r'] = NS.r
    const sect = bodySect(m, true)
    const refs = sect.kids.filter((k) => k.name === 'w:headerReference' || k.name === 'w:footerReference')
    const node = el(`w:${kind}Reference`, { 'w:type': 'default', 'r:id': rid })
    const after = kind === 'header' ? refs.filter((k) => k.name === 'w:headerReference').pop() : refs[refs.length - 1]
    sect.kids.splice(after ? sect.kids.indexOf(after) + 1 : 0, 0, node)
    hf = { part, root: m.pkg.xml(part).root }
  } else if (textValue !== null && descendants(hf.root, HF_LOCKED).length && textValue !== '') {
    throw new Error(`${where}: 这个${kind === 'header' ? '页眉' : '页脚'}含页码、图片或表格，请在 Word/WPS 中编辑，或先清空`)
  }
  const first = children(hf.root, 'w:p')[0]
  const ppr = first && child(first, 'w:pPr') ? clone(child(first, 'w:pPr')) : el('w:pPr', {}, [v('w:jc', 'center')])
  const firstRun = first && descendants(first, 'w:r').find((r) => child(r, 'w:rPr'))
  const rpr = () => (firstRun ? clone(child(firstRun, 'w:rPr')) : null)
  let paras
  if (textValue === null) {
    const fld = (instr, shown) => el('w:fldSimple', { 'w:instr': ` ${instr} ` }, [el('w:r', {}, [rpr(), el('w:t', {}, [text(shown)])])])
    const run = (s) => el('w:r', {}, [rpr(), el('w:t', { 'xml:space': 'preserve' }, [text(s)])])
    const pp = clone(ppr)
    pp.kids = pp.kids.filter((x) => x.name !== 'w:jc')
    insertOrdered(pp, v('w:jc', 'center'), PPR_ORDER)
    paras = [el('w:p', {}, [pp, run('第 '), fld('PAGE', '1'), run(' 页，共 '), fld('NUMPAGES', '1'), run(' 页')])]
  } else {
    paras = textValue.split('\n').map((line) => el('w:p', {}, [clone(ppr), line ? el('w:r', {}, [rpr(), ...textNodes(line)]) : null]))
  }
  hf.root.kids = [...hf.root.kids.filter((k) => isEl(k) && k.name !== 'w:p' && !HF_LOCKED(k) && k.name !== 'w:sdt'), ...paras]
  m.pkg.touch(hf.part)
}

const PAGE_SIZES = { A4: [11906, 16838], A3: [16838, 23811], A5: [8391, 11906], B5: [9979, 14175], Letter: [12240, 15840], Legal: [12240, 20160] }
const MARGINS = { normal: [1440, 1800], narrow: [720, 720], moderate: [1440, 1080], wide: [1440, 2880] }

// setPage changes the last section's paper size, orientation and margins.
function setPage (m, e, where) {
  const sect = bodySect(m, true)
  const order = ['w:headerReference', 'w:footerReference', 'w:footnotePr', 'w:endnotePr', 'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc', 'w:pgBorders', 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid', 'w:printerSettings', 'w:sectPrChange']
  let sz = child(sect, 'w:pgSz')
  if (!sz) { sz = el('w:pgSz', { 'w:w': '11906', 'w:h': '16838' }); insertOrdered(sect, sz, order) }
  let w = +sz.attrs['w:w'] || 11906; let h = +sz.attrs['w:h'] || 16838
  if (e.size) {
    const dims = PAGE_SIZES[e.size]
    if (!dims) throw new Error(`${where}: size must be one of ${Object.keys(PAGE_SIZES).join(', ')}`)
    ;[w, h] = dims
    if (sz.attrs['w:orient'] === 'landscape') [w, h] = [h, w]
  }
  if (e.orient) {
    const land = e.orient === 'landscape'
    if (land !== w > h) [w, h] = [h, w]
    if (land) sz.attrs['w:orient'] = 'landscape'; else delete sz.attrs['w:orient']
  }
  sz.attrs['w:w'] = String(w); sz.attrs['w:h'] = String(h)
  if (e.cols !== undefined) {
    const n = Math.max(1, Math.min(4, +e.cols || 1))
    let cols = child(sect, 'w:cols')
    if (!cols) { cols = el('w:cols'); insertOrdered(sect, cols, order) }
    cols.kids = []
    if (n > 1) Object.assign(cols.attrs, { 'w:num': String(n), 'w:space': '425' }); else { delete cols.attrs['w:num']; cols.attrs['w:space'] = cols.attrs['w:space'] ?? '425' }
    delete cols.attrs['w:equalWidth']
    if (e.sep !== undefined) { if (e.sep) cols.attrs['w:sep'] = '1'; else delete cols.attrs['w:sep'] }
  }
  if (e.margin) {
    const mg = MARGINS[e.margin]
    if (!mg) throw new Error(`${where}: margin must be one of ${Object.keys(MARGINS).join(', ')}`)
    let mar = child(sect, 'w:pgMar')
    if (!mar) { mar = el('w:pgMar', { 'w:header': '851', 'w:footer': '992', 'w:gutter': '0' }); insertOrdered(sect, mar, order) }
    Object.assign(mar.attrs, { 'w:top': String(mg[0]), 'w:bottom': String(mg[0]), 'w:left': String(mg[1]), 'w:right': String(mg[1]) })
  }
}

// prevList finds the list the paragraph before block i belongs to, so a
// paragraph turned into a list item joins it.
function prevList (m, i) {
  const b = i > 0 ? m.blocks[i - 1] : undefined
  if (b?.kind !== 'p') return undefined
  const np = xpath(b.node, 'w:pPr', 'w:numPr')
  const li = m.list(b.node)
  return np && li ? { numId: child(np, 'w:numId')?.attrs['w:val'], ordered: li.ordered } : undefined
}

function writePara (m, W_, p, runs, props, prev, like, opt = {}) {
  const map = charMap(p)
  const likeMap = like ? charMap(like) : undefined
  const markPr = child(child(like ?? p, 'w:pPr'), 'w:rPr')
  const newText = runs.filter((r) => r.t === 'text').map((r) => r.text).join('')
  const old = map.text
  let pre = 0
  while (pre < old.length && pre < newText.length && old[pre] === newText[pre]) pre++
  let suf = 0
  while (suf < old.length - pre && suf < newText.length - pre && old[old.length - 1 - suf] === newText[newText.length - 1 - suf]) suf++
  const ownerAt = (j) => {
    if (!old.length) return undefined
    if (j < pre) return map.owners[j]
    if (j >= newText.length - suf) return map.owners[j - newText.length + old.length]
    return map.owners[Math.max(0, pre - 1)] ?? map.owners[0]
  }
  const baseFor = (j) => {
    const o = ownerAt(j) ?? (likeMap?.owners.length ? likeMap.owners[likeMap.owners.length - 1] : undefined)
    const pr = o ? child(o, 'w:rPr') : markPr
    const c = pr ? clone(pr) : el('w:rPr')
    c.name = 'w:rPr'
    c.kids = c.kids.filter((x) => !['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo', 'w:rPrChange'].includes(x.name))
    return c
  }
  const hyperlinks = new Map()
  const linkFor = (url) => {
    if (hyperlinks.has(url)) return hyperlinks.get(url)
    let attrs
    if (url.startsWith('#')) attrs = { 'w:anchor': url.slice(1) }
    else {
      const rel = m.pkg.rels(m.main).find((r) => r.target === url && /hyperlink$/.test(r.type))
      attrs = { 'r:id': rel ? rel.id : m.pkg.addRel(m.main, NS.hyperlink, url, true), 'w:history': '1' }
    }
    hyperlinks.set(url, attrs)
    return attrs
  }
  // Revisions stay in the file until accepted or rejected, even where the
  // editor has not shown them yet: unchanged text keeps its insertion,
  // and every deletion keeps its place.
  if (map.insOf.size || map.dels.size) {
    const out = []
    let at = 0
    for (const r of runs) {
      if (r.t !== 'text' || !r.text || r.rev) { out.push(r); if (r.t === 'text') at += r.text?.length ?? 0; continue }
      let cur = null
      for (let i = 0; i < r.text.length; i++) {
        const j = at + i
        const oj = j < pre ? j : j >= newText.length - suf ? j - newText.length + old.length : -1
        const o = oj >= 0 ? map.owners[oj] : undefined
        const ins = o && map.insOf.get(o)
        const rev = ins ? 'ins:' + ins.attrs['w:id'] : undefined
        if (cur && cur.rev === rev) cur.text += r.text[i]
        else { cur = { ...r, text: r.text[i], ...(rev ? { rev } : {}) }; out.push(cur) }
      }
      at += r.text.length
    }
    runs = out
    const sent = new Set(runs.filter((r) => r.t === 'del').map((r) => String(r.id)))
    for (const [id, o] of map.delAt) {
      if (sent.has(id)) continue
      const want = o <= pre ? o : o >= old.length - suf ? o + newText.length - old.length : pre
      let pos = 0
      let ix = 0
      for (; ix < runs.length; ix++) {
        const r = runs[ix]
        if (r.t !== 'text') continue
        if (pos + r.text.length > want) {
          if (want > pos) { const cut = want - pos; runs.splice(ix, 1, { ...r, text: r.text.slice(0, cut) }, { ...r, text: r.text.slice(cut) }); ix++ }
          break
        }
        pos += r.text.length
      }
      runs.splice(ix, 0, { t: 'del', id })
    }
  }
  // Tracking changes, new text goes in as an insertion and text taken
  // out of the file's own runs stays as a deletion.
  if (opt.track && (opt.fresh || newText !== old)) {
    const a = opt.fresh ? 0 : pre
    const z = opt.fresh ? newText.length : newText.length - suf
    const split = []
    let at = 0
    for (const r of runs) {
      if (r.t !== 'text' || !r.text) { split.push(r); continue }
      const cuts = [0, Math.max(0, Math.min(r.text.length, a - at)), Math.max(0, Math.min(r.text.length, z - at)), r.text.length]
      for (let i = 0; i < 3; i++) {
        const t = r.text.slice(cuts[i], cuts[i + 1])
        if (t) split.push(i === 1 ? { ...r, text: t, rev: r.rev ?? 'ins:new' } : { ...r, text: t })
      }
      at += r.text.length
    }
    runs = split
    if (!opt.fresh) {
      const gone = []
      for (let j = pre; j < old.length - suf; j++) {
        const o = map.owners[j]
        if (map.insOf.has(o)) continue
        const last = gone[gone.length - 1]
        if (last && last.owner === o) last.text += old[j]; else gone.push({ owner: o, text: old[j] })
      }
      if (gone.length) {
        const del = el('w:del', revAttrs(m, opt), gone.map((g) => {
          const pr = child(g.owner, 'w:rPr')
          return el('w:r', {}, [pr ? clone(pr) : null, ...textNodes(g.text).map((t) => { t.name = t.name === 'w:t' ? 'w:delText' : t.name; return t })])
        }))
        let ix = 0; let seen = 0
        while (ix < runs.length && !(runs[ix].t === 'text' && seen >= a)) { if (runs[ix].t === 'text') seen += runs[ix].text.length; ix++ }
        runs.splice(ix, 0, { t: '_node', node: del })
      }
    }
  }
  const out = []
  let pos = 0
  let curLink
  let curIns
  const newIns = new Map()
  for (const r of runs) {
    let node
    if (r.t === 'text') {
      if (!r.text) continue
      const rpr = enforce(baseFor(pos), r)
      pos += r.text.length
      node = el('w:r', {}, [rpr.kids.length ? rpr : null, ...textNodes(r.text)])
      if (r.rev) {
        // Runs of one insertion share its w:ins.
        if (curIns?.rev === r.rev && (!r.link || curIns.link === r.link)) { (curIns.link ? curIns.h : curIns.node).kids.push(node); continue }
        const id = r.rev.slice(4)
        const had = map.insEls.get(id)
        let attrs
        if (had) attrs = { ...had.attrs }
        else { if (!newIns.has(r.rev)) newIns.set(r.rev, revAttrs(m, opt)); attrs = newIns.get(r.rev) }
        const ins = el('w:ins', attrs, [])
        if (r.link) { const h = el('w:hyperlink', linkFor(r.link), [node]); ins.kids.push(h); curIns = { rev: r.rev, link: r.link, node: ins, h } } else { ins.kids.push(node); curIns = { rev: r.rev, node: ins } }
        curLink = undefined
        out.push(ins)
        continue
      }
      curIns = undefined
      if (r.link) {
        if (curLink?.url === r.link) { curLink.node.kids.push(node); continue }
        const h = el('w:hyperlink', linkFor(r.link), [node])
        curLink = { url: r.link, node: h }
        out.push(h)
        continue
      }
    } else if (r.t === 'img') {
      const d = (r.from !== undefined && m.drawingsOf?.(r.from)?.[r.n]) || (likeMap ?? map).drawings[r.n] || map.drawings[r.n]
      if (!d) continue
      node = clone(d)
      shapeDrawing(node, r)
    } else if (r.t === '_node') node = r.node
    else if (r.t === 'del') {
      const had = map.dels.get(String(r.id ?? ''))
      if (!had) continue
      node = clone(had)
    } else if (r.t === 'fn' || r.t === 'en') {
      const id = String(r.id ?? '')
      if (!id) continue
      const had = map.cm.get(r.t + ':' + id)
      node = had ? clone(had) : el('w:r', {}, [el('w:rPr', {}, [v('w:vertAlign', 'superscript')]), el(r.t === 'fn' ? 'w:footnoteReference' : 'w:endnoteReference', { 'w:id': id })])
    } else if (r.t === 'pb') node = el('w:r', {}, [el('w:br', { 'w:type': 'page' })])
    else if (r.t === 'cs' || r.t === 'ce' || r.t === 'cr') {
      const id = String(r.id ?? '')
      if (!id) continue
      const had = map.cm.get(r.t + ':' + id)
      node = had ? clone(had) : r.t === 'cr' ? el('w:r', {}, [el('w:commentReference', { 'w:id': id })]) : el(r.t === 'cs' ? 'w:commentRangeStart' : 'w:commentRangeEnd', { 'w:id': id })
    } else continue
    curLink = undefined
    curIns = undefined
    out.push(node)
  }
  let ppr = child(p, 'w:pPr')
  if (opt.track && opt.fresh) {
    if (!ppr) { ppr = el('w:pPr'); p.kids.unshift(ppr) }
    let mark = child(ppr, 'w:rPr')
    if (!mark) mark = insertOrdered(ppr, el('w:rPr'), PPR_ORDER)
    if (!child(mark, 'w:ins')) mark.kids.unshift(el('w:ins', revAttrs(m, opt)))
  }
  p.kids = [ppr, ...map.marks.start, ...out, ...map.marks.end].filter(Boolean)
  if (!props) return
  const ensurePPr = () => { if (!ppr) { ppr = el('w:pPr'); p.kids.unshift(ppr) } return ppr }
  if (props.style !== undefined) {
    const st = ensurePPr()
    st.kids = st.kids.filter((x) => x.name !== 'w:pStyle')
    let id = props.style
    if (typeof id === 'string' && id.startsWith('@')) id = W_().styles[id.slice(1)]
    if (id) insertOrdered(st, v('w:pStyle', id), PPR_ORDER)
    // A heading set by style drops a direct outline level that would fight it.
    st.kids = st.kids.filter((x) => x.name !== 'w:outlineLvl')
  }
  if (props.align !== undefined) {
    const st = ensurePPr()
    st.kids = st.kids.filter((x) => x.name !== 'w:jc')
    if (props.align) insertOrdered(st, v('w:jc', props.align), PPR_ORDER)
  }
  if (props.spacing !== undefined) {
    const st = ensurePPr()
    let sp = child(st, 'w:spacing')
    if (!sp) { sp = el('w:spacing'); insertOrdered(st, sp, PPR_ORDER) }
    const s = props.spacing ?? { line: null, before: null, after: null }
    if (s.line !== undefined) { if (s.line === null) { delete sp.attrs['w:line']; delete sp.attrs['w:lineRule'] } else { sp.attrs['w:line'] = String(Math.round(s.line)); sp.attrs['w:lineRule'] = 'auto' } }
    for (const k of ['before', 'after']) if (s[k] !== undefined) { delete sp.attrs[`w:${k}Autospacing`]; delete sp.attrs[`w:${k}Lines`]; if (s[k] === null) delete sp.attrs['w:' + k]; else sp.attrs['w:' + k] = String(Math.max(0, Math.round(s[k]))) }
    if (!Object.keys(sp.attrs).length) remove(st, sp)
  }
  if (props.indent !== undefined) {
    const st = ensurePPr()
    let ind = child(st, 'w:ind')
    if (!ind) { ind = el('w:ind'); insertOrdered(st, ind, PPR_ORDER) }
    const s = props.indent ?? { left: null, firstLine: null }
    if (s.left !== undefined) { delete ind.attrs['w:start']; delete ind.attrs['w:leftChars']; delete ind.attrs['w:startChars']; if (s.left === null) delete ind.attrs['w:left']; else ind.attrs['w:left'] = String(Math.max(0, Math.round(s.left))) }
    if (s.firstLine !== undefined) {
      for (const k of ['w:firstLine', 'w:hanging', 'w:firstLineChars', 'w:hangingChars']) delete ind.attrs[k]
      if (s.firstLine) { if (s.firstLine > 0) ind.attrs['w:firstLine'] = String(Math.round(s.firstLine)); else ind.attrs['w:hanging'] = String(Math.round(-s.firstLine)) }
    }
    if (!Object.keys(ind.attrs).length) remove(st, ind)
  }
  if (props.list !== undefined) {
    const st = ensurePPr()
    const cur = child(st, 'w:numPr')
    if (props.list === null) {
      if (cur) remove(st, cur)
      // A style that numbers paragraphs is switched off directly.
      if (m.list(p)) insertOrdered(st, el('w:numPr', {}, [v('w:ilvl', 0), v('w:numId', 0)]), PPR_ORDER)
    } else {
      const level = Math.max(0, Math.min(8, +props.list.level || 0))
      let numId = props.list.numId
      if (!numId) {
        const curId = cur && child(cur, 'w:numId')?.attrs['w:val']
        const curLi = m.list(p)
        if (curId && curId !== '0' && curLi && curLi.ordered === !!props.list.ordered) numId = curId
        else if (prev && prev.ordered === !!props.list.ordered) numId = prev.numId
        else numId = props.list.ordered ? W_().numbering.numbered(Date.now() % 1e6) : W_().numbering.bullet()
      }
      if (cur) remove(st, cur)
      insertOrdered(st, el('w:numPr', {}, [v('w:ilvl', level), v('w:numId', numId)]), PPR_ORDER)
    }
  }
  if (props.border !== undefined) {
    const st = ensurePPr()
    setPr(st, 'w:pBdr', null)
    const sides = { t: 'top', l: 'left', b: 'bottom', r: 'right' }
    const kids = Object.entries(sides).filter(([k]) => (props.border ?? '').includes(k)).map(([, side]) => el('w:' + side, { 'w:val': 'single', 'w:sz': '4', 'w:space': side === 'left' || side === 'right' ? '4' : '1', 'w:color': 'auto' }))
    if (kids.length) insertOrdered(st, el('w:pBdr', {}, kids), PPR_ORDER)
  }
  if (props.shade !== undefined) {
    const st = ensurePPr()
    setPr(st, 'w:shd', /^[0-9A-F]{6}$/i.test(props.shade ?? '') ? el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': props.shade.toUpperCase() }) : null, PPR_ORDER)
  }
  if (ppr && !ppr.kids.length) remove(p, ppr)
}

// paraBorder names a paragraph's bordered sides: t, l, b, r.
function paraBorder (p) {
  const box = xpath(p, 'w:pPr', 'w:pBdr')
  if (!box) return ''
  return [['t', 'top'], ['l', 'left'], ['b', 'bottom'], ['r', 'right']].filter(([, s]) => { const b = child(box, 'w:' + s) ?? (s === 'left' ? child(box, 'w:start') : s === 'right' ? child(box, 'w:end') : undefined); return b && !['nil', 'none'].includes(b.attrs['w:val']) }).map(([k]) => k).join('')
}

// setDropCap makes the paragraph's first letter a drop cap of lines
// lines, as Word does: a framed paragraph of its own before it. lines 0
// puts the letter back.
function setDropCap (m, b, lines, where) {
  const owner = b.owner
  const isDrop = (p) => ['drop', 'margin'].includes(xpath(p, 'w:pPr', 'w:framePr')?.attrs['w:dropCap'])
  if (!lines) {
    if (!isDrop(b.node)) throw new Error(`${where}: ${b.ref} 不是首字下沉`)
    const at = owner.kids.indexOf(b.node)
    const next = owner.kids.slice(at + 1).find((k) => k.name === 'w:p')
    const runs = children(b.node, 'w:r')
    remove(owner, b.node)
    if (next) {
      for (const r of runs) { const pr = child(r, 'w:rPr'); if (pr) pr.kids = pr.kids.filter((k) => !['w:sz', 'w:szCs', 'w:position'].includes(k.name)) }
      const ppr = child(next, 'w:pPr')
      next.kids.splice(ppr ? next.kids.indexOf(ppr) + 1 : 0, 0, ...runs)
    }
    return
  }
  lines = Math.max(2, Math.min(10, lines | 0))
  if (isDrop(b.node)) { xpath(b.node, 'w:pPr', 'w:framePr').attrs['w:lines'] = String(lines); return }
  const t = descendants(b.node, 'w:t').find((x) => textContent(x).trim())
  if (!t) throw new Error(`${where}: 这一段没有可下沉的文字`)
  const r = parentOf(b.node, t)
  const str = textContent(t)
  const i = str.search(/\S/)
  const ch = Array.from(str.slice(i))[0]
  t.kids = [text(str.slice(0, i) + str.slice(i + ch.length))]
  if (!textContent(t)) remove(r, t)
  const base = runProps(child(r, "w:rPr")).size ?? styleLook(m, xpath(b.node, "w:pPr", "w:pStyle")?.attrs["w:val"]).size ?? 21
  const rpr = child(r, 'w:rPr') ? clone(child(r, 'w:rPr')) : el('w:rPr')
  rpr.kids = rpr.kids.filter((k) => !['w:sz', 'w:szCs', 'w:position', 'w:rPrChange'].includes(k.name))
  insertOrdered(rpr, v('w:position', -Math.round(base * 0.4)), RPR_ORDER)
  insertOrdered(rpr, v('w:sz', Math.round(base * lines * 1.25)), RPR_ORDER)
  insertOrdered(rpr, v('w:szCs', Math.round(base * lines * 1.25)), RPR_ORDER)
  const ppr = el('w:pPr', {}, [])
  const style = xpath(b.node, 'w:pPr', 'w:pStyle')
  if (style) ppr.kids.push(clone(style))
  insertOrdered(ppr, el('w:framePr', { 'w:dropCap': 'drop', 'w:lines': String(lines), 'w:wrap': 'around', 'w:vAnchor': 'text', 'w:hAnchor': 'text' }), PPR_ORDER)
  insertOrdered(ppr, el('w:spacing', { 'w:before': '0', 'w:after': '0', 'w:line': String(Math.round(base * 10 * 1.3 * lines)), 'w:lineRule': 'exact' }), PPR_ORDER)
  insertOrdered(ppr, el('w:textAlignment', { 'w:val': 'baseline' }), PPR_ORDER)
  owner.kids.splice(owner.kids.indexOf(b.node), 0, el('w:p', {}, [ppr, el('w:r', {}, [rpr, ...textNodes(ch)])]))
}


// ---------------------------------------------------------- table layout
// The editor works on a table in grid coordinates: row (1-based) and
// grid column (0-based), as Word lays merged cells out.

const TCPR_ORDER = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark']
const TBLPR_ORDER = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook']
const SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']

// tableGrid lists each row's cells with the grid column each starts at.
function tableGrid (tbl) {
  return children(tbl, 'w:tr').map((tr) => {
    let g = +(xpath(tr, 'w:trPr', 'w:gridBefore')?.attrs['w:val'] ?? 0)
    return cellsOf(tr).map((tc) => {
      const span = +(xpath(tc, 'w:tcPr', 'w:gridSpan')?.attrs['w:val'] ?? 1)
      const vm = xpath(tc, 'w:tcPr', 'w:vMerge')
      const c = { tc, tr, g, span, vm: vm ? (vm.attrs['w:val'] === 'restart' ? 'restart' : 'continue') : undefined }
      g += span
      return c
    })
  })
}

function tcPrOf (tc) {
  let pr = child(tc, 'w:tcPr')
  if (!pr) { pr = el('w:tcPr'); tc.kids.unshift(pr) }
  return pr
}
function tblPrOf (tbl) {
  let pr = child(tbl, 'w:tblPr')
  if (!pr) { pr = el('w:tblPr'); tbl.kids.unshift(pr) }
  return pr
}
// setPr replaces (or with null removes) a property element.
function setPr (pr, name, node, order) {
  for (const k of children(pr, name)) remove(pr, k)
  if (node) insertOrdered(pr, node, order)
}

function gridWidths (tbl) {
  return children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((g) => +(g.attrs['w:w'] ?? 0))
}
function fitCellWidth (tc, g, span, widths) {
  const w = widths.slice(g, g + span).reduce((a, b) => a + b, 0)
  if (w > 0) setPr(tcPrOf(tc), 'w:tcW', el('w:tcW', { 'w:w': String(w), 'w:type': 'dxa' }), TCPR_ORDER)
}

const hasContent = (n) => descendants(n, (x) => (x.name === 'w:t' && textContent(x) !== '') || x.name === 'w:drawing' || x.name === 'w:pict' || x.name === 'w:tbl').length > 0

// absorb moves src's content into dst, leaving src one empty paragraph.
function absorb (dst, src) {
  const blocks = src.kids.filter((k) => k.name === 'w:p' || k.name === 'w:tbl')
  const keep = blocks.find((k) => k.name === 'w:p')
  const moved = blocks.filter(hasContent)
  src.kids = src.kids.filter((k) => k.name !== 'w:p' && k.name !== 'w:tbl')
  const ppr = keep && child(keep, 'w:pPr')
  src.kids.push(el('w:p', {}, ppr ? [clone(ppr)] : []))
  if (!moved.length) return
  if (!hasContent(dst)) dst.kids = dst.kids.filter((k) => k.name !== 'w:p')
  dst.kids.push(...moved)
}

function blankCell (like) {
  const tc = el('w:tc')
  const pr = child(like, 'w:tcPr')
  if (pr) { const c = clone(pr); c.kids = c.kids.filter((k) => k.name !== 'w:gridSpan' && k.name !== 'w:vMerge' && k.name !== 'w:hMerge'); tc.kids.push(c) }
  const p0 = child(like, 'w:p')
  const ppr = p0 && child(p0, 'w:pPr')
  tc.kids.push(el('w:p', {}, ppr ? [clone(ppr)] : []))
  return tc
}

function cellAt (grid, row, g, where) {
  const c = grid[row - 1]?.find((x) => x.g <= g && g < x.g + x.span)
  if (!c) throw new Error(`${where}: 第 ${row} 行没有第 ${g + 1} 列的单元格`)
  return c
}

// mergeCells merges the rectangle rows r1..r2, grid columns g1..g2 into
// its top-left cell, keeping every cell's content.
function mergeCells (tbl, r1, r2, g1, g2, where) {
  const grid = tableGrid(tbl)
  if (r1 < 1 || r2 > grid.length || r1 > r2 || g1 > g2 || g1 < 0) throw new Error(`${where}: 选区超出表格`)
  const rows = []
  for (let r = r1; r <= r2; r++) {
    const cells = grid[r - 1].filter((c) => c.g + c.span > g1 && c.g <= g2)
    if (!cells.length || cells[0].g !== g1 || cells[cells.length - 1].g + cells[cells.length - 1].span - 1 !== g2) throw new Error(`${where}: 选区不是完整的矩形，不能合并`)
    rows.push(cells)
  }
  if (rows[0].some((c) => c.vm === 'continue')) throw new Error(`${where}: 选区切断了已有的纵向合并，不能合并`)
  if (grid[r2]?.some((c) => c.g + c.span > g1 && c.g <= g2 && c.vm === 'continue')) throw new Error(`${where}: 选区切断了已有的纵向合并，不能合并`)
  if (rows.length === 1 && rows[0].length === 1) return
  const widths = gridWidths(tbl)
  const top = rows[0][0].tc
  rows.forEach((cells, i) => {
    const lead = cells[0]
    if (i > 0) absorb(top, lead.tc)
    for (const c of cells.slice(1)) { absorb(top, c.tc); remove(c.tr.kids.includes(c.tc) ? c.tr : parentOf(c.tr, c.tc), c.tc) }
    const pr = tcPrOf(lead.tc)
    setPr(pr, 'w:gridSpan', g2 > g1 ? v('w:gridSpan', g2 - g1 + 1) : null, TCPR_ORDER)
    setPr(pr, 'w:hMerge', null)
    if (i > 0) setPr(pr, 'w:vMerge', el('w:vMerge'), TCPR_ORDER)
    else setPr(pr, 'w:vMerge', r2 > r1 ? v('w:vMerge', 'restart') : null, TCPR_ORDER)
    fitCellWidth(lead.tc, g1, g2 - g1 + 1, widths)
  })
}

// unmergeCell splits a merged cell back into the grid cells it covers.
function unmergeCell (tbl, row, g, where) {
  const grid = tableGrid(tbl)
  const c = cellAt(grid, row, g, where)
  if (c.vm === 'continue') throw new Error(`${where}: 请在合并区域的第一行操作`)
  const widths = gridWidths(tbl)
  const split = (x) => {
    const pr = tcPrOf(x.tc)
    setPr(pr, 'w:vMerge', null)
    if (x.span > 1) {
      setPr(pr, 'w:gridSpan', null)
      const owner = x.tr.kids.includes(x.tc) ? x.tr : parentOf(x.tr, x.tc)
      const extra = Array.from({ length: x.span - 1 }, (_, i) => { const n = blankCell(x.tc); fitCellWidth(n, x.g + 1 + i, 1, widths); return n })
      owner.kids.splice(owner.kids.indexOf(x.tc) + 1, 0, ...extra)
      fitCellWidth(x.tc, x.g, 1, widths)
    }
  }
  const chain = [c]
  if (c.vm === 'restart') {
    for (let r = row + 1; r <= grid.length; r++) {
      const d = grid[r - 1].find((x) => x.g === c.g)
      if (d?.vm !== 'continue') break
      chain.push(d)
    }
  }
  if (chain.length === 1 && c.span === 1) throw new Error(`${where}: 这个单元格没有合并`)
  for (const x of chain) split(x)
}

// splitCell divides a cell into cols × rows cells.
function splitCell (tbl, row, g, cols, rows, where) {
  let grid = tableGrid(tbl)
  const c = cellAt(grid, row, g, where)
  if (c.span > 1 || c.vm) throw new Error(`${where}: 合并过的单元格请先取消合并`)
  cols = Math.max(1, Math.min(10, cols | 0))
  rows = Math.max(1, Math.min(20, rows | 0))
  let targets = [c.tc]
  if (cols > 1) {
    const tg = child(tbl, 'w:tblGrid')
    const gcs = children(tg, 'w:gridCol')
    const gc = gcs[c.g]
    const w = +(gc?.attrs['w:w'] ?? 0)
    const part = Math.max(1, Math.floor(w / cols))
    if (gc) { gc.attrs['w:w'] = String(part); tg.kids.splice(tg.kids.indexOf(gc) + 1, 0, ...Array.from({ length: cols - 1 }, () => el('w:gridCol', { 'w:w': String(part) }))) }
    grid.forEach((cells, i) => {
      if (i === row - 1) return
      const x = cells.find((d) => d.g <= c.g && c.g < d.g + d.span)
      if (x) { setPr(tcPrOf(x.tc), 'w:gridSpan', v('w:gridSpan', x.span + cols - 1), TCPR_ORDER); fitCellWidth(x.tc, x.g, x.span + cols - 1, gridWidths(tbl)) }
    })
    const owner = c.tr.kids.includes(c.tc) ? c.tr : parentOf(c.tr, c.tc)
    const extra = Array.from({ length: cols - 1 }, () => blankCell(c.tc))
    owner.kids.splice(owner.kids.indexOf(c.tc) + 1, 0, ...extra)
    targets = [c.tc, ...extra]
    const widths = gridWidths(tbl)
    targets.forEach((tc, i) => fitCellWidth(tc, c.g + i, 1, widths))
  }
  if (rows > 1) {
    grid = tableGrid(tbl)
    const tr = c.tr
    const cells = grid[row - 1]
    const trPr = child(tr, 'w:trPr')
    const made = []
    for (let k = 1; k < rows; k++) {
      const nr = el('w:tr', {}, trPr ? [clone(trPr)] : [])
      if (child(nr, 'w:trPr')) child(nr, 'w:trPr').kids = child(nr, 'w:trPr').kids.filter((x) => x.name !== 'w:tblHeader')
      for (const x of cells) {
        const n = blankCell(x.tc)
        if (x.span > 1) setPr(tcPrOf(n), 'w:gridSpan', v('w:gridSpan', x.span), TCPR_ORDER)
        if (!targets.includes(x.tc)) {
          setPr(tcPrOf(n), 'w:vMerge', el('w:vMerge'), TCPR_ORDER)
          if (!x.vm) setPr(tcPrOf(x.tc), 'w:vMerge', v('w:vMerge', 'restart'), TCPR_ORDER)
        }
        nr.kids.push(n)
      }
      made.push(nr)
    }
    tbl.kids.splice(tbl.kids.indexOf(tr) + 1, 0, ...made)
  }
}

// setWidths sets the grid columns' widths (twips) and fixes the layout so
// Word keeps them.
function setWidths (tbl, widths, where) {
  const tg = child(tbl, 'w:tblGrid')
  const gcs = children(tg, 'w:gridCol')
  if (!Array.isArray(widths) || widths.length !== gcs.length) throw new Error(`${where}: 需要 ${gcs.length} 个列宽`)
  widths = widths.map((w) => Math.max(200, Math.min(30000, Math.round(+w || 0))))
  gcs.forEach((g, i) => { g.attrs['w:w'] = String(widths[i]) })
  for (const row of tableGrid(tbl)) for (const c of row) fitCellWidth(c.tc, c.g, c.span, widths)
  const pr = tblPrOf(tbl)
  setPr(pr, 'w:tblW', el('w:tblW', { 'w:w': String(widths.reduce((a, b) => a + b, 0)), 'w:type': 'dxa' }), TBLPR_ORDER)
  setPr(pr, 'w:tblLayout', el('w:tblLayout', { 'w:type': 'fixed' }), TBLPR_ORDER)
}

const borderEl = (side, color) => el('w:' + side, color === 'nil' ? { 'w:val': 'nil' } : { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': color || '000000' })

function setBorders (pr, name, sides, order) {
  let box = child(pr, name)
  if (!box) box = insertOrdered(pr, el(name), order)
  for (const [side, color] of Object.entries(sides)) {
    for (const k of children(box, 'w:' + side)) remove(box, k)
    if (side === 'left') for (const k of children(box, 'w:start')) remove(box, k)
    if (side === 'right') for (const k of children(box, 'w:end')) remove(box, k)
    if (color !== undefined) insertOrdered(box, borderEl(side, color), ['w:top', 'w:left', 'w:start', 'w:bottom', 'w:right', 'w:end', 'w:insideH', 'w:insideV', 'w:tl2br', 'w:tr2bl'])
  }
  if (!box.kids.some(isEl)) remove(pr, box)
}

function cellParas (tc) { return descendants(tc, 'w:p') }

function alignParas (tc, align) {
  for (const p of cellParas(tc)) {
    let ppr = child(p, 'w:pPr')
    if (!ppr) { ppr = el('w:pPr'); p.kids.unshift(ppr) }
    setPr(ppr, 'w:jc', align ? v('w:jc', align) : null, PPR_ORDER)
  }
}

// formatCells applies shading, alignment and borders to the cells of a
// rectangle in grid coordinates.
function formatCells (tbl, e, where) {
  const grid = tableGrid(tbl)
  const r1 = +e.r1; const r2 = +e.r2; const g1 = +e.g1; const g2 = +e.g2
  if (!(r1 >= 1 && r2 <= grid.length && r1 <= r2 && g1 >= 0 && g1 <= g2)) throw new Error(`${where}: 选区超出表格`)
  const picked = []
  for (let r = r1; r <= r2; r++) for (const c of grid[r - 1]) if (c.g + c.span > g1 && c.g <= g2 && c.vm !== 'continue') picked.push({ ...c, r })
  for (const c of picked) {
    const pr = tcPrOf(c.tc)
    if (e.fill !== undefined) setPr(pr, 'w:shd', /^[0-9A-F]{6}$/i.test(e.fill) ? el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': e.fill.toUpperCase() }) : null, TCPR_ORDER)
    if (e.valign !== undefined) setPr(pr, 'w:vAlign', ['center', 'bottom'].includes(e.valign) ? v('w:vAlign', e.valign) : null, TCPR_ORDER)
    if (e.align !== undefined) alignParas(c.tc, ['center', 'right', 'both'].includes(e.align) ? e.align : e.align === 'left' ? 'left' : null)
    if (e.border !== undefined) {
      const right = c.g + c.span - 1
      let bottom = c.r
      if (c.vm === 'restart') while (grid[bottom]?.find((x) => x.g === c.g)?.vm === 'continue') bottom++
      const edge = { top: c.r === r1, bottom: bottom >= r2, left: c.g === g1, right: right >= g2 }
      const col = '000000'
      const s = {}
      for (const side of ['top', 'left', 'bottom', 'right']) {
        if (e.border === 'all') s[side] = col
        else if (e.border === 'none') s[side] = 'nil'
        else if (e.border === 'outer') { if (edge[side]) s[side] = col }
        else if (e.border === 'inner') { if (!edge[side]) s[side] = col }
        else if (e.border === side) { if (edge[side]) s[side] = col }
      }
      if (Object.keys(s).length) setBorders(pr, 'w:tcBorders', s, TCPR_ORDER)
    }
  }
}

const TABLE_PRESETS = {
  grid: { border: '000000', inner: '000000' },
  plain: { border: 'nil', inner: 'nil' },
  light: { border: 'BFBFBF', inner: 'BFBFBF', head: 'D9E2F3', headBold: true },
  banded: { border: '4472C4', inner: 'BFBFBF', vertical: 'nil', head: '4472C4', headColor: 'FFFFFF', headBold: true, band: 'DEEAF6' },
  dark: { border: '8EAADB', inner: '8EAADB', head: '1F3864', headColor: 'FFFFFF', headBold: true }
}

// styleTable applies a look: borders, header row and banded rows, as
// direct formatting so every reader shows it the same.
function styleTable (tbl, name, where) {
  const p = TABLE_PRESETS[name]
  if (!p) throw new Error(`${where}: 未知的表格样式 ${name}`)
  const pr = tblPrOf(tbl)
  setPr(pr, 'w:tblBorders', null)
  setBorders(pr, 'w:tblBorders', { top: p.border, left: p.vertical ?? p.border, bottom: p.border, right: p.vertical ?? p.border, insideH: p.inner, insideV: p.vertical ?? p.inner }, TBLPR_ORDER)
  setPr(pr, 'w:tblLook', el('w:tblLook', { 'w:val': p.head ? '0420' : '0400', 'w:firstRow': p.head ? '1' : '0', 'w:lastRow': '0', 'w:firstColumn': '0', 'w:lastColumn': '0', 'w:noHBand': '1', 'w:noVBand': '1' }), TBLPR_ORDER)
  const grid = tableGrid(tbl)
  grid.forEach((cells, i) => {
    const fill = i === 0 && p.head ? p.head : p.band && i % 2 === 1 ? p.band : undefined
    for (const c of cells) {
      const cp = tcPrOf(c.tc)
      setPr(cp, 'w:tcBorders', null)
      setPr(cp, 'w:shd', fill ? el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': fill }) : null, TCPR_ORDER)
      if (i !== 0) continue
      for (const para of cellParas(c.tc)) {
        const rprs = descendants(para, 'w:r').map((r) => { let rp = child(r, 'w:rPr'); if (!rp) { rp = el('w:rPr'); r.kids.unshift(rp) } return rp })
        let ppr = child(para, 'w:pPr')
        if (!ppr) { ppr = el('w:pPr'); para.kids.unshift(ppr) }
        let mark = child(ppr, 'w:rPr')
        if (!mark) mark = insertOrdered(ppr, el('w:rPr'), PPR_ORDER)
        for (const rp of [...rprs, mark]) {
          const have = runProps(rp)
          enforce(rp, { ...have, b: p.head ? (p.headBold ? true : have.b) : have.b, color: p.headColor ?? (have.color === 'FFFFFF' ? undefined : have.color) })
        }
      }
    }
  })
  const tr0 = children(tbl, 'w:tr')[0]
  if (tr0) {
    let trPr = child(tr0, 'w:trPr')
    if (p.head) { if (!trPr) { trPr = el('w:trPr'); tr0.kids.unshift(trPr) } if (!child(trPr, 'w:tblHeader')) trPr.kids.push(el('w:tblHeader')) } else if (trPr) setPr(trPr, 'w:tblHeader', null)
  }
}

// tableLook describes borders, widths and cell formatting for the editor.
function bordersOf (box) {
  if (!box) return undefined
  const out = {}
  for (const side of SIDES) {
    const b = child(box, 'w:' + side) ?? (side === 'left' ? child(box, 'w:start') : side === 'right' ? child(box, 'w:end') : undefined)
    if (!b) continue
    const val = b.attrs['w:val']
    out[side] = !val || val === 'nil' || val === 'none' ? 'none' : (/^[0-9A-F]{6}$/i.test(b.attrs['w:color'] ?? '') ? b.attrs['w:color'].toUpperCase() : '000000') + (+(b.attrs['w:sz'] ?? 4) >= 12 ? ':2' : '')
  }
  return Object.keys(out).length ? out : undefined
}

function tableStyleBorders (m, tbl) {
  let id = xpath(tbl, 'w:tblPr', 'w:tblStyle')?.attrs['w:val']
  const out = {}
  for (let k = 0; id && k < 6; k++) {
    const st = m.styleInfo.get(id)
    if (!st) break
    const b = bordersOf(xpath(st.node, 'w:tblPr', 'w:tblBorders'))
    for (const [s, val] of Object.entries(b ?? {})) out[s] ??= val
    id = st.basedOn
  }
  return out
}


// ------------------------------------------------------------- revisions

const SETTINGS_TAIL = new Set(['w:doNotTrackMoves', 'w:doNotTrackFormatting', 'w:documentProtection', 'w:autoFormatOverride', 'w:styleLockTheme', 'w:styleLockQFSet', 'w:defaultTabStop', 'w:autoHyphenation', 'w:consecutiveHyphenLimit', 'w:hyphenationZone', 'w:doNotHyphenateCaps', 'w:showEnvelope', 'w:summaryLength', 'w:clickAndTypeStyle', 'w:defaultTableStyle', 'w:evenAndOddHeaders', 'w:bookFoldRevPrinting', 'w:bookFoldPrinting', 'w:bookFoldPrintingSheets', 'w:drawingGridHorizontalSpacing', 'w:drawingGridVerticalSpacing', 'w:displayHorizontalDrawingGridEvery', 'w:displayVerticalDrawingGridEvery', 'w:doNotUseMarginsForDrawingGridOrigin', 'w:drawingGridHorizontalOrigin', 'w:drawingGridVerticalOrigin', 'w:doNotShadeFormData', 'w:noPunctuationKerning', 'w:characterSpacingControl', 'w:printTwoOnOne', 'w:strictFirstAndLastChars', 'w:noLineBreaksAfter', 'w:noLineBreaksBefore', 'w:savePreviewPicture', 'w:doNotValidateAgainstSchema', 'w:saveInvalidXml', 'w:ignoreMixedContent', 'w:alwaysShowPlaceholderText', 'w:doNotDemarcateInvalidXml', 'w:saveXmlDataOnly', 'w:useXSLTWhenSaving', 'w:saveThroughXslt', 'w:showXMLTags', 'w:alwaysMergeEmptyNamespace', 'w:updateFields', 'w:hdrShapeDefaults', 'w:footnotePr', 'w:endnotePr', 'w:compat', 'w:docVars', 'w:rsids', 'm:mathPr', 'w:attachedSchema', 'w:themeFontLang', 'w:clrSchemeMapping', 'w:doNotIncludeSubdocsInStats', 'w:doNotAutoCompressPictures', 'w:forceUpgrade', 'w:captions', 'w:readModeInkLockDown', 'w:smartTagType', 'sl:schemaLibrary', 'w:shapeDefaults', 'w:doNotEmbedSmartTags', 'w:decimalSymbol', 'w:listSeparator'])

function settingsRoot (m, create) {
  let part = m.pkg.relsOfType(m.main, 'settings')[0]?.part
  if (part && m.pkg.has(part)) return { part, root: m.pkg.xml(part).root }
  if (!create) return undefined
  part = 'word/settings.xml'
  m.pkg.setXml(part, doc(el('w:settings', { 'xmlns:w': W }, [])))
  m.pkg.setOverride(part, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
  m.pkg.addRel(m.main, REL('settings'), part)
  return { part, root: m.pkg.xml(part).root }
}

function trackOn (m) { const s = settingsRoot(m, false); return !!s && on(child(s.root, 'w:trackRevisions') ?? null) }

function setTrack (m, yes) {
  const s = settingsRoot(m, yes)
  if (!s) return
  for (const k of children(s.root, 'w:trackRevisions')) remove(s.root, k)
  if (yes) {
    const at = s.root.kids.findIndex((k) => SETTINGS_TAIL.has(k.name))
    if (at < 0) s.root.kids.push(el('w:trackRevisions')); else s.root.kids.splice(at, 0, el('w:trackRevisions'))
  }
  m.pkg.touch(s.part)
}

// revAttrs names a new revision: a fresh id, its author and time.
function revAttrs (m, opt = {}) {
  if (m.revSeq === undefined) m.revSeq = Math.max(0, ...descendants(m.body, (n) => n.attrs?.['w:id'] !== undefined && /^\d+$/.test(n.attrs['w:id'])).map((n) => +n.attrs['w:id'])) + 1
  return { 'w:id': String(m.revSeq++), 'w:author': String(opt.author || 'Quire'), 'w:date': wordDate() }
}

// trackDelete marks a whole paragraph deleted: its text, and its mark so
// it joins the next one once accepted. Text the same person inserted
// and never accepted just goes.
function trackDelete (m, p, opt) {
  const attrs = revAttrs(m, opt)
  const out = []
  let del
  const take = (node) => {
    if (node.name === 'w:ins') { return }
    if (node.name === 'w:del') { del = undefined; out.push(node); return }
    if (node.name === 'w:pPr') { out.push(node); return }
    if (node.name === 'w:r' || node.name === 'w:hyperlink') {
      for (const t of descendants(node, 'w:t')) t.name = 'w:delText'
      for (const t of descendants(node, 'w:instrText')) t.name = 'w:delInstrText'
      if (!del) { del = el('w:del', attrs, []); out.push(del) }
      del.kids.push(...(node.name === 'w:hyperlink' ? node.kids.filter((k) => k.name === 'w:r') : [node]))
      return
    }
    del = undefined
    out.push(node)
  }
  for (const k of [...p.kids]) if (isEl(k)) take(k)
  p.kids = out
  let ppr = child(p, 'w:pPr')
  if (!ppr) { ppr = el('w:pPr'); p.kids.unshift(ppr) }
  let mark = child(ppr, 'w:rPr')
  if (!mark) mark = insertOrdered(ppr, el('w:rPr'), PPR_ORDER)
  for (const k of children(mark, 'w:ins')) remove(mark, k)
  if (!child(mark, 'w:del')) mark.kids.unshift(el('w:del', { ...attrs }))
}

// resolveRevisions accepts or rejects tracked changes under scope, all of
// them or the one with id.
function resolveRevisions (m, scope, action, id) {
  const pick = (n) => id === undefined || n.attrs['w:id'] === id
  const replaceWith = (n, kids) => { const parent = parentOf(m.body, n) ?? parentOf(scope, n); if (!parent) return; parent.kids.splice(parent.kids.indexOf(n), 1, ...kids) }
  const accept = action === 'accept'
  for (const n of descendants(scope, (x) => (x.name === 'w:ins' || x.name === 'w:moveTo') && pick(x))) {
    const parent = parentOf(m.body, n) ?? parentOf(scope, n)
    if (parent?.name === 'w:rPr') continue
    replaceWith(n, accept ? n.kids : [])
  }
  for (const n of descendants(scope, (x) => (x.name === 'w:del' || x.name === 'w:moveFrom') && pick(x))) {
    const parent = parentOf(m.body, n) ?? parentOf(scope, n)
    if (parent?.name === 'w:rPr') continue
    if (accept) replaceWith(n, [])
    else {
      for (const t of descendants(n, 'w:delText')) t.name = 'w:t'
      for (const t of descendants(n, 'w:delInstrText')) t.name = 'w:instrText'
      replaceWith(n, n.kids)
    }
  }
  for (const n of descendants(scope, (x) => x.name === 'w:rPrChange' && pick(x))) {
    const pr = parentOf(m.body, n) ?? parentOf(scope, n)
    if (!pr) continue
    if (accept) remove(pr, n)
    else { const was = child(n, 'w:rPr'); pr.kids = was ? was.kids.map(clone) : [] }
  }
  for (const n of descendants(scope, (x) => x.name === 'w:pPrChange' && pick(x))) {
    const pr = parentOf(m.body, n) ?? parentOf(scope, n)
    if (!pr) continue
    if (accept) remove(pr, n)
    else { const was = child(n, 'w:pPr'); const keep = pr.kids.filter((k) => k.name === 'w:rPr' || k.name === 'w:sectPr'); pr.kids = [...(was ? was.kids.map(clone) : []), ...keep] }
  }
  // Paragraph marks: an inserted mark rejected, or a deleted one
  // accepted, joins the paragraph to the next.
  const paras = scope.name === 'w:p' ? [scope] : descendants(scope, 'w:p')
  for (const p of paras) {
    const mark = xpath(p, 'w:pPr', 'w:rPr')
    if (!mark) continue
    for (const kind of ['w:ins', 'w:del']) {
      const x = child(mark, kind)
      if (!x || !pick(x)) continue
      remove(mark, x)
      const join = (kind === 'w:ins') !== accept
      if (!join) continue
      const owner = parentOf(m.body, p)
      if (!owner) continue
      const next = owner.kids.slice(owner.kids.indexOf(p) + 1).find((k) => k.name === 'w:p')
      if (!next) { if (!descendants(p, 'w:t').length && !descendants(p, 'w:drawing').length) remove(owner, p); continue }
      const content = p.kids.filter((k) => k.name !== 'w:pPr')
      const npr = child(next, 'w:pPr')
      next.kids.splice(npr ? next.kids.indexOf(npr) + 1 : 0, 0, ...content)
      remove(owner, p)
    }
    if (!mark.kids.length) remove(child(p, 'w:pPr'), mark)
  }
}

// readRevisions lists who made each tracked change and when.
function readRevisions (m) {
  const out = {}
  for (const n of descendants(m.body, (x) => x.name === 'w:ins' || x.name === 'w:del')) {
    out[n.attrs['w:id']] = { kind: n.name === 'w:ins' ? 'ins' : 'del', author: n.attrs['w:author'] ?? '', date: n.attrs['w:date'] ?? '' }
  }
  return out
}

// ------------------------------------------------------------- notes

const NOTE = {
  fn: { rel: 'footnotes', file: 'footnotes', root: 'w:footnotes', item: 'w:footnote', ref: 'w:footnoteRef', ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml' },
  en: { rel: 'endnotes', file: 'endnotes', root: 'w:endnotes', item: 'w:endnote', ref: 'w:endnoteRef', ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml' }
}

function notesPart (m, kind, create) {
  const k = NOTE[kind]
  let part = m.pkg.relsOfType(m.main, k.rel)[0]?.part
  if (part && m.pkg.has(part)) return { part, root: m.pkg.xml(part).root }
  if (!create) return undefined
  part = `word/${k.file}.xml`
  const sep = (type, id, mark) => el(k.item, { 'w:type': type, 'w:id': id }, [el('w:p', {}, [el('w:pPr', {}, [el('w:spacing', { 'w:after': '0', 'w:line': '240', 'w:lineRule': 'auto' })]), el('w:r', {}, [el(mark)])])])
  m.pkg.setXml(part, doc(el(k.root, { 'xmlns:w': W }, [sep('separator', '-1', 'w:separator'), sep('continuationSeparator', '0', 'w:continuationSeparator')])))
  m.pkg.setOverride(part, k.ct)
  m.pkg.addRel(m.main, REL(k.rel), part)
  return { part, root: m.pkg.xml(part).root }
}

function noteParas (k, text) {
  return text.split('\n').map((line, i) => el('w:p', {}, [
    el('w:pPr', {}, [el('w:spacing', { 'w:after': '0' })]),
    i === 0 ? el('w:r', {}, [el('w:rPr', {}, [v('w:vertAlign', 'superscript')]), el(k.ref)]) : null,
    line || i === 0 ? el('w:r', {}, [el('w:rPr', {}, [v('w:sz', 18), v('w:szCs', 18)]), ...textNodes((i === 0 ? ' ' : '') + line)]) : null
  ]))
}

function addNote (m, kind, text) {
  const k = NOTE[kind]
  const np = notesPart(m, kind, true)
  const id = String(Math.max(0, ...children(np.root, k.item).map((x) => +x.attrs['w:id'] || 0)) + 1)
  np.root.kids.push(el(k.item, { 'w:id': id }, noteParas(k, text)))
  m.pkg.touch(np.part)
  return id
}

function editNote (m, kind, id, text, where) {
  const k = NOTE[kind]
  const np = notesPart(m, kind, false)
  const note = np && children(np.root, k.item).find((x) => x.attrs['w:id'] === id && !x.attrs['w:type'])
  if (!note) throw new Error(`${where}: 找不到${kind === 'fn' ? '脚注' : '尾注'} ${id}`)
  note.kids = noteParas(k, text)
  m.pkg.touch(np.part)
}

function readNotes (m) {
  const out = []
  for (const kind of ['fn', 'en']) {
    const np = notesPart(m, kind, false)
    if (!np) continue
    for (const n of children(np.root, NOTE[kind].item)) {
      if (n.attrs['w:type'] && n.attrs['w:type'] !== 'normal') continue
      const text = children(n, 'w:p').map((p) => descendants(p, (x) => x.name === 'w:t' || x.name === 'w:tab').map((x) => (x.name === 'w:tab' ? '\t' : textContent(x))).join('')).join('\n').replace(/^ /, '')
      out.push({ kind, id: n.attrs['w:id'], text, locked: descendants(n, (x) => x.name === 'w:drawing' || x.name === 'w:fldChar' || x.name === 'w:tbl').length > 0 })
    }
  }
  return out
}

// settleNotes gives notes made in this save their ids in the body and
// drops notes whose reference the person deleted.
function settleNotes (m, noted, fresh) {
  for (const kind of ['fn', 'en']) {
    const refName = kind === 'fn' ? 'w:footnoteReference' : 'w:endnoteReference'
    const refs = descendants(m.body, refName)
    for (const r of refs) { const f = fresh.get(r.attrs['w:id']); if (f && f.kind === kind) r.attrs['w:id'] = f.id }
    const np = notesPart(m, kind, false)
    if (!np) continue
    const now = new Set(descendants(m.body, refName).map((x) => x.attrs['w:id']))
    const items = children(np.root, NOTE[kind].item)
    const known = new Set(items.map((x) => x.attrs['w:id']))
    for (const r of descendants(m.body, refName)) if (!known.has(r.attrs['w:id'])) { const run = parentOf(m.body, r); const p = run && parentOf(m.body, run); if (p) remove(p, run) }
    let touched = false
    for (const x of items) {
      if (x.attrs['w:type']) continue
      const id = x.attrs['w:id']
      const isNew = [...fresh.values()].some((f) => f.kind === kind && f.id === id)
      if ((noted[kind].has(id) || isNew) && !now.has(id)) { remove(np.root, x); touched = true }
    }
    if (touched) m.pkg.touch(np.part)
  }
}

// ------------------------------------------------------------- picture layout

// wrapOf reads how text flows around a picture: inline, or floating
// square, tight, top-and-bottom, behind or in front, and its alignment.
function wrapOf (drawing) {
  const a = descendants(drawing, 'wp:anchor')[0]
  if (!a) return {}
  const kind = child(a, 'wp:wrapSquare') ? 'square' : child(a, 'wp:wrapTight') || child(a, 'wp:wrapThrough') ? 'tight' : child(a, 'wp:wrapTopAndBottom') ? 'topBottom' : a.attrs.behindDoc === '1' ? 'behind' : 'front'
  const al = xpath(a, 'wp:positionH', 'wp:align')
  return { wrap: kind, align: al ? textContent(al) : 'left' }
}

// shapeDrawing sets a picture's size and how text wraps around it.
function shapeDrawing (d, r) {
  const box = descendants(d, (x) => x.name === 'wp:inline' || x.name === 'wp:anchor')[0]
  if (!box) return
  if (r.w > 0 && r.h > 0) {
    const cx = String(Math.round(r.w * EMU_PER_PX)); const cy = String(Math.round(r.h * EMU_PER_PX))
    const ext = child(box, 'wp:extent')
    const cur = ext ? [Math.round(+ext.attrs.cx / EMU_PER_PX), Math.round(+ext.attrs.cy / EMU_PER_PX)] : [0, 0]
    if (Math.abs(cur[0] - r.w) > 1 || Math.abs(cur[1] - r.h) > 1) {
      if (ext) Object.assign(ext.attrs, { cx, cy })
      for (const e of descendants(box, (x) => x.name === 'a:ext' && parentOf(box, x)?.name === 'a:xfrm')) Object.assign(e.attrs, { cx, cy })
    }
  }
  if (r.wrap === undefined) return
  const want = r.wrap || 'inline'
  const have = box.name === 'wp:inline' ? 'inline' : wrapOf(d).wrap
  const align = ['left', 'right', 'center'].includes(r.align) ? r.align : 'left'
  if (want === have && (want === 'inline' || wrapOf(d).align === align)) return
  const keep = (name) => child(box, name)
  const parts = [keep('wp:extent'), keep('wp:effectExtent') ?? el('wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' }), keep('wp:docPr'), keep('wp:cNvGraphicFramePr'), keep('a:graphic')]
  let node
  if (want === 'inline') node = el('wp:inline', { distT: '0', distB: '0', distL: '0', distR: '0' }, parts.filter(Boolean))
  else {
    const wrapEl = want === 'square' ? el('wp:wrapSquare', { wrapText: 'bothSides' })
      : want === 'tight' ? el('wp:wrapTight', { wrapText: 'bothSides' }, [el('wp:wrapPolygon', { edited: '0' }, [el('wp:start', { x: '0', y: '0' }), el('wp:lineTo', { x: '0', y: '21600' }), el('wp:lineTo', { x: '21600', y: '21600' }), el('wp:lineTo', { x: '21600', y: '0' }), el('wp:lineTo', { x: '0', y: '0' })])])
        : want === 'topBottom' ? el('wp:wrapTopAndBottom')
          : el('wp:wrapNone')
    const [extent, effect, docPr, frame, graphic] = parts
    node = el('wp:anchor', { distT: '0', distB: '0', distL: '114300', distR: '114300', simplePos: '0', relativeHeight: '251659264', behindDoc: want === 'behind' ? '1' : '0', locked: '0', layoutInCell: '1', allowOverlap: '1' }, [
      el('wp:simplePos', { x: '0', y: '0' }),
      el('wp:positionH', { relativeFrom: 'column' }, [el('wp:align', {}, [text(align)])]),
      el('wp:positionV', { relativeFrom: 'paragraph' }, [el('wp:posOffset', {}, [text('0')])]),
      extent, effect, wrapEl, docPr, frame, graphic
    ].filter(Boolean))
  }
  const parent = parentOf(d, box) ?? d
  parent.kids.splice(parent.kids.indexOf(box), 1, node)
}
