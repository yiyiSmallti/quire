// A deck PowerPoint-style objects quire did not write itself: a SmartArt
// diagram (data model plus the drawing PowerPoint saves for it), a table and
// a group of text boxes. node smartart.mjs LIB FILE
import fs from 'node:fs'
import path from 'node:path'
const [lib, file] = process.argv.slice(2)
const { createPptx, editPptxUi } = await import(path.join(lib, 'pptx.js'))
const { Package, REL } = await import(path.join(lib, 'pkg.js'))
const { parse, child } = await import(path.join(lib, 'xml.js'))

let buf = (await createPptx([
  { title: '项目流程', layout: 'title-only' },
  { title: '人员分工', table: { rows: [['阶段', '负责人', '周期'], ['调研', '王芳', '2 周'], ['开发', '李雷', '6 周']] } },
  { title: '备注', layout: 'title-only' }
], {})).buffer
buf = (await editPptxUi(buf, [
  { op: 'add_textbox', slide: 3, x: 900000, y: 2000000, cx: 3000000, cy: 600000 },
  { op: 'shape_text', slide: 3, shape: 2, paras: [{ runs: [{ text: '组合里的第一段' }] }] },
  { op: 'add_textbox', slide: 3, x: 900000, y: 2800000, cx: 3000000, cy: 600000 },
  { op: 'shape_text', slide: 3, shape: 3, paras: [{ runs: [{ text: '组合里的第二段' }] }] },
  { op: 'group', slide: 3, items: [{ shape: 2 }, { shape: 3 }] }
])).buffer

const pkg = Package.fromBuffer(buf)
const slide = 'ppt/slides/slide1.xml'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
const DRT = (t) => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagram' + t
// Three steps, the second with a sub-point: modelIds 1..4 are the nodes,
// 11..14 the presentation points that show them.
const steps = [['1', '需求调研', null], ['2', '方案设计', null], ['3', '开发上线', null], ['4', '灰度发布', '3']]
const tx = (s) => `<a:p><a:r><a:rPr lang="zh-CN"/><a:t>${s}</a:t></a:r></a:p>`
const pts = steps.map(([id, s]) => `<dgm:pt modelId="${id}"><dgm:prSet phldrT="[文本]"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>${tx(s)}</dgm:t></dgm:pt>`).join('') +
  steps.map(([id]) => `<dgm:pt modelId="1${id}" type="pres"><dgm:prSet presAssocID="${id}" presName="node" presStyleLbl="node1"/><dgm:spPr/></dgm:pt>`).join('')
const cxns = steps.map(([id, , par], i) => `<dgm:cxn modelId="c${id}" srcId="${par ?? '0'}" destId="${id}" srcOrd="${i}" destOrd="0"/>`).join('')
const data = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></dgm:t></dgm:pt>${pts}</dgm:ptLst><dgm:cxnLst>${cxns}</dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${DSP}" relId="rIdDraw" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>`
const box = (i) => ({ x: 200000 + i * 2700000, y: 300000, cx: 2400000, cy: 1200000 })
const drawSps = steps.slice(0, 3).map(([id, s], i) => {
  const b = box(i)
  const sub = id === '3' ? `<a:p><a:pPr lvl="1"/><a:r><a:rPr lang="zh-CN" sz="1400"/><a:t>灰度发布</a:t></a:r></a:p>` : ''
  return `<dsp:sp modelId="1${id}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="${b.x}" y="${b.y}"/><a:ext cx="${b.cx}" cy="${b.cy}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr><dsp:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="2000"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>${s}</a:t></a:r></a:p>${sub}</dsp:txBody><dsp:txXfrm><a:off x="${b.x}" y="${b.y}"/><a:ext cx="${b.cx}" cy="${b.cy}"/></dsp:txXfrm></dsp:sp>`
}).join('')
const drawing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dgm="${DGM}" xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${drawSps}</dsp:spTree></dsp:drawing>`
const empty = (root, ns) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:${root} xmlns:dgm="${DGM}" xmlns:a="${A}"${ns ?? ''}/>`

const parts = {
  'ppt/diagrams/data1.xml': [data, 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml', DRT('Data')],
  'ppt/diagrams/layout1.xml': [empty('layoutDef', ' uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process1"'), 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml', DRT('Layout')],
  'ppt/diagrams/quickStyle1.xml': [empty('styleDef'), 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml', DRT('QuickStyle')],
  'ppt/diagrams/colors1.xml': [empty('colorsDef'), 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml', DRT('Colors')],
  'ppt/diagrams/drawing1.xml': [drawing, 'application/vnd.ms-office.drawingml.diagramDrawing+xml', 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing']
}
const ids = {}
for (const [name, [xml, ct, type]] of Object.entries(parts)) {
  pkg.setXml(name, parse(xml))
  pkg.setOverride(name, ct)
  ids[name] = pkg.addRel(slide, type, name)
}
// The data model names the drawing by a relationship on the slide.
const dataXml = data.replace('rIdDraw', ids['ppt/diagrams/drawing1.xml'])
pkg.setXml('ppt/diagrams/data1.xml', parse(dataXml))

const frame = parse(`<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${A}" xmlns:r="${REL('').replace(/\/$/, '')}"><p:nvGraphicFramePr><p:cNvPr id="90" name="流程图"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="838200" y="2000000"/><a:ext cx="8400000" cy="1800000"/></p:xfrm><a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="${ids['ppt/diagrams/data1.xml']}" r:lo="${ids['ppt/diagrams/layout1.xml']}" r:qs="${ids['ppt/diagrams/quickStyle1.xml']}" r:cs="${ids['ppt/diagrams/colors1.xml']}"/></a:graphicData></a:graphic></p:graphicFrame>`).root
delete frame.attrs['xmlns:p']; delete frame.attrs['xmlns:a']; delete frame.attrs['xmlns:r']
const sld = pkg.xml(slide)
sld.root.attrs['xmlns:r'] ??= 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
child(child(sld.root, 'p:cSld'), 'p:spTree').kids.push(frame)
pkg.touch(slide)
fs.writeFileSync(file, pkg.toBuffer())
