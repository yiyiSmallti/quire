// Real dsh with the plugin package installed into an sdk-template profile:
// a fake model creates, reads and edits a Word document, an Excel workbook
// and a PowerPoint deck, edits a document another program wrote, and
// exports a preview. Readers that do not go through quire check the files:
// textutil, Quick Look, unzip, and python-docx / openpyxl when installed.
import fs from 'node:fs'
import path from 'node:path'
import { chromePath } from './chrome.mjs'
import {
  Cleanup, Dsh, E2E, REPORT, ROOT, check, chart, clip, findDsh, goList, keep, packPlugin,
  quireTools, run, shortTemp, startModel, tail, unzip, which
} from './harness.mjs'

const reportMD = `# 季度报告（草稿）

本季度收入 **120 万元**，比上季度增长 *12%*。

## 分区数据

| 地区 | 收入 | 增长 |
| --- | ---: | ---: |
| 华东 | 120 | 12% |
| 华北 | 98 | 8% |

- 华东区贡献最大
  - 新客户 30 家
- 成本下降 5%

![趋势图](chart.png)

> 数据截至 2026-09-30。
`

// legacyHTML becomes a .docx written by another program (macOS textutil).
const legacyHTML = `<html><head><meta charset="utf-8"></head><body>
<h1>旧文档</h1><p>第一段，由别的程序写的。</p><p>第二段需要改。</p>
<ul><li>条目一</li><li>条目二</li></ul>
<table border="1"><tr><td>甲</td><td>乙</td></tr><tr><td>丙</td><td>丁</td></tr></table>
</body></html>`

// pyCheck reads the files with python-docx and openpyxl.
const pyCheck = `
import sys, docx, openpyxl
d = docx.Document(sys.argv[1])
print("docx-title=" + d.paragraphs[0].text)
print("docx-cell=" + d.tables[0].cell(2, 1).text)
wb = openpyxl.load_workbook(sys.argv[2], data_only=True)
print("sheets=" + ",".join(wb.sheetnames))
print("total=%d" % wb["汇总"]["B1"].value)
wf = openpyxl.load_workbook(sys.argv[2])
print("formula=" + str(wf["汇总"]["B1"].value))
`

const call = (name, input) => ({ name, input })
const fields = (s) => s.trim().split(/\s+/).filter(Boolean).join(' ')
const shown = (r) => `{Text:${r.text} IsError:${r.isError}}`

function zipHas (data, name) {
  try { return unzip(data).some((e) => e.name === name) } catch { return false }
}

// zipDiff lists the entries of b whose content differs from a, and the
// entries of a that b lacks.
function zipDiff (a, b) {
  const read = (data) => new Map(unzip(data).map((e) => [e.name, e.data()]))
  const ma = read(a)
  const mb = read(b)
  const missing = [...ma.keys()].filter((n) => !mb.has(n))
  const changed = [...mb.keys()].filter((n) => !ma.has(n) || !ma.get(n).equals(mb.get(n)))
  return { changed, missing }
}

export async function vanilla (rep) {
  const s = rep.section('原版 dsh')
  const cleanup = new Cleanup()
  try {
    await suite(s, cleanup)
  } finally {
    await cleanup.run()
  }
}

async function suite (s, cleanup) {
  const dshBin = findDsh()
  if (!dshBin) {
    s.step('找到 dsh', new Error('PATH 上没有 dsh，也没设 QUIRE_E2E_DSH'))
    return
  }
  const work = shortTemp(cleanup)
  const d = new Dsh(dshBin, work)
  let ver = ''
  let err = null
  try {
    const tgz = packPlugin(work)
    const got = await d.install('quire-e2e', tgz)
    ver = got.ver
    err = check(got.dump.includes('id: quire') && got.dump.includes('@precinct/quire'), 'quire entry missing from dump-config')
  } catch (e) { err = e }
  if (!s.step('装进真实 dsh：sdk 模板 profile + quire 包，--dump-config 通过', err, 'dsh ' + ver)) return

  const ws = path.join(work, 'ws')
  fs.mkdirSync(ws, { recursive: true })
  const legacy = path.join(ws, 'legacy.docx')
  let haveTextutil = false
  err = null
  try {
    chart(path.join(ws, 'chart.png'))
    fs.writeFileSync(path.join(ws, 'notes.txt'), '不是 Office 文件\n')
    const lib = path.join(ROOT, 'plugin', 'lib')
    await run({}, work, process.execPath, [path.join(E2E, 'gen', 'review.mjs'), lib, path.join(ws, 'review.docx')])
    await run({}, work, process.execPath, [path.join(E2E, 'gen', 'smartart.mjs'), lib, path.join(ws, 'flow.pptx')])
    if (which('textutil')) {
      const htmlPath = path.join(work, 'legacy.html')
      fs.writeFileSync(htmlPath, legacyHTML)
      await run({}, work, 'textutil', ['-convert', 'docx', '-output', legacy, htmlPath])
      haveTextutil = true
    }
  } catch (e) { err = e }
  const legacyBefore = fs.existsSync(legacy) ? fs.readFileSync(legacy) : Buffer.alloc(0)
  if (!s.step('准备工作区：图表图片、一个文本文件、一份别的程序（textutil）写的 .docx、一份带别人修订的 .docx、一份带 SmartArt 和组合的 .pptx', err, `legacy.docx ${legacyBefore.length} 字节`)) return

  const script = [
    /* 0 */ call('docx_create', { file_path: 'report.docx', markdown: reportMD }),
    /* 1 */ call('office_read', { file_path: 'report.docx' }),
    /* 2 */ call('docx_edit', {
      file_path: 'report.docx',
      edits: [
        { op: 'replace', find: '草稿', replace: '终稿' },
        { op: 'cell', table: 't1', row: 3, col: 2, text: '101' },
        { op: 'insert', after: 'p3', markdown: '本节数据来自**财务系统**。' },
        { op: 'delete', blocks: ['p8'] },
        { op: 'append', markdown: '## 结论\n\n继续保持增长。' }
      ]
    }),
    /* 3 */ call('office_read', { file_path: 'report.docx' }),
    /* 4 */ call('docx_create', { file_path: 'report.docx', markdown: '# 覆盖' }),
    /* 5 */ call('xlsx_create', {
      file_path: 'budget.xlsx',
      sheets: [
        {
          name: '预算',
          header: true,
          rows: [
            ['项目', '单价', '数量', '小计'],
            ['服务器', 8000, 3, '=B2*C2'],
            ['带宽', 1200, 12, '=B3*C3'],
            ['合计', null, null, '=SUM(D2:D3)']
          ]
        },
        { name: '汇总', rows: [['总预算', '=预算!D4'], ['平均单价', '=AVERAGE(预算!B2:B3)']] }
      ]
    }),
    /* 6 */ call('xlsx_edit', {
      file_path: 'budget.xlsx',
      edits: [
        { op: 'set', sheet: '预算', cell: 'C2', value: 5 },
        { op: 'style', sheet: '预算', range: 'B2:D4', format: '#,##0' },
        { op: 'add_sheet', name: '备注', rows: [['说明'], ['由 Agent 生成']] },
        { op: 'rename_sheet', sheet: '预算', name: '明细' }
      ]
    }),
    /* 7 */ call('office_read', { file_path: 'budget.xlsx', sheet: '汇总' }),
    /* 8 */ call('pptx_create', {
      file_path: 'deck.pptx',
      slides: [
        { title: '季度经营回顾', subtitle: '战略部' },
        { title: '核心要点', bullets: ['收入增长 12%', { text: '华东区贡献最大', level: 1 }], notes: '强调华东' },
        { title: '分区数据', table: { rows: [['地区', '收入'], ['华东', '120']] } },
        { title: '趋势', image: { path: 'chart.png' } }
      ],
      theme: { accent: 'C00000' }
    }),
    /* 9 */ call('pptx_edit', {
      file_path: 'deck.pptx',
      edits: [
        { op: 'replace', find: '12%', replace: '13%', all: true },
        { op: 'set_text', shape: 's2.2', text: '收入增长 13%\n  华东区贡献最大\n成本下降 5%' },
        { op: 'notes', slide: 1, text: '开场白' },
        { op: 'move_slide', slide: 4, to: 2 },
        { op: 'delete_slide', slide: 3 },
        { op: 'add_slide', slide: { title: '谢谢', subtitle: 'Q&A' } }
      ]
    }),
    /* 10 */ call('office_read', { file_path: 'deck.pptx' }),
    /* 11 */ call('docx_edit', { file_path: 'legacy.docx', edits: [{ op: 'replace', find: '需要改', replace: '已改好' }] }),
    /* 12 */ call('office_read', { file_path: 'legacy.docx' }),
    /* 13 */ call('office_preview', { file_path: 'report.docx' }),
    /* 14 */ call('xlsx_create', { file_path: '../outside.xlsx', sheets: [{ name: 'S', rows: [[1]] }] }),
    /* 15 */ call('office_read', { file_path: 'notes.txt' }),
    /* 16 */ call('docx_edit', { file_path: 'report.docx', edits: [{ op: 'delete', blocks: ['p99'] }] }),
    /* 17 */ call('xlsx_edit', {
      file_path: 'budget.xlsx',
      edits: [
        { op: 'insert_rows', sheet: '明细', at: 3 },
        { op: 'write', sheet: '明细', at: 'A3', rows: [['存储', 500, 4, '=B3*C3']] },
        { op: 'add_chart', sheet: '明细', range: 'A1:B4', type: 'col', title: '单价对比', at: 'F2', labels: 'value', legend: 'none', y_title: '元' },
        { op: 'edit_chart', sheet: '明细', chart: 1, type: 'bar', gridlines: false, y_min: 0 }
      ]
    }),
    /* 18 */ call('office_read', { file_path: 'budget.xlsx', sheet: '明细' }),
    /* 19 */ call('office_read', { file_path: 'budget.xlsx', sheet: '汇总' }),
    /* 20 */ call('docx_edit', { file_path: 'report.docx', edits: [{ op: 'comment', block: 'p2', find: '120 万元', text: '数字请财务复核', author: '审阅人' }] }),
    /* 21 */ call('docx_edit', { file_path: 'report.docx', edits: [{ op: 'reply', comment: 0, text: '已复核，无误' }, { op: 'resolve', comment: 0 }] }),
    /* 22 */ call('office_read', { file_path: 'report.docx' }),
    /* 23 */ call('office_read', { file_path: 'review.docx' }),
    /* 24 */ call('docx_edit', { file_path: 'review.docx', edits: [{ op: 'accept', block: 'p2' }, { op: 'reject' }, { op: 'track', on: false }] }),
    /* 25 */ call('office_read', { file_path: 'review.docx' }),
    /* 26 */ call('office_read', { file_path: 'flow.pptx' }),
    /* 27 */ call('pptx_edit', {
      file_path: 'flow.pptx',
      edits: [
        { op: 'set_text', shape: 's1.2.4', text: '全量发布' },
        { op: 'set_text', shape: 's1.2.1', text: '需求评审' },
        { op: 'set_text', shape: 's3.3', text: '组合里改过的第二段' }
      ]
    }),
    /* 28 */ call('office_read', { file_path: 'flow.pptx' })
  ]
  if (!haveTextutil) {
    script[11] = call('office_read', { file_path: 'report.docx' })
    script[12] = call('office_read', { file_path: 'report.docx' })
  }
  const model = await startModel(cleanup, script)
  const turn = await d.turn('quire-e2e', ws, model, '写一份季度报告、一张预算表和一份汇报 PPT，然后按要求修改')
  const res = model.results
  s.step('一个回合：假模型依次调用 quire 工具', turn.err, `模型收到 ${res.length} 个工具结果`)
  if (turn.err || res.length < script.length) {
    console.error('dsh stderr: ' + tail(turn.stderr, 40))
    if (!turn.err) s.step('工具结果齐全', new Error(`只有 ${res.length} 个；dsh stderr：${tail(turn.stderr, 8)}`))
    return
  }
  const have = quireTools(model.tools)
  s.step('模型看到八个 quire 工具，系统提示里有 Office 文件的用法',
    check(have.length === 8 && model.system.includes('zip archives: read them with office_read'), `tools = ${goList(have)}`), have.join('、'))

  const ok = (i, cond, name) => s.step(name, check(cond && !res[i].isError, `result = ${shown(res[i])}`), clip(res[i].text, 300))
  const fails = (i, cond, name) => s.step(name, check(res[i].isError && cond, `result = ${shown(res[i])}`), clip(res[i].text, 300))
  const t = res.map((r) => r.text)
  ok(0, t[0].includes('Created ') && t[0].includes('report.docx'), 'docx_create：Markdown 新建 Word 文档（标题、粗斜体、表格、嵌套列表、图片、引用）')
  ok(1, t[1].includes('[p1] # 季度报告（草稿）') && t[1].includes('**120 万元**') &&
    t[1].includes('[t1] table 3 rows × 3 columns') && t[1].includes('| 华北 | 98 | 8% |') &&
    t[1].includes('[p5]   - 新客户 30 家') && t[1].includes('[image: 趋势图'),
  'office_read（Word）：块编号、标题级别、粗体、表格、列表缩进、图片')
  ok(2, t[2].includes('Edited '), 'docx_edit：替换、改单元格、在块后插入、删除块、追加，一次调用')
  const r3 = t[3]
  ok(3, r3.includes('# 季度报告（终稿）') && r3.includes('| 华北 | 101 | 8% |') &&
    r3.includes('本节数据来自**财务系统**。') && !r3.includes('数据截至') && r3.includes('## 结论') &&
    r3.indexOf('## 分区数据') < r3.indexOf('本节数据来自') && r3.indexOf('本节数据来自') < r3.indexOf('[t1]'),
  'office_read：编辑都落在对的位置（插入在 p3 之后、表格之前，引用已删）')
  fails(4, t[4].includes('already exists') && t[4].includes('overwrite: true'), 'docx_create 不覆盖已有文件：报错并提示用 docx_edit 或 overwrite')
  ok(5, t[5].includes('Created ') && t[5].includes('2 sheet(s)'), 'xlsx_create：两张工作表、表头、公式、跨表引用')
  ok(6, t[6].includes('renamed 预算 to 明细') && t[6].includes('added sheet 备注'), 'xlsx_edit：改单元格、设数字格式、加工作表、改名')
  ok(7, t[7].includes('- 明细:') && t[7].includes('- 备注:') && t[7].includes('=明细!D4 → 54400') &&
    t[7].includes('=AVERAGE(明细!B2:B3) → 4600'),
  'office_read（Excel）：重算后汇总 = 8000×5 + 1200×12 = 54400，改名后公式跟着改')
  ok(8, t[8].includes('4 slide(s)'), 'pptx_create：标题页、要点页（两级）、表格页、图片页、备注、主题色')
  ok(9, t[9].includes('moved slide 4 to 2') && t[9].includes('deleted slide 3'), 'pptx_edit：替换、改形状文字、备注、移页、删页、加页')
  const r10 = t[10]
  ok(10, r10.includes('4 slides') && r10.includes('## Slide 2 · Title and Content\n[s2.1] title: 趋势') &&
    r10.includes('收入增长 13%') && r10.includes('成本下降 5%') && r10.includes('notes: 开场白') &&
    !r10.includes('分区数据') && r10.includes('[s4.1] ctrTitle: 谢谢'),
  'office_read（PowerPoint）：页序、改过的文字、备注都对，删掉的页不在')
  if (haveTextutil) {
    ok(11, t[11].includes('replaced 1'), 'docx_edit 别的程序写的文档：替换一处文字')
    ok(12, t[12].includes('第二段已改好') && t[12].includes('第一段，由别的程序写的。') && t[12].includes('丁'),
      'office_read 读回：改了的地方变了，其余内容都在')
  }
  ok(13, t[13].includes('Wrote the preview'), 'office_preview（原版 dsh）：预览写成 HTML 文件')
  fails(14, t[14].includes('sandbox') && t[14].includes('outside the workspace'), '沙箱 workspace-write：工作区外的路径拒绝写入')
  const outside = path.join(work, 'outside.xlsx')
  s.step('工作区外没有留下文件', check(!fs.existsSync(outside), `stat = ${outside} exists`))
  fails(15, t[15].includes('not a Word, Excel or PowerPoint file'), 'office_read 非 Office 文件：明确报错')
  fails(16, t[16].includes('no block p99') && t[16].includes('office_read'), 'docx_edit 编号不存在：报错并提示重新 office_read')
  const report = fs.readFileSync(path.join(ws, 'report.docx'))
  s.step('失败的编辑不动文件', check(report.includes('PK') && zipHas(report, 'word/document.xml'), 'report.docx broken'))
  ok(17, t[17].includes('inserted 明细 row 3') && t[17].includes('added chart 1 on 明细 from A1:B4') && t[17].includes('edited chart 1'),
    'xlsx_edit：插入一行、写入、按区域加图表（数据标签、图例、轴标题），再改成条形图')
  ok(18, t[18].includes('chart 1: bar "单价对比"') && t[18].includes('labels value, legend none, y_title "元", y_min 0, no gridlines') &&
    t[18].includes('| 3 | 存储 | 500 | 4 | =B3*C3 → 2000 |') && t[18].includes('| 4 | 带宽 | 1,200 | 12 | =B4*C4 → 14,400 |') &&
    t[18].includes('=SUM(D2:D4) → 56,400'),
  'office_read（Excel）：插入行后下面的公式跟着下移、合计范围扩大；图表和外观设置都列出来')
  ok(19, t[19].includes('=明细!D5 → 56400') && t[19].includes('=AVERAGE(明细!B2:B4)'), 'office_read：别的表引用插入行下方的单元格，也跟着改')
  ok(20, t[20].includes('Edited '), 'docx_edit：在一段里的「120 万元」上加批注')
  ok(21, t[21].includes('Edited '), 'docx_edit：回复批注并标为已解决')
  ok(22, t[22].includes('comment 0 by 审阅人 (resolved) on p2 "120 万元": 数字请财务复核') && t[22].includes('  - comment 1 by Agent: 已复核，无误') &&
    t[22].includes('# 季度报告（终稿）'),
  'office_read（Word）：批注、作者、锚定文字、回复和解决状态都列出来，正文不变')
  ok(23, t[23].includes('4 tracked changes') && t[23].includes('change 1: deleted by 王芳 in p2 "三"') && t[23].includes('change 4: inserted by 李雷 in p3 "三"'),
    'office_read（Word）：别人的修订按作者、段落、增删列出')
  ok(24, t[24].includes('Edited '), 'docx_edit：接受 p2 的修订，拒绝其余的，关掉修订')
  ok(25, t[25].includes('[p2] 付款期限为六十天。') && t[25].includes('[p3] 违约金按日万分之五计算。') &&
    !t[25].includes('tracked changes') && !t[25].includes('track changes is on'),
  'office_read 读回：接受的留下、拒绝的复原，文档里不再有修订')
  ok(26, t[26].includes('[s1.2] SmartArt "流程图"') && t[26].includes('    - [s1.2.4] 灰度发布') && t[26].includes('组合里的第二段'),
    'office_read（PowerPoint）：SmartArt 按节点层级列出，组合里的文本框也列出来')
  ok(27, !t[27].includes('Error'), 'pptx_edit：改 SmartArt 的子节点、顶层节点和组合里的文本框')
  ok(28, t[28].includes('[s1.2.1] 需求评审') && t[28].includes('    - [s1.2.4] 全量发布') && t[28].includes('组合里改过的第二段') &&
    !t[28].includes('灰度发布'),
  'office_read 读回：SmartArt 节点和组合文字都改了，层级不变')
  let saErr = null
  let saNote = ''
  try {
    const parts = new Map(unzip(fs.readFileSync(path.join(ws, 'flow.pptx'))).map((e) => [e.name, e.data().toString('utf8')]))
    const data = parts.get('ppt/diagrams/data1.xml') ?? ''
    const drawing = parts.get('ppt/diagrams/drawing1.xml') ?? ''
    // PowerPoint shows the saved drawing until it re-lays the diagram out, so
    // the text has to change there too, the child still an indented line.
    saErr = check(data.includes('全量发布') && data.includes('需求评审') && drawing.includes('需求评审') &&
      /<a:pPr[^>]*lvl="1"[^>]*\/?>(?:(?!<\/a:p>).)*全量发布/s.test(drawing) && !drawing.includes('灰度发布'), `drawing = ${clip(drawing, 400)}`)
    saNote = 'data1.xml 与 drawing1.xml 都已更新'
  } catch (e) { saErr = e }
  s.step('SmartArt 的数据模型和 PowerPoint 存下的绘图都改了，子节点仍是缩进一级的段落', saErr, saNote)

  // ------------------------------------------- readers other than quire
  const files = ['report.docx', 'budget.xlsx', 'deck.pptx', 'review.docx', 'flow.pptx']
  if (haveTextutil) files.push('legacy.docx')
  const bad = []
  for (const f of files) {
    try { await run({}, ws, 'unzip', ['-tq', f]) } catch (e) { bad.push(f + ': ' + e.message) }
  }
  s.step('unzip -t：所有文件都是完整的 zip', check(!bad.length, goList(bad)), files.join('、'))

  if (haveTextutil) {
    let txt = ''
    err = null
    try {
      txt = await run({}, ws, 'textutil', ['-convert', 'txt', '-stdout', 'report.docx'])
      err = check(txt.includes('季度报告（终稿）') && txt.includes('101') && txt.includes('本节数据来自财务系统') &&
        txt.includes('继续保持增长') && !txt.includes('数据截至'), `text = ${clip(txt, 400)}`)
    } catch (e) { err = e }
    s.step('textutil（Apple 的 Word 读取器）读 report.docx：内容与编辑一致', err, clip(fields(txt), 200))

    let changed = []
    err = null
    try {
      const diff = zipDiff(legacyBefore, fs.readFileSync(legacy))
      changed = diff.changed
      err = check(!diff.missing.length && changed.length === 1 && changed[0] === 'word/document.xml', `changed = ${goList(changed)}, missing = ${goList(diff.missing)}`)
    } catch (e) { err = e }
    s.step('保真：别的程序写的 .docx 改一处文字后，只有 word/document.xml 变了，其余部件逐字节相同', err, `变化的部件：${goList(changed)}`)
    txt = ''
    err = null
    try {
      txt = await run({}, ws, 'textutil', ['-convert', 'txt', '-stdout', 'legacy.docx'])
      err = check(txt.includes('第二段已改好') && txt.includes('条目二') && txt.includes('丁'), `text = ${txt}`)
    } catch (e) { err = e }
    s.step('textutil 读 legacy.docx：改动生效，列表与表格都在', err, clip(fields(txt), 120))
  }

  const py = which('python3')
  if (py) {
    let out = ''
    let perr = null
    try { out = await run({}, ws, py, ['-I', '-c', pyCheck, path.join(ws, 'report.docx'), path.join(ws, 'budget.xlsx')]) } catch (e) { perr = e; out = e.out ?? '' }
    if (!perr) {
      s.step('python-docx / openpyxl 读：段落与表格、工作表名、公式缓存值（data_only）', check(
        out.includes('docx-title=季度报告（终稿）') && out.includes('docx-cell=101') &&
        out.includes('sheets=明细,汇总,备注') && out.includes('total=56400') && out.includes('formula==明细!D5'), out), out.trim())
    } else if (out.includes('ModuleNotFoundError')) {
      s.appendix += '\n本机 python3 没有 python-docx/openpyxl，跳过这一项核对。\n'
    } else {
      s.step('python-docx / openpyxl 读', perr)
    }
  }

  if (process.platform === 'darwin') {
    const ql = path.join(work, 'ql')
    fs.mkdirSync(ql, { recursive: true })
    const made = []
    let qerr = null
    for (const f of ['report.docx', 'budget.xlsx', 'deck.pptx', 'flow.pptx']) {
      if (!fs.existsSync(path.join(ws, f))) {
        qerr = new Error(`stat ${f}: no such file or directory`)
        continue
      }
      try {
        await run({}, ws, 'qlmanage', ['-t', '-s', '900', '-o', ql, f])
      } catch (e) {
        qerr = e
        continue
      }
      const png = path.join(ql, f + '.png')
      if (!fs.existsSync(png) || fs.statSync(png).size < 2000) {
        qerr = new Error(`${f}: no thumbnail`)
        continue
      }
      keep(png, 'vanilla-ql-' + f + '.png')
      made.push('vanilla-ql-' + f + '.png')
    }
    s.step('Quick Look（macOS 的 Office 渲染器）给四个文件各出一张缩略图', qerr, made.join('、'))
    for (const m of made) s.appendix += `\n![${m}](${m})\n`
  }

  let preview = ''
  const i = t[13].indexOf(' to ')
  if (i >= 0) preview = t[13].slice(i + 4).split('; ')[0].replace(/;$/, '')
  let perr = null
  try {
    const page = fs.readFileSync(preview, 'utf8')
    perr = check(page.includes('季度报告（终稿）') && page.includes('data:image/png;base64,') &&
      !page.includes('<script') && page.includes('Content-Security-Policy'), 'preview content')
  } catch (e) { perr = e }
  s.step('预览 HTML：内容是改后的文档，图片内联，没有脚本，带 CSP', perr, path.basename(preview))
  if (!perr) {
    keep(preview, 'vanilla-preview-report.html')
    fs.rmSync(preview, { force: true })
    const cp = chromePath()
    if (cp) {
      const shot = path.join(work, 'preview.png')
      let err = null
      try {
        await run({}, work, cp, ['--headless=new', '--disable-gpu', '--hide-scrollbars', '--window-size=1100,1500',
          '--user-data-dir=' + path.join(work, 'chrome-shot'), '--screenshot=' + shot, 'file://' + path.join(REPORT, 'vanilla-preview-report.html')])
        err = check(fs.existsSync(shot) && fs.statSync(shot).size > 10000, 'no screenshot')
      } catch (e) { err = e }
      if (s.step('Chrome 打开预览页截图', err, 'vanilla-preview-report.png')) {
        keep(shot, 'vanilla-preview-report.png')
        s.appendix += '\n预览页截图：\n\n![vanilla-preview-report.png](vanilla-preview-report.png)\n'
      }
    }
  }
  for (const f of files) keep(path.join(ws, f), 'vanilla-' + f)

  s.appendix += '\n### 模型收到的工具结果\n\n'
  res.forEach((x, i) => {
    s.appendix += `${i + 1}. \`${script[i].name}\`（is_error=${x.isError}）\n\n\`\`\`\n${clip(x.text, 1400)}\n\`\`\`\n\n`
  })
}
