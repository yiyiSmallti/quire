// The three files the person edits in the web run, written with quire's
// own writers: node fixture.mjs LIB WS
import fs from 'node:fs'
import path from 'node:path'
const [lib, ws] = process.argv.slice(2)
const { createDocx } = await import(path.join(lib, 'docx.js'))
const { createXlsx } = await import(path.join(lib, 'xlsx.js'))
const { createPptx } = await import(path.join(lib, 'pptx.js'))
const io = { loadImage: async (p) => fs.readFileSync(path.resolve(ws, p)) }
const md = '# 季度报告\n\n本季度收入 **120 万元**。\n\n| 地区 | 收入 |\n| --- | ---: |\n| 华东 | 120 |\n| 华北 | 98 |\n\n- 华东区贡献最大\n- 成本下降 5%\n\n![趋势图](chart.png)\n\n## 结论\n\n继续保持增长。\n'
fs.writeFileSync(path.join(ws, '季度报告.docx'), (await createDocx(md, io)).buffer)
const xl = createXlsx([{ name: '明细', header: true, rows: [['项目', '单价', '数量', '小计'], ['服务器', 8000, 5, '=B2*C2'], ['带宽', 1200, 12, '=B3*C3'], ['合计', null, null, '=SUM(D2:D3)']] }])
fs.writeFileSync(path.join(ws, '预算.xlsx'), xl.buffer)
const pp = await createPptx([{ layout: 'title', title: '季度经营回顾', subtitle: '战略部', notes: '开场白' }, { title: '核心要点', bullets: ['收入增长 13%', '成本下降 5%'] }], {}, io)
fs.writeFileSync(path.join(ws, '经营回顾.pptx'), pp.buffer)
