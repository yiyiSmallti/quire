// The sample workspace the README screenshots show, written with quire's
// own writers: node demo.mjs LIB WS (WS holds chart.png).
import fs from 'node:fs'
import path from 'node:path'
const [lib, ws] = process.argv.slice(2)
const { createDocx } = await import(path.join(lib, 'docx.js'))
const { createXlsx, editXlsxUi } = await import(path.join(lib, 'xlsx.js'))
const { createPptx } = await import(path.join(lib, 'pptx.js'))
const io = { loadImage: async (p) => fs.readFileSync(path.resolve(ws, p)) }

const md = `# 2026 年第三季度经营报告

本季度实现营业收入 **1,286 万元**，同比增长 *13.4%*；毛利率 41.5%，比上季度提高 1.3 个百分点。华东、华南两个大区贡献了六成以上的增量，新产品线首次实现盈亏平衡。

## 一、分区经营情况

| 大区 | 收入（万元） | 同比 | 毛利率 |
| --- | ---: | ---: | ---: |
| 华东 | 486 | 18.2% | 43.5% |
| 华南 | 352 | 15.7% | 42.1% |
| 华北 | 268 | 7.9% | 38.6% |
| 西部 | 180 | 6.3% | 37.2% |

华东区受益于渠道下沉，新增签约门店 46 家；华北区增速放缓，主要是大客户订单推迟到第四季度确认。

![各季度收入（万元）](chart.png)

## 二、成本与费用

- 原材料采购价格同比下降 3.1%，带动毛利率回升
- 销售费用率 14.8%，与上季度持平
  - 线上投放占比提高到 42%
  - 线下展会减少两场
- 研发投入 156 万元，占收入 12.1%

## 三、下季度重点

1. 推进华北大客户订单交付，确保年内确认收入
2. 新产品线扩大到 3 个大区试点
3. 控制销售费用率在 15% 以内

> 本报告数据未经审计，最终以年度审计报告为准。
`
fs.writeFileSync(path.join(ws, '季度报告.docx'), (await createDocx(md, io)).buffer)

const months = ['1 月', '2 月', '3 月', '4 月', '5 月', '6 月', '7 月', '8 月', '9 月']
const income = [312, 298, 345, 368, 382, 401, 418, 432, 436]
const cost = [196, 189, 214, 221, 229, 238, 246, 252, 254]
const rows = [['月份', '收入', '成本', '毛利', '毛利率']]
months.forEach((m, i) => {
  const r = i + 2
  rows.push([m, income[i], cost[i], `=B${r}-C${r}`, `=D${r}/B${r}`])
})
rows.push(['合计', '=SUM(B2:B10)', '=SUM(C2:C10)', '=SUM(D2:D10)', '=D11/B11'])
let xl = createXlsx([
  { name: '经营数据', header: true, rows, widths: [10, 12, 12, 12, 12] },
  { name: '预算', header: true, rows: [['项目', '全年预算', '已执行', '执行率'], ['人力', 1680, 1215, '=C2/B2'], ['市场', 620, 468, '=C3/B3'], ['研发', 560, 402, '=C4/B4']] }
]).buffer
const ops = [
  { op: 'style', sheet: '经营数据', range: 'B2:D11', format: '#,##0' },
  { op: 'style', sheet: '经营数据', range: 'E2:E11', format: '0.0%' },
  { op: 'style', sheet: '经营数据', range: 'A11:E11', bold: true, fill: 'F2F2F2' },
  { op: 'style', sheet: '预算', range: 'D2:D4', format: '0.0%' },
  { op: 'add_chart', sheet: '经营数据', range: 'A1:C10', type: 'col', title: '月度收入与成本（万元）', at: { c: 7, r: 2 }, w: 520, h: 300 }
]
const out = editXlsxUi(xl, ops)
xl = out.buffer ?? out
fs.writeFileSync(path.join(ws, '预算.xlsx'), xl)

const pp = await createPptx([
  { layout: 'title', title: '2026 年第三季度经营回顾', subtitle: '战略与运营部 · 2026 年 10 月', notes: '开场：先讲结论，再讲分区和下季度计划。' },
  { title: '核心结论', bullets: ['营业收入 1,286 万元，同比增长 13.4%', { text: '华东、华南贡献六成以上增量', level: 0 }, { text: '华东新增签约门店 46 家', level: 1 }, '毛利率 41.5%，环比提高 1.3 个百分点', '新产品线首次实现盈亏平衡'], notes: '强调毛利率回升的原因是采购成本下降。' },
  { title: '分区经营情况', table: { rows: [['大区', '收入（万元）', '同比', '毛利率'], ['华东', '486', '18.2%', '43.5%'], ['华南', '352', '15.7%', '42.1%'], ['华北', '268', '7.9%', '38.6%'], ['西部', '180', '6.3%', '37.2%']] } },
  { title: '收入走势', image: 'chart.png' },
  { title: '下季度重点', left: ['推进华北大客户交付', '新产品线 3 个大区试点'], right: ['销售费用率控制在 15% 以内', '研发投入保持 12% 左右'] }
], { accent: '1F4E79' }, io)
fs.writeFileSync(path.join(ws, '经营回顾.pptx'), pp.buffer)
