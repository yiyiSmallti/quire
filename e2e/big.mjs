// Large files for the web E2E: node big.mjs LIB WS
// A 20 000-row workbook whose formulas carry no results (as openpyxl writes
// them) plus a sheet that checks tricky functions against Excel's answers,
// a 40-chapter document and a 120-slide deck. Deterministic (seeded).
import fs from 'node:fs'
import path from 'node:path'
const [lib, ws] = process.argv.slice(2)
const { createDocx } = await import(path.join(lib, 'docx.js'))
const { createXlsx } = await import(path.join(lib, 'xlsx.js'))
const { createPptx } = await import(path.join(lib, 'pptx.js'))
const { Package } = await import(path.join(lib, 'pkg.js'))
const { descendants } = await import(path.join(lib, 'xml.js'))
let seed = 7
const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648
const pick = (a) => a[Math.floor(rnd() * a.length)]
const REG = ['华东', '华南', '华北', '西南', '东北']
const N = 20000
const last = N + 1

const prod = [['编码', '名称', '单价']]
for (let i = 1; i <= 200; i++) prod.push([`P${String(i).padStart(3, '0')}`, `产品${String(i).padStart(3, '0')}`, Math.round((5 + rnd() * 495) * 100) / 100])
const det = [['日期', '区域', '编码', '名称', '单价', '数量', '金额', '累计', '等级', '月份', '折扣']]
for (let r = 2; r <= last; r++) {
  det.push([45292 + Math.floor(rnd() * 366), pick(REG), `P${String(1 + Math.floor(rnd() * 200)).padStart(3, '0')}`,
    `=VLOOKUP(C${r},产品!$A$2:$C$201,2,FALSE)`, `=INDEX(产品!$C$2:$C$201,MATCH(C${r},产品!$A$2:$A$201,0))`, 1 + Math.floor(rnd() * 119),
    `=ROUND(E${r}*F${r}*K${r},2)`, r === 2 ? '=G2' : `=H${r - 1}+G${r}`, `=IFS(G${r}>=20000,"A",G${r}>=5000,"B",TRUE,"C")`,
    `=TEXT(A${r},"yyyy-mm")`, `=IF(AND(B${r}="华东",F${r}>50),0.9,IF(OR(B${r}="华南",B${r}="华北"),0.95,1))`])
}
const months = Array.from({ length: 12 }, (_, i) => `2024-${String(i + 1).padStart(2, '0')}`)
const col = (j) => String.fromCharCode(66 + j)
const sum = [['区域', ...months, '合计', '笔数', '均额', '最大单笔', '排名']]
REG.forEach((g, k) => {
  const i = k + 2
  sum.push([g, ...months.map((_, j) => `=SUMIFS(明细!$G$2:$G$${last},明细!$B$2:$B$${last},$A${i},明细!$J$2:$J$${last},${col(j)}$1)`),
    `=SUM(B${i}:M${i})`, `=COUNTIFS(明细!$B$2:$B$${last},$A${i})`, `=ROUND(N${i}/O${i},2)`, `=MAXIFS(明细!$G$2:$G$${last},明细!$B$2:$B$${last},$A${i})`, `=RANK(N${i},$N$2:$N$6)`])
})
sum.push(['总计', ...Array.from({ length: 13 }, (_, j) => `=SUM(${col(j)}2:${col(j)}6)`), '=SUM(O2:O6)'])
sum.push(['核对', `=ROUND(N7-明细!H${last},2)`, '应为 0'])
sum.push(['SUMPRODUCT', `=ROUND(SUMPRODUCT((明细!$B$2:$B$${last}="华东")*明细!$G$2:$G$${last})-N2,2)`, '应为 0'])

const d = (y, m, dd) => Date.UTC(y, m - 1, dd) / 86400000 + 25569
const CASES = [
  ['SUMIFS 单条件', '=SUMIFS(H2:H11,I2:I11,"华东")', 210], ['SUMIF 通配符', '=SUMIF(I2:I11,"*东",H2:H11)', 210],
  ['COUNTIF 前缀通配', '=COUNTIF(I2:I11,"华*")', 8], ['COUNTIF 比较', '=COUNTIF(H2:H11,">=50")', 6],
  ['COUNTIF 不等于', '=COUNTIF(I2:I11,"<>华东")', 6], ['COUNTIFS 拼接条件', '=COUNTIFS(I2:I11,"华东",H2:H11,">"&H4)', 3],
  ['AVERAGEIFS', '=AVERAGEIFS(H2:H11,I2:I11,"华南")', 45], ['MAXIFS', '=MAXIFS(H2:H11,I2:I11,"华北")', 90],
  ['SUMPRODUCT 多数组', '=SUMPRODUCT((I2:I11="华东")*(H2:H11>30)*H2:H11)', 200], ['INDEX+MATCH', '=INDEX(H2:H11,MATCH("西南",I2:I11,0))', 50],
  ['INDEX 二维', '=INDEX(H2:I11,4,2)', '华东'], ['VLOOKUP 近似', '=VLOOKUP(55,H2:H11,1,TRUE)', 50],
  ['MATCH 近似', '=MATCH(35,H2:H11,1)', 3], ['XLOOKUP', '=XLOOKUP("东北",I2:I11,H2:H11,"无")', 80],
  ['IFNA+XLOOKUP', '=IFNA(XLOOKUP("西北",I2:I11,H2:H11),"未找到")', '未找到'], ['IFERROR 除零', '=IFERROR(1/0,"除零")', '除零'],
  ['TEXT 千分位', '=TEXT(1234.567,"#,##0.00")', '1,234.57'], ['TEXT 百分比', '=TEXT(0.256,"0.0%")', '25.6%'],
  ['TEXT 中文日期', '=TEXT(DATE(2024,3,5),"yyyy年m月d日")', '2024年3月5日'], ['TEXT 补零', '=TEXT(7,"000")', '007'],
  ['TEXT 货币与单位', '=TEXT(-1234.5,"¥#,##0.0;(¥#,##0.0)")&TEXT(12.5,"0.00""元""")', '(¥1,234.5)12.50元'],
  ['DATE 月份溢出', '=DATE(2024,14,1)', d(2025, 2, 1)], ['EOMONTH 闰年', '=EOMONTH(DATE(2024,1,31),1)', d(2024, 2, 29)],
  ['EDATE 月末', '=EDATE(DATE(2024,1,31),1)', d(2024, 2, 29)], ['DATEDIF 年', '=DATEDIF(DATE(2020,2,29),DATE(2024,2,28),"Y")', 3],
  ['DATEDIF MD', '=DATEDIF(DATE(2024,1,15),DATE(2024,3,10),"MD")', 24], ['WEEKDAY 周一起', '=WEEKDAY(DATE(2024,1,1),2)', 1],
  ['负号与乘方', '=-2^2', 4], ['乘方左结合', '=2^3^2', 64], ['文本转数字', '="1"+1', 2], ['逻辑值参与运算', '=TRUE+1', 2],
  ['数字拼接', '=1&2', '12'], ['百分号字面量', '=50%*10', 5], ['ROUND 负位数', '=ROUND(1234.5678,-2)', 1200],
  ['ROUND 半数远离零', '=ROUND(-2.5,0)', -3], ['ROUNDUP 负数', '=ROUNDUP(-1.21,1)', -1.3], ['MOD 负数', '=MOD(-7,3)', 2],
  ['INT 负数', '=INT(-3.2)', -4], ['TRUNC 负数', '=TRUNC(-3.7)', -3], ['CEILING', '=CEILING(2.1,0.5)', 2.5], ['FLOOR', '=FLOOR(7.8,2)', 6],
  ['LEN 中文', '=LEN("中文abc")', 5], ['MID 中文', '=MID("季度经营回顾",3,2)', '经营'], ['SUBSTITUTE 第 N 处', '=SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'],
  ['TEXTJOIN 忽略空', '=TEXTJOIN("、",TRUE,"甲","","乙")', '甲、乙'], ['PROPER', '=PROPER("hello world")', 'Hello World'],
  ['TRIM 内部空格', '=TRIM("  a   b  ")', 'a b'], ['VALUE 千分位', '=VALUE("1,234.5")', 1234.5],
  ['ISNUMBER+SEARCH', '=ISNUMBER(SEARCH("经营","季度经营回顾"))', true], ['FIND 起点', '=FIND("b","abcb",3)', 4],
  ['SWITCH', '=SWITCH(3,1,"一",2,"二",3,"三","其它")', '三'], ['IFS', '=IFS(H6>80,"高",H6>40,"中",TRUE,"低")', '中'],
  ['CHOOSE', '=CHOOSE(2,"甲","乙","丙")', '乙'], ['LARGE', '=LARGE(H2:H11,3)', 80], ['RANK', '=RANK(70,H2:H11)', 4],
  ['MEDIAN', '=MEDIAN(H2:H11)', 55], ['STDEV', '=STDEV(H2:H11)', 30.276503540974915], ['LOG 底数', '=LOG(8,2)', 3],
  ['文本比较不分大小写', '="a"<"B"', true], ['AND/OR 嵌套', '=AND(H2>5,OR(I2="华南",I3="华南"))', true],
  ['COUNTBLANK', '=COUNTBLANK(H2:H13)', 2], ['整列引用', '=SUM(H:H)', 550], ['浮点比较', '=0.1+0.2=0.3', true],
  ['ROWS*COLUMNS', '=ROWS(H2:I11)*COLUMNS(H2:I11)', 20],
  ['嵌套组合', '=IF(SUMIF(I2:I11,"华东",H2:H11)>200,TEXT(SUMIF(I2:I11,"华东",H2:H11)/SUM(H2:H11),"0.0%"),"低")', '38.2%'],
  ['跨表引用', '=产品!A2&"/"&COUNTA(产品!A:A)', 'P001/201'], ['OFFSET', '=OFFSET(H2,2,0)', 30], ['INDIRECT', '=INDIRECT("H"&4)', 30],
  ['LET', '=LET(x,H3,x*3)', 60], ['SUM+OFFSET 动态区域', '=SUM(OFFSET(H2,0,0,3,1))', 60], ['ADDRESS', '=ADDRESS(2,8)', '$H$2']
]
const nums = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
const regs = ['华东', '华南', '华北', '华东', '西南', '华东', '华南', '东北', '华北', '华东']
const n = CASES.length
const chk = [['说明', '公式', '期望', '一致', '', `=COUNTIF(D2:D${n + 1},"✓")&"/"&ROWS(D2:D${n + 1})`, '', '数值', '区域']]
CASES.forEach(([label, f, want], k) => {
  const i = k + 2
  chk.push([label, f, want, `=IF(ISERROR(B${i}),"✗ 错误",IF(B${i}=C${i},"✓","✗"))`, null, null, null, nums[k] ?? null, regs[k] ?? null])
})
const xl = createXlsx([
  { name: '明细', header: true, rows: det, widths: [11, 6, 6, 9, 8, 6, 10, 13, 5, 8, 5] },
  { name: '产品', header: true, rows: prod },
  { name: '汇总', header: true, rows: sum },
  { name: '函数核对', header: true, rows: chk, widths: [18, 40, 12, 8, 2, 10] }
])
// Scripts (openpyxl and the like) write formulas without results: drop them.
const pkg = Package.fromBuffer(xl.buffer)
for (const name of pkg.names().filter((x) => /^xl\/worksheets\/sheet\d+\.xml$/.test(x))) {
  const doc = pkg.xml(name)
  for (const c of descendants(doc.root, 'c')) {
    if (!c.kids.some((k) => k.name === 'f')) continue
    c.kids = c.kids.filter((k) => k.name !== 'v')
    delete c.attrs.t
  }
  pkg.touch(name)
}
fs.writeFileSync(path.join(ws, '大表格.xlsx'), pkg.toBuffer())

let md = '# 年度经营分析报告\n\n'
const para = '本季度公司在华东、华南两大区域继续保持增长，渠道结构进一步优化，新产品线贡献的收入占比提升到三成以上。与此同时，成本控制取得阶段性成果，毛利率同比上升一点二个百分点。'
for (let ch = 1; ch <= 40; ch++) {
  md += `# 第 ${ch} 章 区域与产品分析\n\n`
  for (let sec = 1; sec <= 3; sec++) {
    md += `## ${ch}.${sec} 专题\n\n`
    for (let k = 0; k < 6; k++) md += para.repeat(1 + (k % 3)) + (k === 2 ? ' **重点：环比增长 12.5%**' : '') + '\n\n'
    md += '- 要点一：线上占比提升\n- 要点二：华东增长最快\n\n'
  }
  md += '| 月份 | 收入 | 成本 |\n| --- | ---: | ---: |\n' + Array.from({ length: 12 }, (_, m) => `| ${m + 1} 月 | ${1000 + ch * 10 + m} | ${600 + m * 7} |`).join('\n') + '\n\n'
}
fs.writeFileSync(path.join(ws, '长文档.docx'), (await createDocx(md, {})).buffer)

const slides = [{ layout: 'title', title: '年度经营回顾（大文件）', subtitle: '共 120 页' }]
for (let i = 1; slides.length < 120; i++) {
  if (i % 10 === 1) slides.push({ layout: 'section', title: `第 ${Math.ceil(i / 10)} 部分` })
  else if (i % 3 === 0) slides.push({ title: `数据表 ${i}`, table: { header: true, rows: [['区域', 'Q1', 'Q2'], ...REG.map((g, k) => [g, String(100 + i + k), String(200 + i - k)])] }, notes: `第 ${i} 页备注` })
  else slides.push({ title: `要点 ${i}`, bullets: ['渠道结构进一步优化', { text: '线上占比提升到 42%', level: 1 }, '新产品线贡献三成收入'] })
}
fs.writeFileSync(path.join(ws, '大演示.pptx'), (await createPptx(slides, {}, {})).buffer)
