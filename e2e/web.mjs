// dsh web with the plugin, driven by a headless Chrome: the person edits
// Word, Excel and PowerPoint files in the right Sidebar, the files on disk
// follow, and annotations reach the model with the next message.
import { spawn } from 'node:child_process'
import path from 'node:path'
import fs from 'node:fs'
import { Chrome, Page, chromePath, jsStr, pt } from './chrome.mjs'
import {
  Cleanup, Dsh, E2E, ROOT, attempt, chart, count, findDsh, keep, packPlugin,
  partsLike, run, shortTemp, sleep, startModel, unzip, waitFile, which, zipPartOr
} from './harness.mjs'

const sharedOrInline = (file) => zipPartOr(file, 'xl/sharedStrings.xml') + zipPartOr(file, 'xl/worksheets/sheet1.xml')

// notesOf joins the text of every notes page.
function notesOf (file) {
  let out = ''
  try {
    for (const e of unzip(fs.readFileSync(file))) {
      if (!e.name.startsWith('ppt/notesSlides/notesSlide') || !e.name.endsWith('.xml')) continue
      for (const m of e.data().toString('utf8').matchAll(/<a:t>([^<]*)<\/a:t>/g)) out += m[1]
      out += '\n'
    }
  } catch {}
  return out
}

// startWeb runs dsh web in the workspace on a free port and returns its
// signed-in address; the address carries a token, so it is never logged.
export function startWeb (d, ws, model, cleanup) {
  const [cmd, ...rest] = d.argv('--profile', 'quire-web', '--no-open', '--port', '0')
  const child = spawn(cmd, rest, {
    cwd: ws,
    env: { ...process.env, ...d.env, DEEPSEEK_API_KEY: 'sk-quire-e2e', DEEPSEEK_BASE_URL: model.url },
    stdio: ['ignore', 'pipe', 'pipe']
  })
  const exited = new Promise((resolve) => child.on('close', resolve))
  cleanup.add(async () => { child.kill('SIGKILL'); await exited })
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('dsh web 90 秒内没有打印地址')), 90000)
    let buf = ''
    const scan = (d) => {
      buf += d
      const m = /dsh web: (http:\/\/\S+)/.exec(buf)
      if (m) { clearTimeout(timer); resolve(m[1]) }
      if (buf.length > 1 << 20) buf = buf.slice(-4096)
    }
    child.stdout.on('data', scan)
    child.stderr.on('data', scan)
    child.on('error', (err) => { clearTimeout(timer); reject(err) })
  })
}

// verifyWithPython reads the edited files with openpyxl and python-docx
// when they are installed.
async function verifyWithPython (ws) {
  const py = which('python3')
  if (!py) return ''
  const code = `import sys, os
out = []
try:
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(sys.argv[1], '预算.xlsx'))
    out.append('openpyxl D5=' + str(wb.active['D5'].value))
except ImportError: pass
try:
    import docx
    d = docx.Document(os.path.join(sys.argv[1], '季度报告.docx'))
    out.append('python-docx 末段=' + d.paragraphs[-1].text + '，表格 ' + d.tables[0].rows[1].cells[1].text)
except ImportError: pass
print('；'.join(out))`
  return new Promise((resolve) => {
    const child = spawn(py, ['-I', '-c', code, ws], { stdio: ['ignore', 'pipe', 'ignore'] })
    let out = ''
    child.stdout.on('data', (d) => { out += d })
    child.on('error', () => resolve(''))
    child.on('close', (c) => resolve(c === 0 ? out.trim() : ''))
  })
}

// dismissDialogs clears what a fresh home shows first, a preview notice and
// a sign-in prompt; the editor needs neither. They come up a moment after
// load, one after the other.
export async function dismissDialogs (p) {
  for (const end = Date.now() + 8000; Date.now() < end; await sleep(400)) {
    let v
    try {
      v = await p.js(`const d = [...document.querySelectorAll('[role=dialog]')].find((e) => __q.vis(e)); if (!d) return 'none'; const b = [...d.querySelectorAll('button')].find((b) => b.textContent.trim() === '继续') ?? d.querySelector('button[aria-label="关闭"], button[aria-label="Close"]'); return b ? __q.centre(b) : 'esc'`)
    } catch {}
    if (v === 'esc') await p.key('Escape').catch(() => {})
    else if (v && typeof v === 'object') await p.click(pt(v).x, pt(v).y, 1).catch(() => {})
  }
}

// enterSession opens the Session by its title and the right Sidebar's
// workspace files.
export async function enterSession (p, title) {
  let e
  for (let i = 0; i < 3; i++) {
    e = await attempt(async () => {
      await p.clickText(title, false)
      await p.waitJS(`return document.title.includes(${jsStr(title)})`, 4000)
    })
    if (!e) break
  }
  if (e) throw e
  await sleep(1000)
  if (await p.js(`return !!__q.byText('工作区文件', true)`).catch(() => false) !== true) {
    await p.clickJS(`return __q.centre(document.querySelector('button[aria-label="打开右侧边栏"]'))`)
  }
  await p.clickText('工作区文件', true)
}

// opener returns open, which shows a file of the workspace in the right
// Sidebar.
export function opener (p) {
  return async (name) => {
    if (await p.js(`return !!__q.byText(${jsStr(name)}, true)`).catch(() => false) !== true) {
      // With many files open the side bar's tab strip scrolls; bring
      // its first tab (文件) back into view.
      await p.js(`for (const e of document.querySelectorAll('body *')) { if (e.closest('.quire') || ![...e.childNodes].some((n) => n.nodeType === 3 && n.data.trim() === '文件')) continue; let b = e; while (b && !b.getBoundingClientRect().width) b = b.parentElement; b?.scrollIntoView({ block: 'nearest', inline: 'nearest' }) } return true`).catch(() => {})
      await p.clickText('文件', true)
    }
    await p.clickText(name, true)
  }
}

export async function web (rep) {
  const s = rep.section('dsh web：人在右侧栏直接编辑 Word / Excel / PowerPoint')
  const cleanup = new Cleanup()
  try {
    await suite(s, cleanup)
  } finally {
    await cleanup.run()
  }
}

async function suite (s, cleanup) {
  const dshBin = findDsh()
  const chromeBin = chromePath()
  const node = process.execPath
  if (!dshBin || !chromeBin) {
    s.step('找到 dsh、Chrome 和 node', new Error(`缺少 dsh=${jsStr(dshBin)} chrome=${jsStr(chromeBin)} node=${jsStr(node)}`))
    return
  }
  const lib = path.join(ROOT, 'plugin', 'lib')
  const gen = path.join(E2E, 'gen')
  const work = shortTemp(cleanup)
  const ws = path.join(work, 'ws')
  fs.mkdirSync(ws, { recursive: true })
  const d = new Dsh(dshBin, work)

  // Files, a Session in their folder, and the web profile with quire.
  const model = await startModel(cleanup)
  let err = await attempt(async () => {
    chart(path.join(ws, 'chart.png'))
    await run({}, work, node, [path.join(gen, 'fixture.mjs'), lib, ws])
    await run({}, work, node, [path.join(gen, 'big.mjs'), lib, ws])
    await run({}, work, node, [path.join(gen, 'smartart.mjs'), lib, path.join(ws, '流程.pptx')])
    const tgz = packPlugin(work)
    await d.install('quire-e2e', tgz)
    const turn = await d.turn('quire-e2e', ws, model, '看一下这几个文件')
    if (turn.err) throw turn.err
    await d.run('--profile', 'quire-web', '--from-default-profile', 'web', '--dump-config')
    await d.run('plugin', '--profile', 'quire-web', 'add', tgz)
  })
  if (!s.step('准备：用 quire 写出三个文件，sdk profile 跑一轮建出会话，web 模板 profile 装上本包', err)) return
  let url = ''
  err = await attempt(async () => { url = await startWeb(d, ws, model, cleanup) })
  if (!s.step('启动 dsh web（随机端口）', err)) return
  let c
  err = await attempt(async () => { c = await Chrome.start(chromeBin, work, cleanup) })
  if (!s.step('无头 Chrome', err)) return
  await c.send(c.session, 'Emulation.setDeviceMetricsOverride', { width: 1440, height: 900, deviceScaleFactor: 1, mobile: false }).catch(() => {})
  await c.navigate(url).catch(() => {})
  await c.waitLoaded()
  const p = new Page(c)

  err = await attempt(() => p.waitJS(`return !!(window.__DSH_BOOT__?.entries ?? []).find((e) => e.id === '@precinct/quire') && !!document.getElementById('quire-editor-css')`, 30000))
  if (!s.step('客户端模块：boot 清单里有 @precinct/quire，apply 已运行（样式已注入）', err)) {
    await p.shot('web-fail-boot.png')
    return
  }

  // Open the Session, the right Sidebar's files, then each file.
  await dismissDialogs(p)
  const open = opener(p)
  err = await attempt(async () => {
    await enterSession(p, '看一下这几个文件')
    await open('预算.xlsx')
    await p.waitJS(`return !!document.querySelector('.quire .qx-canvas') && document.querySelector('.qx-tab')?.textContent === '明细'`, 20000)
  })
  if (!s.step('在会话的右侧栏打开 预算.xlsx，显示的是 Quire 编辑器', err)) {
    await p.shot('web-fail-open.png')
    return
  }

  await spreadsheet(s, p, ws, work, lib, node)
  await documentSteps(s, p, ws, open)
  await presentation(s, p, ws, open)
  await objects(s, p, ws, open)
  await bigFiles(s, p, ws, open)
  await pins(s, p, open, model)

  for (const f of ['预算.xlsx', '季度报告.docx', '经营回顾.pptx']) keep(path.join(ws, f), 'web-' + f)
  const out = await verifyWithPython(ws)
  if (out) s.step('不经本插件的读者打开改后的文件（openpyxl / python-docx）', null, out)
}

// ------------------------------------------------------------ spreadsheet

async function spreadsheet (s, p, ws, work, lib, node) {
  const xlsx = path.join(ws, '预算.xlsx')
  const sheet1 = 'xl/worksheets/sheet1.xml'
  let err = await attempt(async () => {
    await p.goCell('A5')
    await p.typeText('税费')
    await p.key('Tab')
    await p.goCell('D5')
    await p.typeText('=D4*0.06')
    await p.key('Enter')
    await waitFile(xlsx, sheet1, (x) => /<c r="D5"[^>]*><f>D4\*0\.06<\/f><v>3264<\/v>/.test(x),
      (x) => /<c r="A5"/.test(x) && sharedOrInline(xlsx).includes('税费'))
  })
  s.step('表格：A5 输入文字、D5 输入公式，自动保存；文件里有公式和重算后的缓存值 3264', err)
  await p.shot('web-xlsx.png')

  err = await attempt(async () => {
    await p.goCell('A1')
    await p.goCell('A4:D5')
    const plain = zipPartOr(xlsx, sheet1)
    await p.key('b', 'mod')
    await waitFile(xlsx, sheet1, (x) => x !== plain)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire .q-side button[aria-label^="撤销"]')))`)
    await waitFile(xlsx, sheet1, (x) => x === plain)
  })
  s.step('表格：选区 ⌘B 加粗写进样式；工具栏撤销把文件恢复到加粗前', err)

  err = await attempt(async () => {
    await run({}, work, node, [path.join(E2E, 'gen', 'agent-edit.mjs'), lib, xlsx])
    await p.waitJS(`return (document.querySelector('.qx-name').value = 'B2', document.querySelector('.qx-name').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })), document.querySelector('.qx-fx').value === '9000')`, 20000)
  })
  s.step('Agent 在外部改了 B2：编辑器自动载入新内容（编辑栏显示 9000）', err)

  // Formula help: typing a name offers functions; Tab takes one.
  err = await attempt(async () => {
    await p.goCell('F6')
    await p.typeText('=SUM')
    await p.waitJS(`const a = document.querySelector('.qx-ac'); return !a.hidden && a.querySelector('.qx-ac-item.on b')?.textContent === 'SUM'`, 5000)
    await p.key('Tab')
    await p.typeText('D2:D3)')
    await p.key('Enter')
    await waitFile(xlsx, sheet1, (x) => /<c r="F6"[^>]*><f>SUM\(D2:D3\)<\/f><v>59400<\/v>/.test(x))
  })
  s.step('表格：输入 =SUM 弹出函数提示，Tab 补全后写成 SUM(D2:D3)，缓存值 59400', err)

  // Sort the two data rows by 项目 ascending: 带宽 comes first and the
  // formulas move with their rows.
  err = await attempt(async () => {
    await p.goCell('A1:D3')
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label="升序排序"]')))`)
    await waitFile(xlsx, sheet1, (x) => /<c r="D2"[^>]*><f>B2\*C2<\/f><v>14400<\/v>/.test(x) && /<c r="D3"[^>]*><f>B3\*C3<\/f><v>45000<\/v>/.test(x))
  })
  s.step('表格：选中 A1:D3 升序排序，带宽排到第 2 行，公式随行移动', err)

  // cellBox reads the top-left corners of two cells' boxes.
  const corner = async (ref) => {
    await p.goCell(ref)
    return pt(await p.js(`const r = document.querySelector('.qx-input').getBoundingClientRect(); return { x: r.left, y: r.top }`))
  }
  const cellBox = async (a, b) => {
    const p1 = await corner(a)
    const p2 = await corner(b)
    return [p1.x, p1.y, p2.x, p2.y]
  }

  // Fill handle: drag G1 (第1周) down two rows.
  err = await attempt(async () => {
    await p.goCell('G1')
    await p.typeText('第1周')
    await p.key('Enter')
    // G1's box runs from its top-left to H2's; select G1 last.
    const [hx, hy, , gy] = await cellBox('H2', 'G1')
    const h = hy - gy
    await p.drag(hx - 1, hy - 1, hx - 1, hy + h * 1.5)
    await waitFile(xlsx, sheet1, (x) => /<c r="G3"/.test(x) && sharedOrInline(xlsx).includes('第2周') && sharedOrInline(xlsx).includes('第3周'))
  })
  s.step('表格：拖动填充柄把 G1「第1周」向下填充，得到 第2周、第3周', err)

  // Borders from the right-click menu.
  const styles0 = zipPartOr(xlsx, 'xl/styles.xml')
  let box = []
  err = await attempt(async () => {
    box = await cellBox('A1', 'B2')
    const [ax1, ay1, ax2, ay2] = box
    await p.goCell('A1:D4')
    await p.rclick((ax1 + ax2) / 2, (ay1 + ay2) / 2)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.startsWith('所有框线')); return b && __q.centre(b)`)
    await waitFile(xlsx, 'xl/styles.xml', (x) => x !== styles0 && x.includes('<left style="thin"'))
  })
  s.step('表格：右键菜单「所有框线」给 A1:D4 加上细边框', err)

  // AutoFilter, then hide 带宽 from the column A filter list.
  err = await attempt(async () => {
    const [, ay1, ax2, ay2] = box
    await p.goCell('A1:D5')
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label="筛选"]')))`)
    await waitFile(xlsx, sheet1, (x) => x.includes('<autoFilter ref="A1:D5"'))
    await p.goCell('C3')
    await sleep(500)
    await p.click(ax2 - 8, (ay1 + ay2) / 2, 1)
    await p.waitJS(`const l = [...document.querySelectorAll('.q-filter-item')].find((e) => e.textContent.startsWith('带宽')); if (!l) return false; l.querySelector('input').click(); return true`, 5000)
    await p.clickJS(`return __q.centre(document.querySelector('.q-filter-ok'))`)
    await waitFile(xlsx, sheet1, (x) => /<row r="2"[^>]*hidden="1"/.test(x))
  })
  s.step('表格：开启筛选，在 A 列筛选面板里去掉「带宽」，第 2 行被隐藏', err)
  await p.shot('web-xlsx-more.png')

  // Charts: insert from a range, drag, change type, delete.
  const charts = () => partsLike(xlsx, 'xl/charts/')
  err = await attempt(async () => {
    await p.goCell('A1:C3')
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="插入图表"]')))`)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '柱形图'); return b && __q.centre(b)`)
    await waitFile(xlsx, '', () => {
      const c = charts()
      return c.includes('<c:barChart') && c.includes("'明细'!$B$2:$B$3") && c.includes("'明细'!$A$2:$A$3")
    }, () => partsLike(xlsx, sheet1).includes('<drawing '))
    await p.waitJS(`return document.querySelectorAll('.qx-chart.on svg rect').length >= 4`, 10000)
  })
  s.step('表格：选中 A1:C3 插入柱形图，系列引用 B2:B3、分类 A2:A3，网格上画出图表', err)
  // Row 2 is hidden by the filter: like Excel, the chart leaves it out.
  err = err ?? await attempt(async () => {
    await p.waitJS(`const t = document.querySelector('.qx-chart svg')?.textContent ?? ''; return t.length > 0 && !t.includes('带宽')`, 10000)
  })
  s.step('表格：筛选隐藏的第 2 行（带宽）不画进图表，系列引用不变', err)
  await p.shot('web-xlsx-chart.png')

  const drawing0 = partsLike(xlsx, 'xl/drawings/drawing')
  err = await attempt(async () => {
    const b = pt(await p.js(`const r = document.querySelector('.qx-chart').getBoundingClientRect(); return { X: r.left + r.width / 2, Y: r.top + r.height / 2 }`))
    await p.drag(b.x, b.y, b.x + 130, b.y + 70)
    await waitFile(xlsx, '', () => {
      const dr = partsLike(xlsx, 'xl/drawings/drawing')
      return dr !== drawing0 && dr.includes('twoCellAnchor')
    })
  })
  s.step('表格：拖动图表，锚点写进 drawing 的 twoCellAnchor', err)

  const chartCorner = `const r = document.querySelector('.qx-chart').getBoundingClientRect(); return { X: r.left + 60, Y: r.top + 50 }`
  err = await attempt(async () => {
    await p.waitJS(`return !!document.querySelector('.qx-chart')`, 10000)
    const b = pt(await p.js(chartCorner))
    await p.rclick(b.x, b.y)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim().endsWith('折线图')); return b && __q.centre(b)`)
    await waitFile(xlsx, '', () => {
      const c = charts()
      return c.includes('<c:lineChart') && !c.includes('<c:barChart')
    })
    await p.waitJS(`return document.querySelectorAll('.qx-chart svg path').length >= 2`, 10000)
  })
  s.step('表格：右键图表改成折线图', err)

  err = err ?? await attempt(async () => {
    await p.waitJS(`return !!document.querySelector('.qx-chart')`, 10000)
    const b = pt(await p.js(chartCorner))
    await p.rclick(b.x, b.y)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '图表选项…'); return b && __q.centre(b)`)
    await p.waitJS(`return !!document.querySelector('.q-dlg[aria-label="图表选项"]')`, 5000)
    await p.js(`const rows = [...document.querySelectorAll('.q-dlg .q-dlg-row')]; const f = (l) => rows.find((r) => r.firstChild.textContent === l).querySelector('select, input')
      const set = (n, v) => { if (n.type === 'checkbox') n.checked = v; else n.value = v; n.dispatchEvent(new Event(n.tagName === 'SELECT' || n.type === 'checkbox' ? 'change' : 'input', { bubbles: true })) }
      set(f('数据标签'), 'value'); set(f('数值轴标题'), '金额'); set(f('数值轴最小值'), '0'); set(f('主要网格线'), false); return true`)
    await p.clickJS(`return __q.centre(document.querySelector('.q-dlg .q-dlg-ok'))`)
    await waitFile(xlsx, '', () => {
      const c = charts()
      return /<c:dLbls>(?:(?!<\/c:dLbls>).)*<c:showVal val="1"\/>/s.test(c) && c.includes('金额') && !c.includes('<c:majorGridlines') && c.includes('<c:min val="0"/>')
    })
    await p.waitJS(`return (document.querySelector('.qx-chart svg')?.textContent ?? '').includes('金额')`, 10000)
  })
  s.step('表格：右键「图表选项…」打开数据标签、写数值轴标题和最小值、关掉网格线，写进图表部件并画出来', err)
  await p.shot('web-xlsx-chart-look.png')

  err = err ?? await attempt(async () => {
    await p.waitJS(`return !!document.querySelector('.qx-chart')`, 10000)
    const b = pt(await p.js(chartCorner))
    await p.click(b.x, b.y, 1)
    await p.key('Delete')
    await waitFile(xlsx, '', () => charts() === '')
    await p.waitJS(`return !document.querySelector('.qx-chart')`, 10000)
  })
  s.step('表格：选中图表按 Delete 删除，图表部件一并移除', err)
}

// --------------------------------------------------------------- document

async function documentSteps (s, p, ws, open) {
  const docx = path.join(ws, '季度报告.docx')
  const doc = 'word/document.xml'
  const before = zipPartOr(docx, doc)
  let err = await attempt(async () => {
    await open('季度报告.docx')
    await p.waitJS(`return document.querySelectorAll('.quire .qd-page .qd-p').length > 3`, 20000)
    await p.js(`const ps = document.querySelectorAll('.qd-page > .qd-p'); const last = ps[ps.length - 1]; document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(last); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.typeText('下季度目标 150 万。')
    await p.key('Enter')
    await p.typeText('新增段落')
    await p.js(`const ps = document.querySelectorAll('.qd-page > .qd-p'); const r = document.createRange(); r.selectNodeContents(ps[ps.length - 1]); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="加粗"]')))`)
    await waitFile(docx, doc, (x) => x.includes('继续保持增长。下季度目标 150 万。') && /<w:b\/>.{0,200}?<w:t[^>]*>新增段落<\/w:t>/.test(x),
      (x) => count(x, '<w:drawing') === count(before, '<w:drawing') && count(x, '<w:drawing') > 0)
  })
  s.step('文档：段尾打字、回车分段、选中加粗，自动保存；图片原样保留', err)

  err = await attempt(async () => {
    await p.js(`const td = document.querySelector('.qd-t td[data-r="2"][data-c="2"]'); td.focus(); const r = document.createRange(); r.selectNodeContents(td); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.typeText('130')
    await waitFile(docx, doc, (x) => x.includes('>130<') && !x.includes('>120<'))
  })
  s.step('文档：改表格单元格（120 → 130）', err)

  const selectPara = `const p = [...document.querySelectorAll('.qd-page > .qd-p')].find((e) => e.textContent.includes('本季度收入')); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(p); getSelection().removeAllRanges(); getSelection().addRange(r); return true`
  const setSelect = (label, value) => `const s = document.querySelector('.quire select[aria-label="${label}"]'); s.value = ${jsStr(value)}; s.dispatchEvent(new Event('change', { bubbles: true })); return true`
  err = await attempt(async () => {
    await p.js(selectPara)
    await p.js(setSelect('字体', '黑体'))
    await p.js(selectPara)
    await p.js(setSelect('行距', '360'))
    await waitFile(docx, doc, (x) => /w:eastAsia="黑体"[^>]*\/>.{0,300}?本季度收入/.test(x) && /w:line="360"/.test(x))
  })
  s.step('文档：选中段落改字体为黑体、行距 1.5 倍，写进 w:rFonts 和 w:spacing', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="查找和替换"]')))`)
    // The box starts with the selected text, selected: typing replaces it.
    await p.waitJS(`const i = document.querySelector('.q-find-in[aria-label="查找"]'); return document.activeElement === i && i.selectionEnd - i.selectionStart === i.value.length`, 5000)
    await p.typeText('华东区')
    await p.waitJS(`return document.querySelector('.q-find-n').textContent === '1/1'`, 5000)
    await p.clickJS(`return __q.centre(document.querySelector('.q-find-in[aria-label="替换为"]'))`)
    await p.typeText('华东地区')
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.q-find button')].find((b) => b.textContent === '全部替换'))`)
    await waitFile(docx, doc, (x) => x.includes('华东地区贡献最大'))
  })
  await p.key('Escape').catch(() => {})
  s.step('文档：⌘F 查找栏显示 1/1，全部替换「华东区」→「华东地区」', err)

  const pics = zipPartOr(docx, doc)
  err = await attempt(async () => {
    await p.js(selectPara)
    await p.js(`const c = document.createElement('canvas'); c.width = 80; c.height = 40; const g = c.getContext('2d'); g.fillStyle = '#d33'; g.fillRect(0, 0, 80, 40); const b = await new Promise((r) => c.toBlob(r, 'image/png')); const dt = new DataTransfer(); dt.items.add(new File([b], 'red.png', { type: 'image/png' })); document.querySelector('.qd-page').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return true`)
    await waitFile(docx, doc, (x) => count(x, '<w:drawing') === count(pics, '<w:drawing') + 1)
    await p.waitJS(`return document.querySelectorAll('.qd-page img.qd-img').length === 2`, 10000)
  })
  s.step('文档：粘贴一张图片，插在当前段落后并保存进文件', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelector('.qd-hdr'))`)
    await p.typeText('内部资料')
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.qd-page > .qd-p')].find((e) => e.textContent.includes('结论')))`)
    await waitFile(docx, '', () => partsLike(docx, 'word/header').includes('内部资料'))
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="页码"]')))`)
    await waitFile(docx, '', () => partsLike(docx, 'word/footer').includes('PAGE'))
  })
  s.step('文档：在页眉写「内部资料」，工具栏 # 在页脚插入页码', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="页面设置"]')))`)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim().endsWith('横向')); return b && __q.centre(b)`)
    await waitFile(docx, doc, (x) => x.includes('w:orient="landscape"'))
    await p.waitJS(`const s = document.querySelector('.qd-sheet'); return parseFloat(s.style.width) > 1000`, 10000)
  })
  s.step('文档：页面设置改为横向，纸张变宽', err)

  // Comments: select words, comment, reply, resolve.
  err = await attempt(async () => {
    await p.js(`const p = [...document.querySelectorAll('.qd-page > .qd-p')].find((e) => e.textContent.includes('本季度收入')); const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode()) && !n.data.includes('本季度收入')); const i = n.data.indexOf('本季度收入'); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 5); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="新建批注"]')))`)
    await p.waitJS(`return document.activeElement === document.querySelector('.qd-draft textarea')`, 5000)
    await p.typeText('请核对数字')
    await p.clickJS(`return __q.centre(document.querySelector('.qd-draft .qd-cok'))`)
    await waitFile(docx, 'word/comments.xml', (x) => x.includes('请核对数字'),
      () => /<w:commentRangeStart w:id="(\d+)"\/>.*本季度收入.*<w:commentRangeEnd w:id="(\d+)"\/>.*<w:commentReference w:id="\d+"\/>/.test(zipPartOr(docx, doc)))
    await p.waitJS(`return document.querySelectorAll('.qd-card').length === 1 && !!document.querySelector('.qd-page .qd-cm[data-cm="r"]') && document.querySelector('.qd-card .qd-cquote').textContent === '本季度收入'`, 10000)
  })
  s.step('文档：选中「本季度收入」新建批注，写进 comments.xml，正文有批注范围和引用', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.qd-card button')].find((b) => b.textContent === '回复'))`)
    await p.waitJS(`return document.activeElement === document.querySelector('.qd-card .qd-cta')`, 5000)
    await p.typeText('已核对')
    await p.clickJS(`return __q.centre(document.querySelector('.qd-card .qd-cok'))`)
    await waitFile(docx, 'word/commentsExtended.xml', (x) => x.includes('paraIdParent'), () => zipPartOr(docx, 'word/comments.xml').includes('已核对'))
    await p.waitJS(`return document.querySelectorAll('.qd-card .qd-creply').length === 1`, 10000)
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.qd-card button')].find((b) => b.textContent === '解决'))`)
    await waitFile(docx, 'word/commentsExtended.xml', (x) => x.includes('w15:done="1"'))
    await p.waitJS(`return document.querySelector('.qd-card.done') !== null`, 10000)
  })
  s.step('文档：回复批注写进 commentsExtended（paraIdParent），再标记为已解决', err)
  await p.shot('web-docx-comments.png')

  // Table of contents: insert at the top, rename a heading, update.
  err = await attempt(async () => {
    await p.js(`const p = document.querySelector('.qd-page > .qd-p'); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="插入目录"]')))`)
    await waitFile(docx, doc, (x) => x.includes('TOC \\o "1-3"') && x.includes('_Toc') && x.indexOf('TOC \\o') < x.indexOf('>季度报告<') &&
      count(x, 'PAGEREF _Toc') === 2 && x.includes('w:leader="dot"'))
    await p.waitJS(`return document.querySelectorAll('.qd-p[data-toc]').length >= 3 && [...document.querySelectorAll('.qd-p[data-toc]')].some((e) => e.textContent.includes('结论')) &&
      document.querySelectorAll('.qd-tocp > .qd-leader').length === 2 && [...document.querySelectorAll('.qd-tocp')].every((e) => /\\t\\d+$/.test(e.textContent))`, 10000)
  })
  s.step('文档：在开头插入目录（TOC 域 + 书签），目录列出「季度报告」「结论」，每项带点线和页码（PAGEREF）', err)

  err = await attempt(async () => {
    await p.js(`const p = [...document.querySelectorAll('.qd-page > .qd-p')].find((e) => !e.dataset.toc && e.textContent === '结论'); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.typeText('与展望')
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="插入目录"]')))`)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '更新目录'); return b && __q.centre(b)`)
    await waitFile(docx, doc, (x) => count(x, '结论与展望') >= 2)
    await p.waitJS(`return [...document.querySelectorAll('.qd-p[data-toc]')].some((e) => e.textContent.includes('结论与展望'))`, 10000)
  })
  s.step('文档：标题改成「结论与展望」后更新目录，目录项跟着变', err)

  // A table copied from a web page or another editor comes in as a table.
  const tables0 = count(zipPartOr(docx, doc), '<w:tbl>')
  err = await attempt(async () => {
    await p.js(selectPara)
    await p.js(`const dt = new DataTransfer(); dt.setData('text/html', '<meta charset="utf-8"><table><thead><tr><th>季度</th><th>收入</th></tr></thead><tbody><tr><td>Q1</td><td>110</td></tr><tr><td>Q2</td><td>120<br>（预估）</td></tr></tbody></table>'); dt.setData('text/plain', '季度\\t收入\\nQ1\\t110\\nQ2\\t120'); document.querySelector('.qd-page').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return true`)
    await waitFile(docx, doc, (x) => count(x, '<w:tbl>') === tables0 + 1 && /<w:t[^>]*>Q2<\/w:t>/.test(x) && /120<\/w:t><\/w:r><w:r>(?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:br\/>/.test(x) && x.includes('<w:tblHeader/>'))
    await p.waitJS(`return [...document.querySelectorAll('.qd-t')].some((t) => t.textContent.includes('Q2') && t.textContent.includes('季度'))`, 10000)
  })
  s.step('文档：粘贴网页里复制的 HTML 表格，插成 Word 表格（表头行、单元格内换行）', err)
  await p.shot('web-docx-more.png')
  await p.shot('web-docx.png')
}

// ----------------------------------------------------------- presentation

async function presentation (s, p, ws, open) {
  const pptx = path.join(ws, '经营回顾.pptx')
  const slide1 = 'ppt/slides/slide1.xml'
  let err = await attempt(async () => {
    await open('经营回顾.pptx')
    await p.waitJS(`return document.querySelectorAll('.quire .qp-thumb').length === 2`, 20000)
    await p.clickJS(`const n = [...document.querySelectorAll('.qp-stage .slide [data-t="txt"]')].find((e) => e.textContent.includes('季度经营回顾')); return __q.centre(n)`, 2)
    await p.key('a', 'mod')
    await p.typeText('季度经营回顾（终稿）')
    await p.key('Escape')
    await waitFile(pptx, slide1, (x) => x.includes('>季度经营回顾（终稿）<') && count(x, '季度经营回顾') === 1)
  })
  s.step('演示：双击标题进入编辑，⌘A 全选后改写，Esc 保存', err)

  err = err ?? await attempt(async () => {
    const y0 = zipPartOr(pptx, slide1)
    const at = await p.js(`const n = [...document.querySelectorAll('.qp-stage .slide [data-t="txt"]')].find((e) => e.textContent.includes('战略部')); return __q.centre(n)`).catch(() => null)
    if (!at) throw new Error('副标题不在页面上')
    const { x, y } = pt(at)
    await p.drag(x, y, x + 40, y + 30)
    await waitFile(pptx, slide1, (x) => x !== y0 && count(x, '<a:off ') > count(y0, '<a:off '))
  })
  s.step('演示：拖动副标题，位置写进 a:xfrm', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelector('.qp-notes'))`)
    await p.js(`const n = document.querySelector('.qp-notes'); n.setSelectionRange(n.value.length, n.value.length); return true`)
    await p.typeText('，先讲结论')
    await p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[1])`)
    await waitFile(pptx, '', () => notesOf(pptx).includes('开场白，先讲结论'))
  })
  s.step('演示：改备注，切换页面时保存（原有备注没有被清掉）', err)

  err = await attempt(async () => {
    await p.js(`const s = document.querySelector('.quire select[aria-label="新建幻灯片"]'); s.value = s.options[2].value; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
    await waitFile(pptx, 'ppt/presentation.xml', (x) => count(x, '<p:sldId ') === 3)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire .q-side button[aria-label^="撤销"]')))`)
    await waitFile(pptx, 'ppt/presentation.xml', (x) => count(x, '<p:sldId ') === 2)
  })
  s.step('演示：用文件自己的版式新建幻灯片，再撤销', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[0])`)
    await p.js(`const s = document.querySelector('.quire select[aria-label="插入形状"]'); s.value = 'star5'; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
    await waitFile(pptx, slide1, (x) => count(x, 'prst="star5"') === 1)
    await p.waitJS(`return !document.querySelector('.qp-box').hidden`, 10000)
    await p.js(`const i = document.querySelector('.quire label[aria-label="填充颜色"] input'); i.value = '#ff0000'; i.dispatchEvent(new Event('input', { bubbles: true })); i.dispatchEvent(new Event('change', { bubbles: true })); return true`)
    await waitFile(pptx, slide1, (x) => x.includes('<a:srgbClr val="FF0000"/>'))
  })
  s.step('演示：插入五角星形状，填充改成红色', err)

  err = await attempt(async () => {
    await p.waitJS(`return !document.querySelector('.qp-box').hidden`, 10000)
    await p.clickJS(`return __q.centre(document.querySelector('.qp-box'))`)
    await p.key('d', 'mod')
    await waitFile(pptx, slide1, (x) => count(x, 'prst="star5"') === 2)
  })
  s.step('演示：⌘D 复制出一个相同的形状', err)

  // The copy sits on top, a little lower: the first star's top point is
  // clear of it, so a Shift-click there adds the first star.
  err = await attempt(async () => {
    await p.clickJS(`const ns = [...document.querySelectorAll('.qp-stage [data-k]')]; return __q.centre(ns[ns.length - 1])`)
    const at = await p.js(`const ns = [...document.querySelectorAll('.qp-stage [data-k]')]; const r = ns[ns.length - 2].getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height * 0.07 }`).catch(() => null)
    if (!at) throw new Error('星形不在页面上')
    await p.shiftClick(pt(at).x, pt(at).y)
    await p.waitJS(`return document.querySelectorAll('.qp-ol').length === 2`, 5000)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="组合所选形状"]')))`)
    await waitFile(pptx, slide1, (x) => count(x, '<p:grpSp>') === 1)
    await p.waitJS(`return document.querySelectorAll('.qp-stage [data-g]').length === 2`, 10000)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="取消组合"]')))`)
    await waitFile(pptx, slide1, (x) => !x.includes('<p:grpSp>') && count(x, 'prst="star5"') === 2)
  })
  s.step('演示：Shift 多选两个星形，组合写进 p:grpSp，再取消组合', err)

  err = await attempt(async () => {
    await p.waitJS(`return document.querySelectorAll('.qp-ol').length === 2`, 10000)
    await p.js(`const s = __q.reveal(document.querySelector('.quire select[aria-label="动画效果"]')); s.value = 'fade'; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
    await waitFile(pptx, slide1, (x) => x.includes('<p:timing>') && count(x, 'presetID="10"') === 2)
    await p.waitJS(`return [...document.querySelectorAll('.qp-badge')].length === 2`, 10000)
    await p.js(`const s = __q.reveal(document.querySelector('.quire select[aria-label="切换效果"]')); s.value = 'push'; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
    await waitFile(pptx, slide1, (x) => x.includes('<p:transition') && x.includes('<p:push'))
  })
  s.step('演示：两个星形加「淡入」动画写进 p:timing（编辑区显示序号），切换效果「推入」写进 p:transition', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre([...document.querySelectorAll('button[aria-label="撤销多步"]')].find((b) => __q.vis(b)))`)
    await p.waitJS(`return [...document.querySelectorAll('.q-menu button')].map((b) => b.textContent).join('|').startsWith('撤销 1 步：切换效果|撤销 2 步：动画')`, 5000)
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.q-menu button')][1])`)
    await waitFile(pptx, slide1, (x) => !x.includes('<p:transition') && !x.includes('<p:timing>') && count(x, 'prst="star5"') === 2)
  })
  s.step('通用：撤销下拉按步列出操作，选「撤销 2 步」一次撤掉切换和动画，形状保留', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="编辑母版与版式"]')))`)
    await p.waitJS(`return document.querySelector('.qp').classList.contains('qp-master') && document.querySelectorAll('.qp-thumb.layout').length > 0`, 5000)
    await p.clickJS(`return __q.centre(document.querySelector('.qp-thumb'))`)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label="插入文本框"]')))`)
    await p.waitJS(`return !!document.querySelector('.qp-editing')`, 10000)
    await p.typeText('内部资料')
    await p.key('Escape')
    await waitFile(pptx, 'ppt/slideMasters/slideMaster1.xml', (x) => x.includes('>内部资料<'))
    await p.clickJS(`return __q.centre(document.querySelector('.qp-mbar button'))`)
    await p.waitJS(`return !document.querySelector('.qp').classList.contains('qp-master') && document.querySelector('.qp-stage .slide').textContent.includes('内部资料')`, 10000)
  })
  s.step('演示：幻灯片母版视图里插入文本框，写进 slideMaster1.xml；回到普通视图，幻灯片上也显示', err)
  await p.shot('web-pptx-master.png')

  err = await attempt(async () => {
    await p.waitJS(`return document.querySelector('.qp-stxt')?.textContent.startsWith('幻灯片 1 / 2')`, 5000)
    await p.clickJS(`return __q.centre([...document.querySelector('.qp-status .q-zoom').children][2])`)
    await p.waitJS(`const t = document.querySelector('.qp-status .q-zoom-p').textContent; const m = /scale\\(([\\d.]+)\\)/.exec(document.querySelector('.qp-scaler>.slide').style.transform); return !!m && Math.round(+m[1] * 100) + '%' === t && document.querySelector('.qp').closest('.quire').querySelector('select[aria-label="缩放"]').value + '%' === t`, 5000)
    await p.clickJS(`return __q.centre(document.querySelector('.qp-status .q-zoom-p'))`)
    await p.waitJS(`return document.querySelector('.qp').closest('.quire').querySelector('select[aria-label="缩放"]').value === ''`, 5000)
  })
  s.step('通用：状态栏显示「幻灯片 1 / 2」，缩放滑块 + 放大一档（比例与画面一致），点百分比回到适应窗口', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="从当前幻灯片放映"]')))`)
    await p.waitJS(`const s = document.querySelector('.qp-show'); return !!s && s.textContent.includes('季度经营回顾')`, 5000)
    await p.shot('web-pptx-show.png')
    await p.key('ArrowRight')
    await p.waitJS(`return document.querySelector('.qp-show')?.textContent.includes('核心要点')`, 5000)
    await p.key('Escape')
    await p.waitJS(`return !document.querySelector('.qp-show')`, 5000)
  })
  s.step('演示：放映从当前页开始，→ 翻到第 2 页，Esc 退出', err)
  await p.shot('web-pptx.png')

  // Emphasis and exit effects: the copy flies out, the first star pulses.
  const setAnim = (v) => `const s = __q.reveal(document.querySelector('.quire select[aria-label="动画效果"]')); s.value = ${jsStr(v)}; s.dispatchEvent(new Event('change', { bubbles: true })); return true`
  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[0])`)
    // The two stars are still selected from the steps before.
    await p.key('Escape')
    await p.waitJS(`return document.querySelector('.qp-box').hidden && document.querySelectorAll('.qp-ol').length === 0`, 5000)
    // The master's text box is drawn on the slide too: find the stars by shape.
    const stars = `[...document.querySelectorAll('.qp-stage [data-k]')].filter((n) => n.outerHTML.slice(0, 3000).includes('polygon('))`
    await p.waitJS(`return ${stars}.length === 2`, 5000)
    await p.clickJS(`const ns = ${stars}; return __q.centre(ns[1])`)
    await p.waitJS(`/* copy */ return !document.querySelector('.qp-box').hidden && document.querySelectorAll('.qp-ol').length === 0`, 5000)
    await p.js(setAnim('exit:fly'))
    await waitFile(pptx, slide1, (x) => x.includes('presetClass="exit"') && x.includes('presetID="2"') && x.includes('<p:strVal val="hidden"/>'))
    const at = await p.js(`const r = ${stars}[0].getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height * 0.07 }`)
    await p.click(pt(at).x, pt(at).y, 1)
    await p.waitJS(`/* first */ return !document.querySelector('.qp-box').hidden && document.querySelectorAll('.qp-ol').length === 0`, 5000)
    await p.js(setAnim('emph:pulse'))
    await waitFile(pptx, slide1, (x) => x.includes('presetClass="emph"') && x.includes('presetID="26"') && x.includes('presetClass="exit"') && x.includes('autoRev="1"'))
    await p.waitJS(`return document.querySelectorAll('.qp-badge').length === 2`, 10000)
  })
  s.step('演示：一个星形加「强调 · 脉冲」、另一个加「退出 · 飞出」，写成 emph / exit 动画（退出结束时隐藏）', err)

  // A blocked audience window: the slides go full screen in this window.
  err = await attempt(async () => {
    await p.js(`window.__open = window.open; window.open = () => null; return true`)
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.qp').closest('.quire').querySelector('button[aria-label^="放映时显示备注"]')))`)
    await p.waitJS(`return !!document.querySelector('.qp-show .qp-pv')`, 5000)
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.qp-show button')].find((b) => b.textContent === '观众窗口'))`)
    await p.waitJS(`const s = document.querySelector('.qp-show'); return s.classList.contains('qp-solo') && !__q.vis(s.querySelector('.qp-pv')) && __q.vis(s.querySelector('.qp-solo-back'))`, 5000)
    await p.shot('web-pptx-solo.png')
    await p.clickJS(`return __q.centre(document.querySelector('.qp-solo-back'))`)
    await p.waitJS(`const s = document.querySelector('.qp-show'); return !s.classList.contains('qp-solo') && __q.vis(s.querySelector('.qp-pv'))`, 5000)
    await p.key('Escape')
    await p.waitJS(`return !document.querySelector('.qp-show')`, 5000)
  })
  await p.js(`if (window.__open) window.open = window.__open; return true`).catch(() => {})
  s.step('演示：浏览器拦截观众窗口时，改在本窗口全屏显示观众画面，右下角按钮切回演示者视图', err)

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.qp').closest('.quire').querySelector('button[aria-label^="新建空白"]')))`)
    await p.waitJS(`return [...document.querySelectorAll('.q-menu button')].some((b) => b.textContent.includes('空白 Word 文档'))`, 5000)
    await p.clickJS(`return __q.centre([...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.includes('空白 Word 文档')))`)
    await p.waitJS(`return document.querySelector('.q-dlg [name=name]')?.value === '新建文档.docx'`, 5000)
    await p.clickJS(`return __q.centre(document.querySelector('.q-dlg .q-dlg-ok'))`)
    await waitFile(path.join(ws, '新建文档.docx'), 'word/document.xml', (x) => x.includes('w:body'))
    await p.waitJS(`return [...document.querySelectorAll('.qd-stxt')].some((n) => __q.vis(n) && n.textContent.includes('共 1 页'))`, 20000)
  })
  s.step('通用：文件 › 新建 › 空白 Word 文档，写进工作区并在新标签里用 Quire 打开', err)
  await p.shot('web-new-docx.png')
}

// --------------------------------------------- objects PowerPoint wrote

// objects edits a deck with a SmartArt diagram, a table and a group, the
// way PowerPoint saves them (gen/smartart.mjs).
async function objects (s, p, ws, open) {
  const pptx = path.join(ws, '流程.pptx')
  const data = 'ppt/diagrams/data1.xml'
  const drawing = 'ppt/diagrams/drawing1.xml'
  const editIn = (sel, text) => async () => {
    await p.clickJS(`return __q.centre(${sel})`, 2)
    await p.waitJS(`return !!document.querySelector('.qp-editing')`, 5000)
    await p.key('a', 'mod')
    await p.typeText(text)
    await p.key('Escape')
  }
  const node = (t) => `[...document.querySelectorAll('.qp-stage [data-sa]')].find((e) => e.textContent.includes(${jsStr(t)}))`
  let err = await attempt(async () => {
    await open('流程.pptx')
    await p.waitJS(`return document.querySelectorAll('.quire .qp-thumb').length === 3 && document.querySelectorAll('.qp-stage [data-sa]').length >= 3`, 20000)
    await editIn(node('方案设计'), '方案评审')()
    await waitFile(pptx, data, (x) => x.includes('>方案评审<') && !x.includes('方案设计'), () => zipPartOr(pptx, drawing).includes('>方案评审<'))
  })
  s.step('SmartArt：双击节点「方案设计」改写，数据模型和 PowerPoint 存的绘图都改了', err)

  // 开发上线 draws its child 灰度发布 as an indented second line.
  err = err ?? await attempt(async () => {
    await p.clickJS(`return __q.centre(${node('开发上线')})`, 2)
    await p.waitJS(`return !!document.querySelector('.qp-editing')`, 5000)
    await p.js(`const n = document.querySelector('.qp-editing'); const t = [...n.querySelectorAll('p')].pop(); const w = document.createTreeWalker(t, NodeFilter.SHOW_TEXT); let x, last; while ((x = w.nextNode())) last = x; const r = document.createRange(); r.selectNodeContents(last); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
    await p.typeText('全量发布')
    await p.key('Escape')
    await waitFile(pptx, data, (x) => x.includes('>全量发布<') && x.includes('>开发上线<') && !x.includes('灰度发布'),
      () => /<a:pPr[^>]*lvl="1"[^>]*\/?>(?:(?!<\/a:p>).)*全量发布/s.test(zipPartOr(pptx, drawing)))
  })
  // The line has no colour of its own: the shape style's fontRef (lt1) makes it white.
  err = err ?? await attempt(async () => {
    await p.waitJS(`const s = [...document.querySelectorAll('.qp-stage [data-sa] span')].find((e) => e.textContent === '全量发布'); return !!s && getComputedStyle(s).color === 'rgb(255, 255, 255)'`, 10000)
  })
  s.step('SmartArt：在父节点的形状里改子节点那一行，子节点的数据跟着改，绘图里仍缩进一级；没有自带颜色的文字按形状样式（fontRef）显示为白色', err)
  await p.shot('web-pptx-smartart.png')

  const slide2 = 'ppt/slides/slide2.xml'
  const cell = (t) => `[...document.querySelectorAll('.qp-stage td[data-r]')].find((e) => e.textContent === ${jsStr(t)})`
  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[1])`)
    await p.waitJS(`return !!${cell('王芳')}`, 10000)
    let at = pt(await p.js(`return __q.centre(${cell('王芳')})`))
    await p.rclick(at.x, at.y)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '在下方插入行'); return b && __q.centre(b)`)
    await waitFile(pptx, slide2, (x) => count(x, '<a:tr ') === 4)
    await p.waitJS(`return !!${cell('王芳')} && document.querySelectorAll('.qp-stage tr').length === 4`, 10000)
    at = pt(await p.js(`return __q.centre(${cell('王芳')})`))
    await p.rclick(at.x, at.y)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '在右侧插入列'); return b && __q.centre(b)`)
    await waitFile(pptx, slide2, (x) => count(x, '<a:gridCol ') === 4 && count(x, '<a:tc>') + count(x, '<a:tc ') === 16)
    await p.waitJS(`return !!${cell('周期')}`, 10000)
    at = pt(await p.js(`return __q.centre(${cell('周期')})`))
    await p.rclick(at.x, at.y)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '删除列'); return b && __q.centre(b)`)
    await waitFile(pptx, slide2, (x) => count(x, '<a:gridCol ') === 3 && !x.includes('>周期<') && x.includes('>李雷<'))
  })
  s.step('表格：右键单元格在下方插入行、在右侧插入列，再删掉「周期」列，行列和网格宽度一起写进文件', err)

  const slide3 = 'ppt/slides/slide3.xml'
  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[2])`)
    await p.waitJS(`return [...document.querySelectorAll('.qp-stage [data-g] [data-t="txt"], .qp-stage [data-g][data-t="txt"]')].some((e) => e.textContent.includes('第二段'))`, 10000)
    await editIn(`[...document.querySelectorAll('.qp-stage [data-t="txt"]')].find((e) => e.textContent.includes('第二段'))`, '组合里改过的第二段')()
    await waitFile(pptx, slide3, (x) => x.includes('>组合里改过的第二段<') && count(x, '<p:grpSp>') === 1 && x.includes('>组合里的第一段<'))
  })
  s.step('组合：双击组合里的文本框直接改字，组合保持不变', err)
  keep(pptx, 'web-流程.pptx')
}

// ------------------------------------------------------------ large files

const secs = (t0) => ((Date.now() - t0) / 1000).toFixed(1)

async function bigFiles (s, p, ws, open) {
  const big = path.join(ws, '大表格.xlsx')
  let t0 = Date.now()
  let took = ''
  let err = await attempt(async () => {
    try {
      await open('大表格.xlsx')
      await p.waitJS(`const Q = __q.cur(); return !!Q?.querySelector('.qx-canvas') && [...Q.querySelectorAll('.qx-tab')].some((t) => t.textContent === '函数核对')`, 60000)
    } finally { took = secs(t0) }
    await p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qx-tab')].find((t) => t.textContent === '函数核对'))`)
    await p.clickJS(`return __q.centre(__q.reveal(__q.cur().querySelector('button[aria-label^="查找和替换"]')))`)
    await p.waitJS(`return document.activeElement?.matches('.q-find-in[aria-label="查找"]')`, 5000)
    await p.typeText('71/71')
    // Find looks at what cells show: only the tally cell shows 71/71.
    await p.waitJS(`return __q.cur().querySelector('.q-find-n').textContent === '1/1'`, 10000)
  })
  await p.key('Escape').catch(() => {})
  s.step(`大文件：脚本生成、公式不带结果的 2 万行表格打开时算出全部公式，用时 ${took} 秒；函数核对页 71 个难点函数与 Excel 的答案一致（71/71）`, err)
  await p.shot('web-big-xlsx.png')

  took = '0.0'
  err = await attempt(async () => {
    await p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qx-tab')].find((t) => t.textContent === '明细'))`)
    await p.goCell('F2')
    await p.typeText('7')
    t0 = Date.now()
    try {
      await p.key('Enter')
      await waitFile(big, 'xl/worksheets/sheet1.xml', (x) => /<c r="F2"[^>]*><v>7<\/v>/.test(x))
    } finally { took = secs(t0) }
    await waitFile(big, 'xl/worksheets/sheet3.xml', (x) => /<c r="B8"[^>]*><f>[^<]+<\/f><v>0<\/v>/.test(x) && /<c r="B9"[^>]*><f>[^<]+<\/f><v>0<\/v>/.test(x))
  })
  s.step(`大文件：改 F2 后整本重算并写回（${took} 秒），汇总页 SUMIFS 合计与明细累计、SUMPRODUCT 两项核对都为 0`, err)

  t0 = Date.now()
  err = await attempt(async () => {
    await open('长文档.docx')
    await p.waitJS(`const m = /共 (\\d+) 页/.exec(__q.cur()?.querySelector('.qd-stxt')?.textContent ?? ''); return !!m && +m[1] >= 50`, 60000)
  })
  took = secs(t0)
  const pages = await p.js(`return /共 (\\d+) 页/.exec(__q.cur().querySelector('.qd-stxt').textContent)[1]`).catch(() => '')
  s.step(`大文件：长文档按纸张分成 ${pages} 页，打开并排好版用时 ${took} 秒`, err)

  t0 = Date.now()
  took = ''
  err = await attempt(async () => {
    try {
      await open('大演示.pptx')
      await p.waitJS(`return __q.cur()?.querySelectorAll('.qp-thumb').length === 120`, 60000)
    } finally { took = secs(t0) }
    await p.clickJS(`return __q.centre(__q.cur().querySelectorAll('.qp-thumb')[119])`)
    await p.waitJS(`return __q.cur().querySelector('.qp-stxt').textContent.includes('120 / 120')`, 10000)
  })
  s.step(`大文件：120 页演示文稿打开用时 ${took} 秒，缩略图栏翻到最后一页`, err)
  await p.shot('web-big-pptx.png')
}

// ------------------------------------------------------------------- pins

// Notes on spots in the files, as cards above the chat composer.
async function pins (s, p, open, model) {
  let err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin'))`)
    await p.waitJS(`return __q.cur().matches('.q-pinning') && !!__q.cur().querySelector('.q-pintip') && __q.cur().querySelector('.q-btn.q-pin').matches('.arm')`, 5000)
    await p.dragJS(`const r = __q.cur().querySelector('.qp-stage .slide [data-k="2"]').getBoundingClientRect(); return { x1: r.left - 6, y1: r.top - 6, x2: r.right + 6, y2: r.bottom + 6 }`)
    await p.waitJS(`return document.activeElement?.matches('.q-pin-in') && !__q.cur().matches('.q-pinning') && !!__q.cur().querySelector('.q-pinmarq')`, 5000)
    // The box's picture shows in the popover and goes with the pin.
    await p.waitJS(`return __q.cur().querySelector('.q-pinpop .q-pin-shot')?.src.startsWith('data:image/jpeg')`, 10000)
    await p.shot('web-pin-box.png')
    await p.typeText('这页要点太笼统，补上数字')
    await p.key('Enter')
    await p.waitJS(`const l = __q.pins(); return l.length === 1 && l[0].startsWith('大演示.pptx · 第 120 页 [s120.2]「') && l[0].endsWith('：这页要点太笼统，补上数字') && !__q.draft().includes('大演示') && document.querySelector('.qpd-pill')?.textContent.startsWith('1 条标注') && document.querySelector('.qpd-list').hidden && document.querySelector('.qpd-card .qpd-thumb')?.src.startsWith('data:image/') && __q.cur().querySelectorAll('.qp-pin').length === 1 && __q.cur().querySelector('.q-pin-n').textContent === '1' && !!__q.cur().querySelectorAll('.qp-thumb')[119].matches('.q-pinned')`, 10000)
  })
  s.step('标注：点工具栏的「标注」进入框选，在第 120 页拖框框住文本框后写下问题，弹窗里带上框选区域的截图；回车后聊天框上方只多一个「1 条标注」胶囊，卡片「大演示.pptx · 第 120 页 [s120.2]「原文」：问题」连同截图收在胶囊里，聊天框文字不变；形状描上虚线框、缩略图带圆点', err)

  err = await attempt(async () => {
    await open('长文档.docx')
    await p.waitJS(`const Q = __q.cur(); return !!Q?.querySelector('.qd-nav:not([hidden])') && Q.querySelectorAll('.qd-nav-i').length > 100`, 30000)
    await p.clickJS(`const it = [...__q.cur().querySelectorAll('.qd-nav-i')].find((x) => x.textContent === '第 30 章 区域与产品分析'); it?.scrollIntoView({ block: 'center' }); return __q.centre(it)`)
    await p.waitJS(`const Q = __q.cur(); const m = /第 (\\d+) 页/.exec(Q.querySelector('.qd-stxt').textContent); return Q.querySelector('.qd-nav-i.on')?.textContent === '第 30 章 区域与产品分析' && +m[1] > 50`, 5000)
    await p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qd-nav-i')].find((x) => x.textContent === '第 30 章 区域与产品分析').querySelector('.qd-nav-tw'))`)
    await p.waitJS(`return __q.cur().querySelectorAll('.qd-nav-i[hidden]').length === 3`, 5000)
  })
  const navPage = await p.js(`return /第 (\\d+) 页/.exec(__q.cur().querySelector('.qd-stxt').textContent)[1]`).catch(() => '')
  s.step(`文档：导航窗格列出长文档的 161 个标题，点「第 30 章」跳到第 ${navPage} 页并高亮这一章，折叠后它下面的 3 个小节收起`, err)
  await p.shot('web-docx-nav.png')

  // Box the words in the chapter's first paragraph that holds them.
  let span = ''
  err = await attempt(async () => {
    await p.js(`const Q = __q.cur(); const h = [...Q.querySelectorAll('.qd-page > .qd-p')].find((x) => x.textContent === '第 30 章 区域与产品分析'); let p = h.nextElementSibling; while (p && (p.dataset.h || !p.textContent.includes('毛利率'))) p = p.nextElementSibling; p.scrollIntoView({ block: 'center' }); Q.querySelector('.qd-page').focus({ preventScroll: true }); window.__words = p; return true`)
    await p.key('m', 'mod', 'shift')
    await p.waitJS(`return __q.cur().matches('.q-pinning')`, 5000)
    // A box from the first heading line down into the next paragraph names both blocks.
    await p.dragJS(`const h = [...__q.cur().querySelectorAll('.qd-page > .qd-p')].find((x) => x.textContent === '第 30 章 区域与产品分析'); const a = h.getBoundingClientRect(); const b = h.nextElementSibling.getBoundingClientRect(); return a.top > 0 && b.bottom < innerHeight ? { x1: a.left + 2, y1: a.top + 2, x2: b.left + 80, y2: b.top + 8 } : null`)
    span = await p.waitJS(`const t = __q.cur().querySelector('.q-pinpop .q-pin-h b')?.textContent; return /^\\[p\\d+\\]–\\[p\\d+\\]$/.test(t ?? '') && t`, 5000)
    await p.key('Escape')
    await p.waitJS(`return !__q.cur().querySelector('.q-pinpop') && !__q.cur().querySelector('.q-pinmarq')`, 5000)
  })
  s.step(`标注：⇧⌘M 进入框选，框住标题和下一段，标注写成首尾两个段落编号 ${span}；Esc 关掉不留痕迹`, err)

  err = err ?? await attempt(async () => {
    await p.key('m', 'mod', 'shift')
    await p.dragJS(`const p = window.__words; const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode()) && !n.data.includes('毛利率')); const i = n.data.indexOf('毛利率'); const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 7); const b = r.getBoundingClientRect(); return { x1: b.left + 1, y1: b.top + 2, x2: b.right - 1, y2: b.bottom - 2 }`)
    await p.waitJS(`return document.activeElement?.matches('.q-pin-in') && /^\\[p\\d+\\]$/.test(__q.cur().querySelector('.q-pin-h b').textContent)`, 5000)
    await p.typeText('写出具体百分点')
    await p.key('Enter')
    await p.waitJS(`const l = __q.pins(); return l.length === 2 && /^长文档\\.docx · \\[p\\d+\\]「毛利率同比上升」：写出具体百分点$/.test(l[1]) && document.querySelectorAll('.qpd-card')[1].querySelector('.qpd-thumb')?.src.startsWith('data:image/') && [...(CSS.highlights.get('q-pin') ?? [])].some((r) => r.toString() === '毛利率同比上升')`, 10000)
  })
  s.step('标注：在文档里框住「毛利率同比上升」再标注，胶囊里累积到 2 条，第二条带 office_read 的段落编号 [pN] 和这段文字的截图，原文在文档里标出', err)

  let cells = ''
  err = await attempt(async () => {
    await open('大表格.xlsx')
    await p.waitJS(`return !!__q.cur()?.querySelector('.qx-canvas')`, 30000)
    await p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qx-tab')].find((t) => t.textContent === '汇总'))`)
    // The cell box sits on the active cell: read where B3 and D5 are.
    const cellAt = `const r = __q.cur().querySelector('.qx-input').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }`
    await p.goCell('B3')
    const b3 = pt(await p.js(cellAt))
    await p.goCell('D5')
    const d5 = pt(await p.js(cellAt))
    await p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin'))`)
    // From the middle of B3 to the middle of D5.
    await p.dragJS(`const a = ${JSON.stringify(b3)}; const z = ${JSON.stringify(d5)}; return { x1: a.x, y1: a.y, x2: z.x, y2: z.y }`)
    cells = await p.waitJS(`const t = __q.cur().querySelector('.q-pinpop .q-pin-h b')?.textContent; return /^汇总!B3:D5$/.test(t ?? '') && __q.cur().querySelector('.q-pinpop .q-pin-q')?.textContent && t`, 5000)
    await p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.q-pinpop .q-btn')].find((b) => b.textContent === '取消'))`)
    await p.waitJS(`return !__q.cur().querySelector('.q-pinpop')`, 5000)
  })
  s.step(`标注：表格里点「标注」拖框，框到的单元格选中并写成 ${cells}，带上框内的值；取消后不加卡片`, err)

  err = err ?? await attempt(async () => {
    await p.goCell('N2')
    const at = pt(await p.js(`const r = __q.cur().querySelector('.qx-input').getBoundingClientRect(); return { x: r.left + 10, y: r.top + 8 }`))
    await p.rclick(at.x, at.y)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.startsWith('标注问题')); return b && __q.centre(b)`)
    await p.waitJS(`return document.activeElement?.matches('.q-pin-in')`, 5000)
    await p.typeText('华东合计和上月对不上')
    await p.clickJS(`return __q.centre(__q.cur().querySelector('.q-pin-ok'))`)
    await p.waitJS(`const l = __q.pins(); return l.length === 3 && /^大表格\\.xlsx · 汇总!N2「[\\d,.]+」（公式 =SUM\\(B2:M2\\)）：华东合计和上月对不上$/.test(l[2]) && document.querySelector('.qpd-pill').textContent.startsWith('3 条标注') && document.querySelector('.qpd-list').hidden`, 10000)
  })
  s.step('标注：表格单元格右键「标注问题…」，第三条写明 汇总!N2、显示的值和公式；三个文件的标注都收在聊天框上方的「3 条标注」胶囊里，不占地方', err)
  await p.shot('web-pins.png')

  err = await attempt(async () => {
    await p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin + .q-more'))`)
    await p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.startsWith('移除这个文件的标注')); return b && __q.centre(b)`)
    await p.waitJS(`return __q.pins().length === 2 && __q.cur().querySelector('.q-pin-n').hidden`, 10000)
  })
  s.step('标注：从标注菜单移除这个表格的标注，聊天框上方剩下另外两张卡片，表格上的标记随之消失', err)

  // Open the pill, edit the first card's note, drop the second card, then send.
  err = err ?? await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelector('.qpd-tog'))`)
    await p.waitJS(`return !document.querySelector('.qpd-list').hidden && document.querySelector('.qpd-tog').getAttribute('aria-expanded') === 'true'`, 5000)
    await p.shot('web-pins-open.png')
    await p.clickJS(`return __q.centre(document.querySelector('.qpd-card .qpd-btn[aria-label="修改说明"]'))`)
    await p.waitJS(`return document.activeElement?.matches('.qpd-in')`, 5000)
    await p.waitJS(`const n = document.activeElement; return n.selectionStart === 0 && n.selectionEnd === n.value.length && n.value.length > 0`, 5000)
    await p.typeText('要点改成三条，每条带数字')
    await p.key('Enter')
    await p.waitJS(`const l = __q.pins(); return l.length === 2 && /：要点改成三条，每条带数字$/.test(l[0]) && !l[0].includes('笼统') && !document.querySelector('.qpd-in')`, 10000)
    await p.clickJS(`return __q.centre(document.querySelectorAll('.qpd-card')[1]?.querySelector('.qpd-btn[aria-label="移除这条标注"]'))`)
    await p.waitJS(`return __q.pins().length === 1 && document.querySelector('.qpd-pill').textContent.startsWith('1 条标注')`, 10000)
  })
  s.step('标注卡片：点胶囊展开列表，每条带截图；点 ✎ 改写第一张卡片的说明，回车保存；点第二张卡片的 × 移除它，胶囊变成「1 条标注」', err)

  let seen = ''
  let images = 0
  err = err ?? await attempt(async () => {
    await p.clickJS(`return __q.centre(document.querySelector('[data-lexical-editor="true"]'))`)
    await p.typeText('按标注改一下')
    await p.key('Enter')
    await p.waitJS(`return !document.querySelector('.qpd') && !__q.cur().querySelector('.qp-pin')`, 30000)
    for (const end = Date.now() + 30000; Date.now() < end; await sleep(300)) {
      seen = model.prompt
      images = model.images
      if (seen.includes('按标注改一下')) break
    }
    const want = '1. 大演示.pptx · 第 120 页 [s120.2]「'
    if (!seen.includes('按标注改一下') || !seen.includes(want) || !seen.includes('：要点改成三条，每条带数字') || seen.includes('写出具体百分点')) {
      throw new Error(`模型收到的内容不对：${JSON.stringify(seen)}`)
    }
    if (images !== 1 || !seen.includes('（下面是这处的截图）')) {
      throw new Error(`模型应收到 1 张标注截图，实际 ${images} 张：${JSON.stringify(seen)}`)
    }
  })
  const sent = seen.split('\n').filter((l) => l.startsWith('1. 大演示.pptx')).pop() ?? ''
  s.step(`发送：在聊天框输入「按标注改一下」回车，卡片和编辑器里的标记随即清空；模型在同一轮收到这句话、附带的标注 ${JSON.stringify(sent)} 和 ${images} 张框选截图，删掉的那条没有发出`, err)
  await p.shot('web-pins-sent.png')
}
