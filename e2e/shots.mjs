// The README screenshots: dsh web with the plugin on a sample workspace
// (gen/demo.mjs), shot by a headless Chrome into docs/images/.
//
// node e2e/shots.mjs
// Needs what the web run needs: dsh and Chrome or Edge.
import fs from 'node:fs'
import path from 'node:path'
import { Chrome, Page, chromePath, jsStr } from './chrome.mjs'
import { Cleanup, Dsh, E2E, ROOT, chart, findDsh, packPlugin, run, shortTemp, sleep, startModel } from './harness.mjs'
import { dismissDialogs, enterSession, opener, startWeb } from './web.mjs'

const OUT = path.join(ROOT, 'docs', 'images')
const PROMPT = '看看第三季度的经营数据'
const REPLY = '第三季度（7–9 月）收入 1,286 万元，毛利 534 万元，毛利率 41.5%，比第二季度高 1.3 个百分点；9 月收入 436 万元，是今年单月最高。季度报告和经营回顾里的数字与表格一致。'

const cleanup = new Cleanup()
try {
  await shots()
} finally {
  await cleanup.run()
}

async function shots () {
  const dshBin = findDsh()
  const chromeBin = chromePath()
  if (!dshBin || !chromeBin) throw new Error(`缺少 dsh=${jsStr(dshBin)} chrome=${jsStr(chromeBin)}`)
  const lib = path.join(ROOT, 'plugin', 'lib')
  const work = shortTemp(cleanup)
  const ws = path.join(work, '季度经营')
  fs.mkdirSync(ws, { recursive: true })
  const d = new Dsh(dshBin, work)

  chart(path.join(ws, 'chart.png'))
  await run({}, work, process.execPath, [path.join(E2E, 'gen', 'demo.mjs'), lib, ws])
  const model = await startModel(cleanup, [{ name: 'office_read', input: { file_path: '预算.xlsx' } }], REPLY)
  const tgz = packPlugin(work)
  await d.install('quire-demo', tgz)
  const turn = await d.turn('quire-demo', ws, model, PROMPT)
  if (turn.err) throw turn.err
  await d.run('--profile', 'quire-web', '--from-default-profile', 'web', '--dump-config')
  await d.run('plugin', '--profile', 'quire-web', 'add', tgz)

  const url = await startWeb(d, ws, model, cleanup)
  const c = await Chrome.start(chromeBin, work, cleanup)
  await c.send(c.session, 'Emulation.setDeviceMetricsOverride', { width: 1440, height: 900, deviceScaleFactor: 2, mobile: false })
  await c.navigate(url)
  await c.waitLoaded()
  const p = new Page(c)
  await p.waitJS(`return !!document.getElementById('quire-editor-css')`, 30000)
  await dismissDialogs(p)
  await enterSession(p, PROMPT)
  const open = opener(p)
  fs.mkdirSync(OUT, { recursive: true })
  // The side bar names the file by its absolute path, a temp folder here.
  const real = fs.realpathSync(work)
  const tidy = () => p.js(`for (const e of document.querySelectorAll('body *')) for (const n of e.childNodes) if (n.nodeType === 3) for (const w of ${JSON.stringify([real, work])}) if (n.data.startsWith(w)) n.data = '~' + n.data.slice(w.length); return true`)
  const shot = async (name) => {
    await tidy()
    await sleep(600)
    const { data } = await c.send(c.session, 'Page.captureScreenshot', { format: 'webp', quality: 88 })
    fs.writeFileSync(path.join(OUT, name), Buffer.from(data, 'base64'))
    console.log('wrote docs/images/' + name)
  }
  const full = async (on) => {
    if (await p.js(`return __q.cur().matches('.q-full')`) !== on) await p.clickJS(`return __q.centre(__q.cur().querySelector('button[title="全屏"], button[title^="退出全屏"]'))`)
    await sleep(800)
  }

  // Overview: the chat beside the report, one annotation pinned and open.
  await open('季度报告.docx')
  await p.waitJS(`return !!__q.cur()?.querySelector('.qd-page')`, 20000)
  await sleep(1000)
  await p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin'))`)
  await p.waitJS(`return __q.cur().matches('.q-pinning')`, 5000)
  await p.dragJS(`const q = [...__q.cur().querySelectorAll('.qd-page > .qd-p')].find((x) => x.textContent.startsWith('华东区受益')); q.scrollIntoView({ block: 'center' }); const r = q.getBoundingClientRect(); return { x1: r.left - 4, y1: r.top - 4, x2: r.right + 4, y2: r.bottom + 4 }`)
  await p.waitJS(`return document.activeElement?.matches('.q-pin-in')`, 5000)
  await p.typeText('华北增速放缓这里展开一点，补上推迟订单的金额和预计确认时间')
  await p.key('Enter')
  await p.waitJS(`return __q.pins().length === 1`, 10000)
  await p.clickJS(`return __q.centre(document.querySelector('.qpd-pill'))`)
  await p.waitJS(`return !document.querySelector('.qpd-list').hidden`, 5000)
  await sleep(5000) // the "added" notice fades
  await shot('overview.webp')

  // Each editor full screen, without the chat.
  await open('预算.xlsx')
  await p.waitJS(`return !!__q.cur()?.querySelector('.qx-canvas')`, 20000)
  await full(true)
  await p.goCell('E11')
  await shot('xlsx.webp')
  await full(false)

  await open('季度报告.docx')
  await p.waitJS(`return !!__q.cur()?.querySelector('.qd-page')`, 20000)
  await full(true)
  await p.js(`const r = __q.cur().querySelector('.q-zoom-r'); r.value = '100'; r.dispatchEvent(new Event('input')); __q.cur().querySelector('.qd-page').scrollIntoView({ block: 'start' }); return true`)
  await p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qd-nav-i')].find((x) => x.textContent === '一、分区经营情况'))`)
  await shot('docx.webp')
  await full(false)

  await open('经营回顾.pptx')
  await p.waitJS(`return __q.cur()?.querySelectorAll('.qp-thumb').length === 5`, 20000)
  await full(true)
  await p.clickJS(`return __q.centre(__q.cur().querySelectorAll('.qp-thumb')[1])`)
  await p.clickJS(`return __q.centre(__q.cur().querySelector('.qp-stage .slide [data-k="2"]'))`)
  await shot('pptx.webp')
}
