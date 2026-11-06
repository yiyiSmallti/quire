// A headless Chrome on a throwaway profile, driven over the DevTools
// protocol on a pipe (--remote-debugging-pipe: fd 3 in, fd 4 out, one JSON
// message per NUL), and a page driver that sends real input events.
import { spawn } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { REPORT, clip, sleep, which } from './harness.mjs'

// chromePath finds Chrome or Edge, or QUIRE_E2E_CHROME.
export function chromePath () {
  if (process.env.QUIRE_E2E_CHROME) return process.env.QUIRE_E2E_CHROME
  let cands = []
  if (process.platform === 'darwin') {
    cands = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge']
  } else if (process.platform === 'win32') {
    cands = ['C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe']
  } else {
    cands = ['google-chrome', 'chromium', 'chromium-browser'].map(which).filter(Boolean)
  }
  return cands.find((c) => fs.existsSync(c)) ?? ''
}

export class Chrome {
  next = 0
  pending = new Map()
  loading = false
  session = ''

  static async start (bin, work, cleanup) {
    const c = new Chrome()
    c.proc = spawn(bin, ['--headless=new', '--remote-debugging-pipe', '--no-first-run', '--no-default-browser-check',
      '--disable-extensions', '--disable-background-networking', '--disable-sync', '--window-size=1200,900',
      '--user-data-dir=' + path.join(work, 'chrome-profile'), 'about:blank'],
    { stdio: ['ignore', 'ignore', 'ignore', 'pipe', 'pipe'] })
    const exited = new Promise((resolve) => c.proc.on('close', resolve))
    cleanup.add(async () => { c.proc.kill('SIGKILL'); await exited })
    const err = await new Promise((resolve) => {
      c.proc.on('spawn', () => resolve(null))
      c.proc.on('error', resolve)
    })
    if (err) throw err
    c.w = c.proc.stdio[3]
    c.w.on('error', () => {})
    c.#read(c.proc.stdio[4])
    const { targetId } = await c.send('', 'Target.createTarget', { url: 'about:blank' }, 20000)
    const { sessionId } = await c.send('', 'Target.attachToTarget', { targetId, flatten: true }, 20000)
    c.session = sessionId
    await c.send(c.session, 'Page.enable', {}, 20000)
    return c
  }

  #read (r) {
    let buf = Buffer.alloc(0)
    r.on('data', (d) => {
      buf = Buffer.concat([buf, d])
      let i
      while ((i = buf.indexOf(0)) >= 0) {
        const raw = buf.subarray(0, i).toString('utf8')
        buf = buf.subarray(i + 1)
        let m
        try { m = JSON.parse(raw) } catch { continue }
        if (m.method === 'Page.frameStartedLoading') this.loading = true
        if (m.method === 'Page.loadEventFired' || m.method === 'Page.frameStoppedLoading') this.loading = false
        const p = m.id && this.pending.get(m.id)
        if (p) {
          this.pending.delete(m.id)
          if (m.error) p.reject(new Error(m.error.message))
          else p.resolve(m.result ?? {})
        }
      }
    })
    const gone = () => {
      for (const p of this.pending.values()) p.reject(new Error('chrome exited'))
      this.pending.clear()
    }
    r.on('close', gone)
    r.on('error', gone)
  }

  send (session, method, params = {}, timeout = 60000) {
    const id = ++this.next
    const msg = { id, method, params }
    if (session) msg.sessionId = session
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`${method}: context deadline exceeded`))
      }, timeout)
      this.pending.set(id, {
        resolve: (v) => { clearTimeout(timer); resolve(v) },
        reject: (e) => { clearTimeout(timer); reject(e) }
      })
      this.w.write(JSON.stringify(msg) + '\0')
    })
  }

  // eval runs code as the body of an async function and returns its value.
  async eval (code) {
    const out = await this.send(this.session, 'Runtime.evaluate', {
      expression: '(async () => {\n' + code + '\n})()', awaitPromise: true, returnByValue: true, userGesture: true
    })
    const e = out.exceptionDetails
    if (e) throw new Error('javascript error: ' + (e.exception?.description || e.text))
    return out.result?.value
  }

  // waitLoaded waits for the page to start and finish loading after a
  // navigation.
  async waitLoaded () {
    const deadline = Date.now() + 15000
    for (let i = 0; Date.now() < deadline; i++) {
      let rs = ''
      try { rs = await this.eval('return document.readyState') } catch {}
      if (!this.loading && rs === 'complete' && i >= 2) return
      await sleep(100)
    }
  }

  navigate (url) {
    this.loading = true
    return this.send(this.session, 'Page.navigate', { url })
  }

  async capture () {
    const { data } = await this.send(this.session, 'Page.captureScreenshot', { format: 'png' })
    return Buffer.from(data, 'base64')
  }
}

// Page helpers, run before each step: find a visible element by text or
// selector and report the centre of its box for a real mouse click.
const pageJS = `
window.__q = {
  // reveal opens the ribbon tab a toolbar control sits on.
  reveal (n) { const p = n?.closest?.('.q-panel'); if (p && p.hidden) [...document.querySelectorAll('.quire .q-tab')].find((t) => t.textContent === p.dataset.tab)?.click(); return n },
  vis (e) { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && r.bottom > 0 && r.right > 0 },
  // byText finds host UI by its text; the editor's own tabs (it has a
  // 文件 tab too) never count.
  byText (text, right) {
    const own = (e) => [...e.childNodes].filter((n) => n.nodeType === 3).map((n) => n.data).join('').trim()
    // Text in a display: contents box is shown by the nearest box around it.
    const shown = (e) => { while (e && !this.vis(e)) e = e.parentElement; return e }
    const all = [...document.querySelectorAll('body *')].filter((e) => own(e) === text && !e.closest('.quire')).map(shown).filter(Boolean)
    const pick = right ? all.filter((e) => e.getBoundingClientRect().left > innerWidth / 2) : all
    return pick[pick.length - 1]
  },
  // draft is the chat composer's text.
  draft () { return document.querySelector('[data-lexical-editor="true"]')?.innerText ?? '' },
  // pins is the annotation cards above the composer, as their text.
  pins () { return [...document.querySelectorAll('.qpd-card')].map((c) => c.querySelector('.qpd-where').textContent + '：' + (c.querySelector('.qpd-note')?.textContent ?? '')) },
  // cur is the editor on screen; other tabs keep theirs in the page.
  cur () { return [...document.querySelectorAll('.quire')].find((q) => this.vis(q)) },
  centre (e) { if (!e) return null; e.scrollIntoView({ block: 'nearest' }); const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 } }
}
`

const keyCodes = {
  Enter: ['Enter', 13, '\r'],
  Tab: ['Tab', 9, ''],
  Escape: ['Escape', 27, ''],
  a: ['KeyA', 65, ''],
  b: ['KeyB', 66, ''],
  d: ['KeyD', 68, ''],
  m: ['KeyM', 77, ''],
  ArrowRight: ['ArrowRight', 39, ''],
  Delete: ['Delete', 46, '']
}

// pt reads a point a script returned as {x, y} (or {X, Y}).
export const pt = (v) => ({ x: v?.x ?? v?.X, y: v?.y ?? v?.Y })

export const jsStr = (s) => JSON.stringify(s)

// Page drives the dsh web page with real input events.
export class Page {
  constructor (c) { this.c = c }

  js (code) { return this.c.eval(pageJS + '\n' + code) }

  // waitJS polls code until it returns something other than false, null
  // or nothing.
  async waitJS (code, limit) {
    const deadline = Date.now() + limit
    let last = null
    while (Date.now() < deadline) {
      try {
        const v = await this.js(code)
        if (v !== false && v !== null && v !== undefined) return v
        last = null
      } catch (err) { last = err }
      await sleep(200)
    }
    throw last ?? new Error(`等待超时：${clip(code, 80)}`)
  }

  async #mouse (type, x, y, extra) {
    await this.c.send(this.c.session, 'Input.dispatchMouseEvent', { type, x, y, button: 'left', clickCount: 1, ...extra })
  }

  async click (x, y, n = 1) {
    for (let i = 1; i <= n; i++) {
      for (const type of ['mousePressed', 'mouseReleased']) await this.#mouse(type, x, y, { clickCount: i })
    }
    await sleep(300)
  }

  async rclick (x, y) {
    for (const type of ['mousePressed', 'mouseReleased']) await this.#mouse(type, x, y, { button: 'right' })
    await sleep(300)
  }

  // shiftClick clicks with Shift held, to add to a selection.
  async shiftClick (x, y) {
    for (const type of ['mousePressed', 'mouseReleased']) await this.#mouse(type, x, y, { modifiers: 8 })
    await sleep(300)
  }

  // clickJS clicks the centre the code returns, once or n times.
  async clickJS (code, n = 1) {
    const at = pt(await this.waitJS(code, 15000))
    await this.click(at.x, at.y, n)
  }

  clickText (text, right) {
    return this.clickJS(`return __q.centre(__q.byText(${jsStr(text)}, ${right}))`)
  }

  async drag (x1, y1, x2, y2) {
    const steps = [['mousePressed', x1, y1], ['mouseMoved', (x1 + x2) / 2, (y1 + y2) / 2], ['mouseMoved', x2, y2], ['mouseReleased', x2, y2]]
    for (const [type, x, y] of steps) {
      await this.#mouse(type, x, y, type === 'mouseMoved' ? { buttons: 1 } : {})
      await sleep(60)
    }
    await sleep(300)
  }

  // dragJS drags over the box a script returns as {x1, y1, x2, y2}.
  async dragJS (code) {
    const b = await this.waitJS(code, 15000)
    await this.drag(b.x1 ?? b.X1, b.y1 ?? b.Y1, b.x2 ?? b.X2, b.y2 ?? b.Y2)
  }

  async typeText (text) {
    await this.c.send(this.c.session, 'Input.insertText', { text })
    await sleep(200)
  }

  // key presses a key, with "mod" meaning ⌘ on macOS and Ctrl elsewhere.
  async key (name, ...mods) {
    const [code, vk, text] = keyCodes[name] ?? ['', 0, '']
    let m = 0
    for (const mod of mods) m |= mod === 'shift' ? 8 : process.platform === 'darwin' ? 4 : 2
    const down = { type: 'keyDown', key: name, code, windowsVirtualKeyCode: vk, modifiers: m }
    if (text && m === 0) down.text = text
    await this.c.send(this.c.session, 'Input.dispatchKeyEvent', down)
    await this.c.send(this.c.session, 'Input.dispatchKeyEvent', { type: 'keyUp', key: name, code, windowsVirtualKeyCode: vk, modifiers: m })
    await sleep(250)
  }

  // goCell selects a cell or range through the name box, then gives the
  // grid the keyboard.
  async goCell (ref) {
    await this.clickJS(`return __q.centre(__q.cur().querySelector('.qx-name'))`)
    await this.js(`const n = __q.cur().querySelector('.qx-name'); n.select(); return true`)
    await this.typeText(ref)
    await this.key('Enter')
  }

  async shot (name) {
    try {
      const png = await this.c.capture()
      fs.mkdirSync(REPORT, { recursive: true })
      fs.writeFileSync(path.join(REPORT, name), png)
    } catch {}
  }
}
