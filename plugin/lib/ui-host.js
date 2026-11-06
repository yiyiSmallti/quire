// Host half of the editor in dsh web: two authenticated routes on the
// Connection's /api channel. open returns a document as the editor sees
// it; save applies the person's edits, refusing when the file changed
// since it was opened (the Agent may have edited it meanwhile), and
// writes atomically. Only owner Peers reach these routes; files open
// read-only outside the Session's workspace.
import fs from 'node:fs'
import path from 'node:path'
import crypto from 'node:crypto'
import { kindOf, renderPage } from './render.js'
import { docxUiModel, editDocxUi, createDocx } from './docx.js'
import { xlsxUiModel, editXlsxUi, createXlsx } from './xlsx.js'
import { pptxUiModel, editPptxUi, createPptx } from './pptx.js'

const MAX = 100 * 1024 * 1024
const MAX_IMAGE = 20 * 1024 * 1024

const json = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } })
const versionOf = (buf) => crypto.createHash('sha256').update(buf).digest('hex').slice(0, 24)

function real (p) {
  let head = path.resolve(p)
  const rest = []
  for (;;) {
    try { return path.join(fs.realpathSync.native(head), ...rest) } catch {
      const up = path.dirname(head)
      if (up === head) return path.resolve(p)
      rest.unshift(path.basename(head))
      head = up
    }
  }
}

function modelOf (kind, buf) {
  if (kind === 'docx') return docxUiModel(buf)
  if (kind === 'xlsx') return xlsxUiModel(buf)
  return pptxUiModel(buf)
}

class Failure extends Error {
  constructor (status, code, message) { super(message); this.status = status; this.code = code }
}

const HISTORY = 100

// What an undo step did, in the words of the editor's menus.
const OP_LABELS = {
  input: '输入', inputs: '输入', clear: '清除内容', cell: '编辑单元格', style: '设置格式', xf: '设置格式', border: '设置边框', merge: '合并单元格', unmerge: '取消合并',
  row: '插入/删除行', col: '插入/删除列', insert: '插入', delete: '删除', width: '调整列宽', height: '调整行高', hide: '隐藏', freeze: '冻结窗格', sort: '排序', filter: '筛选', filter_off: '取消筛选',
  autofill: '填充', dedupe: '删除重复项', cf_add: '条件格式', cf_delete: '删除条件格式', cf_clear: '清除条件格式', dv_set: '数据验证', dv_clear: '清除数据验证', name_set: '定义名称', name_delete: '删除名称',
  add_sheet: '新建工作表', delete_sheet: '删除工作表', rename_sheet: '重命名工作表', move_sheet: '移动工作表', sheet_state: '工作表设置', outline: '分组', note_new: '新建批注', note_edit: '编辑批注', note_set: '批注', split: '分列',
  add_chart: '插入图表', chart_edit: '编辑图表', chart_move: '移动图表', chart_delete: '删除图表', pic_add: '插入图片', page: '页面设置', hf: '页眉页脚',
  para: '编辑文字', table: '表格', tcell: '编辑表格', tstyle: '表格样式', twidth: '表格宽度', image: '图片', link_set: '超链接', liststyle: '列表样式', dropcap: '首字下沉', toc: '目录', toc_update: '更新目录', toc_delete: '删除目录',
  comment_new: '新建批注', comment_reply: '回复批注', comment_edit: '编辑批注', comment_done: '解决批注', comment_delete: '删除批注', track: '修订设置', rev: '接受/拒绝修订',
  shape_text: '编辑文字', xfrm: '移动/调整形状', shape_style: '形状格式', add_textbox: '插入文本框', add_shape: '插入形状', add_image: '插入图片', add_table: '插入表格', table_cell: '编辑表格', copy_shape: '粘贴形状', delete_shape: '删除形状', order: '调整层次',
  group: '组合', ungroup: '取消组合', crop: '裁剪图片', body: '文本框选项', text_effect: '艺术字效果', slide_layout: '更换版式', theme: '更改主题', background: '设置背景', transition: '切换效果', animations: '动画',
  new_slide: '新建幻灯片', duplicate_slide: '复制幻灯片', delete_slide: '删除幻灯片', move_slide: '移动幻灯片', hide_slide: '隐藏幻灯片', notes: '编辑备注'
}
const labelOf = (ops) => {
  const names = [...new Set(ops.map((o) => OP_LABELS[o?.op] ?? '编辑'))]
  return names.length > 2 ? names.slice(0, 2).join('、') + '等' : names.join('、')
}
// labels lists the steps most recent first, for the undo and redo menus.
const labels = (list) => list.slice(-20).reverse().map((x) => x.label)
const historyOut = (h) => ({ undo: h?.undo.length ?? 0, redo: h?.redo.length ?? 0, undoLabels: h ? labels(h.undo) : [], redoLabels: h ? labels(h.redo) : [] })

// New files the editor makes: one blank page, sheet or slide.
async function blank (kind) {
  if (kind === 'docx') return (await createDocx('', {})).buffer
  if (kind === 'xlsx') return createXlsx([{ name: 'Sheet1', rows: [] }]).buffer
  return (await createPptx([{ layout: 'title', title: '', subtitle: '' }], {}, {})).buffer
}

export function registerUiRoutes (ctx, pins) {
  const queues = new Map()
  // Undo and redo restore whole files the editor wrote, and only while
  // the file is still exactly what the editor left.
  const history = new Map()
  const historyOf = (abs) => { let h = history.get(abs); if (!h) { h = { undo: [], redo: [] }; history.set(abs, h) } return h }

  // serial runs fn alone for one file.
  async function serial (abs, fn) {
    const prev = queues.get(abs) ?? Promise.resolve()
    let release
    const mine = new Promise((resolve) => { release = resolve })
    const tail = prev.then(() => mine)
    queues.set(abs, tail)
    await prev
    try { return await fn() } finally {
      release()
      if (queues.get(abs) === tail) queues.delete(abs)
    }
  }

  async function sessionHeader (sessionId) {
    let header = ctx.get('sessions')?.get?.(sessionId)?.header
    if (!header) header = (await ctx.get('sessionPersistence')?.stat?.(sessionId).catch(() => undefined))?.header
    if (!header) throw new Failure(404, 'no-session', '找不到这个会话')
    return header
  }

  async function locate (body) {
    const sessionId = body?.sessionId
    const rel = body?.path
    if (typeof sessionId !== 'string' || !sessionId || typeof rel !== 'string' || !rel) throw new Failure(400, 'bad-request', 'sessionId and path are required')
    const header = await sessionHeader(sessionId)
    const root = header.cwd ?? ctx.get('sandboxPolicy')?.workspaceRoot
    if (!root) throw new Failure(400, 'no-workspace', '这个会话没有工作区')
    const abs = path.isAbsolute(rel) ? path.resolve(rel) : path.resolve(root, rel)
    const kind = kindOf(abs)
    if (!kind) throw new Failure(415, 'not-office', 'Quire 只编辑 .docx、.xlsx 和 .pptx')
    const inside = !path.relative(real(root), real(abs)).startsWith('..') && !path.isAbsolute(path.relative(real(root), real(abs)))
    const lock = path.join(path.dirname(abs), '~$' + path.basename(abs).slice(2))
    let why
    if (!inside) why = '文件在工作区之外，只能查看'
    else if (fs.existsSync(lock) || fs.existsSync(path.join(path.dirname(abs), '~$' + path.basename(abs)))) why = '文件正在 Office/WPS 中打开，关掉后才能在这里编辑'
    return { abs, root, kind, writable: !why, why }
  }

  function readFile (abs) {
    let st
    try { st = fs.statSync(abs) } catch { throw new Failure(404, 'not-found', '文件不存在') }
    if (!st.isFile()) throw new Failure(400, 'not-file', '不是普通文件')
    if (st.size > MAX) throw new Failure(413, 'too-large', '文件超过 100 MB')
    return fs.readFileSync(abs)
  }

  function writeFile (abs, buf) {
    const dir = path.dirname(abs)
    const tmp = path.join(dir, `.~quire-${crypto.randomBytes(6).toString('hex')}.tmp`)
    try {
      fs.writeFileSync(tmp, buf)
      try { fs.chmodSync(tmp, fs.statSync(abs).mode & 0o7777) } catch {}
      fs.renameSync(tmp, abs)
    } catch (err) {
      try { fs.rmSync(tmp, { force: true }) } catch {}
      throw new Failure(500, 'write-failed', `保存失败：${err.message}`)
    }
  }

  // Pictures the person inserts come from their computer (sent inline as a
  // data: URI) or from the workspace, like the Agent's.
  const io = (root) => ({
    async loadImage (src) {
      const m = /^data:image\/(png|jpe?g|gif);base64,([A-Za-z0-9+/=\s]+)$/i.exec(src)
      if (m) {
        const data = Buffer.from(m[2], 'base64')
        if (data.length > MAX_IMAGE) throw new Error('图片超过 20 MB')
        return data
      }
      if (/^(https?|data):/i.test(src)) throw new Error('只能插入 PNG、JPEG、GIF 图片')
      const abs = path.resolve(root, src.replace(/^file:\/\//, ''))
      const st = fs.statSync(abs)
      if (st.size > MAX_IMAGE) throw new Error('图片超过 20 MB')
      return fs.readFileSync(abs)
    }
  })

  const handle = (fn) => async (request) => {
    try {
      const body = await request.json().catch(() => { throw new Failure(400, 'bad-json', '请求不是 JSON') })
      return json(200, await fn(body))
    } catch (err) {
      if (err instanceof Failure) return json(err.status, { error: { code: err.code, message: err.message }, ...(err.extra ?? {}) })
      ctx.logger?.warn?.(`quire ui: ${err.stack ?? err}`)
      return json(422, { error: { code: 'edit-failed', message: err.message } })
    }
  }

  ctx.effect(() => ctx.connection.fetch.register({
    path: '/api/quire/open',
    methods: ['POST'],
    requestBody: 'buffered',
    fetch: handle(async (body) => {
      const f = await locate(body)
      const buf = readFile(f.abs)
      const h = history.get(f.abs)
      return { kind: f.kind, version: versionOf(buf), writable: f.writable, why: f.why, model: modelOf(f.kind, buf), ...historyOut(h) }
    })
  }), 'quire: open route')

  ctx.effect(() => ctx.connection.fetch.register({
    path: '/api/quire/save',
    methods: ['POST'],
    requestBody: 'buffered',
    fetch: handle(async (body) => {
      const f = await locate(body)
      if (!f.writable) throw new Failure(403, 'read-only', f.why)
      if (!Array.isArray(body.ops) || !body.ops.length) throw new Failure(400, 'no-ops', '没有改动')
      return serial(f.abs, async () => {
        const buf = readFile(f.abs)
        const version = versionOf(buf)
        if (body.version !== version) {
          const err = new Failure(409, 'conflict', '文件已被其他人或 Agent 修改')
          err.extra = { version, model: modelOf(f.kind, buf) }
          throw err
        }
        // The editor groups ops by the action that made them; each action is
        // its own undo step. Without groups the whole batch is one step.
        const sizes = Array.isArray(body.steps) && body.steps.every((n) => Number.isInteger(n) && n > 0) && body.steps.reduce((a, n) => a + n, 0) === body.ops.length ? body.steps : [body.ops.length]
        const edit = async (b, ops) => f.kind === 'docx' ? await editDocxUi(b, ops, io(f.root))
          : f.kind === 'xlsx' ? editXlsxUi(b, ops)
            : await editPptxUi(b, ops, io(f.root))
        const steps = []
        let cur = buf
        let res
        let at = 0
        for (const n of sizes) {
          const ops = body.ops.slice(at, at + n)
          at += n
          res = await edit(cur, ops)
          steps.push({ before: cur, after: res.buffer, label: labelOf(ops) })
          cur = res.buffer
        }
        writeFile(f.abs, cur)
        const h = historyOf(f.abs)
        h.undo.push(...steps)
        if (h.undo.length > HISTORY) h.undo.splice(0, h.undo.length - HISTORY)
        h.redo = []
        return { version: versionOf(cur), summary: res.summary, model: modelOf(f.kind, cur), ...historyOut(h) }
      })
    })
  }), 'quire: save route')

  // print returns the file as a self-contained page, for printing or
  // saving as PDF in the browser.
  ctx.effect(() => ctx.connection.fetch.register({
    path: '/api/quire/print',
    methods: ['POST'],
    requestBody: 'buffered',
    fetch: handle(async (body) => {
      const f = await locate(body)
      return { html: renderPage(f.abs, readFile(f.abs)).html }
    })
  }), 'quire: print route')

  // name checks a file name the person typed for a new file or a copy.
  const nameFor = (raw, kind) => {
    const n = String(raw ?? '').trim()
    if (!n || /[\\/:*?"<>|\u0000-\u001f]/.test(n) || n.startsWith('.') || n.startsWith('~$')) throw new Failure(400, 'bad-name', '文件名不能为空，也不能含有 \\ / : * ? " < > |')
    return n.toLowerCase().endsWith('.' + kind) ? n : `${n}.${kind}`
  }
  const relOf = (root, abs) => path.relative(real(root), real(abs)).split(path.sep).join('/')

  // new makes a blank document beside the open one; copy saves the open
  // document under another name. Neither overwrites a file.
  ctx.effect(() => ctx.connection.fetch.register({
    path: '/api/quire/new',
    methods: ['POST'],
    requestBody: 'buffered',
    fetch: handle(async (body) => {
      const kind = body?.kind
      if (!['docx', 'xlsx', 'pptx'].includes(kind)) throw new Failure(400, 'bad-kind', '只能新建 Word、Excel 或 PowerPoint 文件')
      const f = await locate(body)
      if (!f.writable) throw new Failure(403, 'read-only', f.why)
      const abs = path.join(path.dirname(f.abs), nameFor(body.name, kind))
      if (fs.existsSync(abs)) throw new Failure(409, 'exists', '已有同名文件，请换个名字')
      writeFile(abs, await blank(kind))
      return { path: relOf(f.root, abs) }
    })
  }), 'quire: new route')

  ctx.effect(() => ctx.connection.fetch.register({
    path: '/api/quire/copy',
    methods: ['POST'],
    requestBody: 'buffered',
    fetch: handle(async (body) => {
      const f = await locate(body)
      const inside = !relOf(f.root, path.dirname(f.abs)).startsWith('..')
      if (!inside) throw new Failure(403, 'read-only', '文件在工作区之外，不能在旁边另存')
      const abs = path.join(path.dirname(f.abs), nameFor(body.name, f.kind))
      if (fs.existsSync(abs)) throw new Failure(409, 'exists', '已有同名文件，请换个名字')
      writeFile(abs, readFile(f.abs))
      return { path: relOf(f.root, abs) }
    })
  }), 'quire: copy route')

  // Annotations waiting for the next message (lib/pins.js).
  ctx.effect(() => ctx.connection.fetch.register({
    path: '/api/quire/pins',
    methods: ['POST'],
    requestBody: 'buffered',
    fetch: handle(async (body) => {
      const sid = body?.sessionId
      if (typeof sid !== 'string' || !sid) throw new Failure(400, 'bad-request', 'sessionId is required')
      await sessionHeader(sid)
      if (body.op === 'add') return pins.add(sid, body.pin)
      if (body.op === 'edit') return pins.edit(sid, String(body.id ?? ''), body.note)
      if (body.op === 'remove') return pins.remove(sid, { id: typeof body.id === 'string' ? body.id : '', path: typeof body.path === 'string' ? body.path : '' })
      if (body.op === 'clear') return pins.clear(sid)
      return pins.list(sid)
    })
  }), 'quire: pins route')

  for (const dir of ['undo', 'redo']) {
    ctx.effect(() => ctx.connection.fetch.register({
      path: `/api/quire/${dir}`,
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: handle(async (body) => {
        const f = await locate(body)
        if (!f.writable) throw new Failure(403, 'read-only', f.why)
        return serial(f.abs, async () => {
          const h = historyOf(f.abs)
          const from = dir === 'undo' ? h.undo : h.redo
          const to = dir === 'undo' ? h.redo : h.undo
          // count steps back (or forward) at once, as the undo menu asks.
          const count = Math.max(1, Math.min(from.length, Number.isInteger(body.count) ? body.count : 1))
          const step = from[from.length - 1]
          if (!step) throw new Failure(409, 'nothing', dir === 'undo' ? '没有可撤销的操作' : '没有可重做的操作')
          const buf = readFile(f.abs)
          const want = dir === 'undo' ? step.after : step.before
          if (versionOf(buf) !== versionOf(want)) {
            history.delete(f.abs)
            throw new Failure(409, 'stale', '文件在这期间被修改过，无法撤销')
          }
          const moved = from.splice(from.length - count, count).reverse()
          const last = moved[moved.length - 1]
          const next = dir === 'undo' ? last.before : last.after
          writeFile(f.abs, next)
          to.push(...moved)
          return { version: versionOf(next), model: modelOf(f.kind, next), ...historyOut(h) }
        })
      })
    }), `quire: ${dir} route`)
  }
}
