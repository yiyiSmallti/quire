// The preview page: a self-contained, script-free HTML document.
import path from 'node:path'
import { renderDocx } from './docx.js'
import { renderXlsx } from './xlsx.js'
import { renderPptx } from './pptx.js'

const escH = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export function kindOf (file) {
  const ext = path.extname(file).toLowerCase()
  if (['.docx', '.docm', '.dotx', '.dotm'].includes(ext)) return 'docx'
  if (['.xlsx', '.xlsm', '.xltx', '.xltm'].includes(ext)) return 'xlsx'
  if (['.pptx', '.pptm', '.potx', '.potm', '.ppsx'].includes(ext)) return 'pptx'
  return undefined
}

export function renderPage (file, buf, when = new Date()) {
  const kind = kindOf(file)
  const r = kind === 'docx' ? renderDocx(buf) : kind === 'xlsx' ? renderXlsx(buf) : renderPptx(buf)
  const name = path.basename(file)
  const time = when.toLocaleString('zh-CN', { hour12: false })
  const accent = { docx: '#2B579A', xlsx: '#107C41', pptx: '#C43E1C' }[kind]
  const html = `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escH(name)}</title>
<style>
html,body{margin:0;background:#f0f0f0}
.bar{position:sticky;top:0;z-index:5;display:flex;align-items:center;gap:10px;padding:8px 16px;background:#fff;border-bottom:1px solid #ddd;font:13px/1.4 -apple-system,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;color:#333}
.bar .tag{color:#fff;background:${accent};border-radius:3px;padding:1px 7px;font-size:12px}
.bar .name{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar .time{margin-left:auto;color:#888;white-space:nowrap}
${r.css}
</style></head>
<body><div class="bar"><span class="tag">${r.kind}</span><span class="name" title="${escH(file)}">${escH(name)}</span><span class="time">更新于 ${escH(time)}</span></div>
${r.body}
</body></html>`
  return { title: name, html }
}
