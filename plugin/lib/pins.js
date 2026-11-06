// Annotations the person adds in the editor wait here, per Session, until
// the next message they send. They never enter the draft: the composer
// folds them into a pill (client dock), and when a typed message enters
// a step the pre-step listener appends them to that step as one notice
// message, each pin's picture as an image block, then forgets them. Kept in memory: a dsh restart drops pins
// that were never sent.
import crypto from 'node:crypto'

const MAX_PINS = 50
const MAX_IMAGE = 8 * 1024 * 1024
// Pictures beyond this many in one message go as text only.
const MAX_PICTURES = 8
const MEDIA = new Set(['image/png', 'image/jpeg', 'image/webp'])
const cap = (v, n) => typeof v === 'string' ? v.slice(0, n) : ''
// A pin's picture bytes stay on the server; the composer gets the preview.
const pub = ({ image, mediaType, ...p }) => ({ ...p, hasImage: !!image })

export function createPinStore () {
  const sessions = new Map()
  const slot = (sid) => {
    let s = sessions.get(sid)
    if (!s) { s = { rev: 0, items: [] }; sessions.set(sid, s) }
    return s
  }
  const out = (sid) => {
    const s = sessions.get(sid)
    return { rev: s?.rev ?? 0, items: (s?.items ?? []).map(pub) }
  }
  const bump = (sid, s) => {
    s.rev++
    return out(sid)
  }
  return {
    list: out,
    add (sid, p) {
      const s = slot(sid)
      if (s.items.length >= MAX_PINS) return out(sid)
      const pin = {
        id: crypto.randomUUID(),
        path: cap(p?.path, 500),
        where: cap(p?.where, 300),
        quote: cap(p?.quote, 200),
        extra: cap(p?.extra, 300),
        note: cap(p?.note, 4000),
        thumb: typeof p?.thumb === 'string' && p.thumb.startsWith('data:image/') && p.thumb.length < 200000 ? p.thumb : ''
      }
      if (!pin.path || !pin.where) return out(sid)
      if (MEDIA.has(p?.image?.mediaType) && typeof p.image.data === 'string') {
        const data = Buffer.from(p.image.data, 'base64')
        if (data.length && data.length <= MAX_IMAGE) { pin.image = data; pin.mediaType = p.image.mediaType }
      }
      s.items = [...s.items, pin]
      return bump(sid, s)
    },
    edit (sid, id, note) {
      const s = sessions.get(sid)
      if (!s) return out(sid)
      s.items = s.items.map((p) => p.id === id ? { ...p, note: cap(note, 4000) } : p)
      return bump(sid, s)
    },
    // remove drops one pin by id, or every pin of one file by path.
    remove (sid, { id, path }) {
      const s = sessions.get(sid)
      if (!s) return out(sid)
      s.items = s.items.filter((p) => id ? p.id !== id : p.path !== path)
      return bump(sid, s)
    },
    clear (sid) {
      const s = sessions.get(sid)
      if (!s) return out(sid)
      s.items = []
      return bump(sid, s)
    },
    // take hands the pins over to a step and forgets them.
    take (sid) {
      const s = sessions.get(sid)
      if (!s?.items.length) return []
      const items = s.items
      s.items = []
      bump(sid, s)
      return items
    }
  }
}

// A part of a slide is shown by its picture, not by percentages.
const whereOf = (p) => p.image ? p.where.replace(/ 区域 [\d%,.–-]+$/, ' 框选区域') : p.where

function pinLine (p, i, pictured) {
  let s = `${i + 1}. ${p.path} · ${whereOf(p)}`
  if (p.quote) s += `「${p.quote}」`
  if (p.extra) s += `（${p.extra}）`
  s += '：' + (p.note || '（没有写说明）')
  if (pictured === true) s += '\n（下面是这处的截图）'
  else if (typeof pictured === 'string') s += `\n（截图没能附上：${pictured}）`
  return s
}

const INTRO = '用户在文件预览里给这些位置加了标注，和用户这条消息一起处理。方括号里是 office_read 的 ref；带截图的标注，截图就是用户框选的范围：'

export function pinText (items) {
  return INTRO + '\n' + items.map((p, i) => pinLine(p, i)).join('\n')
}

const freeze = (o) => {
  for (const v of Object.values(o)) if (v && typeof v === 'object') freeze(v)
  return Object.freeze(o)
}

// pinContent is the message body: each pin as a line, its picture after
// it as an image the attachment store keeps.
async function pinContent (items, attachments, logger) {
  const pictured = items.filter((p) => p.image).slice(0, MAX_PICTURES)
  const refs = new Map()
  const why = new Map()
  for (const p of items) if (p.image && !pictured.includes(p)) why.set(p, `一条消息最多带 ${MAX_PICTURES} 张截图`)
  if (pictured.length && !attachments) for (const p of pictured) why.set(p, '这个 dsh 没有附件服务')
  else {
    for (const [i, p] of pictured.entries()) {
      try {
        refs.set(p, await attachments.saveImage({ data: new Uint8Array(p.image), mediaType: p.mediaType, name: `标注${i + 1}.${p.mediaType.split('/')[1]}` }))
      } catch (err) {
        logger?.warn?.(`quire pins: picture not saved: ${err?.message ?? err}`)
        why.set(p, err?.message ?? String(err))
      }
    }
  }
  const content = [{ type: 'text', text: INTRO }]
  for (const [i, p] of items.entries()) {
    const ref = refs.get(p)
    content.push({ type: 'text', text: pinLine(p, i, ref ? true : why.get(p)) })
    if (ref) content.push({ type: 'image', attachment: ref })
  }
  return content
}

// injectPins appends the Session's pins to a step that a typed message
// opens. Steps without the person's own words (tool rounds, job wakeups)
// leave the pins waiting.
export function injectPins (ctx, store) {
  ctx.on('agent/pre-step', async ({ agent, signal }, next) => {
    const decision = await next()
    if (decision.kind === 'reject' || signal.aborted) return decision
    if (!decision.messages.some((m) => m?.source?.kind === 'user')) return decision
    const sid = agent?.session?.id
    if (!sid) return decision
    const items = store.take(sid)
    if (!items.length) return decision
    const files = [...new Set(items.map((p) => p.path))]
    const message = freeze({
      id: crypto.randomUUID(),
      role: 'user',
      content: await pinContent(items, ctx.get('attachments'), ctx.logger),
      source: { kind: 'quire', form: 'notice', summary: `${items.length} 条文件标注 · ${files.join('、')}`.slice(0, 120) }
    })
    return { ...decision, messages: [...decision.messages, message] }
  }, { prepend: true })
}
