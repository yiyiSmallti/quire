// Zip archives, the container of every Office Open XML file: read all
// entries into memory, write them back. Stored and deflated entries only,
// which is what Office, WPS and LibreOffice write; no zip64 (Office files
// stay far below 4 GB) and no encryption.
import zlib from 'node:zlib'

const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  return t
})()

export function crc32 (buf) {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

export class ZipError extends Error {}

// readZip returns the entries in archive order as [{ name, data }].
export function readZip (input) {
  const buf = Buffer.isBuffer(input) ? input : Buffer.from(input)
  // The end-of-central-directory record sits in the last 64 KB + 22 bytes.
  let eocd = -1
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) throw new ZipError('not a zip archive (no end of central directory): the file is not an Office Open XML document, or it is damaged')
  const count = buf.readUInt16LE(eocd + 10)
  let p = buf.readUInt32LE(eocd + 16)
  if (count === 0xffff || p === 0xffffffff) throw new ZipError('zip64 archives are not supported')
  const out = []
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new ZipError('damaged zip central directory')
    const flags = buf.readUInt16LE(p + 8)
    const method = buf.readUInt16LE(p + 10)
    const csize = buf.readUInt32LE(p + 20)
    const nameLen = buf.readUInt16LE(p + 28)
    const extraLen = buf.readUInt16LE(p + 30)
    const commentLen = buf.readUInt16LE(p + 32)
    const local = buf.readUInt32LE(p + 42)
    const name = buf.toString(flags & 0x800 ? 'utf8' : 'latin1', p + 46, p + 46 + nameLen)
    p += 46 + nameLen + extraLen + commentLen
    if (flags & 1) throw new ZipError(`"${name}" is encrypted: password-protected Office files cannot be opened`)
    if (buf.readUInt32LE(local) !== 0x04034b50) throw new ZipError(`damaged zip entry "${name}"`)
    const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28)
    const raw = buf.subarray(start, start + csize)
    let data
    if (method === 0) data = Buffer.from(raw)
    else if (method === 8) data = zlib.inflateRawSync(raw)
    else throw new ZipError(`"${name}" uses zip compression method ${method}, which is not supported`)
    if (name.endsWith('/')) continue
    out.push({ name, data })
  }
  return out
}

const STORE = /\.(png|jpe?g|gif|zip|xlsx|docx|pptx|mp3|mp4|m4a)$/i

// writeZip builds an archive from [{ name, data }], keeping the order.
export function writeZip (entries) {
  const chunks = []
  const central = []
  let offset = 0
  // A fixed timestamp keeps output reproducible: 2026-01-01 00:00.
  const time = 0
  const date = ((2026 - 1980) << 9) | (1 << 5) | 1
  for (const { name, data } of entries) {
    const nameBuf = Buffer.from(name, 'utf8')
    const crc = crc32(data)
    let method = 8
    let body = STORE.test(name) ? data : zlib.deflateRawSync(data, { level: 6 })
    if (body.length >= data.length) { method = 0; body = data }
    const head = Buffer.alloc(30)
    head.writeUInt32LE(0x04034b50, 0)
    head.writeUInt16LE(20, 4)
    head.writeUInt16LE(0x800, 6)
    head.writeUInt16LE(method, 8)
    head.writeUInt16LE(time, 10)
    head.writeUInt16LE(date, 12)
    head.writeUInt32LE(crc, 14)
    head.writeUInt32LE(body.length, 18)
    head.writeUInt32LE(data.length, 22)
    head.writeUInt16LE(nameBuf.length, 26)
    head.writeUInt16LE(0, 28)
    chunks.push(head, nameBuf, body)
    const c = Buffer.alloc(46)
    c.writeUInt32LE(0x02014b50, 0)
    c.writeUInt16LE(20, 4)
    c.writeUInt16LE(20, 6)
    c.writeUInt16LE(0x800, 8)
    c.writeUInt16LE(method, 10)
    c.writeUInt16LE(time, 12)
    c.writeUInt16LE(date, 14)
    c.writeUInt32LE(crc, 16)
    c.writeUInt32LE(body.length, 20)
    c.writeUInt32LE(data.length, 24)
    c.writeUInt16LE(nameBuf.length, 28)
    c.writeUInt32LE(offset, 42)
    central.push(c, nameBuf)
    offset += 30 + nameBuf.length + body.length
  }
  const cd = Buffer.concat(central)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(entries.length, 8)
  end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(cd.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...chunks, cd, end])
}
