// An Open Packaging Conventions package: the parts of a .docx/.xlsx/.pptx,
// their relationships and content types. Parts are kept as bytes and
// parsed on first use; only parts that were parsed and marked changed are
// serialized again, so everything else is written back byte for byte.
import path from 'node:path/posix'
import { readZip, writeZip } from './zip.js'
import { parse, serialize, doc, el, children, child } from './xml.js'

export const NS = {
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
}

export const REL = (t) => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + t

const MEDIA_CT = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff' }

export class Package {
  constructor (entries) {
    this.order = []
    this.bytes = new Map()
    this.docs = new Map()
    this.dirty = new Set()
    for (const { name, data } of entries) {
      this.order.push(name)
      this.bytes.set(name, data)
    }
  }

  static fromBuffer (buf) { return new Package(readZip(buf)) }

  has (name) { return this.bytes.has(name) || this.docs.has(name) }
  names () { return [...this.order] }

  get (name) {
    if (this.docs.has(name) && this.dirty.has(name)) return Buffer.from(serialize(this.docs.get(name)), 'utf8')
    return this.bytes.get(name)
  }

  // xml returns the parsed part (cached), or undefined when absent.
  xml (name) {
    if (this.docs.has(name)) return this.docs.get(name)
    const b = this.bytes.get(name)
    if (!b) return undefined
    const d = parse(b)
    this.docs.set(name, d)
    return d
  }

  // touch marks a parsed part changed.
  touch (name) { this.dirty.add(name) }

  setXml (name, d) {
    if (!this.has(name)) this.order.push(name)
    this.docs.set(name, d)
    this.dirty.add(name)
    this.bytes.delete(name)
  }

  setBytes (name, data) {
    if (!this.has(name)) this.order.push(name)
    this.bytes.set(name, data)
    this.docs.delete(name)
    this.dirty.delete(name)
  }

  delete (name) {
    this.bytes.delete(name)
    this.docs.delete(name)
    this.dirty.delete(name)
    this.order = this.order.filter((n) => n !== name)
  }

  toBuffer () {
    // [Content_Types].xml first, as Office writes it.
    const names = [...this.order].sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0))
    return writeZip(names.map((name) => ({ name, data: this.get(name) })))
  }

  // ---------------------------------------------------------- relationships

  relsName (part) {
    const dir = path.dirname(part)
    return (dir === '.' ? '' : dir + '/') + '_rels/' + path.basename(part) + '.rels'
  }

  relsDoc (part, create = false) {
    const name = part === '' ? '_rels/.rels' : this.relsName(part)
    let d = this.xml(name)
    if (!d && create) {
      d = doc(el('Relationships', { xmlns: NS.rel }))
      this.setXml(name, d)
    }
    return d ? { name, d } : undefined
  }

  // rels lists part's relationships with targets resolved to part names.
  rels (part) {
    const r = this.relsDoc(part)
    if (!r) return []
    return children(r.d.root, 'Relationship').map((n) => ({
      id: n.attrs.Id,
      type: n.attrs.Type,
      target: n.attrs.Target,
      external: n.attrs.TargetMode === 'External',
      part: n.attrs.TargetMode === 'External' ? undefined : resolveTarget(part, n.attrs.Target),
      node: n
    }))
  }

  rel (part, id) { return this.rels(part).find((r) => r.id === id) }
  relsOfType (part, type) { return this.rels(part).filter((r) => r.type === type || r.type.endsWith('/' + type)) }

  // addRel adds a relationship from part to target (a part name, or a URL
  // with external) and returns its id.
  addRel (part, type, target, external = false) {
    const { name, d } = this.relsDoc(part, true)
    const ids = new Set(children(d.root, 'Relationship').map((n) => n.attrs.Id))
    let i = ids.size + 1
    while (ids.has('rId' + i)) i++
    const id = 'rId' + i
    const attrs = { Id: id, Type: type, Target: external ? target : relativeTarget(part, target) }
    if (external) attrs.TargetMode = 'External'
    d.root.kids.push(el('Relationship', attrs))
    this.touch(name)
    return id
  }

  removeRel (part, id) {
    const r = this.relsDoc(part)
    if (!r) return
    r.d.root.kids = r.d.root.kids.filter((n) => n.attrs?.Id !== id)
    this.touch(r.name)
  }

  // ---------------------------------------------------------- content types

  contentTypes () { return this.xml('[Content_Types].xml') }

  contentType (part) {
    const ct = this.contentTypes()
    if (!ct) return undefined
    const o = children(ct.root, 'Override').find((n) => n.attrs.PartName === '/' + part)
    if (o) return o.attrs.ContentType
    const ext = path.extname(part).slice(1).toLowerCase()
    return children(ct.root, 'Default').find((n) => n.attrs.Extension.toLowerCase() === ext)?.attrs.ContentType
  }

  setOverride (part, type) {
    const ct = this.contentTypes()
    const o = children(ct.root, 'Override').find((n) => n.attrs.PartName === '/' + part)
    if (o) o.attrs.ContentType = type
    else ct.root.kids.push(el('Override', { PartName: '/' + part, ContentType: type }))
    this.touch('[Content_Types].xml')
  }

  removeOverride (part) {
    const ct = this.contentTypes()
    ct.root.kids = ct.root.kids.filter((n) => n.attrs?.PartName !== '/' + part)
    this.touch('[Content_Types].xml')
  }

  ensureDefault (ext, type) {
    const ct = this.contentTypes()
    ext = ext.toLowerCase()
    if (children(ct.root, 'Default').some((n) => n.attrs.Extension.toLowerCase() === ext)) return
    const at = ct.root.kids.findIndex((n) => n.name === 'Override')
    const node = el('Default', { Extension: ext, ContentType: type })
    if (at < 0) ct.root.kids.push(node)
    else ct.root.kids.splice(at, 0, node)
    this.touch('[Content_Types].xml')
  }

  // addMedia stores image bytes under dir (e.g. word/media) and returns
  // the new part name.
  addMedia (dir, ext, data) {
    ext = ext === 'jpg' ? 'jpeg' : ext
    let i = 1
    while (this.has(`${dir}/image${i}.${ext}`)) i++
    const name = `${dir}/image${i}.${ext}`
    this.setBytes(name, data)
    this.ensureDefault(ext, MEDIA_CT[ext] ?? 'application/octet-stream')
    return name
  }

  // uniqueName finds pattern with the lowest free number: 'ppt/slides/slide#.xml'.
  uniqueName (pattern) {
    let i = 1
    while (this.has(pattern.replace('#', i))) i++
    return pattern.replace('#', i)
  }

  // mainPart is the target of the package's officeDocument relationship.
  mainPart () {
    const r = this.rels('').find((x) => x.type === NS.officeDocument || x.type.endsWith('/officeDocument'))
    return r?.part
  }
}

export function resolveTarget (fromPart, target) {
  if (target.startsWith('/')) return target.slice(1)
  const base = fromPart === '' ? '' : path.dirname(fromPart)
  return path.normalize(path.join(base === '.' ? '' : base, target)).replace(/^\.\//, '')
}

export function relativeTarget (fromPart, toPart) {
  const base = fromPart === '' ? '.' : path.dirname(fromPart)
  return path.relative(base, toPart) || path.basename(toPart)
}

export function mediaType (part) {
  return MEDIA_CT[path.extname(part).slice(1).toLowerCase()]
}

export function dataUri (pkg, part) {
  const b = pkg.get(part)
  const t = mediaType(part)
  if (!b || !t || !/png|jpeg|gif|svg|bmp/.test(t)) return undefined
  return `data:${t};base64,${b.toString('base64')}`
}

export { child }
