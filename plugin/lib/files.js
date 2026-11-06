// Paths and bytes for the tools. Paths resolve like dsh's own file tools:
// relative to the Session's current directory, through ctx.fs. ctx.fs
// writes text only, so documents are written here, atomically (a temp file
// in the same directory, then rename), under the Session's sandbox policy.
import fs from 'node:fs'
import path from 'node:path'
import crypto from 'node:crypto'

const MAX_READ = 100 * 1024 * 1024
const MAX_IMAGE = 20 * 1024 * 1024

// real resolves symlinks in p's longest existing prefix (macOS /tmp is
// /private/tmp), so paths compare the way the OS sees them.
function real (p) {
  let head = path.resolve(p)
  const rest = []
  for (;;) {
    try {
      return path.join(fs.realpathSync.native(head), ...rest)
    } catch {
      const up = path.dirname(head)
      if (up === head) return path.resolve(p)
      rest.unshift(path.basename(head))
      head = up
    }
  }
}

export function makeIO (ctx, exec) {
  const signal = exec?.signal
  const hasFs = !!ctx.fs

  async function cwdOpts () {
    let cwd
    try {
      if (exec?.agent && ctx.workingDirectory) cwd = await ctx.workingDirectory.ensure(exec.agent, signal)
    } catch {}
    return { ...(cwd ? { cwd } : {}), signal }
  }

  async function resolve (p) {
    if (typeof p !== 'string' || !p.trim()) throw new Error('file_path must be a non-empty string')
    if (!hasFs) {
      const abs = path.resolve(process.cwd(), p)
      return { abs, display: p, target: undefined }
    }
    const target = await ctx.fs.resolve(p, await cwdOpts())
    return { abs: ctx.fs.processPath(target), display: target.displayPath ?? p, target }
  }

  async function exists (r) {
    if (r.target) return (await ctx.fs.stat(r.target, signal)) !== undefined
    return fs.existsSync(r.abs)
  }

  async function read (r, max = MAX_READ) {
    if (r.target) {
      const info = await ctx.fs.stat(r.target, signal)
      if (!info) throw new Error(`cannot read "${r.display}": not found`)
      if (info.type !== 'file') throw new Error(`cannot read "${r.display}": not a regular file`)
      return Buffer.from(await ctx.fs.readBytes(r.target, signal, max))
    }
    if (!fs.existsSync(r.abs)) throw new Error(`cannot read "${r.display}": not found`)
    const st = fs.statSync(r.abs)
    if (st.size > max) throw new Error(`cannot read "${r.display}": larger than ${max >> 20} MB`)
    return fs.readFileSync(r.abs)
  }

  // allowWrite applies the Session's sandbox mode, when a confining
  // filesystem is mounted.
  function allowWrite (r) {
    if (!hasFs || ctx.fs.sandboxMode === undefined) return
    let policy
    try { policy = ctx.get?.('sandboxPolicy')?.resolve({ ...(exec?.agent ? { session: exec.agent.session } : {}) }) } catch {}
    const mode = policy?.mode ?? ctx.fs.sandboxMode
    if (mode === 'read-only') throw new Error(`[sandbox: file access denied under read-only mode] cannot write "${r.display}"`)
    if (mode === 'workspace-write' && policy?.workspaceRoot) {
      const rel = path.relative(real(policy.workspaceRoot), real(r.abs))
      if (rel.startsWith('..') || path.isAbsolute(rel)) throw new Error(`[sandbox: file access denied under workspace-write mode] "${r.display}" is outside the workspace ${policy.workspaceRoot}`)
    }
  }

  async function write (r, buf) {
    allowWrite(r)
    const dir = path.dirname(r.abs)
    fs.mkdirSync(dir, { recursive: true })
    const tmp = path.join(dir, `.~quire-${crypto.randomBytes(6).toString('hex')}.tmp`)
    try {
      fs.writeFileSync(tmp, buf)
      try {
        const mode = fs.statSync(r.abs).mode
        fs.chmodSync(tmp, mode & 0o7777)
      } catch {}
      fs.renameSync(tmp, r.abs)
    } catch (err) {
      try { fs.rmSync(tmp, { force: true }) } catch {}
      throw new Error(`cannot write "${r.display}": ${err.message}`)
    }
    if (r.target) {
      // Let dsh's read-before-write policy know the file's new version.
      try {
        const info = await ctx.fs.stat(r.target, signal)
        if (info) ctx.emit?.('fs/observed', r.target, { kind: 'present', version: info.version }, exec)
      } catch {}
    }
  }

  async function loadImage (src) {
    if (/^https?:\/\//i.test(src)) throw new Error(`image "${src}": download it into the workspace first, then give its path`)
    const r = await resolve(src.replace(/^file:\/\//, ''))
    return read(r, MAX_IMAGE)
  }

  return { resolve, exists, read, write, loadImage }
}
