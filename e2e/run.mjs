// End-to-end runs of quire, the only test mechanism of this repository.
// They drive the real pieces and write report/quire-e2e.md, with the
// generated documents, previews and screenshots beside it:
//
//   - vanilla: real dsh with the plugin package; a fake model calls every
//     Agent tool, and readers other than quire check the files.
//   - web: dsh web with the plugin, driven by a headless Chrome.
//
// node e2e/run.mjs [vanilla] [web]   (no argument runs both)
// Needs node and dsh (on PATH, or QUIRE_E2E_DSH: an executable or bin.js,
// run with QUIRE_E2E_NODE); the web run also Chrome or Edge
// (QUIRE_E2E_CHROME).
import { Report } from './harness.mjs'
import { vanilla } from './vanilla.mjs'
import { web } from './web.mjs'

const suites = { vanilla, web }
const pick = process.argv.slice(2)
for (const name of pick) {
  if (!suites[name]) {
    console.error(`unknown suite ${name}; have: ${Object.keys(suites).join(', ')}`)
    process.exit(2)
  }
}
const rep = new Report()
for (const [name, fn] of Object.entries(suites)) {
  if (pick.length && !pick.includes(name)) continue
  try {
    await fn(rep)
  } catch (err) {
    rep.section(name).step('运行到底', err)
  }
}
rep.write()
const total = rep.sections.reduce((n, s) => n + s.lines.length, 0)
console.log(`\n${total} steps, ${rep.fails} failed; report: e2e/report/quire-e2e.md`)
process.exit(rep.fails ? 1 : 0)
