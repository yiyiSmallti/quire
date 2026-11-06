// @precinct/quire — Word, Excel and PowerPoint files for the dsh agent.
//
// Runs inside the dsh process. The tools read .docx/.xlsx/.pptx into text
// with refs, create them from Markdown or plain data, and edit them in
// place, touching only the parts an edit names; everything else in the
// file (styles, comments, charts, masters…) is written back as it was.
// Parsing and writing need nothing but Node (lib/zip.js, lib/xml.js).
//
// office_preview writes the preview page to a file; in dsh web the file
// preview sidebar shows the editor (lib/client.js) instead.
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import crypto from 'node:crypto'
import { makeIO } from './files.js'
import { createPinStore, injectPins } from './pins.js'
import { kindOf, renderPage } from './render.js'
import { readDocx, createDocx, editDocx } from './docx.js'
import { readXlsx, createXlsx, editXlsx } from './xlsx.js'
import { readPptx, createPptx, editPptx } from './pptx.js'
import { registerUiRoutes } from './ui-host.js'

export const name = 'quire'
export const inject = ['tools', 'fs', 'workingDirectory']

// ------------------------------------------------------------ helpers

const FORMATS = {
  docx: { read: readDocx, label: 'Word document' },
  xlsx: { read: readXlsx, label: 'Excel workbook' },
  pptx: { read: readPptx, label: 'PowerPoint presentation' }
}

function need (file, kind) {
  const k = kindOf(file)
  if (!k) throw new Error(`"${file}" is not a Word, Excel or PowerPoint file (.docx, .xlsx or .pptx)`)
  if (kind && k !== kind) throw new Error(`"${file}" is a ${FORMATS[k].label}; use the ${k}_* tools for it`)
  return k
}

// Office keeps an owner file "~$name" next to a document it has open.
function openElsewhere (abs) {
  const dir = path.dirname(abs)
  const base = path.basename(abs)
  for (const n of [`~$${base}`, `~$${base.slice(2)}`]) {
    if (n.length > 2 && fs.existsSync(path.join(dir, n))) return true
  }
  return false
}

const text = (_args, value) => [{ type: 'text', text: typeof value === 'string' ? value : value.text }]
const STR = { type: 'string' }
const FILE = (what) => ({ type: 'string', description: `Path of the ${what}, relative to the current directory or absolute.` })
const OVERWRITE = { type: 'boolean', description: 'Replace the file if it already exists (default false: refuse).' }
const CELL_VALUE = {
  description: 'A string, number, boolean or null; a string starting with = is a formula. Or an object { value?, formula?, format?, bold?, italic?, color?, fill?, align?, wrap? }: format is an Excel number format such as "#,##0.00", "0%" or "yyyy-mm-dd"; color and fill are 6-digit hex; align is left, center or right.'
}
const ROWS = { type: 'array', items: { type: 'array', items: CELL_VALUE }, description: 'Rows of cells, top to bottom.' }
const SLIDE = {
  type: 'object',
  description: 'One slide. layout: title, section, content, two_column or blank (inferred when omitted). bullets: strings or { text, level } (level 0–4; **bold** and *italic* work). left/right: bullets of the two columns. table: { rows: [[…]], header? }. image: { path } (PNG/JPEG). notes: speaker notes.',
  properties: {
    layout: { type: 'string', enum: ['title', 'section', 'content', 'two_column', 'blank'] },
    title: STR,
    subtitle: STR,
    bullets: { type: 'array', items: {} },
    left: { type: 'array', items: {} },
    right: { type: 'array', items: {} },
    table: { type: 'object', properties: { rows: { type: 'array', items: { type: 'array', items: {} } }, header: { type: 'boolean' } }, required: ['rows'] },
    image: { type: 'object', properties: { path: STR }, required: ['path'] },
    notes: STR
  }
}

// ------------------------------------------------------------ apply

export function apply (ctx) {
  async function writeNew (exec, args, kind, build) {
    const io = makeIO(ctx, exec)
    need(args.file_path, kind)
    const r = await io.resolve(args.file_path)
    const existed = await io.exists(r)
    if (existed && !args.overwrite) throw new Error(`"${r.display}" already exists; read it with office_read and change it with ${kind}_edit, or pass overwrite: true to replace it`)
    if (existed && openElsewhere(r.abs)) throw new Error(`"${r.display}" is open in Office (its ~$ owner file exists); ask the user to close it first`)
    const { buffer, summary } = await build(io)
    await io.write(r, buffer)
    return { text: `${existed ? 'Replaced' : 'Created'} ${r.abs} (${summary}, ${buffer.length} bytes).`, path: r.abs }
  }

  async function edit (exec, args, kind, run) {
    const io = makeIO(ctx, exec)
    need(args.file_path, kind)
    if (!Array.isArray(args.edits) || !args.edits.length) throw new Error('edits must be a non-empty array')
    const r = await io.resolve(args.file_path)
    if (openElsewhere(r.abs)) throw new Error(`"${r.display}" is open in Office (its ~$ owner file exists); ask the user to close it first, or the user's copy will overwrite these edits`)
    const before = await io.read(r)
    const { buffer, summary } = await run(before, io)
    await io.write(r, buffer)
    return { text: `Edited ${r.abs}: ${summary}.`, path: r.abs }
  }

  const TOOLS = [
    {
      name: 'office_read',
      description: 'Read a Word (.docx), Excel (.xlsx) or PowerPoint (.pptx) file as text. Word: blocks in order, [p3] paragraphs and [t1] tables as Markdown. ' +
        'Excel: the sheets with their used ranges, then one sheet as a grid with column letters and row numbers, formulas shown as =SUM(B2:B4) → 60. ' +
        'PowerPoint: each slide with its shapes as [s3.2] and speaker notes. The refs are what docx_edit and pptx_edit take. Long content comes in parts; the result says which from to pass next. ' +
        'Use this, not read, for these files: they are zip archives.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('document'),
          from: { type: 'integer', minimum: 1, description: 'Continue from here, as the previous part said: a block number (Word), row number (Excel) or slide number (PowerPoint).' },
          sheet: { type: 'string', description: 'Excel: the sheet to read (default the first).' },
          range: { type: 'string', description: 'Excel: the cells to read, like A1:F40 (default the used range, at most 200 rows × 40 columns).' },
          slides: { type: 'string', description: 'PowerPoint: the slides to read, like "3" or "2-5" (default all).' }
        },
        required: ['file_path'],
        additionalProperties: false
      },
      async execute (args, exec) {
        const io = makeIO(ctx, exec)
        const kind = need(args.file_path)
        const r = await io.resolve(args.file_path)
        const buf = await io.read(r)
        const body = FORMATS[kind].read(buf, { from: args.from, sheet: args.sheet, range: args.range, slides: args.slides })
        return { text: `${r.abs}\n${body}`, path: r.abs }
      }
    },
    {
      name: 'docx_create',
      description: 'Create a Word document (.docx) from Markdown: # headings, paragraphs, **bold**, *italic*, ~~strike~~, `code`, [links](url), - and 1. lists (two spaces per level), > quotes, GFM tables (first row is the header), ``` code blocks, ![caption](image.png) for local PNG/JPEG images, and a line \\pagebreak for a page break.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('new .docx file'),
          markdown: { type: 'string', description: 'The whole document as Markdown.' },
          overwrite: OVERWRITE
        },
        required: ['file_path', 'markdown'],
        additionalProperties: false
      },
      execute: (args, exec) => writeNew(exec, args, 'docx', (io) => createDocx(String(args.markdown ?? ''), io))
    },
    {
      name: 'docx_edit',
      description: 'Edit a Word document in place; untouched content and formatting stay as they are. Refs [p12] and [t1] come from office_read and all refer to the document as it was before this call. Ops:\n' +
        '{op:"replace", find, replace, all?} text in body, tables, headers and footers, even across formatting; keeps the formatting where the match starts; fails when find is not unique unless all: true.\n' +
        '{op:"insert", after|before: "p12", markdown} new blocks next to a block. {op:"append", markdown} at the end.\n' +
        '{op:"set", block: "p12", markdown} replace a block. {op:"delete", blocks: ["p3","t1"]}.\n' +
        '{op:"cell", table: "t1", row, col, text} set a table cell (1-based), keeping its formatting.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('.docx file'),
          edits: {
            type: 'array',
            minItems: 1,
            items: {
              type: 'object',
              properties: {
                op: { type: 'string', enum: ['replace', 'insert', 'append', 'set', 'delete', 'cell'] },
                find: STR,
                replace: STR,
                all: { type: 'boolean' },
                after: STR,
                before: STR,
                block: STR,
                blocks: { type: 'array', items: STR },
                markdown: STR,
                table: STR,
                row: { type: 'integer', minimum: 1 },
                col: { type: 'integer', minimum: 1 },
                text: STR
              },
              required: ['op']
            }
          }
        },
        required: ['file_path', 'edits'],
        additionalProperties: false
      },
      execute: (args, exec) => edit(exec, args, 'docx', (buf, io) => editDocx(buf, args.edits, io))
    },
    {
      name: 'xlsx_create',
      description: 'Create an Excel workbook (.xlsx). Each sheet is rows of cells; a string starting with = is a formula (calculated and cached, so office_read and previews show the result; Excel recalculates on open). header: true makes the first row bold, shaded and frozen.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('new .xlsx file'),
          sheets: {
            type: 'array',
            minItems: 1,
            items: {
              type: 'object',
              properties: {
                name: { type: 'string', description: 'Sheet name (at most 31 characters, none of : \\ / ? * [ ]).' },
                rows: ROWS,
                header: { type: 'boolean' },
                widths: { type: 'array', items: { type: 'number' }, description: 'Column widths in characters, from column A (estimated from the content when omitted).' },
                freeze: { type: 'string', description: 'Freeze panes above and left of this cell, like B2.' }
              },
              required: ['name', 'rows']
            }
          },
          overwrite: OVERWRITE
        },
        required: ['file_path', 'sheets'],
        additionalProperties: false
      },
      execute: (args, exec) => writeNew(exec, args, 'xlsx', async () => createXlsx(args.sheets))
    },
    {
      name: 'xlsx_edit',
      description: 'Edit an Excel workbook in place; other cells, styles, charts and sheets stay as they are. Every formula is recalculated after the edits. sheet defaults to the first sheet. Ops:\n' +
        '{op:"set", sheet?, cell: "B3", value?|formula?, format?, bold?, italic?, color?, fill?, align?, wrap?} one cell (a value string starting with = is a formula).\n' +
        '{op:"write", sheet?, at: "A1", rows: [[…]]} a block of cells, like xlsx_create rows.\n' +
        '{op:"style", sheet?, range: "A1:D1", bold?, italic?, color?, fill?, format?, align?, wrap?, width?} format cells; width sets the columns\' width in characters.\n' +
        '{op:"clear", sheet?, range}. {op:"freeze", sheet?, cell: "B2"}. {op:"add_sheet", name, rows?}. {op:"rename_sheet", sheet, name} (formulas follow). {op:"delete_sheet", sheet}.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('.xlsx file'),
          edits: {
            type: 'array',
            minItems: 1,
            items: {
              type: 'object',
              properties: {
                op: { type: 'string', enum: ['set', 'write', 'style', 'clear', 'freeze', 'add_sheet', 'rename_sheet', 'delete_sheet'] },
                sheet: STR,
                cell: STR,
                at: STR,
                range: STR,
                rows: ROWS,
                name: STR,
                value: {},
                formula: STR,
                format: STR,
                bold: { type: 'boolean' },
                italic: { type: 'boolean' },
                color: STR,
                fill: STR,
                align: { type: 'string', enum: ['left', 'center', 'right'] },
                wrap: { type: 'boolean' },
                width: { type: 'number' }
              },
              required: ['op']
            }
          }
        },
        required: ['file_path', 'edits'],
        additionalProperties: false
      },
      execute: (args, exec) => edit(exec, args, 'xlsx', async (buf) => editXlsx(buf, args.edits))
    },
    {
      name: 'pptx_create',
      description: 'Create a PowerPoint presentation (.pptx, 16:9) from a list of slides: title, section, bullet, two-column and blank slides, with tables, local PNG/JPEG images and speaker notes.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('new .pptx file'),
          slides: { type: 'array', minItems: 1, items: SLIDE },
          theme: {
            type: 'object',
            description: 'font: the typeface; accent, background: 6-digit hex colours.',
            properties: { font: STR, accent: STR, background: STR },
            additionalProperties: false
          },
          overwrite: OVERWRITE
        },
        required: ['file_path', 'slides'],
        additionalProperties: false
      },
      execute: (args, exec) => writeNew(exec, args, 'pptx', (io) => createPptx(args.slides, args.theme ?? {}, io))
    },
    {
      name: 'pptx_edit',
      description: 'Edit a PowerPoint presentation in place; masters, layouts, animations and untouched slides stay as they are. Shape refs [s3.2] come from office_read; slide numbers start at 1 and refer to the deck as it was before this call. Ops:\n' +
        '{op:"replace", find, replace, all?, slides?} text on slides. {op:"set_text", shape: "s3.2", text} one paragraph per line, two leading spaces per level, keeping the paragraph formatting.\n' +
        '{op:"add_slide", at?, slide: {…like pptx_create}} using the deck\'s own layouts. {op:"duplicate_slide", slide, at?}. {op:"delete_slide", slide}. {op:"move_slide", slide, to}. {op:"notes", slide, text}.',
      parameters: {
        type: 'object',
        properties: {
          file_path: FILE('.pptx file'),
          edits: {
            type: 'array',
            minItems: 1,
            items: {
              type: 'object',
              properties: {
                op: { type: 'string', enum: ['replace', 'set_text', 'add_slide', 'duplicate_slide', 'delete_slide', 'move_slide', 'notes'] },
                find: STR,
                replace: STR,
                all: { type: 'boolean' },
                slides: { description: 'Slide numbers to limit replace to, like "2-4" or [2, 3].' },
                shape: STR,
                text: STR,
                slide: { description: 'A slide number, or for add_slide the new slide.' },
                at: { type: 'integer', minimum: 1 },
                to: { type: 'integer', minimum: 1 }
              },
              required: ['op']
            }
          }
        },
        required: ['file_path', 'edits'],
        additionalProperties: false
      },
      execute: (args, exec) => edit(exec, args, 'pptx', (buf, io) => editPptx(buf, args.edits, io))
    },
    {
      name: 'office_preview',
      description: 'Write a rendered preview of a Word, Excel or PowerPoint file as an HTML file and return its path. The preview is approximate (no Office fonts or effects) and read-only.',
      parameters: {
        type: 'object',
        properties: { file_path: FILE('document') },
        required: ['file_path'],
        additionalProperties: false
      },
      async execute (args, exec) {
        const io = makeIO(ctx, exec)
        need(args.file_path)
        const r = await io.resolve(args.file_path)
        const buf = await io.read(r)
        const page = renderPage(r.abs, buf)
        const dir = path.join(os.tmpdir(), 'quire')
        fs.mkdirSync(dir, { recursive: true })
        const id = crypto.createHash('sha256').update(r.abs).digest('hex').slice(0, 12)
        const out = path.join(dir, `${path.basename(r.abs).replace(/[^\w.一-鿿-]+/g, '_')}-${id}.html`)
        fs.writeFileSync(out, page.html)
        return { text: `Wrote the preview of ${r.abs} to ${out}; open it in a browser.`, path: r.abs, preview: out }
      }
    }
  ]

  const ANY = { type: 'object', additionalProperties: true }
  for (const t of TOOLS) {
    const definition = {
      name: t.name,
      description: t.description,
      parameters: t.parameters,
      output: { schema: ANY, render: text },
      async execute (args, exec) {
        try {
          return await t.execute(args ?? {}, exec)
        } catch (err) {
          throw new Error(err?.message ?? String(err))
        }
      }
    }
    ctx.effect(() => ctx.tools.register(definition))
  }

  // dsh web: the editor in the file preview reads and saves through these
  // routes. Compositions without a browser Connection skip them.
  // Annotations from the editor ride along with the next typed message.
  const pins = createPinStore()
  ctx.inject(['connection'], (cctx) => registerUiRoutes(cctx, pins))
  ctx.inject(['agents'], (actx) => injectPins(actx, pins))

  ctx.inject(['systemPrompt'], (pctx) => {
    pctx.systemPrompt.section({
      name: 'tool:quire',
      order: (pctx.systemPrompt.getSectionOrder('TOOL_EDIT') ?? 1300) + 50,
      text: ({ scope }) => ctx.tools.get('office_read', scope) === undefined
        ? ''
        : 'Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) files are zip archives: read them with office_read, create them with docx_create, xlsx_create or pptx_create, and change them with docx_edit, xlsx_edit or pptx_edit, never with read, write or edit. Read a file before editing it so the refs are current.'
    })
  })
}
