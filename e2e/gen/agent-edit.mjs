// Changes a cell the way the Agent's tools would, while the file is open
// in the editor: node agent-edit.mjs LIB FILE
import fs from 'node:fs'
import path from 'node:path'
const [lib, file] = process.argv.slice(2)
const { editXlsxUi } = await import(path.join(lib, 'xlsx.js'))
fs.writeFileSync(file, editXlsxUi(fs.readFileSync(file), [{ op: 'input', sheet: '明细', cell: 'B2', text: '9000' }]).buffer)
