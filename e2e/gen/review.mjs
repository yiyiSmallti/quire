// A Word document with tracked changes by another author, written the way
// the editor writes them with Track Changes on: node review.mjs LIB FILE
import fs from 'node:fs'
import path from 'node:path'
const [lib, file] = process.argv.slice(2)
const { createDocx, editDocxUi } = await import(path.join(lib, 'docx.js'))
let buf = (await createDocx('# 合同草稿\n\n付款期限为三十天。\n\n违约金按日万分之五计算。\n')).buffer
buf = (await editDocxUi(buf, [
  { op: 'track', on: true },
  { op: 'para', block: 'p2', runs: [{ t: 'text', text: '付款期限为六十天。' }], author: '王芳' },
  { op: 'para', block: 'p3', runs: [{ t: 'text', text: '违约金按日万分之三计算。' }], author: '李雷' }
])).buffer
fs.writeFileSync(file, buf)
