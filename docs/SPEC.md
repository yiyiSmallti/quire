# Quire 执行 Spec

版本 0.1.1 · 2026-11-06 · 依赖：dsh 0.2.1-rc.x（原版即可）

## 1. 目标

让 dsh 里的 Agent 像处理文本文件一样处理 Word 文档（.docx）、Excel 表格（.xlsx）和 PowerPoint 演示文稿（.pptx）：读出结构化内容、从零新建、按需修改，改动只碰要改的地方，其余内容（样式、批注、图表、母版……）原样保留。

- 原版 dsh：装上插件包即有全部读写工具，不需要任何服务端。
- dsh web：插件带一个 web 客户端模块，在右侧边栏里接管 Office 文件的预览，换成人能直接操作的编辑器（见 §5）。

## 2. 组件与数据流

```
Agent 工具调用 ─▶ dsh 进程内 bundle（@precinct/quire，纯 JS，零依赖）─ ctx.fs 解析路径、读字节；原子写回

dsh web 侧边栏 ─ client.js（编辑器）─ POST /api/quire/open|save|undo|redo ─▶ ui-host.js（同一 bundle）─ 读、改、原子写回
```

- 解析与生成全在 bundle 里：自带 zip（node:zlib）与保真 XML 解析器（不认识的节点、属性、命名空间原样写回），不依赖 npm 包，离线环境装得上。
- 文件被 Office 打开着（同目录有 `~$` 锁文件）时拒绝修改或覆盖，提示用户先关掉，免得互相覆盖。
- 路径：经 `ctx.fs.resolve`（相对路径按会话当前目录）与 `ctx.fs.processPath`；读用 `ctx.fs.readBytes`（上限 100 MB）。`ctx.fs` 只有文本写入，写二进制由 bundle 在同目录写临时文件再 rename（原子替换）。挂了沙箱时按会话策略把关：`read-only` 拒绝写，`workspace-write` 只能写工作区根目录以内（两边都先解析符号链接再比较，macOS 的 `/tmp` 即 `/private/tmp`），`danger-full-access` 不限。
- `office_preview` 的预览页由 bundle 渲染成自包含 HTML（图片内联为 data URI），写到系统临时目录的 `quire/` 下并返回路径。
## 3. Agent 工具

`file_path` 相对会话当前目录。块与形状的编号（`[p12]`、`[t2]`、`[s3.2]`）来自最近一次 `office_read`，批量编辑时所有编号都按调用前的文档解析，前面的编辑不会让后面的编号错位。

| 工具 | 参数 | 行为 |
| --- | --- | --- |
| `office_read` | `file_path`、`from?`、`sheet?`、`range?`、`slides?` | 按扩展名读 .docx/.xlsx/.pptx（见 3.1）；每次最多约 16000 字，更长时结尾告诉 Agent 下一次的 `from` |
| `docx_create` | `file_path`、`markdown`、`overwrite?` | 用 Markdown 新建 Word 文档（见 3.2）；文件已存在且没给 `overwrite: true` 时拒绝 |
| `docx_edit` | `file_path`、`edits[]` | 替换文字、在块前后插入、改写或删除块、改表格单元格、追加（见 3.3） |
| `xlsx_create` | `file_path`、`sheets[]`、`overwrite?` | 新建工作簿：每个工作表给二维数组，`=` 开头为公式，可带格式（见 3.4） |
| `xlsx_edit` | `file_path`、`edits[]` | 写单元格或区域、设格式与列宽、清空、增删改名工作表（见 3.5） |
| `pptx_create` | `file_path`、`slides[]`、`theme?`、`overwrite?` | 新建演示文稿：标题页、章节页、要点页、两栏、空白，可带表格、图片、备注（见 3.6） |
| `pptx_edit` | `file_path`、`edits[]` | 替换文字、改形状文字、加页、复制页、删页、移页、改备注（见 3.7） |
| `office_preview` | `file_path` | 把预览写成 HTML 文件并返回路径 |

成功结果是一段简短文字：做了什么、文件路径、改动计数；失败抛出带恢复提示的错误（编号不存在时提示重新 `office_read`）。

### 3.1 读出的样子

- Word：按文档顺序，段落 `[p3] 文字`，标题 `[p1] # 标题`（按级别），列表项带 `- ` 或 `1. ` 和缩进，粗体、斜体、删除线、链接写成 Markdown；表格 `[t1]` 写成带行列号的 Markdown 表；图片 `[image: 名称 宽×高]`；分页符 `[page break]`。页眉页脚、脚注只计数提示，不展开。
- Excel：先列出所有工作表及已用区域，再读一个工作表（`sheet`，默认第一个）的一块区域（`range`，默认已用区域，最多 200 行 × 40 列）为表格，列标与行号齐全；公式单元格写成 `=SUM(B2:B4) → 60`；合并区域、冻结窗格单独一行说明。`from` 为接着读的起始行。
- PowerPoint：逐页 `## Slide 3 · 版式名`，每个形状 `[s3.2] title: 文字`（占位符类型或形状名），段落按级别缩进，表格写成 Markdown 表，图片 `[image]`，备注 `notes:`。`slides` 选页（`3`、`2-5`）。

### 3.2 Markdown 子集（docx_create、docx_edit 共用）

标题 `#`–`######`、段落、`**粗体**`、`*斜体*`、`~~删除线~~`、`` `代码` ``、`[链接](url)`、无序与有序列表（两个空格一级缩进）、`>` 引用、GFM 表格（首行为表头）、` ``` ` 代码块、`![说明](图片路径)`（PNG/JPEG，相对会话目录，按原尺寸、最宽 16 cm）、单独一行 `\pagebreak` 为分页符。中文默认用微软雅黑/苹方回退，西文 Calibri。

### 3.3 docx_edit

- `{ op: "replace", find, replace, all? }`：正文、表格、页眉页脚里的文字替换，跨格式片段也能匹配，替换后的文字沿用匹配起点的格式；不唯一且没给 `all` 时报错。
- `{ op: "insert", after | before: "p12" | "t1", markdown }`、`{ op: "append", markdown }`。
- `{ op: "set", block, markdown }`：用新内容替换整块；`{ op: "delete", blocks: [...] }`。
- `{ op: "cell", table: "t1", row, col, text }`：行列从 1 起，保留单元格格式。

### 3.4 xlsx_create

`sheets: [{ name, rows: [[cell…]…], header?, widths?, freeze? }]`。单元格可为字符串、数字、布尔、null，或 `{ value?, formula?, format?, bold?, italic?, color?, fill?, align?, wrap? }`；字符串以 `=` 开头即公式。`header: true` 首行加粗、加底色并冻结；`widths` 为列宽（字符数），不给时按内容估算。

### 3.5 xlsx_edit

`{ op: "set", sheet?, cell, …同单元格对象 }`、`{ op: "write", sheet?, at, rows }`、`{ op: "style", sheet?, range, bold?, italic?, color?, fill?, format?, align?, wrap?, width? }`、`{ op: "clear", sheet?, range }`、`{ op: "freeze", sheet?, cell }`、`{ op: "add_sheet", name, rows? }`、`{ op: "rename_sheet", sheet, name }`（公式与定义名里对旧名的引用跟着改；像 Excel 一样，只有不是普通词或像单元格地址的表名才加引号）、`{ op: "delete_sheet", sheet }`。

公式：自带求值器（约 110 个函数：数学与统计、条件聚合 SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS、查找 VLOOKUP/HLOOKUP/XLOOKUP/INDEX/MATCH/OFFSET/INDIRECT/ADDRESS、逻辑 IF/IFS/SWITCH/LET、文本与日期，完整列表见 `formula.js` 的 `FUNCTION_NAMES`；运算符对区域逐元素计算，单行或单列自动扩展，所以 `SUMPRODUCT((B2:B9="华东")*G2:G9)` 这类数组写法可用；比较时相差不到 15 位有效数字的数视为相等，`0.1+0.2=0.3` 为 TRUE；TEXT 支持 General、`正;负;零` 分段、`0`/`#`/`?` 占位、千分位、百分比和带引号字面量）。每次写入后重算整个工作簿并写缓存值，单次重算里每个单元格和每个区域只求一次，2 万行 14 万个公式约 2 秒；打开一份公式没有结果的文件（openpyxl 等脚本写出的，`<v>` 为空）时先整本算一遍再显示。读出与预览即可看到结果；不认识的函数不写缓存值，读出标注“打开时计算”。工作簿一律设 `fullCalcOnLoad`，Excel/WPS 打开时会全部重算。

### 3.6 pptx_create

`slides: [{ layout?, title?, subtitle?, bullets?, left?, right?, table?, image?, notes? }]`，`layout` 为 `title`、`section`、`content`、`two_column`、`blank`，省略时按内容推断。`bullets` 每项为字符串或 `{ text, level }`；`table: { rows, header? }`；`image: { path }` 放在内容区内按比例居中。`theme: { font?, accent?, background? }`（颜色为 6 位十六进制）。版式 16:9。

### 3.7 pptx_edit

- `{ op: "replace", find, replace, all?, slides? }`；`{ op: "set_text", shape: "s3.2", text }`：每行一个段落，行首每两个空格降一级，保留原段落格式。
- `{ op: "add_slide", at?, slide }`：内容同 3.6，用该文件自己的版式（按类型找标题、标题和内容、节标题、两栏、空白），没有对应版式时用最接近的。
- `{ op: "duplicate_slide", slide, at? }`、`{ op: "delete_slide", slide }`、`{ op: "move_slide", slide, to }`、`{ op: "notes", slide, text }`。页码从 1 起。

## 4. 预览渲染

- Word：段落、标题、列表、表格（含合并单元格）、粗斜体下划线删除线、颜色、对齐、图片，按 A4 纸面宽度排版。
- Excel：工作表标签切换，网格带列标行号，显示值按数字格式（常用的千分位、小数位、百分比、日期），粗斜体、字体颜色、底色、对齐、合并单元格、列宽。
- PowerPoint：每页按幻灯片尺寸等比缩放，形状位置大小来自形状自身或其版式、母版的占位符，文字字号、粗细、颜色、项目符号，纯色填充与背景，图片，表格；备注列在页下。
- 页面顶部是文件名、类型、更新时间。渲染求“看得出改了什么”，不追求与 Office 像素一致。

## 5. dsh web 编辑器

### 5.1 接入

- `package.json` 的 `dsh.client: { platform: "web", inject: ["@deepseek-ai/dsh-client-ui-sidebar-documentpreview"] }` 与 `exports["./client"]`：dsh web 启动时把 `lib/client.js` 作为客户端模块加载（服务端缓存，改了要重启 dsh web）。模块注入 `slots` 与 `documentPreviews`，为 .docx/.xlsx/.pptx 注册预览组件，取代官方只读预览；其它类型不受影响。
- Host 侧在 `connection` 服务上注册七个 POST 路由，只有 owner Peer 能到：

| 路由 | 请求 | 结果 |
| --- | --- | --- |
| `/api/quire/open` | `{ sessionId, path }` | `{ kind, version, writable, why?, model, undo, redo, undoLabels, redoLabels }` |
| `/api/quire/save` | `{ sessionId, path, version, ops[], steps? }` | `{ version, summary, model, undo, redo, undoLabels, redoLabels }`；`steps` 是各步的 op 数（和须等于 `ops.length`），每步各成一条撤销记录，文件只写一次；版本不符时 409 `conflict`，带最新 `{ version, model }` |
| `/api/quire/undo`、`/redo` | `{ sessionId, path, count? }` | `{ version, model, undo, redo, undoLabels, redoLabels }`；`count` 一次退/进多步；文件在这期间被别人改过时 409 `stale` 并清空历史 |
| `/api/quire/new` | `{ sessionId, path, kind, name }` | `{ path }`：在当前文件所在文件夹建一份空白 docx/xlsx/pptx；同名文件已存在时 409 `exists`，文件名不合法时 400 `bad-name` |
| `/api/quire/copy` | `{ sessionId, path, name }` | `{ path }`：把当前文件复制一份到同一文件夹，冲突规则同上 |
| `/api/quire/print` | `{ sessionId, path }` | `{ html }`：与预览同一套渲染的独立页面，编辑器放进隐藏 iframe 调起浏览器打印（可存 PDF） |

- 路径相对会话的工作区；工作区以外的文件、同目录有 `~$` 锁文件的文件只读打开（`writable: false` 加原因）。`version` 是文件内容哈希；同一文件的保存与撤销串行执行，写入同样是同目录临时文件再 rename。撤销/重做保存整份文件的前后字节，每个文件最多 100 步，只存在内存里；每步带一个按 op 类型取的中文名（“切换效果”“动画”“插入形状”等），撤销下拉列表按最近 20 步列出。新建和另存副本的文件用宿主的 `sidebarRight.openResource` 在新标签里打开。

### 5.2 编辑器

三个编辑器共用一套接口（`load`/`collect`/`saved`/`hasChanges`/`dispose`，查找用 `findAll`/`showMatch`/`replaceMatch`/`replaceAll`）：改动先积在客户端，约 0.7 秒无操作或切换时打包成 `ops` 保存；保存返回新 `model`，编辑器保留光标与选区原地更新。409 冲突时载入最新内容，提示最后一次改动没有保存。会话里 Agent 改了文件，编辑器收到文件变化后重载。

共用外壳：顶行是功能区选项卡与撤销/重做（各带 ▾ 下拉，按步列出、一次撤销多步）、查找、打印、全屏，下面一行是当前选项卡的按钮（窄时换行）。每个编辑器的第一个选项卡都是「文件」：新建空白文档/表格/演示文稿、另存副本（先保存再复制）、导出 PDF、打印、键盘快捷键；打开时默认停在第一个非「文件」选项卡，之后记住上次的选项卡（localStorage `quire.tab.<kind>`）。选项卡——文档：开始、插入、表格、布局、审阅、视图；表格：开始、插入、页面布局、公式、数据、审阅、视图；演示：开始、插入、设计、形状格式、切换、动画、放映、视图。底部状态栏：文档显示“第 N 页，共 M 页”与字数/选中字数，表格显示选区的求和、平均值、计数，演示显示“幻灯片 N / M”、版式与选中对象数；三者右侧都是缩放滑块（− 滑块 + 百分比，点百分比回到适应窗口或 100%），与视图选项卡的缩放下拉同步。快捷键：⌘S 立即保存，⌘/ 打开快捷键列表（当前编辑器的加通用的）。右键菜单、查找替换栏（⌘F / ⇧⌘H，大小写不敏感；表格查当前工作表的显示值与输入，文档查段落和可编辑单元格，演示查所有幻灯片的文本框和表格；当前结果用 CSS Highlight 标出）、打印/导出 PDF（⌘P）。插入的图片以 `data:image/(png|jpeg|gif);base64` 发给 Host，单张不超过 20 MB。

- 表格（canvas 网格）：工作表标签（右键插入、重命名、删除、左右移、隐藏/取消隐藏）、名称框跳转、公式栏、冻结窗格、合并单元格、列宽行高拖动、50–200% 缩放（⌘ 滚轮）；键盘与输入法输入、方向键/Tab/回车移动、选区、复制粘贴（制表符分列的文本，或网页/Excel 的 HTML 表格）。公式输入时按前缀提示函数（↑↓ 选、Tab/回车补全），光标在括号里时显示该函数的参数。填充柄拖动、⌘D 向下填充；自动筛选的表头带下拉按钮，面板里排序、搜索、勾选值。ops：`input`/`inputs`（值或 `=` 公式，写后重算）、`xf`（粗斜下划线、颜色、底色、对齐、数字格式）、`width`、`height`、`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`（公式、定义名、合并、条件格式、数据验证、超链接、筛选范围、图表系列的引用和图表锚点跟着移）、`merge`/`unmerge`、`border`（`all|outside|thick|top|bottom|left|right|none`）、`autofill`（数字按等差、单个日期逐日、末尾带数字的文本递增、公式平移引用）、`sort`（按区域内的绝对列号，可带标题行；区域里有合并单元格时拒绝；公式随行平移）、`filter`（建立 `autoFilter`，按列保留的值列表隐藏行，`''` 表示空白）、`filter_off`、`hide`（行或列的隐藏/取消隐藏）、`move_sheet`、`sheet_state`、`add_chart`（`range` 按 Excel 的规则排成系列：首行是文字时作系列名，首列是文字时作分类，行比列多时按列成系列；散点图首列作 X 值；`type` 为 `col|bar|line|area|pie|doughnut|scatter`，`at` 是左上角单元格，默认放在区域右边两列）、`chart_move`（`chart` 是工作表上图表的序号，`from`/`to` 是 `{c, r, dx, dy}`，1 起的行列加 100% 缩放下的像素偏移，写成 twoCellAnchor）、`chart_edit`（`title`，空串去掉标题；单一类型的图表可改 `type` 与 `stacked: 'stacked'|'percent'|null`，系列数据保留）、`chart_delete`（连同图表部件和其关系），以及 3.5 的 `clear`、`freeze`、`add_sheet`、`rename_sheet`、`delete_sheet`。
- 文档（contenteditable 纸面，上下是可直接编辑的页眉页脚区）：段落与表格块带客户端编号，保存时与基线比对，只发改了的块与改了的段落属性。ops：`para`（整段文字与字符格式：粗斜下划线删除线、颜色、字号、字体、高亮、上下标、链接；段落样式、对齐、列表、`spacing {line, before, after}`（twips，line 240 为单倍）、`indent {left, firstLine}`（负的 firstLine 为悬挂），null 清除）、`insert`（在块前后新建段落，沿用相邻段落格式）、`delete`、`cell`、`row`（上/下插入行、删除行）、`col`（左/右插入列、删除列；有合并单元格的表拒绝）、`table`（插入 rows×cols 表格）、`image`（居中的图片段落）、`hf`（默认页眉/页脚写成纯文本行，或 `pageNumber: true` 写“第 PAGE 页，共 NUMPAGES 页”；含复杂内容的页眉页脚只读）、`page`（纸张 A4/A3/A5/B5/Letter/Legal、纵横、页边距 普通/窄/适中/宽）、`comment_new`（`{id, text, author}`，`id` 是客户端临时编号，同一批 `para` 里放好 `{t:'cs'|'ce'|'cr', id}` 三个标记 run，保存时换成正式编号）、`comment_reply`（`{parent, text, author}`，写进 commentsExtended 的 paraIdParent，标记紧跟父批注）、`comment_edit`、`comment_done`（commentsExtended 的 `w15:done`）、`comment_delete`（连同回复）、`toc`（`before`/`after` 块，`levels` 默认 3：TOC 域放进内容控件，给标题加 `_Toc` 书签，目录项是到书签的超链接；页码不写，`w:dirty` 让 Word 打开时刷新）、`toc_update`、`toc_delete`。批注标记在段落里是不可编辑的 run，批注的范围用 CSS Highlight 标出；正文里引用标记没了（删掉了被批注的文字）时，批注一起删除。目录段落只读，点目录项跳到对应标题。图片、分节、未识别内容原样保留；结构性改动（表格、图片、页面、批注、目录）先保存已有改动再发，保存后整页重绘。
- 表格的其它 ops：`cf_add`/`cf_delete`/`cf_clear`（条件格式：大于小于介于等于、文本包含、重复值、前 N 项，数据条、色阶，按单元格当前值画在网格上）、`dv_set`/`dv_clear`（数据验证：下拉列表、整数小数范围，输入不合规时提示，下拉单元格带箭头）、`name_set`/`name_delete`（名称管理器）、`dedupe`（删除重复项）、`split`（分列）、`outline`（行列分组，折叠与展开）、`note_set`（批注，悬停显示）、`link_set`、`pic_add`/`pic_move`/`pic_delete`（工作表图片）、`protect`/`unprotect`（锁定的单元格只读）、`page`（纸张方向大小、页边距、打印区域、打印标题行、缩放、网格线），另有插入函数对话框、自动求和与格式刷。
- 标注（三个编辑器共用）：点顶行橙色的「标注」或按 ⇧⌘M 进入框选（编辑区描橙色虚线、光标变十字，再点一次或 Esc 退出），在文件上拖一个框：表格选中框到的单元格；文档从框左上角所在行选到右下角所在行（角落落在文字外时取整段），跨段时位置写首尾两段 `[p12]–[p15]`；演示取框住一半以上的形状，没有就取框所在的那个形状，都没有则记成幻灯片上的一块 `第 3 页 框选区域`（截图失败时退回按页宽高的百分比 `区域 10%,20%–60%,50%`）。单击则标注点到的那一处。右键「标注问题…」直接标注当前选中的内容。松开后在框旁弹出框里写问题，回车保存为一条标注，不写进聊天框文字。弹出时同时给框选范围截图（拖框就截框住的范围，单击或右键截选中的单元格、段落或形状，外扩 8px）：表格直接裁 canvas，文档和演示把可见的 DOM 连同计算后的样式克隆进 SVG foreignObject 再画到 canvas（长边不超过 1600px，图片资源转成 data URL），截图缩略图显示在弹出框里；截图失败时照常保存，只是不带图。每条标注记 `{文件, 位置, 原文, 补充, 说明, 截图}`，截图以 PNG（过大时 JPEG）连同不超过 240px 的 JPEG 缩略图一起提交。位置沿用 office_read 给 Agent 的编号：表格 `汇总!B4`/`汇总!B4:C6`（补充写公式或行列数），文档 `[p12]`、表格单元格 `[t2] 第 3 行第 2 列`，演示 `第 3 页 [s3.2]`（多个形状并列 `[s3.2] [s3.5]`）或 `第 3 页`；原文最多 200 字。标注按会话存在 dsh 进程内存里（`lib/pins.js`，每会话最多 50 条，截图只收 png/jpeg/webp、单张不超过 8MB，原图留在服务端，列表只回缩略图和 `hasImage`），经 `/api/quire/pins` 增删改查（op：list/add/edit/remove/clear，只有 owner 能访问）；dsh 重启会丢掉还没发送的标注。客户端在宿主的 `conversation.input.dock` 插槽（聊天框上方）登记一块标注区，默认只显示一行：「📌 N 条标注 ▴ ×」胶囊（× 清空全部）、提示文字，聊天框为空时多一个「直接发送」。点胶囊展开列表（再点、点别处或 Esc 收起），每条一张卡片，左侧是截图缩略图（没有截图时是图钉图标），右侧显示文件 · 位置「原文」（补充）和说明，✎ 就地改说明（回车保存、Esc 放弃），× 移除。点卡片在已打开的编辑器里跳到原处。发送：服务端在 `agent/pre-step` 上挂一个前置监听，当某一步带有用户亲手输入的消息（`source.kind === 'user'`）时，取走该会话的全部标注，追加一条 `source: {kind: 'quire', form: 'notice'}` 的用户消息（摘要「N 条文件标注 · 文件名」，正文先是一句说明，再逐条列出 `序号. 文件 · 位置「原文」（补充）：说明`，带截图的条目后面紧跟一个图片块：截图经宿主的 `attachments.saveImage` 存成附件，以 `{type: 'image', attachment}` 交给模型，一条消息最多 8 张，存不下或超出时在该条后注明「截图没能附上：原因」；模型不支持图片时由宿主处理），交出即从服务端清空（这一步若在调用模型前被取消，这些标注不会恢复）；工具往返、任务唤醒等没有用户输入的步骤不取。客户端在输入状态、队列或会话运行状态变化时重新拉取，卡片和编辑器里的标记随之消失。编辑器在原处标出本文件待发送的标注（表格橙色虚线框、文档文字高亮、演示形状虚线框和缩略图圆点），按钮上显示条数；标注按钮的 ▾ 列出本文件的标注（点击跳到原处）、移除本文件的标注、清空全部。只读文件也能标注。
- 文档的导航窗格：编辑区左侧按文档顺序列出各级标题（含按样式套用的标题，不含目录里的条目），点击跳到该标题并把光标放在开头，正文滚动时高亮当前所在的章节，有下级的标题可折叠（折叠后由它代表其中的小节）。视图选项卡和状态栏左端的「导航」按钮开关，选择记在本机；没设过时文档有标题就默认打开。
- 文档的其它 ops：按纸张尺寸与页边距真实分页（分页符另起一页，状态栏显示页数）、`merge`/`unmerge`/`split`（表格合并拆分单元格）、`tcell`（单元格底纹、边框、对齐）、`tstyle`（表格样式）、`twidth`（平均分布各列）、`liststyle`（项目符号、编号与多级列表样式）、`dropcap`（首字下沉）、`note_new`/`note_edit`（脚注、尾注）、`track`（修订开关，之后的改动写成 `w:ins`/`w:del`）、`rev`（接受/拒绝修订），以及段落边框底纹、分栏、格式刷。
- 表格的图表：从工作表的 drawing 关系读出各锚点（twoCell、oneCell、absolute）里的图表，柱形/条形、折线、面积、饼、圆环、散点按单元格的当前值（缓存兜底）画成 SVG 叠在网格上，颜色取系列的 srgbClr 或主题色。单击选中带 8 个手柄，拖动与缩放发 `chart_move`，右键改类型、堆积、标题、删除，Delete 删除；一次只发一个图表改动，等页面回来再发下一个（图表按序号指认）。保存前把所有图表系列的缓存重写成单元格的当前值，别的阅读器打开看到的也是新数据。组合了多种类型或 3D 的图表只能移动、改标题和删除。
- 演示：缩略图栏（右键新建、复制、删除、移动、隐藏、从此放映）、按幻灯片尺寸缩放的舞台（或固定比例缩放）、选择框带 8 个手柄拖动缩放（位置大小换算成 EMU）、双击进入文字编辑或表格单元格编辑（Tab 跳格）、备注框、全屏放映（跳过隐藏页，方向键/空格/单击翻页，Esc 退出）。形状剪贴板：⌘C 记下形状，⌘V 用 `copy_shape` 粘到当前页（同页错开一点），⌘X 粘贴时连原形状一起删；系统剪贴板里的图片或文字粘成图片或文本框。ops：`shape_text`（run 可带 `font`）、`xfrm`、`add_textbox`、`add_shape`（rect、roundRect、ellipse、triangle、diamond、pentagon、star5、rightArrow、leftArrow）、`add_image`、`add_table`、`table_cell`、`shape_style`（`fill`/`line` 为颜色或 null 表示无）、`order`（front/back/forward/backward）、`copy_shape`（新 id，跨页复制关系，去掉占位符身份）、`delete_shape`、`new_slide`（用文件自己的版式）、`hide_slide`、`notes`，以及 3.7 的 `duplicate_slide`、`delete_slide`、`move_slide`。新形状的填充与轮廓取主题色（`schemeClr accent1`，轮廓再加深 50%），换主题颜色时跟着变。
- 演示的形状编辑：选择框上方的圆形控点旋转（Shift 以 15° 为步长，写进 `a:xfrm rot`），旋转菜单可左右转 90°、转 180°、重置或输入角度；拖动时对齐其它形状和幻灯片的边与中线，显示参考线并吸附；从空白处拖出框选、Shift 单击加减选，多选一起移动；`group`/`ungroup` 写成或拆开 `p:grpSp`（⌘G / ⇧⌘G），单击组合选中整组，再单击选中组里的形状；`shape_style` 另有 `grad`（线性渐变）与 `shadow`（外阴影）；`text_effect`（阴影、轮廓、发光、渐变等艺术字效果，写进 run 属性）；`body`（文本框四边边距、不自动调整/溢出时缩小文字/根据文字调整形状、横排竖排的文字方向）；`crop`（图片 `a:srcRect`，裁剪时黑色控点调四边，Enter 完成，可按比例）。段落：项目符号与编号（`a:buChar` / `a:buAutoNum`）、列表级别、行距（`a:lnSpc`）。
- 演示的设计与动画：`slide_layout`（更换当前页版式，按占位符类型把内容搬到新版式）、`theme`（主题颜色与字体方案，改 `theme1.xml`；标题用主要字体 `+mj`，正文用次要字体）、`background`（纯色、渐变、图片，写当前页或母版的 `p:bg`）。幻灯片母版视图：缩略图栏列出母版和各版式，在上面增删改形状直接写母版/版式部件，用到它们的幻灯片跟着变。`transition`（淡入淡出、推入、擦除、分割、覆盖、揭开、缩放、溶解，带方向、速度、自动换片时间，可全部应用）与 `animations`（进入动画：出现、淡入、飞入、擦除、缩放，单击时/与上一动画同时/上一动画之后，时长，调整顺序；写成 `p:timing` 主序列），动画窗格列出本页动画，编辑区在形状旁显示序号。放映时按顺序播放切换与动画；演示者视图显示当前页、下一张、备注和计时，可另开观众窗口。

## 6. 测试

端到端为唯一机制（`e2e/`，报告 `e2e/report/quire-e2e.md`）：

- `TestVanillaDsh`：原版 dsh（sdk 模板 profile + 本包），假模型在一个回合里依次新建 docx/xlsx/pptx、读回、编辑、再读回、导出预览 HTML；生成的文件再由不经本插件的读者核对：macOS `textutil`（docx 文本）、Quick Look `qlmanage`（三种格式各出一张缩略图，证明 Office 兼容的渲染器能打开）、`unzip -t`（zip 完整），并保留到 `e2e/report/`。
- `TestWebEditor`：真实 dsh web（以 web 模板新建隔离 profile 并装本包，独立 `DSH_HOME`），先用 sdk profile 跑一个回合建出会话，再用无头 Chrome（CDP pipe，1440×900）打开会话侧边栏里的三份文件：表格用名称框和公式栏输入 `=D4*0.06`、⌘B 加粗后用工具栏撤销、再在编辑器外用同一套编辑代码改 B2（模拟 Agent），确认编辑器重载并显示新值，接着输入 `=SUM` 用函数提示补全、升序排序、拖填充柄、右键加框线、开筛选并去掉一项，从 A1:C3 插入柱形图后拖动、右键改成折线图、Delete 删除；文档打字、回车分段、选中加粗、改表格单元格、改字体与行距、查找全部替换、粘贴图片、写页眉、加页码、改横向、选中文字新建批注后回复并解决、在开头插入目录、改标题后更新目录；演示双击改标题、拖动副标题、改备注后切页、新建幻灯片后撤销、插入五角星并填色、⌘D 复制、Shift 多选两个星形后组合再取消组合、加淡入动画和推入切换、从撤销下拉一次撤销两步、在母版视图插入文本框并回到普通视图核对、状态栏缩放、放映翻页后退出；最后从「文件」新建空白 Word 文档，核对文件落盘并在新标签里打开。标注：点「标注」在 120 页演示文稿的最后一页拖框框住文本框、在长文档里 ⇧⌘M 框住标题和下一段（核对首尾段落编号后 Esc 取消）再框住一句话标注、在大表格上框选 B3:D5（核对位置后取消）再右键单元格标注，核对弹出框里的截图、聊天框上方默认收起的「N 条标注」胶囊里累积的三条（位置编号、原文、公式、缩略图）和编辑器里的标记，再从菜单移除表格的标注，点胶囊展开，用 ✎ 改写一张卡片的说明、用 × 移除另一张，在聊天框输入一句话发送，核对卡片清空、假模型在同一轮收到这句话、附带的标注和 1 张截图图片块（被移除的那条不在其中）；长文档的导航窗格列出 161 个标题，点第 30 章跳到对应页并高亮、折叠后收起 3 个小节。大文件：脚本生成的 2 万行表格（公式不带结果）打开即算出全部结果，「函数核对」页 71 个难点函数与 Excel 的答案一致，改一个单元格后整本重算写回、汇总核对为 0；40 章长文档分页、120 页演示文稿打开并翻到最后一页，各记用时。每一步都解开磁盘上的文件核对 XML，最后存截图与文件副本，并交给 python-docx / openpyxl（装了的话）复核。
## 7. 待做

- xlsx 插入/删除行列、图表给 Agent 工具也开放（网页编辑器已支持）。
- 网页编辑器：文档的富文本粘贴（表格已支持粘贴 HTML 表格）、目录页码（现由 Word 打开时刷新）、演示表格的增删行列；图表的坐标轴与数据标签设置、按筛选隐藏的行不画；动画只做进入效果（强调、退出、路径动画原样保留但锁定不改）；观众窗口在不允许弹窗的环境里打不开。
- docx 批注与修订给 Agent 工具开放（网页编辑器已支持）。
- pptx 的 SmartArt、组合形状内部文字的编辑。
