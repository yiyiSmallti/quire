# Quire

让 dsh 里的 Agent 像处理文本文件一样处理 Word 文档（.docx）、Excel 表格（.xlsx）和 PowerPoint 演示文稿（.pptx）：读出带编号的结构化内容，从零新建，按需修改。只改要改的地方，样式、批注、图表、母版这些没碰到的内容原样写回。

- **原版 dsh 就能用**：装上插件包即可，纯 JS、零依赖，不需要 Office、LibreOffice 或任何服务端。
- **dsh web 里能直接编辑**：右侧边栏打开工作区里的 Office 文件时，用 Quire 编辑器代替官方只读预览，人可以直接改、给 Agent 做标注。

Quire（一叠对折的纸，书的一帖）。

## 安装

dsh 还没有官方插件市场，下面几种方式任选其一，装进任何带文件系统插件的 profile（例如 web、sdk 模板）。

**npm**：dsh web 侧栏「插件」→「添加插件」输入 `@precinct/quire`，或

```bash
dsh plugin --profile web add @precinct/quire
```

**GitHub**（包在仓库的 `plugin/` 子目录）：

```bash
dsh plugin --profile web add "git+https://github.com/yiyiSmallti/quire.git#main&path:/plugin"
```

**插件市场索引**：在「插件」页「实验性插件」里启用「自定义插件市场」，到「发现」标签页添加索引源 `https://raw.githubusercontent.com/yiyiSmallti/quire/main/market/index.json`，即可搜索并一键安装。索引文件是 [market/index.json](market/index.json)（dsh 插件市场索引 format 1）。

**本地源码**：

```bash
(cd plugin && pnpm pack)
```

```bash
dsh plugin --profile <profile> add plugin/precinct-quire-0.1.1.tgz
```

## Agent 工具

| 工具 | 做什么 |
| --- | --- |
| `office_read` | 读 .docx/.xlsx/.pptx：Word 按块 `[p3]`、`[t1]` 写成 Markdown；Excel 列出工作表，再把一块区域读成带列标行号的表，公式写成 `=SUM(B2:B4) → 60`；PowerPoint 逐页列出形状 `[s3.2]` 和备注。长内容分段读 |
| `docx_create` / `docx_edit` | Markdown 新建 Word（标题、粗斜体、列表、表格、图片、引用、分页）；替换文字（跨格式片段也能匹配）、在块前后插入、改写或删除块、改单元格、追加 |
| `xlsx_create` / `xlsx_edit` | 二维数组新建工作簿，`=` 开头是公式；写单元格或区域、设数字格式和样式、列宽、冻结、清空、增删改名工作表（公式跟着改名）。自带公式求值器（约 110 个函数，支持数组运算），写完立即整本重算并缓存结果；打开不带公式结果的文件时先算一遍 |
| `pptx_create` / `pptx_edit` | 标题页、章节页、要点页、两栏页、空白页，可带表格、图片、备注；替换文字、改形状文字、加页（用文件自己的版式）、复制、删除、移动、改备注 |
| `office_preview` | 把预览写成 HTML 文件并返回路径 |

写入是原子的（同目录临时文件再 rename），遵守 dsh 的沙箱策略（`read-only` 不写，`workspace-write` 只写工作区内）；文件被 Office 打开着（有 `~$` 锁文件）时拒绝修改，提示先关掉。

## dsh web 编辑器

在 dsh web 右侧边栏打开工作区里的 .docx/.xlsx/.pptx 就会进入 Quire 编辑器。

**界面**：顶部是 Office 式的功能区选项卡，Word 蓝、Excel 绿、PowerPoint 红三种主色，跟随 dsh 的明暗主题，记住上次用的选项卡。最前面是「文件」：新建空白 Word/Excel/PowerPoint、另存副本、导出 PDF、打印、键盘快捷键（⌘/）。三种文件都有右键菜单、查找替换（⌘F）、⌘S 立即保存，底部状态栏带页数/字数/选区统计和缩放滑块。

**保存与撤销**：改动自动保存回原文件；撤销/重做按操作步骤记录（最多 100 步），下拉列表能一次撤销多步；Agent 同时改了文件时编辑器自动重载，不会互相覆盖。

**标注**：点顶行橙色的「标注」（或 ⇧⌘M）进入框选，在文件上拖框框住要说的单元格、文字或形状（单击标一处），写下问题，同时自动截下框选的范围。标注收在聊天框上方的「N 条标注」胶囊里，点开才列出每条的截图、文件、位置编号和原文，可改写说明、单条移除或一键清空，不占用聊天框文字。下一条消息发出时，这些标注连同截图（以图片形式）随消息一起交给 Agent；还没发送的标注在编辑器原处标出。

**表格**（开始、插入、页面布局、公式、数据、审阅、视图）：

- 公式栏与函数提示补全、插入函数、自动求和、名称框与名称管理器。
- 冻结窗格、合并、行列增删隐藏、框线、单元格样式、格式刷。
- 条件格式（突出显示、数据条、色阶）与数据验证（下拉列表、数值范围）。
- 排序（含多列）、自动筛选、删除重复项、分列、行列分组折叠。
- 批注、链接、图片、保护工作表；纸张方向大小、页边距、打印区域、打印标题行与缩放。
- 柱形、条形、折线、面积、饼、圆环、散点图，可拖动缩放、改类型、堆积、改标题。

**文档**（开始、插入、表格、布局、审阅、视图）：

- 左侧导航窗格按标题跳转、高亮当前章节、可折叠；按纸张真实分页显示。
- 字体字号、删除线上下标、段落边框底纹、首行缩进、列表样式与多级列表、格式刷。
- 图片、表格（合并拆分单元格、底纹边框、对齐、表格样式、平均分布列）、分页符、分栏、首字下沉、脚注尾注、链接。
- 页眉页脚、页码，纸张、方向、页边距。
- 批注与修订（接受、拒绝、上一处/下一处）；按标题生成目录并一键更新。

**演示**（开始、插入、设计、形状格式、切换、动画、放映、视图）：

- 拖动缩放旋转形状（对齐参考线吸附），框选与 Shift 多选、组合/取消组合。
- 填充、渐变、轮廓、阴影、艺术字，文本框边距与自动调整，图片裁剪；项目符号、编号与行距。
- 更换版式、主题颜色与字体、背景（纯色、渐变、图片）、幻灯片母版视图。
- 切换效果与进入动画（动画窗格、编辑区序号、预览）。
- 全屏放映与演示者视图（计时、下一张、备注）。

## 仓库结构

| 部分 | 位置 | 说明 |
| --- | --- | --- |
| dsh bundle | `plugin/lib/` | 工具注册（`index.js`）、路径与原子写（`files.js`）、预览页（`render.js`）；自带 zip（`zip.js`）与保真 XML（`xml.js`）、OPC 包模型（`pkg.js`）、Markdown 子集（`md.js`）、三种格式（`docx.js`、`xlsx.js`、`pptx.js`）和公式求值器（`formula.js`） |
| 插件包 `@precinct/quire` | `plugin/package.json` | `dsh.bundle.patch` 把 bundle 装进 profile，`dsh.client` 声明 web 客户端模块 |
| 网页编辑器 | `plugin/lib/client.js`、`plugin/lib/ui-host.js` | `dsh.client` 声明的 web 客户端模块，接管侧边栏的文档预览；Host 侧在 Connection 的 `/api/quire/open\|save\|undo\|redo\|print` 上读写文件 |
| 端到端测试 | `e2e/` | `run.mjs` 入口，`vanilla.mjs`、`web.mjs` 两组用例，`gen/` 生成测试文件，报告写在 `e2e/report/` |
| 市场索引 | `market/index.json` | dsh 插件市场索引 |

行为细节与待做事项见 [docs/SPEC.md](docs/SPEC.md)。

## 测试

端到端是唯一的测试机制，只用 Node 内置模块：

```bash
node e2e/run.mjs
```

只跑一组时加 `vanilla` 或 `web`。报告、生成的文件、预览页和截图写在 `e2e/report/`（`quire-e2e.md`）。

- **vanilla**：真实 dsh（sdk 模板 + 本包），假模型在一个回合里新建、读回、编辑三种文件，改一份别的程序写的 .docx，导出预览，并验证沙箱、错误提示；生成的文件再交给不经本插件的读者核对（textutil、Quick Look、unzip，装了的话还有 python-docx / openpyxl）。
- **web**：真实 dsh web（web 模板 + 本包），无头 Chrome 像人一样在侧边栏里编辑三种文件，每一步解开磁盘上的文件核对：
  - 表格：输入公式（含函数补全）、加粗后撤销、排序、拖填充柄、右键加框线、筛选，插入图表后拖动、改类型、删除；在编辑器外改文件（模拟 Agent），确认编辑器跟着重载。
  - 文档：打字分段、改单元格、改字体行距、查找替换、粘贴图片、写页眉加页码、改横向、批注回复并解决、插入并更新目录。
  - 演示：改标题、拖动形状、改备注、新建幻灯片、插入并填色形状、⌘D 复制、组合、动画与切换、母版、放映翻页。
  - 标注：在三种文件上框选标注，核对截图、默认收起的标注胶囊与编辑器里的标记，展开后改写、移除卡片再发送一条消息，核对模型在同一轮收到附带的标注和截图。
  - 大文件：2 万行公式表格（含 71 个函数核对）、长文档（导航窗格跳转与折叠）、120 页演示文稿的打开、编辑与用时。

需要 node 和 PATH 上的 dsh 0.2.1-rc.x（或 `QUIRE_E2E_DSH`，可指向 bin.js 配 `QUIRE_E2E_NODE`）；web 还需要 Chrome 或 Edge（或 `QUIRE_E2E_CHROME`）。

## 许可

MIT
