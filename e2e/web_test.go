package e2e

import (
	"archive/zip"
	"bufio"
	"context"
	_ "embed"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"testing"
	"time"
)

// bigJS writes the large files (see big.mjs).
//
//go:embed big.mjs
var bigJS string

// fixtureJS writes the three files the person edits, with quire's own
// writers: argv is the plugin's lib directory and the workspace.
const fixtureJS = `
import fs from 'node:fs'
import path from 'node:path'
const [lib, ws] = process.argv.slice(2)
const { createDocx } = await import(path.join(lib, 'docx.js'))
const { createXlsx } = await import(path.join(lib, 'xlsx.js'))
const { createPptx } = await import(path.join(lib, 'pptx.js'))
const io = { loadImage: async (p) => fs.readFileSync(path.resolve(ws, p)) }
const md = '# 季度报告\n\n本季度收入 **120 万元**。\n\n| 地区 | 收入 |\n| --- | ---: |\n| 华东 | 120 |\n| 华北 | 98 |\n\n- 华东区贡献最大\n- 成本下降 5%\n\n![趋势图](chart.png)\n\n## 结论\n\n继续保持增长。\n'
fs.writeFileSync(path.join(ws, '季度报告.docx'), (await createDocx(md, io)).buffer)
const xl = createXlsx([{ name: '明细', header: true, rows: [['项目', '单价', '数量', '小计'], ['服务器', 8000, 5, '=B2*C2'], ['带宽', 1200, 12, '=B3*C3'], ['合计', null, null, '=SUM(D2:D3)']] }])
fs.writeFileSync(path.join(ws, '预算.xlsx'), xl.buffer)
const pp = await createPptx([{ layout: 'title', title: '季度经营回顾', subtitle: '战略部', notes: '开场白' }, { title: '核心要点', bullets: ['收入增长 13%', '成本下降 5%'] }], {}, io)
fs.writeFileSync(path.join(ws, '经营回顾.pptx'), pp.buffer)
`

// agentEditJS changes a cell the way the Agent's tools would, while the
// file is open in the editor.
const agentEditJS = `
import fs from 'node:fs'
import path from 'node:path'
const [lib, file] = process.argv.slice(2)
const { editXlsxUi } = await import(path.join(lib, 'xlsx.js'))
fs.writeFileSync(file, editXlsxUi(fs.readFileSync(file), [{ op: 'input', sheet: '明细', cell: 'B2', text: '9000' }]).buffer)
`

// Page helpers, run before each step: find a visible element by text or
// selector and report the centre of its box for a real mouse click.
const pageJS = `
window.__q = {
  // reveal opens the ribbon tab a toolbar control sits on.
  reveal (n) { const p = n?.closest?.('.q-panel'); if (p && p.hidden) [...document.querySelectorAll('.quire .q-tab')].find((t) => t.textContent === p.dataset.tab)?.click(); return n },
  vis (e) { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && r.bottom > 0 && r.right > 0 },
  // byText finds host UI by its text; the editor's own tabs (it has a
  // 文件 tab too) never count.
  byText (text, right) {
    const own = (e) => [...e.childNodes].filter((n) => n.nodeType === 3).map((n) => n.data).join('').trim()
    // Text in a display: contents box is shown by the nearest box around it.
    const shown = (e) => { while (e && !this.vis(e)) e = e.parentElement; return e }
    const all = [...document.querySelectorAll('body *')].filter((e) => own(e) === text && !e.closest('.quire')).map(shown).filter(Boolean)
    const pick = right ? all.filter((e) => e.getBoundingClientRect().left > innerWidth / 2) : all
    return pick[pick.length - 1]
  },
  // draft is the chat composer's text.
  draft () { return document.querySelector('[data-lexical-editor="true"]')?.innerText ?? '' },
  // pins is the annotation cards above the composer, as their text.
  pins () { return [...document.querySelectorAll('.qpd-card')].map((c) => c.querySelector('.qpd-where').textContent + '：' + (c.querySelector('.qpd-note')?.textContent ?? '')) },
  // cur is the editor on screen; other tabs keep theirs in the page.
  cur () { return [...document.querySelectorAll('.quire')].find((q) => this.vis(q)) },
  centre (e) { if (!e) return null; e.scrollIntoView({ block: 'nearest' }); const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 } }
}
`

func TestWebEditor(t *testing.T) {
	s := rep.section(t, "dsh web：人在右侧栏直接编辑 Word / Excel / PowerPoint")
	dshBin := findDsh()
	chromeBin := chromePath()
	node, _ := exec.LookPath("node")
	if dshBin == "" || chromeBin == "" || node == "" {
		s.step("找到 dsh、Chrome 和 node", fmt.Errorf("缺少 dsh=%q chrome=%q node=%q", dshBin, chromeBin, node), "")
		return
	}
	lib, _ := filepath.Abs(filepath.Join("..", "plugin", "lib"))
	work := shortTemp(t)
	ws := filepath.Join(work, "ws")
	os.MkdirAll(ws, 0o755)
	d := newDsh(dshBin, work)

	// Files, a Session in their folder, and the web profile with quire.
	err := chart(filepath.Join(ws, "chart.png"))
	if err == nil {
		os.WriteFile(filepath.Join(work, "fixture.mjs"), []byte(fixtureJS), 0o644)
		_, err = run(nil, work, node, filepath.Join(work, "fixture.mjs"), lib, ws)
	}
	if err == nil {
		os.WriteFile(filepath.Join(work, "big.mjs"), []byte(bigJS), 0o644)
		_, err = run(nil, work, node, filepath.Join(work, "big.mjs"), lib, ws)
	}
	tgz, perr := packPlugin(work)
	if err == nil {
		err = perr
	}
	if err == nil {
		_, _, err = d.install("quire-e2e", tgz)
	}
	model := startModel(t, nil)
	if err == nil {
		_, err = d.turn("quire-e2e", ws, model, "看一下这几个文件")
	}
	if err == nil {
		_, err = d.run("--profile", "quire-web", "--from-default-profile", "web", "--dump-config")
	}
	if err == nil {
		_, err = d.run("plugin", "--profile", "quire-web", "add", tgz)
	}
	if !s.step("准备：用 quire 写出三个文件，sdk profile 跑一轮建出会话，web 模板 profile 装上本包", err, "") {
		return
	}
	url, err := startWeb(t, d, ws, model)
	if !s.step("启动 dsh web（随机端口）", err, "") {
		return
	}
	c, err := startChrome(t, chromeBin, work)
	if !s.step("无头 Chrome", err, "") {
		return
	}
	ctx, cancel := context.WithTimeout(context.Background(), 6*time.Minute)
	defer cancel()
	c.send(ctx, c.session, "Emulation.setDeviceMetricsOverride", map[string]any{"width": 1440, "height": 900, "deviceScaleFactor": 1, "mobile": false}, nil)
	c.navigate(ctx, url)
	c.waitLoaded(ctx)
	p := &page{c: c, ctx: ctx}

	_, err = p.waitJS(`return !!(window.__DSH_BOOT__?.entries ?? []).find((e) => e.id === '@precinct/quire') && !!document.getElementById('quire-editor-css')`, 30*time.Second)
	if !s.step("客户端模块：boot 清单里有 @precinct/quire，apply 已运行（样式已注入）", err, "") {
		p.shot("web-fail-boot.png")
		return
	}

	// Open the Session, the right Sidebar's files, then each file.
	// A fresh home first shows a preview notice and a sign-in prompt;
	// the editor needs neither.
	// They come up a moment after load, one after the other.
	for end := time.Now().Add(8 * time.Second); time.Now().Before(end); time.Sleep(400 * time.Millisecond) {
		v, _ := p.js(`const d = [...document.querySelectorAll('[role=dialog]')].find((e) => __q.vis(e)); if (!d) return 'none'; const b = [...d.querySelectorAll('button')].find((b) => b.textContent.trim() === '继续') ?? d.querySelector('button[aria-label="关闭"], button[aria-label="Close"]'); return b ? __q.centre(b) : 'esc'`)
		switch string(v) {
		case `"none"`, "":
		case `"esc"`:
			p.key("Escape")
		default:
			var at struct{ X, Y float64 }
			json.Unmarshal(v, &at)
			p.click(at.X, at.Y, 1)
		}
	}
	for i := 0; i < 3; i++ {
		if err = p.clickText("看一下这几个文件", false); err == nil {
			if _, err = p.waitJS(`return document.title.includes('看一下这几个文件')`, 4*time.Second); err == nil {
				break
			}
		}
	}
	if err == nil {
		time.Sleep(1000 * time.Millisecond)
		if v, _ := p.js(`return !!__q.byText('工作区文件', true)`); string(v) != "true" {
			err = p.clickJS(`return __q.centre(document.querySelector('button[aria-label="打开右侧边栏"]'))`)
		}
	}
	if err == nil {
		err = p.clickText("工作区文件", true)
	}
	open := func(name string) error {
		if v, _ := p.js(`return !!__q.byText(` + jsStr(name) + `, true)`); string(v) != "true" {
			// With many files open the side bar's tab strip scrolls; bring
			// its first tab (文件) back into view.
			p.js(`for (const e of document.querySelectorAll('body *')) { if (e.closest('.quire') || ![...e.childNodes].some((n) => n.nodeType === 3 && n.data.trim() === '文件')) continue; let b = e; while (b && !b.getBoundingClientRect().width) b = b.parentElement; b?.scrollIntoView({ block: 'nearest', inline: 'nearest' }) } return true`)
			if err := p.clickText("文件", true); err != nil {
				return err
			}
		}
		return p.clickText(name, true)
	}
	if err == nil {
		err = open("预算.xlsx")
	}
	if err == nil {
		_, err = p.waitJS(`return !!document.querySelector('.quire .qx-canvas') && document.querySelector('.qx-tab')?.textContent === '明细'`, 20*time.Second)
	}
	if !s.step("在会话的右侧栏打开 预算.xlsx，显示的是 Quire 编辑器", err, "") {
		p.shot("web-fail-open.png")
		return
	}

	// --- spreadsheet
	xlsx := filepath.Join(ws, "预算.xlsx")
	err = p.goCell("A5")
	if err == nil {
		err = p.typeText("税费")
	}
	if err == nil {
		err = p.key("Tab")
	}
	if err == nil {
		err = p.goCell("D5")
	}
	if err == nil {
		err = p.typeText("=D4*0.06")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool {
			return regexp.MustCompile(`<c r="D5"[^>]*><f>D4\*0\.06</f><v>3264</v>`).MatchString(x)
		}, func(x string) bool {
			return regexp.MustCompile(`<c r="A5"`).MatchString(x) && strings.Contains(sharedOrInline(xlsx), "税费")
		})
	}
	s.step("表格：A5 输入文字、D5 输入公式，自动保存；文件里有公式和重算后的缓存值 3264", err, "")
	p.shot("web-xlsx.png")

	err = p.goCell("A1")
	if err == nil {
		err = p.goCell("A4:D5")
	}
	plain, _ := zipPart(xlsx, "xl/worksheets/sheet1.xml")
	if err == nil {
		err = p.key("b", "mod")
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool { return x != plain })
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire .q-side button[aria-label^="撤销"]')))`)
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool { return x == plain })
	}
	s.step("表格：选区 ⌘B 加粗写进样式；工具栏撤销把文件恢复到加粗前", err, "")

	_, err = run(nil, work, node, writeTemp(work, "agent.mjs", agentEditJS), lib, xlsx)
	if err == nil {
		_, err = p.waitJS(`return (document.querySelector('.qx-name').value = 'B2', document.querySelector('.qx-name').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })), document.querySelector('.qx-fx').value === '9000')`, 20*time.Second)
	}
	s.step("Agent 在外部改了 B2：编辑器自动载入新内容（编辑栏显示 9000）", err, "")

	// Formula help: typing a name offers functions; Tab takes one.
	err = p.goCell("F6")
	if err == nil {
		err = p.typeText("=SUM")
	}
	if err == nil {
		_, err = p.waitJS(`const a = document.querySelector('.qx-ac'); return !a.hidden && a.querySelector('.qx-ac-item.on b')?.textContent === 'SUM'`, 5*time.Second)
	}
	if err == nil {
		err = p.key("Tab")
	}
	if err == nil {
		err = p.typeText("D2:D3)")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool {
			return regexp.MustCompile(`<c r="F6"[^>]*><f>SUM\(D2:D3\)</f><v>59400</v>`).MatchString(x)
		})
	}
	s.step("表格：输入 =SUM 弹出函数提示，Tab 补全后写成 SUM(D2:D3)，缓存值 59400", err, "")

	// Sort the two data rows by 项目 ascending: 带宽 comes first and the
	// formulas move with their rows.
	err = p.goCell("A1:D3")
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label="升序排序"]')))`)
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool {
			return regexp.MustCompile(`<c r="D2"[^>]*><f>B2\*C2</f><v>14400</v>`).MatchString(x) && regexp.MustCompile(`<c r="D3"[^>]*><f>B3\*C3</f><v>45000</v>`).MatchString(x)
		})
	}
	s.step("表格：选中 A1:D3 升序排序，带宽排到第 2 行，公式随行移动", err, "")

	// Fill handle: drag G1 (第1周) down two rows.
	cellBox := func(a, b string) (x1, y1, x2, y2 float64, err error) {
		var p1, p2 struct{ X, Y float64 }
		read := func(ref string, at *struct{ X, Y float64 }) error {
			if err := p.goCell(ref); err != nil {
				return err
			}
			v, err := p.js(`const r = document.querySelector('.qx-input').getBoundingClientRect(); return { x: r.left, y: r.top }`)
			if err != nil {
				return err
			}
			return json.Unmarshal(v, at)
		}
		if err = read(a, &p1); err == nil {
			err = read(b, &p2)
		}
		return p1.X, p1.Y, p2.X, p2.Y, err
	}
	err = p.goCell("G1")
	if err == nil {
		err = p.typeText("第1周")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		// G1's box runs from its top-left to H2's; select G1 last.
		var hx, hy, gy float64
		hx, hy, _, gy, err = cellBox("H2", "G1")
		if err == nil {
			h := hy - gy
			err = p.drag(hx-1, hy-1, hx-1, hy+h*1.5)
		}
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool {
			return regexp.MustCompile(`<c r="G3"`).MatchString(x) && strings.Contains(sharedOrInline(xlsx), "第2周") && strings.Contains(sharedOrInline(xlsx), "第3周")
		})
	}
	s.step("表格：拖动填充柄把 G1「第1周」向下填充，得到 第2周、第3周", err, "")

	// Borders from the right-click menu.
	styles0, _ := zipPart(xlsx, "xl/styles.xml")
	ax1, ay1, ax2, ay2, err := cellBox("A1", "B2")
	if err == nil {
		err = p.goCell("A1:D4")
	}
	if err == nil {
		err = p.rclick((ax1+ax2)/2, (ay1+ay2)/2)
	}
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.startsWith('所有框线')); return b && __q.centre(b)`)
	}
	if err == nil {
		err = waitFile(xlsx, "xl/styles.xml", func(x string) bool { return x != styles0 && strings.Contains(x, `<left style="thin"`) })
	}
	s.step("表格：右键菜单「所有框线」给 A1:D4 加上细边框", err, "")

	// AutoFilter, then hide 带宽 from the column A filter list.
	err = p.goCell("A1:D5")
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label="筛选"]')))`)
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool { return strings.Contains(x, `<autoFilter ref="A1:D5"`) })
	}
	if err == nil {
		err = p.goCell("C3")
	}
	if err == nil {
		time.Sleep(500 * time.Millisecond)
		err = p.click(ax2-8, (ay1+ay2)/2, 1)
	}
	if err == nil {
		_, err = p.waitJS(`const l = [...document.querySelectorAll('.q-filter-item')].find((e) => e.textContent.startsWith('带宽')); if (!l) return false; l.querySelector('input').click(); return true`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.q-filter-ok'))`)
	}
	if err == nil {
		err = waitFile(xlsx, "xl/worksheets/sheet1.xml", func(x string) bool { return regexp.MustCompile(`<row r="2"[^>]*hidden="1"`).MatchString(x) })
	}
	s.step("表格：开启筛选，在 A 列筛选面板里去掉「带宽」，第 2 行被隐藏", err, "")
	p.shot("web-xlsx-more.png")

	// Charts: insert from a range, drag, change type, delete.
	charts := func() string { return partsLike(xlsx, "xl/charts/") }
	err = p.goCell("A1:C3")
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="插入图表"]')))`)
	}
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '柱形图'); return b && __q.centre(b)`)
	}
	if err == nil {
		err = waitFile(xlsx, "", func(string) bool {
			c := charts()
			return strings.Contains(c, "<c:barChart") && strings.Contains(c, "'明细'!$B$2:$B$3") && strings.Contains(c, "'明细'!$A$2:$A$3")
		}, func(string) bool { return strings.Contains(partsLike(xlsx, "xl/worksheets/sheet1.xml"), "<drawing ") })
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qx-chart.on svg rect').length >= 4`, 10*time.Second)
	}
	s.step("表格：选中 A1:C3 插入柱形图，系列引用 B2:B3、分类 A2:A3，网格上画出图表", err, "")
	p.shot("web-xlsx-chart.png")

	drawing0 := partsLike(xlsx, "xl/drawings/drawing")
	var box struct{ X, Y float64 }
	raw, err := p.js(`const r = document.querySelector('.qx-chart').getBoundingClientRect(); return { X: r.left + r.width / 2, Y: r.top + r.height / 2 }`)
	if err == nil {
		err = json.Unmarshal(raw, &box)
	}
	if err == nil {
		err = p.drag(box.X, box.Y, box.X+130, box.Y+70)
	}
	if err == nil {
		err = waitFile(xlsx, "", func(string) bool {
			d := partsLike(xlsx, "xl/drawings/drawing")
			return d != drawing0 && strings.Contains(d, "twoCellAnchor")
		})
	}
	s.step("表格：拖动图表，锚点写进 drawing 的 twoCellAnchor", err, "")

	_, err = p.waitJS(`return !!document.querySelector('.qx-chart')`, 10*time.Second)
	if err == nil {
		raw, err = p.js(`const r = document.querySelector('.qx-chart').getBoundingClientRect(); return { X: r.left + 60, Y: r.top + 50 }`)
	}
	if err == nil {
		err = json.Unmarshal(raw, &box)
	}
	if err == nil {
		err = p.rclick(box.X, box.Y)
	}
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim().endsWith('折线图')); return b && __q.centre(b)`)
	}
	if err == nil {
		err = waitFile(xlsx, "", func(string) bool {
			c := charts()
			return strings.Contains(c, "<c:lineChart") && !strings.Contains(c, "<c:barChart")
		})
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qx-chart svg path').length >= 2`, 10*time.Second)
	}
	s.step("表格：右键图表改成折线图", err, "")

	if err == nil {
		_, err = p.waitJS(`return !!document.querySelector('.qx-chart')`, 10*time.Second)
	}
	if err == nil {
		raw, err = p.js(`const r = document.querySelector('.qx-chart').getBoundingClientRect(); return { X: r.left + 60, Y: r.top + 50 }`)
	}
	if err == nil {
		err = json.Unmarshal(raw, &box)
	}
	if err == nil {
		err = p.click(box.X, box.Y, 1)
	}
	if err == nil {
		err = p.key("Delete")
	}
	if err == nil {
		err = waitFile(xlsx, "", func(string) bool { return charts() == "" })
	}
	if err == nil {
		_, err = p.waitJS(`return !document.querySelector('.qx-chart')`, 10*time.Second)
	}
	s.step("表格：选中图表按 Delete 删除，图表部件一并移除", err, "")

	// --- document
	docx := filepath.Join(ws, "季度报告.docx")
	before, _ := zipPart(docx, "word/document.xml")
	err = open("季度报告.docx")
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.quire .qd-page .qd-p').length > 3`, 20*time.Second)
	}
	if err == nil {
		_, err = p.js(`const ps = document.querySelectorAll('.qd-page > .qd-p'); const last = ps[ps.length - 1]; document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(last); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
	}
	if err == nil {
		err = p.typeText("下季度目标 150 万。")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		err = p.typeText("新增段落")
	}
	if err == nil {
		_, err = p.js(`const ps = document.querySelectorAll('.qd-page > .qd-p'); const r = document.createRange(); r.selectNodeContents(ps[ps.length - 1]); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="加粗"]')))`)
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool {
			return strings.Contains(x, "继续保持增长。下季度目标 150 万。") && regexp.MustCompile(`<w:b/>.{0,200}?<w:t[^>]*>新增段落</w:t>`).MatchString(x)
		}, func(x string) bool {
			return strings.Count(x, "<w:drawing") == strings.Count(before, "<w:drawing") && strings.Count(x, "<w:drawing") > 0
		})
	}
	s.step("文档：段尾打字、回车分段、选中加粗，自动保存；图片原样保留", err, "")

	{
		_, err = p.js(`const td = document.querySelector('.qd-t td[data-r="2"][data-c="2"]'); td.focus(); const r = document.createRange(); r.selectNodeContents(td); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
	}
	if err == nil {
		err = p.typeText("130")
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool { return strings.Contains(x, ">130<") && !strings.Contains(x, ">120<") })
	}
	s.step("文档：改表格单元格（120 → 130）", err, "")

	selectPara := `const p = [...document.querySelectorAll('.qd-page > .qd-p')].find((e) => e.textContent.includes('本季度收入')); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(p); getSelection().removeAllRanges(); getSelection().addRange(r); return true`
	setSelect := func(label, value string) string {
		return `const s = document.querySelector('.quire select[aria-label="` + label + `"]'); s.value = ` + jsStr(value) + `; s.dispatchEvent(new Event('change', { bubbles: true })); return true`
	}
	_, err = p.js(selectPara)
	if err == nil {
		_, err = p.js(setSelect("字体", "黑体"))
	}
	if err == nil {
		_, err = p.js(selectPara)
	}
	if err == nil {
		_, err = p.js(setSelect("行距", "360"))
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool {
			return regexp.MustCompile(`w:eastAsia="黑体"[^>]*/>.{0,300}?本季度收入`).MatchString(x) && regexp.MustCompile(`w:line="360"`).MatchString(x)
		})
	}
	s.step("文档：选中段落改字体为黑体、行距 1.5 倍，写进 w:rFonts 和 w:spacing", err, "")

	err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="查找和替换"]')))`)
	if err == nil {
		// The box starts with the selected text, selected: typing replaces it.
		_, err = p.waitJS(`const i = document.querySelector('.q-find-in[aria-label="查找"]'); return document.activeElement === i && i.selectionEnd - i.selectionStart === i.value.length`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("华东区")
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelector('.q-find-n').textContent === '1/1'`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.q-find-in[aria-label="替换为"]'))`)
	}
	if err == nil {
		err = p.typeText("华东地区")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...document.querySelectorAll('.q-find button')].find((b) => b.textContent === '全部替换'))`)
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool { return strings.Contains(x, "华东地区贡献最大") })
	}
	p.key("Escape")
	s.step("文档：⌘F 查找栏显示 1/1，全部替换「华东区」→「华东地区」", err, "")

	pics, _ := zipPart(docx, "word/document.xml")
	_, err = p.js(selectPara)
	if err == nil {
		_, err = p.js(`const c = document.createElement('canvas'); c.width = 80; c.height = 40; const g = c.getContext('2d'); g.fillStyle = '#d33'; g.fillRect(0, 0, 80, 40); const b = await new Promise((r) => c.toBlob(r, 'image/png')); const dt = new DataTransfer(); dt.items.add(new File([b], 'red.png', { type: 'image/png' })); document.querySelector('.qd-page').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return true`)
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool { return strings.Count(x, "<w:drawing") == strings.Count(pics, "<w:drawing")+1 })
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qd-page img.qd-img').length === 2`, 10*time.Second)
	}
	s.step("文档：粘贴一张图片，插在当前段落后并保存进文件", err, "")

	err = p.clickJS(`return __q.centre(document.querySelector('.qd-hdr'))`)
	if err == nil {
		err = p.typeText("内部资料")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...document.querySelectorAll('.qd-page > .qd-p')].find((e) => e.textContent.includes('结论')))`)
	}
	if err == nil {
		err = waitFile(docx, "", func(string) bool { return strings.Contains(partsLike(docx, "word/header"), "内部资料") })
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="页码"]')))`)
	}
	if err == nil {
		err = waitFile(docx, "", func(string) bool { return strings.Contains(partsLike(docx, "word/footer"), "PAGE") })
	}
	s.step("文档：在页眉写「内部资料」，工具栏 # 在页脚插入页码", err, "")

	err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="页面设置"]')))`)
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim().endsWith('横向')); return b && __q.centre(b)`)
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool { return strings.Contains(x, `w:orient="landscape"`) })
	}
	if err == nil {
		_, err = p.waitJS(`const s = document.querySelector('.qd-sheet'); return parseFloat(s.style.width) > 1000`, 10*time.Second)
	}
	s.step("文档：页面设置改为横向，纸张变宽", err, "")

	// Comments: select words, comment, reply, resolve.
	_, err = p.js(`const p = [...document.querySelectorAll('.qd-page > .qd-p')].find((e) => e.textContent.includes('本季度收入')); const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode()) && !n.data.includes('本季度收入')); const i = n.data.indexOf('本季度收入'); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 5); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="新建批注"]')))`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.activeElement === document.querySelector('.qd-draft textarea')`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("请核对数字")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qd-draft .qd-cok'))`)
	}
	if err == nil {
		err = waitFile(docx, "word/comments.xml", func(x string) bool { return strings.Contains(x, "请核对数字") }, func(string) bool {
			d, _ := zipPart(docx, "word/document.xml")
			return regexp.MustCompile(`<w:commentRangeStart w:id="(\d+)"/>.*本季度收入.*<w:commentRangeEnd w:id="(\d+)"/>.*<w:commentReference w:id="\d+"/>`).MatchString(d)
		})
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qd-card').length === 1 && !!document.querySelector('.qd-page .qd-cm[data-cm="r"]') && document.querySelector('.qd-card .qd-cquote').textContent === '本季度收入'`, 10*time.Second)
	}
	s.step("文档：选中「本季度收入」新建批注，写进 comments.xml，正文有批注范围和引用", err, "")

	err = p.clickJS(`return __q.centre([...document.querySelectorAll('.qd-card button')].find((b) => b.textContent === '回复'))`)
	if err == nil {
		_, err = p.waitJS(`return document.activeElement === document.querySelector('.qd-card .qd-cta')`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("已核对")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qd-card .qd-cok'))`)
	}
	if err == nil {
		err = waitFile(docx, "word/commentsExtended.xml", func(x string) bool { return strings.Contains(x, "paraIdParent") }, func(string) bool {
			c, _ := zipPart(docx, "word/comments.xml")
			return strings.Contains(c, "已核对")
		})
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qd-card .qd-creply').length === 1`, 10*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...document.querySelectorAll('.qd-card button')].find((b) => b.textContent === '解决'))`)
	}
	if err == nil {
		err = waitFile(docx, "word/commentsExtended.xml", func(x string) bool { return strings.Contains(x, `w15:done="1"`) })
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelector('.qd-card.done') !== null`, 10*time.Second)
	}
	s.step("文档：回复批注写进 commentsExtended（paraIdParent），再标记为已解决", err, "")
	p.shot("web-docx-comments.png")

	// Table of contents: insert at the top, rename a heading, update.
	_, err = p.js(`const p = document.querySelector('.qd-page > .qd-p'); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="插入目录"]')))`)
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool {
			return strings.Contains(x, `TOC \o "1-3"`) && strings.Contains(x, "_Toc") && strings.Index(x, "TOC \\o") < strings.Index(x, ">季度报告<")
		})
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qd-p[data-toc]').length >= 3 && [...document.querySelectorAll('.qd-p[data-toc]')].some((e) => e.textContent.includes('结论'))`, 10*time.Second)
	}
	s.step("文档：在开头插入目录（TOC 域 + 书签），目录列出「季度报告」「结论」", err, "")

	_, err = p.js(`const p = [...document.querySelectorAll('.qd-page > .qd-p')].find((e) => !e.dataset.toc && e.textContent === '结论'); document.querySelector('.qd-page').focus(); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); return true`)
	if err == nil {
		err = p.typeText("与展望")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="插入目录"]')))`)
	}
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.trim() === '更新目录'); return b && __q.centre(b)`)
	}
	if err == nil {
		err = waitFile(docx, "word/document.xml", func(x string) bool { return strings.Count(x, "结论与展望") >= 2 })
	}
	if err == nil {
		_, err = p.waitJS(`return [...document.querySelectorAll('.qd-p[data-toc]')].some((e) => e.textContent.includes('结论与展望'))`, 10*time.Second)
	}
	s.step("文档：标题改成「结论与展望」后更新目录，目录项跟着变", err, "")
	p.shot("web-docx-more.png")
	p.shot("web-docx.png")

	// --- presentation
	pptx := filepath.Join(ws, "经营回顾.pptx")
	err = open("经营回顾.pptx")
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.quire .qp-thumb').length === 2`, 20*time.Second)
	}
	if err == nil {
		err = p.clickJS(`const n = [...document.querySelectorAll('.qp-stage .slide [data-t="txt"]')].find((e) => e.textContent.includes('季度经营回顾')); return __q.centre(n)`, 2)
	}
	if err == nil {
		err = p.key("a", "mod")
	}
	if err == nil {
		err = p.typeText("季度经营回顾（终稿）")
	}
	if err == nil {
		err = p.key("Escape")
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool {
			return strings.Contains(x, ">季度经营回顾（终稿）<") && strings.Count(x, "季度经营回顾") == 1
		})
	}
	s.step("演示：双击标题进入编辑，⌘A 全选后改写，Esc 保存", err, "")

	var y0 string
	if err == nil {
		x, _ := zipPart(pptx, "ppt/slides/slide1.xml")
		y0 = x
		var at struct{ X, Y float64 }
		v, jerr := p.js(`const n = [...document.querySelectorAll('.qp-stage .slide [data-t="txt"]')].find((e) => e.textContent.includes('战略部')); return __q.centre(n)`)
		if jerr == nil && json.Unmarshal(v, &at) == nil {
			err = p.drag(at.X, at.Y, at.X+40, at.Y+30)
		} else {
			err = fmt.Errorf("副标题不在页面上")
		}
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return x != y0 && strings.Count(x, "<a:off ") > strings.Count(y0, "<a:off ") })
	}
	s.step("演示：拖动副标题，位置写进 a:xfrm", err, "")

	{
		err = p.clickJS(`return __q.centre(document.querySelector('.qp-notes'))`)
	}
	if err == nil {
		_, err = p.js(`const n = document.querySelector('.qp-notes'); n.setSelectionRange(n.value.length, n.value.length); return true`)
	}
	if err == nil {
		err = p.typeText("，先讲结论")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[1])`)
	}
	if err == nil {
		err = waitFile(pptx, "", func(string) bool { return strings.Contains(notesOf(pptx), "开场白，先讲结论") })
	}
	s.step("演示：改备注，切换页面时保存（原有备注没有被清掉）", err, "")

	{
		_, err = p.js(`const s = document.querySelector('.quire select[aria-label="新建幻灯片"]'); s.value = s.options[2].value; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/presentation.xml", func(x string) bool { return strings.Count(x, "<p:sldId ") == 3 })
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire .q-side button[aria-label^="撤销"]')))`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/presentation.xml", func(x string) bool { return strings.Count(x, "<p:sldId ") == 2 })
	}
	s.step("演示：用文件自己的版式新建幻灯片，再撤销", err, "")

	err = p.clickJS(`return __q.centre(document.querySelectorAll('.qp-thumb')[0])`)
	if err == nil {
		_, err = p.js(`const s = document.querySelector('.quire select[aria-label="插入形状"]'); s.value = 'star5'; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return strings.Count(x, `prst="star5"`) == 1 })
	}
	if err == nil {
		_, err = p.waitJS(`return !document.querySelector('.qp-box').hidden`, 10*time.Second)
	}
	if err == nil {
		_, err = p.js(`const i = document.querySelector('.quire label[aria-label="填充颜色"] input'); i.value = '#ff0000'; i.dispatchEvent(new Event('input', { bubbles: true })); i.dispatchEvent(new Event('change', { bubbles: true })); return true`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return strings.Contains(x, `<a:srgbClr val="FF0000"/>`) })
	}
	s.step("演示：插入五角星形状，填充改成红色", err, "")

	_, err = p.waitJS(`return !document.querySelector('.qp-box').hidden`, 10*time.Second)
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qp-box'))`)
	}
	if err == nil {
		err = p.key("d", "mod")
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return strings.Count(x, `prst="star5"`) == 2 })
	}
	s.step("演示：⌘D 复制出一个相同的形状", err, "")

	// The copy sits on top, a little lower: the first star's top point is
	// clear of it, so a Shift-click there adds the first star.
	err = p.clickJS(`const ns = [...document.querySelectorAll('.qp-stage [data-k]')]; return __q.centre(ns[ns.length - 1])`)
	if err == nil {
		var at struct{ X, Y float64 }
		v, jerr := p.js(`const ns = [...document.querySelectorAll('.qp-stage [data-k]')]; const r = ns[ns.length - 2].getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height * 0.07 }`)
		if jerr == nil && json.Unmarshal(v, &at) == nil {
			err = p.shiftClick(at.X, at.Y)
		} else {
			err = fmt.Errorf("星形不在页面上")
		}
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qp-ol').length === 2`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="组合所选形状"]')))`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return strings.Count(x, "<p:grpSp>") == 1 })
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelectorAll('.qp-stage [data-g]').length === 2`, 10*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="取消组合"]')))`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return !strings.Contains(x, "<p:grpSp>") && strings.Count(x, `prst="star5"`) == 2 })
	}
	s.step("演示：Shift 多选两个星形，组合写进 p:grpSp，再取消组合", err, "")

	_, err = p.waitJS(`return document.querySelectorAll('.qp-ol').length === 2`, 10*time.Second)
	if err == nil {
		_, err = p.js(`const s = __q.reveal(document.querySelector('.quire select[aria-label="动画效果"]')); s.value = 'fade'; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool {
			return strings.Contains(x, "<p:timing>") && strings.Count(x, `presetID="10"`) == 2
		})
	}
	if err == nil {
		_, err = p.waitJS(`return [...document.querySelectorAll('.qp-badge')].length === 2`, 10*time.Second)
	}
	if err == nil {
		_, err = p.js(`const s = __q.reveal(document.querySelector('.quire select[aria-label="切换效果"]')); s.value = 'push'; s.dispatchEvent(new Event('change', { bubbles: true })); return true`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool { return strings.Contains(x, "<p:transition") && strings.Contains(x, "<p:push") })
	}
	s.step("演示：两个星形加「淡入」动画写进 p:timing（编辑区显示序号），切换效果「推入」写进 p:transition", err, "")

	err = p.clickJS(`return __q.centre([...document.querySelectorAll('button[aria-label="撤销多步"]')].find((b) => __q.vis(b)))`)
	if err == nil {
		_, err = p.waitJS(`return [...document.querySelectorAll('.q-menu button')].map((b) => b.textContent).join('|').startsWith('撤销 1 步：切换效果|撤销 2 步：动画')`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...document.querySelectorAll('.q-menu button')][1])`)
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slides/slide1.xml", func(x string) bool {
			return !strings.Contains(x, "<p:transition") && !strings.Contains(x, "<p:timing>") && strings.Count(x, `prst="star5"`) == 2
		})
	}
	s.step("通用：撤销下拉按步列出操作，选「撤销 2 步」一次撤掉切换和动画，形状保留", err, "")

	err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="编辑母版与版式"]')))`)
	if err == nil {
		_, err = p.waitJS(`return document.querySelector('.qp').classList.contains('qp-master') && document.querySelectorAll('.qp-thumb.layout').length > 0`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qp-thumb'))`)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label="插入文本框"]')))`)
	}
	if err == nil {
		_, err = p.waitJS(`return !!document.querySelector('.qp-editing')`, 10*time.Second)
	}
	if err == nil {
		err = p.typeText("内部资料")
	}
	if err == nil {
		err = p.key("Escape")
	}
	if err == nil {
		err = waitFile(pptx, "ppt/slideMasters/slideMaster1.xml", func(x string) bool { return strings.Contains(x, ">内部资料<") })
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qp-mbar button'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return !document.querySelector('.qp').classList.contains('qp-master') && document.querySelector('.qp-stage .slide').textContent.includes('内部资料')`, 10*time.Second)
	}
	s.step("演示：幻灯片母版视图里插入文本框，写进 slideMaster1.xml；回到普通视图，幻灯片上也显示", err, "")
	p.shot("web-pptx-master.png")

	_, err = p.waitJS(`return document.querySelector('.qp-stxt')?.textContent.startsWith('幻灯片 1 / 2')`, 5*time.Second)
	if err == nil {
		err = p.clickJS(`return __q.centre([...document.querySelector('.qp-status .q-zoom').children][2])`)
	}
	if err == nil {
		_, err = p.waitJS(`const t = document.querySelector('.qp-status .q-zoom-p').textContent; const m = /scale\(([\d.]+)\)/.exec(document.querySelector('.qp-scaler>.slide').style.transform); return !!m && Math.round(+m[1] * 100) + '%' === t && document.querySelector('.qp').closest('.quire').querySelector('select[aria-label="缩放"]').value + '%' === t`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qp-status .q-zoom-p'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelector('.qp').closest('.quire').querySelector('select[aria-label="缩放"]').value === ''`, 5*time.Second)
	}
	s.step("通用：状态栏显示「幻灯片 1 / 2」，缩放滑块 + 放大一档（比例与画面一致），点百分比回到适应窗口", err, "")

	err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.quire button[aria-label^="从当前幻灯片放映"]')))`)
	if err == nil {
		_, err = p.waitJS(`const s = document.querySelector('.qp-show'); return !!s && s.textContent.includes('季度经营回顾')`, 5*time.Second)
	}
	if err == nil {
		p.shot("web-pptx-show.png")
		err = p.key("ArrowRight")
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelector('.qp-show')?.textContent.includes('核心要点')`, 5*time.Second)
	}
	if err == nil {
		err = p.key("Escape")
	}
	if err == nil {
		_, err = p.waitJS(`return !document.querySelector('.qp-show')`, 5*time.Second)
	}
	s.step("演示：放映从当前页开始，→ 翻到第 2 页，Esc 退出", err, "")
	p.shot("web-pptx.png")

	err = p.clickJS(`return __q.centre(__q.reveal(document.querySelector('.qp').closest('.quire').querySelector('button[aria-label^="新建空白"]')))`)
	if err == nil {
		_, err = p.waitJS(`return [...document.querySelectorAll('.q-menu button')].some((b) => b.textContent.includes('空白 Word 文档'))`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.includes('空白 Word 文档')))`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.querySelector('.q-dlg [name=name]')?.value === '新建文档.docx'`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.q-dlg .q-dlg-ok'))`)
	}
	if err == nil {
		err = waitFile(filepath.Join(ws, "新建文档.docx"), "word/document.xml", func(x string) bool { return strings.Contains(x, "w:body") })
	}
	if err == nil {
		_, err = p.waitJS(`return [...document.querySelectorAll('.qd-stxt')].some((n) => __q.vis(n) && n.textContent.includes('共 1 页'))`, 20*time.Second)
	}
	s.step("通用：文件 › 新建 › 空白 Word 文档，写进工作区并在新标签里用 Quire 打开", err, "")
	p.shot("web-new-docx.png")

	// --- large files
	big := filepath.Join(ws, "大表格.xlsx")
	t0 := time.Now()
	err = open("大表格.xlsx")
	if err == nil {
		_, err = p.waitJS(`const Q = __q.cur(); return !!Q?.querySelector('.qx-canvas') && [...Q.querySelectorAll('.qx-tab')].some((t) => t.textContent === '函数核对')`, 60*time.Second)
	}
	took := time.Since(t0)
	if err == nil {
		err = p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qx-tab')].find((t) => t.textContent === '函数核对'))`)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.reveal(__q.cur().querySelector('button[aria-label^="查找和替换"]')))`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.activeElement?.matches('.q-find-in[aria-label="查找"]')`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("71/71")
	}
	if err == nil {
		// Find looks at what cells show: only the tally cell shows 71/71.
		_, err = p.waitJS(`return __q.cur().querySelector('.q-find-n').textContent === '1/1'`, 10*time.Second)
	}
	p.key("Escape")
	s.step(fmt.Sprintf("大文件：脚本生成、公式不带结果的 2 万行表格打开时算出全部公式，用时 %.1f 秒；函数核对页 71 个难点函数与 Excel 的答案一致（71/71）", took.Seconds()), err, "")
	p.shot("web-big-xlsx.png")

	err = p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qx-tab')].find((t) => t.textContent === '明细'))`)
	if err == nil {
		err = p.goCell("F2")
	}
	if err == nil {
		err = p.typeText("7")
	}
	t0 = time.Now()
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		err = waitFile(big, "xl/worksheets/sheet1.xml", func(x string) bool {
			return regexp.MustCompile(`<c r="F2"[^>]*><v>7</v>`).MatchString(x)
		})
	}
	took = time.Since(t0)
	if err == nil {
		err = waitFile(big, "xl/worksheets/sheet3.xml", func(x string) bool {
			return regexp.MustCompile(`<c r="B8"[^>]*><f>[^<]+</f><v>0</v>`).MatchString(x) && regexp.MustCompile(`<c r="B9"[^>]*><f>[^<]+</f><v>0</v>`).MatchString(x)
		})
	}
	s.step(fmt.Sprintf("大文件：改 F2 后整本重算并写回（%.1f 秒），汇总页 SUMIFS 合计与明细累计、SUMPRODUCT 两项核对都为 0", took.Seconds()), err, "")

	t0 = time.Now()
	err = open("长文档.docx")
	if err == nil {
		_, err = p.waitJS(`const m = /共 (\d+) 页/.exec(__q.cur()?.querySelector('.qd-stxt')?.textContent ?? ''); return !!m && +m[1] >= 50`, 60*time.Second)
	}
	took = time.Since(t0)
	pages, _ := p.js(`return /共 (\d+) 页/.exec(__q.cur().querySelector('.qd-stxt').textContent)[1]`)
	s.step(fmt.Sprintf("大文件：长文档按纸张分成 %s 页，打开并排好版用时 %.1f 秒", strings.Trim(string(pages), `"`), took.Seconds()), err, "")

	t0 = time.Now()
	err = open("大演示.pptx")
	if err == nil {
		_, err = p.waitJS(`return __q.cur()?.querySelectorAll('.qp-thumb').length === 120`, 60*time.Second)
	}
	took = time.Since(t0)
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.cur().querySelectorAll('.qp-thumb')[119])`)
	}
	if err == nil {
		_, err = p.waitJS(`return __q.cur().querySelector('.qp-stxt').textContent.includes('120 / 120')`, 10*time.Second)
	}
	s.step(fmt.Sprintf("大文件：120 页演示文稿打开用时 %.1f 秒，缩略图栏翻到最后一页", took.Seconds()), err, "")
	p.shot("web-big-pptx.png")

	// --- pins: notes on spots in the files, as cards above the chat composer
	err = p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin'))`)
	if err == nil {
		_, err = p.waitJS(`return __q.cur().matches('.q-pinning') && !!__q.cur().querySelector('.q-pintip') && __q.cur().querySelector('.q-btn.q-pin').matches('.arm')`, 5*time.Second)
	}
	if err == nil {
		err = p.dragJS(`const r = __q.cur().querySelector('.qp-stage .slide [data-k="2"]').getBoundingClientRect(); return { x1: r.left - 6, y1: r.top - 6, x2: r.right + 6, y2: r.bottom + 6 }`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.activeElement?.matches('.q-pin-in') && !__q.cur().matches('.q-pinning') && !!__q.cur().querySelector('.q-pinmarq')`, 5*time.Second)
	}
	if err == nil {
		// The box's picture shows in the popover and goes with the pin.
		_, err = p.waitJS(`return __q.cur().querySelector('.q-pinpop .q-pin-shot')?.src.startsWith('data:image/jpeg')`, 10*time.Second)
	}
	if err == nil {
		p.shot("web-pin-box.png")
	}
	if err == nil {
		err = p.typeText("这页要点太笼统，补上数字")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		_, err = p.waitJS(`const l = __q.pins(); return l.length === 1 && l[0].startsWith('大演示.pptx · 第 120 页 [s120.2]「') && l[0].endsWith('：这页要点太笼统，补上数字') && !__q.draft().includes('大演示') && document.querySelector('.qpd-pill')?.textContent.startsWith('1 条标注') && document.querySelector('.qpd-list').hidden && document.querySelector('.qpd-card .qpd-thumb')?.src.startsWith('data:image/') && __q.cur().querySelectorAll('.qp-pin').length === 1 && __q.cur().querySelector('.q-pin-n').textContent === '1' && !!__q.cur().querySelectorAll('.qp-thumb')[119].matches('.q-pinned')`, 10*time.Second)
	}
	s.step("标注：点工具栏的「标注」进入框选，在第 120 页拖框框住文本框后写下问题，弹窗里带上框选区域的截图；回车后聊天框上方只多一个「1 条标注」胶囊，卡片「大演示.pptx · 第 120 页 [s120.2]「原文」：问题」连同截图收在胶囊里，聊天框文字不变；形状描上虚线框、缩略图带圆点", err, "")

	err = open("长文档.docx")
	if err == nil {
		_, err = p.waitJS(`const Q = __q.cur(); return !!Q?.querySelector('.qd-nav:not([hidden])') && Q.querySelectorAll('.qd-nav-i').length > 100`, 30*time.Second)
	}
	if err == nil {
		err = p.clickJS(`const it = [...__q.cur().querySelectorAll('.qd-nav-i')].find((x) => x.textContent === '第 30 章 区域与产品分析'); it?.scrollIntoView({ block: 'center' }); return __q.centre(it)`)
	}
	if err == nil {
		_, err = p.waitJS(`const Q = __q.cur(); const m = /第 (\d+) 页/.exec(Q.querySelector('.qd-stxt').textContent); return Q.querySelector('.qd-nav-i.on')?.textContent === '第 30 章 区域与产品分析' && +m[1] > 50`, 5*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qd-nav-i')].find((x) => x.textContent === '第 30 章 区域与产品分析').querySelector('.qd-nav-tw'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return __q.cur().querySelectorAll('.qd-nav-i[hidden]').length === 3`, 5*time.Second)
	}
	navPage, _ := p.js(`return /第 (\d+) 页/.exec(__q.cur().querySelector('.qd-stxt').textContent)[1]`)
	s.step(fmt.Sprintf("文档：导航窗格列出长文档的 161 个标题，点「第 30 章」跳到第 %s 页并高亮这一章，折叠后它下面的 3 个小节收起", strings.Trim(string(navPage), `"`)), err, "")
	p.shot("web-docx-nav.png")

	// Box the words in the chapter's first paragraph that holds them.
	_, err = p.js(`const Q = __q.cur(); const h = [...Q.querySelectorAll('.qd-page > .qd-p')].find((x) => x.textContent === '第 30 章 区域与产品分析'); let p = h.nextElementSibling; while (p && (p.dataset.h || !p.textContent.includes('毛利率'))) p = p.nextElementSibling; p.scrollIntoView({ block: 'center' }); Q.querySelector('.qd-page').focus({ preventScroll: true }); window.__words = p; return true`)
	if err == nil {
		err = p.key("m", "mod", "shift")
	}
	if err == nil {
		_, err = p.waitJS(`return __q.cur().matches('.q-pinning')`, 5*time.Second)
	}
	if err == nil {
		// A box from the first heading line down into the next paragraph names both blocks.
		err = p.dragJS(`const h = [...__q.cur().querySelectorAll('.qd-page > .qd-p')].find((x) => x.textContent === '第 30 章 区域与产品分析'); const a = h.getBoundingClientRect(); const b = h.nextElementSibling.getBoundingClientRect(); return a.top > 0 && b.bottom < innerHeight ? { x1: a.left + 2, y1: a.top + 2, x2: b.left + 80, y2: b.top + 8 } : null`)
	}
	var span string
	if err == nil {
		var v json.RawMessage
		v, err = p.waitJS(`const t = __q.cur().querySelector('.q-pinpop .q-pin-h b')?.textContent; return /^\[p\d+\]–\[p\d+\]$/.test(t ?? '') && t`, 5*time.Second)
		_ = json.Unmarshal(v, &span)
	}
	if err == nil {
		err = p.key("Escape")
	}
	if err == nil {
		_, err = p.waitJS(`return !__q.cur().querySelector('.q-pinpop') && !__q.cur().querySelector('.q-pinmarq')`, 5*time.Second)
	}
	s.step(fmt.Sprintf("标注：⇧⌘M 进入框选，框住标题和下一段，标注写成首尾两个段落编号 %s；Esc 关掉不留痕迹", span), err, "")

	if err == nil {
		err = p.key("m", "mod", "shift")
	}
	if err == nil {
		err = p.dragJS(`const p = window.__words; const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode()) && !n.data.includes('毛利率')); const i = n.data.indexOf('毛利率'); const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 7); const b = r.getBoundingClientRect(); return { x1: b.left + 1, y1: b.top + 2, x2: b.right - 1, y2: b.bottom - 2 }`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.activeElement?.matches('.q-pin-in') && /^\[p\d+\]$/.test(__q.cur().querySelector('.q-pin-h b').textContent)`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("写出具体百分点")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		_, err = p.waitJS(`const l = __q.pins(); return l.length === 2 && /^长文档\.docx · \[p\d+\]「毛利率同比上升」：写出具体百分点$/.test(l[1]) && document.querySelectorAll('.qpd-card')[1].querySelector('.qpd-thumb')?.src.startsWith('data:image/') && [...(CSS.highlights.get('q-pin') ?? [])].some((r) => r.toString() === '毛利率同比上升')`, 10*time.Second)
	}
	s.step("标注：在文档里框住「毛利率同比上升」再标注，胶囊里累积到 2 条，第二条带 office_read 的段落编号 [pN] 和这段文字的截图，原文在文档里标出", err, "")

	err = open("大表格.xlsx")
	if err == nil {
		_, err = p.waitJS(`return !!__q.cur()?.querySelector('.qx-canvas')`, 30*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.qx-tab')].find((t) => t.textContent === '汇总'))`)
	}
	// The cell box sits on the active cell: read where B3 and D5 are.
	var b3, d5 json.RawMessage
	cellAt := `const r = __q.cur().querySelector('.qx-input').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }`
	if err == nil {
		if err = p.goCell("B3"); err == nil {
			b3, err = p.js(cellAt)
		}
	}
	if err == nil {
		if err = p.goCell("D5"); err == nil {
			d5, err = p.js(cellAt)
		}
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin'))`)
	}
	if err == nil {
		// From the middle of B3 to the middle of D5.
		err = p.dragJS(fmt.Sprintf(`const a = %s; const z = %s; return { x1: a.x, y1: a.y, x2: z.x, y2: z.y }`, b3, d5))
	}
	var cells string
	if err == nil {
		var v json.RawMessage
		v, err = p.waitJS(`const t = __q.cur().querySelector('.q-pinpop .q-pin-h b')?.textContent; return /^汇总!B3:D5$/.test(t ?? '') && __q.cur().querySelector('.q-pinpop .q-pin-q')?.textContent && t`, 5*time.Second)
		_ = json.Unmarshal(v, &cells)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre([...__q.cur().querySelectorAll('.q-pinpop .q-btn')].find((b) => b.textContent === '取消'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return !__q.cur().querySelector('.q-pinpop')`, 5*time.Second)
	}
	s.step(fmt.Sprintf("标注：表格里点「标注」拖框，框到的单元格选中并写成 %s，带上框内的值；取消后不加卡片", cells), err, "")

	if err == nil {
		err = p.goCell("N2")
	}
	var at struct{ X, Y float64 }
	if err == nil {
		var v json.RawMessage
		if v, err = p.js(`const r = __q.cur().querySelector('.qx-input').getBoundingClientRect(); return { x: r.left + 10, y: r.top + 8 }`); err == nil {
			err = json.Unmarshal(v, &at)
		}
	}
	if err == nil {
		err = p.rclick(at.X, at.Y)
	}
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.startsWith('标注问题')); return b && __q.centre(b)`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.activeElement?.matches('.q-pin-in')`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("华东合计和上月对不上")
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(__q.cur().querySelector('.q-pin-ok'))`)
	}
	if err == nil {
		_, err = p.waitJS(`const l = __q.pins(); return l.length === 3 && /^大表格\.xlsx · 汇总!N2「[\d,.]+」（公式 =SUM\(B2:M2\)）：华东合计和上月对不上$/.test(l[2]) && document.querySelector('.qpd-pill').textContent.startsWith('3 条标注') && document.querySelector('.qpd-list').hidden`, 10*time.Second)
	}
	s.step("标注：表格单元格右键「标注问题…」，第三条写明 汇总!N2、显示的值和公式；三个文件的标注都收在聊天框上方的「3 条标注」胶囊里，不占地方", err, "")
	p.shot("web-pins.png")

	err = p.clickJS(`return __q.centre(__q.cur().querySelector('.q-btn.q-pin + .q-more'))`)
	if err == nil {
		err = p.clickJS(`const b = [...document.querySelectorAll('.q-menu button')].find((b) => b.textContent.startsWith('移除这个文件的标注')); return b && __q.centre(b)`)
	}
	if err == nil {
		_, err = p.waitJS(`return __q.pins().length === 2 && __q.cur().querySelector('.q-pin-n').hidden`, 10*time.Second)
	}
	s.step("标注：从标注菜单移除这个表格的标注，聊天框上方剩下另外两张卡片，表格上的标记随之消失", err, "")

	// Open the pill, edit the first card's note, drop the second card, then send.
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('.qpd-tog'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return !document.querySelector('.qpd-list').hidden && document.querySelector('.qpd-tog').getAttribute('aria-expanded') === 'true'`, 5*time.Second)
	}
	if err == nil {
		p.shot("web-pins-open.png")
		err = p.clickJS(`return __q.centre(document.querySelector('.qpd-card .qpd-btn[aria-label="修改说明"]'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return document.activeElement?.matches('.qpd-in')`, 5*time.Second)
	}
	if err == nil {
		_, err = p.waitJS(`const n = document.activeElement; return n.selectionStart === 0 && n.selectionEnd === n.value.length && n.value.length > 0`, 5*time.Second)
	}
	if err == nil {
		err = p.typeText("要点改成三条，每条带数字")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		_, err = p.waitJS(`const l = __q.pins(); return l.length === 2 && /：要点改成三条，每条带数字$/.test(l[0]) && !l[0].includes('笼统') && !document.querySelector('.qpd-in')`, 10*time.Second)
	}
	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelectorAll('.qpd-card')[1]?.querySelector('.qpd-btn[aria-label="移除这条标注"]'))`)
	}
	if err == nil {
		_, err = p.waitJS(`return __q.pins().length === 1 && document.querySelector('.qpd-pill').textContent.startsWith('1 条标注')`, 10*time.Second)
	}
	s.step("标注卡片：点胶囊展开列表，每条带截图；点 ✎ 改写第一张卡片的说明，回车保存；点第二张卡片的 × 移除它，胶囊变成「1 条标注」", err, "")

	if err == nil {
		err = p.clickJS(`return __q.centre(document.querySelector('[data-lexical-editor="true"]'))`)
	}
	if err == nil {
		err = p.typeText("按标注改一下")
	}
	if err == nil {
		err = p.key("Enter")
	}
	if err == nil {
		_, err = p.waitJS(`return !document.querySelector('.qpd') && !__q.cur().querySelector('.qp-pin')`, 30*time.Second)
	}
	var seen string
	var images int
	if err == nil {
		for end := time.Now().Add(30 * time.Second); time.Now().Before(end); time.Sleep(300 * time.Millisecond) {
			if seen, images = model.lastPrompt(); strings.Contains(seen, "按标注改一下") {
				break
			}
		}
		want := "1. 大演示.pptx · 第 120 页 [s120.2]「"
		if !strings.Contains(seen, "按标注改一下") || !strings.Contains(seen, want) || !strings.Contains(seen, "：要点改成三条，每条带数字") || strings.Contains(seen, "写出具体百分点") {
			err = fmt.Errorf("模型收到的内容不对：%q", seen)
		} else if images != 1 || !strings.Contains(seen, "（下面是这处的截图）") {
			err = fmt.Errorf("模型应收到 1 张标注截图，实际 %d 张：%q", images, seen)
		}
	}
	var sent string
	for _, l := range strings.Split(seen, "\n") {
		if strings.HasPrefix(l, "1. 大演示.pptx") {
			sent = l
		}
	}
	s.step(fmt.Sprintf("发送：在聊天框输入「按标注改一下」回车，卡片和编辑器里的标记随即清空；模型在同一轮收到这句话、附带的标注 %q 和 %d 张框选截图，删掉的那条没有发出", sent, images), err, "")
	p.shot("web-pins-sent.png")

	for _, f := range []string{"预算.xlsx", "季度报告.docx", "经营回顾.pptx"} {
		keep(filepath.Join(ws, f), "web-"+f)
	}
	if out, verr := verifyWithPython(ws); verr == nil && out != "" {
		s.step("不经本插件的读者打开改后的文件（openpyxl / python-docx）", nil, out)
	}
}

// startWeb runs dsh web in the workspace on a free port and returns its
// signed-in address; the address carries a token, so it is never logged.
func startWeb(t *testing.T, d *dshEnv, ws string, model *fakeModel) (string, error) {
	argv := d.argv("--profile", "quire-web", "--no-open", "--port", "0")
	cmd := exec.Command(argv[0], argv[1:]...)
	cmd.Dir = ws
	cmd.Env = append(os.Environ(), d.env...)
	cmd.Env = append(cmd.Env, "DEEPSEEK_API_KEY=sk-quire-e2e", "DEEPSEEK_BASE_URL="+model.URL)
	out, _ := cmd.StdoutPipe()
	cmd.Stderr = cmd.Stdout
	if err := cmd.Start(); err != nil {
		return "", err
	}
	t.Cleanup(func() { cmd.Process.Kill(); cmd.Wait() })
	found := make(chan string, 1)
	go func() {
		sc := bufio.NewScanner(out)
		sc.Buffer(make([]byte, 1<<20), 1<<20)
		re := regexp.MustCompile(`dsh web: (http://\S+)`)
		for sc.Scan() {
			if m := re.FindStringSubmatch(sc.Text()); m != nil {
				select {
				case found <- m[1]:
				default:
				}
			}
		}
	}()
	select {
	case u := <-found:
		return u, nil
	case <-time.After(90 * time.Second):
		return "", fmt.Errorf("dsh web 90 秒内没有打印地址")
	}
}

// page drives the dsh web page with real input events.
type page struct {
	c   *chrome
	ctx context.Context
}

func (p *page) js(code string) (json.RawMessage, error) {
	return p.c.eval(p.ctx, pageJS+"\n"+code)
}

func (p *page) waitJS(code string, limit time.Duration) (json.RawMessage, error) {
	deadline := time.Now().Add(limit)
	var last error
	for time.Now().Before(deadline) {
		v, err := p.js(code)
		if err == nil && string(v) != "false" && string(v) != "null" && len(v) > 0 {
			return v, nil
		}
		last = err
		time.Sleep(200 * time.Millisecond)
	}
	if last != nil {
		return nil, last
	}
	return nil, fmt.Errorf("等待超时：%s", clip(code, 80))
}

func (p *page) click(x, y float64, count int) error {
	for i := 1; i <= count; i++ {
		for _, typ := range []string{"mousePressed", "mouseReleased"} {
			if err := p.c.send(p.ctx, p.c.session, "Input.dispatchMouseEvent", map[string]any{"type": typ, "x": x, "y": y, "button": "left", "clickCount": i}, nil); err != nil {
				return err
			}
		}
	}
	time.Sleep(300 * time.Millisecond)
	return nil
}

// clickJS clicks the centre the code returns, once or count times.
func (p *page) rclick(x, y float64) error {
	for _, typ := range []string{"mousePressed", "mouseReleased"} {
		if err := p.c.send(p.ctx, p.c.session, "Input.dispatchMouseEvent", map[string]any{"type": typ, "x": x, "y": y, "button": "right", "clickCount": 1}, nil); err != nil {
			return err
		}
	}
	time.Sleep(300 * time.Millisecond)
	return nil
}

func (p *page) clickJS(code string, count ...int) error {
	v, err := p.waitJS(code, 15*time.Second)
	if err != nil {
		return err
	}
	var at struct{ X, Y float64 }
	if err := json.Unmarshal(v, &at); err != nil {
		return err
	}
	n := 1
	if len(count) > 0 {
		n = count[0]
	}
	return p.click(at.X, at.Y, n)
}

func (p *page) clickText(text string, right bool) error {
	return p.clickJS(fmt.Sprintf(`return __q.centre(__q.byText(%s, %v))`, jsStr(text), right))
}

// shiftClick clicks with Shift held, to add to a selection.
func (p *page) shiftClick(x, y float64) error {
	for _, typ := range []string{"mousePressed", "mouseReleased"} {
		if err := p.c.send(p.ctx, p.c.session, "Input.dispatchMouseEvent", map[string]any{"type": typ, "x": x, "y": y, "button": "left", "clickCount": 1, "modifiers": 8}, nil); err != nil {
			return err
		}
	}
	time.Sleep(300 * time.Millisecond)
	return nil
}

func (p *page) drag(x1, y1, x2, y2 float64) error {
	steps := []struct {
		typ  string
		x, y float64
	}{{"mousePressed", x1, y1}, {"mouseMoved", (x1 + x2) / 2, (y1 + y2) / 2}, {"mouseMoved", x2, y2}, {"mouseReleased", x2, y2}}
	for _, s := range steps {
		params := map[string]any{"type": s.typ, "x": s.x, "y": s.y, "button": "left", "clickCount": 1}
		if s.typ == "mouseMoved" {
			params["buttons"] = 1
		}
		if err := p.c.send(p.ctx, p.c.session, "Input.dispatchMouseEvent", params, nil); err != nil {
			return err
		}
		time.Sleep(60 * time.Millisecond)
	}
	time.Sleep(300 * time.Millisecond)
	return nil
}

// dragJS drags over the box a script returns as {x1, y1, x2, y2}.
func (p *page) dragJS(code string) error {
	v, err := p.waitJS(code, 15*time.Second)
	if err != nil {
		return err
	}
	var b struct{ X1, Y1, X2, Y2 float64 }
	if err := json.Unmarshal(v, &b); err != nil {
		return err
	}
	return p.drag(b.X1, b.Y1, b.X2, b.Y2)
}

func (p *page) typeText(s string) error {
	err := p.c.send(p.ctx, p.c.session, "Input.insertText", map[string]any{"text": s}, nil)
	time.Sleep(200 * time.Millisecond)
	return err
}

var keyCodes = map[string]struct {
	code string
	vk   int
	text string
}{"Enter": {"Enter", 13, "\r"}, "Tab": {"Tab", 9, ""}, "Escape": {"Escape", 27, ""}, "a": {"KeyA", 65, ""}, "b": {"KeyB", 66, ""}, "d": {"KeyD", 68, ""}, "m": {"KeyM", 77, ""}, "ArrowRight": {"ArrowRight", 39, ""}, "Delete": {"Delete", 46, ""}}

// key presses a key, with "mod" meaning ⌘ on macOS and Ctrl elsewhere.
func (p *page) key(name string, mods ...string) error {
	k := keyCodes[name]
	m := 0
	for _, mod := range mods {
		switch {
		case mod == "shift":
			m |= 8
		case runtime.GOOS == "darwin":
			m |= 4
		default:
			m |= 2
		}
	}
	down := map[string]any{"type": "keyDown", "key": name, "code": k.code, "windowsVirtualKeyCode": k.vk, "modifiers": m}
	if k.text != "" && m == 0 {
		down["text"] = k.text
	}
	if err := p.c.send(p.ctx, p.c.session, "Input.dispatchKeyEvent", down, nil); err != nil {
		return err
	}
	err := p.c.send(p.ctx, p.c.session, "Input.dispatchKeyEvent", map[string]any{"type": "keyUp", "key": name, "code": k.code, "windowsVirtualKeyCode": k.vk, "modifiers": m}, nil)
	time.Sleep(250 * time.Millisecond)
	return err
}

// goCell selects a cell or range through the name box, then gives the
// grid the keyboard.
func (p *page) goCell(ref string) error {
	if err := p.clickJS(`return __q.centre(__q.cur().querySelector('.qx-name'))`); err != nil {
		return err
	}
	if _, err := p.js(`const n = __q.cur().querySelector('.qx-name'); n.select(); return true`); err != nil {
		return err
	}
	if err := p.typeText(ref); err != nil {
		return err
	}
	return p.key("Enter")
}

func (p *page) shot(name string) {
	if png, err := p.c.capture(p.ctx); err == nil {
		os.MkdirAll("report", 0o755)
		os.WriteFile(filepath.Join("report", name), png, 0o644)
	}
}

func jsStr(s string) string { b, _ := json.Marshal(s); return string(b) }

func writeTemp(dir, name, body string) string {
	p := filepath.Join(dir, name)
	os.WriteFile(p, []byte(body), 0o644)
	return p
}

func zipPart(file, name string) (string, error) {
	zr, err := zip.OpenReader(file)
	if err != nil {
		return "", err
	}
	defer zr.Close()
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			return string(b), err
		}
	}
	return "", fmt.Errorf("%s 里没有 %s", filepath.Base(file), name)
}

// waitFile waits until the part satisfies every condition (the editor
// saves on its own, a moment after the change).
func waitFile(file, part string, conds ...func(string) bool) error {
	deadline := time.Now().Add(15 * time.Second)
	var x string
	for time.Now().Before(deadline) {
		x = ""
		if part != "" {
			x, _ = zipPart(file, part)
		}
		ok := part == "" || x != ""
		for _, c := range conds {
			ok = ok && c(x)
		}
		if ok {
			return nil
		}
		time.Sleep(300 * time.Millisecond)
	}
	return fmt.Errorf("%s 没有按预期保存（%s 片段：%s）", filepath.Base(file), part, clip(x, 160))
}

// partsLike joins every part whose name starts with prefix.
func partsLike(file, prefix string) string {
	zr, err := zip.OpenReader(file)
	if err != nil {
		return ""
	}
	defer zr.Close()
	var b strings.Builder
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, prefix) {
			x, _ := zipPart(file, f.Name)
			b.WriteString(x)
		}
	}
	return b.String()
}

func sharedOrInline(file string) string {
	a, _ := zipPart(file, "xl/sharedStrings.xml")
	b, _ := zipPart(file, "xl/worksheets/sheet1.xml")
	return a + b
}

// notesOf joins the text of every notes page.
func notesOf(file string) string {
	zr, err := zip.OpenReader(file)
	if err != nil {
		return ""
	}
	defer zr.Close()
	var b strings.Builder
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/notesSlides/notesSlide") && strings.HasSuffix(f.Name, ".xml") {
			x, _ := zipPart(file, f.Name)
			for _, m := range regexp.MustCompile(`<a:t>([^<]*)</a:t>`).FindAllStringSubmatch(x, -1) {
				b.WriteString(m[1])
			}
			b.WriteString("\n")
		}
	}
	return b.String()
}

// verifyWithPython reads the edited files with openpyxl and python-docx
// when they are installed.
func verifyWithPython(ws string) (string, error) {
	py, err := exec.LookPath("python3")
	if err != nil {
		return "", err
	}
	code := `import sys, os
out = []
try:
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(sys.argv[1], '预算.xlsx'))
    out.append('openpyxl D5=' + str(wb.active['D5'].value))
except ImportError: pass
try:
    import docx
    d = docx.Document(os.path.join(sys.argv[1], '季度报告.docx'))
    out.append('python-docx 末段=' + d.paragraphs[-1].text + '，表格 ' + d.tables[0].rows[1].cells[1].text)
except ImportError: pass
print('；'.join(out))`
	cmd := exec.Command(py, "-I", "-c", code, ws)
	b, err := cmd.Output()
	return strings.TrimSpace(string(b)), err
}
