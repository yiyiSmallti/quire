package e2e

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

const reportMD = `# 季度报告（草稿）

本季度收入 **120 万元**，比上季度增长 *12%*。

## 分区数据

| 地区 | 收入 | 增长 |
| --- | ---: | ---: |
| 华东 | 120 | 12% |
| 华北 | 98 | 8% |

- 华东区贡献最大
  - 新客户 30 家
- 成本下降 5%

![趋势图](chart.png)

> 数据截至 2026-09-30。
`

// legacyHTML becomes a .docx written by another program (macOS textutil).
const legacyHTML = `<html><head><meta charset="utf-8"></head><body>
<h1>旧文档</h1><p>第一段，由别的程序写的。</p><p>第二段需要改。</p>
<ul><li>条目一</li><li>条目二</li></ul>
<table border="1"><tr><td>甲</td><td>乙</td></tr><tr><td>丙</td><td>丁</td></tr></table>
</body></html>`

func TestVanillaDsh(t *testing.T) {
	s := rep.section(t, "原版 dsh")
	dshBin := findDsh()
	if dshBin == "" {
		s.step("找到 dsh", fmt.Errorf("PATH 上没有 dsh，也没设 QUIRE_E2E_DSH"), "")
		return
	}
	work := shortTemp(t)
	d := newDsh(dshBin, work)
	tgz, err := packPlugin(work)
	var ver, dump string
	if err == nil {
		ver, dump, err = d.install("quire-e2e", tgz)
	}
	if err == nil {
		err = check(strings.Contains(dump, "id: quire") && strings.Contains(dump, "@precinct/quire"), "quire entry missing from dump-config")
	}
	if !s.step("装进真实 dsh：sdk 模板 profile + quire 包，--dump-config 通过", err, "dsh "+ver) {
		return
	}

	ws := filepath.Join(work, "ws")
	os.MkdirAll(ws, 0o755)
	err = chart(filepath.Join(ws, "chart.png"))
	if err == nil {
		err = os.WriteFile(filepath.Join(ws, "notes.txt"), []byte("不是 Office 文件\n"), 0o644)
	}
	legacy := filepath.Join(ws, "legacy.docx")
	haveTextutil := false
	if _, lerr := exec.LookPath("textutil"); lerr == nil && err == nil {
		htmlPath := filepath.Join(work, "legacy.html")
		os.WriteFile(htmlPath, []byte(legacyHTML), 0o644)
		_, err = run(nil, work, "textutil", "-convert", "docx", "-output", legacy, htmlPath)
		haveTextutil = err == nil
	}
	legacyBefore, _ := os.ReadFile(legacy)
	if !s.step("准备工作区：图表图片、一个文本文件、一份别的程序（textutil）写的 .docx", err, fmt.Sprintf("legacy.docx %d 字节", len(legacyBefore))) {
		return
	}

	script := []toolCall{
		/* 0 */ {"docx_create", map[string]any{"file_path": "report.docx", "markdown": reportMD}},
		/* 1 */ {"office_read", map[string]any{"file_path": "report.docx"}},
		/* 2 */ {"docx_edit", map[string]any{"file_path": "report.docx", "edits": []any{
			map[string]any{"op": "replace", "find": "草稿", "replace": "终稿"},
			map[string]any{"op": "cell", "table": "t1", "row": 3, "col": 2, "text": "101"},
			map[string]any{"op": "insert", "after": "p3", "markdown": "本节数据来自**财务系统**。"},
			map[string]any{"op": "delete", "blocks": []any{"p8"}},
			map[string]any{"op": "append", "markdown": "## 结论\n\n继续保持增长。"},
		}}},
		/* 3 */ {"office_read", map[string]any{"file_path": "report.docx"}},
		/* 4 */ {"docx_create", map[string]any{"file_path": "report.docx", "markdown": "# 覆盖"}},
		/* 5 */ {"xlsx_create", map[string]any{"file_path": "budget.xlsx", "sheets": []any{
			map[string]any{"name": "预算", "header": true, "rows": []any{
				[]any{"项目", "单价", "数量", "小计"},
				[]any{"服务器", 8000, 3, "=B2*C2"},
				[]any{"带宽", 1200, 12, "=B3*C3"},
				[]any{"合计", nil, nil, "=SUM(D2:D3)"},
			}},
			map[string]any{"name": "汇总", "rows": []any{
				[]any{"总预算", "=预算!D4"},
				[]any{"平均单价", "=AVERAGE(预算!B2:B3)"},
			}},
		}}},
		/* 6 */ {"xlsx_edit", map[string]any{"file_path": "budget.xlsx", "edits": []any{
			map[string]any{"op": "set", "sheet": "预算", "cell": "C2", "value": 5},
			map[string]any{"op": "style", "sheet": "预算", "range": "B2:D4", "format": "#,##0"},
			map[string]any{"op": "add_sheet", "name": "备注", "rows": []any{[]any{"说明"}, []any{"由 Agent 生成"}}},
			map[string]any{"op": "rename_sheet", "sheet": "预算", "name": "明细"},
		}}},
		/* 7 */ {"office_read", map[string]any{"file_path": "budget.xlsx", "sheet": "汇总"}},
		/* 8 */ {"pptx_create", map[string]any{"file_path": "deck.pptx", "slides": []any{
			map[string]any{"title": "季度经营回顾", "subtitle": "战略部"},
			map[string]any{"title": "核心要点", "bullets": []any{"收入增长 12%", map[string]any{"text": "华东区贡献最大", "level": 1}}, "notes": "强调华东"},
			map[string]any{"title": "分区数据", "table": map[string]any{"rows": []any{[]any{"地区", "收入"}, []any{"华东", "120"}}}},
			map[string]any{"title": "趋势", "image": map[string]any{"path": "chart.png"}},
		}, "theme": map[string]any{"accent": "C00000"}}},
		/* 9 */ {"pptx_edit", map[string]any{"file_path": "deck.pptx", "edits": []any{
			map[string]any{"op": "replace", "find": "12%", "replace": "13%", "all": true},
			map[string]any{"op": "set_text", "shape": "s2.2", "text": "收入增长 13%\n  华东区贡献最大\n成本下降 5%"},
			map[string]any{"op": "notes", "slide": 1, "text": "开场白"},
			map[string]any{"op": "move_slide", "slide": 4, "to": 2},
			map[string]any{"op": "delete_slide", "slide": 3},
			map[string]any{"op": "add_slide", "slide": map[string]any{"title": "谢谢", "subtitle": "Q&A"}},
		}}},
		/* 10 */ {"office_read", map[string]any{"file_path": "deck.pptx"}},
		/* 11 */ {"docx_edit", map[string]any{"file_path": "legacy.docx", "edits": []any{
			map[string]any{"op": "replace", "find": "需要改", "replace": "已改好"},
		}}},
		/* 12 */ {"office_read", map[string]any{"file_path": "legacy.docx"}},
		/* 13 */ {"office_preview", map[string]any{"file_path": "report.docx"}},
		/* 14 */ {"xlsx_create", map[string]any{"file_path": "../outside.xlsx", "sheets": []any{map[string]any{"name": "S", "rows": []any{[]any{1}}}}}},
		/* 15 */ {"office_read", map[string]any{"file_path": "notes.txt"}},
		/* 16 */ {"docx_edit", map[string]any{"file_path": "report.docx", "edits": []any{map[string]any{"op": "delete", "blocks": []any{"p99"}}}}},
	}
	if !haveTextutil {
		script[11] = toolCall{"office_read", map[string]any{"file_path": "report.docx"}}
		script[12] = toolCall{"office_read", map[string]any{"file_path": "report.docx"}}
	}
	model := startModel(t, script)
	stderr, err := d.turn("quire-e2e", ws, model, "写一份季度报告、一张预算表和一份汇报 PPT，然后按要求修改")
	results, tools, system := model.snapshot()
	s.step("一个回合：假模型依次调用 quire 工具", err, fmt.Sprintf("模型收到 %d 个工具结果", len(results)))
	if err != nil || len(results) < len(script) {
		t.Logf("dsh stderr: %s", tail(stderr, 40))
		if err == nil {
			s.step("工具结果齐全", fmt.Errorf("只有 %d 个；dsh stderr：%s", len(results), tail(stderr, 8)), "")
		}
		return
	}
	have := quireTools(tools)
	s.step("模型看到八个 quire 工具，系统提示里有 Office 文件的用法", check(len(have) == 8 && strings.Contains(system, "zip archives: read them with office_read"),
		"tools = %v", have), strings.Join(have, "、"))

	res := results
	ok := func(i int, cond bool, name string) {
		s.step(name, check(cond && !res[i].IsError, "result = %+v", res[i]), clip(res[i].Text, 300))
	}
	fails := func(i int, cond bool, name string) {
		s.step(name, check(res[i].IsError && cond, "result = %+v", res[i]), clip(res[i].Text, 300))
	}
	ok(0, strings.Contains(res[0].Text, "Created ") && strings.Contains(res[0].Text, "report.docx"), "docx_create：Markdown 新建 Word 文档（标题、粗斜体、表格、嵌套列表、图片、引用）")
	ok(1, strings.Contains(res[1].Text, "[p1] # 季度报告（草稿）") && strings.Contains(res[1].Text, "**120 万元**") &&
		strings.Contains(res[1].Text, "[t1] table 3 rows × 3 columns") && strings.Contains(res[1].Text, "| 华北 | 98 | 8% |") &&
		strings.Contains(res[1].Text, "[p5]   - 新客户 30 家") && strings.Contains(res[1].Text, "[image: 趋势图"),
		"office_read（Word）：块编号、标题级别、粗体、表格、列表缩进、图片")
	ok(2, strings.Contains(res[2].Text, "Edited "), "docx_edit：替换、改单元格、在块后插入、删除块、追加，一次调用")
	r3 := res[3].Text
	ok(3, strings.Contains(r3, "# 季度报告（终稿）") && strings.Contains(r3, "| 华北 | 101 | 8% |") &&
		strings.Contains(r3, "本节数据来自**财务系统**。") && !strings.Contains(r3, "数据截至") && strings.Contains(r3, "## 结论") &&
		strings.Index(r3, "## 分区数据") < strings.Index(r3, "本节数据来自") && strings.Index(r3, "本节数据来自") < strings.Index(r3, "[t1]"),
		"office_read：编辑都落在对的位置（插入在 p3 之后、表格之前，引用已删）")
	fails(4, strings.Contains(res[4].Text, "already exists") && strings.Contains(res[4].Text, "overwrite: true"),
		"docx_create 不覆盖已有文件：报错并提示用 docx_edit 或 overwrite")
	ok(5, strings.Contains(res[5].Text, "Created ") && strings.Contains(res[5].Text, "2 sheet(s)"), "xlsx_create：两张工作表、表头、公式、跨表引用")
	ok(6, strings.Contains(res[6].Text, "renamed 预算 to 明细") && strings.Contains(res[6].Text, "added sheet 备注"), "xlsx_edit：改单元格、设数字格式、加工作表、改名")
	ok(7, strings.Contains(res[7].Text, "- 明细:") && strings.Contains(res[7].Text, "- 备注:") && strings.Contains(res[7].Text, "=明细!D4 → 54400") &&
		strings.Contains(res[7].Text, "=AVERAGE(明细!B2:B3) → 4600"),
		"office_read（Excel）：重算后汇总 = 8000×5 + 1200×12 = 54400，改名后公式跟着改")
	ok(8, strings.Contains(res[8].Text, "4 slide(s)"), "pptx_create：标题页、要点页（两级）、表格页、图片页、备注、主题色")
	ok(9, strings.Contains(res[9].Text, "moved slide 4 to 2") && strings.Contains(res[9].Text, "deleted slide 3"), "pptx_edit：替换、改形状文字、备注、移页、删页、加页")
	r10 := res[10].Text
	ok(10, strings.Contains(r10, "4 slides") && strings.Contains(r10, "## Slide 2 · Title and Content\n[s2.1] title: 趋势") &&
		strings.Contains(r10, "收入增长 13%") && strings.Contains(r10, "成本下降 5%") && strings.Contains(r10, "notes: 开场白") &&
		!strings.Contains(r10, "分区数据") && strings.Contains(r10, "[s4.1] ctrTitle: 谢谢"),
		"office_read（PowerPoint）：页序、改过的文字、备注都对，删掉的页不在")
	if haveTextutil {
		ok(11, strings.Contains(res[11].Text, "replaced 1"), "docx_edit 别的程序写的文档：替换一处文字")
		ok(12, strings.Contains(res[12].Text, "第二段已改好") && strings.Contains(res[12].Text, "第一段，由别的程序写的。") && strings.Contains(res[12].Text, "丁"),
			"office_read 读回：改了的地方变了，其余内容都在")
	}
	ok(13, strings.Contains(res[13].Text, "Wrote the preview"), "office_preview（原版 dsh）：预览写成 HTML 文件")
	fails(14, strings.Contains(res[14].Text, "sandbox") && strings.Contains(res[14].Text, "outside the workspace"),
		"沙箱 workspace-write：工作区外的路径拒绝写入")
	_, outsideErr := os.Stat(filepath.Join(work, "outside.xlsx"))
	s.step("工作区外没有留下文件", check(os.IsNotExist(outsideErr), "stat = %v", outsideErr), "")
	fails(15, strings.Contains(res[15].Text, "not a Word, Excel or PowerPoint file"), "office_read 非 Office 文件：明确报错")
	fails(16, strings.Contains(res[16].Text, "no block p99") && strings.Contains(res[16].Text, "office_read"), "docx_edit 编号不存在：报错并提示重新 office_read")
	report, _ := os.ReadFile(filepath.Join(ws, "report.docx"))
	s.step("失败的编辑不动文件", check(bytes.Contains(report, []byte("PK")) && zipHas(report, "word/document.xml"), "report.docx broken"), "")

	// ------------------------------------------- readers other than quire
	files := []string{"report.docx", "budget.xlsx", "deck.pptx"}
	if haveTextutil {
		files = append(files, "legacy.docx")
	}
	var bad []string
	for _, f := range files {
		if _, err := run(nil, ws, "unzip", "-tq", f); err != nil {
			bad = append(bad, f+": "+err.Error())
		}
	}
	s.step("unzip -t：四个文件都是完整的 zip", check(len(bad) == 0, "%v", bad), strings.Join(files, "、"))

	if haveTextutil {
		txt, err := run(nil, ws, "textutil", "-convert", "txt", "-stdout", "report.docx")
		if err == nil {
			err = check(strings.Contains(txt, "季度报告（终稿）") && strings.Contains(txt, "101") && strings.Contains(txt, "本节数据来自财务系统") &&
				strings.Contains(txt, "继续保持增长") && !strings.Contains(txt, "数据截至"), "text = %s", clip(txt, 400))
		}
		s.step("textutil（Apple 的 Word 读取器）读 report.docx：内容与编辑一致", err, clip(strings.Join(strings.Fields(txt), " "), 200))

		after, _ := os.ReadFile(legacy)
		changed, missing, err := zipDiff(legacyBefore, after)
		if err == nil {
			err = check(len(missing) == 0 && len(changed) == 1 && changed[0] == "word/document.xml", "changed = %v, missing = %v", changed, missing)
		}
		s.step("保真：别的程序写的 .docx 改一处文字后，只有 word/document.xml 变了，其余部件逐字节相同", err, fmt.Sprintf("变化的部件：%v", changed))
		txt, err = run(nil, ws, "textutil", "-convert", "txt", "-stdout", "legacy.docx")
		if err == nil {
			err = check(strings.Contains(txt, "第二段已改好") && strings.Contains(txt, "条目二") && strings.Contains(txt, "丁"), "text = %s", txt)
		}
		s.step("textutil 读 legacy.docx：改动生效，列表与表格都在", err, clip(strings.Join(strings.Fields(txt), " "), 120))
	}

	if py, err := exec.LookPath("python3"); err == nil {
		out, perr := run(nil, ws, py, "-I", "-c", pyCheck, filepath.Join(ws, "report.docx"), filepath.Join(ws, "budget.xlsx"))
		switch {
		case perr == nil:
			s.step("python-docx / openpyxl 读：段落与表格、工作表名、公式缓存值（data_only）", check(
				strings.Contains(out, "docx-title=季度报告（终稿）") && strings.Contains(out, "docx-cell=101") &&
					strings.Contains(out, "sheets=明细,汇总,备注") && strings.Contains(out, "total=54400") && strings.Contains(out, "formula==明细!D4"),
				"%s", out), strings.TrimSpace(out))
		case strings.Contains(out, "ModuleNotFoundError"):
			s.appendix.WriteString("\n本机 python3 没有 python-docx/openpyxl，跳过这一项核对。\n")
		default:
			s.step("python-docx / openpyxl 读", perr, "")
		}
	}

	if runtime.GOOS == "darwin" {
		ql := filepath.Join(work, "ql")
		os.MkdirAll(ql, 0o755)
		var made []string
		var qerr error
		for _, f := range []string{"report.docx", "budget.xlsx", "deck.pptx"} {
			if _, err := os.Stat(filepath.Join(ws, f)); err != nil {
				qerr = err
				continue
			}
			if _, err := run(nil, ws, "qlmanage", "-t", "-s", "900", "-o", ql, f); err != nil {
				qerr = err
				continue
			}
			png := filepath.Join(ql, f+".png")
			if st, err := os.Stat(png); err != nil || st.Size() < 2000 {
				qerr = fmt.Errorf("%s: no thumbnail", f)
				continue
			}
			keep(png, "vanilla-ql-"+f+".png")
			made = append(made, "vanilla-ql-"+f+".png")
		}
		s.step("Quick Look（macOS 的 Office 渲染器）给三个文件各出一张缩略图", qerr, strings.Join(made, "、"))
		for _, m := range made {
			fmt.Fprintf(&s.appendix, "\n![%s](%s)\n", m, m)
		}
	}

	var preview string
	if i := strings.Index(res[13].Text, " to "); i >= 0 {
		preview = strings.TrimSuffix(strings.SplitN(res[13].Text[i+4:], "; ", 2)[0], ";")
	}
	page, perr := os.ReadFile(preview)
	if perr == nil {
		perr = check(bytes.Contains(page, []byte("季度报告（终稿）")) && bytes.Contains(page, []byte("data:image/png;base64,")) &&
			!bytes.Contains(page, []byte("<script")) && bytes.Contains(page, []byte("Content-Security-Policy")), "preview content")
	}
	s.step("预览 HTML：内容是改后的文档，图片内联，没有脚本，带 CSP", perr, filepath.Base(preview))
	if perr == nil {
		keep(preview, "vanilla-preview-report.html")
		os.Remove(preview)
		if cp := chromePath(); cp != "" {
			shot := filepath.Join(work, "preview.png")
			abs, _ := filepath.Abs(filepath.Join("report", "vanilla-preview-report.html"))
			_, err := run(nil, work, cp, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--window-size=1100,1500",
				"--user-data-dir="+filepath.Join(work, "chrome-shot"), "--screenshot="+shot, "file://"+abs)
			if err == nil {
				st, serr := os.Stat(shot)
				err = check(serr == nil && st.Size() > 10000, "no screenshot: %v", serr)
			}
			if s.step("Chrome 打开预览页截图", err, "vanilla-preview-report.png") {
				keep(shot, "vanilla-preview-report.png")
				s.appendix.WriteString("\n预览页截图：\n\n![vanilla-preview-report.png](vanilla-preview-report.png)\n")
			}
		}
	}
	for _, f := range files {
		keep(filepath.Join(ws, f), "vanilla-"+f)
	}

	s.appendix.WriteString("\n### 模型收到的工具结果\n\n")
	for i, x := range res {
		fmt.Fprintf(&s.appendix, "%d. `%s`（is_error=%v）\n\n```\n%s\n```\n\n", i+1, script[i].Name, x.IsError, clip(x.Text, 1400))
	}
}

// pyCheck reads the files with python-docx and openpyxl.
const pyCheck = `
import sys, docx, openpyxl
d = docx.Document(sys.argv[1])
print("docx-title=" + d.paragraphs[0].text)
print("docx-cell=" + d.tables[0].cell(2, 1).text)
wb = openpyxl.load_workbook(sys.argv[2], data_only=True)
print("sheets=" + ",".join(wb.sheetnames))
print("total=%d" % wb["汇总"]["B1"].value)
wf = openpyxl.load_workbook(sys.argv[2])
print("formula=" + str(wf["汇总"]["B1"].value))
`

func zipHas(data []byte, name string) bool {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return false
	}
	for _, f := range zr.File {
		if f.Name == name {
			return true
		}
	}
	return false
}

// zipDiff lists the entries of b whose content differs from a, and the
// entries of a that b lacks.
func zipDiff(a, b []byte) (changed, missing []string, err error) {
	read := func(data []byte) (map[string][]byte, []string, error) {
		zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			return nil, nil, err
		}
		m := map[string][]byte{}
		var names []string
		for _, f := range zr.File {
			rc, err := f.Open()
			if err != nil {
				return nil, nil, err
			}
			m[f.Name], _ = io.ReadAll(rc)
			rc.Close()
			names = append(names, f.Name)
		}
		return m, names, nil
	}
	ma, na, err := read(a)
	if err != nil {
		return nil, nil, err
	}
	mb, nb, err := read(b)
	if err != nil {
		return nil, nil, err
	}
	for _, n := range na {
		if _, ok := mb[n]; !ok {
			missing = append(missing, n)
		}
	}
	for _, n := range nb {
		if old, ok := ma[n]; !ok || !bytes.Equal(old, mb[n]) {
			changed = append(changed, n)
		}
	}
	return changed, missing, nil
}
