// End-to-end runs of quire, the only test mechanism of this repository.
// They drive the real pieces and write report/quire-e2e.md, with the
// generated documents, previews and screenshots beside it:
//
//   - TestVanillaDsh: real dsh with the plugin package installed into an
//     sdk-template profile; a fake model creates, reads and edits a Word
//     document, an Excel workbook and a PowerPoint deck, edits a document
//     another program wrote, and exports a preview. Readers that do not go
//     through quire check the files: textutil, Quick Look, unzip, and
//     python-docx / openpyxl when installed.
//   - TestWebEditor: dsh web with the plugin, driven by a headless Chrome.
//
// Needs go, node and dsh (on PATH, or QUIRE_E2E_DSH: an executable or
// bin.js, run with QUIRE_E2E_NODE); the web run also Chrome or Edge.
package e2e

import (
	"archive/tar"
	"bytes"
	"compress/gzip"
	"context"
	"encoding/json"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"io"
	"io/fs"
	"net"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/yiyiSmallti/quire/internal/jsonrpc"
)

// ---------------------------------------------------------------- report

// report collects the steps of every test; TestMain writes it.
type report struct {
	mu       sync.Mutex
	sections []*section
}

type section struct {
	t        *testing.T
	title    string
	lines    []string
	fails    int
	appendix strings.Builder
}

var rep report

func (r *report) section(t *testing.T, title string) *section {
	s := &section{t: t, title: title}
	r.mu.Lock()
	r.sections = append(r.sections, s)
	r.mu.Unlock()
	return s
}

func (s *section) step(name string, err error, detail string) bool {
	mark := "✅"
	if err != nil {
		mark = "❌"
		s.fails++
		s.t.Errorf("%s: %v", name, err)
		if detail != "" {
			detail += "；"
		}
		detail += "错误：" + err.Error()
	}
	detail = strings.ReplaceAll(strings.ReplaceAll(detail, "|", "\\|"), "\n", "<br>")
	s.lines = append(s.lines, fmt.Sprintf("| %s | %s | %s |", mark, name, detail))
	return err == nil
}

func TestMain(m *testing.M) {
	code := m.Run()
	if len(rep.sections) > 0 {
		os.MkdirAll("report", 0o755)
		var b strings.Builder
		total, fails := 0, 0
		for _, s := range rep.sections {
			total += len(s.lines)
			fails += s.fails
		}
		fmt.Fprintf(&b, "# quire 端到端报告\n\n%s · %s/%s\n\n结果：%d 步，失败 %d。\n", time.Now().Format("2006-01-02 15:04:05 MST"), runtime.GOOS, runtime.GOARCH, total, fails)
		for _, s := range rep.sections {
			fmt.Fprintf(&b, "\n## %s\n\n| | 步骤 | 详情 |\n| --- | --- | --- |\n%s\n%s", s.title, strings.Join(s.lines, "\n"), s.appendix.String())
		}
		os.WriteFile(filepath.Join("report", "quire-e2e.md"), []byte(b.String()), 0o644)
	}
	os.Exit(code)
}

func check(cond bool, format string, a ...any) error {
	if cond {
		return nil
	}
	return fmt.Errorf(format, a...)
}

// keep copies a generated file into the report directory.
func keep(src, name string) {
	os.MkdirAll("report", 0o755)
	if data, err := os.ReadFile(src); err == nil {
		os.WriteFile(filepath.Join("report", name), data, 0o644)
	}
}

// ----------------------------------------------------------------- dsh

type dshEnv struct {
	bin  string
	work string
	env  []string
}

func findDsh() string {
	if p := os.Getenv("QUIRE_E2E_DSH"); p != "" {
		return p
	}
	p, _ := exec.LookPath("dsh")
	return p
}

func newDsh(bin, work string) *dshEnv {
	return &dshEnv{bin: bin, work: work, env: []string{"DSH_HOME=" + filepath.Join(work, "dsh-home")}}
}

func (d *dshEnv) argv(args ...string) []string {
	if strings.HasSuffix(d.bin, ".js") {
		node := os.Getenv("QUIRE_E2E_NODE")
		if node == "" {
			node = "node"
		}
		return append([]string{node, d.bin}, args...)
	}
	return append([]string{d.bin}, args...)
}

func (d *dshEnv) run(args ...string) (string, error) {
	cmd := d.argv(args...)
	return run(d.env, d.work, cmd[0], cmd[1:]...)
}

// install makes an sdk-template profile with the packages and returns
// dsh's version and the profile's dump-config.
func (d *dshEnv) install(profile string, tgzs ...string) (string, string, error) {
	ver, _ := d.run("--version")
	if _, err := d.run("--profile", profile, "--from-default-profile", "sdk", "--dump-config"); err != nil {
		return ver, "", err
	}
	for _, tgz := range tgzs {
		if _, err := d.run("plugin", "--profile", profile, "add", tgz); err != nil {
			return ver, "", err
		}
	}
	dump, err := d.run("--profile", profile, "--dump-config")
	return strings.TrimSpace(ver), dump, err
}

// turn runs one dsh session turn in cwd with the fake model, over the SDK
// protocol on stdio, and returns dsh's stderr.
func (d *dshEnv) turn(profile, cwd string, model *fakeModel, prompt string) (string, error) {
	argv := d.argv("--profile", profile)
	cmd := exec.Command(argv[0], argv[1:]...)
	cmd.Dir = cwd
	cmd.Env = append(os.Environ(), d.env...)
	cmd.Env = append(cmd.Env, "DEEPSEEK_API_KEY=sk-quire-e2e", "DEEPSEEK_BASE_URL="+model.URL)
	var stderr strings.Builder
	cmd.Stderr = &stderr
	stdin, _ := cmd.StdinPipe()
	stdout, _ := cmd.StdoutPipe()
	if err := cmd.Start(); err != nil {
		return "", err
	}
	var asked []string
	var askMu sync.Mutex
	conn := jsonrpc.New(stdout, stdin, jsonrpc.HandlerFuncs{
		OnRequest: func(ctx context.Context, method string, raw json.RawMessage) (any, error) {
			askMu.Lock()
			asked = append(asked, method)
			askMu.Unlock()
			return nil, jsonrpc.Errorf(jsonrpc.CodeMethodNotFound, "the e2e client has no %s", method)
		},
		OnNotify: func(string, json.RawMessage) {},
	})
	go conn.Serve(context.Background())
	defer func() {
		ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		conn.Call(ctx, "shutdown", nil, nil)
		cancel()
		done := make(chan struct{})
		go func() { cmd.Wait(); close(done) }()
		select {
		case <-done:
		case <-time.After(8 * time.Second):
			cmd.Process.Kill()
		}
	}()
	ctx, cancel := context.WithTimeout(context.Background(), 4*time.Minute)
	defer cancel()
	err := conn.Call(ctx, "initialize", map[string]any{"cwd": cwd, "provider": "deepseek-official", "model": "deepseek-flash"}, nil)
	if err == nil {
		err = conn.Call(ctx, "session/prompt", map[string]any{"sessionId": "main-quire-e2e",
			"contentBlocks": []any{map[string]string{"type": "text", "text": prompt}}}, nil)
	}
	if err == nil {
		err = conn.Call(ctx, "session/wait", map[string]any{"sessionId": "main-quire-e2e"}, nil)
	}
	askMu.Lock()
	if len(asked) > 0 {
		stderr.WriteString("\n[e2e] dsh asked the client: " + strings.Join(asked, ", "))
	}
	askMu.Unlock()
	return stderr.String(), err
}

// ----------------------------------------------------------- fake model

// fakeModel speaks the Anthropic Messages API dsh's DeepSeek adapter uses:
// it calls the scripted tools in order, one per request, then answers
// with text. It records each tool_result it receives, the tools it was
// offered and the system prompt.
type fakeModel struct {
	URL    string
	script []toolCall

	mu      sync.Mutex
	results []toolResult
	tools   []string
	system  string
	prompt  string // the text of every message in the latest request
	images  int    // image blocks in the latest request
}

type toolCall struct {
	Name  string
	Input map[string]any
}

type toolResult struct {
	Text    string
	IsError bool
}

func startModel(t *testing.T, script []toolCall) *fakeModel {
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	m := &fakeModel{URL: "http://" + ln.Addr().String(), script: script}
	srv := &http.Server{Handler: m}
	go srv.Serve(ln)
	t.Cleanup(func() { srv.Close() })
	return m
}

func (m *fakeModel) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	if strings.Contains(r.URL.Path, "/files") {
		// No Files API here: images go inline in the request.
		http.NotFound(w, r)
		return
	}
	if !strings.HasSuffix(r.URL.Path, "/messages") {
		w.Header().Set("content-type", "application/json")
		w.Write([]byte(`{"data":[]}`))
		return
	}
	body, _ := io.ReadAll(r.Body)
	var req struct {
		Model    string          `json:"model"`
		System   json.RawMessage `json:"system"`
		Messages []struct {
			Content json.RawMessage `json:"content"`
		} `json:"messages"`
		Tools []struct {
			Name string `json:"name"`
		} `json:"tools"`
	}
	json.Unmarshal(body, &req)
	var results []toolResult
	var texts []string
	images := 0
	for _, msg := range req.Messages {
		texts = append(texts, contentText(msg.Content))
		var blocks []struct {
			Type    string          `json:"type"`
			Content json.RawMessage `json:"content"`
			IsError bool            `json:"is_error"`
		}
		if json.Unmarshal(msg.Content, &blocks) != nil {
			continue
		}
		for _, b := range blocks {
			if b.Type == "image" {
				images++
			}
			if b.Type == "tool_result" {
				results = append(results, toolResult{Text: contentText(b.Content), IsError: b.IsError})
			}
		}
	}
	m.mu.Lock()
	m.results = results
	m.tools = nil
	for _, t := range req.Tools {
		m.tools = append(m.tools, t.Name)
	}
	m.system = contentText(req.System)
	m.prompt = strings.Join(texts, "\n")
	m.images = images
	m.mu.Unlock()

	w.Header().Set("content-type", "text/event-stream")
	ev := func(name string, data any) {
		b, _ := json.Marshal(data)
		fmt.Fprintf(w, "event: %s\ndata: %s\n\n", name, b)
	}
	ev("message_start", map[string]any{"type": "message_start", "message": map[string]any{
		"id": fmt.Sprintf("msg_%d", time.Now().UnixNano()), "type": "message", "role": "assistant", "model": req.Model,
		"content": []any{}, "stop_reason": nil, "usage": map[string]int{"input_tokens": 10, "output_tokens": 0},
	}})
	stop := "end_turn"
	if n := len(results); n < len(m.script) {
		input, _ := json.Marshal(m.script[n].Input)
		ev("content_block_start", map[string]any{"type": "content_block_start", "index": 0, "content_block": map[string]any{
			"type": "tool_use", "id": fmt.Sprintf("toolu_%d_%d", n, time.Now().UnixNano()), "name": m.script[n].Name, "input": map[string]any{}}})
		ev("content_block_delta", map[string]any{"type": "content_block_delta", "index": 0, "delta": map[string]any{"type": "input_json_delta", "partial_json": string(input)}})
		stop = "tool_use"
	} else {
		ev("content_block_start", map[string]any{"type": "content_block_start", "index": 0, "content_block": map[string]any{"type": "text", "text": ""}})
		ev("content_block_delta", map[string]any{"type": "content_block_delta", "index": 0, "delta": map[string]any{"type": "text_delta", "text": "done"}})
	}
	ev("content_block_stop", map[string]any{"type": "content_block_stop", "index": 0})
	ev("message_delta", map[string]any{"type": "message_delta", "delta": map[string]any{"stop_reason": stop}, "usage": map[string]int{"output_tokens": 5}})
	ev("message_stop", map[string]any{"type": "message_stop"})
}

func contentText(raw json.RawMessage) string {
	var s string
	if json.Unmarshal(raw, &s) == nil {
		return s
	}
	var blocks []struct{ Text string }
	json.Unmarshal(raw, &blocks)
	var out []string
	for _, b := range blocks {
		out = append(out, b.Text)
	}
	return strings.Join(out, "\n")
}

// lastPrompt is the text the model saw in its latest request and how
// many images came with it.
func (m *fakeModel) lastPrompt() (string, int) {
	m.mu.Lock()
	defer m.mu.Unlock()
	return m.prompt, m.images
}

func (m *fakeModel) snapshot() ([]toolResult, []string, string) {
	m.mu.Lock()
	defer m.mu.Unlock()
	return append([]toolResult(nil), m.results...), append([]string(nil), m.tools...), m.system
}

// quireTools is the tools the model was offered that quire registers.
func quireTools(tools []string) []string {
	var have []string
	for _, n := range tools {
		if strings.HasPrefix(n, "office_") || strings.HasPrefix(n, "docx_") || strings.HasPrefix(n, "xlsx_") || strings.HasPrefix(n, "pptx_") {
			have = append(have, n)
		}
	}
	return have
}

// ----------------------------------------------------------------- helpers

// packDir tars dir as an npm package (package/…), skipping skip.
func packDir(dir, out string, skip func(rel string) bool) (string, error) {
	var buf bytes.Buffer
	gz := gzip.NewWriter(&buf)
	tw := tar.NewWriter(gz)
	err := filepath.WalkDir(dir, func(path string, d fs.DirEntry, err error) error {
		if err != nil || d.IsDir() {
			return err
		}
		rel, _ := filepath.Rel(dir, path)
		if skip != nil && skip(filepath.ToSlash(rel)) {
			return nil
		}
		data, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		tw.WriteHeader(&tar.Header{Name: "package/" + filepath.ToSlash(rel), Mode: 0o644, Size: int64(len(data)), ModTime: time.Unix(499162500, 0), Typeflag: tar.TypeReg})
		_, err = tw.Write(data)
		return err
	})
	if err != nil {
		return "", err
	}
	tw.Close()
	gz.Close()
	path := filepath.Join(out, filepath.Base(dir)+".tgz")
	return path, os.WriteFile(path, buf.Bytes(), 0o644)
}

// packPlugin packs ../plugin as npm would publish it.
func packPlugin(work string) (string, error) {
	return packDir(filepath.Join("..", "plugin"), work, func(rel string) bool {
		return strings.HasSuffix(rel, ".tgz")
	})
}

func run(env []string, dir string, name string, args ...string) (string, error) {
	ctx, cancel := context.WithTimeout(context.Background(), 3*time.Minute)
	defer cancel()
	cmd := exec.CommandContext(ctx, name, args...)
	cmd.Env = append(os.Environ(), env...)
	cmd.Dir = dir
	out, err := cmd.CombinedOutput()
	if err != nil {
		return string(out), fmt.Errorf("%s %s: %v\n%s", name, strings.Join(args, " "), err, tail(string(out), 20))
	}
	return string(out), nil
}

func tail(s string, n int) string {
	lines := strings.Split(strings.TrimRight(s, "\n"), "\n")
	if len(lines) > n {
		lines = lines[len(lines)-n:]
	}
	return strings.Join(lines, "\n")
}

// shortTemp keeps socket paths under macOS's 104-byte limit.
func shortTemp(t *testing.T) string {
	base := ""
	if runtime.GOOS != "windows" {
		base = "/tmp"
	}
	dir, err := os.MkdirTemp(base, "quire-e2e-")
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { os.RemoveAll(dir) })
	return dir
}

func clip(s string, n int) string {
	if r := []rune(s); len(r) > n {
		return string(r[:n]) + "…"
	}
	return s
}

// chart writes a small bar chart PNG for the documents to embed.
func chart(path string) error {
	img := image.NewRGBA(image.Rect(0, 0, 320, 180))
	for y := 0; y < 180; y++ {
		for x := 0; x < 320; x++ {
			img.Set(x, y, color.White)
		}
	}
	bars := []struct {
		h int
		c color.RGBA
	}{{60, color.RGBA{0x44, 0x72, 0xC4, 0xff}}, {95, color.RGBA{0x44, 0x72, 0xC4, 0xff}}, {130, color.RGBA{0xED, 0x7D, 0x31, 0xff}}}
	for i, b := range bars {
		for y := 170 - b.h; y < 170; y++ {
			for x := 40 + i*90; x < 100+i*90; x++ {
				img.Set(x, y, b.c)
			}
		}
	}
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	return png.Encode(f, img)
}
