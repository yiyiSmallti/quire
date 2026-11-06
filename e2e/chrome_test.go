package e2e

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sync"
	"testing"
	"time"
)

// chrome is a headless Chrome on a throwaway profile, driven over the
// DevTools protocol on a pipe (--remote-debugging-pipe: fd 3 in, fd 4 out,
// one JSON message per NUL). It opens dsh web for the editor runs.
type chrome struct {
	cmd     *exec.Cmd
	w       *os.File
	session string

	mu      sync.Mutex
	next    int
	pending map[int]chan cdpReply
	loading bool
	title   string
	history []string // navigations, for state
}

type cdpReply struct {
	Result json.RawMessage
	Err    error
}

// chromePath finds Chrome or Edge, or QUIRE_E2E_CHROME.
func chromePath() string {
	if p := os.Getenv("QUIRE_E2E_CHROME"); p != "" {
		return p
	}
	var cands []string
	switch runtime.GOOS {
	case "darwin":
		cands = []string{"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"}
	case "windows":
		cands = []string{`C:\Program Files\Google\Chrome\Application\chrome.exe`, `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`}
	default:
		for _, n := range []string{"google-chrome", "chromium", "chromium-browser"} {
			if p, err := exec.LookPath(n); err == nil {
				cands = append(cands, p)
			}
		}
	}
	for _, c := range cands {
		if _, err := os.Stat(c); err == nil {
			return c
		}
	}
	return ""
}

func startChrome(t *testing.T, bin, work string) (*chrome, error) {
	inR, inW, err := os.Pipe() // Chrome reads commands from fd 3
	if err != nil {
		return nil, err
	}
	outR, outW, err := os.Pipe() // and writes replies to fd 4
	if err != nil {
		return nil, err
	}
	c := &chrome{w: inW, pending: map[int]chan cdpReply{}}
	c.cmd = exec.Command(bin, "--headless=new", "--remote-debugging-pipe", "--no-first-run", "--no-default-browser-check",
		"--disable-extensions", "--disable-background-networking", "--disable-sync", "--window-size=1200,900",
		"--user-data-dir="+filepath.Join(work, "chrome-profile"), "about:blank")
	c.cmd.ExtraFiles = []*os.File{inR, outW}
	if err := c.cmd.Start(); err != nil {
		return nil, err
	}
	inR.Close()
	outW.Close()
	t.Cleanup(func() {
		c.cmd.Process.Kill()
		c.cmd.Wait()
	})
	go c.read(outR)

	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	var created struct{ TargetID string }
	if err := c.send(ctx, "", "Target.createTarget", map[string]any{"url": "about:blank"}, &created); err != nil {
		return nil, err
	}
	var attached struct{ SessionID string }
	if err := c.send(ctx, "", "Target.attachToTarget", map[string]any{"targetId": created.TargetID, "flatten": true}, &attached); err != nil {
		return nil, err
	}
	c.session = attached.SessionID
	if err := c.send(ctx, c.session, "Page.enable", nil, nil); err != nil {
		return nil, err
	}
	return c, nil
}

func (c *chrome) read(r *os.File) {
	br := bufio.NewReader(r)
	for {
		msg, err := br.ReadBytes(0)
		if err != nil {
			c.mu.Lock()
			for _, ch := range c.pending {
				ch <- cdpReply{Err: errors.New("chrome exited")}
			}
			c.pending = map[int]chan cdpReply{}
			c.mu.Unlock()
			return
		}
		var m struct {
			ID     int             `json:"id"`
			Method string          `json:"method"`
			Params json.RawMessage `json:"params"`
			Result json.RawMessage `json:"result"`
			Error  *struct {
				Message string `json:"message"`
			} `json:"error"`
		}
		if json.Unmarshal(msg[:len(msg)-1], &m) != nil {
			continue
		}
		c.mu.Lock()
		switch m.Method {
		case "Page.frameStartedLoading":
			c.loading = true
		case "Page.loadEventFired", "Page.frameStoppedLoading":
			c.loading = false
		}
		if ch := c.pending[m.ID]; m.ID != 0 && ch != nil {
			delete(c.pending, m.ID)
			rep := cdpReply{Result: m.Result}
			if m.Error != nil {
				rep.Err = errors.New(m.Error.Message)
			}
			ch <- rep
		}
		c.mu.Unlock()
	}
}

func (c *chrome) send(ctx context.Context, session, method string, params, out any) error {
	c.mu.Lock()
	c.next++
	id := c.next
	ch := make(chan cdpReply, 1)
	c.pending[id] = ch
	c.mu.Unlock()
	msg := map[string]any{"id": id, "method": method, "params": params}
	if params == nil {
		msg["params"] = map[string]any{}
	}
	if session != "" {
		msg["sessionId"] = session
	}
	b, _ := json.Marshal(msg)
	if _, err := c.w.Write(append(b, 0)); err != nil {
		return err
	}
	select {
	case rep := <-ch:
		if rep.Err != nil || out == nil {
			return rep.Err
		}
		return json.Unmarshal(rep.Result, out)
	case <-ctx.Done():
		return ctx.Err()
	}
}

// eval runs code as the body of an async function and returns its value
// as JSON.
func (c *chrome) eval(ctx context.Context, code string) (json.RawMessage, error) {
	var out struct {
		Result struct {
			Value json.RawMessage `json:"value"`
		} `json:"result"`
		ExceptionDetails *struct {
			Text      string `json:"text"`
			Exception struct {
				Description string `json:"description"`
			} `json:"exception"`
		} `json:"exceptionDetails"`
	}
	err := c.send(ctx, c.session, "Runtime.evaluate", map[string]any{
		"expression": "(async () => {\n" + code + "\n})()", "awaitPromise": true, "returnByValue": true, "userGesture": true,
	}, &out)
	if err != nil {
		return nil, err
	}
	if e := out.ExceptionDetails; e != nil {
		msg := e.Exception.Description
		if msg == "" {
			msg = e.Text
		}
		return nil, fmt.Errorf("javascript error: %s", msg)
	}
	return out.Result.Value, nil
}

// waitLoaded waits for the page to start and finish loading after a
// navigation.
func (c *chrome) waitLoaded(ctx context.Context) {
	deadline := time.Now().Add(15 * time.Second)
	for i := 0; time.Now().Before(deadline); i++ {
		c.mu.Lock()
		loading := c.loading
		c.mu.Unlock()
		var rs string
		if v, err := c.eval(ctx, "return document.readyState"); err == nil {
			json.Unmarshal(v, &rs)
		}
		if !loading && rs == "complete" && i >= 2 {
			return
		}
		time.Sleep(100 * time.Millisecond)
	}
}

func (c *chrome) navigate(ctx context.Context, url string) error {
	c.mu.Lock()
	c.loading = true
	c.mu.Unlock()
	return c.send(ctx, c.session, "Page.navigate", map[string]any{"url": url}, nil)
}

func (c *chrome) capture(ctx context.Context) ([]byte, error) {
	var out struct {
		Data []byte `json:"data"`
	}
	err := c.send(ctx, c.session, "Page.captureScreenshot", map[string]any{"format": "png"}, &out)
	return out.Data, err
}
