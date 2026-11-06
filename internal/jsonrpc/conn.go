// Package jsonrpc is a bidirectional JSON-RPC 2.0 endpoint over a byte
// stream, one frame per line: the framing of dsh's SDK protocol on stdio,
// which the end-to-end runs speak as the client.
// A frame with id and method is a request, id alone a response, method alone
// a notification. Lines that are not JSON are ignored.
package jsonrpc

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"strconv"
	"sync"
)

// Standard and shared error codes.
const (
	CodeInvalidParams  = -32602
	CodeMethodNotFound = -32601
	CodeInternal       = -32603
)

// Error is a JSON-RPC error object; handlers return it to choose the code.
type Error struct {
	Code    int             `json:"code"`
	Message string          `json:"message"`
	Data    json.RawMessage `json:"data,omitempty"`
}

func (e *Error) Error() string { return fmt.Sprintf("jsonrpc %d: %s", e.Code, e.Message) }

// Errorf builds an *Error.
func Errorf(code int, format string, a ...any) *Error {
	return &Error{Code: code, Message: fmt.Sprintf(format, a...)}
}

// IsCode reports whether err is a JSON-RPC error with code.
func IsCode(err error, code int) bool {
	var e *Error
	return errors.As(err, &e) && e.Code == code
}

// ErrClosed is returned for calls on a closed connection, and for calls that
// were pending when it closed.
var ErrClosed = errors.New("jsonrpc: connection closed")

// Handler serves the peer's requests and notifications. Requests run on their
// own goroutine; notifications run in order on the reading goroutine, so they
// must not block.
type Handler interface {
	Request(ctx context.Context, method string, params json.RawMessage) (any, error)
	Notify(method string, params json.RawMessage)
}

// HandlerFuncs adapts two functions to Handler; either may be nil.
type HandlerFuncs struct {
	OnRequest func(ctx context.Context, method string, params json.RawMessage) (any, error)
	OnNotify  func(method string, params json.RawMessage)
}

func (h HandlerFuncs) Request(ctx context.Context, method string, params json.RawMessage) (any, error) {
	if h.OnRequest == nil {
		return nil, Errorf(CodeMethodNotFound, "method not found: %s", method)
	}
	return h.OnRequest(ctx, method, params)
}

func (h HandlerFuncs) Notify(method string, params json.RawMessage) {
	if h.OnNotify != nil {
		h.OnNotify(method, params)
	}
}

type frame struct {
	JSONRPC string           `json:"jsonrpc"`
	ID      *json.RawMessage `json:"id,omitempty"`
	Method  string           `json:"method,omitempty"`
	Params  json.RawMessage  `json:"params,omitempty"`
	Result  json.RawMessage  `json:"result,omitempty"`
	Error   *Error           `json:"error,omitempty"`
}

// Conn is one endpoint. Start reading with Serve.
type Conn struct {
	r       io.Reader
	w       io.Writer
	handler Handler

	wmu     sync.Mutex
	mu      sync.Mutex
	next    int64
	pending map[string]chan frame
	closed  bool
	done    chan struct{}
}

// New wraps a stream. h may be nil (every request gets method-not-found).
func New(r io.Reader, w io.Writer, h Handler) *Conn {
	if h == nil {
		h = HandlerFuncs{}
	}
	return &Conn{r: r, w: w, handler: h, pending: map[string]chan frame{}, done: make(chan struct{})}
}

// Serve reads frames until the stream ends, then fails pending calls. ctx is
// passed to request handlers. It returns nil on EOF.
func (c *Conn) Serve(ctx context.Context) error {
	sc := bufio.NewScanner(c.r)
	sc.Buffer(make([]byte, 64*1024), 64<<20)
	for sc.Scan() {
		var f frame
		if json.Unmarshal(sc.Bytes(), &f) != nil {
			continue
		}
		switch {
		case f.ID != nil && f.Method != "":
			go c.answer(ctx, f)
		case f.ID != nil:
			key := string(*f.ID)
			c.mu.Lock()
			ch := c.pending[key]
			delete(c.pending, key)
			c.mu.Unlock()
			if ch != nil {
				ch <- f
			}
		case f.Method != "":
			c.handler.Notify(f.Method, f.Params)
		}
	}
	c.shut()
	return sc.Err()
}

// Done is closed once the connection stopped reading.
func (c *Conn) Done() <-chan struct{} { return c.done }

func (c *Conn) shut() {
	c.mu.Lock()
	defer c.mu.Unlock()
	if c.closed {
		return
	}
	c.closed = true
	for k, ch := range c.pending {
		close(ch)
		delete(c.pending, k)
	}
	close(c.done)
}

// abort closes the stream, when it can be closed, after a write was cut short:
// a partly written frame leaves it unusable.
func (c *Conn) abort() {
	for _, x := range []any{c.w, c.r} {
		if cl, ok := x.(io.Closer); ok {
			cl.Close()
		}
	}
}

// Then wraps a request result with a function run after the response frame
// was written, e.g. to start talking only once a handshake answer is out.
type Then struct {
	Result any
	After  func()
}

func (c *Conn) answer(ctx context.Context, f frame) {
	out := frame{ID: f.ID}
	res, err := c.handler.Request(ctx, f.Method, f.Params)
	var after func()
	if t, ok := res.(Then); ok {
		res, after = t.Result, t.After
	}
	if after != nil {
		defer after()
	}
	if err == nil {
		var b []byte
		if b, err = json.Marshal(res); err == nil {
			out.Result = b
		}
	}
	if err != nil {
		var je *Error
		if !errors.As(err, &je) {
			je = &Error{Code: CodeInternal, Message: err.Error()}
		}
		out.Error = je
	}
	c.send(out)
}

func (c *Conn) send(f frame) error {
	f.JSONRPC = "2.0"
	b, err := json.Marshal(f)
	if err != nil {
		return err
	}
	c.wmu.Lock()
	defer c.wmu.Unlock()
	_, err = c.w.Write(append(b, '\n'))
	return err
}

// Call sends a request and decodes its result into out (nil to discard).
func (c *Conn) Call(ctx context.Context, method string, params, out any) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	c.mu.Lock()
	if c.closed {
		c.mu.Unlock()
		return ErrClosed
	}
	c.next++
	key := strconv.FormatInt(c.next, 10)
	ch := make(chan frame, 1)
	c.pending[key] = ch
	c.mu.Unlock()
	drop := func() {
		c.mu.Lock()
		delete(c.pending, key)
		c.mu.Unlock()
	}

	id := json.RawMessage(key)
	f := frame{ID: &id, Method: method}
	if params != nil {
		b, err := json.Marshal(params)
		if err != nil {
			drop()
			return err
		}
		f.Params = b
	}
	// A peer that stops reading can block the write; the deadline still ends
	// the call, and the stream is closed so the stuck writer returns too.
	sent := make(chan error, 1)
	go func() { sent <- c.send(f) }()
	select {
	case err := <-sent:
		if err != nil {
			drop()
			return err
		}
	case <-ctx.Done():
		drop()
		c.abort()
		return ctx.Err()
	}
	select {
	case r, ok := <-ch:
		if !ok {
			return ErrClosed
		}
		if r.Error != nil {
			return r.Error
		}
		if out != nil && len(r.Result) > 0 {
			return json.Unmarshal(r.Result, out)
		}
		return nil
	case <-ctx.Done():
		drop()
		return ctx.Err()
	}
}

// Notify sends a notification.
func (c *Conn) Notify(method string, params any) error {
	f := frame{Method: method}
	if params != nil {
		b, err := json.Marshal(params)
		if err != nil {
			return err
		}
		f.Params = b
	}
	return c.send(f)
}
