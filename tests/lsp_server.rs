//! End-to-end tests for the LSP server.
//!
//! These run the *real* `Vb6LanguageServer` behind a `tower_lsp::LspService`
//! over an in-memory duplex, and drive it with actual JSON-RPC traffic (proper
//! `Content-Length` framing, real request/response correlation, real
//! notifications). Nothing is stubbed: `initialize`, `didOpen`, incremental
//! `didChange`, navigation, formatting, rename, semantic tokens and diagnostics
//! all flow through the same path the editor uses — which is exactly where
//! integration bugs (offset conversion, document sync, capability wiring,
//! diagnostic publishing) surface.

use std::time::Duration;

use serde_json::{json, Value};
use tokio::io::{
    duplex, split, AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader, ReadHalf, WriteHalf,
};
use tokio::io::DuplexStream;
use tower_lsp::{LspService, Server};
use vb6_lsp::lsp::Vb6LanguageServer;

const URI: &str = "file:///test/demo.bas";

/// A JSON-RPC client speaking to a live server over an in-memory pipe.
struct Harness {
    writer: WriteHalf<DuplexStream>,
    reader: BufReader<ReadHalf<DuplexStream>>,
    next_id: i64,
    _server: tokio::task::JoinHandle<()>,
}

impl Harness {
    /// Start a server and complete the LSP `initialize`/`initialized` handshake.
    /// Returns the harness plus the `initialize` result for capability checks.
    async fn start() -> (Self, Value) {
        let (client_end, server_end) = duplex(256 * 1024);
        let (s_read, s_write) = split(server_end);
        let (c_read, c_write) = split(client_end);

        let (service, socket) = LspService::new(Vb6LanguageServer::new);
        let server = tokio::spawn(async move {
            Server::new(s_read, s_write, socket).serve(service).await;
        });

        let mut h = Harness {
            writer: c_write,
            reader: BufReader::new(c_read),
            next_id: 1,
            _server: server,
        };

        let init = h
            .request("initialize", json!({ "capabilities": {}, "processId": null }))
            .await;
        h.notify("initialized", json!({})).await;
        let result = init["result"].clone();
        (h, result)
    }

    /// Frame and send a raw JSON value (LSP `Content-Length` framing).
    async fn send(&mut self, value: Value) {
        let body = serde_json::to_vec(&value).unwrap();
        let header = format!("Content-Length: {}\r\n\r\n", body.len());
        self.writer.write_all(header.as_bytes()).await.unwrap();
        self.writer.write_all(&body).await.unwrap();
        self.writer.flush().await.unwrap();
    }

    /// Read exactly one framed message, failing the test on timeout.
    async fn read_message(&mut self) -> Value {
        let fut = async {
            let mut content_len = 0usize;
            loop {
                let mut line = String::new();
                let n = self.reader.read_line(&mut line).await.unwrap();
                assert_ne!(n, 0, "server closed the connection unexpectedly");
                if line == "\r\n" || line == "\n" {
                    break;
                }
                if let Some(v) = line.to_ascii_lowercase().strip_prefix("content-length:") {
                    content_len = v.trim().parse().unwrap();
                }
            }
            let mut buf = vec![0u8; content_len];
            self.reader.read_exact(&mut buf).await.unwrap();
            serde_json::from_slice::<Value>(&buf).unwrap()
        };
        tokio::time::timeout(Duration::from_secs(10), fut)
            .await
            .expect("timed out waiting for a server message")
    }

    /// Send a request and return the matching response, skipping (discarding)
    /// any interleaved notifications.
    async fn request(&mut self, method: &str, params: Value) -> Value {
        let id = self.next_id;
        self.next_id += 1;
        self.send(json!({ "jsonrpc": "2.0", "id": id, "method": method, "params": params }))
            .await;
        loop {
            let msg = self.read_message().await;
            if msg.get("id").and_then(Value::as_i64) == Some(id) {
                return msg;
            }
            // otherwise it's a notification (logMessage / publishDiagnostics) — skip
        }
    }

    /// Send a notification (no response expected).
    async fn notify(&mut self, method: &str, params: Value) {
        self.send(json!({ "jsonrpc": "2.0", "method": method, "params": params })).await;
    }

    /// Read messages until a notification with `method` arrives; return its params.
    async fn wait_notification(&mut self, method: &str) -> Value {
        loop {
            let msg = self.read_message().await;
            if msg.get("id").is_none() && msg.get("method").and_then(Value::as_str) == Some(method) {
                return msg["params"].clone();
            }
        }
    }

    /// Open a document and return the diagnostics published for it.
    async fn open(&mut self, uri: &str, text: &str) -> Value {
        self.notify(
            "textDocument/didOpen",
            json!({ "textDocument": { "uri": uri, "languageId": "vb6", "version": 1, "text": text } }),
        )
        .await;
        self.wait_notification("textDocument/publishDiagnostics").await
    }
}

const DEMO: &str = "Option Explicit\n\nPrivate Sub Demo()\n    Dim count As Long\n    count = count + 1\n    MsgBox count\n    bad = 1\nEnd Sub\n";

#[tokio::test]
async fn initialize_advertises_expected_capabilities() {
    let (_h, init) = Harness::start().await;
    let caps = &init["capabilities"];
    assert_eq!(caps["definitionProvider"], json!(true));
    assert_eq!(caps["hoverProvider"], json!(true));
    assert_eq!(caps["referencesProvider"], json!(true));
    assert_eq!(caps["documentSymbolProvider"], json!(true));
    assert_eq!(caps["documentFormattingProvider"], json!(true));
    assert_eq!(caps["renameProvider"], json!(true));
    assert!(caps["semanticTokensProvider"].is_object());
    assert_eq!(init["serverInfo"]["name"], json!("vb6-lsp"));
}

#[tokio::test]
async fn did_open_publishes_diagnostics_for_undefined_variable() {
    let (mut h, _) = Harness::start().await;
    let diags = h.open(URI, DEMO).await;
    assert_eq!(diags["uri"], json!(URI));
    let items = diags["diagnostics"].as_array().unwrap();
    assert!(!items.is_empty(), "expected at least one diagnostic for `bad = 1`");
    // `bad` is on line 6 (0-based); under Option Explicit it is undefined.
    assert!(
        items.iter().any(|d| d["range"]["start"]["line"] == json!(6)),
        "expected a diagnostic on the `bad = 1` line: {items:#?}"
    );
}

#[tokio::test]
async fn goto_definition_resolves_a_local_use() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    // `count` use inside `MsgBox count` (line 5, char 11) → its `Dim` (line 3, char 8).
    let resp = h
        .request(
            "textDocument/definition",
            json!({ "textDocument": { "uri": URI }, "position": { "line": 5, "character": 11 } }),
        )
        .await;
    let loc = &resp["result"];
    assert_eq!(loc["uri"], json!(URI), "definition resp: {resp}");
    assert_eq!(loc["range"]["start"], json!({ "line": 3, "character": 8 }));
}

#[tokio::test]
async fn hover_returns_the_declaration_signature() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    let resp = h
        .request(
            "textDocument/hover",
            json!({ "textDocument": { "uri": URI }, "position": { "line": 3, "character": 8 } }),
        )
        .await;
    let value = resp["result"]["contents"]["value"].as_str().unwrap_or("");
    assert!(value.contains("count"), "hover value: {value:?}");
}

#[tokio::test]
async fn references_include_all_uses() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    let resp = h
        .request(
            "textDocument/references",
            json!({
                "textDocument": { "uri": URI },
                "position": { "line": 5, "character": 11 },
                "context": { "includeDeclaration": true }
            }),
        )
        .await;
    let refs = resp["result"].as_array().unwrap();
    // declaration + three uses (count =, = count +, MsgBox count).
    assert!(refs.len() >= 4, "expected >= 4 references, got {}: {refs:#?}", refs.len());
}

#[tokio::test]
async fn document_symbols_list_the_procedure() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    let resp = h
        .request("textDocument/documentSymbol", json!({ "textDocument": { "uri": URI } }))
        .await;
    let syms = resp["result"].as_array().unwrap();
    assert!(
        syms.iter().any(|s| s["name"] == json!("Demo")),
        "expected `Demo` in document symbols: {syms:#?}"
    );
}

#[tokio::test]
async fn workspace_symbols_match_a_query() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    let resp = h.request("workspace/symbol", json!({ "query": "Demo" })).await;
    let syms = resp["result"].as_array().unwrap();
    assert!(syms.iter().any(|s| s["name"] == json!("Demo")), "{syms:#?}");
}

#[tokio::test]
async fn semantic_tokens_full_returns_data() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    let resp = h
        .request("textDocument/semanticTokens/full", json!({ "textDocument": { "uri": URI } }))
        .await;
    let data = resp["result"]["data"].as_array().unwrap();
    assert!(!data.is_empty(), "expected semantic-token data");
    assert_eq!(data.len() % 5, 0, "semantic tokens are 5 ints each");
}

#[tokio::test]
async fn rename_produces_edits_for_every_occurrence() {
    let (mut h, _) = Harness::start().await;
    h.open(URI, DEMO).await;
    let resp = h
        .request(
            "textDocument/rename",
            json!({
                "textDocument": { "uri": URI },
                "position": { "line": 3, "character": 8 },
                "newName": "total"
            }),
        )
        .await;
    let edits = resp["result"]["changes"][URI].as_array().unwrap();
    assert!(edits.len() >= 4, "expected >= 4 rename edits, got {}", edits.len());
    assert_eq!(edits[0]["newText"], json!("total"));
}

#[tokio::test]
async fn formatting_normalizes_keyword_case_and_indentation() {
    let (mut h, _) = Harness::start().await;
    // Lowercase keyword + missing indentation → the formatter must emit edits.
    h.open("file:///fmt.bas", "sub Main()\nDim x As Long\nend sub\n").await;
    let resp = h
        .request(
            "textDocument/formatting",
            json!({
                "textDocument": { "uri": "file:///fmt.bas" },
                "options": { "tabSize": 4, "insertSpaces": true }
            }),
        )
        .await;
    let edits = resp["result"].as_array().unwrap();
    assert!(!edits.is_empty(), "expected formatting edits for messy source");
}

#[tokio::test]
async fn incremental_did_change_is_reflected_in_later_queries() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///edit.bas";
    h.open(uri, "Public alpha As Long\n").await;

    // Sanity: the original module variable is reported.
    let before = h
        .request("textDocument/documentSymbol", json!({ "textDocument": { "uri": uri } }))
        .await;
    assert!(before["result"].as_array().unwrap().iter().any(|s| s["name"] == json!("alpha")));

    // Incremental edit: replace `alpha` (line 0, chars 7..12) with `gamma`.
    h.notify(
        "textDocument/didChange",
        json!({
            "textDocument": { "uri": uri, "version": 2 },
            "contentChanges": [{
                "range": { "start": { "line": 0, "character": 7 }, "end": { "line": 0, "character": 12 } },
                "text": "gamma"
            }]
        }),
    )
    .await;
    // didChange republishes diagnostics; drain that notification.
    let _ = h.wait_notification("textDocument/publishDiagnostics").await;

    let after = h
        .request("textDocument/documentSymbol", json!({ "textDocument": { "uri": uri } }))
        .await;
    let names: Vec<&str> =
        after["result"].as_array().unwrap().iter().filter_map(|s| s["name"].as_str()).collect();
    assert!(names.contains(&"gamma"), "edit not reflected: {names:?}");
    assert!(!names.contains(&"alpha"), "stale symbol after edit: {names:?}");
}

// ── Completion ──────────────────────────────────────────────────────────────────

const COMP_SRC: &str =
    "Sub Adder(x As Long, y As Long)\nEnd Sub\nSub Main()\n    Dim total As Long\n    total = \nEnd Sub\n";

#[tokio::test]
async fn completion_returns_items_including_locals_and_procs() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///comp.bas";
    h.open(uri, COMP_SRC).await;

    // Cursor at the end of `total = ` (inside Main, after module-level procs are visible)
    let resp = h
        .request(
            "textDocument/completion",
            json!({ "textDocument": { "uri": uri }, "position": { "line": 4, "character": 12 } }),
        )
        .await;

    let items = resp["result"].as_array().unwrap();
    let labels: Vec<&str> = items.iter().filter_map(|i| i["label"].as_str()).collect();
    assert!(!labels.is_empty(), "expected completion items, got none");
    assert!(labels.contains(&"total"), "local `total` not in completion: {labels:?}");
    assert!(labels.contains(&"Adder"), "proc `Adder` not in completion: {labels:?}");
    assert!(labels.contains(&"Main"), "proc `Main` not in completion: {labels:?}");
    // VB6 keywords must also appear
    assert!(labels.iter().any(|l| l.eq_ignore_ascii_case("Dim")), "keyword Dim missing: {labels:?}");
}

// ── Signature help ──────────────────────────────────────────────────────────────

const SIG_SRC: &str =
    "Sub Paint(x As Long, y As Long)\nEnd Sub\nSub Main()\n    Paint(\nEnd Sub\n";

#[tokio::test]
async fn signature_help_shows_first_parameter() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///sig.bas";
    h.open(uri, SIG_SRC).await;

    // Cursor inside `Paint(` argument list: line 3 is `    Paint(\n`
    // `(` is at char 9, so char 10 places the cursor just inside the arg list
    let resp = h
        .request(
            "textDocument/signatureHelp",
            json!({ "textDocument": { "uri": uri }, "position": { "line": 3, "character": 10 } }),
        )
        .await;

    let sigs = resp["result"]["signatures"].as_array().expect("expected signatures array");
    assert!(!sigs.is_empty(), "expected at least one signature");
    let label = sigs[0]["label"].as_str().unwrap_or("");
    assert!(label.contains("Paint"), "signature label should contain proc name: {label}");
    let params = sigs[0]["parameters"].as_array().expect("expected parameters");
    assert_eq!(params.len(), 2, "expected 2 params in signature");
    assert_eq!(resp["result"]["activeParameter"], json!(0), "first param should be active");
}

// ── Document highlight ──────────────────────────────────────────────────────────

const HL_SRC: &str =
    "Sub Foo()\n    Dim counter As Long\n    counter = 1\n    counter = counter + 1\nEnd Sub\n";

#[tokio::test]
async fn document_highlight_finds_all_occurrences() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///hl.bas";
    h.open(uri, HL_SRC).await;

    // Cursor on `counter` in `Dim counter` (line 1, char 8)
    let resp = h
        .request(
            "textDocument/documentHighlight",
            json!({ "textDocument": { "uri": uri }, "position": { "line": 1, "character": 8 } }),
        )
        .await;

    let highlights = resp["result"].as_array().expect("expected highlights array");
    // `counter` appears as: Dim counter, counter=1, counter=(lhs), counter (rhs) → ≥3
    assert!(
        highlights.len() >= 3,
        "expected ≥3 highlights for `counter`, got {}: {highlights:#?}",
        highlights.len()
    );
}

// ── Folding range ───────────────────────────────────────────────────────────────

const FOLD_SRC: &str =
    "Sub Alpha()\n    Dim x As Long\n    If x > 0 Then\n        x = 1\n    End If\nEnd Sub\nSub Beta()\nEnd Sub\n";

#[tokio::test]
async fn folding_range_covers_procedures_and_blocks() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///fold.bas";
    h.open(uri, FOLD_SRC).await;

    let resp = h
        .request("textDocument/foldingRange", json!({ "textDocument": { "uri": uri } }))
        .await;

    let ranges = resp["result"].as_array().expect("expected folding ranges array");
    assert!(ranges.len() >= 3, "expected ≥3 folds (Alpha, If block, Beta), got {}: {ranges:#?}", ranges.len());

    // Each range must have startLine < endLine
    for r in ranges {
        let start = r["startLine"].as_u64().unwrap();
        let end = r["endLine"].as_u64().unwrap();
        assert!(start < end, "zero-length fold: {r}");
    }

    // Alpha should cover line 0 to line 5 (inclusive)
    let alpha_fold = ranges.iter().find(|r| r["startLine"] == json!(0));
    assert!(alpha_fold.is_some(), "expected fold starting at line 0 for Alpha");
    assert_eq!(alpha_fold.unwrap()["endLine"], json!(5));
}

// ── Call hierarchy ──────────────────────────────────────────────────────────────

const HIER_SRC: &str =
    "Sub Worker()\nEnd Sub\nSub Dispatcher()\n    Worker\n    Worker\nEnd Sub\nSub Main()\n    Dispatcher\nEnd Sub\n";

#[tokio::test]
async fn call_hierarchy_prepare_and_incoming() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///hier.bas";
    h.open(uri, HIER_SRC).await;

    // Prepare on `Worker` (line 0, char 4)
    let prep = h
        .request(
            "textDocument/prepareCallHierarchy",
            json!({ "textDocument": { "uri": uri }, "position": { "line": 0, "character": 4 } }),
        )
        .await;

    let items = prep["result"].as_array().expect("expected prepareCallHierarchy items");
    assert_eq!(items.len(), 1, "expected 1 hierarchy item for Worker");
    assert_eq!(items[0]["name"], json!("Worker"));

    // Incoming calls for Worker: should list Dispatcher
    let incoming = h
        .request("callHierarchy/incomingCalls", json!({ "item": items[0] }))
        .await;

    let calls = incoming["result"].as_array().expect("expected incomingCalls array");
    assert_eq!(calls.len(), 1, "expected 1 incoming call (Dispatcher), got {}: {calls:#?}", calls.len());
    assert_eq!(calls[0]["from"]["name"], json!("Dispatcher"));
    // Dispatcher calls Worker twice → 2 call sites
    let from_ranges = calls[0]["fromRanges"].as_array().unwrap();
    assert_eq!(from_ranges.len(), 2, "expected 2 call sites in Dispatcher");
}

#[tokio::test]
async fn call_hierarchy_outgoing() {
    let (mut h, _) = Harness::start().await;
    let uri = "file:///hier2.bas";
    h.open(uri, HIER_SRC).await;

    // Prepare on `Dispatcher` (line 2, char 4)
    let prep = h
        .request(
            "textDocument/prepareCallHierarchy",
            json!({ "textDocument": { "uri": uri }, "position": { "line": 2, "character": 4 } }),
        )
        .await;

    let items = prep["result"].as_array().expect("expected items");
    assert_eq!(items[0]["name"], json!("Dispatcher"));

    // Outgoing calls: Dispatcher calls Worker twice
    let outgoing = h
        .request("callHierarchy/outgoingCalls", json!({ "item": items[0] }))
        .await;

    let calls = outgoing["result"].as_array().expect("expected outgoingCalls");
    assert_eq!(calls.len(), 1, "expected 1 unique callee (Worker), got {}: {calls:#?}", calls.len());
    assert_eq!(calls[0]["to"]["name"], json!("Worker"));
    let from_ranges = calls[0]["fromRanges"].as_array().unwrap();
    assert_eq!(from_ranges.len(), 2, "expected 2 call sites to Worker in Dispatcher");
}

#[tokio::test]
async fn requests_on_an_unopened_document_return_null() {
    let (mut h, _) = Harness::start().await;
    let resp = h
        .request(
            "textDocument/definition",
            json!({ "textDocument": { "uri": "file:///never-opened.bas" }, "position": { "line": 0, "character": 0 } }),
        )
        .await;
    assert!(resp["result"].is_null(), "expected null for an unopened document: {resp}");
}
