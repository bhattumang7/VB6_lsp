//! Shared JSON-RPC harness for the end-to-end LSP tests: the real `Vb6LanguageServer` behind a
//! `tower_lsp::LspService` over an in-memory duplex, driven with framed JSON-RPC.

#![allow(dead_code, unused_imports)]

use std::time::Duration;

use serde_json::{json, Value};
use tokio::io::{
    duplex, split, AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader, ReadHalf, WriteHalf,
};
use tokio::io::DuplexStream;
use tower_lsp::{LspService, Server};
use vb6_lsp::lsp::Vb6LanguageServer;

/// A JSON-RPC client speaking to a live server over an in-memory pipe.
pub struct Harness {
    pub writer: WriteHalf<DuplexStream>,
    pub reader: BufReader<ReadHalf<DuplexStream>>,
    pub next_id: i64,
    pub _server: tokio::task::JoinHandle<()>,
}

impl Harness {
    /// Start a server and complete the LSP `initialize`/`initialized` handshake.
    /// Returns the harness plus the `initialize` result for capability checks.
    pub async fn start() -> (Self, Value) {
        Self::start_with(json!({ "capabilities": {}, "processId": null })).await
    }

    /// Like [`start`](Self::start), with the given `initialize` parameters.
    pub async fn start_with(params: Value) -> (Self, Value) {
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

        let init = h.request("initialize", params).await;
        h.notify("initialized", json!({})).await;
        let result = init["result"].clone();
        (h, result)
    }

    /// Frame and send a raw JSON value (LSP `Content-Length` framing).
    pub async fn send(&mut self, value: Value) {
        let body = serde_json::to_vec(&value).unwrap();
        let header = format!("Content-Length: {}\r\n\r\n", body.len());
        self.writer.write_all(header.as_bytes()).await.unwrap();
        self.writer.write_all(&body).await.unwrap();
        self.writer.flush().await.unwrap();
    }

    /// Read exactly one framed message, failing the test on timeout.
    pub async fn read_message(&mut self) -> Value {
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
    pub async fn request(&mut self, method: &str, params: Value) -> Value {
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
    pub async fn notify(&mut self, method: &str, params: Value) {
        self.send(json!({ "jsonrpc": "2.0", "method": method, "params": params })).await;
    }

    /// Read messages until a notification with `method` arrives; return its params.
    pub async fn wait_notification(&mut self, method: &str) -> Value {
        loop {
            let msg = self.read_message().await;
            if msg.get("id").is_none() && msg.get("method").and_then(Value::as_str) == Some(method) {
                return msg["params"].clone();
            }
        }
    }

    /// Open a document and return the diagnostics published for it.
    pub async fn open(&mut self, uri: &str, text: &str) -> Value {
        self.notify(
            "textDocument/didOpen",
            json!({ "textDocument": { "uri": uri, "languageId": "vb6", "version": 1, "text": text } }),
        )
        .await;
        self.wait_notification("textDocument/publishDiagnostics").await
    }
}

/// A scratch directory holding a project's files, removed when dropped.
pub struct TempDir {
    root: std::path::PathBuf,
}

impl TempDir {
    pub fn new(label: &str) -> Self {
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("vb6lsp-{label}-{}-{nanos}", std::process::id()));
        std::fs::create_dir_all(&root).unwrap();
        // Resolve the spelling the file system reports (short names, symlinks).
        let root = std::fs::canonicalize(&root).map(strip_verbatim).unwrap_or(root);
        TempDir { root }
    }

    pub fn path(&self, rel: &str) -> std::path::PathBuf {
        self.root.join(rel)
    }

    /// Write `content` (Windows-1252 text; ASCII in these tests) to `rel`, creating folders.
    pub fn write(&self, rel: &str, content: &str) -> std::path::PathBuf {
        let path = self.path(rel);
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, content).unwrap();
        path
    }

    pub fn uri(&self, rel: &str) -> String {
        tower_lsp::lsp_types::Url::from_file_path(self.path(rel)).unwrap().to_string()
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.root);
    }
}

/// `\?\C:\x` → `C:\x`
fn strip_verbatim(path: std::path::PathBuf) -> std::path::PathBuf {
    match path.to_string_lossy().strip_prefix(r"\\?\") {
        Some(rest) => std::path::PathBuf::from(rest),
        None => path,
    }
}

impl Harness {
    /// Read messages until a `vb6/projectStatus` notification with `state` arrives.
    pub async fn wait_status(&mut self, state: &str) -> Value {
        loop {
            let status = self.wait_notification("vb6/projectStatus").await;
            if status["state"] == json!(state) {
                return status;
            }
        }
    }

    /// Read messages until diagnostics for `uri` are published; return the diagnostics.
    pub async fn diagnostics_of(&mut self, uri: &str) -> Vec<Value> {
        loop {
            let params = self.wait_notification("textDocument/publishDiagnostics").await;
            if params["uri"] == json!(uri) {
                return params["diagnostics"].as_array().unwrap().clone();
            }
        }
    }

    /// Send a request; return its response and every notification received before it.
    pub async fn request_collect(&mut self, method: &str, params: Value) -> (Value, Vec<Value>) {
        let id = self.next_id;
        self.next_id += 1;
        self.send(json!({ "jsonrpc": "2.0", "id": id, "method": method, "params": params })).await;
        let mut notifications = Vec::new();
        loop {
            let msg = self.read_message().await;
            if msg.get("id").and_then(Value::as_i64) == Some(id) {
                return (msg, notifications);
            }
            notifications.push(msg);
        }
    }

    pub async fn execute(&mut self, command: &str, argument: Option<&str>) -> Value {
        let arguments: Vec<Value> = argument.map(|a| json!(a)).into_iter().collect();
        self.request("workspace/executeCommand", json!({ "command": command, "arguments": arguments })).await
    }

    pub async fn change_full(&mut self, uri: &str, version: i32, text: &str) {
        self.notify(
            "textDocument/didChange",
            json!({ "textDocument": { "uri": uri, "version": version }, "contentChanges": [{ "text": text }] }),
        )
        .await;
    }

    pub async fn close(&mut self, uri: &str) {
        self.notify("textDocument/didClose", json!({ "textDocument": { "uri": uri } })).await;
    }

    pub async fn definition(&mut self, uri: &str, line: u32, character: u32) -> Value {
        self.request(
            "textDocument/definition",
            json!({ "textDocument": { "uri": uri }, "position": { "line": line, "character": character } }),
        )
        .await["result"]
            .clone()
    }
}

/// Publish notifications in `messages` for `uri`, as (version, diagnostics) in arrival order.
pub fn published_for(messages: &[Value], uri: &str) -> Vec<Vec<Value>> {
    messages
        .iter()
        .filter(|m| m["method"] == json!("textDocument/publishDiagnostics") && m["params"]["uri"] == json!(uri))
        .map(|m| m["params"]["diagnostics"].as_array().unwrap().clone())
        .collect()
}

impl Harness {
    /// Read messages until a `vb6/projectStatus` with `state`; return every message read,
    /// the status last.
    pub async fn until_status(&mut self, state: &str) -> Vec<Value> {
        let mut seen = Vec::new();
        loop {
            let msg = self.read_message().await;
            let done = msg["method"] == json!("vb6/projectStatus") && msg["params"]["state"] == json!(state);
            seen.push(msg);
            if done {
                return seen;
            }
        }
    }

    /// Run `command`; return its response and all notifications that arrived while it ran,
    /// including those that trail the response (the server does not order the two).
    pub async fn execute_settled(&mut self, command: &str, argument: Option<&str>) -> (Value, Vec<Value>) {
        let arguments: Vec<Value> = argument.map(|a| json!(a)).into_iter().collect();
        let (resp, mut notes) = self
            .request_collect("workspace/executeCommand", json!({ "command": command, "arguments": arguments }))
            .await;
        let (_, trailing) = self.request_collect("workspace/executeCommand", json!({ "command": "vb6.projectInfo" })).await;
        notes.extend(trailing);
        (resp, notes)
    }
}

/// The message for a library that is neither registered nor at `tried`, as this platform words it.
pub fn missing_library_message(what: &str, tried: &str) -> String {
    if cfg!(windows) {
        format!("Type library {what} is not registered and was not found at {tried}. Its types are not available to the analysis")
    } else {
        format!("Type library {what} cannot be read on this platform. Its types are not available to the analysis")
    }
}
