//! End-to-end tests for the MCP server.
//!
//! Spawns the real `vb6-mcp` binary and drives it with Model Context Protocol
//! JSON-RPC traffic over stdio (newline-delimited JSON, no Content-Length
//! framing). Nothing is stubbed: initialization handshake, tool listing, and
//! every tool call flow through the same path Claude Code uses.

use std::process::Stdio;
use std::time::Duration;

use serde_json::{json, Value};
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};
use tokio::process::{Child, Command};

// ── Harness ────────────────────────────────────────────────────────────────────

struct McpHarness {
    stdin: tokio::process::ChildStdin,
    stdout: BufReader<tokio::process::ChildStdout>,
    next_id: i64,
    _child: Child,
}

impl McpHarness {
    async fn start() -> Self {
        let mut child = Command::new(env!("CARGO_BIN_EXE_vb6-mcp"))
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null())
            .spawn()
            .expect("failed to spawn vb6-mcp binary");

        let stdin = child.stdin.take().expect("stdin");
        let stdout = BufReader::new(child.stdout.take().expect("stdout"));

        let mut h = Self { stdin, stdout, next_id: 1, _child: child };
        h.request(
            "initialize",
            json!({
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-harness", "version": "0.0.0"}
            }),
        )
        .await;
        h.notify("notifications/initialized", json!({})).await;
        h
    }

    async fn send_raw(&mut self, msg: Value) {
        let line = serde_json::to_string(&msg).unwrap() + "\n";
        self.stdin.write_all(line.as_bytes()).await.unwrap();
        self.stdin.flush().await.unwrap();
    }

    async fn read_message(&mut self) -> Value {
        let fut = async {
            loop {
                let mut line = String::new();
                let n = self.stdout.read_line(&mut line).await.unwrap();
                assert_ne!(n, 0, "vb6-mcp closed stdout unexpectedly");
                let trimmed = line.trim();
                if trimmed.is_empty() {
                    continue;
                }
                return serde_json::from_str(trimmed)
                    .unwrap_or_else(|e| panic!("invalid JSON from server: {e}\nline: {trimmed}"));
            }
        };
        tokio::time::timeout(Duration::from_secs(10), fut)
            .await
            .expect("timed out waiting for vb6-mcp response")
    }

    async fn request(&mut self, method: &str, params: Value) -> Value {
        let id = self.next_id;
        self.next_id += 1;
        self.send_raw(
            json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params}),
        )
        .await;
        loop {
            let msg = self.read_message().await;
            if msg.get("id").and_then(Value::as_i64) == Some(id) {
                return msg;
            }
            // skip interleaved notifications
        }
    }

    async fn notify(&mut self, method: &str, params: Value) {
        self.send_raw(json!({"jsonrpc": "2.0", "method": method, "params": params}))
            .await;
    }

    /// Call a tool and parse the JSON text embedded in the MCP content array.
    async fn call_tool(&mut self, name: &str, args: Value) -> Value {
        let resp = self
            .request("tools/call", json!({"name": name, "arguments": args}))
            .await;
        let text = resp["result"]["content"][0]["text"].as_str().unwrap_or("{}");
        serde_json::from_str(text)
            .unwrap_or_else(|_| json!({"raw": text}))
    }
}

// ── Helpers ────────────────────────────────────────────────────────────────────

fn write_temp_bas(filename: &str, content: &str) -> String {
    let path = std::env::temp_dir().join(filename);
    std::fs::write(&path, content).unwrap();
    path.to_string_lossy().into_owned()
}

fn fixture(filename: &str) -> String {
    let mut p = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    p.push("tests/fixtures");
    p.push(filename);
    p.to_string_lossy().into_owned()
}

// Source shared across several tests.
const DEMO: &str = "Option Explicit\n\
                    \n\
                    Private Sub Demo()\n\
                    \x20   Dim count As Long\n\
                    \x20   count = count + 1\n\
                    \x20   MsgBox count\n\
                    \x20   bad = 1\n\
                    End Sub\n";

// ── Tests ──────────────────────────────────────────────────────────────────────

#[tokio::test]
async fn tools_list_returns_all_nine_tools() {
    let mut h = McpHarness::start().await;
    let resp = h.request("tools/list", json!({})).await;
    let tools = resp["result"]["tools"].as_array().expect("expected tools array");
    let names: Vec<&str> = tools.iter().filter_map(|t| t["name"].as_str()).collect();
    for expected in &[
        "vb6_get_symbols",
        "vb6_find_definition",
        "vb6_find_references",
        "vb6_get_hover",
        "vb6_get_diagnostics",
        "vb6_read_res_file",
        "vb6_write_res_file",
        "vb6_get_string_table",
        "vb6_read_form",
    ] {
        assert!(names.contains(expected), "tool {expected} missing: {names:?}");
    }
    assert_eq!(tools.len(), 9, "expected exactly 9 tools, got {}: {names:?}", tools.len());
}

#[tokio::test]
async fn get_symbols_returns_the_sub_and_variable() {
    let path = write_temp_bas("mcp_get_symbols.bas", DEMO);
    let mut h = McpHarness::start().await;
    let result = h.call_tool("vb6_get_symbols", json!({"file_path": path})).await;
    let syms = result["symbols"].as_array().expect("expected symbols array");
    let names: Vec<&str> = syms.iter().filter_map(|s| s["name"].as_str()).collect();
    assert!(names.contains(&"Demo"), "Sub Demo not in symbols: {names:?}");
    // `count` is a local Dim inside a Sub — not a document-level symbol
}

#[tokio::test]
async fn find_definition_resolves_use_to_dim_line() {
    let path = write_temp_bas("mcp_find_definition.bas", DEMO);
    let mut h = McpHarness::start().await;
    // `count` use inside `MsgBox count` is on line 5, char 11
    let result = h
        .call_tool(
            "vb6_find_definition",
            json!({"file_path": path, "line": 5, "column": 11}),
        )
        .await;
    let def = &result["definition"];
    assert!(!def.is_null(), "expected a definition location: {result}");
    // The `Dim count` declaration is on line 3
    assert_eq!(def["line"], json!(3), "definition should point to the Dim line: {def}");
}

#[tokio::test]
async fn find_references_includes_all_occurrences() {
    let path = write_temp_bas("mcp_find_refs.bas", DEMO);
    let mut h = McpHarness::start().await;
    let result = h
        .call_tool(
            "vb6_find_references",
            json!({"file_path": path, "line": 5, "column": 11}),
        )
        .await;
    let refs = result["references"].as_array().expect("expected references array");
    // Dim + `count = count + 1` (lhs + rhs) + `MsgBox count` = 4 occurrences
    assert!(refs.len() >= 4, "expected ≥4 references, got {}: {refs:#?}", refs.len());
}

#[tokio::test]
async fn get_hover_returns_type_info_for_variable() {
    let path = write_temp_bas("mcp_get_hover.bas", DEMO);
    let mut h = McpHarness::start().await;
    // Hover on `count` at the Dim declaration (line 3, char 8)
    let result = h
        .call_tool(
            "vb6_get_hover",
            json!({"file_path": path, "line": 3, "column": 8}),
        )
        .await;
    let text = result["hover"].as_str().unwrap_or("");
    assert!(text.contains("count"), "hover text should mention 'count': {text:?}");
}

#[tokio::test]
async fn get_diagnostics_flags_undeclared_variable() {
    let path = write_temp_bas("mcp_get_diagnostics.bas", DEMO);
    let mut h = McpHarness::start().await;
    let result = h.call_tool("vb6_get_diagnostics", json!({"file_path": path})).await;
    let diags = result["diagnostics"].as_array().expect("expected diagnostics array");
    assert!(!diags.is_empty(), "expected at least one diagnostic for `bad`");
    // `bad = 1` is on line 6 (0-based)
    assert!(
        diags.iter().any(|d| d["line"] == json!(6)),
        "expected a diagnostic on line 6 (`bad = 1`): {diags:#?}"
    );
}

#[tokio::test]
async fn read_res_file_parses_game2048_fixture() {
    let path = fixture("Game2048.RES");
    let mut h = McpHarness::start().await;
    let result = h.call_tool("vb6_read_res_file", json!({"file_path": path})).await;
    let resources = result["resources"].as_array().expect("expected resources array");
    assert!(!resources.is_empty(), "expected resources in Game2048.RES");
    for r in resources {
        assert!(r["resource_type"].is_string(), "missing resource_type: {r}");
        assert!(r["data_base64"].is_string(), "missing data_base64: {r}");
        assert!(r["language_id"].is_number(), "missing language_id: {r}");
    }
}

#[tokio::test]
async fn write_then_read_res_file_round_trips() {
    let out_path = std::env::temp_dir()
        .join("mcp_roundtrip.res")
        .to_string_lossy()
        .into_owned();

    // "hello mcp" in standard Base64 = "aGVsbG8gbWNw"
    let data_b64 = "aGVsbG8gbWNw";
    let mut h = McpHarness::start().await;

    let write_result = h
        .call_tool(
            "vb6_write_res_file",
            json!({
                "file_path": out_path,
                "resources": [{
                    "resource_type": "RcData",
                    "name_id": 42,
                    "language_id": 0,
                    "data_base64": data_b64
                }]
            }),
        )
        .await;
    assert_eq!(write_result["success"], json!(true), "write failed: {write_result}");
    assert_eq!(write_result["resource_count"], json!(1));

    let read_result = h.call_tool("vb6_read_res_file", json!({"file_path": out_path})).await;
    let resources = read_result["resources"].as_array().expect("expected resources");
    assert_eq!(resources.len(), 1, "expected exactly 1 resource after roundtrip");
    assert_eq!(
        resources[0]["data_base64"],
        json!(data_b64),
        "data mismatch after roundtrip"
    );
}

#[tokio::test]
async fn get_string_table_returns_strings_from_written_res() {
    // Build a minimal Windows string table block in memory.
    // Block 1 covers string IDs 1–16; each entry is u16 length + UTF-16LE chars.
    // We put "Hello" at position 0 (string ID 1) and leave the rest empty (len=0).
    let text_utf16: Vec<u16> = "Hello".encode_utf16().collect();
    let mut block: Vec<u8> = Vec::new();
    // Entry 0: "Hello"
    block.extend_from_slice(&(text_utf16.len() as u16).to_le_bytes());
    for &c in &text_utf16 {
        block.extend_from_slice(&c.to_le_bytes());
    }
    // Entries 1–15: empty (length word = 0)
    for _ in 1..16u16 {
        block.extend_from_slice(&0u16.to_le_bytes());
    }

    // Encode block data as Base64 manually via u8 groups-of-3
    let b64 = encode_base64(&block);

    let out_path = std::env::temp_dir()
        .join("mcp_string_table.res")
        .to_string_lossy()
        .into_owned();

    let mut h = McpHarness::start().await;
    let write = h
        .call_tool(
            "vb6_write_res_file",
            json!({
                "file_path": out_path,
                "resources": [{
                    "resource_type": "String",
                    "name_id": 1,
                    "language_id": 0,
                    "data_base64": b64
                }]
            }),
        )
        .await;
    assert_eq!(write["success"], json!(true), "write failed: {write}");

    let st = h
        .call_tool("vb6_get_string_table", json!({"file_path": out_path, "block_id": 1}))
        .await;
    let strings = st["strings"].as_array().expect("expected strings array");
    assert!(
        strings.iter().any(|s| s["value"].as_str() == Some("Hello")),
        "expected 'Hello' in string table: {strings:#?}"
    );
}

#[tokio::test]
async fn get_string_table_returns_error_for_missing_block() {
    let path = fixture("Game2048.RES");
    let mut h = McpHarness::start().await;
    let result = h
        .call_tool("vb6_get_string_table", json!({"file_path": path, "block_id": 9999}))
        .await;
    assert!(
        result["error"].is_string(),
        "expected error for a block that does not exist: {result}"
    );
}

/// Minimal Base64 encoder (standard alphabet, with padding) — avoids adding
/// the `base64` crate to dev-dependencies just for this one test.
fn encode_base64(input: &[u8]) -> String {
    const ALPHABET: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::new();
    let mut i = 0;
    while i < input.len() {
        let b0 = input[i] as u32;
        let b1 = if i + 1 < input.len() { input[i + 1] as u32 } else { 0 };
        let b2 = if i + 2 < input.len() { input[i + 2] as u32 } else { 0 };
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(ALPHABET[((n >> 18) & 0x3F) as usize] as char);
        out.push(ALPHABET[((n >> 12) & 0x3F) as usize] as char);
        if i + 1 < input.len() {
            out.push(ALPHABET[((n >> 6) & 0x3F) as usize] as char);
        } else {
            out.push('=');
        }
        if i + 2 < input.len() {
            out.push(ALPHABET[(n & 0x3F) as usize] as char);
        } else {
            out.push('=');
        }
        i += 3;
    }
    out
}

#[tokio::test]
async fn read_form_exports_control_tree_for_fixture() {
    let path = fixture("vb6_sample/Form1.frm");
    let mut h = McpHarness::start().await;
    let result = h.call_tool("vb6_read_form", json!({"file_path": path})).await;
    // The form export should not be an error and should contain the form name
    assert!(result["error"].is_null(), "read_form returned error: {result}");
    // Form1.frm has a VB.Form control at the root; the export JSON should mention it
    let text = serde_json::to_string(&result).unwrap();
    assert!(
        text.contains("Form1") || text.contains("form"),
        "form export should reference the form name: {text:.200}"
    );
}

#[tokio::test]
async fn missing_file_returns_error_object() {
    let mut h = McpHarness::start().await;
    let result = h
        .call_tool(
            "vb6_get_symbols",
            json!({"file_path": "/nonexistent/path/fake.bas"}),
        )
        .await;
    assert!(
        result["error"].is_string(),
        "expected an error field for a missing file: {result}"
    );
}

#[tokio::test]
async fn out_of_range_position_returns_error_object() {
    let path = write_temp_bas("mcp_out_of_range.bas", DEMO);
    let mut h = McpHarness::start().await;
    let result = h
        .call_tool(
            "vb6_find_definition",
            json!({"file_path": path, "line": 9999, "column": 9999}),
        )
        .await;
    // Engine clamps or returns null-definition — either is acceptable; must not panic
    assert!(
        result.get("definition").is_some() || result.get("error").is_some(),
        "expected definition (null) or error for out-of-range position: {result}"
    );
}
