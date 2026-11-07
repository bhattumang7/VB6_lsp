//! VB6 MCP Server
//!
//! Exposes VB6 analysis, resource, and form tools over the Model Context
//! Protocol using stdio transport. Run with no arguments; communicate via
//! JSON-RPC on stdin/stdout.

use base64::Engine as _;
use rmcp::{handler::server::wrapper::Parameters, schemars, tool, tool_router, ServiceExt};
use rmcp::transport::stdio;
use serde::Deserialize;
use serde_json::json;
use vb6_engine::session::{Session, Position as EngPos};
use vb6_formats::res_json::{
    find_string_table_block, parse_resource_type_name, resource_entry_to_json,
    string_entries_to_json, ResourceTypeNameError,
};
use vb6_formats::res::{
    parse_string_table, read_res_file, write_res_file, ResourceEntry, ResourceId, ResourceType,
};

// ── Parameter types ────────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, schemars::JsonSchema)]
struct FileParams {
    /// Absolute path to the VB6 source file (.bas, .cls, .frm, etc.)
    file_path: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
struct PositionParams {
    /// Absolute path to the VB6 source file
    file_path: String,
    /// Zero-based line number
    line: u32,
    /// Zero-based column (character) offset
    column: u32,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
struct ResEntry {
    /// Resource type name (e.g. "Bitmap", "Icon", "String", "RcData")
    resource_type: String,
    /// Numeric resource ID (use this or name_str)
    name_id: Option<u16>,
    /// String resource name (use this or name_id)
    name_str: Option<String>,
    /// Windows language ID
    language_id: u16,
    /// Resource data encoded as Base64
    data_base64: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
struct WriteResParams {
    /// Absolute path to write the .res file
    file_path: String,
    /// Resource entries to write
    resources: Vec<ResEntry>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
struct StringTableParams {
    /// Absolute path to the .res file
    file_path: String,
    /// String table block ID (numeric resource ID)
    block_id: u16,
}

// ── Server ─────────────────────────────────────────────────────────────────────

#[derive(Clone)]
struct Vb6McpServer;

fn load_session(file_path: &str) -> Result<(Session, usize), String> {
    let bytes = std::fs::read(file_path)
        .map_err(|e| format!("Cannot read {}: {}", file_path, e))?;
    let session = Session::from_sources(vec![(file_path.to_string(), bytes)]);
    let m = session
        .module_of(file_path)
        .ok_or_else(|| format!("Not a recognized VB6 file: {}", file_path))?;
    Ok((session, m))
}

fn pos_to_offset(session: &Session, m: usize, line: u32, column: u32) -> Option<u32> {
    let li = session.line_index(m)?;
    Some(li.offset(EngPos { line, character: column }))
}

fn parse_res_type(s: &str) -> Result<ResourceType, String> {
    parse_resource_type_name(s).map_err(|e| match e {
        ResourceTypeNameError::InvalidCustomId(_) => format!("Invalid Custom id: {}", s),
        ResourceTypeNameError::Unknown => format!("Unknown resource type: {}", s),
    })
}

#[tool_router(server_handler)]
impl Vb6McpServer {
    #[tool(description = "Get all symbols (variables, procedures, types, enums, constants) defined in a VB6 source file")]
    fn vb6_get_symbols(&self, Parameters(p): Parameters<FileParams>) -> String {
        let (session, m) = match load_session(&p.file_path) {
            Ok(v) => v,
            Err(e) => return json!({"error": e}).to_string(),
        };
        let symbols = session.document_symbols(m);
        let li = match session.line_index(m) {
            Some(li) => li,
            None => return json!({"symbols": []}).to_string(),
        };
        let result: Vec<_> = symbols
            .iter()
            .map(|s| {
                let r = li.range(s.location.span);
                json!({
                    "name": s.name,
                    "kind": format!("{:?}", s.kind),
                    "line": r.start.line,
                    "column": r.start.character,
                    "end_line": r.end.line,
                    "end_column": r.end.character,
                })
            })
            .collect();
        serde_json::to_string_pretty(&json!({"symbols": result})).unwrap_or_default()
    }

    #[tool(description = "Find the definition of a symbol at a given position in a VB6 source file")]
    fn vb6_find_definition(&self, Parameters(p): Parameters<PositionParams>) -> String {
        let (session, m) = match load_session(&p.file_path) {
            Ok(v) => v,
            Err(e) => return json!({"error": e}).to_string(),
        };
        let Some(off) = pos_to_offset(&session, m, p.line, p.column) else {
            return json!({"error": "position out of range"}).to_string();
        };
        match session.definition(m, off) {
            None => json!({"definition": null}).to_string(),
            Some(loc) => {
                let path = session.module_path(loc.module).unwrap_or_default();
                let li = session.line_index(loc.module).unwrap();
                let r = li.range(loc.span);
                json!({
                    "definition": {
                        "file": path,
                        "line": r.start.line,
                        "column": r.start.character,
                        "end_line": r.end.line,
                        "end_column": r.end.character,
                    }
                })
                .to_string()
            }
        }
    }

    #[tool(description = "Find all references to a symbol at a given position in a VB6 source file")]
    fn vb6_find_references(&self, Parameters(p): Parameters<PositionParams>) -> String {
        let (session, m) = match load_session(&p.file_path) {
            Ok(v) => v,
            Err(e) => return json!({"error": e}).to_string(),
        };
        let Some(off) = pos_to_offset(&session, m, p.line, p.column) else {
            return json!({"error": "position out of range"}).to_string();
        };
        let refs = session.references(m, off, true);
        let result: Vec<_> = refs
            .iter()
            .filter_map(|loc| {
                let path = session.module_path(loc.module)?;
                let li = session.line_index(loc.module)?;
                let r = li.range(loc.span);
                Some(json!({
                    "file": path,
                    "line": r.start.line,
                    "column": r.start.character,
                    "end_line": r.end.line,
                    "end_column": r.end.character,
                }))
            })
            .collect();
        serde_json::to_string_pretty(&json!({"references": result})).unwrap_or_default()
    }

    #[tool(description = "Get hover information for a symbol at a given position in a VB6 source file")]
    fn vb6_get_hover(&self, Parameters(p): Parameters<PositionParams>) -> String {
        let (session, m) = match load_session(&p.file_path) {
            Ok(v) => v,
            Err(e) => return json!({"error": e}).to_string(),
        };
        let Some(off) = pos_to_offset(&session, m, p.line, p.column) else {
            return json!({"error": "position out of range"}).to_string();
        };
        match session.hover(m, off) {
            None => json!({"hover": null}).to_string(),
            Some(h) => json!({"hover": h.text}).to_string(),
        }
    }

    #[tool(description = "Get all syntax and semantic diagnostics for a VB6 source file")]
    fn vb6_get_diagnostics(&self, Parameters(p): Parameters<FileParams>) -> String {
        let (session, m) = match load_session(&p.file_path) {
            Ok(v) => v,
            Err(e) => return json!({"error": e}).to_string(),
        };
        let diags = session.diagnostics(m);
        let li = match session.line_index(m) {
            Some(li) => li,
            None => return json!({"diagnostics": []}).to_string(),
        };
        let result: Vec<_> = diags
            .iter()
            .map(|d| {
                let r = li.range(d.span);
                json!({
                    "code": d.code,
                    "message": d.message().unwrap_or_else(|| "syntax error".to_string()),
                    "line": r.start.line,
                    "column": r.start.character,
                    "end_line": r.end.line,
                    "end_column": r.end.character,
                })
            })
            .collect();
        serde_json::to_string_pretty(&json!({"diagnostics": result})).unwrap_or_default()
    }

    #[tool(description = "Read a VB6 resource file (.res) and return all resource entries with their data encoded as Base64")]
    fn vb6_read_res_file(&self, Parameters(p): Parameters<FileParams>) -> String {
        match read_res_file(&p.file_path) {
            Err(e) => json!({"error": e.to_string()}).to_string(),
            Ok(resources) => {
                let result: Vec<_> = resources.iter().map(resource_entry_to_json).collect();
                serde_json::to_string_pretty(&json!({"resources": result})).unwrap_or_default()
            }
        }
    }

    #[tool(description = "Read a VB6 form file (.frm, .ctl, .pag, .dob) and return its complete design as JSON including the control tree, all properties, and companion resources")]
    fn vb6_read_form(&self, Parameters(p): Parameters<FileParams>) -> String {
        use vb6_formats::controls::export::build_form_export;
        let path = std::path::Path::new(&p.file_path);
        match build_form_export(path) {
            Err(e) => json!({"error": e.to_string()}).to_string(),
            Ok(export) => serde_json::to_string_pretty(&export).unwrap_or_default(),
        }
    }

    #[tool(description = "Write resource entries to a VB6 resource file (.res). Each entry needs a resource_type, either name_id (numeric) or name_str (string), language_id, and Base64-encoded data")]
    fn vb6_write_res_file(&self, Parameters(p): Parameters<WriteResParams>) -> String {
        let mut entries = Vec::new();
        for r in &p.resources {
            let resource_type = match parse_res_type(&r.resource_type) {
                Ok(t) => t,
                Err(e) => return json!({"error": e}).to_string(),
            };
            let name = match (r.name_id, &r.name_str) {
                (Some(id), _) => ResourceId::Id(id),
                (None, Some(n)) => ResourceId::Name(n.clone()),
                (None, None) => {
                    return json!({"error": "each resource needs name_id or name_str"}).to_string();
                }
            };
            let data = match base64::engine::general_purpose::STANDARD.decode(&r.data_base64) {
                Ok(d) => d,
                Err(e) => return json!({"error": format!("invalid base64: {}", e)}).to_string(),
            };
            entries.push(ResourceEntry::new(resource_type, name, r.language_id, data));
        }
        match write_res_file(&p.file_path, &entries) {
            Err(e) => json!({"error": e.to_string()}).to_string(),
            Ok(()) => json!({
                "success": true,
                "file": p.file_path,
                "resource_count": entries.len(),
            })
            .to_string(),
        }
    }

    #[tool(description = "Parse a string table block from a VB6 resource file (.res) and return the individual string entries")]
    fn vb6_get_string_table(&self, Parameters(p): Parameters<StringTableParams>) -> String {
        let resources = match read_res_file(&p.file_path) {
            Err(e) => return json!({"error": e.to_string()}).to_string(),
            Ok(r) => r,
        };
        let block_id = p.block_id;
        let res = find_string_table_block(&resources, block_id);
        let Some(string_resource) = res else {
            return json!({"error": format!("string table block {} not found", block_id)}).to_string();
        };
        match parse_string_table(&string_resource.data, block_id) {
            Err(e) => json!({"error": e.to_string()}).to_string(),
            Ok(strings) => {
                let result = string_entries_to_json(&strings);
                serde_json::to_string_pretty(&json!({"strings": result})).unwrap_or_default()
            }
        }
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    tracing_subscriber::fmt().with_writer(std::io::stderr).init();

    let server = Vb6McpServer.serve(stdio()).await?;
    server.waiting().await?;
    Ok(())
}
