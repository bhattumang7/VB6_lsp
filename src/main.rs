//! VB6 Language Server - Main Entry Point
//!
//! A Language Server Protocol implementation for Visual Basic 6
//! with Claude AI integration for intelligent code assistance.

use std::env;
use base64::Engine as _;
use tower_lsp::{LspService, Server};
use tracing_subscriber::{layer::SubscriberExt, util::SubscriberInitExt};

#[cfg(windows)]
mod com_decoder;

use vb6_lsp::controls;
use vb6_lsp::lsp::Vb6LanguageServer;
use vb6_lsp::workspace::res_json::{
    find_string_table_block, parse_resource_type_name, resource_entry_to_json,
    string_entries_to_json, ResourceTypeNameError,
};
use vb6_lsp::workspace::{
    parse_string_table, read_res_file, write_res_file, ResourceEntry, ResourceId, ResourceType,
};

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let args: Vec<String> = env::args().collect();

    // Check if we should start LSP server or handle CLI commands
    // --stdio is the standard LSP flag, no args means LSP mode too
    if args.len() > 1 && args[1] != "--stdio" {
        return handle_cli_command(&args[1..]);
    }

    // Initialize logging - MUST write to stderr, not stdout!
    // LSP uses stdout for JSON-RPC communication
    tracing_subscriber::registry()
        .with(tracing_subscriber::EnvFilter::new(
            std::env::var("RUST_LOG").unwrap_or_else(|_| "vb6_lsp=info".into()),
        ))
        .with(tracing_subscriber::fmt::layer().with_writer(std::io::stderr))
        .init();

    tracing::info!("Starting VB6 Language Server...");

    // Create the LSP service
    let stdin = tokio::io::stdin();
    let stdout = tokio::io::stdout();

    let (service, socket) = LspService::new(Vb6LanguageServer::new);

    // Run the server
    Server::new(stdin, stdout, socket).serve(service).await;

    Ok(())
}

/// Build the form export IR, optionally decoding Tier-3 proprietary bags live via
/// COM (Windows only, where a licensed control is installed). Falls back to opaque
/// Tier-3 when `--com-decode` is off or the bridge is unavailable.
fn build_form_export_cli(
    path: &std::path::Path,
    com_decode: bool,
) -> anyhow::Result<controls::export::FormExport> {
    if com_decode {
        #[cfg(windows)]
        {
            if let Some(dec) = com_decoder::ScriptedComDecoder::new() {
                return Ok(controls::export::build_form_export_with(
                    path,
                    controls::resources::OcxBagPolicy::ComDecode,
                    Some(&dec),
                )?);
            }
            eprintln!("--com-decode: COM bridge unavailable; using opaque Tier-3");
        }
        #[cfg(not(windows))]
        eprintln!("--com-decode is Windows-only; using opaque Tier-3");
    }
    Ok(controls::export::build_form_export(path)?)
}

/// Resolve a form's companion resources for the CLI, optionally decoding
/// proprietary control bags via COM (Windows + a registered/licensed control).
/// Mirrors [`build_form_export_cli`]'s decoder selection but returns the raw
/// resolved resources, for callers (e.g. `extract-form`) that need the values.
fn resolve_form_cli(
    path: &std::path::Path,
    com_decode: bool,
) -> anyhow::Result<Vec<controls::resources::ResolvedResource>> {
    use controls::resources::{resolve_form_resources, OcxBagPolicy};
    if com_decode {
        #[cfg(windows)]
        {
            if let Some(dec) = com_decoder::ScriptedComDecoder::new() {
                return Ok(resolve_form_resources(path, OcxBagPolicy::ComDecode, Some(&dec))?);
            }
            eprintln!("--com-decode: COM bridge unavailable; using opaque Tier-3");
        }
        #[cfg(not(windows))]
        eprintln!("--com-decode is Windows-only; using opaque Tier-3");
    }
    Ok(resolve_form_resources(path, OcxBagPolicy::Opaque, None)?)
}

/// Handle CLI commands for resource file operations
fn handle_cli_command(args: &[String]) -> anyhow::Result<()> {
    match args[0].as_str() {
        "read-res" => cmd_read_res(args),
        "read-form" => cmd_read_form(args),
        "extract-form" => cmd_extract_form(args),
        "read-bags" => cmd_read_bags(args),
        "write-res" => cmd_write_res(args),
        "parse-string-table" => cmd_parse_string_table(args),
        _ => {
            eprintln!("Unknown command: {}", args[0]);
            eprintln!("Available commands:");
            eprintln!("  read-res <file.res>                    - Read a .res file");
            eprintln!("  read-form <form> [--com-decode]        - Control tree with .frx/.ctx data resolved");
            eprintln!("  extract-form <form> <dir> [--com-decode] - Write the form's pictures and bags");
            eprintln!("  read-bags <form>                       - Decode the VB property bags of third-party controls as JSON");
            eprintln!("  write-res <input.json> <output.res>    - Write a .res file");
            eprintln!("  parse-string-table <file.res> <id>     - Parse string table");
            std::process::exit(1);
        }
    }
}

fn cmd_read_res(args: &[String]) -> anyhow::Result<()> {
    if args.len() < 2 {
        eprintln!("Usage: vb6-lsp read-res <file.res>");
        std::process::exit(1);
    }

    let file_path = &args[1];
    let resources = read_res_file(file_path)?;

    // Convert to JSON-friendly format
    let json_resources: Vec<serde_json::Value> =
        resources.iter().map(resource_entry_to_json).collect();

    println!("{}", serde_json::json!({
        "resources": json_resources
    }));

    Ok(())
}

fn cmd_read_form(args: &[String]) -> anyhow::Result<()> {
    let positional: Vec<&String> =
        args[1..].iter().filter(|a| !a.starts_with("--")).collect();
    if positional.is_empty() {
        eprintln!("Usage: vb6-lsp read-form <file.frm|.ctl|.pag|.dob> [--com-decode]");
        std::process::exit(1);
    }
    let path = std::path::Path::new(positional[0].as_str());
    let com_decode = args.iter().any(|a| a == "--com-decode");
    let export = build_form_export_cli(path, com_decode)?;
    println!("{}", serde_json::to_string_pretty(&export)?);
    Ok(())
}

fn cmd_read_bags(args: &[String]) -> anyhow::Result<()> {
    let positional: Vec<&String> = args[1..].iter().filter(|a| !a.starts_with("--")).collect();
    if positional.is_empty() {
        eprintln!("Usage: vb6-lsp read-bags <file.frm|.ctl|.pag|.dob>");
        std::process::exit(1);
    }
    let path = std::path::Path::new(positional[0].as_str());
    let resolved = resolve_form_cli(path, false)?;
    let mut bags: Vec<serde_json::Value> = Vec::new();
    for r in &resolved {
        if let Ok(controls::frx::FrxValue::OcxBag { clsid, data }) = &r.value {
            let decoded = controls::propbag::decode(data);
            bags.push(serde_json::json!({
                "control": r.reference.control_path,
                "control_type": r.reference.control_type,
                "property": r.reference.property,
                "clsid": clsid.map(|c| format!("{:02x?}", c)),
                "bytes": data.len(),
                "bag": decoded,
            }));
        }
    }
    println!("{}", serde_json::to_string_pretty(&serde_json::json!({ "form": positional[0], "bags": bags }))?);
    Ok(())
}

fn cmd_extract_form(args: &[String]) -> anyhow::Result<()> {
    let positional: Vec<&String> =
        args[1..].iter().filter(|a| !a.starts_with("--")).collect();
    if positional.len() < 2 {
        eprintln!(
            "Usage: vb6-lsp extract-form <file.frm|.ctl> <output-dir> [--com-decode]"
        );
        std::process::exit(1);
    }
    let path = std::path::Path::new(positional[0].as_str());
    let outdir = std::path::Path::new(positional[1].as_str());
    let com_decode = args.iter().any(|a| a == "--com-decode");
    std::fs::create_dir_all(outdir)?;

    let resolved = resolve_form_cli(path, com_decode)?;

    let mut extracted: Vec<serde_json::Value> = Vec::new();
    for r in &resolved {
        if let Some(entry) = extract_resource(r, outdir)? {
            extracted.push(entry);
        }
    }

    println!("{}", serde_json::json!({
        "form": positional[0], "outdir": positional[1], "extracted": extracted
    }));
    Ok(())
}

/// Write a single resolved resource to `outdir` and return its JSON manifest
/// entry (or `None` for resource kinds that are skipped).
fn extract_resource(
    r: &controls::resources::ResolvedResource,
    outdir: &std::path::Path,
) -> anyhow::Result<Option<serde_json::Value>> {
    let base = format!("{}.{}", r.reference.control_path, r.reference.property);
    let entry = match &r.value {
        Ok(controls::frx::FrxValue::Picture { format, data, .. }) => {
            let fname = format!("{}.{}", base, format.ext());
            std::fs::write(outdir.join(&fname), data)?;
            serde_json::json!({
                "property": r.reference.property, "kind": "picture",
                "format": format!("{:?}", format), "file": fname, "bytes": data.len()
            })
        }
        Ok(controls::frx::FrxValue::OcxBag { clsid, data }) => {
            let fname = format!("{}.bin", base);
            std::fs::write(outdir.join(&fname), data)?;
            if let Some(g) = clsid {
                std::fs::write(
                    outdir.join(format!("{}.clsid.txt", base)),
                    controls::export::format_guid(*g),
                )?;
            }
            serde_json::json!({
                "property": r.reference.property, "kind": "ocx_bag",
                "clsid": (*clsid).map(controls::export::format_guid),
                "file": fname, "bytes": data.len()
            })
        }
        Ok(controls::frx::FrxValue::DecodedBag { clsid, properties }) => {
            let fname = format!("{}.properties.json", base);
            let doc = serde_json::json!({
                "clsid": (*clsid).map(controls::export::format_guid),
                "properties": properties
                    .iter()
                    .map(|(k, v)| serde_json::json!([k, v]))
                    .collect::<Vec<_>>(),
            });
            std::fs::write(outdir.join(&fname), serde_json::to_string_pretty(&doc)?)?;
            serde_json::json!({
                "property": r.reference.property, "kind": "decoded_bag",
                "clsid": (*clsid).map(controls::export::format_guid),
                "file": fname, "properties": properties.len()
            })
        }
        Ok(controls::frx::FrxValue::Text(s)) => {
            let fname = format!("{}.txt", base);
            std::fs::write(outdir.join(&fname), s)?;
            serde_json::json!({
                "property": r.reference.property, "kind": "text", "file": fname
            })
        }
        Err(e) => serde_json::json!({
            "property": r.reference.property, "kind": "error",
            "message": e.to_string()
        }),
        _ => return Ok(None),
    };
    Ok(Some(entry))
}

fn cmd_write_res(args: &[String]) -> anyhow::Result<()> {
    if args.len() < 3 {
        eprintln!("Usage: vb6-lsp write-res <input.json> <output.res>");
        std::process::exit(1);
    }

    let json_file = &args[1];
    let output_file = &args[2];

    // Read JSON input
    let json_content = std::fs::read_to_string(json_file)?;
    let json: serde_json::Value = serde_json::from_str(&json_content)?;

    // Parse resources from JSON
    let mut resources = Vec::new();
    if let Some(res_array) = json["resources"].as_array() {
        for res in res_array {
            resources.push(parse_resource_entry(res)?);
        }
    }

    // Write the .res file
    write_res_file(output_file, &resources)?;

    println!("{}", serde_json::json!({
        "success": true,
        "file": output_file,
        "resourceCount": resources.len()
    }));

    Ok(())
}

/// Build a single [`ResourceEntry`] from its JSON description.
fn parse_resource_entry(res: &serde_json::Value) -> anyhow::Result<ResourceEntry> {
    let resource_type_str = res["resource_type"].as_str()
        .ok_or_else(|| anyhow::anyhow!("Missing resource_type"))?;

    let resource_type = parse_resource_type(resource_type_str)?;

    let name = if res["name"]["type"].as_str() == Some("Id") {
        ResourceId::Id(res["name"]["value"].as_u64()
            .ok_or_else(|| anyhow::anyhow!("Invalid ID value"))? as u16)
    } else {
        ResourceId::Name(res["name"]["value"].as_str()
            .ok_or_else(|| anyhow::anyhow!("Invalid Name value"))?.to_string())
    };

    let language_id = res["language_id"].as_u64()
        .ok_or_else(|| anyhow::anyhow!("Missing language_id"))? as u16;

    let data_base64 = res["data_base64"].as_str()
        .ok_or_else(|| anyhow::anyhow!("Missing data_base64"))?;
    let data = base64::engine::general_purpose::STANDARD.decode(data_base64)?;

    Ok(ResourceEntry::new(resource_type, name, language_id, data))
}

fn cmd_parse_string_table(args: &[String]) -> anyhow::Result<()> {
    if args.len() < 3 {
        eprintln!("Usage: vb6-lsp parse-string-table <file.res> <block_id>");
        std::process::exit(1);
    }

    let file_path = &args[1];
    let block_id: u16 = args[2].parse()
        .map_err(|_| anyhow::anyhow!("Invalid block_id: must be a number"))?;

    // Read the .res file
    let resources = read_res_file(file_path)?;

    // Find the string table resource
    let string_resource = find_string_table_block(&resources, block_id)
        .ok_or_else(|| anyhow::anyhow!("String table block {} not found", block_id))?;

    // Parse the string table
    let strings = parse_string_table(&string_resource.data, block_id)?;

    let json_strings = string_entries_to_json(&strings);

    println!("{}", serde_json::json!({
        "strings": json_strings
    }));

    Ok(())
}

/// Parse a resource type string (e.g., "Bitmap", "Icon", "Named(\"CUSTOM\")")
fn parse_resource_type(s: &str) -> anyhow::Result<ResourceType> {
    parse_resource_type_name(s).map_err(|e| match e {
        ResourceTypeNameError::InvalidCustomId(parse_err) => anyhow::Error::from(parse_err),
        ResourceTypeNameError::Unknown => anyhow::anyhow!("Unknown resource type: {}", s),
    })
}
