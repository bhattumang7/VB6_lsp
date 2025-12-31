//! VB6 Language Server - Main Entry Point
//!
//! A Language Server Protocol implementation for Visual Basic 6
//! with Claude AI integration for intelligent code assistance.

use std::env;
use tower_lsp::{LspService, Server};
use tracing_subscriber::{layer::SubscriberExt, util::SubscriberInitExt};

mod analysis;
mod claude;
mod controls;
mod lsp;
mod parser;
mod utils;
mod workspace;

use lsp::Vb6LanguageServer;
use workspace::{read_res_file, write_res_file, parse_string_table, ResourceEntry, ResourceId, ResourceType};

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

    let (service, socket) = LspService::new(|client| Vb6LanguageServer::new(client));

    // Run the server
    Server::new(stdin, stdout, socket).serve(service).await;

    Ok(())
}

/// Handle CLI commands for resource file operations
fn handle_cli_command(args: &[String]) -> anyhow::Result<()> {
    match args[0].as_str() {
        "read-res" => {
            if args.len() < 2 {
                eprintln!("Usage: vb6-lsp read-res <file.res>");
                std::process::exit(1);
            }

            let file_path = &args[1];
            let resources = read_res_file(file_path)?;

            // Convert to JSON-friendly format
            let json_resources: Vec<serde_json::Value> = resources.iter().map(|r| {
                serde_json::json!({
                    "resource_type": format!("{:?}", r.resource_type),
                    "name": match &r.name {
                        ResourceId::Id(id) => serde_json::json!({
                            "type": "Id",
                            "value": id
                        }),
                        ResourceId::Name(name) => serde_json::json!({
                            "type": "Name",
                            "value": name
                        })
                    },
                    "language_id": r.language_id,
                    "data_size": r.data.len(),
                    "data_base64": base64::encode(&r.data)
                })
            }).collect();

            println!("{}", serde_json::json!({
                "resources": json_resources
            }));

            Ok(())
        }

        "write-res" => {
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
                    let data = base64::decode(data_base64)?;

                    resources.push(ResourceEntry::new(resource_type, name, language_id, data));
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

        "parse-string-table" => {
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
            let string_resource = resources.iter()
                .find(|r| {
                    r.resource_type == ResourceType::String &&
                    matches!(r.name, ResourceId::Id(id) if id == block_id)
                })
                .ok_or_else(|| anyhow::anyhow!("String table block {} not found", block_id))?;

            // Parse the string table
            let strings = parse_string_table(&string_resource.data, block_id)?;

            let json_strings: Vec<serde_json::Value> = strings.iter().map(|s| {
                serde_json::json!({
                    "id": s.id,
                    "value": s.value
                })
            }).collect();

            println!("{}", serde_json::json!({
                "strings": json_strings
            }));

            Ok(())
        }

        _ => {
            eprintln!("Unknown command: {}", args[0]);
            eprintln!("Available commands:");
            eprintln!("  read-res <file.res>                    - Read a .res file");
            eprintln!("  write-res <input.json> <output.res>    - Write a .res file");
            eprintln!("  parse-string-table <file.res> <id>     - Parse string table");
            std::process::exit(1);
        }
    }
}

/// Parse a resource type string (e.g., "Bitmap", "Icon", "Named(\"CUSTOM\")")
fn parse_resource_type(s: &str) -> anyhow::Result<ResourceType> {
    Ok(match s {
        "Cursor" => ResourceType::Cursor,
        "Bitmap" => ResourceType::Bitmap,
        "Icon" => ResourceType::Icon,
        "Menu" => ResourceType::Menu,
        "Dialog" => ResourceType::Dialog,
        "String" => ResourceType::String,
        "FontDir" => ResourceType::FontDir,
        "Font" => ResourceType::Font,
        "Accelerator" => ResourceType::Accelerator,
        "RcData" => ResourceType::RcData,
        "MessageTable" => ResourceType::MessageTable,
        "GroupCursor" => ResourceType::GroupCursor,
        "GroupIcon" => ResourceType::GroupIcon,
        "Version" => ResourceType::Version,
        "DlgInclude" => ResourceType::DlgInclude,
        "PlugPlay" => ResourceType::PlugPlay,
        "Vxd" => ResourceType::Vxd,
        "AniCursor" => ResourceType::AniCursor,
        "AniIcon" => ResourceType::AniIcon,
        "Html" => ResourceType::Html,
        "Manifest" => ResourceType::Manifest,
        "Toolbar" => ResourceType::Toolbar,
        "DlgInit" => ResourceType::DlgInit,
        s if s.starts_with("Custom(") => {
            let id_str = s.trim_start_matches("Custom(").trim_end_matches(")");
            let id: u16 = id_str.parse()?;
            ResourceType::Custom(id)
        }
        s if s.starts_with("Named(") => {
            let name = s.trim_start_matches("Named(\"").trim_end_matches("\")");
            ResourceType::Named(name.to_string())
        }
        _ => anyhow::bail!("Unknown resource type: {}", s)
    })
}
