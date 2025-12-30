//! VB6 Language Server - Main Entry Point
//!
//! A Language Server Protocol implementation for Visual Basic 6
//! with Claude AI integration for intelligent code assistance.

use tower_lsp::{LspService, Server};
use tracing_subscriber::{layer::SubscriberExt, util::SubscriberInitExt};

mod analysis;
mod claude;
mod lsp;
mod parser;

use lsp::Vb6LanguageServer;

#[tokio::main]
async fn main() -> anyhow::Result<()> {
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
