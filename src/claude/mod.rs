//! Claude AI Integration
//!
//! Provides AI-powered code assistance using Claude Sonnet.

use serde::{Deserialize, Serialize};

/// Claude API client
pub struct ClaudeClient {
    api_key: String,
    client: reqwest::Client,
    model: String,
}

impl ClaudeClient {
    pub fn new(api_key: String) -> Self {
        Self {
            api_key,
            client: reqwest::Client::new(),
            model: "claude-sonnet-4-20250514".to_string(),
        }
    }

    /// Get code completion suggestions from Claude
    pub async fn get_completion_suggestion(
        &self,
        code_context: &str,
        cursor_position: &str,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let prompt = format!(
            "You are a Visual Basic 6 expert. Given the following VB6 code context, suggest a completion at the cursor position.\n\nCode:\n{}\n\nCursor at: {}\n\nProvide only the suggested completion, no explanations.",
            code_context, cursor_position
        );

        self.send_message(&prompt).await
    }

    /// Explain code using Claude
    pub async fn explain_code(&self, code: &str) -> Result<String, Box<dyn std::error::Error>> {
        let prompt = format!(
            "Explain the following Visual Basic 6 code in a concise way:\n\n{}",
            code
        );

        self.send_message(&prompt).await
    }

    /// Suggest refactoring using Claude
    pub async fn suggest_refactoring(
        &self,
        code: &str,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let prompt = format!(
            "Suggest improvements or refactorings for this Visual Basic 6 code. Be specific and practical:\n\n{}",
            code
        );

        self.send_message(&prompt).await
    }

    /// Explain an error using Claude
    pub async fn explain_error(
        &self,
        error_message: &str,
        code_context: &str,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let prompt = format!(
            "Explain this VB6 error and how to fix it:\n\nError: {}\n\nCode context:\n{}",
            error_message, code_context
        );

        self.send_message(&prompt).await
    }

    /// Generate documentation using Claude
    pub async fn generate_documentation(
        &self,
        code: &str,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let prompt = format!(
            "Generate concise documentation comments for this Visual Basic 6 code:\n\n{}",
            code
        );

        self.send_message(&prompt).await
    }

    /// Suggest migration to VB.NET/C# using Claude
    pub async fn suggest_migration(
        &self,
        code: &str,
        target_language: &str,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let prompt = format!(
            "Convert this Visual Basic 6 code to {}. Explain any important changes:\n\n{}",
            target_language, code
        );

        self.send_message(&prompt).await
    }

    /// Send a message to Claude API
    async fn send_message(&self, prompt: &str) -> Result<String, Box<dyn std::error::Error>> {
        let request = ClaudeRequest {
            model: self.model.clone(),
            max_tokens: 1024,
            messages: vec![Message {
                role: "user".to_string(),
                content: prompt.to_string(),
            }],
        };

        let response = self
            .client
            .post("https://api.anthropic.com/v1/messages")
            .header("x-api-key", &self.api_key)
            .header("anthropic-version", "2023-06-01")
            .header("content-type", "application/json")
            .json(&request)
            .send()
            .await?;

        if !response.status().is_success() {
            let error_text = response.text().await?;
            return Err(format!("Claude API error: {}", error_text).into());
        }

        let claude_response: ClaudeResponse = response.json().await?;

        // Extract text from first content block
        if let Some(content) = claude_response.content.first() {
            Ok(content.text.clone())
        } else {
            Err("No response from Claude".into())
        }
    }
}

#[derive(Debug, Serialize)]
struct ClaudeRequest {
    model: String,
    max_tokens: u32,
    messages: Vec<Message>,
}

#[derive(Debug, Serialize, Deserialize)]
struct Message {
    role: String,
    content: String,
}

#[derive(Debug, Deserialize)]
struct ClaudeResponse {
    content: Vec<ContentBlock>,
}

#[derive(Debug, Deserialize)]
struct ContentBlock {
    text: String,
}

/// Utility to get code context around a position
pub fn get_code_context(full_text: &str, line: usize, character: usize, context_lines: usize) -> String {
    let lines: Vec<&str> = full_text.lines().collect();

    let start_line = line.saturating_sub(context_lines);
    let end_line = (line + context_lines + 1).min(lines.len());

    let mut context = String::new();
    for (i, line_text) in lines[start_line..end_line].iter().enumerate() {
        let current_line = start_line + i;
        if current_line == line {
            context.push_str(&format!("{:4} > {}\n", current_line + 1, line_text));
            context.push_str(&format!("      {: >width$}^\n", "", width = character));
        } else {
            context.push_str(&format!("{:4}   {}\n", current_line + 1, line_text));
        }
    }

    context
}
