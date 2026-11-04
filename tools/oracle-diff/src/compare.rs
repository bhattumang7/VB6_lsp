//! Joins the compiler's verdict with the parser's diagnostics and classifies
//! the outcome.

use crate::oracle::Verdict;
use crate::ours::OurDiagnostic;
use std::collections::HashMap;

/// How the two sides relate for one case.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Outcome {
    /// Compiler accepted, parser reported nothing.
    AgreeAccept,
    /// Both rejected and the compiler's message is among the parser's messages.
    AgreeRejectSameMessage,
    /// Both rejected but no parser message equals the compiler's.
    AgreeRejectDifferentMessage,
    /// Compiler accepted the whole module, parser reported diagnostics.
    FalseReject,
    /// The compiler stopped at a semantic error and the parser reported a
    /// diagnostic. The compiler does not read declarations and bodies in
    /// source order, so its stop does not show that it read the parser's
    /// position without a syntax error: it gives no verdict on the diagnostic.
    /// Needs a closer look.
    Unverifiable,
    /// Compiler rejected with a message classed as syntax, parser reported nothing.
    FalseAccept,
    /// Compiler rejected with a message the taxonomy classes as semantic; the
    /// parser reported nothing, which is expected.
    SemanticOnly,
    /// Compiler rejected with a message that has no taxonomy entry yet, parser
    /// reported nothing. Must be classified before it can count either way.
    Unclassified,
    /// Compiler gave no usable verdict (timeout, project-level error or unrecognised log).
    NoVerdict,
}

/// Message class table: compiler message text to `syntax` / `semantic` /
/// `project`. Loaded from a checked-in TSV (`message<TAB>class`); a message
/// ending in `*` stands for every message that begins with the text before it.
#[derive(Default)]
pub struct Taxonomy {
    classes: HashMap<String, String>,
}

impl Taxonomy {
    pub fn parse(text: &str) -> Result<Taxonomy, String> {
        let mut classes = HashMap::new();
        for (n, line) in text.lines().enumerate() {
            if line.trim().is_empty() || line.starts_with('#') {
                continue;
            }
            let (message, class) = line
                .split_once('\t')
                .ok_or_else(|| format!("taxonomy line {}: expected `message<TAB>class`", n + 1))?;
            if !matches!(class, "syntax" | "semantic" | "project") {
                return Err(format!("taxonomy line {}: unknown class `{class}`", n + 1));
            }
            if classes.insert(message.to_string(), class.to_string()).is_some() {
                return Err(format!("taxonomy line {}: duplicate message `{message}`", n + 1));
            }
        }
        Ok(Taxonomy { classes })
    }

    /// The class of `message`: an exact entry wins, otherwise the longest entry
    /// ending in `*` that is a prefix of the message.
    pub fn class_of(&self, message: &str) -> Option<&str> {
        if let Some(class) = self.classes.get(message) {
            return Some(class);
        }
        self.classes
            .iter()
            .filter(|(key, _)| key.strip_suffix('*').is_some_and(|prefix| message.starts_with(prefix)))
            .max_by_key(|(key, _)| key.len())
            .map(|(_, class)| class.as_str())
    }
}

pub fn classify(verdict: &Verdict, ours: &[OurDiagnostic], taxonomy: &Taxonomy) -> Outcome {
    match (verdict, ours.is_empty()) {
        (Verdict::Accepted, true) => Outcome::AgreeAccept,
        (Verdict::Accepted, false) => Outcome::FalseReject,
        (Verdict::Rejected { message, .. }, false) => {
            if ours.iter().any(|d| &d.message == message) {
                Outcome::AgreeRejectSameMessage
            } else if taxonomy.class_of(message) == Some("semantic") {
                Outcome::Unverifiable
            } else {
                Outcome::AgreeRejectDifferentMessage
            }
        }
        (Verdict::Rejected { message, .. }, true) => match taxonomy.class_of(message) {
            Some("syntax") => Outcome::FalseAccept,
            Some(_) => Outcome::SemanticOnly,
            None => Outcome::Unclassified,
        },
        (Verdict::TimedOut | Verdict::Unrecognised { .. } | Verdict::ProjectError { .. }, _) => {
            Outcome::NoVerdict
        }
    }
}
