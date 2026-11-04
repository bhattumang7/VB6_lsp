//! Template-by-vocabulary matrices: every template is compiled once per word,
//! and the compiler's answer is kept as a fixture the parser tests replay.

use crate::compare::Taxonomy;
use crate::oracle::{Case, ModuleKind, Oracle, Verdict};

/// One template block of a templates file.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Template {
    pub name: String,
    pub kind: ModuleKind,
    /// Module text with `{k}` where the word goes.
    pub text: String,
}

/// Parses a templates file: `=== <name> <bas|cls>` starts a block, `#` lines
/// before the first block are comments.
pub fn parse_templates(text: &str) -> Result<Vec<Template>, String> {
    let mut out: Vec<Template> = Vec::new();
    for line in text.lines() {
        let line = line.trim_end_matches('\r');
        if let Some(header) = line.strip_prefix("=== ") {
            let mut parts = header.split_whitespace();
            let (Some(name), Some(kind), None) = (parts.next(), parts.next(), parts.next()) else {
                return Err(format!("bad template header `{line}`"));
            };
            let kind = match kind {
                "bas" => ModuleKind::Bas,
                "cls" => ModuleKind::Cls,
                other => return Err(format!("unknown module kind `{other}`")),
            };
            if out.iter().any(|t| t.name == name) {
                return Err(format!("duplicate template `{name}`"));
            }
            out.push(Template { name: name.to_string(), kind, text: String::new() });
        } else if let Some(last) = out.last_mut() {
            last.text.push_str(line);
            last.text.push('\n');
        } else if !line.starts_with('#') && !line.trim().is_empty() {
            return Err(format!("text before the first template: `{line}`"));
        }
    }
    Ok(out)
}

/// Parses a words file: one word per line, blank lines and `#` lines ignored.
pub fn parse_words(text: &str) -> Vec<String> {
    text.lines()
        .map(|l| l.trim_end_matches('\r').trim())
        .filter(|l| !l.is_empty() && !l.starts_with('#'))
        .map(str::to_string)
        .collect()
}

/// The source of `template` with `word` in place.
pub fn instantiate(template: &Template, word: &str) -> String {
    template.text.replace("{k}", &unescape(word))
}

/// `\n` and `\t` in a word stand for a line break and a tab.
pub fn unescape(word: &str) -> String {
    word.replace("\\n", "\n").replace("\\t", "\t")
}

/// What the compiler made of one instantiated template.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Fate {
    Accepted,
    /// Rejected with a message the taxonomy classes as semantic: the form parsed.
    Semantic(String),
    /// Rejected with a message the taxonomy classes as syntax.
    Syntax(String),
    /// The compiler terminated abnormally on this input (exit code recorded), so
    /// it gives no answer.
    Unanswered(String),
}

impl Fate {
    pub fn column(&self) -> String {
        match self {
            Fate::Accepted => "accepted\t".to_string(),
            Fate::Semantic(m) => format!("semantic\t{m}"),
            Fate::Syntax(m) => format!("syntax\t{m}"),
            Fate::Unanswered(m) => format!("unanswered\t{m}"),
        }
    }

    pub fn parse(class: &str, message: &str) -> Option<Fate> {
        Some(match class {
            "accepted" => Fate::Accepted,
            "semantic" => Fate::Semantic(message.to_string()),
            "syntax" => Fate::Syntax(message.to_string()),
            "unanswered" => Fate::Unanswered(message.to_string()),
            _ => return None,
        })
    }

    /// `Some(true)` when the compiler found nothing wrong with the form,
    /// `Some(false)` when it rejected the form itself, `None` when it gave no answer.
    pub fn parses(&self) -> Option<bool> {
        match self {
            Fate::Accepted | Fate::Semantic(_) => Some(true),
            Fate::Syntax(_) => Some(false),
            Fate::Unanswered(_) => None,
        }
    }
}

/// Runs every template with every word. Fails when the compiler gives no
/// verdict or a message the taxonomy has no class for.
pub fn run(
    oracle: &Oracle,
    taxonomy: &Taxonomy,
    templates: &[Template],
    words: &[String],
) -> Result<Vec<(String, String, Fate)>, String> {
    let mut cases = Vec::new();
    let mut keys = Vec::new();
    for t in templates {
        for w in words {
            cases.push(Case::from_text(format!("{}/{w}", t.name), t.kind, &instantiate(t, w)));
            keys.push((t.name.clone(), w.clone()));
        }
    }
    let verdicts = oracle.run(&cases);
    let mut rows = Vec::new();
    let mut unknown = std::collections::BTreeSet::new();
    for ((name, word), verdict) in keys.into_iter().zip(verdicts) {
        let fate = match verdict {
            Verdict::Accepted => Fate::Accepted,
            Verdict::Rejected { message, .. } => match taxonomy.class_of(&message) {
                Some("syntax") => Fate::Syntax(message),
                Some(_) => Fate::Semantic(message),
                None => {
                    unknown.insert(message);
                    continue;
                }
            },
            Verdict::Unrecognised { exit_code: Some(code), log } if code < 0 && log.trim().is_empty() => {
                Fate::Unanswered(format!("compiler terminated abnormally, exit {code}"))
            }
            other => return Err(format!("{name}/{word}: no verdict from the compiler: {other:?}")),
        };
        rows.push((name, word, fate));
    }
    if !unknown.is_empty() {
        return Err(format!("messages without a class: {unknown:?}"));
    }
    Ok(rows)
}
