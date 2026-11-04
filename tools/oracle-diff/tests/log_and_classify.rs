use oracle_diff::compare::{classify, Outcome, Taxonomy};
use oracle_diff::oracle::{parse_log, Case, ModuleKind, Verdict};
use oracle_diff::ours::OurDiagnostic;

fn diag(message: &str) -> OurDiagnostic {
    OurDiagnostic { code: 1, line: 1, message: message.to_string() }
}

#[test]
fn success_log_is_accepted() {
    let log = "\r\nBuild of 'p.exe' succeeded.\r\n";
    assert_eq!(parse_log(log, Some(0)), Verdict::Accepted);
}

#[test]
fn compile_error_log_keeps_file_line_and_message() {
    let log = "\r\nCompile Error in File 'C:\\w\\m.bas', Line 4 : Block If without End If\r\nBuild of 'p.exe' failed.\r\n";
    assert_eq!(
        parse_log(log, Some(1)),
        Verdict::Rejected {
            file: "C:\\w\\m.bas".to_string(),
            line: Some(4),
            message: "Block If without End If".to_string(),
        }
    );
}

#[test]
fn project_level_failure_is_a_project_error() {
    let log = "\r\nAmbiguous name detected: Main\r\nBuild of 'p.exe' failed.\r\n";
    assert_eq!(
        parse_log(log, Some(1)),
        Verdict::ProjectError { message: "Ambiguous name detected: Main".to_string() }
    );
}

#[test]
fn success_with_a_failing_exit_code_is_unrecognised() {
    let log = "Build of 'p.exe' succeeded.\r\n";
    assert_eq!(
        parse_log(log, Some(1)),
        Verdict::Unrecognised { exit_code: Some(1), log: log.to_string() }
    );
}

#[test]
fn empty_log_is_unrecognised() {
    assert_eq!(
        parse_log("", Some(0)),
        Verdict::Unrecognised { exit_code: Some(0), log: String::new() }
    );
}

#[test]
fn case_from_text_normalises_every_break_to_crlf() {
    let case = Case::from_text("c", ModuleKind::Bas, "a\nb\r\nc\rd");
    assert_eq!(case.source, b"a\r\nb\r\nc\r\nd".to_vec());
}

#[test]
fn taxonomy_rejects_unknown_classes_and_duplicates() {
    assert!(Taxonomy::parse("msg\tbogus\n").is_err());
    assert!(Taxonomy::parse("msg\tsyntax\nmsg\tsemantic\n").is_err());
    assert!(Taxonomy::parse("no tab here\n").is_err());
    let ok = Taxonomy::parse("# comment\n\nSyntax error\tsyntax\nVariable not defined\tsemantic\n").unwrap();
    assert_eq!(ok.class_of("Syntax error"), Some("syntax"));
    assert_eq!(ok.class_of("Variable not defined"), Some("semantic"));
    assert_eq!(ok.class_of("Other"), None);
}

fn rejected(message: &str) -> Verdict {
    Verdict::Rejected { file: "m.bas".into(), line: Some(0), message: message.into() }
}

#[test]
fn every_outcome_is_reachable() {
    let taxonomy = Taxonomy::parse("Syntax error\tsyntax\nVariable not defined\tsemantic\n").unwrap();
    assert_eq!(classify(&Verdict::Accepted, &[], &taxonomy), Outcome::AgreeAccept);
    assert_eq!(classify(&Verdict::Accepted, &[diag("x")], &taxonomy), Outcome::FalseReject);
    assert_eq!(
        classify(&rejected("Syntax error"), &[diag("Syntax error")], &taxonomy),
        Outcome::AgreeRejectSameMessage
    );
    assert_eq!(
        classify(&rejected("Syntax error"), &[diag("Expected: =")], &taxonomy),
        Outcome::AgreeRejectDifferentMessage
    );
    assert_eq!(classify(&rejected("Syntax error"), &[], &taxonomy), Outcome::FalseAccept);
    assert_eq!(classify(&rejected("Variable not defined"), &[], &taxonomy), Outcome::SemanticOnly);
    assert_eq!(classify(&rejected("Never seen"), &[], &taxonomy), Outcome::Unclassified);
    assert_eq!(classify(&Verdict::TimedOut, &[], &taxonomy), Outcome::NoVerdict);
    assert_eq!(
        classify(&Verdict::ProjectError { message: "m".into() }, &[], &taxonomy),
        Outcome::NoVerdict
    );
}
