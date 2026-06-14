//! W8: conditional-compilation sanity.
//!
//! Confirms the parser handles `#If`/`#ElseIf`/`#Else`/`#End If` and `#Const`
//! without emitting false diagnostics, evaluates the predefined constants
//! (Win32/Win64/VB6), and skips dead branches (so unparseable text in an
//! inactive branch does not surface as an error).

use vb6_core::session::Session;

fn session(src: &str) -> Session {
    Session::from_sources(vec![("M.bas".to_string(), src.as_bytes().to_vec())])
}

fn no_parse_errors(s: &Session) -> bool {
    s.diagnostics(0).is_empty()
}

fn has_sub(s: &Session, name: &str) -> bool {
    s.document_symbols(0).iter().any(|x| x.name.eq_ignore_ascii_case(name))
}

#[test]
fn predefined_win32_branch_is_taken() {
    // Win32 is predefined true; Win64 false. The active branch defines `A`.
    let src = "#If Win32 Then\n\
               Sub A()\nEnd Sub\n\
               #Else\n\
               Sub B()\nEnd Sub\n\
               #End If\n";
    let s = session(src);
    assert!(no_parse_errors(&s), "no errors: {:?}", s.diagnostics(0));
    assert!(has_sub(&s, "A"), "Win32 branch should define A");
    assert!(!has_sub(&s, "B"), "Win64/#Else branch should be inactive");
}

#[test]
fn dead_branch_content_is_not_diagnosed() {
    // Win64 is predefined false; the dead branch holds text that would otherwise
    // be a parse error. Skipping it must produce no diagnostics.
    let src = "#If Win64 Then\n\
               Dim Dim Dim garbage !!!\n\
               #End If\n\
               Sub Ok()\nEnd Sub\n";
    let s = session(src);
    assert!(no_parse_errors(&s), "dead branch leaked errors: {:?}", s.diagnostics(0));
    assert!(has_sub(&s, "Ok"));
}

#[test]
fn user_const_drives_branch_selection() {
    let src = "#Const Flag = 1\n\
               #If Flag Then\n\
               Sub Yes()\nEnd Sub\n\
               #Else\n\
               Sub No()\nEnd Sub\n\
               #End If\n";
    let s = session(src);
    assert!(no_parse_errors(&s), "no errors: {:?}", s.diagnostics(0));
    assert!(has_sub(&s, "Yes"));
    assert!(!has_sub(&s, "No"));
}

#[test]
fn elseif_chain_selects_first_true_branch() {
    let src = "#Const Mode = 2\n\
               #If Mode = 1 Then\n\
               Sub One()\nEnd Sub\n\
               #ElseIf Mode = 2 Then\n\
               Sub Two()\nEnd Sub\n\
               #Else\n\
               Sub Other()\nEnd Sub\n\
               #End If\n";
    let s = session(src);
    assert!(no_parse_errors(&s), "no errors: {:?}", s.diagnostics(0));
    assert!(has_sub(&s, "Two"));
    assert!(!has_sub(&s, "One"));
    assert!(!has_sub(&s, "Other"));
}
