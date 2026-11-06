//! A module may declare a name once, apart from the Get, Let and Set of one property and the
//! members of enumerations; a second declaration is reported where it stands, with the name as
//! it was last written. A variable declared without a type takes the type its first letter has
//! under the module's `Def<Type>` statements.

mod common;
use common::*;
use vb6_syntax::frontend::declared_names::{DeclaredType, NameKind};

const AMBIGUOUS: u32 = 0x802c;

/// The ambiguity reports of `src` as `(offset, length, message)`.
fn ambiguities(src: &str) -> Vec<(u32, u32, String)> {
    let p = parse(src);
    p.found
        .iter()
        .zip(&p.messages)
        .filter(|((code, _, _), _)| *code == AMBIGUOUS)
        .map(|((_, at, len), m)| (*at, *len, m.clone()))
        .collect()
}

#[test]
fn a_second_procedure_of_the_same_name_is_reported_at_its_name_in_the_case_it_was_written() {
    let src = "Sub A()\nEnd Sub\nSub a()\nEnd Sub\n";
    assert_eq!(ambiguities(src), vec![(20, 1, "Ambiguous name detected: a".to_string())]);
    let p = parse(src);
    assert_eq!(p.found, vec![(AMBIGUOUS, 20, 1)]);
}

#[test]
fn a_procedure_and_a_variable_of_one_name_clash_wherever_the_variable_stands() {
    assert_eq!(ambiguities("Dim x As Long\nSub X()\nEnd Sub\n"), vec![(4, 1, "Ambiguous name detected: X".to_string())]);
    assert_eq!(ambiguities("Const c = 1\nDim C As Long\n"), vec![(16, 1, "Ambiguous name detected: C".to_string())]);
}

#[test]
fn the_get_let_and_set_of_one_property_may_share_a_name() {
    let src = "Property Get P() As Long\nEnd Property\nProperty Let P(v As Long)\nEnd Property\n";
    assert_eq!(parse(src).found, vec![]);
    let src = "Property Get P() As Long\nEnd Property\nProperty Set P(v As Object)\nEnd Property\nProperty Let P(v As Long)\nEnd Property\n";
    assert_eq!(parse(src).found, vec![]);
}

#[test]
fn two_gets_of_one_property_are_ambiguous() {
    let src = "Property Get P() As Long\nEnd Property\nProperty Get P() As Long\nEnd Property\n";
    assert_eq!(ambiguities(src), vec![(51, 1, "Ambiguous name detected: P".to_string())]);
}

#[test]
fn a_repeated_enumeration_member_is_reported_at_the_first_of_its_declarations() {
    let src = "Enum E\nm\nM\nEnd Enum\n";
    assert_eq!(ambiguities(src), vec![(7, 1, "Ambiguous name detected: M".to_string())]);
}

#[test]
fn enumeration_members_of_different_enumerations_may_share_a_name() {
    assert_eq!(parse("Enum E\nm\nEnd Enum\nEnum F\nm\nEnd Enum\n").found, vec![]);
}

#[test]
fn a_repeated_member_of_a_type_is_reported_at_the_first_of_its_declarations() {
    let src = "Type T\na As Long\nA As Long\nEnd Type\n";
    assert_eq!(ambiguities(src), vec![(7, 1, "Ambiguous name detected: A".to_string())]);
}

#[test]
fn a_repeated_event_is_reported_at_the_second_declaration() {
    let p = parse_as("Public Event Ev()\nPublic Event ev()\n", vb6_syntax::frontend::parser::ModuleKind::Class);
    assert_eq!(p.found, vec![(AMBIGUOUS, 31, 2)]);
    assert_eq!(p.messages, vec!["Ambiguous name detected: ev".to_string()]);
}

#[test]
fn an_untyped_variable_takes_the_type_its_first_letter_has_under_def_statements() {
    let type_of = |src: &str, name: &str| {
        let p = parse(src);
        assert_eq!(p.found, vec![], "{src:?}");
        let sym = p.names.entries.iter().find(|e| p.name(e.sym).eq_ignore_ascii_case(name)).expect("the name is declared");
        assert_eq!(sym.kind, NameKind::Variable, "{src:?}");
        sym.declared_type
    };
    assert_eq!(type_of("DefInt a-z\nDim x\n", "x"), Some(DeclaredType::Integer));
    assert_eq!(type_of("DefStr s\nDim sx\nDim ax\n", "sx"), Some(DeclaredType::String));
    assert_eq!(type_of("DefStr s\nDim sx\nDim ax\n", "ax"), Some(DeclaredType::Variant));
    assert_eq!(type_of("DefLng a-c, x\nDim bb\n", "bb"), Some(DeclaredType::Long));
    assert_eq!(type_of("DefLng a-c, x\nDim x\n", "x"), Some(DeclaredType::Long));
    assert_eq!(type_of("DefLng a-c, x\nDim d\n", "d"), Some(DeclaredType::Variant));
    assert_eq!(type_of("DefInt a-z\nDim x As Double\n", "x"), Some(DeclaredType::Double));
    assert_eq!(type_of("Dim x\n", "x"), Some(DeclaredType::Variant));
}
