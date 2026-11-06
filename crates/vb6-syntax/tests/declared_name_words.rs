//! Which words may be declared as names: the module-header words are reserved like any
//! other, a keyword spelled with a trailing `$` is its word plus the string character, a name
//! in brackets is refused only where the first name of a declaration is wanted, and a word
//! cannot begin with an underscore.

mod common;
use common::{offset_of, parse, parse_as};
use vb6_syntax::frontend::ast::{AstLit, ExprNode, ProcKind};
use vb6_syntax::frontend::parser::ModuleKind;
use vb6_syntax::frontend::token::TypeSuffix;

const EXPECTED: u32 = 0x9c6f;
const SYNTAX_ERROR: u32 = 0x9c70;
const INVALID_CHARACTER: u32 = 0x9c6c;
const TYPE_CHARACTER_MISMATCH: u32 = 0x9c86;
const INVALID_OUTSIDE_PROCEDURE: u32 = 0x9c73;

const HEADER_WORDS: [&str; 25] = [
    "Attribute",
    "VB_Name",
    "VB_GlobalNameSpace",
    "VB_Creatable",
    "VB_PredeclaredId",
    "VB_Exposed",
    "VB_Description",
    "VB_Ext_KEY",
    "VB_HelpID",
    "VB_Invoke_Property",
    "VB_Invoke_PropertyPut",
    "VB_Invoke_PropertyPutRef",
    "VB_MemberFlags",
    "VB_ProcData",
    "VB_UserMemId",
    "VB_VarDescription",
    "VB_VarHelpID",
    "VB_VarMemberFlags",
    "VB_VarProcData",
    "VB_VarUserMemId",
    "VB_Base",
    "VB_Control",
    "VB_Customizable",
    "VB_TemplateDerived",
    "VB_Invoke_Func",
];

/// The first diagnostic of `src` as `(code, start, message)`.
fn first_problem(src: &str, kind: ModuleKind) -> (u32, u32, String) {
    let p = parse_as(src, kind);
    assert!(!p.found.is_empty(), "no diagnostic for {src:?}");
    (p.found[0].0, p.found[0].1, p.messages[0].clone())
}

fn class_source(body: &str) -> String {
    format!(
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"Class1\"\n\
         Attribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\n\
         Attribute VB_PredeclaredId = False\nAttribute VB_Exposed = False\n{body}"
    )
}

// ── Header words ─────────────────────────────────────────────────────────────

#[test]
fn a_header_word_cannot_name_a_declaration_at_module_level() {
    for word in HEADER_WORDS {
        for template in [
            "Sub {k}()\nEnd Sub\n",
            "Function {k}() As Long\nEnd Function\n",
            "Sub S({k} As Long)\nEnd Sub\n",
            "Sub S(a As Long, {k} As Long)\nEnd Sub\n",
            "Dim {k} As Long\n",
            "Enum {k}\nA\nEnd Enum\n",
            "Type {k}\na As Long\nEnd Type\n",
            "Const {k} = 1\n",
        ] {
            let src = template.replace("{k}", word);
            let (code, at, message) = first_problem(&src, ModuleKind::Standard);
            assert_eq!(code, EXPECTED, "for {src:?}");
            assert_eq!(at, offset_of(&src, word, 0), "for {src:?}");
            assert_eq!(message, "Expected: identifier", "for {src:?}");
        }
    }
}

#[test]
fn a_header_word_cannot_name_an_event() {
    for word in HEADER_WORDS {
        let src = class_source(&format!("Event {word}()\n"));
        let (code, at, message) = first_problem(&src, ModuleKind::Class);
        assert_eq!(code, EXPECTED, "for {word}");
        assert_eq!(at, offset_of(&src, &format!("Event {word}"), 0) + 6, "for {word}");
        assert_eq!(message, "Expected: identifier", "for {word}");
    }
}

#[test]
fn a_header_word_cannot_name_anything_inside_a_procedure() {
    for word in HEADER_WORDS {
        // `Attribute` alone at the start of a statement is read as a name there: a label or
        // the target of an assignment.
        let leads_statement = word != "Attribute";
        for template in [
            "Dim {k} As Long",
            "{k}:",
            "{k} = 1",
            "Call {k}",
            "Set {k} = Nothing",
            "ReDim {k}(3)",
            "GoTo {k}",
            "For {k} = 1 To 2\nNext",
            "For Each {k} In x\nNext",
        ] {
            if !leads_statement && (template == "{k}:" || template == "{k} = 1") {
                continue;
            }
            let src = format!("Sub S()\n{}\nEnd Sub\n", template.replace("{k}", word));
            let (code, at, message) = first_problem(&src, ModuleKind::Standard);
            assert_eq!(code, SYNTAX_ERROR, "for {src:?}");
            assert_eq!(at, offset_of(&src, word, 0), "for {src:?}");
            assert_eq!(message, "Syntax error", "for {src:?}");
        }
    }
}

#[test]
fn a_header_word_is_still_a_member_name() {
    for word in HEADER_WORDS {
        let src = format!("Sub S()\nx.{word} = 1\ny = x.{word}\nWith x\n.{word} = 2\nEnd With\nEnd Sub\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![], "for {word}");
        let body = p.proc_body(0);
        assert_eq!(body.len(), 3, "for {word}");
        let ExprNode::Assign { target, .. } = p.node(body[0]) else { panic!("expected an assignment") };
        let ExprNode::MemberAccess { member, bang: false, .. } = p.node(*target) else { panic!("expected a member") };
        assert_eq!(p.name(*member), word);
    }
}

#[test]
fn header_lines_still_read_as_attributes() {
    let p = parse("Attribute VB_Name = \"Module1\"\nSub S()\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::Attribute { target: None, name, values } = p.node(p.top[0]) else { panic!("expected Attribute") };
    assert_eq!(p.name(*name), "VB_Name");
    assert_eq!(values.len(), 1);
    assert!(matches!(p.node(values[0]), ExprNode::Literal { lit: AstLit::Str(s) } if &**s == "Module1"));
}

// ── Keywords spelled with a trailing dollar ─────────────────────────────────

const PLAIN_DOLLAR_WORDS: [&str; 6] = ["Mid$", "MidB$", "Dir$", "Error$", "Format$", "CurDir$"];
const RESERVED_DOLLAR_WORDS: [&str; 4] = ["Date$", "String$", "Input$", "InputB$"];

#[test]
fn a_dollar_word_cannot_name_a_sub_or_an_event_or_a_declared_sub() {
    for word in PLAIN_DOLLAR_WORDS {
        for template in ["Sub {k}()\nEnd Sub\n", "Declare Sub {k} Lib \"k\" ()\n"] {
            let src = template.replace("{k}", word);
            let (code, at, message) = first_problem(&src, ModuleKind::Standard);
            assert_eq!(code, EXPECTED, "for {src:?}");
            assert_eq!(at, offset_of(&src, word, 0), "for {src:?}");
            assert_eq!(message, "Expected: Identifier without type suffix", "for {src:?}");
        }
        let src = class_source(&format!("Event {word}()\n"));
        let (code, at, message) = first_problem(&src, ModuleKind::Class);
        assert_eq!(code, EXPECTED, "for {word}");
        assert_eq!(at, offset_of(&src, &format!("Event {word}"), 0) + 6, "for {word}");
        assert_eq!(message, "Expected: Identifier without type suffix", "for {word}");
    }
}

#[test]
fn a_dollar_word_function_is_a_string_function_that_takes_no_type() {
    for word in PLAIN_DOLLAR_WORDS {
        let src = format!("Function {word}()\nEnd Function\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![], "for {word}");
        let ExprNode::ProcDecl { kind: ProcKind::Function, ret_type: Some(t), .. } = p.node(p.top[0]) else {
            panic!("expected a function with a return type")
        };
        assert!(matches!(p.node(*t), ExprNode::StringType { fixed_len: None }), "for {word}");

        let src = format!("Function {word}() As Long\nEnd Function\n");
        let (code, at, message) = first_problem(&src, ModuleKind::Standard);
        assert_eq!(code, EXPECTED, "for {word}");
        assert_eq!(at, offset_of(&src, word, 0), "for {word}");
        assert_eq!(message, "Expected: end of statement", "for {word}");
    }
}

#[test]
fn a_reserved_dollar_word_is_a_reserved_word_like_any_other() {
    for word in RESERVED_DOLLAR_WORDS {
        for template in ["Sub {k}()\nEnd Sub\n", "Function {k}() As Long\nEnd Function\n", "Dim {k} As Long\n", "Const {k} = 1\n"] {
            let src = template.replace("{k}", word);
            let (code, at, message) = first_problem(&src, ModuleKind::Standard);
            assert_eq!(code, EXPECTED, "for {src:?}");
            assert_eq!(at, offset_of(&src, word, 0), "for {src:?}");
            assert_eq!(message, "Expected: identifier", "for {src:?}");
        }
    }
}

#[test]
fn a_dollar_word_still_names_variables_constants_parameters_and_types() {
    for word in ["Mid$", "MidB$", "Format$", "CurDir$", "Error$"] {
        for template in [
            "Dim {k} As Long\n",
            "Dim a As Long, {k} As Long\n",
            "Const {k} = 1\n",
            "Sub S({k} As Long)\nEnd Sub\n",
            "Property Get {k}() As Long\nEnd Property\n",
            "Type {k}\na As Long\nEnd Type\n",
            "Enum {k}\nA\nEnd Enum\n",
            "Type T\n{k} As Long\nEnd Type\n",
            "Sub S()\nDim {k} As Long\nEnd Sub\n",
        ] {
            let src = template.replace("{k}", word);
            assert_eq!(parse(&src).found, vec![], "for {src:?}");
        }
    }
}

#[test]
fn an_enumeration_member_with_a_value_takes_no_type_character_but_the_long_one() {
    for (name, rejected) in [
        ("Foo%", true),
        ("Foo&", false),
        ("Foo!", true),
        ("Foo#", true),
        ("Foo@", true),
        ("Foo$", true),
        ("Mid$", true),
        ("Error$", true),
        ("Date$", true),
        ("Foo", false),
    ] {
        let src = format!("Enum E\n{name} = 1\nEnd Enum\n");
        let p = parse(&src);
        if rejected {
            assert_eq!(p.found, vec![(TYPE_CHARACTER_MISMATCH, offset_of(&src, name, 0), name.len() as u32)], "for {name}");
            assert_eq!(p.messages, vec!["Type-declaration character does not match declared data type".to_string()]);
        } else {
            assert_eq!(p.found, vec![], "for {name}");
        }
        // Without a value the member has no type to disagree with; `Mid$` alone is the start of
        // the `Mid$` statement, which needs its brackets.
        let src = format!("Enum E\n{name}\nEnd Enum\n");
        if name == "Mid$" {
            assert!(!parse(&src).found.is_empty(), "for {name} alone");
        } else {
            assert_eq!(parse(&src).found, vec![], "for {name} alone");
        }
    }
}

#[test]
fn an_enumeration_member_keeps_its_name_and_value() {
    let p = parse("Enum E\nFoo& = 7\nBar\nEnd Enum\n");
    assert_eq!(p.found, vec![]);
    let ExprNode::EnumDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Enum") };
    assert_eq!(members.len(), 2);
    let ExprNode::DimItem { name, init: Some(init), type_node, .. } = p.node(members[0]) else { panic!("expected a member") };
    assert_eq!(p.name(*name), "Foo");
    assert!(matches!(p.node(*init), ExprNode::Literal { lit: AstLit::Int(7) }));
    assert!(matches!(type_node.map(|t| p.node(t)), Some(ExprNode::BuiltinType { kind: 3 })));
    let ExprNode::DimItem { name, init: None, .. } = p.node(members[1]) else { panic!("expected a member") };
    assert_eq!(p.name(*name), "Bar");
}

#[test]
fn a_type_member_that_begins_an_assignment_cannot_be_a_member() {
    for word in ["Me", "String$"] {
        let src = format!("Type T\n{word} As Long\nEnd Type\n");
        let (code, at, message) = first_problem(&src, ModuleKind::Standard);
        assert_eq!(code, EXPECTED, "for {word}");
        assert_eq!(at, offset_of(&src, word, 0), "for {word}");
        assert_eq!(message, "Expected: =", "for {word}");
    }
    // Every other reserved word, and the other dollar words, may name a member.
    for word in ["Static", "Abs", "Date", "Integer", "String", "Date$", "Input$", "Mid$", "Attribute", "VB_Name"] {
        let src = format!("Type T\n{word} As Long\nEnd Type\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![], "for {word}");
        let ExprNode::TypeDecl { members, .. } = p.node(p.top[0]) else { panic!("expected Type") };
        assert_eq!(members.len(), 1, "for {word}");
    }
}

// ── Static outside a procedure ────────────────────────────────────────────────

#[test]
fn static_outside_a_procedure_reads_the_variable_list_first() {
    // A list that reads well is refused as a whole, at `Static`.
    for (src, name_at) in [
        ("Static x As Long\n", None),
        ("Static x\n", None),
        ("Static Mid$ As Long\n", None),
        ("Static Object As Long\n", None),
        ("Static a As Long, b As Long\n", None),
        ("Static Foo% As Long\n", Some("Foo%")),
    ] {
        let (code, at, message) = first_problem(src, ModuleKind::Standard);
        match name_at {
            None => {
                assert_eq!((code, at, message.as_str()), (INVALID_OUTSIDE_PROCEDURE, 0, "Invalid outside procedure"), "for {src:?}");
            }
            Some(_) => {
                // A type character and `As` together are wrong before the list is judged.
                assert_eq!(code, EXPECTED, "for {src:?}");
                assert_eq!(message, "Expected: end of statement", "for {src:?}");
            }
        }
    }
    // What cannot be a name is reported as such.
    for src in [
        "Static Const x = 1\n",
        "Static Dim x As Long\n",
        "Static Enum E\nA\nEnd Enum\n",
        "Static Type T\na As Long\nEnd Type\n",
        "Static Declare Sub F Lib \"k\" ()\n",
        "Static Event F()\n",
        "Static If\n",
        "Static 1\n",
        "Static\n",
        "Static [x] As Long\n",
        "Static Date$ As Long\n",
        "Static Abs As Long\n",
    ] {
        let (code, at, message) = first_problem(src, ModuleKind::Standard);
        assert_eq!(code, EXPECTED, "for {src:?}");
        assert_eq!(message, "Expected: identifier", "for {src:?}");
        assert!(at >= 7 || src == "Static\n", "for {src:?}");
    }
    // `Static` still qualifies a procedure.
    let p = parse("Static Sub S()\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}

// ── Bracketed names ───────────────────────────────────────────────────────────

#[test]
fn an_enumeration_may_be_named_with_a_bracketed_word() {
    for (word, name) in [("[Foo]", "Foo"), ("[Date]", "Date"), ("[Foo Bar]", "Foo Bar"), ("[Integer]", "Integer"), ("[Error]", "Error")] {
        for access in ["", "Public ", "Private "] {
            let p = parse(&format!("{access}Enum {word}\nA\nEnd Enum\n"));
            assert_eq!(p.found, vec![], "for {access}{word}");
            let ExprNode::EnumDecl { name: n, members } = p.node(p.top[0]) else { panic!("expected Enum") };
            assert_eq!(p.name(*n), name);
            assert_eq!(members.len(), 1);
        }
    }
}

#[test]
fn a_type_character_after_a_bracketed_enumeration_name_is_an_invalid_character() {
    let src = "Enum [Foo]%\nA\nEnd Enum\n";
    let p = parse(src);
    assert_eq!(p.messages[0], "Invalid character");
    assert_eq!(p.found[0], (INVALID_CHARACTER, offset_of(src, "%", 0), 1));
}

#[test]
fn only_the_first_name_of_a_variable_list_refuses_brackets() {
    let src = "Public [a] As Long\n";
    let (code, at, message) = first_problem(src, ModuleKind::Standard);
    assert_eq!((code, at, message.as_str()), (EXPECTED, offset_of(src, "[a]", 0), "Expected: identifier"));

    for src in ["Public a1 As Long, [a2] As Long\n", "Dim a1 As Long, [a2] As Long\n", "Private a1, [Date] As Long, [Foo Bar]\n"] {
        let p = parse(src);
        assert_eq!(p.found, vec![], "for {src:?}");
        let ExprNode::Block { stmts } = p.node(p.top[0]) else { panic!("expected a declaration list") };
        assert!(stmts.len() >= 2);
    }
    let p = parse("Dim a1 As Long, [a2] As Long\n");
    let ExprNode::Block { stmts } = p.node(p.top[0]) else { panic!("expected a declaration list") };
    let ExprNode::DimItem { name, .. } = p.node(stmts[1]) else { panic!("expected an item") };
    assert_eq!(p.name(*name), "a2");

    // A constant list refuses brackets at every name.
    let src = "Const a1 = 1, [a2] = 2\n";
    let (code, at, message) = first_problem(src, ModuleKind::Standard);
    assert_eq!((code, at, message.as_str()), (EXPECTED, offset_of(src, "[a2]", 0), "Expected: identifier"));
}

#[test]
fn a_reserved_word_after_a_comma_of_a_variable_list_wants_a_variable() {
    for (src, word) in [("Dim a As Long, Abs As Long\n", "Abs"), ("Public a As Long, Rem As Long\n", "Rem")] {
        let (code, at, message) = first_problem(src, ModuleKind::Standard);
        assert_eq!((code, at, message.as_str()), (EXPECTED, offset_of(src, word, 0), "Expected: variable"), "for {src:?}");
    }
    let src = "Sub S()\nDim a As Long, Abs As Long\nEnd Sub\n";
    let (code, at, message) = first_problem(src, ModuleKind::Standard);
    assert_eq!((code, at, message.as_str()), (SYNTAX_ERROR, offset_of(src, "Abs", 0), "Syntax error"));
}

// ── Qualified type names ────────────────────────────────────────────────────

#[test]
fn rem_cannot_follow_the_dot_of_a_type_name() {
    for template in [
        "Private o As T.Rem\n",
        "Private o As New T.Rem\n",
        "Private Sub P(o As T.Rem)\nEnd Sub\n",
        "Private Function F() As T.Rem\nEnd Function\n",
    ] {
        let (code, at, message) = first_problem(template, ModuleKind::Standard);
        assert_eq!(code, EXPECTED, "for {template:?}");
        assert_eq!(at, offset_of(template, "Rem", 0), "for {template:?}");
        assert_eq!(message, "Expected: identifier", "for {template:?}");
    }
    // Any other word, a reserved one included, names a type there.
    for word in ["If", "Foo", "Date", "Dim"] {
        let src = format!("Private o As T.{word}\n");
        let p = parse(&src);
        assert_eq!(p.found, vec![], "for {word}");
        let ExprNode::Block { stmts } = p.node(p.top[0]) else { panic!("expected a declaration") };
        let ExprNode::DimItem { type_node: Some(t), .. } = p.node(stmts[0]) else { panic!("expected an item") };
        let ExprNode::UserType { name, child: Some(c) } = p.node(*t) else { panic!("expected a qualified type") };
        assert_eq!(p.name(*name), "T");
        let ExprNode::NameRef { sym, suffix: TypeSuffix::None } = p.node(*c) else { panic!("expected the type name") };
        assert_eq!(p.name(*sym).to_ascii_lowercase(), word.to_ascii_lowercase());
    }
}

// ── Underscore ─────────────────────────────────────────────────────────────────

#[test]
fn a_word_cannot_begin_with_an_underscore() {
    for template in [
        "Public Sub _Foo()\nEnd Sub\n",
        "Public Sub P(_a As Long)\nEnd Sub\n",
        "Enum _E\nA\nEnd Enum\n",
        "Enum E\n_a\nEnd Enum\n",
        "Type T\n_a As Long\nEnd Type\n",
        "Const _a = 1\n",
        "Declare Sub _a Lib \"k\" ()\n",
    ] {
        let (code, at, message) = first_problem(template, ModuleKind::Standard);
        assert_eq!(code, INVALID_CHARACTER, "for {template:?}");
        assert_eq!(at, offset_of(template, "_", 0), "for {template:?}");
        assert_eq!(message, "Invalid character", "for {template:?}");
    }
    // Inside a procedure the general syntax error is all there is.
    for template in ["Dim _x As Long", "x._y = 1", "GoTo _x"] {
        let src = format!("Sub S()\n{template}\nEnd Sub\n");
        let (code, at, message) = first_problem(&src, ModuleKind::Standard);
        assert_eq!(code, SYNTAX_ERROR, "for {template:?}");
        assert_eq!(at, offset_of(&src, "_", 0), "for {template:?}");
        assert_eq!(message, "Syntax error", "for {template:?}");
    }
    // A name may contain and end with underscores.
    let p = parse("Public Sub a_b_()\nEnd Sub\n");
    assert_eq!(p.found, vec![]);
}
