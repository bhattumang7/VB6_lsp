//! Parser acceptance tests restored from the ANTLR grammar example corpus.
//!
//! The ANTLR examples were real VB6 snippets used to verify the old
//! tree-sitter grammar against actual VB6 behaviour.  They were deleted when
//! the project moved to a Rust-based parser but their coverage was not carried
//! over.  Each test below corresponds to one or more of those example files and
//! asserts that the Rust parser accepts the same construct without errors *and*
//! that the expected AST node shape was produced.

use vb6_core::frontend::ast::{
    AstLit, BinOpKind, DoKind, ExprArena, ExprNode, FileIoKind, LabelRef, NodeId, OnErrorKind,
    ProcKind, UnOpKind,
};

// ── structural assertion helpers ────────────────────────────────────────────────

/// Return a reference to the node at `id`.
fn at(arena: &ExprArena, id: NodeId) -> &ExprNode {
    arena.get(id)
}

/// Assert the node at `id` is a `Block` and return its statement ids.
fn block_stmts<'a>(arena: &'a ExprArena, id: NodeId) -> &'a [NodeId] {
    match arena.get(id) {
        ExprNode::Block { stmts } => stmts,
        other => panic!("expected Block, got {}", kind_name(other)),
    }
}

/// Assert the node at `id` is an `ArgList` and return its arg ids.
fn arglist_args<'a>(arena: &'a ExprArena, id: NodeId) -> &'a [NodeId] {
    match arena.get(id) {
        ExprNode::ArgList { args } => args,
        other => panic!("expected ArgList, got {}", kind_name(other)),
    }
}

/// A short discriminant name for a node, for assertion messages.
fn kind_name(n: &ExprNode) -> &'static str {
    use ExprNode::*;
    match n {
        Generic { .. } => "Generic",
        ForRange { .. } => "ForRange",
        TypeSpec { .. } => "TypeSpec",
        UdtTypeSpec { .. } => "UdtTypeSpec",
        Literal { .. } => "Literal",
        NameRef { .. } => "NameRef",
        Me => "Me",
        Nothing => "Nothing",
        WithContext => "WithContext",
        Paren { .. } => "Paren",
        BinOp { .. } => "BinOp",
        UnOp { .. } => "UnOp",
        MemberAccess { .. } => "MemberAccess",
        AddressOf { .. } => "AddressOf",
        Call { .. } => "Call",
        Assign { .. } => "Assign",
        SetAssign { .. } => "SetAssign",
        RangeAssign { .. } => "RangeAssign",
        MidAssign { .. } => "MidAssign",
        ProcDecl { .. } => "ProcDecl",
        Block { .. } => "Block",
        ArgList { .. } => "ArgList",
        If { .. } => "If",
        While { .. } => "While",
        Do { .. } => "Do",
        For { .. } => "For",
        ForEach { .. } => "ForEach",
        With { .. } => "With",
        SelectCase { .. } => "SelectCase",
        CaseBlock { .. } => "CaseBlock",
        CaseElse { .. } => "CaseElse",
        GoTo { .. } => "GoTo",
        GoSub { .. } => "GoSub",
        ReturnStmt => "ReturnStmt",
        OnError { .. } => "OnError",
        Resume { .. } => "Resume",
        Stop => "Stop",
        EndStmt => "EndStmt",
        OnGo { .. } => "OnGo",
        ExitStmt { .. } => "ExitStmt",
        CallStmt { .. } => "CallStmt",
        Erase { .. } => "Erase",
        RaiseEvent { .. } => "RaiseEvent",
        DebugPrint { .. } => "DebugPrint",
        ErrorStmt { .. } => "ErrorStmt",
        DimItem { .. } => "DimItem",
        ReDimItem { .. } => "ReDimItem",
        ParamDef { .. } => "ParamDef",
        TypeDecl { .. } => "TypeDecl",
        EnumDecl { .. } => "EnumDecl",
        EventDecl { .. } => "EventDecl",
        Implements { .. } => "Implements",
        BuiltinType { .. } => "BuiltinType",
        StringType { .. } => "StringType",
        UserType { .. } => "UserType",
        RangeTo { .. } => "RangeTo",
        TypeOf { .. } => "TypeOf",
        New { .. } => "New",
        CaseIs { .. } => "CaseIs",
        MissingArg => "MissingArg",
        NamedArg { .. } => "NamedArg",
        OptionExplicit => "OptionExplicit",
        OptionBase { .. } => "OptionBase",
        OptionCompare { .. } => "OptionCompare",
        DeclareDecl { .. } => "DeclareDecl",
        FileIoStmt { .. } => "FileIoStmt",
        Label { .. } => "Label",
        DefType { .. } => "DefType",
        Attribute { .. } => "Attribute",
    }
}
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;

// ── helpers ───────────────────────────────────────────────────────────────────

fn ctx() -> ScannerContext {
    let mut c = ScannerContext::new(1, 1, 0x0409);
    c.intern_keywords();
    c
}

fn parse_module(src: &[u8]) -> (Vec<NodeId>, ExprArena, bool) {
    let mut c = ctx();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut c, src);
    let nodes = parser.parse_module(&mut arena);
    let has_errors = parser.diagnostics.has_errors();
    (nodes, arena, has_errors)
}

fn parse_class_module(src: &[u8]) -> (Vec<NodeId>, ExprArena, bool) {
    let mut c = ctx();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut c, src, ModuleKind::Class);
    let nodes = parser.parse_module(&mut arena);
    let has_errors = parser.diagnostics.has_errors();
    (nodes, arena, has_errors)
}

/// Wrap bare statements in a Sub so the module parser sees them.
fn parse_stmts(src: &[u8]) -> (Vec<NodeId>, ExprArena, bool) {
    let wrapped = format!("Sub Test()\n{}\nEnd Sub", std::str::from_utf8(src).unwrap());
    parse_module(wrapped.as_bytes())
}

fn find_node<'a>(arena: &'a ExprArena, pred: impl Fn(&ExprNode) -> bool) -> Option<&'a ExprNode> {
    (0..arena.len())
        .map(|i| arena.get(NodeId(i as u32)))
        .find(|n| pred(n))
}

fn count_nodes(arena: &ExprArena, pred: impl Fn(&ExprNode) -> bool) -> usize {
    (0..arena.len())
        .filter(|&i| pred(arena.get(NodeId(i as u32))))
        .count()
}

/// Find the single `FileIoStmt` of the given kind and return
/// `(channel_is_some, args_len)`. Panics unless exactly one such node exists.
fn fileio(arena: &ExprArena, want: FileIoKind) -> (bool, usize) {
    let n = count_nodes(arena, |n| matches!(n, ExprNode::FileIoStmt { kind, .. } if *kind == want));
    assert_eq!(n, 1, "expected exactly one FileIoStmt of the requested kind, found {n}");
    match find_node(arena, |n| matches!(n, ExprNode::FileIoStmt { kind, .. } if *kind == want)).unwrap() {
        ExprNode::FileIoStmt { channel, args, .. } => (channel.is_some(), args.len()),
        _ => unreachable!(),
    }
}

// ── Intrinsic statements: Beep, AppActivate, Shell, Randomize, Unload, Error ─

#[test]
fn beep_accepts() {
    let (top, arena, errors) = parse_stmts(b"Beep");
    assert!(!errors, "Beep must parse without errors");
    // Beep with no arguments is a bare name-reference at statement level — and
    // exactly that: one NameRef, no spurious CallStmt and no extra nodes.
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::NameRef { .. })),
        1,
        "Beep must produce exactly one NameRef"
    );
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })),
        0,
        "zero-arg Beep must NOT be wrapped in a CallStmt"
    );
    // The Sub body is a single statement: the Beep NameRef.
    let body = match at(&arena, top[0]) {
        ExprNode::ProcDecl { body, .. } => *body,
        other => panic!("expected ProcDecl, got {}", kind_name(other)),
    };
    let stmts = block_stmts(&arena, body);
    assert_eq!(stmts.len(), 1, "Beep body must have exactly one statement");
    assert!(matches!(at(&arena, stmts[0]), ExprNode::NameRef { .. }), "the lone statement is the Beep NameRef");
}

#[test]
fn beep_in_for_loop() {
    let (_, arena, errors) = parse_stmts(b"For J = 1 To 2\n  Beep\nNext J");
    assert!(!errors, "Beep inside For/Next must parse without errors");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::For { .. })),
        1,
        "exactly one For node"
    );
    let for_node = find_node(&arena, |n| matches!(n, ExprNode::For { .. })).unwrap();
    let (start, end, step, body) = match for_node {
        ExprNode::For { start, end, step, body, .. } => (*start, *end, *step, *body),
        _ => unreachable!(),
    };
    assert!(step.is_none(), "no Step clause → step must be None");
    assert!(matches!(at(&arena, start), ExprNode::Literal { lit: AstLit::Int(1) }), "start = 1");
    assert!(matches!(at(&arena, end), ExprNode::Literal { lit: AstLit::Int(2) }), "end = 2");
    // The body must actually contain the Beep (a NameRef), not be empty.
    let body_stmts = block_stmts(&arena, body);
    assert_eq!(body_stmts.len(), 1, "For body has exactly one statement (Beep)");
    assert!(matches!(at(&arena, body_stmts[0]), ExprNode::NameRef { .. }), "For body statement is the Beep NameRef");
}

#[test]
fn app_activate_string_literal() {
    let (_, arena, errors) = parse_stmts(b"AppActivate \"Microsoft Word\"");
    assert!(!errors, "AppActivate with string literal must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let (callee, args) = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { callee, args } => (*callee, *args),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, callee), ExprNode::NameRef { .. }), "callee is the AppActivate name");
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the title string literal");
}

#[test]
fn app_activate_with_wait_flag() {
    let (_, arena, errors) = parse_stmts(b"AppActivate \"Visual Basic\", 1");
    assert!(!errors, "AppActivate with wait-flag argument must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 2, "title + wait-flag → 2 args");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg0 is the title string");
    assert!(matches!(at(&arena, args[1]), ExprNode::Literal { lit: AstLit::Int(1) }), "arg1 is the wait flag 1");
}

#[test]
fn shell_statement() {
    let (_, arena, errors) = parse_stmts(b"Shell \"C:\\WINDOWS\\NOTEPAD.EXE\", 1");
    assert!(!errors, "Shell must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 2, "path + window-style → 2 args");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg0 is the path string");
    assert!(matches!(at(&arena, args[1]), ExprNode::Literal { lit: AstLit::Int(1) }), "arg1 is window style 1");
}

#[test]
fn randomize_no_arg() {
    let (_, arena, errors) = parse_stmts(b"Randomize");
    assert!(!errors, "Randomize without argument must parse without errors");
    // Randomize with no args is a bare name-reference at statement level — exactly
    // one NameRef and no CallStmt.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::NameRef { .. })), 1, "exactly one NameRef");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 0, "no CallStmt for zero-arg Randomize");
}

#[test]
fn unload_me() {
    let (_, arena, errors) = parse_stmts(b"Unload Me");
    assert!(!errors, "Unload Me must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let (callee, args) = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { callee, args } => (*callee, *args),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, callee), ExprNode::NameRef { .. }), "callee is the Unload name");
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "Unload Me has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Me), "the argument is the Me keyword");
}

#[test]
fn error_number_statement() {
    let (_, arena, errors) = parse_stmts(b"Error 12");
    assert!(!errors, "Error <number> statement must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ErrorStmt { .. })), 1, "one ErrorStmt");
    let expr = match find_node(&arena, |n| matches!(n, ExprNode::ErrorStmt { .. })).unwrap() {
        ExprNode::ErrorStmt { expr } => *expr,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, expr), ExprNode::Literal { lit: AstLit::Int(12) }), "Error operand is the literal 12");
}

// ── File-system statements: ChDir, ChDrive, MkDir, RmDir, Kill, FileCopy, Name

#[test]
fn chdir_string_literal() {
    let (_, arena, errors) = parse_stmts(b"ChDir \"D:\\TMP\"");
    assert!(!errors, "ChDir with string literal must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "ChDir has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the path string");
}

#[test]
fn chdir_variable() {
    let (_, arena, errors) = parse_stmts(b"Dim Path As String\nPath = \"C:\\Tmp\"\nChDir Path");
    assert!(!errors, "ChDir with variable must parse without errors");
    // One assignment (Path = "...") and one ChDir call.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt (ChDir Path)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "one Assign (Path = ...)");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "ChDir Path has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::NameRef { .. }), "arg is the Path variable reference");
}

#[test]
fn chdir_concat_expr() {
    let (_, arena, errors) = parse_stmts(b"ChDir \"c:\\\" & \"Tmp\"");
    assert!(!errors, "ChDir with string-concat expression must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::BinOp { op: BinOpKind::Cat, .. })),
        1,
        "exactly one BinOp::Cat for the concat argument"
    );
    // The single CallStmt arg must BE the Cat node whose operands are both strings.
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "ChDir with one concat expr → one argument");
    let (op, l, r) = match at(&arena, args[0]) {
        ExprNode::BinOp { op, lhs, rhs } => (*op, *lhs, *rhs),
        other => panic!("ChDir arg must be the Cat BinOp, got {}", kind_name(other)),
    };
    assert_eq!(op, BinOpKind::Cat, "operator is string concat (&)");
    assert!(matches!(at(&arena, l), ExprNode::Literal { lit: AstLit::Str(_) }), "lhs is a string literal");
    assert!(matches!(at(&arena, r), ExprNode::Literal { lit: AstLit::Str(_) }), "rhs is a string literal");
}

#[test]
fn chdrive_literal() {
    let (_, arena, errors) = parse_stmts(b"ChDrive \"C\"");
    assert!(!errors, "ChDrive with string literal must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "ChDrive has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the drive string");
}

#[test]
fn chdrive_variable() {
    let (_, arena, errors) = parse_stmts(b"Dim D\nChDrive D");
    assert!(!errors, "ChDrive with variable must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "ChDrive D has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::NameRef { .. }), "arg is the D variable reference");
}

#[test]
fn mkdir_literal() {
    let (_, arena, errors) = parse_stmts(b"MkDir \"c:/tmp\"");
    assert!(!errors, "MkDir must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "MkDir has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the path string");
}

#[test]
fn rmdir_literal() {
    let (_, arena, errors) = parse_stmts(b"RmDir \"c:/tmp\"");
    assert!(!errors, "RmDir must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "RmDir has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the path string");
}

#[test]
fn kill_string_literal() {
    let (_, arena, errors) = parse_stmts(b"Kill \"c:/file1.txt\"");
    assert!(!errors, "Kill with string literal must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "Kill has one argument");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the path string");
}

#[test]
fn kill_wildcard_path() {
    let (_, arena, errors) = parse_stmts(b"Kill \"c:/*.bat\"");
    assert!(!errors, "Kill with wildcard path must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 1, "Kill has one argument");
    // The wildcard is just part of the string literal — it must not be split.
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "arg is the wildcard path string");
}

#[test]
fn filecopy_two_literals() {
    let (_, arena, errors) = parse_stmts(b"FileCopy \"c:/File1.txt\", \"c:/File2.txt\"");
    assert!(!errors, "FileCopy must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { args, .. } => *args,
        _ => unreachable!(),
    };
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 2, "FileCopy src, dst → 2 args");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "src is a string literal");
    assert!(matches!(at(&arena, args[1]), ExprNode::Literal { lit: AstLit::Str(_) }), "dst is a string literal");
}

#[test]
fn name_file_rename() {
    let (_, arena, errors) = parse_stmts(b"Dim F1, F2\nName F1 As F2");
    assert!(!errors, "Name <src> As <dst> (file-rename) must parse without errors");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Name, .. })),
        1,
        "one FileIoStmt {{ kind: Name }}"
    );
    let (channel, nargs) = match find_node(&arena, |n| matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Name, .. })).unwrap() {
        ExprNode::FileIoStmt { channel, args, .. } => (channel.is_some(), args.len()),
        _ => unreachable!(),
    };
    assert!(!channel, "Name is not channel-based → channel is None");
    assert_eq!(nargs, 2, "Name src As dst → src and dst as two args");
}

#[test]
fn name_file_rename_builds_fileio_node() {
    let (_, arena, errors) = parse_stmts(b"Dim F1, F2\nName F1 As F2");
    assert!(!errors, "Name <src> As <dst> must parse without errors");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Name, .. })) {
        Some(ExprNode::FileIoStmt { args, channel, .. }) => {
            assert!(channel.is_none(), "Name has no channel");
            args.clone()
        }
        _ => panic!("Name statement must build a FileIoStmt {{ kind: Name }}"),
    };
    assert_eq!(args.len(), 2, "Name carries src and dst");
    assert!(matches!(at(&arena, args[0]), ExprNode::NameRef { .. }), "src is the F1 reference");
    assert!(matches!(at(&arena, args[1]), ExprNode::NameRef { .. }), "dst is the F2 reference");
}

// ── File I/O: Open, Close, Print #, Write #, Line Input, Width # ──────────────

#[test]
fn open_for_output() {
    let (_, arena, errors) = parse_stmts(b"Open \"c:/test.txt\" For Output As #1");
    assert!(!errors, "Open For Output must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Open);
    assert!(has_channel, "Open ... As #1 records the channel");
    assert_eq!(nargs, 1, "Open For Output records the filename as its single arg");
}

#[test]
fn open_for_binary() {
    let (_, arena, errors) = parse_stmts(b"Open \"c:/Tmp/File\" For Binary As #2");
    assert!(!errors, "Open For Binary must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Open);
    assert!(has_channel, "Open ... As #2 records the channel");
    assert_eq!(nargs, 1, "Open For Binary records the filename as its single arg");
}

#[test]
fn open_for_append_with_access_lock_len() {
    let src = b"Open \"c:/f\" For Append Access Read Write Lock Read Write As #3 Len = 2";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Open with Access, Lock, and Len clauses must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Open);
    assert!(has_channel, "Open ... As #3 records the channel");
    // Beyond the filename, the `Len = 2` record length is captured as a second arg.
    assert_eq!(nargs, 2, "Open with Len clause records filename + record length");
}

#[test]
fn open_for_input() {
    let (_, arena, errors) = parse_stmts(b"Open \"C:/file\" For Input As #1");
    assert!(!errors, "Open For Input must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Open);
    assert!(has_channel, "Open ... As #1 records the channel");
    assert_eq!(nargs, 1, "Open For Input records the filename as its single arg");
}

#[test]
fn open_builds_fileio_node() {
    let (_, arena, errors) = parse_stmts(b"Open \"c:/f\" For Output As #1");
    assert!(!errors);
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Open);
    assert!(has_channel, "Open records the channel");
    assert_eq!(nargs, 1, "Open records the filename as its single arg");
}

#[test]
fn close_channel() {
    let (_, arena, errors) = parse_stmts(b"Close #1");
    assert!(!errors, "Close #n must parse without errors");
    // Close takes its channel list as args (a Close can list several channels),
    // not in the dedicated `channel` field.
    let (_has_channel, nargs) = fileio(&arena, FileIoKind::Close);
    assert_eq!(nargs, 1, "Close #1 lists exactly one channel");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Close, .. })).unwrap() {
        ExprNode::FileIoStmt { args, .. } => args.clone(),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Int(1) }), "the channel is #1");
}

#[test]
fn close_no_arg() {
    let (_, arena, errors) = parse_stmts(b"Close");
    assert!(!errors, "Close with no argument must parse without errors");
    // Bare Close (close all files) → no channel, no args.
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Close);
    assert!(!has_channel, "bare Close has no channel");
    assert_eq!(nargs, 0, "bare Close lists no channels");
}

#[test]
fn line_input_statement() {
    let (_, arena, errors) = parse_stmts(b"Dim s As String\nLine Input #1, s");
    assert!(!errors, "Line Input #n, var must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::LineInput);
    assert!(has_channel, "Line Input #1 records the channel");
    assert_eq!(nargs, 1, "Line Input reads into one target variable");
}

#[test]
fn line_input_builds_fileio_node() {
    let (_, arena, errors) = parse_stmts(b"Dim s As String\nLine Input #1, s");
    assert!(!errors);
    let (has_channel, nargs) = fileio(&arena, FileIoKind::LineInput);
    assert!(has_channel, "Line Input records the channel");
    assert_eq!(nargs, 1, "Line Input reads into one target variable");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::LineInput, .. })).unwrap() {
        ExprNode::FileIoStmt { args, .. } => args.clone(),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, args[0]), ExprNode::NameRef { .. }), "target is the variable s");
}

#[test]
fn print_with_channel() {
    let (_, arena, errors) = parse_stmts(b"Print #1, \"Cell 1\", 123, \"Cell 3\"");
    assert!(!errors, "Print #n, ... must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Print);
    assert!(has_channel, "Print #1 records the channel");
    assert_eq!(nargs, 3, "three comma-separated print items → 3 args");
}

#[test]
fn print_blank_to_channel() {
    let (_, arena, errors) = parse_stmts(b"Print #1,");
    assert!(!errors, "Print #n (blank) must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Print);
    assert!(has_channel, "Print #1 records the channel even with no items");
    assert_eq!(nargs, 0, "blank Print prints no items → 0 args");
}

#[test]
fn write_with_channel() {
    let (_, arena, errors) = parse_stmts(b"Write #1, \"Cell 1\", 123, \"Cell 3\"");
    assert!(!errors, "Write #n, ... must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Write);
    assert!(has_channel, "Write #1 records the channel");
    assert_eq!(nargs, 3, "three comma-separated write items → 3 args");
}

#[test]
fn write_blank_to_channel() {
    let (_, arena, errors) = parse_stmts(b"Write #1,");
    assert!(!errors, "Write #n (blank) must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Write);
    assert!(has_channel, "Write #1 records the channel even with no items");
    assert_eq!(nargs, 0, "blank Write writes no items → 0 args");
}

#[test]
fn write_semicolon_separator() {
    let (_, arena, errors) = parse_stmts(b"Write #1, 123; 456");
    assert!(!errors, "Write #n with semicolon separator must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Write);
    assert!(has_channel, "Write #1 records the channel");
    assert_eq!(nargs, 2, "semicolon-separated 123; 456 → two write items");
}

#[test]
fn write_builds_fileio_node() {
    let (_, arena, errors) = parse_stmts(b"Write #1, \"x\"");
    assert!(!errors);
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Write);
    assert!(has_channel, "Write #1 records the channel");
    assert_eq!(nargs, 1, "Write #1, \"x\" writes one item");
}

#[test]
fn width_file_statement() {
    let (_, arena, errors) = parse_stmts(b"Width #1, 10");
    assert!(!errors, "Width #n, cols must parse without errors");
    let (has_channel, nargs) = fileio(&arena, FileIoKind::Width);
    assert!(has_channel, "Width #1 records the channel");
    assert_eq!(nargs, 1, "Width records the column count as its single arg");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Width, .. })).unwrap() {
        ExprNode::FileIoStmt { args, .. } => args.clone(),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Int(10) }), "the width is 10");
}

// ── Module options ────────────────────────────────────────────────────────────

#[test]
fn option_explicit_accepts() {
    let (_, arena, errors) = parse_module(b"Option Explicit\nSub F()\nEnd Sub");
    assert!(!errors, "Option Explicit must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionExplicit)), 1, "one OptionExplicit");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Sub, .. })),
        1,
        "the trailing Sub must still parse"
    );
}

#[test]
fn option_base_zero() {
    let (_, arena, errors) = parse_module(b"Option Base 0\nSub F()\nEnd Sub");
    assert!(!errors, "Option Base 0 must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionBase { value: 0 })), 1, "one OptionBase{{value:0}}");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionBase { value: 1 })), 0, "no OptionBase{{value:1}}");
}

#[test]
fn option_base_one() {
    let (_, arena, errors) = parse_module(b"Option Base 1\nSub F()\nEnd Sub");
    assert!(!errors, "Option Base 1 must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionBase { value: 1 })), 1, "one OptionBase{{value:1}}");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionBase { value: 0 })), 0, "no OptionBase{{value:0}}");
}

#[test]
fn option_compare_binary() {
    let (_, arena, errors) = parse_module(b"Option Compare Binary\nSub F()\nEnd Sub");
    assert!(!errors, "Option Compare Binary must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionCompare { mode: 0 })), 1, "one OptionCompare{{mode:0}}");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionCompare { mode: 1 })), 0, "not Text mode");
}

#[test]
fn option_compare_text() {
    let (_, arena, errors) = parse_module(b"Option Compare Text\nSub F()\nEnd Sub");
    assert!(!errors, "Option Compare Text must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionCompare { mode: 1 })), 1, "one OptionCompare{{mode:1}}");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionCompare { mode: 0 })), 0, "not Binary mode");
}

#[test]
fn option_private_module() {
    let (_, arena, errors) = parse_module(b"Option Private Module\nSub F()\nEnd Sub");
    assert!(!errors, "Option Private Module must parse without errors");
    // Option Private Module is treated as an unknown Option variant (no dedicated
    // node); verify the Sub that follows still parsed correctly.
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Sub, .. })),
        1,
        "Sub after Option Private Module must still produce exactly one ProcDecl"
    );
}

#[test]
fn multiple_option_directives() {
    let src = b"Option Explicit\nOption Compare Text\nOption Base 0\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Multiple Option directives must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionExplicit)), 1, "exactly one OptionExplicit");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionCompare { mode: 1 })), 1, "exactly one OptionCompare Text");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OptionBase { value: 0 })), 1, "exactly one OptionBase 0");
}

// ── Declare statements ────────────────────────────────────────────────────────

#[test]
fn declare_sub_lib_alias_paramarray() {
    let src = b"Private Declare Sub subName Lib \"MyLib\" Alias \"alias1\" (arg1, arg2, ParamArray arg3)\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Declare Sub Lib Alias with ParamArray must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Sub, .. })), 1, "one Declare Sub");
    let (alias, ret, params) = match find_node(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Sub, .. })).unwrap() {
        ExprNode::DeclareDecl { alias, ret_type, params, .. } => (alias.is_some(), ret_type.is_some(), *params),
        _ => unreachable!(),
    };
    assert!(alias, "Alias \"alias1\" must be recorded");
    assert!(!ret, "a Declare Sub has no return type");
    // Three params incl. the ParamArray; the ParamArray flag (0x20) must be set on one.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 3, "arg1, arg2, ParamArray arg3 → 3 ParamDef");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { flags, .. } if flags & 0x20 != 0)), 1, "exactly one ParamArray param");
    assert!(matches!(params.map(|p| at(&arena, p)), Some(ExprNode::ArgList { .. })), "params is an ArgList");
}

#[test]
fn declare_function_lib_optional_byval() {
    let src = b"Private Declare Function Foo Lib \"MyLib\" (Optional arg1, ByVal arg2, arg3) As Currency\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Declare Function Lib with Optional/ByVal must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })), 1, "one Declare Function");
    let (alias, ret) = match find_node(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })).unwrap() {
        ExprNode::DeclareDecl { alias, ret_type, .. } => (alias.is_some(), *ret_type),
        _ => unreachable!(),
    };
    assert!(!alias, "no Alias clause given");
    // `As Currency` return type → a Currency BuiltinType (kind 6).
    assert!(matches!(ret.map(|r| at(&arena, r)), Some(ExprNode::BuiltinType { kind: 6 })), "return type is Currency (kind 6)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 3, "Optional arg1, ByVal arg2, arg3 → 3 ParamDef");
}

#[test]
fn declare_function_empty_arg_list() {
    let src = b"Public Declare Function Foo Lib \"MyLib\" Alias \"alias3\" ()\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Declare Function with empty arg list must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })), 1, "one Declare Function");
    let alias = match find_node(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })).unwrap() {
        ExprNode::DeclareDecl { alias, .. } => alias.is_some(),
        _ => unreachable!(),
    };
    assert!(alias, "Alias \"alias3\" must be recorded");
    // Empty () → no params.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 0, "empty arg list → zero ParamDef");
}

#[test]
fn declare_api_real_world() {
    let src = b"Private Declare Function GetComputerName Lib \"kernel32\" Alias \"GetComputerNameA\" (ByVal lpBuffer As String, nSize As Long) As Long\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Real-world API Declare must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })), 1, "one Declare Function");
    let (alias, ret) = match find_node(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })).unwrap() {
        ExprNode::DeclareDecl { alias, ret_type, .. } => (alias.is_some(), *ret_type),
        _ => unreachable!(),
    };
    assert!(alias, "Alias \"GetComputerNameA\" must be recorded");
    assert!(matches!(ret.map(|r| at(&arena, r)), Some(ExprNode::BuiltinType { kind: 3 })), "return type is Long (kind 3)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 2, "lpBuffer + nSize → 2 ParamDef");
}

#[test]
fn declare_with_line_continuation() {
    let src = b"Declare Function SetComputerName Lib \"kernel32\" _\n  Alias \"SetComputerNameA\" ( _\n  ByVal lpComputerName As String) As Long\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Declare with line continuation must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })), 1, "one Declare Function");
    let ret = match find_node(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })).unwrap() {
        ExprNode::DeclareDecl { ret_type, .. } => *ret_type,
        _ => unreachable!(),
    };
    // The continuation must not lose the `As Long` return type or the single param.
    assert!(matches!(ret.map(|r| at(&arena, r)), Some(ExprNode::BuiltinType { kind: 3 })), "return type is Long (kind 3)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 1, "single lpComputerName param");
}

#[test]
fn declare_with_type_suffix_param() {
    let src = b"Declare Function Func Lib \"Foo.dll\" (a$)\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Declare with type-suffix parameter must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DeclareDecl { kind: ProcKind::Function, .. })), 1, "one Declare Function");
    // `a$` is one type-suffixed parameter.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 1, "one param a$");
}

// ── Const extensions ──────────────────────────────────────────────────────────

#[test]
fn const_multi_declarator() {
    let src = b"Const A = \"test\", B As Integer = 3, C As String = \"test3\"";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Const with multiple declarators must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::DimItem { is_const: true, .. }));
    assert_eq!(n, 3, "Three-declarator Const must produce 3 DimItem {{ is_const: true }} nodes");
    // Every declarator is initialised; the typed ones (B As Integer, C As String)
    // must carry a type_node while the untyped first (A = "test") must not.
    assert_eq!(
        count_nodes(&arena, |node| matches!(node, ExprNode::DimItem { is_const: true, init: Some(_), .. })),
        3,
        "all three Const declarators have an initialiser"
    );
    assert_eq!(
        count_nodes(&arena, |node| matches!(node, ExprNode::DimItem { is_const: true, type_node: Some(_), .. })),
        2,
        "B As Integer and C As String carry a type_node; A does not"
    );
}

#[test]
fn const_private_at_module_level() {
    let (_, arena, errors) = parse_module(b"Private Const X = 345\nSub F()\nEnd Sub");
    assert!(!errors, "Private Const at module level must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })), 1, "one const DimItem");
    // X = 345 → an Int initialiser, no explicit type.
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })).unwrap();
    match item {
        ExprNode::DimItem { type_node, init, .. } => {
            assert!(type_node.is_none(), "untyped Const has no type_node");
            assert!(matches!(init.map(|i| at(&arena, i)), Some(ExprNode::Literal { lit: AstLit::Int(345) })), "init is 345");
        }
        _ => unreachable!(),
    }
}

#[test]
fn const_public_typed_at_module_level() {
    let (_, arena, errors) = parse_module(b"Public Const X As Double = 567.1\nSub F()\nEnd Sub");
    assert!(!errors, "Public Const with type annotation at module level must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })), 1, "one const DimItem");
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })).unwrap();
    match item {
        ExprNode::DimItem { type_node, init, .. } => {
            // As Double → BuiltinType kind 5; = 567.1 → Double literal.
            assert!(matches!(type_node.map(|t| at(&arena, t)), Some(ExprNode::BuiltinType { kind: 5 })), "type is Double (kind 5)");
            assert!(matches!(init.map(|i| at(&arena, i)), Some(ExprNode::Literal { lit: AstLit::Double(_) })), "init is a Double literal");
        }
        _ => unreachable!(),
    }
}

#[test]
fn const_inside_sub() {
    let (_, arena, errors) = parse_stmts(b"Const PI As Double = 3.1415");
    assert!(!errors, "Const inside Sub must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })), 1, "one const DimItem");
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })).unwrap();
    match item {
        ExprNode::DimItem { type_node, init, .. } => {
            assert!(matches!(type_node.map(|t| at(&arena, t)), Some(ExprNode::BuiltinType { kind: 5 })), "PI As Double → kind 5");
            assert!(matches!(init.map(|i| at(&arena, i)), Some(ExprNode::Literal { lit: AstLit::Double(_) })), "init is a Double literal");
        }
        _ => unreachable!(),
    }
}

#[test]
fn const_embedded_quotes_in_string() {
    let (_, arena, errors) = parse_stmts(b"Const Msg = \"Hello \"\"quoted\"\" world\"");
    assert!(!errors, "Const with embedded double-quotes must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })), 1, "one const DimItem");
    let item = find_node(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })).unwrap();
    match item {
        ExprNode::DimItem { init, .. } => match init.map(|i| at(&arena, i)) {
            // The doubled "" escapes must collapse to single quotes in the value.
            Some(ExprNode::Literal { lit: AstLit::Str(s) }) => {
                assert_eq!(&**s, "Hello \"quoted\" world", "embedded quotes decoded correctly");
            }
            other => panic!("init must be a string literal, got {:?}", other.map(kind_name)),
        },
        _ => unreachable!(),
    }
}

// ── ReDim extensions ──────────────────────────────────────────────────────────

#[test]
fn redim_preserve() {
    let (_, arena, errors) = parse_stmts(b"Dim A() As Integer\nReDim Preserve A(40)");
    assert!(!errors, "ReDim Preserve must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ReDimItem { preserve: true, .. })), 1, "one ReDimItem{{preserve:true}}");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ReDimItem { preserve: false, .. })), 0, "no non-preserve ReDimItem");
    // The ReDim must carry array bounds A(40).
    let bounds = match find_node(&arena, |n| matches!(n, ExprNode::ReDimItem { preserve: true, .. })).unwrap() {
        ExprNode::ReDimItem { bounds, .. } => *bounds,
        _ => unreachable!(),
    };
    assert!(bounds.is_some(), "ReDim Preserve A(40) carries bounds");
}

// ── Enum extensions ───────────────────────────────────────────────────────────

#[test]
fn enum_with_negative_member_value() {
    let (_, arena, errors) = parse_module(b"Enum E\nA = -1\nEnd Enum\nSub F()\nEnd Sub");
    assert!(!errors, "Enum with negative member value must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EnumDecl { .. })), 1, "one EnumDecl");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::EnumDecl { members, .. } if members.len() == 1)),
        1,
        "EnumDecl has exactly 1 member"
    );
    // A = -1 → the value is a Neg unary op over a literal.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::UnOp { op: UnOpKind::Neg, .. })), 1, "negative member value → one Neg UnOp");
}

#[test]
fn enum_with_hex_literal_member_value() {
    let (_, arena, errors) = parse_module(b"Enum E\nA = &H123ABC&\nEnd Enum\nSub F()\nEnd Sub");
    assert!(!errors, "Enum with hex-literal member value must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EnumDecl { .. })), 1, "one EnumDecl");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::EnumDecl { members, .. } if members.len() == 1)),
        1,
        "EnumDecl has exactly 1 member"
    );
    // &H123ABC& is a Long-typed hex literal → an integer Literal in the value.
    assert!(
        find_node(&arena, |n| matches!(n, ExprNode::Literal { lit: AstLit::Int(_) | AstLit::Long(_) })).is_some(),
        "hex member value parsed as an integer literal"
    );
}

#[test]
fn public_enum_accepts() {
    let src = b"Public Enum Days\n  Monday\n  Tuesday\n  Wednesday\nEnd Enum\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Public Enum must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EnumDecl { .. })), 1, "one EnumDecl");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::EnumDecl { members, .. } if members.len() == 3)),
        1,
        "Days enum has exactly 3 members (Monday, Tuesday, Wednesday)"
    );
}

// ── DefType full set ──────────────────────────────────────────────────────────

#[test]
fn deftype_full_alphabet() {
    let src = b"\
DefBool A-B\n\
DefByte C-D\n\
DefInt E-F\n\
DefLng G-H\n\
DefCur I-J\n\
DefSng K-L\n\
DefDbl M\n\
DefDec N,O-P\n\
DefDate Q,R-S,T\n\
DefStr U-V\n\
DefObj W-X\n\
DefVar Y-Z\n\
Sub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "All twelve DefType variants must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::DefType { .. }));
    assert_eq!(n, 12, "Twelve DefType directives must build 12 DefType nodes");
}

// ── Type (UDT) extensions ─────────────────────────────────────────────────────

#[test]
fn private_udt_with_typed_members() {
    let src = b"Private Type T\n  V1 As Integer\n  V2 As Double\nEnd Type\nSub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Private Type with typed members must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::TypeDecl { .. })), 1, "one TypeDecl");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::TypeDecl { members, .. } if members.len() == 2)),
        1,
        "TypeDecl has exactly 2 members (V1, V2)"
    );
    // Both members are typed: V1 As Integer (kind 2), V2 As Double (kind 5).
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::BuiltinType { kind: 2 })), 1, "V1 As Integer");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::BuiltinType { kind: 5 })), 1, "V2 As Double");
}

#[test]
fn public_udt_with_untyped_and_nested_member() {
    let src = b"\
Type T1\n  X As Integer\nEnd Type\n\
Public Type T2\n  V1 As Currency\n  V2 As String\n  V3\n  V4 As T1\nEnd Type\n\
Sub F()\nEnd Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Public Type with untyped member and UDT-typed member must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::TypeDecl { .. }));
    assert_eq!(n, 2, "Two Type declarations must produce 2 TypeDecl nodes");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::TypeDecl { members, .. } if members.len() == 1)),
        1,
        "T1 has exactly 1 member (X)"
    );
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::TypeDecl { members, .. } if members.len() == 4)),
        1,
        "T2 has exactly 4 members (V1..V4, incl. the untyped V3 and UDT-typed V4)"
    );
}

#[test]
fn udt_member_access_in_stmt() {
    let (_, arena, errors) = parse_stmts(b"Dim t As Object\nt.Variable1 = 1");
    assert!(!errors, "UDT member access assignment must parse without errors");
    // t.Variable1 = 1 → an Assign whose target is a (non-bang) MemberAccess.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MemberAccess { bang: false, .. })), 1, "one dotted MemberAccess");
    let (target, value) = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).expect("an Assign") {
        ExprNode::Assign { target, value } => (*target, *value),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, target), ExprNode::MemberAccess { bang: false, .. }), "assignment target is t.Variable1");
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Int(1) }), "assigned value is 1");
}

#[test]
fn udt_array_element_member_access() {
    let (_, arena, errors) = parse_stmts(b"Dim arr(10) As Object\narr(1).Field = \"hello world\"");
    assert!(!errors, "UDT array element member access must parse without errors");
    // arr(1).Field = "..." → MemberAccess whose base is a Call (the arr(1) index).
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MemberAccess { bang: false, .. })), 1, "one dotted MemberAccess");
    let base = match find_node(&arena, |n| matches!(n, ExprNode::MemberAccess { bang: false, .. })).unwrap() {
        ExprNode::MemberAccess { base, .. } => *base,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, base), ExprNode::Call { .. }), "the member base is the arr(1) index call");
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).expect("an Assign") {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Str(_) }), "assigned value is the string");
}

// ── Event declarations ────────────────────────────────────────────────────────

#[test]
fn event_no_args() {
    let (_, arena, errors) = parse_class_module(b"Public Event SomeEvent()");
    assert!(!errors, "Event with empty arg list must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EventDecl { .. })), 1, "one EventDecl");
    let params = match find_node(&arena, |n| matches!(n, ExprNode::EventDecl { .. })).unwrap() {
        ExprNode::EventDecl { params, .. } => *params,
        _ => unreachable!(),
    };
    // Empty () → an ArgList with no params.
    assert!(arglist_args(&arena, params).is_empty(), "SomeEvent() has no parameters");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 0, "no ParamDef nodes");
}

#[test]
fn event_with_byval_string_arg() {
    let (_, arena, errors) = parse_class_module(b"Public Event SomeOtherEvent(ByVal Message As String)");
    assert!(!errors, "Event with ByVal argument must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EventDecl { .. })), 1, "one EventDecl");
    // One parameter: ByVal Message As String → a ParamDef with the ByVal flag (0x2)
    // and a String type.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 1, "one parameter");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { flags, .. } if flags & 0x2 != 0)), 1, "the param is ByVal (flag 0x2)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::StringType { .. })), 1, "As String → a StringType");
}

#[test]
fn event_with_type_suffix_param() {
    let (_, arena, errors) = parse_class_module(b"Event SomeEvent(a$)");
    assert!(!errors, "Event with type-suffix parameter must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EventDecl { .. })), 1, "one EventDecl");
    // a$ is one parameter conveyed via type suffix.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 1, "one param a$");
}

// ── Control flow extensions ───────────────────────────────────────────────────

#[test]
fn do_loop_until() {
    let (_, arena, errors) = parse_stmts(b"Do\nLoop Until i >= 1");
    assert!(!errors, "Do/Loop Until must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Do { .. })), 1, "one Do node");
    let cond = match find_node(&arena, |n| matches!(n, ExprNode::Do { .. })).unwrap() {
        ExprNode::Do { kind, cond, .. } => {
            assert_eq!(*kind, DoKind::PostUntil, "Loop Until → PostUntil");
            *cond
        }
        _ => unreachable!(),
    };
    // i >= 1 → a Ge comparison as the loop condition.
    assert!(matches!(cond.map(|c| at(&arena, c)), Some(ExprNode::BinOp { op: BinOpKind::Ge, .. })), "condition is i >= 1");
}

#[test]
fn do_while_not_condition() {
    let (_, arena, errors) = parse_stmts(b"Dim I\nI = 5\nDo While Not I = 0\n  I = I - 1\nLoop");
    assert!(!errors, "Do While Not <cond> must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Do { .. })), 1, "one Do node");
    let cond = match find_node(&arena, |n| matches!(n, ExprNode::Do { .. })).unwrap() {
        ExprNode::Do { kind, cond, .. } => {
            assert_eq!(*kind, DoKind::PreWhile, "Do While → PreWhile");
            *cond
        }
        _ => unreachable!(),
    };
    // Not I = 0 → a Not unary over the equality.
    assert!(matches!(cond.map(|c| at(&arena, c)), Some(ExprNode::UnOp { op: UnOpKind::Not, .. })), "condition is a Not expression");
}

#[test]
fn do_until_condition() {
    let (_, arena, errors) = parse_stmts(b"Dim I\nDo Until I = 0\n  I = I - 1\nLoop");
    assert!(!errors, "Do Until must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Do { .. })), 1, "one Do node");
    let cond = match find_node(&arena, |n| matches!(n, ExprNode::Do { .. })).unwrap() {
        ExprNode::Do { kind, cond, .. } => {
            assert_eq!(*kind, DoKind::PreUntil, "Do Until → PreUntil");
            *cond
        }
        _ => unreachable!(),
    };
    // I = 0 → an Eq comparison.
    assert!(matches!(cond.map(|c| at(&arena, c)), Some(ExprNode::BinOp { op: BinOpKind::Eq, .. })), "condition is I = 0");
}

#[test]
fn nested_while_wend() {
    let src = b"Dim I\nI = 100\nWhile I > 0\n  While I Mod 15 <> 0\n    I = I - 1\n  Wend\nWend";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Nested While/Wend must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::While { .. }));
    assert_eq!(n, 2, "Nested While/Wend must build 2 While nodes");
    // The inner While must actually be nested in the outer While's body, not a sibling.
    let outer = find_node(&arena, |node| {
        matches!(node, ExprNode::While { body, .. } if {
            block_stmts(&arena, *body).iter().any(|s| matches!(at(&arena, *s), ExprNode::While { .. }))
        })
    });
    assert!(outer.is_some(), "the outer While body must contain the inner While");
}

#[test]
fn nested_select_case() {
    let src = b"\
Dim A, B\n\
Select Case A\n\
  Case \"A\"\n\
    Select Case B\n\
      Case \"A\"\n\
        Beep\n\
      Case Else\n\
        Beep\n\
    End Select\n\
  Case Else\n\
    Beep\n\
End Select";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Nested Select Case must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::SelectCase { .. }));
    assert_eq!(n, 2, "Nested Select Case must build 2 SelectCase nodes");
}

#[test]
fn select_case_mixed_range_is_list() {
    let src = b"\
Dim P\n\
Select Case P\n\
  Case 0 To 25\n    Beep\n\
  Case 26 To 49, 50\n\
  Case 51 To 75\n    Beep\n\
  Case 76 To 80, 81, 82 To 89, Is > 90\n\
  Case Else\n    Beep\n\
End Select";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Select Case with To, Is, and list combinations must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::SelectCase { .. })), 1, "one SelectCase");
    // Four Case arms + one Case Else.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CaseBlock { .. })), 4, "four Case arms");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CaseElse { .. })), 1, "one Case Else");
    // The `To` ranges: 0 To 25, 26 To 49, 51 To 75, 76 To 80, 82 To 89 → 5 RangeTo.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::RangeTo { .. })), 5, "five To-ranges");
    // `Is > 90` → one CaseIs with a Gt operator.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CaseIs { op: BinOpKind::Gt, .. })), 1, "one `Is > 90` CaseIs");
}

#[test]
fn select_case_multiple_else_clauses() {
    let src = b"\
Dim Grade As String\n\
Select Case Grade\n\
  Case \"A\"\n    Beep\n\
  Case Else\n    Beep\n\
  Case \"B\"\n    Beep\n\
  Case Else\n    Beep\n\
End Select";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Select Case with multiple Case Else clauses must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::CaseElse { .. }));
    assert_eq!(n, 2, "Two Case Else clauses must build 2 CaseElse nodes");
}

#[test]
fn for_next_step_negative_type_suffix_counter() {
    let (_, arena, errors) = parse_stmts(b"Dim i%\nFor i% = 10 To 1 Step -1\n  Beep\nNext i%");
    assert!(!errors, "For/Next with negative Step and type-suffix counter must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::For { .. })), 1, "one For node");
    let step = match find_node(&arena, |n| matches!(n, ExprNode::For { .. })).unwrap() {
        ExprNode::For { step, .. } => *step,
        _ => unreachable!(),
    };
    // Step -1 → a Neg unary over the literal 1.
    assert!(matches!(step.map(|s| at(&arena, s)), Some(ExprNode::UnOp { op: UnOpKind::Neg, .. })), "Step -1 is a Neg unary");
}

#[test]
fn for_next_member_access_limit_and_step() {
    let src = b"Dim I, J, K\nFor I = 0 To K.Value\n  For J = 1 To 20 Step 2\n    Beep\n  Next\nNext";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "Nested For/Next with member-access limit and Step must parse without errors");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::For { .. }));
    assert_eq!(n, 2, "Two nested For loops must build 2 For nodes");
    // The outer For limit `K.Value` must be a MemberAccess (and exactly one exists).
    assert_eq!(count_nodes(&arena, |node| matches!(node, ExprNode::MemberAccess { .. })), 1, "one MemberAccess (K.Value)");
    let outer = find_node(&arena, |node| matches!(node, ExprNode::For { end, .. } if matches!(at(&arena, *end), ExprNode::MemberAccess { .. })));
    assert!(outer.is_some(), "the outer For's end limit is K.Value");
    // Exactly one of the two For loops carries a Step (the inner one).
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::For { step: Some(_), .. })), 1, "only the inner For has a Step");
}

#[test]
fn for_next_colon_on_one_line() {
    let (_, arena, errors) = parse_stmts(b"Dim i\nFor i = 1 To 32000: Next i");
    assert!(!errors, "For/Next on one line with colon separator must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::For { .. })), 1, "one For node");
    let (start, end, step, body) = match find_node(&arena, |n| matches!(n, ExprNode::For { .. })).unwrap() {
        ExprNode::For { start, end, step, body, .. } => (*start, *end, *step, *body),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, start), ExprNode::Literal { lit: AstLit::Int(1) }), "start = 1");
    assert!(matches!(at(&arena, end), ExprNode::Literal { lit: AstLit::Int(32000) }), "end = 32000");
    assert!(step.is_none(), "no Step");
    // The empty body (Next immediately follows the colon) → an empty Block.
    assert!(block_stmts(&arena, body).is_empty(), "the colon body is empty");
}

#[test]
fn elseif_single_line_body() {
    let src = b"\
Dim i\n\
If i = 1 Then\n  Beep\n\
ElseIf i = 2 Then Beep\n\
ElseIf i = 3 Then Beep\n\
End If";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "ElseIf with single-line body must parse without errors");
    // Two ElseIf arms desugar to nested If: 3 If nodes total. The two leading Ifs
    // carry an else_body (the next arm); the innermost ElseIf has no Else → None.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::If { .. })), 3, "If + 2 ElseIf → 3 If nodes");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::If { else_body: Some(_), .. })), 2, "two Ifs chain into an else_body");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::If { else_body: None, .. })), 1, "the last ElseIf has no Else");
    // The chained else_body of the outer If must itself be an If (the first ElseIf).
    let outer = find_node(&arena, |n| matches!(n, ExprNode::If { cond, .. } if matches!(at(&arena, *cond), ExprNode::BinOp { rhs, .. } if matches!(at(&arena, *rhs), ExprNode::Literal { lit: AstLit::Int(1) }))));
    let outer_else = match outer.expect("the `i = 1` If") {
        ExprNode::If { else_body, .. } => else_body.expect("outer If has else_body"),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, outer_else), ExprNode::If { .. }), "the outer If's else_body is the first ElseIf (a nested If)");
}

#[test]
fn goto_multiple_labels() {
    let src = b"\
GoTo LineLabel1\n\
LineLabel1:\n  Dim x\n\
LineLabel2:\n  x = 2";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "GoTo with multiple line labels must parse without errors");
    // Exactly one GoTo, targeting a NAMED label (LineLabel1) — this is the
    // regression guard: a mis-parse would drop the GoTo entirely.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::GoTo { .. })), 1, "one GoTo node");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::GoTo { target: LabelRef::Name(_) })),
        1,
        "the GoTo targets a named label, not a numeric line"
    );
    // Two label definitions, both named.
    assert_eq!(count_nodes(&arena, |node| matches!(node, ExprNode::Label { .. })), 2, "two Label nodes");
    assert_eq!(
        count_nodes(&arena, |node| matches!(node, ExprNode::Label { target: LabelRef::Name(_) })),
        2,
        "both labels are named labels"
    );
}

#[test]
fn on_error_goto_label_with_colon() {
    let (_, arena, errors) = parse_stmts(b"On Error GoTo ErrorHandler:\nErrorHandler:");
    assert!(!errors, "On Error GoTo where label target has a colon suffix must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::OnError { .. })), 1, "one OnError node");
    // It must be a GoTo-to-named-label handler (not Disable / ResumeNext), and the
    // ErrorHandler: definition must still be recognised as a Label.
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::OnError { kind: OnErrorKind::Goto(LabelRef::Name(_)) })),
        1,
        "On Error GoTo ErrorHandler → Goto(Name)"
    );
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Label { .. })), 1, "the ErrorHandler: definition is a Label");
}

// ── With block extensions ─────────────────────────────────────────────────────

#[test]
fn with_new_instance() {
    let (_, arena, errors) = parse_stmts(b"With New clsFoo\n  .Display\nEnd With");
    assert!(!errors, "With New <class> must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::With { .. })), 1, "one With node");
    // The With object is a `New clsFoo` instantiation.
    let obj = match find_node(&arena, |n| matches!(n, ExprNode::With { .. })).unwrap() {
        ExprNode::With { obj, .. } => *obj,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, obj), ExprNode::New { .. }), "With object is a New expression");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::New { .. })), 1, "one New node");
}

#[test]
fn with_property_set_and_method_call() {
    let src = b"Dim o As Variant\nWith o\n  .MemberProp = \"SomeValue\"\n  .MemberCall\nEnd With";
    let (_, arena, errors) = parse_stmts(src);
    assert!(!errors, "With member property set and method call must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::With { .. })), 1, "one With node");
    // Inside: `.MemberProp = "SomeValue"` → one Assign; `.MemberCall` → no extra Assign.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "one property-set Assign in the With body");
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).unwrap() {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Str(_) }), "the assigned value is \"SomeValue\"");
}

// ── Literals ──────────────────────────────────────────────────────────────────

#[test]
fn octal_literal() {
    // &O55 = octal 55 = 45 decimal; fits in i16 so scanner emits IntLit.
    let (_, arena, errors) = parse_stmts(b"Dim foo As Long\nfoo = &O55");
    assert!(!errors, "Octal literal &O55 must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Literal { lit: AstLit::Int(45) })), 1, "one Int(45) literal");
    // It must be the value of the `foo = &O55` assignment.
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).expect("an Assign") {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Int(45) }), "&O55 decodes to 45 as the assigned value");
}

#[test]
fn time_literal() {
    let (_, arena, errors) = parse_stmts(b"Dim t\nt = #8:00:00 AM#");
    assert!(!errors, "Time literal #8:00:00 AM# must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Literal { lit: AstLit::Date(_) })), 1, "one Date literal");
    // It is the assigned value of `t = #...#`.
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).expect("an Assign") {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Date(_) }), "the time literal is the assigned value");
}

#[test]
fn date_literal_named_month() {
    let (_, arena, errors) = parse_stmts(b"Dim d\nd = #December, 24, 2000#");
    assert!(!errors, "Date literal with named month must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Literal { lit: AstLit::Date(_) })), 1, "one Date literal");
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).expect("an Assign") {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Date(_) }), "the date literal is the assigned value");
}

#[test]
fn time_property_assign() {
    let (_, arena, errors) = parse_stmts(b"Dim t\nt = #8:00:00 AM#\nTime = t");
    assert!(!errors, "Assignment to built-in Time property must parse without errors");
    // `t = #...#` then `Time = t` → two assignments; both targets are NameRefs.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 2, "two Assign nodes");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::Assign { target, .. } if matches!(at(&arena, *target), ExprNode::NameRef { .. }))),
        2,
        "both assignment targets are name references (t and Time)"
    );
}

#[test]
fn date_property_assign() {
    let (_, arena, errors) = parse_stmts(b"Dim d\nd = #December, 24, 2000#\nDate = d");
    assert!(!errors, "Assignment to built-in Date property must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 2, "two Assign nodes");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::Assign { target, .. } if matches!(at(&arena, *target), ExprNode::NameRef { .. }))),
        2,
        "both assignment targets are name references (d and Date)"
    );
}

#[test]
fn date_dollar_function_ref() {
    let (_, arena, errors) = parse_stmts(b"Dim s As String\ns = Date$");
    assert!(!errors, "Date$ function reference must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "one Assign");
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).unwrap() {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    // Date$ is a bare name reference (the `$` is a type-decl suffix), not a call.
    assert!(matches!(at(&arena, value), ExprNode::NameRef { .. }), "Date$ is a NameRef value");
}

// ── Mid (no $) assignment ─────────────────────────────────────────────────────

#[test]
fn mid_no_dollar_with_length() {
    let (_, arena, errors) = parse_stmts(b"Dim s\nMid(s, 0, 1) = \"L\"");
    assert!(!errors, "Mid (no $) assignment with length must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MidAssign { .. })), 1, "one MidAssign");
    let (byte_oriented, dollar, args, value) = match find_node(&arena, |n| matches!(n, ExprNode::MidAssign { .. })).unwrap() {
        ExprNode::MidAssign { byte_oriented, dollar, args, value } => (*byte_oriented, *dollar, *args, *value),
        _ => unreachable!(),
    };
    assert!(!byte_oriented, "plain Mid is character-oriented");
    assert!(!dollar, "no $ spelling");
    // (s, 0, 1) → a 3-arg tuple, distinguishing the with-length form.
    assert_eq!(arglist_args(&arena, args).len(), 3, "Mid(s, 0, 1) → 3 args (with length)");
    assert!(matches!(at(&arena, value), ExprNode::Literal { lit: AstLit::Str(_) }), "replacement value is \"L\"");
}

#[test]
fn mid_no_dollar_without_length() {
    let (_, arena, errors) = parse_stmts(b"Dim s\nMid(s, 5) = \"   \"");
    assert!(!errors, "Mid (no $) assignment without length must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MidAssign { dollar: false, .. })), 1, "one MidAssign dollar=false");
    let args = match find_node(&arena, |n| matches!(n, ExprNode::MidAssign { .. })).unwrap() {
        ExprNode::MidAssign { args, .. } => *args,
        _ => unreachable!(),
    };
    // (s, 5) → a 2-arg tuple, distinguishing the without-length form.
    assert_eq!(arglist_args(&arena, args).len(), 2, "Mid(s, 5) → 2 args (no length)");
}

// ── Type-suffix parameters ────────────────────────────────────────────────────

#[test]
fn sub_type_suffix_params() {
    let (_, arena, errors) = parse_module(b"Public Sub Foo(a$, b&, c!, d#, e@)\nEnd Sub");
    assert!(!errors, "Sub with type-suffix parameters must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Sub, .. })), 1, "one Sub ProcDecl");
    let n = count_nodes(&arena, |node| matches!(node, ExprNode::ParamDef { .. }));
    assert_eq!(n, 5, "Five type-suffix params must build 5 ParamDef nodes");
    // Each type suffix is lowered to a type node on its ParamDef, so all 5 carry one.
    assert_eq!(
        count_nodes(&arena, |node| matches!(node, ExprNode::ParamDef { type_node: Some(_), .. })),
        5,
        "every type-suffix param carries a lowered type_node"
    );
}

#[test]
fn function_type_suffix_with_array_param_and_return() {
    let src = b"Public Function Foo(a$, b&, c!, d#, e@, f$()) As Boolean\nEnd Function";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Function with type-suffix, array param, and typed return must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Function, ret_type: Some(_), .. })), 1, "one Function with a return type");
    let ret = match find_node(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Function, .. })).unwrap() {
        ExprNode::ProcDecl { ret_type, .. } => *ret_type,
        _ => unreachable!(),
    };
    // As Boolean → BuiltinType kind 11.
    assert!(matches!(ret.map(|r| at(&arena, r)), Some(ExprNode::BuiltinType { kind: 11 })), "return type is Boolean (kind 11)");
    // Five scalar suffix params + one array param f$() → 6 ParamDef.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 6, "a$..e@ + f$() → 6 ParamDef");
}

#[test]
fn function_local_type_suffix_array_dim() {
    let src = b"Public Function Foo(a$) As Boolean\n  Dim arr$()\nEnd Function";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Dim with type-suffix array inside function must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Function, .. })), 1, "one Function ProcDecl");
    // The local `Dim arr$()` produces one DimItem carrying a (suffix-derived) type_node.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { .. })), 1, "one local DimItem (arr$())");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { type_node: Some(_), .. })), 1, "the $ suffix gives arr a type_node");
}

#[test]
fn paramarray_param() {
    let src = b"Function Foo(ByVal FirstArg As Double, ParamArray AdditionalArgs())\nEnd Function";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "ParamArray parameter must parse without errors");
    // Two params: ByVal FirstArg (flag 0x2) and ParamArray AdditionalArgs (flag 0x20).
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { .. })), 2, "two ParamDef");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { flags, .. } if flags & 0x20 != 0)), 1, "exactly one ParamArray (flag 0x20)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ParamDef { flags, .. } if flags & 0x2 != 0)), 1, "exactly one ByVal (flag 0x2)");
}

// ── Call variants ─────────────────────────────────────────────────────────────

#[test]
fn implicit_call_no_parens_no_args() {
    let (_, arena, errors) = parse_stmts(b"Beep");
    assert!(!errors, "Implicit call with no keyword, no parens, no args must parse without errors");
    // Zero-arg implicit call → a bare NameRef, never wrapped in a CallStmt.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::NameRef { .. })), 1, "one NameRef");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 0, "no CallStmt for a zero-arg implicit call");
}

#[test]
fn implicit_call_with_args_no_parens() {
    let (_, arena, errors) = parse_stmts(b"Sub2 1, 2");
    assert!(!errors, "Implicit call with args and no parens must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    let (callee, args) = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { callee, args } => (*callee, *args),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, callee), ExprNode::NameRef { .. }), "callee is Sub2");
    let args = arglist_args(&arena, args);
    assert_eq!(args.len(), 2, "Sub2 1, 2 → two args");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Int(1) }), "arg0 = 1");
    assert!(matches!(at(&arena, args[1]), ExprNode::Literal { lit: AstLit::Int(2) }), "arg1 = 2");
}

#[test]
fn member_call_explicit_keyword() {
    let (_, arena, errors) = parse_stmts(b"Dim m As Object\nCall m.Function1");
    assert!(!errors, "Call on member (explicit Call keyword) must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    // The callee is the member access m.Function1.
    let callee = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { callee, .. } => *callee,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, callee), ExprNode::MemberAccess { .. }), "callee is m.Function1");
}

#[test]
fn member_call_with_args() {
    let (_, arena, errors) = parse_stmts(b"Dim m As Object\nCall m.Function2(1, 2)");
    assert!(!errors, "Call on member with arguments must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    // Call m.Function2(1, 2): the CallStmt callee is the inner Call (m.Function2
    // applied to its arg list); the (1, 2) attaches to that Call, not the CallStmt.
    let callee = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { callee, .. } => *callee,
        _ => unreachable!(),
    };
    let inner_args = match at(&arena, callee) {
        ExprNode::Call { func, args } => {
            assert!(matches!(at(&arena, *func), ExprNode::MemberAccess { .. }), "the call target is m.Function2");
            *args
        }
        other => panic!("callee must be a Call, got {}", kind_name(other)),
    };
    assert_eq!(arglist_args(&arena, inner_args).len(), 2, "Function2(1, 2) → two args on the inner Call");
}

#[test]
fn member_call_implicit() {
    let (_, arena, errors) = parse_stmts(b"Dim m As Object\nm.Function1");
    assert!(!errors, "Implicit member call (no Call keyword) must parse without errors");
    // No args: the statement is the MemberAccess itself, never wrapped in a CallStmt.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MemberAccess { .. })), 1, "one MemberAccess");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 0, "no CallStmt for a zero-arg implicit member call");
}

#[test]
fn chained_member_call_with_parens() {
    let (_, arena, errors) = parse_stmts(b"Dim m As Object\nCall m.GetModule().Function1");
    assert!(!errors, "Chained member call must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::CallStmt { .. })), 1, "one CallStmt");
    // m.GetModule().Function1: the callee is .Function1 over the GetModule() call,
    // so there must be both a Call (GetModule()) and two member accesses.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Call { .. })), 1, "one Call (GetModule())");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MemberAccess { .. })), 2, "two member accesses (.GetModule, .Function1)");
    let callee = match find_node(&arena, |n| matches!(n, ExprNode::CallStmt { .. })).unwrap() {
        ExprNode::CallStmt { callee, .. } => *callee,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, callee), ExprNode::MemberAccess { .. }), "the CallStmt callee is the .Function1 member access");
}

#[test]
fn assignment_from_member_call() {
    let (_, arena, errors) = parse_stmts(b"Dim m As Object\nDim I\nI = m.Function1");
    assert!(!errors, "Assignment from member call must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "one Assign");
    let (target, value) = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).unwrap() {
        ExprNode::Assign { target, value } => (*target, *value),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, target), ExprNode::NameRef { .. }), "target is I");
    assert!(matches!(at(&arena, value), ExprNode::MemberAccess { .. }), "value is m.Function1");
}

#[test]
fn assignment_from_chained_member_call() {
    let (_, arena, errors) = parse_stmts(b"Dim m As Object\nDim I\nI = m.GetModule().Function1()");
    assert!(!errors, "Assignment from chained member call must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "one Assign");
    // I = m.GetModule().Function1() → value is the outer Function1() Call over two
    // member accesses and the inner GetModule() call: 2 Call nodes, 2 MemberAccess.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Call { .. })), 2, "two Call nodes (GetModule(), Function1())");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::MemberAccess { .. })), 2, "two member accesses");
    let value = match find_node(&arena, |n| matches!(n, ExprNode::Assign { .. })).unwrap() {
        ExprNode::Assign { value, .. } => *value,
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, value), ExprNode::Call { .. }), "assigned value is the outer Function1() Call");
}

// ── String / line continuation ────────────────────────────────────────────────

#[test]
fn string_concat_continuation_before_ampersand() {
    let (_, arena, errors) = parse_stmts(b"Dim x\nx = \"foo\" _\n& \"bar\"");
    assert!(!errors, "String concat with line continuation before & must parse without errors");
    // The continuation must reconstruct a single "foo" & "bar" expression.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::BinOp { op: BinOpKind::Cat, .. })), 1, "one Cat BinOp");
    let (l, r) = match find_node(&arena, |n| matches!(n, ExprNode::BinOp { op: BinOpKind::Cat, .. })).unwrap() {
        ExprNode::BinOp { lhs, rhs, .. } => (*lhs, *rhs),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, l), ExprNode::Literal { lit: AstLit::Str(_) }), "lhs is \"foo\"");
    assert!(matches!(at(&arena, r), ExprNode::Literal { lit: AstLit::Str(_) }), "rhs is \"bar\"");
}

#[test]
fn string_concat_continuation_after_ampersand() {
    let (_, arena, errors) = parse_stmts(b"Dim x\nx = \"foo\" & _\n\"bar\"");
    assert!(!errors, "String concat with line continuation after & must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::BinOp { op: BinOpKind::Cat, .. })), 1, "one Cat BinOp");
    let (l, r) = match find_node(&arena, |n| matches!(n, ExprNode::BinOp { op: BinOpKind::Cat, .. })).unwrap() {
        ExprNode::BinOp { lhs, rhs, .. } => (*lhs, *rhs),
        _ => unreachable!(),
    };
    assert!(matches!(at(&arena, l), ExprNode::Literal { lit: AstLit::Str(_) }), "lhs is \"foo\"");
    assert!(matches!(at(&arena, r), ExprNode::Literal { lit: AstLit::Str(_) }), "rhs is \"bar\"");
}

// ── Class module header ───────────────────────────────────────────────────────

#[test]
fn class_module_full_header() {
    let src = b"\
VERSION 1.0 CLASS\n\
BEGIN\n\
  MultiUse = -1\n\
  Persistable = 0\n\
  DataBindingBehavior = 0\n\
  DataSourceBehavior = 0\n\
  MTSTransactionMode = 0\n\
END\n\
Attribute VB_Name = \"MyClass\"\n\
Attribute VB_GlobalNameSpace = False\n\
Attribute VB_Creatable = True\n\
Attribute VB_PredeclaredId = False\n\
Attribute VB_Exposed = False\n\
Sub F()\nEnd Sub";
    let (_, arena, errors) = parse_class_module(src);
    assert!(!errors, "Full class module header must parse without errors");
    // VERSION/BEGIN..END are skipped; 5 Attribute lines + Sub F must produce nodes.
    let attrs = count_nodes(&arena, |node| matches!(node, ExprNode::Attribute { .. }));
    assert_eq!(attrs, 5, "Five Attribute lines must build 5 Attribute nodes");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Sub, .. })),
        1,
        "exactly one Sub F ProcDecl after the header"
    );
    // The VERSION/BEGIN..END preamble must be fully skipped, not parsed as members.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 0, "BEGIN..END property lines are not parsed as Assigns");
}

#[test]
fn class_module_const_as_conditional_compilation_guard() {
    // pr578.vb: a module-level Const used as the condition in a #If block.
    let src = b"\
Const EvalDirective = 0\n\
Public Function EvalConstDirective() As Long\n\
#If EvalDirective Then\n\
  EvalConstDirective = 1\n\
#Else\n\
  EvalConstDirective = 0\n\
#End If\n\
End Function";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Const used as #If guard must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Function, .. })), 1, "one Function ProcDecl");
    // The module-level Const must be recognised; #If EvalDirective selects exactly
    // one of the two branches, so exactly one body Assign survives (not both, not zero).
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::DimItem { is_const: true, .. })), 1, "one module-level Const");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "#If selects one branch → one Assign in the body");
}

// ── Real-world form event handlers ───────────────────────────────────────────

#[test]
fn form_event_handler_with_with_block() {
    let src = b"\
Private Sub cmdHello_Click()\n\
  txtHello.Text = \"Hello World!\"\n\
  With txtHello\n\
    .Font = \"Arial\"\n\
    .FontSize = 16\n\
    .ForeColor = vbBlue\n\
  End With\n\
End Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Form event handler with With block must parse without errors");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::ProcDecl { kind: ProcKind::Sub, .. })), 1, "one event-handler Sub");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::With { .. })), 1, "one With block");
    // One Assign outside the With (txtHello.Text = ...) plus three inside (.Font,
    // .FontSize, .ForeColor) → 4 Assigns total.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 4, "1 outer + 3 With-member assignments");
    // The With body really contains its three member assignments.
    let body = match find_node(&arena, |n| matches!(n, ExprNode::With { .. })).unwrap() {
        ExprNode::With { body, .. } => *body,
        _ => unreachable!(),
    };
    let with_assigns = block_stmts(&arena, body).iter().filter(|s| matches!(at(&arena, **s), ExprNode::Assign { .. })).count();
    assert_eq!(with_assigns, 3, "the With body holds the three member-property assignments");
}

#[test]
fn multiple_form_event_handlers() {
    let src = b"\
Private Sub cmdClear_Click()\n\
  txtHello.Text = \"\"\n\
End Sub\n\
Private Sub cmdExit_Click()\n\
  End\n\
End Sub";
    let (_, arena, errors) = parse_module(src);
    assert!(!errors, "Multiple form event handlers must parse without errors");
    let n = count_nodes(&arena, |node| {
        matches!(node, ExprNode::ProcDecl { kind: ProcKind::Sub, .. })
    });
    assert_eq!(n, 2, "Two Sub event handlers must build 2 ProcDecl {{ kind: Sub }} nodes");
    // cmdClear sets txtHello.Text = "" → one Assign; cmdExit's `End` → one EndStmt.
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::Assign { .. })), 1, "one Assign (in cmdClear)");
    assert_eq!(count_nodes(&arena, |n| matches!(n, ExprNode::EndStmt)), 1, "the `End` statement → one EndStmt");
}

// ── oracle-verified gaps: variable/multi file channels and On Local Error ───────

#[test]
fn close_multiple_channels() {
    // VB6 (VB6.EXE /make oracle): a single Close may list several channels.
    let (_, arena, errors) = parse_stmts(b"Close #1, #2, #3");
    assert!(!errors, "Close #1, #2, #3 must parse without errors");
    let args = match find_node(&arena, |n| {
        matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Close, .. })
    })
    .unwrap()
    {
        ExprNode::FileIoStmt { channel, args, .. } => {
            assert!(channel.is_none(), "Close keeps its channels in args, not the channel slot");
            args.clone()
        }
        _ => unreachable!(),
    };
    assert_eq!(args.len(), 3, "three listed channels → 3 args");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Int(1) }), "first channel is #1");
    assert!(matches!(at(&arena, args[1]), ExprNode::Literal { lit: AstLit::Int(2) }), "second channel is #2");
    assert!(matches!(at(&arena, args[2]), ExprNode::Literal { lit: AstLit::Int(3) }), "third channel is #3");
}

#[test]
fn print_variable_channel() {
    // VB6 (VB6.EXE /make oracle): the file channel may be a variable, not a literal.
    // The channel must be captured as the variable's NameRef, never folded into an
    // opaque literal token.
    let (_, arena, errors) = parse_stmts(b"Dim f As Integer\nf = 1\nPrint #f, \"x\"");
    assert!(!errors, "Print #f, ... (variable channel) must parse without errors");
    let (channel, args) = match find_node(&arena, |n| {
        matches!(n, ExprNode::FileIoStmt { kind: FileIoKind::Print, .. })
    })
    .unwrap()
    {
        ExprNode::FileIoStmt { channel, args, .. } => (*channel, args.clone()),
        _ => unreachable!(),
    };
    let ch = channel.expect("Print #f records a channel");
    assert!(
        matches!(at(&arena, ch), ExprNode::NameRef { .. }),
        "the channel is the variable f (NameRef), not a literal",
    );
    assert_eq!(args.len(), 1, "one print item");
    assert!(matches!(at(&arena, args[0]), ExprNode::Literal { lit: AstLit::Str(_) }), "the item is the string \"x\"");
}

#[test]
fn on_local_error_goto_label() {
    // VB6 (VB6.EXE /make oracle): `On Local Error GoTo` is the explicit form of
    // `On Error GoTo` and must build the same handler node.
    let (_, arena, errors) = parse_stmts(b"On Local Error GoTo h\nExit Sub\nh:");
    assert!(!errors, "On Local Error GoTo must parse without errors");
    assert_eq!(
        count_nodes(&arena, |n| matches!(
            n,
            ExprNode::OnError { kind: OnErrorKind::Goto(LabelRef::Name(_)) }
        )),
        1,
        "On Local Error GoTo <label> → one OnError {{ Goto(Name) }}",
    );
}

#[test]
fn on_local_error_resume_next() {
    // VB6 (VB6.EXE /make oracle): `On Local Error Resume Next` is accepted.
    let (_, arena, errors) = parse_stmts(b"On Local Error Resume Next");
    assert!(!errors, "On Local Error Resume Next must parse without errors");
    assert_eq!(
        count_nodes(&arena, |n| matches!(n, ExprNode::OnError { kind: OnErrorKind::ResumeNext })),
        1,
        "On Local Error Resume Next → one OnError {{ ResumeNext }}",
    );
}
