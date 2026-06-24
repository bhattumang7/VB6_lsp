use std::collections::HashMap;

use vb6_sema::sema::{
    BoundModule, BoundParam, BoundProc, BoundVar, NameResolution, ParamFlags, VbaType,
};
use vb6_syntax::frontend::ast::{BinOpKind, ExprArena, ExprNode, ProcKind};
use vb6_syntax::frontend::token::{Span, TypeSuffix};
use vb6_syntax::support::arena::NodeId;

use super::lower_proc;

// ── Helpers ───────────────────────────────────────────────────────────────────

fn long_var(sym_id: u32) -> BoundVar {
    BoundVar {
        sym_id,
        vba_type: VbaType::Long,
        is_const: false,
        const_value: None,
        const_lit: None,
        fixed_string_len: None,
        array_dims: None,
        is_static: false,
        is_public: false,
        name_span: Span::default(),
    }
}

fn integer_var(sym_id: u32) -> BoundVar {
    BoundVar {
        sym_id,
        vba_type: VbaType::Integer,
        is_const: false,
        const_value: None,
        const_lit: None,
        fixed_string_len: None,
        array_dims: None,
        is_static: false,
        is_public: false,
        name_span: Span::default(),
    }
}

fn byval_long_param(sym_id: u32) -> BoundParam {
    BoundParam {
        sym_id,
        vba_type: VbaType::Long,
        flags: ParamFlags {
            by_val: true,
            by_ref: false,
            optional: false,
            is_array: false,
            param_array: false,
        },
        name_span: Span::default(),
    }
}

fn byref_long_param(sym_id: u32) -> BoundParam {
    BoundParam {
        sym_id,
        vba_type: VbaType::Long,
        flags: ParamFlags {
            by_val: false,
            by_ref: true,
            optional: false,
            is_array: false,
            param_array: false,
        },
        name_span: Span::default(),
    }
}

fn make_proc(
    locals: Vec<BoundVar>,
    params: Vec<BoundParam>,
    body: NodeId,
) -> BoundProc {
    BoundProc {
        sym_id: 0,
        kind: ProcKind::Sub,
        params,
        ret_type: VbaType::Variant,
        locals,
        body: body.0,
        is_public: false,
        name_span: Span::default(),
    }
}

fn name_ref(ea: &mut ExprArena, sym: u32) -> NodeId {
    ea.alloc(ExprNode::NameRef { sym, suffix: TypeSuffix::None })
}

fn binop_node(ea: &mut ExprArena, op: BinOpKind, lhs: NodeId, rhs: NodeId) -> NodeId {
    ea.alloc(ExprNode::BinOp { op, lhs, rhs })
}

fn assign_node(ea: &mut ExprArena, target: NodeId, value: NodeId) -> NodeId {
    ea.alloc(ExprNode::Assign { target, value })
}

fn block_node(ea: &mut ExprArena, stmts: Vec<NodeId>) -> NodeId {
    ea.alloc(ExprNode::Block { stmts })
}

// ── Tests ─────────────────────────────────────────────────────────────────────

// Oracle byte sequence for `r = a + b` with three Long locals.
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-144 (0xff70).
// Bytes: load a, load b, Long-Add, store r.
#[test]
fn lower_local_long_add_matches_oracle() {
    let mut ea = ExprArena::new();
    let a = name_ref(&mut ea, 0);   // NodeId(0)
    let b = name_ref(&mut ea, 1);   // NodeId(1)
    let add = binop_node(&mut ea, BinOpKind::Add, a, b);   // NodeId(2)
    let r = name_ref(&mut ea, 2);   // NodeId(3)
    let stmt = assign_node(&mut ea, r, add);  // NodeId(4)
    let body = block_node(&mut ea, vec![stmt]);  // NodeId(5)

    let mut resolutions = HashMap::new();
    resolutions.insert(a.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });
    resolutions.insert(b.0, NameResolution::Local { proc_idx: 0, local_idx: 1 });
    resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 2 });

    let mut types = HashMap::new();
    types.insert(add.0, VbaType::Long);

    let module = BoundModule {
        procs: vec![make_proc(
            vec![long_var(0), long_var(1), long_var(2)],
            vec![],
            body,
        )],
        resolutions,
        types,
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff]
    );
}

// Oracle byte sequence for `r = a - b` with three Long locals.
// Frame: a=-136, b=-140, r=-144. Bytes: load a, load b, Long-Sub, store r.
#[test]
fn lower_local_long_sub_matches_oracle() {
    let mut ea = ExprArena::new();
    let a = name_ref(&mut ea, 0);
    let b = name_ref(&mut ea, 1);
    let sub = binop_node(&mut ea, BinOpKind::Sub, a, b);
    let r = name_ref(&mut ea, 2);
    let stmt = assign_node(&mut ea, r, sub);
    let body = block_node(&mut ea, vec![stmt]);

    let mut resolutions = HashMap::new();
    resolutions.insert(a.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });
    resolutions.insert(b.0, NameResolution::Local { proc_idx: 0, local_idx: 1 });
    resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 2 });

    let mut types = HashMap::new();
    types.insert(sub.0, VbaType::Long);

    let module = BoundModule {
        procs: vec![make_proc(
            vec![long_var(0), long_var(1), long_var(2)],
            vec![],
            body,
        )],
        resolutions,
        types,
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xae, 0x71, 0x70, 0xff]
    );
}

// Oracle byte sequence for `r = (a = b)` — Long comparison stored into Integer.
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-142 (0xff72).
// Bytes: load a, load b, Long-Eq, Integer-store r.
// (Integer frame size = 2, so r follows b at -140 - 2 = -142.)
#[test]
fn lower_local_long_eq_into_integer_matches_oracle() {
    let mut ea = ExprArena::new();
    let a = name_ref(&mut ea, 0);
    let b = name_ref(&mut ea, 1);
    let eq = binop_node(&mut ea, BinOpKind::Eq, a, b);
    let r = name_ref(&mut ea, 2);
    let stmt = assign_node(&mut ea, r, eq);
    let body = block_node(&mut ea, vec![stmt]);

    let mut resolutions = HashMap::new();
    resolutions.insert(a.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });
    resolutions.insert(b.0, NameResolution::Local { proc_idx: 0, local_idx: 1 });
    resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 2 });

    // Comparison nodes do not carry a type in BoundModule.types (type_tag = 0).
    let module = BoundModule {
        procs: vec![make_proc(
            vec![long_var(0), long_var(1), integer_var(2)],
            vec![],
            body,
        )],
        resolutions,
        types: HashMap::new(),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc7, 0x70, 0x72, 0xff]
    );
}

// Oracle: ByVal Long param `p` at frame +12, local `r` Long at -136.
// `r = p` → load ByVal param [0x6c, 0x0c, 0x00], store Long local [0x71, 0x78, 0xff].
#[test]
fn lower_byval_long_param_load_matches_oracle() {
    let mut ea = ExprArena::new();
    let p = name_ref(&mut ea, 0);   // ByVal param p
    let r = name_ref(&mut ea, 1);   // local r
    let stmt = assign_node(&mut ea, r, p);
    let body = block_node(&mut ea, vec![stmt]);

    let mut resolutions = HashMap::new();
    resolutions.insert(p.0, NameResolution::Param { proc_idx: 0, param_idx: 0 });
    resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });

    let module = BoundModule {
        procs: vec![make_proc(
            vec![long_var(1)],
            vec![byval_long_param(0)],
            body,
        )],
        resolutions,
        types: HashMap::new(),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x6c, 0x0c, 0x00, 0x71, 0x78, 0xff]
    );
}

// Oracle: ByRef Long param `p` at frame +12, local `r` Long at -136.
// `r = p` → ByRef load [0x80, 0x0c, 0x00] (RT_LOAD_BY_CTX[2]+0x14 = 0x6c+0x14 = 0x80),
// store Long local [0x71, 0x78, 0xff].
#[test]
fn lower_byref_long_param_load_matches_oracle() {
    let mut ea = ExprArena::new();
    let p = name_ref(&mut ea, 0);
    let r = name_ref(&mut ea, 1);
    let stmt = assign_node(&mut ea, r, p);
    let body = block_node(&mut ea, vec![stmt]);

    let mut resolutions = HashMap::new();
    resolutions.insert(p.0, NameResolution::Param { proc_idx: 0, param_idx: 0 });
    resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });

    let module = BoundModule {
        procs: vec![make_proc(
            vec![long_var(1)],
            vec![byref_long_param(0)],
            body,
        )],
        resolutions,
        types: HashMap::new(),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x80, 0x0c, 0x00, 0x71, 0x78, 0xff]
    );
}

// Oracle: module-level Long global `g` at field_offset=0, module_desc=0x0008.
// Local Long `r` at frame -136.
// `r = g` → global load [0x94, 0x08, 0x00, 0x00, 0x00]
//           (RT_LOAD_BY_CTX[2]+0x28 = 0x6c+0x28 = 0x94, then module_desc LE, field_offset LE),
//           store Long local [0x71, 0x78, 0xff].
#[test]
fn lower_global_long_load_matches_oracle() {
    let mut ea = ExprArena::new();
    let g = name_ref(&mut ea, 0);
    let r = name_ref(&mut ea, 1);
    let stmt = assign_node(&mut ea, r, g);
    let body = block_node(&mut ea, vec![stmt]);

    let mut resolutions = HashMap::new();
    resolutions.insert(g.0, NameResolution::ModuleVar(0));
    resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });

    let module = BoundModule {
        procs: vec![make_proc(vec![long_var(1)], vec![], body)],
        module_vars: vec![long_var(0)],
        resolutions,
        types: HashMap::new(),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x94, 0x08, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff]
    );
}

// Two sequential assignment statements in a Block.
// `a = b` then `b = a` (Long locals) swaps their values through two stores.
// Frame: a=-136, b=-140.
// Bytes: load b, store a, load a, store b.
#[test]
fn lower_two_sequential_assigns() {
    let mut ea = ExprArena::new();
    let a0 = name_ref(&mut ea, 0);  // load a
    let b0 = name_ref(&mut ea, 1);  // store target b (in first stmt)
    let b1 = name_ref(&mut ea, 1);  // load b
    let a1 = name_ref(&mut ea, 0);  // store target a (in second stmt)
    let stmt1 = assign_node(&mut ea, b0, a0);  // b = a
    let stmt2 = assign_node(&mut ea, a1, b1);  // a = b
    let body = block_node(&mut ea, vec![stmt1, stmt2]);

    let mut resolutions = HashMap::new();
    // b = a: load a (local 0), store b (local 1)
    resolutions.insert(a0.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });
    resolutions.insert(b0.0, NameResolution::Local { proc_idx: 0, local_idx: 1 });
    // a = b: load b (local 1), store a (local 0)
    resolutions.insert(b1.0, NameResolution::Local { proc_idx: 0, local_idx: 1 });
    resolutions.insert(a1.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });

    let module = BoundModule {
        procs: vec![make_proc(vec![long_var(0), long_var(1)], vec![], body)],
        resolutions,
        types: HashMap::new(),
        ..BoundModule::default()
    };

    // b = a: load a [0x6c, 0x78, 0xff], store b [0x71, 0x74, 0xff]
    // a = b: load b [0x6c, 0x74, 0xff], store a [0x71, 0x78, 0xff]
    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[
            0x6c, 0x78, 0xff, 0x71, 0x74, 0xff,
            0x6c, 0x74, 0xff, 0x71, 0x78, 0xff,
        ]
    );
}
