use std::collections::HashMap;

use vb6_sema::sema::{
    BoundModule, BoundParam, BoundProc, BoundTypeDecl, BoundTypeMember, BoundVar, ExternalClass,
    NameResolution, ParamFlags, VbaType,
};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprArena, ExprNode, ProcKind};
use vb6_syntax::frontend::token::{Span, TypeSuffix};
use vb6_syntax::support::arena::NodeId;

use super::{lower_proc, lower_proc_with_classes};

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

fn udt_var(sym_id: u32, type_sym: u32) -> BoundVar {
    BoundVar {
        sym_id,
        vba_type: VbaType::UserDefined(type_sym),
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

/// `Type Point : X As Long : Y As Long : End Type`, as a `BoundTypeDecl`.
fn point_type_decl(type_sym: u32, x_sym: u32, y_sym: u32) -> BoundTypeDecl {
    BoundTypeDecl {
        sym_id: type_sym,
        members: vec![
            BoundTypeMember { sym_id: x_sym, vba_type: VbaType::Long, name_span: Span::default() },
            BoundTypeMember { sym_id: y_sym, vba_type: VbaType::Long, name_span: Span::default() },
        ],
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

fn member_access_node(ea: &mut ExprArena, base: NodeId, member: u32) -> NodeId {
    ea.alloc(ExprNode::MemberAccess { base, member, bang: false })
}

fn int_lit_node(ea: &mut ExprArena, v: i32) -> NodeId {
    ea.alloc(ExprNode::Literal { lit: AstLit::Int(v) })
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

// ── UDT (Type...End Type) field access ───────────────────────────────────────
//
// `Type Point : X As Long : Y As Long : End Type` — the milestone-1 fixture's
// declaration. Frame: `t` (2 Long fields, uniform size 4) allocates first —
// base -140 (0xff74), so X (field 0) is at -140, Y (field 1) at -136 — then
// `y As Long` at -144 (0xff70).

#[test]
fn lower_udt_field_load_matches_isolated_probe() {
    // `y = t.X` in isolation (no preceding store) — oracle-confirmed
    // (e2e_udt_field_scalar_access): `6c 74 ff` (load t.X at -140) then
    // `71 70 ff` (store y at -144).
    let mut ea = ExprArena::new();
    let t = name_ref(&mut ea, 0);
    let field_x = member_access_node(&mut ea, t, 10 /* sym for X */);
    let y = name_ref(&mut ea, 1);
    let stmt = assign_node(&mut ea, y, field_x);
    let body = block_node(&mut ea, vec![stmt]);

    let mut resolutions = HashMap::new();
    resolutions.insert(t.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });
    resolutions.insert(y.0, NameResolution::Local { proc_idx: 0, local_idx: 1 });

    let mut types = HashMap::new();
    types.insert(field_x.0, VbaType::Long);

    let module = BoundModule {
        procs: vec![make_proc(
            vec![udt_var(0, /* type_sym */ 100), long_var(1)],
            vec![],
            body,
        )],
        type_decls: vec![point_type_decl(100, 10, 11)],
        resolutions,
        types,
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0x6c, 0x74, 0xff, 0x71, 0x70, 0xff]
    );
}

#[test]
fn lower_udt_field_store_matches_current_pipeline_output() {
    // `t.X = 1` in isolation — goes through the real 0x2c/resolver chain (a
    // UDT field has no bypass frame slot). Oracle-confirmed
    // (e2e_udt_field_scalar_access): `f5 01 00 00 00` (push Long literal 1),
    // `71 74 ff` (store t.X at -140) — no trailing reload. An earlier version
    // of `emit_reference`'s nOp=4/f_flags&0x40 remap path (opcode-index
    // 0x1f2 → byte 0x71 unconditionally remapped to 0x1e2 → byte 0x6c) added
    // a spurious extra `6c 74 ff`; the real compiler's output — captured here
    // for the first time by real source reaching this path — proved that
    // wrong, and the remap arm was removed.
    let mut ea = ExprArena::new();
    let t = name_ref(&mut ea, 0);
    let field_x = member_access_node(&mut ea, t, 10 /* sym for X */);
    let one = int_lit_node(&mut ea, 1);
    let stmt = assign_node(&mut ea, field_x, one);
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(vec![udt_var(0, 100)], vec![], body)],
        type_decls: vec![point_type_decl(100, 10, 11)],
        resolutions: HashMap::from([(t.0, NameResolution::Local { proc_idx: 0, local_idx: 0 })]),
        types: HashMap::from([(field_x.0, VbaType::Long)]),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_proc(&module, 0, &ea, 0x0008).unwrap(),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x74, 0xff]
    );
}

// ── Class-instance local frame layout (object-base unblock) ─────────────────

#[test]
fn class_instance_local_gets_a_plain_4byte_object_slot() {
    // `Dim o As New Class1` allocates a plain 4-byte object-reference slot
    // (ctx 0, like any Object local) — NOT an embedded struct like a UDT.
    // `Dim x As Long` declared right after must land at the SAME offset a
    // plain 4-byte-typed local would (-140), proving the class local doesn't
    // consume UDT-style struct space.
    let mut ea = ExprArena::new();
    let body = block_node(&mut ea, vec![]);

    let o_var = udt_var(0, 100); // reuse: sym_id 0, type_sym 100 (VbaType::UserDefined(100))
    let x_var = long_var(1);

    let mut class_field_info = HashMap::new();
    class_field_info.insert(100u32, ExternalClass { fields: vec![("F".to_string(), VbaType::Long)] });

    let module = BoundModule {
        procs: vec![make_proc(vec![o_var, x_var], vec![], body)],
        class_field_info,
        ..BoundModule::default()
    };

    let known_classes: HashMap<String, ExternalClass> = HashMap::new();
    let bytes = lower_proc_with_classes(&module, 0, &ea, 0x0008, &known_classes).unwrap();
    // Empty body: no bytes emitted, but lowering must succeed (frame builds
    // without hitting the UDT `type_decls` lookup / UnsupportedType error a
    // class-typed local would otherwise trigger).
    assert_eq!(bytes, Vec::<u8>::new());
}
