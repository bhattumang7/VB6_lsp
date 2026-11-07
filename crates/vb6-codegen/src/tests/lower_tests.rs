use std::collections::HashMap;

use vb6_sema::sema::{
    AccessorKind, BoundModule, BoundParam, BoundProc, BoundTypeDecl, BoundTypeMember, BoundVar,
    ClassMemberSlot, ExternalClass, NameResolution, ParamFlags, ResolvedClassMember, VbaType,
};
use vb6_syntax::frontend::ast::{AstLit, BinOpKind, ExprArena, ExprNode, ProcKind};
use vb6_syntax::frontend::token::{Span, TypeSuffix};
use vb6_syntax::support::arena::NodeId;

use super::{lower_proc, lower_proc_with_classes, LowerError};

// ── Helpers ───────────────────────────────────────────────────────────────────

fn var_of(sym_id: u32, vba_type: VbaType) -> BoundVar {
    BoundVar {
        sym_id,
        vba_type,
        is_const: false,
        const_value: None,
        const_lit: None,
        fixed_string_len: None,
        array_dims: None,
        is_static: false,
        is_public: false,
        is_new: false,
        name_span: Span::default(),
    }
}

fn long_var(sym_id: u32) -> BoundVar {
    var_of(sym_id, VbaType::Long)
}

fn udt_var(sym_id: u32, type_sym: u32) -> BoundVar {
    var_of(sym_id, VbaType::UserDefined(type_sym))
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

fn double_var(sym_id: u32) -> BoundVar {
    var_of(sym_id, VbaType::Double)
}

fn string_var(sym_id: u32) -> BoundVar {
    var_of(sym_id, VbaType::String)
}

/// `Dim <sym_id> As Object` — a plain object-typed local (no specific class).
fn object_var(sym_id: u32) -> BoundVar {
    var_of(sym_id, VbaType::Object)
}

fn integer_var(sym_id: u32) -> BoundVar {
    var_of(sym_id, VbaType::Integer)
}

fn long_param(sym_id: u32, by_val: bool) -> BoundParam {
    BoundParam {
        sym_id,
        vba_type: VbaType::Long,
        flags: ParamFlags {
            by_val,
            by_ref: !by_val,
            optional: false,
            is_array: false,
            param_array: false,
        },
        name_span: Span::default(),
    }
}

fn byval_long_param(sym_id: u32) -> BoundParam {
    long_param(sym_id, true)
}

fn byref_long_param(sym_id: u32) -> BoundParam {
    long_param(sym_id, false)
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
    ea.alloc(ExprNode::MemberAccess { base, member, bang: false, bracketed: false })
}

fn int_lit_node(ea: &mut ExprArena, v: i32) -> NodeId {
    ea.alloc(ExprNode::Literal { lit: AstLit::Int(v) })
}

fn set_assign_node(ea: &mut ExprArena, target: NodeId, value: NodeId) -> NodeId {
    ea.alloc(ExprNode::SetAssign { target, value })
}

fn new_node(ea: &mut ExprArena, class_sym: u32) -> NodeId {
    let type_spec = ea.alloc(ExprNode::UserType { name: class_sym, child: None });
    ea.alloc(ExprNode::New { type_spec })
}

fn nothing_node(ea: &mut ExprArena) -> NodeId {
    ea.alloc(ExprNode::Nothing)
}

/// A local of the procedure under test, by its index in the locals list.
fn local(local_idx: usize) -> NameResolution {
    NameResolution::Local { proc_idx: 0, local_idx }
}

/// A parameter of the procedure under test, by its index in the parameter list.
fn param(param_idx: usize) -> NameResolution {
    NameResolution::Param { proc_idx: 0, param_idx }
}

fn resolutions_of(entries: &[(NodeId, NameResolution)]) -> HashMap<u32, NameResolution> {
    entries.iter().map(|(node, res)| (node.0, res.clone())).collect()
}

fn types_of(entries: &[(NodeId, VbaType)]) -> HashMap<u32, VbaType> {
    entries.iter().map(|(node, ty)| (node.0, ty.clone())).collect()
}

/// Lowers procedure 0 of `module` with no external classes.
fn lower_ok(module: &BoundModule, ea: &ExprArena) -> Vec<u8> {
    lower_proc(module, 0, ea, 0x0008).unwrap()
}

fn lower_with_known_classes(
    module: &BoundModule,
    ea: &ExprArena,
) -> Result<Vec<u8>, LowerError> {
    let known_classes: HashMap<String, ExternalClass> = HashMap::new();
    lower_proc_with_classes(module, 0, ea, 0x0008, &known_classes)
}

fn lower_classes_ok(module: &BoundModule, ea: &ExprArena) -> Vec<u8> {
    lower_with_known_classes(module, ea).unwrap()
}

/// Class symbol 100 with the given members.
fn class_100(members: Vec<ClassMemberSlot>) -> HashMap<u32, ExternalClass> {
    HashMap::from([(100u32, ExternalClass { members })])
}

fn field_slot(name: &str) -> ClassMemberSlot {
    ClassMemberSlot::Field { name: name.to_string(), vba_type: VbaType::Long, is_object: false }
}

fn property_slot(name: &str, vba_type: VbaType, kind: AccessorKind) -> ClassMemberSlot {
    ClassMemberSlot::PropertyAccessor { name: name.to_string(), vba_type, kind }
}

fn resolved_member(
    get_slot: Option<u16>,
    let_slot: Option<u16>,
    is_property: bool,
) -> ResolvedClassMember {
    ResolvedClassMember {
        get_slot,
        let_slot,
        set_slot: None,
        method_slot: None,
        method_ret_type: None,
        method_params: Vec::new(),
        is_property,
    }
}

/// `<target> = <local 0>.P` (or `Set` when `use_set`), where `P` is a Property
/// Get returning `prop_ty` and the target is `target_var` (local 1).
fn lower_property_get(prop_ty: VbaType, target_var: BoundVar, use_set: bool) -> Vec<u8> {
    let mut ea = ExprArena::new();
    let o = name_ref(&mut ea, 0);
    let get_access = member_access_node(&mut ea, o, 10 /* sym for P */);
    let x = name_ref(&mut ea, 1);
    let stmt = if use_set {
        set_assign_node(&mut ea, x, get_access)
    } else {
        assign_node(&mut ea, x, get_access)
    };
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(vec![udt_var(0, 100), target_var], vec![], body)],
        class_field_info: class_100(vec![property_slot("P", prop_ty.clone(), AccessorKind::Get)]),
        class_member_slots: HashMap::from([(get_access.0, resolved_member(Some(0x1c), None, true))]),
        resolutions: resolutions_of(&[(o, local(0)), (x, local(1))]),
        types: types_of(&[(get_access, prop_ty)]),
        ..BoundModule::default()
    };
    lower_classes_ok(&module, &ea)
}

/// `<local 0>.P = <literal>` where `P` is a Property Let of type `prop_ty` and
/// the literal has its own natural type `lit_ty`.
fn lower_property_let(prop_ty: VbaType, lit: AstLit, lit_ty: VbaType) -> Vec<u8> {
    let mut ea = ExprArena::new();
    let o = name_ref(&mut ea, 0);
    let let_access = member_access_node(&mut ea, o, 10 /* sym for P */);
    let value = ea.alloc(ExprNode::Literal { lit });
    let stmt = assign_node(&mut ea, let_access, value);
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(vec![udt_var(0, 100)], vec![], body)],
        class_field_info: class_100(vec![property_slot("P", prop_ty.clone(), AccessorKind::Let)]),
        class_member_slots: HashMap::from([(let_access.0, resolved_member(None, Some(0x1c), true))]),
        resolutions: resolutions_of(&[(o, local(0))]),
        // The literal's own natural type: the real binder always populates it,
        // so this hand-built module must too.
        types: types_of(&[(let_access, prop_ty), (value, lit_ty)]),
        ..BoundModule::default()
    };
    lower_classes_ok(&module, &ea)
}

/// What a `Set o = ...` statement assigns in [`lower_set_into_object`].
enum SetSource {
    New(u32),
    Local(usize),
    Nothing,
}

/// `Set <local 0> = <source>` where every local is an object of class 100 and
/// `classes` are the class symbols the module knows about.
fn lower_set_into_object(
    locals: Vec<BoundVar>,
    classes: &[u32],
    source: SetSource,
) -> Result<Vec<u8>, LowerError> {
    let mut ea = ExprArena::new();
    let o = name_ref(&mut ea, 0);
    let mut resolutions = vec![(o, local(0))];
    let value = match source {
        SetSource::New(class_sym) => new_node(&mut ea, class_sym),
        SetSource::Local(idx) => {
            let other = name_ref(&mut ea, idx as u32);
            resolutions.push((other, local(idx)));
            other
        }
        SetSource::Nothing => nothing_node(&mut ea),
    };
    let stmt = set_assign_node(&mut ea, o, value);
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(locals, vec![], body)],
        class_field_info: classes.iter().map(|c| (*c, ExternalClass::default())).collect(),
        resolutions: resolutions_of(&resolutions),
        ..BoundModule::default()
    };
    lower_with_known_classes(&module, &ea)
}

/// `r = a <op> b` over three locals; `result_type` is the type recorded for
/// the operator node (comparisons carry none).
fn lower_binop_assign(
    op: BinOpKind,
    locals: Vec<BoundVar>,
    result_type: Option<VbaType>,
) -> Vec<u8> {
    let mut ea = ExprArena::new();
    let a = name_ref(&mut ea, 0);
    let b = name_ref(&mut ea, 1);
    let operation = binop_node(&mut ea, op, a, b);
    let r = name_ref(&mut ea, 2);
    let stmt = assign_node(&mut ea, r, operation);
    let body = block_node(&mut ea, vec![stmt]);

    let types = match result_type {
        Some(ty) => types_of(&[(operation, ty)]),
        None => HashMap::new(),
    };
    let module = BoundModule {
        procs: vec![make_proc(locals, vec![], body)],
        resolutions: resolutions_of(&[(a, local(0)), (b, local(1)), (r, local(2))]),
        types,
        ..BoundModule::default()
    };
    lower_ok(&module, &ea)
}

/// `r = <source>` where `r` is a Long local (local 0, sym 1) and the source is
/// a name resolving to `source_res`.
fn lower_read_into_long_local(
    source_res: NameResolution,
    params: Vec<BoundParam>,
    module_vars: Vec<BoundVar>,
) -> Vec<u8> {
    let mut ea = ExprArena::new();
    let p = name_ref(&mut ea, 0);
    let r = name_ref(&mut ea, 1);
    let stmt = assign_node(&mut ea, r, p);
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(vec![long_var(1)], params, body)],
        module_vars,
        resolutions: resolutions_of(&[(p, source_res), (r, local(0))]),
        ..BoundModule::default()
    };
    lower_ok(&module, &ea)
}

/// `t.X` / `t.X = 1` over a `Point` UDT local `t`. With `store` false the
/// statement is `y = t.X` (y a Long local); with `store` true it is `t.X = 1`.
fn lower_udt_field(store: bool) -> Vec<u8> {
    let mut ea = ExprArena::new();
    let t = name_ref(&mut ea, 0);
    let field_x = member_access_node(&mut ea, t, 10 /* sym for X */);
    let mut resolutions = vec![(t, local(0))];
    let (stmt, locals) = if store {
        let one = int_lit_node(&mut ea, 1);
        (assign_node(&mut ea, field_x, one), vec![udt_var(0, 100)])
    } else {
        let y = name_ref(&mut ea, 1);
        resolutions.push((y, local(1)));
        (assign_node(&mut ea, y, field_x), vec![udt_var(0, 100), long_var(1)])
    };
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(locals, vec![], body)],
        type_decls: vec![point_type_decl(100, 10, 11)],
        resolutions: resolutions_of(&resolutions),
        types: types_of(&[(field_x, VbaType::Long)]),
        ..BoundModule::default()
    };
    lower_ok(&module, &ea)
}

// ── Tests ─────────────────────────────────────────────────────────────────────

// Oracle byte sequence for `r = a + b` with three Long locals.
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-144 (0xff70).
// Bytes: load a, load b, Long-Add, store r.
#[test]
fn lower_local_long_add_matches_oracle() {
    let bytes = lower_binop_assign(
        BinOpKind::Add,
        vec![long_var(0), long_var(1), long_var(2)],
        Some(VbaType::Long),
    );
    assert_eq!(bytes, &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xaa, 0x71, 0x70, 0xff, 0x14]);
}

// Oracle byte sequence for `r = a - b` with three Long locals.
// Frame: a=-136, b=-140, r=-144. Bytes: load a, load b, Long-Sub, store r.
#[test]
fn lower_local_long_sub_matches_oracle() {
    let bytes = lower_binop_assign(
        BinOpKind::Sub,
        vec![long_var(0), long_var(1), long_var(2)],
        Some(VbaType::Long),
    );
    assert_eq!(bytes, &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xae, 0x71, 0x70, 0xff, 0x14]);
}

// Oracle byte sequence for `r = (a = b)` — Long comparison stored into Integer.
// Frame: a=-136 (0xff78), b=-140 (0xff74), r=-142 (0xff72).
// Bytes: load a, load b, Long-Eq, Integer-store r.
// (Integer frame size = 2, so r follows b at -140 - 2 = -142.)
// Comparison nodes do not carry a type in BoundModule.types (type_tag = 0).
#[test]
fn lower_local_long_eq_into_integer_matches_oracle() {
    let bytes = lower_binop_assign(
        BinOpKind::Eq,
        vec![long_var(0), long_var(1), integer_var(2)],
        None,
    );
    assert_eq!(bytes, &[0x6c, 0x78, 0xff, 0x6c, 0x74, 0xff, 0xc7, 0x70, 0x72, 0xff, 0x14]);
}

// Oracle: ByVal Long param `p` at frame +12, local `r` Long at -136.
// `r = p` → load ByVal param [0x6c, 0x0c, 0x00], store Long local [0x71, 0x78, 0xff].
#[test]
fn lower_byval_long_param_load_matches_oracle() {
    let bytes = lower_read_into_long_local(param(0), vec![byval_long_param(0)], vec![]);
    assert_eq!(bytes, &[0x6c, 0x0c, 0x00, 0x71, 0x78, 0xff, 0x14]);
}

// Oracle: ByRef Long param `p` at frame +12, local `r` Long at -136.
// `r = p` → ByRef load [0x80, 0x0c, 0x00] (RT_LOAD_BY_CTX[2]+0x14 = 0x6c+0x14 = 0x80),
// store Long local [0x71, 0x78, 0xff].
#[test]
fn lower_byref_long_param_load_matches_oracle() {
    let bytes = lower_read_into_long_local(param(0), vec![byref_long_param(0)], vec![]);
    assert_eq!(bytes, &[0x80, 0x0c, 0x00, 0x71, 0x78, 0xff, 0x14]);
}

// Oracle: module-level Long global `g` at field_offset=0, module_desc=0x0008.
// Local Long `r` at frame -136.
// `r = g` → global load [0x94, 0x08, 0x00, 0x00, 0x00]
//           (RT_LOAD_BY_CTX[2]+0x28 = 0x6c+0x28 = 0x94, then module_desc LE, field_offset LE),
//           store Long local [0x71, 0x78, 0xff].
#[test]
fn lower_global_long_load_matches_oracle() {
    let bytes =
        lower_read_into_long_local(NameResolution::ModuleVar(0), vec![], vec![long_var(0)]);
    assert_eq!(bytes, &[0x94, 0x08, 0x00, 0x00, 0x00, 0x71, 0x78, 0xff, 0x14]);
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

    let module = BoundModule {
        procs: vec![make_proc(vec![long_var(0), long_var(1)], vec![], body)],
        // b = a: load a (local 0), store b (local 1); a = b: load b (local 1), store a (local 0)
        resolutions: resolutions_of(&[(a0, local(0)), (b0, local(1)), (b1, local(1)), (a1, local(0))]),
        ..BoundModule::default()
    };

    // b = a: load a [0x6c, 0x78, 0xff], store b [0x71, 0x74, 0xff]
    // a = b: load b [0x6c, 0x74, 0xff], store a [0x71, 0x78, 0xff]
    assert_eq!(
        lower_ok(&module, &ea),
        &[
            0x6c, 0x78, 0xff, 0x71, 0x74, 0xff,
            0x6c, 0x74, 0xff, 0x71, 0x78, 0xff,
            0x14,
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
    assert_eq!(lower_udt_field(false), &[0x6c, 0x74, 0xff, 0x71, 0x70, 0xff, 0x14]);
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
    assert_eq!(
        lower_udt_field(true),
        &[0xf5, 0x01, 0x00, 0x00, 0x00, 0x71, 0x74, 0xff, 0x14]
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

    let module = BoundModule {
        // sym_id 0, type_sym 100 (VbaType::UserDefined(100))
        procs: vec![make_proc(vec![udt_var(0, 100), long_var(1)], vec![], body)],
        class_field_info: class_100(vec![field_slot("F")]),
        ..BoundModule::default()
    };

    // Empty body: no statement bytes, but lowering must succeed (frame builds
    // without hitting the UDT `type_decls` lookup / UnsupportedType error a
    // class-typed local would otherwise trigger) — and the proc still gets
    // its unconditional trailing implicit-return `0x14`.
    assert_eq!(lower_classes_ok(&module, &ea), vec![0x14]);
}

// ── Class-member vtable dispatch (0x24 resolve-object + 0x0d vtable-call) ────
//
// Oracle-captured for `Class1 : Public F As Long` accessed from
// `Dim o As New Class1` as `o.F = 1` then `x = o.F`. Frame: o (Object, 4
// bytes) at -136, x (Long) at -140, the hidden class-Get temp at -144.
// Full raw bytes (re_lab recon, byte-exact):
//   f5 01 00 00 00 04 78 ff 24 00 00 0d 20 00 01 00
//   04 70 ff 04 78 ff 24 00 00 0d 1c 00 01 00 6c 70 ff 71 74 ff

#[test]
fn class_field_store_then_load_matches_oracle_recon() {
    let mut ea = ExprArena::new();
    let o = name_ref(&mut ea, 0);
    let field_f_store = member_access_node(&mut ea, o, 10 /* sym for F */);
    let one = int_lit_node(&mut ea, 1);
    let stmt1 = assign_node(&mut ea, field_f_store, one);

    let o2 = name_ref(&mut ea, 0);
    let field_f_load = member_access_node(&mut ea, o2, 10);
    let x = name_ref(&mut ea, 1);
    let stmt2 = assign_node(&mut ea, x, field_f_load);

    let body = block_node(&mut ea, vec![stmt1, stmt2]);

    let resolved_f = resolved_member(Some(0x1c), Some(0x20), false);
    let module = BoundModule {
        procs: vec![make_proc(vec![udt_var(0, 100), long_var(1)], vec![], body)],
        class_field_info: class_100(vec![field_slot("F")]),
        class_member_slots: HashMap::from([
            (field_f_store.0, resolved_f.clone()),
            (field_f_load.0, resolved_f),
        ]),
        resolutions: resolutions_of(&[(o, local(0)), (o2, local(0)), (x, local(1))]),
        types: types_of(&[(field_f_store, VbaType::Long), (field_f_load, VbaType::Long)]),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_classes_ok(&module, &ea),
        &[
            0xf5, 0x01, 0x00, 0x00, 0x00, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x20, 0x00,
            0x01, 0x00, 0x04, 0x70, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x1c, 0x00,
            0x01, 0x00, 0x6c, 0x70, 0xff, 0x71, 0x74, 0xff, 0x14,
        ]
    );
}

// ── Property Get value type (Double) ─────────────────────────────────────────
//
// Oracle-captured (`c1_get_double`; see the `e2e_class_property_get_double`
// fixture for the full byte-exact sequence): a Property Get returning
// `Double` reads its out-param temp back with `0x6f` (8-byte-sized), not
// `0x6c` (4-byte `Long`/`Object`) — the class-member scratch temp's frame
// size must track the property's real return type.

#[test]
fn class_property_get_double_matches_oracle_bytes() {
    assert_eq!(
        lower_property_get(VbaType::Double, double_var(1), false),
        &[
            0x04, 0x68, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x1c, 0x00, 0x01, 0x00,
            0x6f, 0x68, 0xff, 0x74, 0x70, 0xff, 0x14,
        ]
    );
}

#[test]
fn class_property_get_mixed_types_in_one_proc_uses_separate_regions() {
    // Two Get accesses of DIFFERENT type-contexts (Long then Double) sharing
    // the same proc: an earlier pass of this port modeled the class-member
    // scratch area as ONE shared temp for the whole proc and gated this
    // shape as ungrounded. Two fresh oracle captures this session (slice
    // #7's `e2e_class_mixed_double_get_and_method_call` and
    // `e2e_class_mixed_long_and_string_get` — see `decl::ClassMemberRegion`)
    // directly disproved the single-shared-temp model: each distinct
    // type-context gets its OWN separate, non-overlapping frame region, not
    // a fixed-size compromise. This test locks in that behavior for the
    // Long+Double combination specifically (analogous to the captured
    // Long+String case, same underlying mechanism) — no longer gated.
    let mut ea = ExprArena::new();
    let o = name_ref(&mut ea, 0);
    let get_long = member_access_node(&mut ea, o, 10);
    let get_double = member_access_node(&mut ea, o, 11);
    let x = name_ref(&mut ea, 1);
    let y = name_ref(&mut ea, 2);
    let stmt1 = assign_node(&mut ea, x, get_long);
    let stmt2 = assign_node(&mut ea, y, get_double);
    let body = block_node(&mut ea, vec![stmt1, stmt2]);

    let module = BoundModule {
        procs: vec![make_proc(vec![udt_var(0, 100), long_var(1), double_var(2)], vec![], body)],
        class_field_info: class_100(vec![
            property_slot("PL", VbaType::Long, AccessorKind::Get),
            property_slot("PD", VbaType::Double, AccessorKind::Get),
        ]),
        class_member_slots: HashMap::from([
            (get_long.0, resolved_member(Some(0x1c), None, true)),
            (get_double.0, resolved_member(Some(0x20), None, true)),
        ]),
        resolutions: resolutions_of(&[(o, local(0)), (x, local(1)), (y, local(2))]),
        types: types_of(&[(get_long, VbaType::Long), (get_double, VbaType::Double)]),
        ..BoundModule::default()
    };

    assert_eq!(
        lower_classes_ok(&module, &ea),
        &[
            // Get(Long): its own region's temp, read back with 0x6c.
            0x04, 0x68, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x1c, 0x00, 0x01, 0x00,
            0x6c, 0x68, 0xff, 0x71, 0x74, 0xff,
            // Get(Double): a SEPARATE region/temp (offset 0x60, not 0x68),
            // read back with 0x6f — the shared member-type-descriptor pool
            // index (`01 00`) is reused, matching every other vtable call.
            0x04, 0x60, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x20, 0x00, 0x01, 0x00,
            0x6f, 0x60, 0xff, 0x74, 0x6c, 0xff, 0x14,
        ]
    );
}

/// Two-field class: `Public F As Long` then `Public G As Long`. `G`'s Get
/// must land at `F`'s Get(0x1c)+8 = 0x24 (F consumes Get=0x1c,Let=0x20; the
/// second field starts right after) — the general multi-member slot rule.
#[test]
fn class_field_access_with_two_fields_computes_second_fields_offset_slot() {
    let mut ea = ExprArena::new();
    let o = name_ref(&mut ea, 0);
    let field_access = member_access_node(&mut ea, o, 10 /* sym for G */);
    let x = name_ref(&mut ea, 1);
    let stmt = assign_node(&mut ea, x, field_access);
    let body = block_node(&mut ea, vec![stmt]);

    let module = BoundModule {
        procs: vec![make_proc(vec![udt_var(0, 100), long_var(1)], vec![], body)],
        class_field_info: class_100(vec![field_slot("F"), field_slot("G")]),
        class_member_slots: HashMap::from([(
            field_access.0,
            resolved_member(Some(0x24), Some(0x28), false),
        )]),
        resolutions: resolutions_of(&[(o, local(0)), (x, local(1))]),
        types: types_of(&[(field_access, VbaType::Long)]),
        ..BoundModule::default()
    };

    // Get-only read of G: LdAddr(temp), LdAddr(o), resolve-object, vtable-call
    // at slot 0x24, load temp into x.
    assert_eq!(
        lower_classes_ok(&module, &ea),
        &[
            0x04, 0x70, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x24, 0x00, 0x01, 0x00,
            0x6c, 0x70, 0xff, 0x71, 0x74, 0xff, 0x14,
        ]
    );
}

// ── `Set` assignment to a plain object local (New/Nothing) ──────────────────
//
// Oracle-captured (`c7_set_new_reassign_nothing`; see the
// `e2e_set_new_reassign_nothing` fixture for the full byte-exact three-
// statement sequence). These unit tests isolate the GATED branches the
// fixture doesn't exercise: a `New` of a class other than the target's
// declared type, and a Set-source local that was never declared `As New`.

#[test]
fn set_new_class_matches_oracle_bytes() {
    let bytes =
        lower_set_into_object(vec![udt_var(0, 100)], &[100], SetSource::New(100)).unwrap();
    // `o` is the proc's only local: frame offset -136 (0xff78).
    assert_eq!(bytes, &[0xfd, 0xf4, 0x00, 0x00, 0x19, 0x78, 0xff, 0x14]);
}

#[test]
fn set_new_of_mismatched_class_is_gated() {
    // `o` is declared As class 100, not 200.
    let err = lower_set_into_object(vec![udt_var(0, 100)], &[100, 200], SetSource::New(200));
    assert!(matches!(err, Err(LowerError::UnsupportedType)));
}

#[test]
fn set_from_non_as_new_object_local_is_gated() {
    // `Dim other As Class1` (no `New`) read as a `Set` source has no oracle
    // capture — only the `As New` lazy-fetch shape (opcode 0x56) is grounded.
    let err = lower_set_into_object(
        vec![udt_var(0, 100), udt_var(1, 100)],
        &[100],
        SetSource::Local(1),
    );
    assert!(matches!(err, Err(LowerError::UnsupportedType)));
}

#[test]
fn set_nothing_matches_oracle_bytes() {
    let bytes =
        lower_set_into_object(vec![udt_var(0, 100)], &[100], SetSource::Nothing).unwrap();
    // `fc 63` (push 0), `3d 00 00` (coerce to class 100's type-desc, its own
    // first class-const-table entry -> index 0), `19 78 ff` (AddRef-store).
    assert_eq!(bytes, &[0xfc, 0x63, 0x3d, 0x00, 0x00, 0x19, 0x78, 0xff, 0x14]);
}

// ── Property Get value type (String) ─────────────────────────────────────────
//
// Oracle-captured (`c1_get_string`; see the `e2e_class_property_get_string`
// fixture): a Property Get returning `String` reads its out-param temp back
// with the STEAL opcode `0x3e` (push the temp's BSTR pointer, zero the temp
// — no separate release needed) rather than a plain typed load, and the
// target receives the MOVE store `0x31` (ctx 9) rather than the refcounted
// copy-store `0x43` a plain string-variable source would use.

#[test]
fn class_property_get_string_matches_oracle_bytes() {
    assert_eq!(
        lower_property_get(VbaType::String, string_var(1), false),
        &[
            0x04, 0x70, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x1c, 0x00, 0x01, 0x00,
            0x3e, 0x70, 0xff, 0x31, 0x74, 0xff, 0x14,
        ]
    );
}

// ── Property Get value type (Object) ─────────────────────────────────────────
//
// Oracle-captured (`c1_get_object`; see the `e2e_class_property_get_object`
// fixture): `Set x = o.P` where `P` returns `Object`. The Get-temp read-back
// uses `0x51` (a distinct plain 4-byte pointer read from the load-context
// table's own `0x6c` a variable read would use), and — because the client
// spelling is `Set`, not a plain `Assign` — the result is stored with the
// refcounted AddRef-store `0x19`, the same store `Set o = New`/`Set o =
// Nothing` use, not a typed variable store.

#[test]
fn class_property_get_object_matches_oracle_bytes() {
    assert_eq!(
        lower_property_get(VbaType::Object, object_var(1), true),
        &[
            0x04, 0x70, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d, 0x1c, 0x00, 0x01, 0x00,
            0x51, 0x70, 0xff, 0x19, 0x74, 0xff, 0x14,
        ]
    );
}

// ── Property Let value type (Double) ─────────────────────────────────────────
//
// Oracle-captured (`c2_let_double`; see the `e2e_class_property_let_double`
// fixture): a Property Let taking a `Double` stages its argument with the
// FPU-aware store `0xfd 0xc9` (pop FPU-top, store as Double with overflow
// check), not the plain `0x59` a `Long` Let uses — and the Integer literal
// `1` coerced to that Double parameter is pushed via the compact 1-byte form
// `f4 01` then converted with the Integer(word)->Double FPU-load `0xeb`
// (distinct from Long(dword)->Double's `0xec`), both via the ALREADY-existing
// general literal-coercion pipeline (`coerce_assign_value` + `emit_
// conversion`) — no new coercion code needed, only the staging opcode itself.

#[test]
fn class_property_let_double_matches_oracle_bytes() {
    assert_eq!(
        lower_property_let(VbaType::Double, AstLit::Int(1), VbaType::Integer),
        &[
            0xf4, 0x01, 0xeb, 0xfd, 0xc9, 0x70, 0xff, 0x04, 0x78, 0xff, 0x24, 0x00, 0x00, 0x0d,
            0x1c, 0x00, 0x01, 0x00, 0x14,
        ]
    );
}

// ── Property Let value type (String) ─────────────────────────────────────────
//
// Oracle-captured (`c2_let_string`; see the `e2e_class_property_let_string`
// fixture): a Property Let taking a `String` copy-stores the pushed literal
// into the shared class-member temp (`0x43`, the same refcounted store a
// plain `Dim s As String: s = "x"` would use — properly owning the BSTR),
// then passes the temp's ADDRESS (`0x04`) rather than a staged value, and
// releases the temp copy (`0x2f`) after the vtable call returns. This test
// ALSO exercises the const-pool-sharing fix: the string literal "x" claims
// pool index 0, so the class-create entry (normally index 0 in every other
// shipped slice's single-class, no-string proc) lands at index 1 instead —
// and the vtable call's own member-type-descriptor operand, previously
// hardcoded to a fixed `1`, must now correctly land at index 2.

#[test]
fn class_property_let_string_matches_oracle_bytes() {
    assert_eq!(
        lower_property_let(VbaType::String, AstLit::Str("x".into()), VbaType::String),
        &[
            0x1b, 0x00, 0x00, 0x43, 0x74, 0xff, 0x04, 0x74, 0xff, 0x04, 0x78, 0xff, 0x24, 0x01,
            0x00, 0x0d, 0x1c, 0x00, 0x02, 0x00, 0x2f, 0x74, 0xff, 0x14,
        ]
    );
}

// ── Project-level resolutions are not lowered ─────────────────────────────────

/// `r = a` and `a = 1` where `a` resolves to `res`, with `r` an ordinary Long.
fn lower_with_project_level(res: NameResolution) -> (Result<Vec<u8>, LowerError>, Result<Vec<u8>, LowerError>) {
    let mut results = Vec::new();
    for read in [true, false] {
        let mut ea = ExprArena::new();
        let a = name_ref(&mut ea, 0);
        let r = name_ref(&mut ea, 1);
        let one = int_lit_node(&mut ea, 1);
        let stmt = if read { assign_node(&mut ea, r, a) } else { assign_node(&mut ea, a, one) };
        let body = block_node(&mut ea, vec![stmt]);
        let mut resolutions = HashMap::new();
        resolutions.insert(a.0, res.clone());
        resolutions.insert(r.0, NameResolution::Local { proc_idx: 0, local_idx: 0 });
        let module = BoundModule {
            procs: vec![make_proc(vec![long_var(1)], vec![], body)],
            resolutions,
            ..BoundModule::default()
        };
        results.push(lower_proc(&module, 0, &ea, 0x0008));
    }
    let write = results.pop().unwrap();
    let read = results.pop().unwrap();
    (read, write)
}

#[test]
fn project_level_resolutions_are_unsupported_nodes() {
    use vb6_sema::sema::ExternalDecl;
    let kinds = [
        NameResolution::Module(0),
        NameResolution::Project,
        NameResolution::Library(0),
        NameResolution::LibraryType { lib: 0, ty: 0 },
        NameResolution::LibraryMember { lib: 0, ty: 0, member: 0 },
        NameResolution::External { module: 0, decl: ExternalDecl::Control(0) },
        NameResolution::External { module: 0, decl: ExternalDecl::TypeMember { type_idx: 0, member_idx: 0 } },
    ];
    for res in kinds {
        let (read, write) = lower_with_project_level(res.clone());
        assert_eq!(read, Err(LowerError::UnsupportedNode), "read of {res:?}");
        assert_eq!(write, Err(LowerError::UnsupportedNode), "write of {res:?}");
    }
}

#[test]
fn unresolved_and_cross_module_variable_stay_unresolved() {
    use vb6_sema::sema::ExternalDecl;
    let (read, write) = lower_with_project_level(NameResolution::Unresolved);
    assert_eq!(read, Err(LowerError::Unresolved));
    assert_eq!(write, Err(LowerError::Unresolved));
    let (read, write) =
        lower_with_project_level(NameResolution::External { module: 1, decl: ExternalDecl::Var(0) });
    assert_eq!(read, Err(LowerError::Unresolved));
    assert_eq!(write, Err(LowerError::Unresolved));
}
