//! AST-to-P-code lowering: walks a `BoundModule`/`BoundProc` from vb6-sema and
//! emits the runtime P-code byte stream via the reference emitter.
//!
//! The entry point is [`lower_proc`].  It builds local/param/global frame
//! layouts (matching VB6's exact slot ordering), walks the body
//! [`ExprNode::Block`] recursively, and for each statement that can be lowered
//! builds a [`NodeArena`] sub-tree and calls [`Emitter::emit_expr`] followed
//! by the appropriate typed store.
//!
//! Only the code paths whose P-code bytes have been oracle-confirmed are
//! implemented.  Unhandled constructs return [`LowerError::UnsupportedNode`]
//! or [`LowerError::UnsupportedType`] — never a silently wrong byte.

use std::cell::{Cell, RefCell};
use std::collections::HashMap;

use vb6_sema::sema::{
    BoundModule, BoundProc, BuiltinCall, ExternalClass, RtcArg, UnaryIntrinsic, VbaType,
    NameResolution,
};
use vb6_syntax::frontend::ast::{ExprArena, ExprNode, AstLit, BinOpKind, UnOpKind, DoKind, ExitKind, LabelRef, OnErrorKind, ResumeTarget};
use vb6_syntax::support::arena::NodeId;

use crate::bind::{GlobalFrame, GlobalVar, LocalVar, ParamVar, ProcFrame, UdtLocal};
use crate::bridge::{load_store_ctx, param_frame_from_types, type_ctx, UnsupportedType};
use crate::emit::Emitter;
use crate::node::{NodeArena, NodeRef};
use crate::tables::RT_STORE_BY_CTX;

mod argcoerce;
mod assign;
mod decl;
mod expr;
mod intrinsics;
mod stmt;

use decl::*;
use stmt::*;

/// Errors that can arise while lowering a proc.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LowerError {
    /// A variable or literal has a type whose P-code emission is not yet
    /// oracle-confirmed (e.g. String, Variant, user-defined type).
    UnsupportedType,
    /// A name reference has no entry in `BoundModule.resolutions`.
    Unresolved,
    /// An AST node kind the lowering pass cannot yet handle.
    UnsupportedNode,
    /// `proc_idx` exceeds `module.procs.len()`.
    ProcIndexOutOfRange,
}

impl From<UnsupportedType> for LowerError {
    fn from(_: UnsupportedType) -> Self {
        LowerError::UnsupportedType
    }
}

/// Map a `VbaType` to the VB6 internal type tag stored in the high 16 bits of
/// an expression node's `word[0]`.
///
/// Node type tags (the high-16 of word[0]), grounded from VB6's kind->VARTYPE
/// table (`DAT_0fa92778[kind] = VARTYPE`): Integer=6, Long=8, Single=0xa,
/// Double=0xb, Date=0xc, Currency=0xd, Variant=0xf, String=0x10, Object=0x16.
/// A Boolean *value* is operated on as Integer (tag 6) — VB6 selects opcodes
/// by the Integer class for Boolean — so it shares tag 6 here (its
/// declaration kind 3 is a separate namespace). Object's tag (22/0x16) is
/// cross-confirmed two ways: a live TTD trace of `Set o.PG = y` shows
/// VBA6.DLL's `EbEmitExpression2` called with `nType=0x16` for the `y`
/// reference, and the already-extracted `RT_RESULT_TYPE[0x16] = 4` matches
/// Object's 4-byte pointer size independently.
fn vba_type_to_node_tag(ty: &VbaType) -> Option<u16> {
    match ty {
        VbaType::Integer | VbaType::Boolean => Some(6),
        VbaType::Long => Some(8),
        VbaType::Single => Some(10),
        VbaType::Double => Some(11),
        VbaType::Currency => Some(0xd),
        VbaType::Date => Some(0xc),
        VbaType::Byte => Some(5),
        VbaType::String => Some(0x10),
        VbaType::Variant => Some(0xf),
        VbaType::Object => Some(0x16),
        _ => None,
    }
}

/// Map a `BinOpKind` to the oracle-confirmed bound-node opcode.
/// Returns `None` for operators whose P-code bytes are not yet confirmed.
fn binop_node_opcode(op: BinOpKind) -> Option<u16> {
    Some(match op {
        BinOpKind::Add => 0x16,
        BinOpKind::Sub => 0x17,
        BinOpKind::Mul => 0x18,
        // Div (`/`) is the arithmetic binop occupying 0x19 — the gap in the
        // contiguous +/-/*//^ block (0x16..0x1a): RT_BINOP_BASE[0x19]=0xb6 is a
        // valid base and the stmt jump table routes 0x19 to the generic-binop
        // emitter (stmt_case_0fab1da9), like Add/Sub.
        BinOpKind::Div => 0x19,
        // The multiplicative and logical operators the front-end operator table
        // assigns by precedence (consumed by the generic operation emitter):
        //   `\`  (integer divide) -> 0x1e   Mod -> 0x1d
        //   Eqv -> 0x20                      Imp -> 0x1f
        // (precedence ladder * / > \ > Mod > + - and And > Or > Xor > Eqv > Imp).
        BinOpKind::IDiv => 0x1e,
        BinOpKind::Mod => 0x1d,
        BinOpKind::Or  => 0x21,
        BinOpKind::Xor => 0x22,
        BinOpKind::And => 0x23,
        BinOpKind::Eqv => 0x20,
        BinOpKind::Imp => 0x1f,
        // String concatenation (`&`): node 0x24; the String-tagged emitter case
        // emits the concat opcode (0x2a). Result is a fresh string temp.
        BinOpKind::Cat => 0x24,
        // String `Like` pattern match: bound opcode 0x25 (comparison-dispatch);
        // for a String LHS, base 0x77 + offset 7 -> 0x7e -> fb 7e.
        BinOpKind::Like => 0x25,
        BinOpKind::Eq  => 0x26,
        BinOpKind::Ne  => 0x27,
        BinOpKind::Le  => 0x28,
        BinOpKind::Ge  => 0x29,
        BinOpKind::Lt  => 0x2a,
        BinOpKind::Gt  => 0x2b,
        _ => return None,
    })
}

fn is_comparison_op(op: BinOpKind) -> bool {
    matches!(
        op,
        BinOpKind::Eq | BinOpKind::Ne | BinOpKind::Lt | BinOpKind::Le | BinOpKind::Gt | BinOpKind::Ge
            | BinOpKind::Like
    )
}

/// Allocate global slots from module-level variable types in declaration order.
/// Mirrors [`frame_from_local_types`] for the global data block.
pub fn global_frame_from_types(
    types: &[VbaType],
    module_desc: u16,
) -> Result<Vec<GlobalVar>, UnsupportedType> {
    let mut frame = GlobalFrame::new(module_desc);
    let mut out = Vec::with_capacity(types.len());
    for ty in types {
        let ctx = type_ctx(ty).ok_or(UnsupportedType)?;
        out.push(frame.declare_anon_global(ctx));
    }
    Ok(out)
}

/// Count the number of For loops directly or indirectly in an AST subtree.
/// Each For loop needs 2 hidden Long slots in the frame.
fn count_for_loops(node_id: NodeId, expr_arena: &ExprArena) -> usize {
    match expr_arena.get(node_id) {
        ExprNode::For { body, .. } => 1 + count_for_loops(*body, expr_arena),
        ExprNode::Block { stmts } => {
            stmts.iter().map(|&id| count_for_loops(id, expr_arena)).sum()
        }
        ExprNode::If { then_body, else_body, .. } => {
            count_for_loops(*then_body, expr_arena)
                + else_body.map(|id| count_for_loops(id, expr_arena)).unwrap_or(0)
        }
        ExprNode::While { body, .. } => count_for_loops(*body, expr_arena),
        ExprNode::Do { body, .. } => count_for_loops(*body, expr_arena),
        ExprNode::SelectCase { cases, .. } => {
            cases.iter().map(|&id| count_for_loops(id, expr_arena)).sum()
        }
        ExprNode::CaseBlock { body, .. } => count_for_loops(*body, expr_arena),
        ExprNode::CaseElse { body } => count_for_loops(*body, expr_arena),
        _ => 0,
    }
}

/// Collect, in statement order, the subject expression of every `Select Case` in
/// a subtree. Each needs one hidden frame slot (typed as the subject) to hold the
/// evaluated subject across the per-case comparisons.
fn collect_select_subjects(node_id: NodeId, expr_arena: &ExprArena, out: &mut Vec<NodeId>) {
    match expr_arena.get(node_id) {
        ExprNode::SelectCase { subject, cases, .. } => {
            out.push(*subject);
            for &id in cases {
                collect_select_subjects(id, expr_arena, out);
            }
        }
        ExprNode::CaseBlock { body, .. } => collect_select_subjects(*body, expr_arena, out),
        ExprNode::CaseElse { body } => collect_select_subjects(*body, expr_arena, out),
        ExprNode::Block { stmts } => {
            for &id in stmts {
                collect_select_subjects(id, expr_arena, out);
            }
        }
        ExprNode::If { then_body, else_body, .. } => {
            collect_select_subjects(*then_body, expr_arena, out);
            if let Some(id) = else_body {
                collect_select_subjects(*id, expr_arena, out);
            }
        }
        ExprNode::While { body, .. } => collect_select_subjects(*body, expr_arena, out),
        ExprNode::Do { body, .. } => collect_select_subjects(*body, expr_arena, out),
        ExprNode::For { body, .. } => collect_select_subjects(*body, expr_arena, out),
        _ => {}
    }
}

/// Lower a single `BoundProc` to its P-code byte vector.
///
/// Frame layout follows VB6's exact convention: locals at negative offsets
/// from -136 downward (4 bytes per Integer/Long/Single/Object, 8 bytes per
/// Double/Currency), params at positive offsets from +12 upward.
///
/// For loops each need 2 hidden Long slots allocated below all user locals.
/// These are pre-allocated here by scanning the body first.
///
/// `module_desc` is the compiled module-object descriptor word — `0x0008` for
/// the primary module in a single-module project (oracle-confirmed).
/// Lower every procedure of a module in declaration order, sharing one
/// module-global string pool so string-literal indices are assigned across the
/// whole module (proc 0's strings first, then proc 1's, deduped by value).
/// Returns the per-procedure p-code byte streams.
pub fn lower_module(
    module: &BoundModule,
    expr_arena: &ExprArena,
    module_desc: u16,
) -> Result<Vec<Vec<u8>>, LowerError> {
    let empty = HashMap::new();
    lower_module_with_classes(module, expr_arena, module_desc, &empty)
}

/// Like [`lower_module`], but with the known-external-classes table (`Dim o As
/// New ClassName` / `o.Field`) needed to lower a class instance's frame slot
/// and member access — see [`bind_with_classes`](vb6_sema::sema::bind_with_classes)
/// (the sema-side counterpart that types `o.Field`; this is codegen's, keyed
/// the same way, needed for the field's vtable dispatch slot / frame layout).
pub fn lower_module_with_classes(
    module: &BoundModule,
    expr_arena: &ExprArena,
    module_desc: u16,
    known_classes: &HashMap<String, ExternalClass>,
) -> Result<Vec<Vec<u8>>, LowerError> {
    let mut pool: Vec<ModuleConstEntry> = Vec::new();
    let mut static_base: u16 = 0;
    let mut procs = Vec::with_capacity(module.procs.len());
    for idx in 0..module.procs.len() {
        let (bytes, next_pool, next_static) = lower_proc_pooled(
            module, idx, expr_arena, module_desc, pool, static_base, known_classes,
        )?;
        pool = next_pool;
        static_base = next_static;
        procs.push(bytes);
    }
    Ok(procs)
}

pub fn lower_proc(
    module: &BoundModule,
    proc_idx: usize,
    expr_arena: &ExprArena,
    module_desc: u16,
) -> Result<Vec<u8>, LowerError> {
    let empty = HashMap::new();
    lower_proc_with_classes(module, proc_idx, expr_arena, module_desc, &empty)
}

/// Like [`lower_proc`], but with the known-external-classes table — see
/// [`lower_module_with_classes`].
pub fn lower_proc_with_classes(
    module: &BoundModule,
    proc_idx: usize,
    expr_arena: &ExprArena,
    module_desc: u16,
    known_classes: &HashMap<String, ExternalClass>,
) -> Result<Vec<u8>, LowerError> {
    // A standalone single-procedure lowering starts with an empty string pool.
    let (bytes, _pool, _static) = lower_proc_pooled(
        module, proc_idx, expr_arena, module_desc, Vec::new(), 0, known_classes,
    )?;
    Ok(bytes)
}

/// Lower one procedure, threading the module-global constant/global-pointer
/// pool (string literals AND class-constant entries, one shared index space
/// — see `ModuleConstEntry`) in and out so its entries continue across
/// procedures. `pool_in` carries the entries interned by earlier procedures;
/// the returned pool adds this one's.
fn lower_proc_pooled(
    module: &BoundModule,
    proc_idx: usize,
    expr_arena: &ExprArena,
    module_desc: u16,
    pool_in: Vec<ModuleConstEntry>,
    static_base: u16,
    known_classes: &HashMap<String, ExternalClass>,
) -> Result<(Vec<u8>, Vec<ModuleConstEntry>, u16), LowerError> {
    let proc = module.procs.get(proc_idx).ok_or(LowerError::ProcIndexOutOfRange)?;

    let user_local_count = proc.locals.len();

    // Build the local frame directly so `Const` locals can be skipped: a const
    // occupies NO frame space (VB6 folds it to a literal at each use site), but
    // keeps an index-aligned placeholder slot so `NameResolution::Local`'s
    // `local_idx` still maps directly. For-loop hidden slots and Select-subject
    // temps are declared on the same frame, after the user locals.
    let mut frame = ProcFrame::new();
    let mut local_slots: Vec<LocalVar> = Vec::with_capacity(proc.locals.len());
    // A `Type...End Type`-typed local's UDT binding, parallel to `local_slots`
    // (which carries an index-aligned placeholder for these, like Const/Static).
    let mut local_udts: Vec<Option<UdtLocal>> = Vec::with_capacity(proc.locals.len());
    // A class-instance local's (`Dim o As New ClassName`) field list, parallel
    // to `local_slots` (which carries the real 4-byte object-reference
    // binding for these, not a placeholder — see the loop below).
    let mut local_classes: Vec<Option<ExternalClass>> = Vec::with_capacity(proc.locals.len());
    // Static locals live in a per-procedure static block (not the frame): each is
    // assigned a byte offset within that block, packed by type size in declaration
    // order. `static_offsets[i]` is meaningful only when `locals[i].is_static`.
    let mut static_offsets: Vec<u16> = Vec::with_capacity(proc.locals.len());
    let mut static_cursor: u16 = static_base;
    for v in &proc.locals {
        if v.is_static {
            static_offsets.push(static_cursor);
            static_cursor += static_var_size(&v.vba_type);
            // No frame slot: a placeholder keeps local_idx aligned.
            local_slots.push(LocalVar { type_ctx: 0, frame_offset: 0 });
            local_udts.push(None);
            local_classes.push(None);
            continue;
        }
        static_offsets.push(0);
        if v.is_const {
            local_slots.push(LocalVar { type_ctx: 0, frame_offset: 0 });
            local_udts.push(None);
            local_classes.push(None);
        } else if let Some(n) = v.fixed_string_len {
            // A fixed-length string (`As String * n`) holds an inline Unicode
            // buffer of `n` chars = 2*n bytes (oracle-confirmed for n=1,4,8,10,16,20).
            let size = 2 * (n as i16);
            local_slots.push(frame.declare_anon_bytes(size));
            local_udts.push(None);
            local_classes.push(None);
        } else if matches!(v.vba_type, VbaType::Array(_)) {
            match v.array_dims {
                // A fixed array is a SAFEARRAY descriptor (size-independent of the
                // element count — data is heap-allocated): 20 bytes + 8 per
                // dimension (28 for 1-D, 36 for 2-D); the LdAddr target sits 4
                // bytes above the slot bottom.
                Some(dims) => {
                    let mut slot = frame.declare_anon_bytes(20 + 8 * dims as i16);
                    slot.frame_offset += 4;
                    local_slots.push(slot);
                }
                // A dynamic array (`Dim a()`) is a 4-byte pointer slot; the array
                // is allocated by `ReDim`.
                None => local_slots.push(frame.declare_anon_bytes(4)),
            }
            local_udts.push(None);
            local_classes.push(None);
        } else if let Some(class) = match &v.vba_type {
            VbaType::UserDefined(sym) => module.class_field_info.get(sym),
            _ => None,
        } {
            // A class-instance local (`Dim o As New ClassName`): a plain
            // 4-byte object reference (ctx 0 — the same "Object/untyped
            // pointer" frame class already mapped in bridge.rs::type_ctx),
            // NOT an embedded struct like a UDT — the field's storage lives
            // in the (separately allocated) object instance, reached via the
            // vtable-dispatch mechanism (emit/reference.rs's class-member
            // path), not a frame offset. `local_slots` keeps an
            // index-aligned placeholder; the real binding lives in
            // `local_classes`.
            local_slots.push(frame.declare_anon(0));
            local_udts.push(None);
            local_classes.push(Some(class.clone()));
        } else if let VbaType::UserDefined(type_sym) = &v.vba_type {
            // A `Type...End Type`-typed local: lay out one frame slot sized for
            // every declared field (uniform-size fields only — see
            // `ProcFrame::declare_udt_local`). `local_slots` keeps an
            // index-aligned placeholder; the real binding lives in `local_udts`.
            let decl = module
                .type_decls
                .iter()
                .find(|d| d.sym_id == *type_sym)
                .ok_or(LowerError::Unresolved)?;
            let field_ctxs: Vec<usize> = decl
                .members
                .iter()
                .map(|m| type_ctx(&m.vba_type).ok_or(LowerError::UnsupportedType))
                .collect::<Result<_, _>>()?;
            let udt = frame.declare_anon_udt(&field_ctxs);
            local_slots.push(LocalVar { type_ctx: 0, frame_offset: 0 });
            local_udts.push(Some(udt));
            local_classes.push(None);
        } else {
            let tctx = type_ctx(&v.vba_type).ok_or(LowerError::UnsupportedType)?;
            local_slots.push(frame.declare_anon(tctx));
            local_udts.push(None);
            local_classes.push(None);
        }
    }

    // 2 Long hidden slots per For loop.
    let for_count = count_for_loops(NodeId(proc.body), expr_arena);
    for _ in 0..(for_count * 2) {
        local_slots.push(frame.declare_anon(2));
    }

    // One hidden slot per Select Case, typed as its subject.
    let select_base = local_slots.len();
    let mut select_subjects = Vec::new();
    collect_select_subjects(NodeId(proc.body), expr_arena, &mut select_subjects);
    for &subj in &select_subjects {
        let ty = module.types.get(&subj.0).cloned().unwrap_or(VbaType::Long);
        let tctx = type_ctx(&ty).ok_or(LowerError::UnsupportedType)?;
        local_slots.push(frame.declare_anon(tctx));
    }

    // One hidden 16-byte Variant temp per Variant-target assignment.
    let variant_base = local_slots.len();
    let variant_temps = count_variant_assigns(module, NodeId(proc.body), expr_arena);
    for _ in 0..variant_temps {
        local_slots.push(frame.declare_anon(10));
    }

    // One hidden 4-byte string temp per intermediate result in a concat chain.
    let concat_base = local_slots.len();
    let concat_temps = count_concat_temps(module, proc, NodeId(proc.body), expr_arena);
    for _ in 0..concat_temps {
        local_slots.push(frame.declare_anon(5));
    }

    // One hidden 16-byte string-result temp per String-returning runtime intrinsic
    // call (Chr/Space), used to receive the runtime function's result.
    let string_rtc_base = local_slots.len();
    let string_rtc_temps = count_string_rtc_temps(module, NodeId(proc.body), expr_arena);
    for _ in 0..string_rtc_temps {
        local_slots.push(frame.declare_anon(10));
    }

    // One hidden 4-byte BSTR temp per runtime-string call argument (the owned result
    // copied for a ByVal String parameter), placed after the 16-byte string temps.
    let owned_copy_base = local_slots.len();
    let owned_copy_temps = count_owned_copy_temps(module, NodeId(proc.body), expr_arena);
    for _ in 0..owned_copy_temps {
        local_slots.push(frame.declare_anon(5));
    }

    // Hidden scratch temps for the proc's class-member vtable dispatch use:
    // a Get access (`x = o.F`/`x = o.P`) writes its out-parameter through
    // one, a Property-Let call (`o.P = v`) stages its argument through one
    // (`0x59 <offset>`) before the vtable call, a Property-Set call stages
    // its argument through one (`fd 9c <offset>`), and a class-method call
    // stages each of its own arguments (plus a Function's result, in value
    // position) through one apiece — a plain-field store needs no staging.
    // One REGION per distinct frame type-context (`ClassMemberRegion`), not
    // one shared region for the whole proc: oracle-confirmed this session
    // that a `Double` Get-temp region and an unrelated `Sub`-call-argument
    // region (`Long`) are SEPARATE, non-overlapping frame areas, and
    // likewise two DIFFERENT Get types (`Long`, `String`) never share one
    // area either, even though same-typed repeated accesses within ONE
    // region still correctly reuse its slots (`e2e_class_multi_field_and_
    // property`/`e2e_class_property_let_before_get`) — see
    // `decl::class_member_regions`'s doc comment for the full derivation.
    let mut class_member_bases: std::collections::HashMap<usize, usize> = std::collections::HashMap::new();
    for region in class_member_regions(module, NodeId(proc.body), expr_arena)? {
        class_member_bases.insert(region.type_ctx, local_slots.len());
        for _ in 0..region.slots {
            local_slots.push(frame.declare_anon(region.type_ctx));
        }
    }

    // `ParamArray` is the variadic inter-procedure-call argument mechanism (the
    // callee receives a packed array of the caller's trailing arguments); it
    // belongs to the procedure-call tier, which is out of scope. Gate it cleanly
    // rather than emit a wrong frame.
    if proc.params.iter().any(|p| p.flags.param_array) {
        return Err(LowerError::UnsupportedNode);
    }
    let param_types: Vec<VbaType> = proc.params.iter().map(|p| p.vba_type.clone()).collect();
    let param_byref: Vec<bool> = proc.params.iter().map(|p| !p.flags.by_val).collect();
    let global_types: Vec<VbaType> =
        module.module_vars.iter().map(|v| v.vba_type.clone()).collect();

    let param_slots = param_frame_from_types(&param_types, &param_byref)?;
    let global_slots = global_frame_from_types(&global_types, module_desc)?;

    let ctx = LowerCtx {
        module,
        proc,
        local_slots,
        local_udts,
        local_classes,
        known_classes,
        param_slots,
        global_slots,
        user_local_count,
        for_next_pair: Cell::new(0),
        select_base,
        select_next: Cell::new(0),
        variant_base,
        variant_next: Cell::new(0),
        concat_base,
        concat_next: Cell::new(0),
        string_rtc_base,
        string_rtc_next: Cell::new(0),
        owned_copy_base,
        owned_copy_next: Cell::new(0),
        class_member_bases,
        call_next: Cell::new(0),
        labels: RefCell::new(Vec::new()),
        goto_patches: RefCell::new(Vec::new()),
        exit_stack: RefCell::new(Vec::new()),
        const_pool: RefCell::new(pool_in),
        module_desc,
        static_offsets,
        line_tracking: proc_needs_line_tracking(NodeId(proc.body), expr_arena),
        line_markers: RefCell::new(Vec::new()),
        member_symbol: RefCell::new(None),
    };

    let mut out = Vec::new();
    // When the procedure needs line tracking, the table opens with a header marker
    // (0x00 + delta to the first statement marker).
    if ctx.line_tracking {
        ctx.line_markers.borrow_mut().push(out.len());
        out.extend_from_slice(&[0x00, 0x00]);
    }
    lower_block(&ctx, NodeId(proc.body), expr_arena, &mut out)?;
    // …and closes with a trailer marker (delta 0), then the whole marker set is
    // backpatched into a forward-delta chain: each marker's 2-byte operand is the
    // byte distance to the next marker.
    if ctx.line_tracking {
        ctx.line_markers.borrow_mut().push(out.len());
        out.extend_from_slice(&[0x00, 0x00]);
        let markers = ctx.line_markers.borrow();
        for w in markers.windows(2) {
            // Marker = 0x00 opcode + 1-byte delta to the next marker.
            let delta = w[1] - w[0];
            if delta > 0xff {
                return Err(LowerError::UnsupportedNode);
            }
            out[w[0] + 1] = delta as u8;
        }
    }

    // Resolve forward/backward `GoTo` jumps now that every label's byte offset
    // is known.
    let labels = ctx.labels.borrow();
    for (target, patch) in ctx.goto_patches.borrow().iter() {
        let off = labels
            .iter()
            .find(|(l, _)| l == target)
            .map(|(_, o)| *o)
            .ok_or(LowerError::Unresolved)?;
        out[*patch..*patch + 2].copy_from_slice(&off.to_le_bytes());
    }
    drop(labels);
    // Every procedure's p-code ends with exactly one implicit-return opcode
    // `0x14`, unconditionally — oracle-confirmed (six independent real VB6
    // recaptures spanning a scalar assign, a class field access, a
    // Property Get/Let, a class-method call, a `GoTo`/line-tracking body, and
    // a `Function` return) to be appended by the compiler regardless of what
    // the body's last statement was, with NO deduplication when that last
    // statement is itself an explicit `Exit Sub`/`Exit Function` (which
    // already emits its own `0x14` via `ExitStmt` in `lower/stmt.rs`): a bare
    // `Sub Main() : Exit Sub : End Sub` compiles to `14 14`, two bytes, not
    // one. So this append is unconditional, never a "last byte already 0x14"
    // check.
    out.push(0x14);
    let pool_out = ctx.const_pool.into_inner();
    Ok((out, pool_out, static_cursor))
}

// ── Internal lowering context ─────────────────────────────────────────────────

struct LowerCtx<'m> {
    module: &'m BoundModule,
    proc: &'m BoundProc,
    local_slots: Vec<LocalVar>,
    /// A `Type...End Type`-typed local's UDT binding, parallel to
    /// `local_slots` (`None` for every non-UDT local).
    local_udts: Vec<Option<UdtLocal>>,
    /// A class-instance local's field list, parallel to `local_slots` (`None`
    /// for every non-class local).
    local_classes: Vec<Option<ExternalClass>>,
    /// The known-external-classes table this proc was lowered with (see
    /// [`lower_module_with_classes`]) — not indexed by local, kept for
    /// completeness/future multi-class lookups.
    known_classes: &'m HashMap<String, ExternalClass>,
    param_slots: Vec<ParamVar>,
    global_slots: Vec<GlobalVar>,
    /// Number of user-declared locals (hidden For-loop slots come after).
    user_local_count: usize,
    /// Which hidden-slot pair the next For loop should use.
    for_next_pair: Cell<usize>,
    /// Frame index of the first Select-subject temp slot.
    select_base: usize,
    /// Which Select-subject temp slot the next Select Case should use.
    select_next: Cell<usize>,
    /// Frame index of the first Variant-assignment temp slot.
    variant_base: usize,
    /// Which Variant temp slot the next Variant assignment should use.
    variant_next: Cell<usize>,
    /// Frame index of the first concat-chain string temp slot.
    concat_base: usize,
    /// Which concat temp slot the next concat-chain intermediate should use.
    concat_next: Cell<usize>,
    /// Frame index of the first String-returning-runtime-intrinsic result temp.
    string_rtc_base: usize,
    /// Which string-result temp slot the next such intrinsic should use.
    string_rtc_next: Cell<usize>,
    /// Frame index of the first 4-byte owned-string call-argument copy temp.
    owned_copy_base: usize,
    /// Which owned-copy temp slot the next runtime-string call argument should use.
    owned_copy_next: Cell<usize>,
    /// Frame index of each class-member vtable-dispatch scratch REGION's
    /// first slot, keyed by frame type-context (`crate::bridge::type_ctx`'s
    /// index space) — see `decl::ClassMemberRegion` and the allocation
    /// comment in `lower_proc`. One entry per distinct context the proc
    /// actually needs; absent entirely when the proc has no class-member
    /// access at all.
    class_member_bases: std::collections::HashMap<usize, usize>,
    /// Sequential index of the next call site within this procedure (each call's
    /// 2-byte callee-reference operand is its emission-order index, 0,1,2,…).
    call_next: Cell<usize>,
    /// Label definitions: `(label, byte offset)`, filled as labels are emitted.
    labels: RefCell<Vec<(LabelRef, u16)>>,
    /// Pending `GoTo` jumps: `(target label, patch offset)`, patched at proc end.
    goto_patches: RefCell<Vec<(LabelRef, usize)>>,
    /// Stack of `Exit For`/`Exit Do` patch lists — one per active loop; each entry
    /// is a byte offset to backpatch with the loop-end offset.
    exit_stack: RefCell<Vec<Vec<usize>>>,
    /// Module-wide constant/global-pointer table: string literals AND
    /// class-constant entries (the operand the `New`/lazy-`As New`-fetch/
    /// typed-`Nothing`/object-resolve opcodes consume) share ONE sequential
    /// index space, in first-use emission order — oracle-confirmed
    /// (`c2_let_string`: a string literal pushed BEFORE a class resolve in
    /// the same statement gets index 0, and the class resolve gets index 1,
    /// not a fresh 0 of its own; `e2e_module_global_string_pool__2` already
    /// established separately that this table is MODULE-wide, not
    /// per-procedure — a second procedure's first string literal continues
    /// the index sequence rather than restarting). Deduped independently
    /// within each entry kind (two identical string literals share a slot;
    /// two class-const entries share a slot only when `(kind, class sym)`
    /// match — see `intern_string`/`intern_class_const`).
    const_pool: RefCell<Vec<ModuleConstEntry>>,
    /// Compiled module-object descriptor word (used to address module globals and
    /// per-procedure static storage).
    module_desc: u16,
    /// Byte offset of each local within the procedure's static block; meaningful
    /// only for entries whose `BoundVar.is_static` is set.
    static_offsets: Vec<u16>,
    /// True when the procedure needs the statement line-number table (it has a
    /// numeric line label, a `Resume`, or `On Error Resume Next`). When set, a
    /// `0x00 <delta>` marker is threaded before each code-emitting statement.
    line_tracking: bool,
    /// Byte offsets of every emitted line-table marker (header, per-statement, and
    /// trailer), in emission (byte) order — backpatched into a forward-delta chain.
    line_markers: RefCell<Vec<usize>>,
    /// Scratch side-channel: the `SymbolContext` a just-lowered UDT field
    /// reference (`lower_udt_field_access`) needs its `Emitter` to carry.
    /// Building the node and resolving its context happen in the same pass,
    /// but the `Emitter` is constructed by the statement-level caller after
    /// the whole tree is built — the caller takes this (via
    /// `member_symbol.borrow_mut().take()`) and attaches it before emitting,
    /// same pattern as this context's other `Cell`/`RefCell` scratch state.
    member_symbol: RefCell<Option<crate::emit::SymbolContext>>,
}

/// A class-constant-table entry kind — see `ModuleConstEntry::Class`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ClassConstKind {
    /// The `New ClassName` / lazy-`As New`-fetch create-descriptor entry
    /// (opcodes `fd f4` and `56`) — ALSO the entry an object-resolve (`0x24`)
    /// consumes for its own lazy-New fallback, oracle-confirmed sharing this
    /// same slot (`c2_let_string`).
    Create,
    /// The typed-`Nothing` coercion entry (opcode `fc 63`).
    TypeDesc,
}

/// One entry in the module-wide constant/global-pointer table — see
/// `LowerCtx::const_pool`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum ModuleConstEntry {
    Str(String),
    Class(ClassConstKind, u32),
    /// A vtable-call's own second operand (the 2-byte field right after the
    /// Get/Let/Set/method slot) — read only on the runtime's error path (a
    /// type-mismatch message), per the `0x0d` handler's disassembly
    /// (`6610a43b`: the operand is unused on the success path, only pushed
    /// as a const-table lookup when the vtable call itself faults) — but its
    /// INDEX still consumes a real pool slot unconditionally, so it must be
    /// interned like every other entry, not hardcoded.
    ///
    /// NOT keyed by the callee member's TYPE — an earlier pass of this port
    /// assumed it was (deduped by `vba_type_to_node_tag`), which happened to
    /// match every fixture shipped at the time because none of them mixed
    /// TWO DIFFERENT member types in one proc against the SAME class. A
    /// dedicated fresh capture (`Get`-ing a `Double` property THEN calling a
    /// `Sub` with no return value, against ONE class instance, in one proc)
    /// disproved the type-keyed hypothesis directly: both vtable calls land
    /// on the SAME operand index despite one having a `Double` return and
    /// the other none at all.
    ///
    /// IS KEYED BY THE CALLEE'S CLASS, though — a later pass (also assumed
    /// module-global-shared, since every fixture up to that point only ever
    /// involved ONE class) was disproved by a fresh two-class capture: `Sub
    /// Main` creating a `Class1` instance and a `Class2` instance, each
    /// calling its own `DoIt(x As Long)`, shows the FIRST call's operand at
    /// index 1 and the SECOND call's operand at index 3 (not both at index
    /// 1) — genuinely different indices for genuinely different classes.
    /// Every prior "shared index" finding (the Double/Sub-call pair, the
    /// Long-Get/String-Get pair, `c2_let_string`'s index-2 landing,
    /// `e2e_class_multi_field_and_property`'s six same-typed accesses
    /// landing on index 1) is STILL consistent with this corrected rule —
    /// every one of those captures only ever involved ONE class instance,
    /// so they never had a second, DIFFERENT class's vtable call in the same
    /// proc to reveal the class-dependence. Deduped by `(kind, class_sym)`
    /// like `Class`, not a bare unit value.
    MemberType(u32),
}

impl LowerCtx<'_> {
    /// Intern a string literal, returning its pool index (deduped by value
    /// among `Str` entries; shares the index space with `Class` entries —
    /// see `const_pool`'s field doc).
    fn intern_string(&self, s: &str) -> u16 {
        let mut pool = self.const_pool.borrow_mut();
        if let Some(i) = pool.iter().position(|p| matches!(p, ModuleConstEntry::Str(v) if v == s))
        {
            return i as u16;
        }
        pool.push(ModuleConstEntry::Str(s.to_string()));
        (pool.len() - 1) as u16
    }

    /// Intern a class-constant-table entry, returning its 16-bit operand
    /// index (deduped by `(kind, class_sym)` among `Class` entries; shares
    /// the index space with `Str` entries — see `const_pool`'s field doc).
    pub(super) fn intern_class_const(&self, kind: ClassConstKind, class_sym: u32) -> u16 {
        let mut pool = self.const_pool.borrow_mut();
        if let Some(i) = pool.iter().position(
            |p| matches!(p, ModuleConstEntry::Class(k, s) if *k == kind && *s == class_sym),
        ) {
            return i as u16;
        }
        pool.push(ModuleConstEntry::Class(kind, class_sym));
        (pool.len() - 1) as u16
    }

    /// Intern a vtable-call's member-type-descriptor entry (the call
    /// opcode's own second operand) for the given callee class — deduped by
    /// `(MemberType, class_sym)`, so every vtable call against the SAME
    /// class shares one index (Get/Let/Set/method, any type, all share it),
    /// but a DIFFERENT class gets its own. Allocated lazily the first time
    /// any vtable call against that class needs it. Shares the index space
    /// with `Str`/`Class` entries. See `ModuleConstEntry::MemberType`.
    pub(super) fn intern_member_type_const(&self, class_sym: u32) -> u16 {
        let mut pool = self.const_pool.borrow_mut();
        if let Some(i) = pool
            .iter()
            .position(|p| matches!(p, ModuleConstEntry::MemberType(s) if *s == class_sym))
        {
            return i as u16;
        }
        pool.push(ModuleConstEntry::MemberType(class_sym));
        (pool.len() - 1) as u16
    }
}

impl<'m> LowerCtx<'m> {
    fn local_type(&self, idx: usize) -> &VbaType {
        &self.proc.locals[idx].vba_type
    }
    fn local_udt(&self, idx: usize) -> Option<UdtLocal> {
        self.local_udts[idx].clone()
    }
    fn local_class(&self, idx: usize) -> Option<&ExternalClass> {
        self.local_classes[idx].as_ref()
    }
    fn param_type(&self, idx: usize) -> &VbaType {
        &self.proc.params[idx].vba_type
    }
    fn global_type(&self, idx: usize) -> &VbaType {
        &self.module.module_vars[idx].vba_type
    }
    /// The frame offset of a class-member scratch region's FIRST slot for
    /// the given frame type-context — see `class_member_bases`. Every call
    /// site that reaches this has already had its own need for this exact
    /// context counted by `decl::class_member_regions` during frame
    /// construction, so the region is always present; a missing entry is an
    /// internal inconsistency between the counting pass and the emission
    /// pass, not a reachable user-facing error.
    pub(super) fn class_member_slot(&self, type_ctx: usize, index: usize) -> i16 {
        let base = *self
            .class_member_bases
            .get(&type_ctx)
            .expect("class_member_regions did not reserve a region for a context an emitter needed");
        self.local_slots[base + index].frame_offset
    }
}

#[cfg(test)]
#[path = "../tests/lower_tests.rs"]
mod tests;
