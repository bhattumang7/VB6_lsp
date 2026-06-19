//! The binder front-half: resolving a reference's context node to the bind
//! handle the resolver reads (convention kind + by-reference mode).
//!
//! `EbResolveIdentRef`'s category-4 path calls `EbResolveExprNode(pNode[5])` to
//! turn the context node into a small handle (the `auStack_24` scratch the source
//! threads through `EbGetBindResult`), then reads its convention kind
//! (`EbGetTypeKind2`) and by-reference mode (`EbGetByRefFlag`). This module ports
//! that chain.
//!
//! The bind structures (context nodes, symbol nodes) are byte-addressed in a
//! buffer (`bmem`); the handle is a small value struct. Disc-5 (resolved module
//! member) reads its convention from the compiled member record — see
//! [`crate::sym_record::MemberRecord`].

use crate::sym_record::MemberRecord;

/// The bind handle (`EbResolveExprNode`'s result — the `auStack_24` scratch).
/// `disc` is the kind tag the readers switch on (`1` = symbol/type-lib item,
/// `3`/`6` = context, `5` = resolved module member); the other slots carry the
/// node / symbol-base / member-offset the readers dereference.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct BindHandle {
    /// `[0]` — the discriminator.
    pub disc: i32,
    /// `[1]` — context pointer (disc 3/6).
    pub w1: u32,
    /// `[2]` — symbol base (disc 5: the records-heap base pointer).
    pub w2: u32,
    /// `[3]` — the bound node (disc 1) or member byte offset (disc 5).
    pub w3: u32,
    /// `[4]`.
    pub w4: u32,
    /// `[5]`.
    pub w5: u32,
}

fn u32_at(bmem: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([bmem[off], bmem[off + 1], bmem[off + 2], bmem[off + 3]])
}

/// Port of `EbGetBindResult`: the handle is valid only when its discriminator is
/// non-zero (`-(uint)(*pResult != 0) & pResult`).
pub fn get_bind_result(h: BindHandle) -> Option<BindHandle> {
    if h.disc != 0 {
        Some(h)
    } else {
        None
    }
}

/// Port of `EbSetupResult`: build a disc-1 handle for the bound `node` (a symbol
/// AST node at offset `node_off` in `bmem`). A null node clears the handle (the
/// `EbClearErrorContext` path). A node whose type nibble (`node+8 & 0xf`) is `9`
/// suppresses the auxiliary slots.
pub fn setup_result(bmem: &[u8], node_off: u32, p2: u32, p3: u32) -> BindHandle {
    if node_off == 0 {
        return BindHandle::default();
    }
    let mut h = BindHandle {
        disc: 1,
        w3: node_off,
        ..BindHandle::default()
    };
    let node_type = bmem[node_off as usize + 8] & 0xf;
    if node_type != 9 {
        h.w4 = p2;
        h.w5 = p3;
    }
    h
}

/// Port of `EbMapTypeToCallConv`: the convention kind for a disc-1 (symbol) bind
/// handle. Reads the symbol's type descriptor (`*(symbol+0x20)`) flags word at
/// `+0x12`, sign-extends its low 3 bits, and maps `1→1, 2→2, 3→6`, everything
/// else (including `0` and the negative values) → `4`.
pub fn map_type_to_call_conv(bmem: &[u8], symbol_off: usize) -> i32 {
    let type_desc = u32_at(bmem, symbol_off + 0x20) as usize;
    let word = u32_at(bmem, type_desc + 0x12) as i32;
    let v = (word << 29) >> 29; // sign-extend the low 3 bits
    match v {
        1 => 1,
        2 => 2,
        3 => 6,
        _ => 4,
    }
}

/// Port of `EbGetTypeKind2`: the convention kind for a bind handle. Disc 1 maps
/// via [`map_type_to_call_conv`]; disc 5 reads the resolved member record; any
/// other discriminator is the plain-value kind `4`.
pub fn handle_kind(h: &BindHandle, bmem: &[u8], member: Option<&MemberRecord>) -> i32 {
    match h.disc {
        1 => map_type_to_call_conv(bmem, h.w3 as usize),
        5 => member
            .expect("disc-5 bind handle requires its member record")
            .kind(),
        _ => 4,
    }
}

/// Port of `EbGetByRefFlag`: the by-reference mode for a bind handle. Only a
/// disc-5 (resolved member) handle carries one (from its record); every other
/// discriminator is by-value (`0`).
pub fn handle_byref(h: &BindHandle, member: Option<&MemberRecord>) -> i32 {
    match h.disc {
        5 => member
            .expect("disc-5 bind handle requires its member record")
            .byref(),
        _ => 0,
    }
}

/// Port of `EbSetupContext`: a disc-3 context handle for `context`. The
/// convention readers treat disc 3 as the plain-value kind (`4`), so only the
/// discriminator and the context pointer matter to them; the document-context
/// slot (`+0xe0 → +8`) the source also stores is a non-emitting heap side-effect
/// (`EbGetDocumentContext`) left unmodelled here (it never feeds the convention
/// kind/byref readers).
pub fn setup_context(context: u32) -> BindHandle {
    BindHandle {
        disc: 3,
        w1: context,
        ..BindHandle::default()
    }
}

/// Port of `EbResolveExprNode`: resolve a context node to its bind handle.
///
/// * disc 1 — a name/type reference: dereference `pNode[3]` to the symbol and
///   build a disc-1 handle ([`setup_result`]). The `node_type == 10` sub-case
///   (clear + empty result) yields `None`.
/// * disc 2 — empty result → `None`.
/// * disc 3 — `EbInitializeContext2`: a disc-2 handle carrying
///   `*(*(pNode[1]+0x24)+0x5c)`.
/// * disc 4 — `pNode[4] != -1` is the COM `ITypeInfo` slot path
///   (`EbEnsureSlotResolved2`, gated); `pNode[5] != -1` is the error path
///   (`EbRaiseException2`, → `None`); otherwise the context path
///   ([`setup_context`]).
/// * disc 5/6 — the context path ([`setup_context`]).
///
/// `ctx_off` is the context node's offset in `bmem`; its dwords are `pNode[0..]`.
/// The convention readers map every context discriminator (2/3/6) to the
/// plain-value kind 4, so the document-context allocation those paths perform in
/// the source — which emits no p-code — is not reproduced here.
pub fn resolve_expr_node(bmem: &[u8], ctx_off: u32) -> Option<BindHandle> {
    let base = ctx_off as usize;
    let disc = u32_at(bmem, base) as i32;
    match disc {
        1 => {
            let p3 = u32_at(bmem, base + 12);
            let symbol = u32_at(bmem, p3 as usize); // EbDereference(pNode[3]) = *pNode[3]
            let node_type = bmem[symbol as usize + 8] & 0xf;
            if node_type == 10 {
                return None; // EbClearErrorContext path → empty result
            }
            let p4 = u32_at(bmem, base + 16);
            get_bind_result(setup_result(bmem, symbol, p4, 0))
        }
        2 => None,
        3 => {
            // EbInitializeContext2(*(*(pNode[1]+0x24)+0x5c)) → disc-2 handle.
            let p1 = u32_at(bmem, base + 4) as usize;
            let a = u32_at(bmem, p1 + 0x24) as usize;
            let init = u32_at(bmem, a + 0x5c);
            get_bind_result(BindHandle {
                disc: 2,
                w3: init,
                ..BindHandle::default()
            })
        }
        4 => {
            let p5 = u32_at(bmem, base + 20) as i32;
            if p5 != -1 {
                return None; // EbRaiseException2 error path
            }
            let p4 = u32_at(bmem, base + 16) as i32;
            if p4 != -1 {
                unimplemented!(
                    "EbResolveExprNode disc-4 slot: EbEnsureSlotResolved2 (COM \
                     ITypeInfo); Phase 6"
                );
            }
            let p1 = u32_at(bmem, base + 4);
            get_bind_result(setup_context(p1))
        }
        5 | 6 => {
            let p1 = u32_at(bmem, base + 4);
            get_bind_result(setup_context(p1))
        }
        _ => None,
    }
}

#[cfg(test)]
#[path = "tests/binder_tests.rs"]
mod tests;
