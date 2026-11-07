//! The order and shape rules for the parameters of one list: `Optional` parameters follow the
//! required ones, `ParamArray` is last and an array, events take neither, and an array cannot be
//! `ByVal` or `Optional`.

use super::*;

/// A rule a parameter breaks: the error code and, for a missing-token error, what was expected.
pub(super) type Violation = (u32, Option<&'static str>);

/// What the rules need to know about one parameter of a list.
pub(super) struct ParamSlot {
    pub(super) owner: ParamOwner,
    pub(super) flags: u16,
    /// The place of the parameter in the list.
    pub(super) index: usize,
    /// The number of parameters in the list.
    pub(super) count: usize,
}

impl ParamSlot {
    fn is_optional(&self) -> bool {
        self.flags & 0x01 != 0
    }

    fn is_by_val(&self) -> bool {
        self.flags & 0x02 != 0
    }

    fn is_array(&self) -> bool {
        self.flags & 0x08 != 0
    }

    fn is_param_array(&self) -> bool {
        self.flags & 0x20 != 0
    }

    /// The owner of the list has a value parameter, which is its last.
    fn is_value_owner(&self) -> bool {
        matches!(self.owner, ParamOwner::PropertyValue | ParamOwner::PropertySet)
    }

    fn is_value_param(&self) -> bool {
        self.is_value_owner() && self.is_last()
    }

    fn is_last(&self) -> bool {
        self.index + 1 == self.count
    }

    /// The section the list is in after this parameter: 0 = none yet, 1 = an Optional seen,
    /// 2 = a ParamArray seen.
    pub(super) fn section_after(&self, section: u8) -> u8 {
        if self.is_param_array() {
            2
        } else if self.is_optional() {
            section.max(1)
        } else {
            section
        }
    }
}

/// The first rule that the parameter `param` of the list `args` breaks, given the `section` the
/// list is in. Only the first violation of a list is reported, as the compiler stops at the
/// first error of a line.
pub(super) fn first_violation(
    arena: &ExprArena,
    args: &[NodeId],
    param: NodeId,
    slot: &ParamSlot,
    section: u8,
) -> Option<Violation> {
    duplicate_name(arena, args, param, slot)
        .or_else(|| property_set_value(arena, param, slot))
        .or_else(|| shape_violation(arena, param, slot, section))
        .or_else(|| by_val_array(slot))
}

/// A name already used by an earlier parameter.
fn duplicate_name(arena: &ExprArena, args: &[NodeId], param: NodeId, slot: &ParamSlot) -> Option<Violation> {
    if !matches!(slot.owner, ParamOwner::Procedure | ParamOwner::PropertyValue | ParamOwner::PropertySet) {
        return None;
    }
    let ExprNode::ParamDef { name: param_name, .. } = arena.get(param) else { return None };
    let used = args[..slot.index].iter().any(|&earlier| {
        matches!(arena.get(earlier), ExprNode::ParamDef { name, .. } if name == param_name)
    });
    used.then_some((ERR_DUPLICATE_DECLARATION, None))
}

/// The value of a `Property Set` is an object: untyped, Variant, Object or a class.
fn property_set_value(arena: &ExprArena, param: NodeId, slot: &ParamSlot) -> Option<Violation> {
    if slot.owner != ParamOwner::PropertySet
        || !slot.is_value_param()
        || slot.is_optional()
        || slot.is_param_array()
    {
        return None;
    }
    let ExprNode::ParamDef { type_node, .. } = arena.get(param) else { return None };
    let object_typed = !slot.is_array()
        && type_node.is_none_or(|t| {
            matches!(arena.get(t), ExprNode::BuiltinType { kind: 12 | 9 } | ExprNode::UserType { .. })
        });
    (!object_typed).then_some((ERR_PROPERTY_PROCEDURES_INCONSISTENT, None))
}

/// The rules for the kind of parameter this is: event, `ParamArray`, `Optional` or required.
fn shape_violation(arena: &ExprArena, param: NodeId, slot: &ParamSlot, section: u8) -> Option<Violation> {
    if slot.owner == ParamOwner::Event && (slot.is_optional() || slot.is_param_array()) {
        return Some((ERR_SYNTAX_ERROR, None));
    }
    if slot.is_param_array() {
        return param_array_violation(arena, param, slot, section);
    }
    if slot.is_optional() {
        return optional_violation(slot, section);
    }
    if section != 0 && !slot.is_value_param() {
        return Some((ERR_UNEXPECTED_TOKEN, Some(if section == 1 { "Optional" } else { ")" })));
    }
    None
}

fn param_array_violation(arena: &ExprArena, param: NodeId, slot: &ParamSlot, section: u8) -> Option<Violation> {
    if section == 2 {
        return Some((ERR_UNEXPECTED_TOKEN, Some("identifier")));
    }
    if section == 1 {
        return Some((ERR_PARAM_ARRAY_WITH_OPTIONAL, None));
    }
    if !slot.is_array() || (!slot.is_value_owner() && !param_type_is_variant(arena, param)) {
        return Some((ERR_PARAM_ARRAY_NOT_ARRAY, None));
    }
    if !slot.is_last() && !slot.is_value_owner() {
        return Some((ERR_UNEXPECTED_TOKEN, Some(")")));
    }
    None
}

fn optional_violation(slot: &ParamSlot, section: u8) -> Option<Violation> {
    if section == 2 {
        return Some((ERR_UNEXPECTED_TOKEN, Some(")")));
    }
    if slot.is_array() {
        return Some((ERR_OPTIONAL_ARRAY, None));
    }
    None
}

/// An array is always passed by reference.
fn by_val_array(slot: &ParamSlot) -> Option<Violation> {
    (slot.is_array() && slot.is_by_val()).then_some((ERR_ARRAY_ARGUMENT_MUST_BE_BYREF, None))
}
