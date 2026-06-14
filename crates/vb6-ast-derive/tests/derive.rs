//! Expansion behaviour of `#[derive(Children)]`.

use vb6_ast_derive::Children;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct NodeId(pub u32);

// Non-child fields below exist precisely to verify the derive ignores them.
#[allow(dead_code)]
#[derive(Children)]
enum Sample {
    Leaf,
    One {
        child: NodeId,
        scalar: u32,
    },
    OptAndVec {
        maybe: Option<NodeId>,
        many: Vec<NodeId>,
        // non-child collections / scalars are ignored
        ranges: Vec<(u32, u32)>,
        flag: bool,
    },
    Tuple(NodeId, u32, Option<NodeId>),
}

fn kids(n: &Sample) -> Vec<u32> {
    let mut out = Vec::new();
    n.for_each_child(&mut |id| out.push(id.0));
    out
}

#[test]
fn unit_has_no_children() {
    assert_eq!(kids(&Sample::Leaf), Vec::<u32>::new());
}

#[test]
fn named_direct_child_only() {
    assert_eq!(kids(&Sample::One { child: NodeId(7), scalar: 99 }), vec![7]);
}

#[test]
fn option_vec_and_ignored_fields() {
    let n = Sample::OptAndVec {
        maybe: Some(NodeId(1)),
        many: vec![NodeId(2), NodeId(3)],
        ranges: vec![(0, 9)],
        flag: true,
    };
    assert_eq!(kids(&n), vec![1, 2, 3]);

    let none = Sample::OptAndVec {
        maybe: None,
        many: vec![],
        ranges: vec![],
        flag: false,
    };
    assert_eq!(kids(&none), Vec::<u32>::new());
}

#[test]
fn tuple_variant_children_in_order() {
    assert_eq!(kids(&Sample::Tuple(NodeId(4), 0, Some(NodeId(5)))), vec![4, 5]);
}
