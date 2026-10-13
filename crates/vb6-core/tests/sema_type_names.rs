//! Declared type names: qualified names resolve by their last part, the named
//! types behind declarations are exported module-independently, a type-name
//! qualifier is not a variable use, and module attribute lines are recorded.

use std::collections::HashMap;

use vb6_core::frontend::ast::{ExprArena, NodeId};
use vb6_core::frontend::parser::{ModuleKind, Parser};
use vb6_core::frontend::scanner::ScannerContext;
use vb6_core::sema::{
    bind, bind_with_classes, unbound_namerefs, BoundModule, ExternalClass, TypeName, VbaType,
};

struct Bound {
    ctx: ScannerContext,
    arena: ExprArena,
    module: BoundModule,
}

fn bind_src(src: &str, classes: &HashMap<String, ExternalClass>) -> Bound {
    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::with_module_kind(&mut ctx, src.as_bytes(), ModuleKind::Class);
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = if classes.is_empty() {
        bind(&ctx, &arena, &top, &spans, &vis)
    } else {
        bind_with_classes(&ctx, &arena, &top, &spans, &vis, classes)
    };
    Bound { ctx, arena, module }
}

fn tn(q: Option<&str>, n: &str) -> Option<TypeName> {
    Some(TypeName { qualifier: q.map(String::from), name: n.to_string() })
}

#[test]
fn qualified_module_variable_type_is_the_last_name() {
    let b = bind_src("Public o As Proj.clsTrip\n", &HashMap::new());
    let VbaType::UserDefined(sym) = b.module.module_vars[0].vba_type else { panic!("expected UserDefined") };
    assert_eq!(b.ctx.symbol(sym as usize).name, "clsTrip");
    assert_eq!(b.module.decl_types.module_vars, vec![tn(Some("proj"), "clstrip")]);
}

#[test]
fn qualified_type_tail_is_not_a_variable_use() {
    let b = bind_src("Public o As Proj.clsTrip\nSub T()\n    Dim p As New Proj.clsTrip\nEnd Sub\n", &HashMap::new());
    assert!(b.module.resolutions.is_empty(), "no name was resolved: {:?}", b.module.resolutions);
    assert_eq!(unbound_namerefs(&b.arena, &b.module.resolutions), Vec::<NodeId>::new());
}

#[test]
fn unqualified_and_builtin_names() {
    let b = bind_src(
        "Public a As clsTrip\nPublic n As Long\nPublic Function F() As Proj.clsTrip\nEnd Function\nPublic Sub S()\nEnd Sub\n",
        &HashMap::new(),
    );
    assert_eq!(b.module.decl_types.module_vars, vec![tn(None, "clstrip"), None]);
    assert_eq!(b.module.decl_types.proc_rets, vec![tn(Some("proj"), "clstrip"), None]);
}

#[test]
fn type_member_names_align_with_members() {
    let b = bind_src(
        "Public Type TTrip\n    Id As Long\n    Owner As Proj.clsPerson\n    Dest As TPlace\nEnd Type\n",
        &HashMap::new(),
    );
    assert_eq!(
        b.module.decl_types.type_members,
        vec![vec![None, tn(Some("proj"), "clsperson"), tn(None, "tplace")]]
    );
}

#[test]
fn qualified_type_keys_class_info_by_the_class_name() {
    let mut classes = HashMap::new();
    classes.insert("clstrip".to_string(), ExternalClass::default());
    let b = bind_src("Dim o As Proj.clsTrip\n", &classes);
    let keys: Vec<String> =
        b.module.class_field_info.keys().map(|s| b.ctx.symbol(*s as usize).name.clone()).collect();
    assert_eq!(keys, vec!["clsTrip".to_string()]);
}

#[test]
fn module_attributes_are_recorded() {
    let src = "Attribute VB_Name = \"clsTrips\"\nAttribute VB_PredeclaredId = True\n\
               Attribute VB_GlobalNameSpace = True\nPublic Function Item(i As Long) As Long\n\
               Attribute Item.VB_UserMemId = 0\nEnd Function\n";
    let b = bind_src(src, &HashMap::new());
    assert_eq!(b.module.vb_name.as_deref(), Some("clsTrips"));
    assert!(b.module.predeclared_id);
    assert!(b.module.global_namespace);
    assert_eq!(b.module.default_member.as_deref(), Some("item"));
}

#[test]
fn attribute_defaults_when_absent_or_false() {
    let b = bind_src(
        "Attribute VB_PredeclaredId = False\nAttribute VB_GlobalNameSpace = False\nPublic Sub S()\nEnd Sub\n",
        &HashMap::new(),
    );
    assert_eq!(b.module.vb_name, None);
    assert!(!b.module.predeclared_id);
    assert!(!b.module.global_namespace);
    assert_eq!(b.module.default_member, None);
}

#[test]
fn default_variable_member() {
    let b = bind_src("Public Value As Long\nAttribute Value.VB_VarUserMemId = 0\n", &HashMap::new());
    assert_eq!(b.module.default_member.as_deref(), Some("value"));
}
