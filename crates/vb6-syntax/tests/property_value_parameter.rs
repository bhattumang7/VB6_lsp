//! A `Property Let` or `Set` takes its value as the parameter after every Optional or
//! ParamArray one; a list that ends without a required value is "Argument not optional".

mod common;
use common::{offset_of, parse};

const ARGUMENT_NOT_OPTIONAL: u32 = 0x1c1;

fn let_prop(params: &str) -> String {
    format!("Property Let Prp({params})\nEnd Property\n")
}

#[test]
fn a_list_ending_in_optional_or_paramarray_is_rejected_at_the_closing_bracket() {
    for params in [
        "Optional v As Long",
        "ParamArray v() As Variant",
        "a, Optional b",
        "ByVal Optional a As Long",
        "Optional a = 1, Optional b = 2",
    ] {
        let src = let_prop(params);
        let p = parse(&src);
        assert_eq!(p.found, vec![(ARGUMENT_NOT_OPTIONAL, offset_of(&src, ")\n", 0), 1)], "{params}");
    }
    let src = "Property Set Prp(Optional v As Object)\nEnd Property\n";
    assert_eq!(parse(src).found, vec![(ARGUMENT_NOT_OPTIONAL, offset_of(src, ")\n", 0), 1)]);
}

#[test]
fn a_required_value_after_an_optional_parameter_is_accepted() {
    for params in ["Optional a, v", "a, Optional b, v As Long", "v As Long", "a, v"] {
        let p = parse(&let_prop(params));
        assert_eq!(p.found, vec![], "{params}");
    }
}

#[test]
fn property_get_may_take_optional_parameters() {
    let p = parse("Property Get Prp(Optional a As Long) As Long\nEnd Property\n");
    assert_eq!(p.found, vec![]);
}
