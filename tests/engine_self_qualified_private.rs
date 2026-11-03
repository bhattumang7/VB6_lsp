//! A module reaches its own `Private` members through its own name
//! (`ThisModule.Helper`), while other modules still cannot.

use vb6_engine::session::Session;

const METHOD_OR_DATA_MEMBER_NOT_FOUND: u32 = 461;

fn codes(files: &[(&str, &str)], module: usize) -> Vec<u32> {
    let s = Session::from_sources(
        files.iter().map(|(n, c)| (n.to_string(), c.as_bytes().to_vec())).collect(),
    );
    s.diagnostics(module).iter().map(|d| d.code).collect()
}

const THIS: &str = "Attribute VB_Name = \"ThisModule\"\nOption Explicit\n\nPrivate mCount As Long\nPrivate Const LIMIT As Long = 3\nPrivate Enum Mode\n    ModeA = 1\nEnd Enum\n\nPublic Sub RunAll()\n    Call ThisModule.Helper\n    ThisModule.Helper\n    ThisModule.mCount = ThisModule.Compute(1) + ThisModule.LIMIT\n    Dim m As Long\n    m = ThisModule.ModeA\nEnd Sub\n\nPrivate Sub Helper()\nEnd Sub\n\nPrivate Function Compute(ByVal n As Long) As Long\n    Compute = n\nEnd Function\n";

#[test]
fn own_private_members_through_own_module_name() {
    assert_eq!(codes(&[("ThisModule.bas", THIS)], 0), Vec::<u32>::new());
}

#[test]
fn other_module_cannot_reach_private_members() {
    let other = "Attribute VB_Name = \"Other\"\nOption Explicit\n\nPublic Sub Go()\n    ThisModule.Helper\n    ThisModule.mCount = 1\nEnd Sub\n";
    let c = codes(&[("ThisModule.bas", THIS), ("Other.bas", other)], 1);
    assert_eq!(c, vec![METHOD_OR_DATA_MEMBER_NOT_FOUND, METHOD_OR_DATA_MEMBER_NOT_FOUND]);
}

#[test]
fn unknown_member_through_own_name_is_still_reported() {
    let src = "Attribute VB_Name = \"ThisModule\"\nPublic Sub A()\n    ThisModule.Nope\nEnd Sub\n";
    assert_eq!(codes(&[("ThisModule.bas", src)], 0), vec![METHOD_OR_DATA_MEMBER_NOT_FOUND]);
}

#[test]
fn class_private_members_stay_unreachable_through_an_object() {
    // The class's own name or an instance of it exposes the public interface
    // only, from inside the class as well.
    let cls = "Attribute VB_Name = \"Counter\"\nOption Explicit\nPrivate mN As Long\nPrivate Sub Bump()\nEnd Sub\nPublic Sub Copy(ByVal o As Counter)\n    o.Bump\n    mN = o.mN\nEnd Sub\n";
    assert_eq!(
        codes(&[("Counter.cls", cls)], 0),
        vec![METHOD_OR_DATA_MEMBER_NOT_FOUND, METHOD_OR_DATA_MEMBER_NOT_FOUND]
    );
}
