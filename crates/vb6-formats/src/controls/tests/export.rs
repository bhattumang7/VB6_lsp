use super::*;
use std::io::Write;

fn write_temp(name: &str, bytes: &[u8]) -> std::path::PathBuf {
    let mut p = std::env::temp_dir();
    p.push(format!("frxexp_{}_{}", std::process::id(), name));
    std::fs::File::create(&p).unwrap().write_all(bytes).unwrap();
    p
}

#[test]
fn builds_and_serializes_full_export() {
    // companion: one [u32 12][lt\0\0][u32 4][BM..] picture, fully covered.
    let mut frx = Vec::new();
    frx.extend_from_slice(&12u32.to_le_bytes());
    frx.extend_from_slice(b"lt\0\0");
    frx.extend_from_slice(&4u32.to_le_bytes());
    frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
    let frx_path = write_temp("exp.frx", &frx);
    let frx_name = frx_path.file_name().unwrap().to_string_lossy().to_string();
    let frm = format!("Begin VB.Form frmX\n   Icon = \"{}\":0000\nEnd\n", frx_name);
    let frm_path = write_temp("exp.frm", frm.as_bytes());

    let exp = build_form_export(&frm_path).unwrap();
    let root = exp.root.as_ref().expect("root control");
    assert_eq!(root.type_name, "VB.Form");
    assert!(root.is_intrinsic);
    // exactly one resolved resource, decoded as a BMP picture
    assert_eq!(exp.resources.len(), 1);
    assert_eq!(exp.resources[0].property, "Icon");
    assert!(matches!(
        exp.resources[0].value,
        ResourceValueView::Picture { .. }
    ));
    // coverage proves the companion is fully accounted for
    assert_eq!(exp.coverage.len(), 1);
    assert!(exp.coverage[0].complete);
    assert_eq!(exp.coverage[0].covered, exp.coverage[0].total);

    // serialises to JSON with the tagged resource value
    let json = serde_json::to_string(&exp).unwrap();
    assert!(json.contains("\"type\":\"Picture\""));
    assert!(json.contains("\"format\":\"Bmp\""));

    let _ = std::fs::remove_file(frx_path);
    let _ = std::fs::remove_file(frm_path);
}
