use super::*;
use std::io::Write;

fn write_temp(name: &str, bytes: &[u8]) -> std::path::PathBuf {
    let mut p = std::env::temp_dir();
    p.push(format!("frxcov_{}_{}", std::process::id(), name));
    std::fs::File::create(&p).unwrap().write_all(bytes).unwrap();
    p
}

#[test]
fn full_coverage_of_a_single_picture() {
    // companion: [u32 12][lt\0\0][u32 4][BM\0\0] = 16 bytes, fully covered by one ref
    let mut frx = Vec::new();
    frx.extend_from_slice(&12u32.to_le_bytes());
    frx.extend_from_slice(b"lt\0\0");
    frx.extend_from_slice(&4u32.to_le_bytes());
    frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
    let frx_path = write_temp("cov.frx", &frx);
    let name = frx_path.file_name().unwrap().to_string_lossy().to_string();
    let frm = format!("Begin VB.Form f\n   Icon = \"{}\":0000\nEnd\n", name);
    let frm_path = write_temp("cov.frm", frm.as_bytes());

    let reports = coverage_for_form(&frm_path).unwrap();
    assert_eq!(reports.len(), 1);
    let r = &reports[0];
    assert_eq!(r.total_len, 16);
    assert_eq!(r.covered, 16);
    assert!(r.is_complete(), "expected complete, got {:?}", r.gaps);
    assert!(r.overlaps.is_empty());
    let _ = std::fs::remove_file(frx_path);
    let _ = std::fs::remove_file(frm_path);
}

#[test]
fn unexplained_gap_is_flagged() {
    // 16-byte picture + 8 trailing non-zero bytes that no ref points at.
    let mut frx = Vec::new();
    frx.extend_from_slice(&12u32.to_le_bytes());
    frx.extend_from_slice(b"lt\0\0");
    frx.extend_from_slice(&4u32.to_le_bytes());
    frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
    frx.extend_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]); // unreferenced
    let frx_path = write_temp("gap.frx", &frx);
    let name = frx_path.file_name().unwrap().to_string_lossy().to_string();
    let frm = format!("Begin VB.Form f\n   Icon = \"{}\":0000\nEnd\n", name);
    let frm_path = write_temp("gap.frm", frm.as_bytes());

    let reports = coverage_for_form(&frm_path).unwrap();
    let r = &reports[0];
    assert!(!r.is_complete());
    assert_eq!(r.unexplained_bytes(), 8);
    let _ = std::fs::remove_file(frx_path);
    let _ = std::fs::remove_file(frm_path);
}
