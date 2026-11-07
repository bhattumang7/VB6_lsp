use super::*;
use std::io::Write;

/// A COM decoder stub that always reports the license as missing (hard error),
/// to exercise the no-soft-fallback contract.
struct UnlicensedDecoder;
impl ComDecoder for UnlicensedDecoder {
    fn decode_bag(
        &self,
        _control: &BagControl,
        property: &str,
        _bag: &[u8],
    ) -> Result<FrxValue, ResolveError> {
        // A stub that never resolves a coclass, so it can't name one.
        Err(ResolveError::LicensedBagUnavailable {
            clsid: None,
            property: property.to_string(),
        })
    }
}

fn write_temp(name: &str, bytes: &[u8]) -> PathBuf {
    let mut p = std::env::temp_dir();
    p.push(format!("frxtest_{}_{}", std::process::id(), name));
    let mut f = std::fs::File::create(&p).unwrap();
    f.write_all(bytes).unwrap();
    p
}

#[test]
fn missing_companion_is_explicit() {
    let frm = "Begin VB.Form f\n   Icon = \"nope.frx\":0000\nEnd\n";
    let frm_path = write_temp("missing.frm", frm.as_bytes());
    let res = resolve_form_resources(&frm_path, OcxBagPolicy::Opaque, None).unwrap();
    assert_eq!(res.len(), 1);
    assert!(matches!(
        res[0].value,
        Err(ResolveError::MissingCompanion { .. })
    ));
    let _ = std::fs::remove_file(frm_path);
}

#[test]
fn decodes_real_companion_picture() {
    // Build a valid picture blob: [u32 12][lt\0\0][u32 4]["BM\0\0"]
    let mut frx = Vec::new();
    frx.extend_from_slice(&12u32.to_le_bytes());
    frx.extend_from_slice(b"lt\0\0");
    frx.extend_from_slice(&4u32.to_le_bytes());
    frx.extend_from_slice(&[0x42, 0x4D, 0x00, 0x00]);
    let frx_path = write_temp("pic.frx", &frx);
    let frx_name = frx_path.file_name().unwrap().to_string_lossy().to_string();
    let frm = format!("Begin VB.Form f\n   Icon = \"{}\":0000\nEnd\n", frx_name);
    let frm_path = write_temp("pic.frm", frm.as_bytes());

    let res = resolve_form_resources(&frm_path, OcxBagPolicy::Opaque, None).unwrap();
    assert_eq!(res.len(), 1);
    match &res[0].value {
        Ok(FrxValue::Picture { format, .. }) => {
            assert_eq!(*format, frx::ImageFormat::Bmp)
        }
        other => panic!("expected BMP picture, got {:?}", other),
    }
    let _ = std::fs::remove_file(frx_path);
    let _ = std::fs::remove_file(frm_path);
}

#[test]
fn com_decode_hard_errors_without_license() {
    // A control bag reference + ComDecode policy + an unlicensed decoder => hard error.
    let mut frx = Vec::new();
    frx.extend_from_slice(&8u32.to_le_bytes());
    frx.extend_from_slice(&[0u8; 8]);
    let frx_path = write_temp("bag.frx", &frx);
    let frx_name = frx_path.file_name().unwrap().to_string_lossy().to_string();
    let frm = format!("Begin VB.Form f\n   _GridInfo = \"{}\":0000\nEnd\n", frx_name);
    let frm_path = write_temp("bag.frm", frm.as_bytes());

    let dec = UnlicensedDecoder;
    let res =
        resolve_form_resources(&frm_path, OcxBagPolicy::ComDecode, Some(&dec)).unwrap();
    assert_eq!(res.len(), 1);
    assert!(
        matches!(res[0].value, Err(ResolveError::LicensedBagUnavailable { .. })),
        "ComDecode without a license must hard-error, got {:?}",
        res[0].value
    );
    let _ = std::fs::remove_file(frx_path);
    let _ = std::fs::remove_file(frm_path);
}

#[test]
fn opaque_policy_passes_bag_through() {
    let mut frx = Vec::new();
    frx.extend_from_slice(&4u32.to_le_bytes());
    frx.extend_from_slice(&[0xDE, 0xAD, 0xBE, 0xEF]);
    let frx_path = write_temp("op.frx", &frx);
    let frx_name = frx_path.file_name().unwrap().to_string_lossy().to_string();
    let frm = format!("Begin VB.Form f\n   Bands = \"{}\":0000\nEnd\n", frx_name);
    let frm_path = write_temp("op.frm", frm.as_bytes());
    let res = resolve_form_resources(&frm_path, OcxBagPolicy::Opaque, None).unwrap();
    assert!(matches!(res[0].value, Ok(FrxValue::OcxBag { .. })));
    let _ = std::fs::remove_file(frx_path);
    let _ = std::fs::remove_file(frm_path);
}
