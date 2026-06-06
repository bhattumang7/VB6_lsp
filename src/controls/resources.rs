//! Resolve a form's designer references to decoded companion blobs.
//!
//! Ties Tier 1 ([`super::form_designer`]) to Tier 2 ([`super::frx`]) and the Tier 3
//! opaque/COM policy. Two invariants the feature requires:
//!  * **never silently mis-read** — a missing companion or an unreadable blob is
//!    reported explicitly, not swallowed;
//!  * for proprietary control bags, the default is an opaque pass-through, but the
//!    opt-in COM-decode path **hard-errors when the control's license is absent**
//!    (rather than degrading to opaque).

use std::path::{Path, PathBuf};

use super::form_designer::{self, ResourceRef};
use super::frx::{self, FrxError, FrxValue, PropKind};

/// How to treat a proprietary (vendor) control bag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OcxBagPolicy {
    /// Surface the raw bytes opaquely (works everywhere, needs no control).
    Opaque,
    /// Decode via the live control (COM). Requires the control + its license.
    ComDecode,
}

impl Default for OcxBagPolicy {
    fn default() -> Self {
        OcxBagPolicy::Opaque
    }
}

/// Everything we know about the control that produced a bag, used by a
/// [`ComDecoder`] to instantiate the right coclass.
///
/// The GUID on a `.frm` `Object=` line is the control's **type-library** id, not
/// the **coclass** CLSID that `CoCreateInstance` needs — so the decoder is given
/// the OCX file and class name as well, and resolves the coclass from the typelib.
/// Identity of the control that produced a bag. `ocx_files`, `typelib_clsids`, and
/// `versions` are **index-aligned** — entry _i_ describes one `Object=` line.
#[derive(Debug, Default, Clone)]
pub struct BagControl {
    /// Control class name, e.g. `"MSChart"` (the segment after the dot in the
    /// designer control type `"MSChart20Lib.MSChart"`).
    pub class_name: Option<String>,
    /// Control library name, e.g. `"MSChart20Lib"` (the segment before the dot).
    /// Used to pick the *right* candidate typelib by its library name, so a
    /// coclass-name collision across two referenced OCXs can't mis-resolve.
    pub lib_name: Option<String>,
    /// Every OCX file named by an `Object=` line, e.g. `["MShflxgd.ocx",
    /// "MSChrt20.ocx"]`. The decoder loads each typelib and uses the one whose
    /// library/coclass match — the library→file name is too fuzzy to match in-process.
    pub ocx_files: Vec<String>,
    /// The `Object=` type-library GUIDs (aligned with `ocx_files`). Used to load the
    /// typelib from the registry when the OCX isn't in a system directory.
    pub typelib_clsids: Vec<String>,
    /// The `Object=` type-library versions, e.g. `"2.0"` (aligned with `ocx_files`).
    pub versions: Vec<String>,
    /// A CLSID embedded in the bag body itself, if one was detectable.
    pub embedded_clsid: Option<String>,
}

/// A pluggable bridge to a live-control COM decoder (implemented out-of-process on
/// Windows). Kept as a trait so the core stays testable and platform-independent.
pub trait ComDecoder {
    /// Decode a control's persisted bag. The implementation must return
    /// [`ResolveError::LicensedBagUnavailable`] (a hard error) when the control
    /// exists but its license is not present — never a silent opaque fallback.
    fn decode_bag(
        &self,
        control: &BagControl,
        property: &str,
        bag: &[u8],
    ) -> Result<FrxValue, ResolveError>;
}

/// Why a single reference could not be resolved.
#[derive(Debug)]
pub enum ResolveError {
    /// The `.frm` referenced a companion file that is not on disk.
    MissingCompanion { file: String, path: PathBuf },
    /// The companion file could not be read.
    Read(std::io::Error),
    /// The blob framing was invalid for the expected kind.
    Decode(FrxError),
    /// A proprietary bag required COM decode but the control's license is absent.
    /// This is intentionally fatal for the COM path (per project policy).
    LicensedBagUnavailable { clsid: Option<String>, property: String },
    /// COM decode was requested but no decoder bridge was supplied.
    NoComDecoder,
    /// COM decode was attempted but failed for a non-license reason (control not
    /// registered, load failed, bridge error). Also a hard error — never opaque.
    ComDecodeFailed { property: String, message: String },
}

impl std::fmt::Display for ResolveError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ResolveError::MissingCompanion { file, path } => {
                write!(f, "companion file `{}` not found at {}", file, path.display())
            }
            ResolveError::Read(e) => write!(f, "reading companion failed: {}", e),
            ResolveError::Decode(e) => write!(f, "decode failed: {}", e),
            ResolveError::LicensedBagUnavailable { clsid, property } => write!(
                f,
                "proprietary bag `{}` (control {}) requires a control license that is not installed",
                property,
                clsid.as_deref().unwrap_or("<unknown clsid>")
            ),
            ResolveError::NoComDecoder => write!(f, "COM decode requested but no decoder available"),
            ResolveError::ComDecodeFailed { property, message } => {
                write!(f, "COM decode of `{}` failed: {}", property, message)
            }
        }
    }
}
impl std::error::Error for ResolveError {}

/// One resolved (or failed) reference.
#[derive(Debug)]
pub struct ResolvedResource {
    pub reference: ResourceRef,
    pub kind: PropKind,
    pub value: Result<FrxValue, ResolveError>,
}

/// Read a `.frm`/`.ctl`, parse its designer block, and resolve every companion
/// reference. Companion files are read relative to the form's directory.
pub fn resolve_form_resources(
    frm_path: &Path,
    policy: OcxBagPolicy,
    com: Option<&dyn ComDecoder>,
) -> std::io::Result<Vec<ResolvedResource>> {
    let raw = std::fs::read(frm_path)?;
    let source = String::from_utf8_lossy(&raw);
    let designer = form_designer::parse_designer(&source);
    let dir = frm_path.parent().unwrap_or_else(|| Path::new("."));

    // Cache companion file contents so each is read at most once.
    let mut cache: std::collections::HashMap<String, std::io::Result<Vec<u8>>> =
        std::collections::HashMap::new();

    let mut out = Vec::new();
    for reference in designer.resource_refs() {
        let kind = frx::kind_for_property(&reference.property, reference.frx.dollar);
        let value = resolve_one(dir, &reference, kind, &designer.objects, policy, com, &mut cache);
        out.push(ResolvedResource { reference, kind, value });
    }
    Ok(out)
}

#[allow(clippy::too_many_arguments)]
fn resolve_one(
    dir: &Path,
    reference: &ResourceRef,
    kind: PropKind,
    objects: &[form_designer::ObjectEntry],
    policy: OcxBagPolicy,
    com: Option<&dyn ComDecoder>,
    cache: &mut std::collections::HashMap<String, std::io::Result<Vec<u8>>>,
) -> Result<FrxValue, ResolveError> {
    let file = &reference.frx.file;
    let companion = dir.join(file);
    let entry = cache
        .entry(file.clone())
        .or_insert_with(|| std::fs::read(&companion));
    let bytes = match entry {
        Ok(b) => b,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            return Err(ResolveError::MissingCompanion {
                file: file.clone(),
                path: companion,
            })
        }
        Err(e) => {
            return Err(ResolveError::Read(std::io::Error::new(e.kind(), e.to_string())))
        }
    };

    let offset = reference.frx.offset as usize;

    if kind == PropKind::OcxBag {
        // Tier 3: opaque by default; COM decode hard-errors if license absent.
        match policy {
            OcxBagPolicy::Opaque => Ok(frx::decode_ocx_bag(bytes, offset)),
            OcxBagPolicy::ComDecode => {
                let com = com.ok_or(ResolveError::NoComDecoder)?;
                // Hand the bag body to the live-control decoder (it owns the
                // hard-error-on-missing-license contract). The decoder resolves the
                // coclass from the OCX typelib + class name; the `Object=` GUID and
                // any CLSID embedded in the bag are passed as fallbacks.
                let (bag, embedded) = match frx::decode_ocx_bag(bytes, offset) {
                    FrxValue::OcxBag { data, clsid } => (data, clsid),
                    other => return Ok(other),
                };
                let (lib_name, class_name) = match reference.control_type.split_once('.') {
                    Some((l, c)) => (Some(l.to_string()), Some(c.to_string())),
                    None => (None, Some(reference.control_type.clone())),
                };
                let control = BagControl {
                    class_name,
                    lib_name,
                    ocx_files: objects.iter().map(|o| o.file.clone()).collect(),
                    typelib_clsids: objects.iter().map(|o| o.clsid.clone()).collect(),
                    versions: objects.iter().map(|o| o.version.clone()).collect(),
                    embedded_clsid: embedded.map(format_guid16),
                };
                com.decode_bag(&control, &reference.property, &bag)
            }
        }
    } else {
        frx::decode(bytes, offset, kind).map_err(ResolveError::Decode)
    }
}

/// Format a 16-byte CLSID into `{D1-D2-D3-D4hi-D4lo}` (Data1/2/3 LE, Data4 BE).
fn format_guid16(g: [u8; 16]) -> String {
    let d1 = u32::from_le_bytes([g[0], g[1], g[2], g[3]]);
    let d2 = u16::from_le_bytes([g[4], g[5]]);
    let d3 = u16::from_le_bytes([g[6], g[7]]);
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
        d1, d2, d3, g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15]
    )
}

#[cfg(test)]
mod tests {
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
}
