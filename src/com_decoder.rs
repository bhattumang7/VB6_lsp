//! Tier-3 live COM decoder (Windows only).
//!
//! Implements [`crate::controls::resources::ComDecoder`] by hosting the control
//! out-of-process in 32-bit PowerShell + COM (`scripts/com_bag_decode.ps1` +
//! `scripts/ComBag.cs`). The control is created license-aware; when it exists but
//! its license is absent the bridge reports `NOTLICENSED` and we return a HARD
//! [`ResolveError::LicensedBagUnavailable`] — never a silent opaque fallback.
//!
//! This lives in the binary (not the library) so the core crate stays platform
//! independent: the `ComDecoder` trait is the seam, this is the Windows impl.
#![cfg(windows)]

use std::io::Write;
use std::path::PathBuf;
use std::process::Command;

use crate::controls::frx::FrxValue;
use crate::controls::resources::{ComDecoder, ResolveError};

/// Decodes OCX bags by driving the control via 32-bit PowerShell + COM.
pub struct OracleComDecoder {
    ps32: PathBuf,
    script: PathBuf,
}

impl OracleComDecoder {
    /// Locate 32-bit PowerShell and the bundled bridge script. Returns `None`
    /// when either is unavailable (caller then keeps the opaque policy).
    pub fn new() -> Option<Self> {
        let sysroot = std::env::var("SystemRoot").unwrap_or_else(|_| "C:\\Windows".to_string());
        let ps32 = PathBuf::from(format!(
            "{}\\SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe",
            sysroot
        ));
        if !ps32.exists() {
            return None;
        }
        let script = locate_script()?;
        Some(OracleComDecoder { ps32, script })
    }
}

/// Find `com_bag_decode.ps1`: explicit override, then next to the exe, then the
/// source tree (dev runs).
fn locate_script() -> Option<PathBuf> {
    if let Ok(p) = std::env::var("VB6_COM_BRIDGE") {
        let p = PathBuf::from(p);
        if p.exists() {
            return Some(p);
        }
    }
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let p = dir.join("scripts").join("com_bag_decode.ps1");
            if p.exists() {
                return Some(p);
            }
        }
    }
    let p = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("scripts")
        .join("com_bag_decode.ps1");
    if p.exists() {
        return Some(p);
    }
    None
}

impl ComDecoder for OracleComDecoder {
    fn decode_bag(
        &self,
        clsid: Option<&str>,
        property: &str,
        bag: &[u8],
    ) -> Result<FrxValue, ResolveError> {
        let clsid = clsid.ok_or_else(|| ResolveError::ComDecodeFailed {
            property: property.to_string(),
            message: "no CLSID available for the control".to_string(),
        })?;

        let tmp = std::env::temp_dir().join(format!(
            "vb6bag_{}_{}.bin",
            std::process::id(),
            property.replace(|c: char| !c.is_ascii_alphanumeric(), "_")
        ));
        std::fs::File::create(&tmp)
            .and_then(|mut f| f.write_all(bag))
            .map_err(|e| ResolveError::ComDecodeFailed {
                property: property.to_string(),
                message: e.to_string(),
            })?;

        let output = Command::new(&self.ps32)
            .args(["-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File"])
            .arg(&self.script)
            .arg(clsid)
            .arg(&tmp)
            .output();
        let _ = std::fs::remove_file(&tmp);

        let output = output.map_err(|e| ResolveError::ComDecodeFailed {
            property: property.to_string(),
            message: format!("spawn bridge: {}", e),
        })?;
        let stdout = String::from_utf8_lossy(&output.stdout);
        let line = stdout
            .lines()
            .rev()
            .find(|l| l.trim_start().starts_with('{'))
            .unwrap_or("")
            .trim();
        let v: serde_json::Value =
            serde_json::from_str(line).map_err(|e| ResolveError::ComDecodeFailed {
                property: property.to_string(),
                message: format!("bad bridge output ({}): {}", e, line),
            })?;

        if v["ok"].as_bool() == Some(true) {
            let properties = v["properties"]
                .as_array()
                .map(|a| {
                    a.iter()
                        .filter_map(|p| {
                            let pair = p.as_array()?;
                            Some((
                                pair.first()?.as_str()?.to_string(),
                                pair.get(1)?.as_str()?.to_string(),
                            ))
                        })
                        .collect()
                })
                .unwrap_or_default();
            Ok(FrxValue::DecodedBag {
                clsid: parse_guid(clsid),
                properties,
            })
        } else {
            match v["error"].as_str().unwrap_or("unknown") {
                "NOTLICENSED" => Err(ResolveError::LicensedBagUnavailable {
                    clsid: Some(clsid.to_string()),
                    property: property.to_string(),
                }),
                other => Err(ResolveError::ComDecodeFailed {
                    property: property.to_string(),
                    message: other.to_string(),
                }),
            }
        }
    }
}

/// Parse `{D1-D2-D3-D4hi-D4lo}` into the 16-byte layout (Data1/2/3 LE, Data4 BE).
fn parse_guid(s: &str) -> Option<[u8; 16]> {
    let t = s.trim().trim_start_matches('{').trim_end_matches('}');
    let parts: Vec<&str> = t.split('-').collect();
    if parts.len() != 5 {
        return None;
    }
    let d1 = u32::from_str_radix(parts[0], 16).ok()?;
    let d2 = u16::from_str_radix(parts[1], 16).ok()?;
    let d3 = u16::from_str_radix(parts[2], 16).ok()?;
    let d4a = parse_hex_bytes(parts[3])?; // 2 bytes
    let d4b = parse_hex_bytes(parts[4])?; // 6 bytes
    if d4a.len() != 2 || d4b.len() != 6 {
        return None;
    }
    let mut g = [0u8; 16];
    g[0..4].copy_from_slice(&d1.to_le_bytes());
    g[4..6].copy_from_slice(&d2.to_le_bytes());
    g[6..8].copy_from_slice(&d3.to_le_bytes());
    g[8..10].copy_from_slice(&d4a);
    g[10..16].copy_from_slice(&d4b);
    Some(g)
}

fn parse_hex_bytes(s: &str) -> Option<Vec<u8>> {
    if s.len() % 2 != 0 {
        return None;
    }
    (0..s.len())
        .step_by(2)
        .map(|i| u8::from_str_radix(&s[i..i + 2], 16).ok())
        .collect()
}
