use crate::frx::{FrxError, FrxReader};
use super::clsid::Clsid;

/// Design-time license key for an ActiveX control.
///
/// VB6 calls `IClassFactory2::CreateInstanceLic` with this key at design time
/// when the control requires licensing (error string: "In order to use '|1',
/// you must specify a license string…").  At runtime, `IClassFactory2::CreateInstance`
/// is used without a key.
///
/// On-disk layout:
/// ```text
/// key_len: u32 LE           — byte length of the license string
/// key:     [u8; key_len]    — ANSI/wide license key bytes
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LicenseKey {
    pub key: Vec<u8>,
}

impl LicenseKey {
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let key = r.read_len_prefixed_bytes()?.to_vec();
        Ok(LicenseKey { key })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        let len = self.key.len() as u32;
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(&self.key);
    }
}

/// Control creation metadata — identifies a control and its license.
///
/// This record groups the control's CLSID, ProgID, and optional design-time license key
/// so the container can instantiate the control before loading its persisted state.
///
/// On-disk layout:
/// ```text
/// clsid:          [u8; 16]  — control's CLSID
/// progid_len:     u32 LE    — byte length of ProgID string (0 if unknown)
/// progid:         [u8; progid_len]  — ANSI ProgID (e.g. b"MSComCtl.Slider.2")
/// has_license:    u8        — 0=no license required, 1=license key follows
/// license:        LicenseKey   (present iff has_license == 1)
/// ```
#[derive(Debug)]
pub struct ControlCreateRecord {
    pub clsid: Clsid,
    /// ProgID bytes (ANSI).  May be empty if the CLSID alone is sufficient.
    pub prog_id: Vec<u8>,
    /// Design-time license key, or `None` for unlicensed controls.
    pub license: Option<LicenseKey>,
}

impl ControlCreateRecord {
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let at = r.pos();
        let clsid   = Clsid::read(r)?;
        let prog_id = r.read_len_prefixed_bytes()?.to_vec();
        let has_license = r.read_u8()?;
        let license = match has_license {
            0 => None,
            1 => Some(LicenseKey::read(r)?),
            other => return Err(FrxError::BadMagic { pos: at, expected: 0, got: other as u16 }),
        };
        Ok(ControlCreateRecord { clsid, prog_id, license })
    }

    pub fn write(&self, out: &mut Vec<u8>) {
        self.clsid.write(out);
        let pid_len = self.prog_id.len() as u32;
        out.extend_from_slice(&pid_len.to_le_bytes());
        out.extend_from_slice(&self.prog_id);
        match &self.license {
            None => out.push(0),
            Some(lic) => {
                out.push(1);
                lic.write(out);
            }
        }
    }
}
