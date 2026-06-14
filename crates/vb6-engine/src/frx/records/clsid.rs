use crate::frx::{FrxError, FrxReader};

/// A 128-bit COM CLSID / IID (stored as four fields, all LE in FRX).
///
/// On-disk layout (16 bytes):
/// ```text
/// data1: u32 LE
/// data2: u16 LE
/// data3: u16 LE
/// data4: [u8; 8]
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Clsid {
    pub data1: u32,
    pub data2: u16,
    pub data3: u16,
    pub data4: [u8; 8],
}

impl Clsid {
    /// Read a 16-byte CLSID from the stream.
    pub fn read(r: &mut FrxReader<'_>) -> Result<Self, FrxError> {
        let data1 = r.read_u32_le()?;
        let data2 = r.read_u16_le()?;
        let data3 = r.read_u16_le()?;
        let bytes = r.read_bytes(8)?;
        let mut data4 = [0u8; 8];
        data4.copy_from_slice(bytes);
        Ok(Clsid { data1, data2, data3, data4 })
    }

    /// Write a 16-byte CLSID to a byte buffer.
    pub fn write(&self, out: &mut Vec<u8>) {
        out.extend_from_slice(&self.data1.to_le_bytes());
        out.extend_from_slice(&self.data2.to_le_bytes());
        out.extend_from_slice(&self.data3.to_le_bytes());
        out.extend_from_slice(&self.data4);
    }

    /// Format as `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.
    pub fn to_braced_string(&self) -> String {
        format!(
            "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            self.data1, self.data2, self.data3,
            self.data4[0], self.data4[1],
            self.data4[2], self.data4[3], self.data4[4],
            self.data4[5], self.data4[6], self.data4[7],
        )
    }

    /// IPersistStream  {00000109-0000-0000-C000-000000000046}
    pub const IID_IPERSIST_STREAM: Clsid = Clsid {
        data1: 0x00000109, data2: 0x0000, data3: 0x0000,
        data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
    };
    /// IPersistStreamInit  {7FD52380-4E07-101B-AE2D-08002B2EC713}
    pub const IID_IPERSIST_STREAM_INIT: Clsid = Clsid {
        data1: 0x7FD52380, data2: 0x4E07, data3: 0x101B,
        data4: [0xAE, 0x2D, 0x08, 0x00, 0x2B, 0x2E, 0xC7, 0x13],
    };
    /// IPersistStorage  {0000010A-0000-0000-C000-000000000046}
    pub const IID_IPERSIST_STORAGE: Clsid = Clsid {
        data1: 0x0000010A, data2: 0x0000, data3: 0x0000,
        data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
    };
    /// IPersistPropertyBag  {37D84F60-42CB-11CE-8135-00AA004BB851}
    pub const IID_IPERSIST_PROPERTY_BAG: Clsid = Clsid {
        data1: 0x37D84F60, data2: 0x42CB, data3: 0x11CE,
        data4: [0x81, 0x35, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51],
    };
    /// IPicture  {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    pub const IID_IPICTURE: Clsid = Clsid {
        data1: 0x7BF80980, data2: 0xBF32, data3: 0x101A,
        data4: [0x8B, 0xBB, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB],
    };
    /// IFont  {BEF6E002-A874-101A-8BBA-00AA00300CAB}
    pub const IID_IFONT: Clsid = Clsid {
        data1: 0xBEF6E002, data2: 0xA874, data3: 0x101A,
        data4: [0x8B, 0xBA, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB],
    };
}
