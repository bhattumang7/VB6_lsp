//! Finding the type library inside a file: a `.tlb`/`.olb` is one, and a DLL, OCX or EXE carries
//! one or more as `TYPELIB` resources.

const MSFT_MAGIC: &[u8; 4] = b"MSFT";

/// Why the type library of a file could not be located.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LocateError {
    /// The file is neither a type library nor a PE image.
    NotAnImage,
    /// The PE image is damaged: a table points outside the file.
    Damaged(&'static str),
    /// The image has no `TYPELIB` resource, or none with the requested index.
    NoTypeLibrary,
}

impl std::fmt::Display for LocateError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            LocateError::NotAnImage => write!(f, "not a type library or an executable image"),
            LocateError::Damaged(what) => write!(f, "damaged executable image ({what})"),
            LocateError::NoTypeLibrary => write!(f, "no type library resource"),
        }
    }
}

fn u16_at(data: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes(data.get(at..at.checked_add(2)?)?.try_into().ok()?))
}

fn u32_at(data: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes(data.get(at..at.checked_add(4)?)?.try_into().ok()?))
}

struct Section {
    virtual_address: u32,
    virtual_size: u32,
    raw_offset: u32,
    raw_size: u32,
}

/// The bytes of the type library in `data`. For a PE image, `index` selects the resource
/// (`file.dll\2` is index 2; `None` and 0 mean the first one).
pub fn type_library_bytes(data: &[u8], index: Option<u32>) -> Result<&[u8], LocateError> {
    if data.get(..4) == Some(MSFT_MAGIC) {
        return Ok(data);
    }
    if data.get(..2) != Some(b"MZ") {
        return Err(LocateError::NotAnImage);
    }
    let pe = u32_at(data, 0x3C).ok_or(LocateError::Damaged("DOS header"))? as usize;
    if data.get(pe..pe + 4) != Some(b"PE\0\0") {
        return Err(LocateError::NotAnImage);
    }
    let coff = pe + 4;
    let section_count = u16_at(data, coff + 2).ok_or(LocateError::Damaged("file header"))? as usize;
    let optional_size = u16_at(data, coff + 16).ok_or(LocateError::Damaged("file header"))? as usize;
    let optional = coff + 20;
    let directories = match u16_at(data, optional) {
        Some(0x10b) => optional + 96,
        Some(0x20b) => optional + 112,
        _ => return Err(LocateError::Damaged("optional header")),
    };
    let resource_rva = u32_at(data, directories + 2 * 8).ok_or(LocateError::Damaged("data directories"))?;
    if resource_rva == 0 {
        return Err(LocateError::NoTypeLibrary);
    }
    let table = optional + optional_size;
    let mut sections = Vec::with_capacity(section_count);
    for i in 0..section_count {
        let at = table + i * 40;
        sections.push(Section {
            virtual_size: u32_at(data, at + 8).ok_or(LocateError::Damaged("section table"))?,
            virtual_address: u32_at(data, at + 12).ok_or(LocateError::Damaged("section table"))?,
            raw_size: u32_at(data, at + 16).ok_or(LocateError::Damaged("section table"))?,
            raw_offset: u32_at(data, at + 20).ok_or(LocateError::Damaged("section table"))?,
        });
    }
    let file_offset = |rva: u32| -> Option<usize> {
        sections.iter().find_map(|s| {
            let span = s.virtual_size.max(s.raw_size);
            (rva >= s.virtual_address && rva - s.virtual_address < span)
                .then(|| s.raw_offset as usize + (rva - s.virtual_address) as usize)
        })
    };
    let root = file_offset(resource_rva).ok_or(LocateError::Damaged("resource directory"))?;

    // Level 1: resource type. `TYPELIB` is named.
    let type_entry = directory_entries(data, root, root)?
        .into_iter()
        .find(|e| e.name.as_deref() == Some("TYPELIB"))
        .ok_or(LocateError::NoTypeLibrary)?;
    // Level 2: resource name or id; level 3: language.
    let names = directory_entries(data, root, root + type_entry.target as usize)?;
    let wanted = index.filter(|&i| i != 0).unwrap_or(1);
    let chosen = names
        .iter()
        .find(|e| e.name.is_none() && e.id == wanted)
        .or(if index.is_none() { names.first() } else { None })
        .ok_or(LocateError::NoTypeLibrary)?;
    let languages = directory_entries(data, root, root + chosen.target as usize)?;
    let leaf = languages.first().ok_or(LocateError::NoTypeLibrary)?;
    let entry = root + leaf.target as usize;
    let rva = u32_at(data, entry).ok_or(LocateError::Damaged("resource data entry"))?;
    let size = u32_at(data, entry + 4).ok_or(LocateError::Damaged("resource data entry"))? as usize;
    let start = file_offset(rva).ok_or(LocateError::Damaged("resource data"))?;
    data.get(start..start.checked_add(size).ok_or(LocateError::Damaged("resource data"))?)
        .ok_or(LocateError::Damaged("resource data"))
}

struct Entry {
    id: u32,
    name: Option<String>,
    /// Offset from the start of the resource section: a subdirectory, or (for a leaf) a data entry.
    target: u32,
}

/// The entries of the resource directory at `at`; `root` is the start of the resource section.
fn directory_entries(data: &[u8], root: usize, at: usize) -> Result<Vec<Entry>, LocateError> {
    let named = u16_at(data, at + 12).ok_or(LocateError::Damaged("resource directory"))? as usize;
    let ids = u16_at(data, at + 14).ok_or(LocateError::Damaged("resource directory"))? as usize;
    let mut out = Vec::with_capacity(named + ids);
    for i in 0..named + ids {
        let entry = at + 16 + i * 8;
        let key = u32_at(data, entry).ok_or(LocateError::Damaged("resource entry"))?;
        let value = u32_at(data, entry + 4).ok_or(LocateError::Damaged("resource entry"))?;
        let name = if key & 0x8000_0000 != 0 {
            let at = root + (key & 0x7FFF_FFFF) as usize;
            let length = u16_at(data, at).ok_or(LocateError::Damaged("resource name"))? as usize;
            let units: Vec<u16> = (0..length)
                .map(|n| u16_at(data, at + 2 + n * 2))
                .collect::<Option<_>>()
                .ok_or(LocateError::Damaged("resource name"))?;
            Some(String::from_utf16_lossy(&units))
        } else {
            None
        };
        out.push(Entry { id: key & 0x7FFF_FFFF, name, target: value & 0x7FFF_FFFF });
    }
    Ok(out)
}
