//! The Windows registry side of finding a type library by its GUID.

use super::locate::{
    language_candidates, parse_version_key, select_version, version_key, RegisteredVersion,
};
use windows_registry::{Key, CLASSES_ROOT};

const KEY_WOW64_64KEY: u32 = 0x0100;
const KEY_WOW64_32KEY: u32 = 0x0200;

/// The registry value naming the file of the registered library `guid` at
/// `major.minor` in language `lcid`, looking in the native and the 32-bit view
/// of the registry (VB6 libraries are 32-bit).
pub fn registered_value(guid: &str, major: u16, minor: u16, lcid: u32) -> Option<String> {
    let guid = guid.trim().trim_matches(|c| c == '{' || c == '}');
    [0u32, KEY_WOW64_32KEY, KEY_WOW64_64KEY]
        .into_iter()
        .find_map(|view| value_in_view(guid, major, minor, lcid, view))
}

fn open(view: u32, path: &str) -> Option<Key> {
    CLASSES_ROOT.options().read().access(view).open(path).ok()
}

fn names(key: &Key) -> Vec<String> {
    key.keys().map(|keys| keys.collect()).unwrap_or_default()
}

fn value_in_view(guid: &str, major: u16, minor: u16, lcid: u32, view: u32) -> Option<String> {
    let library = open(view, &format!("TypeLib\\{{{guid}}}"))?;
    let registered: Vec<RegisteredVersion> = names(&library).iter().filter_map(|k| parse_version_key(k)).collect();
    let version = select_version(&registered, major, minor)?;
    let version_path = format!("TypeLib\\{{{guid}}}\\{}", version_key(version));
    let version_key = open(view, &version_path)?;
    let languages = names(&version_key);
    for language in language_candidates(lcid, &languages) {
        let language_path = format!("{version_path}\\{language}");
        let Some(language_key) = open(view, &language_path) else { continue };
        let platforms = names(&language_key);
        let mut order: Vec<&str> = Vec::new();
        for wanted in ["win32", "win64", "win16"] {
            if let Some(found) = platforms.iter().find(|p| p.eq_ignore_ascii_case(wanted)) {
                order.push(found);
            }
        }
        for platform in order {
            if let Some(value) = open(view, &format!("{language_path}\\{platform}")).and_then(|k| k.get_string("").ok()) {
                if !value.trim().is_empty() {
                    return Some(value);
                }
            }
        }
        if let Ok(value) = language_key.get_string("") {
            if !value.trim().is_empty() {
                return Some(value);
            }
        }
    }
    None
}
