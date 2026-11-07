//! Readers and writers for the VB6 file formats.
//!
//! * [`vbp`] / [`vbg`] — project and project-group files.
//! * [`frm`] — the designer section of `.frm`/`.ctl`/`.pag`/`.dob` files.
//! * [`frx`] — the binary companion files (`.frx`/`.ctx`) the designer section points into.
//! * [`controls`] — the resources a form's designer section references, decoded.
//! * [`res`] / [`res_json`] — compiled resource (`.res`) files.
//! * [`encoding`] — reading text files in their legacy code page.
//!
//! Nothing here depends on the language server or the analysis engine.

pub mod controls;
pub mod encoding;
pub mod frm;
pub mod frx;
pub mod res;
pub mod res_json;
pub mod vbg;
pub mod vbp;
