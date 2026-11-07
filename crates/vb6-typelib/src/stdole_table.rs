//! The types of the OLE automation library (`stdole`) as other libraries import them, for when
//! the library file itself is not at hand. An alias is listed as the type it stands for.

use super::imports::ImportedType;

/// (index in the library, GUID, library name (empty for a plain VB6 type), name, is an object).
pub const STDOLE_TYPES: &[(usize, Option<&str>, &str, &str, bool)] = &[
    (0, None, "stdole", "GUID", false),
    (1, None, "stdole", "DISPPARAMS", false),
    (2, None, "stdole", "EXCEPINFO", false),
    (3, Some("00000000-0000-0000-c000-000000000046"), "stdole", "IUnknown", true),
    (4, Some("00020400-0000-0000-c000-000000000046"), "stdole", "IDispatch", true),
    (5, Some("00020404-0000-0000-c000-000000000046"), "stdole", "IEnumVARIANT", true),
    (6, Some("66504301-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (7, Some("66504302-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (8, Some("66504303-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (9, Some("66504304-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (10, Some("66504305-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (11, Some("66504306-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (12, Some("66504307-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (13, Some("66504308-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (14, Some("66504309-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (15, Some("bf030640-9069-101b-ae2d-08002b2ec713"), "", "Single", false),
    (16, Some("bf030641-9069-101b-ae2d-08002b2ec713"), "", "Single", false),
    (17, Some("bf030642-9069-101b-ae2d-08002b2ec713"), "", "Single", false),
    (18, Some("bf030643-9069-101b-ae2d-08002b2ec713"), "", "Single", false),
    (19, Some("66504313-be0f-101a-8bbb-00aa00300cab"), "", "Long", false),
    (20, Some("6650430b-be0f-101a-8bbb-00aa00300cab"), "", "Boolean", false),
    (21, Some("bf030644-9069-101b-ae2d-08002b2ec713"), "", "Boolean", false),
    (22, Some("bf030645-9069-101b-ae2d-08002b2ec713"), "", "Boolean", false),
    (23, Some("6650430a-be0f-101a-8bbb-00aa00300cab"), "stdole", "OLE_TRISTATE", false),
    (24, Some("6650430d-be0f-101a-8bbb-00aa00300cab"), "", "String", false),
    (25, Some("6650430e-be0f-101a-8bbb-00aa00300cab"), "", "Currency", false),
    (26, Some("6650430f-be0f-101a-8bbb-00aa00300cab"), "", "Boolean", false),
    (27, Some("66504310-be0f-101a-8bbb-00aa00300cab"), "", "Boolean", false),
    (28, Some("66504311-be0f-101a-8bbb-00aa00300cab"), "", "Boolean", false),
    (29, Some("66504312-be0f-101a-8bbb-00aa00300cab"), "", "Boolean", false),
    (30, Some("bef6e002-a874-101a-8bba-00aa00300cab"), "stdole", "IFont", true),
    (31, Some("bef6e003-a874-101a-8bba-00aa00300cab"), "stdole", "Font", true),
    (32, None, "stdole", "Font", true),
    (33, Some("0be35203-8f91-11ce-9de3-00aa004bb851"), "stdole", "StdFont", true),
    (34, Some("7bf80980-bf32-101a-8bbb-00aa00300cab"), "stdole", "IPicture", true),
    (35, Some("7bf80981-bf32-101a-8bbb-00aa00300cab"), "stdole", "Picture", true),
    (36, None, "stdole", "Picture", true),
    (37, Some("0be35204-8f91-11ce-9de3-00aa004bb851"), "stdole", "StdPicture", true),
    (38, Some("e6c8fa08-bd9f-11d0-985e-00c04fc29993"), "stdole", "LoadPictureConstants", false),
    (39, Some("91209ac0-60f6-11cf-9c5d-00aa00c1489e"), "stdole", "StdFunctions", false),
    (40, Some("4ef6100a-af88-11d0-9846-00c04fc29993"), "stdole", "FontEvents", true),
    (41, None, "stdole", "FontEvents", true),
];

/// The type of `stdole` with the given GUID or index.
pub fn stdole_type(type_guid: Option<&str>, type_index: Option<usize>) -> Option<ImportedType> {
    STDOLE_TYPES
        .iter()
        .find(|(index, guid, ..)| match (type_guid, type_index) {
            (Some(wanted), _) => *guid == Some(wanted),
            (None, Some(wanted)) => *index == wanted,
            (None, None) => false,
        })
        .map(|(_, _, library, name, object)| ImportedType {
            library: (*library).to_string(),
            name: (*name).to_string(),
            object: *object,
        })
}
