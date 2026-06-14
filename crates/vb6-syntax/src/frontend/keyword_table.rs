//! VB6 keyword/operator table — the behavioral ground truth for keyword
//! interning. Generated table; do not hand-edit.
//!
//! Combines the two parallel 0x10F-entry keyword tables, walked in lockstep
//! during keyword interning:
//!   * the keyword name strings.
//!   * three dwords `{w0, w1, w2}` per entry.
//!
//! Field meaning (as used by the engine):
//!   * `w0` low 16 bits = the token id, which always equals the array index;
//!     the high 16 bits are scanner attributes read by other functions.
//!     `intern_keywords` uses `w0 & 0xffff`.
//!   * `w1` = flags; `intern_keywords` reads only bit 4 (`w1 >> 4 & 1`, the
//!     "reserved word" marker).
//!   * `w2` = auxiliary value (help/dispid); unused by `intern_keywords`.
//!
//! All three dwords are preserved verbatim so no information is lost.

/// One keyword/operator entry: name plus the three raw metadata dwords.
#[derive(Clone, Copy, Debug)]
pub struct KeywordEntry {
    /// Keyword text exactly as stored by VB6 (ASCII).
    pub name: &'static str,
    /// dword 0: token id in low 16 bits (== index), attrs in high 16.
    pub w0: u32,
    /// dword 1: flags. Bit 4 (`0x10`) = reserved-word marker.
    pub w1: u32,
    /// dword 2: auxiliary value; unused by `intern_keywords`.
    pub w2: u32,
}

impl KeywordEntry {
    /// The token id (`w0 & 0xffff`) — equals this entry's index in the table.
    #[inline]
    pub const fn token(&self) -> u16 {
        (self.w0 & 0xffff) as u16
    }
}

/// The keyword/operator name string for token id `token`
/// (`KEYWORD_TABLE[token].name`); the keyword string table is indexed by token
/// id. Returns `""` for an out-of-range id.
pub fn keyword_string(token: u16) -> &'static str {
    KEYWORD_TABLE.get(token as usize).map(|e| e.name).unwrap_or("")
}

/// The 271 keyword/operator entries, in VB6 table order (index == token).
pub static KEYWORD_TABLE: [KeywordEntry; 271] = [
    KeywordEntry { name: "0", w0: 0x00000000, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "Abs", w0: 0x00000001, w1: 0x00000050, w2: 0x000f64d2 },
    KeywordEntry { name: "Access", w0: 0x00000002, w1: 0x00000000, w2: 0x000f64a0 },
    KeywordEntry { name: "AddressOf", w0: 0x00000003, w1: 0x00000010, w2: 0x0010d741 },
    KeywordEntry { name: "Alias", w0: 0x00000004, w1: 0x00000000, w2: 0x000f648c },
    KeywordEntry { name: "And", w0: 0x00230005, w1: 0x00000015, w2: 0x000f64d4 },
    KeywordEntry { name: "Any", w0: 0x00000006, w1: 0x00000010, w2: 0x000f64d6 },
    KeywordEntry { name: "Append", w0: 0x00000007, w1: 0x00000000, w2: 0x000f64a1 },
    KeywordEntry { name: "Array", w0: 0x00000008, w1: 0x00000010, w2: 0x000f7386 },
    KeywordEntry { name: "As", w0: 0x00000009, w1: 0x00000010, w2: 0x000f64d9 },
    KeywordEntry { name: "Assert", w0: 0x0000000a, w1: 0x00000000, w2: 0x0010d742 },
    KeywordEntry { name: "B", w0: 0x0000000b, w1: 0x00000000, w2: 0x000f6e8e },
    KeywordEntry { name: "Base", w0: 0x0000000c, w1: 0x00000000, w2: 0x000f64a4 },
    KeywordEntry { name: "BF", w0: 0x0000000d, w1: 0x00000000, w2: 0x000f6e8f },
    KeywordEntry { name: "Binary", w0: 0x0000000e, w1: 0x00000000, w2: 0x000f6d10 },
    KeywordEntry { name: "Boolean", w0: 0x0000000f, w1: 0x00001810, w2: 0x000f667e },
    KeywordEntry { name: "ByRef", w0: 0x00000010, w1: 0x00000010, w2: 0x000f64ac },
    KeywordEntry { name: "Byte", w0: 0x00000011, w1: 0x00002810, w2: 0x000f7e40 },
    KeywordEntry { name: "ByVal", w0: 0x00000012, w1: 0x00000010, w2: 0x000f64de },
    KeywordEntry { name: "Call", w0: 0x00000013, w1: 0x00000210, w2: 0x000f64df },
    KeywordEntry { name: "Case", w0: 0x00000014, w1: 0x00000210, w2: 0x000f64aa },
    KeywordEntry { name: "CBool", w0: 0x00000015, w1: 0x00000050, w2: 0x000f64b4 },
    KeywordEntry { name: "CByte", w0: 0x00000016, w1: 0x00000050, w2: 0x000f7e5b },
    KeywordEntry { name: "CCur", w0: 0x00000017, w1: 0x00000050, w2: 0x000f64c4 },
    KeywordEntry { name: "CDate", w0: 0x00000018, w1: 0x00000050, w2: 0x000f64f3 },
    KeywordEntry { name: "CDec", w0: 0x00000019, w1: 0x00000010, w2: 0x0010f771 },
    KeywordEntry { name: "CDbl", w0: 0x0000001a, w1: 0x00000050, w2: 0x000f64c5 },
    KeywordEntry { name: "CDecl", w0: 0x0000001b, w1: 0x00000010, w2: 0x000f6eac },
    KeywordEntry { name: "ChDir", w0: 0x0000001c, w1: 0x00000000, w2: 0x000f7386 },
    KeywordEntry { name: "CInt", w0: 0x0000001d, w1: 0x00000050, w2: 0x000f64c7 },
    KeywordEntry { name: "Circle", w0: 0x0000001e, w1: 0x00000310, w2: 0x000f7386 },
    KeywordEntry { name: "CLng", w0: 0x0000001f, w1: 0x00000050, w2: 0x000f64c8 },
    KeywordEntry { name: "Close", w0: 0x00000020, w1: 0x00000210, w2: 0x000f64e8 },
    KeywordEntry { name: "Compare", w0: 0x00000021, w1: 0x00000000, w2: 0x000f64a6 },
    KeywordEntry { name: "Const", w0: 0x00000022, w1: 0x00000210, w2: 0x000f64ed },
    KeywordEntry { name: "CSng", w0: 0x00000023, w1: 0x00000050, w2: 0x000f64c9 },
    KeywordEntry { name: "CStr", w0: 0x00000024, w1: 0x00000050, w2: 0x000f6d56 },
    KeywordEntry { name: "CurDir", w0: 0x00000025, w1: 0x00000000, w2: 0x000f64f1 },
    KeywordEntry { name: "CurDir$", w0: 0x00000026, w1: 0x00000000, w2: 0x000f6e91 },
    KeywordEntry { name: "CVar", w0: 0x00000027, w1: 0x00000050, w2: 0x000f64cb },
    KeywordEntry { name: "CVDate", w0: 0x00000028, w1: 0x00000000, w2: 0x000f6e7a },
    KeywordEntry { name: "CVErr", w0: 0x00000029, w1: 0x00000050, w2: 0x000f64b5 },
    KeywordEntry { name: "Currency", w0: 0x0000002a, w1: 0x00006810, w2: 0x000f64f2 },
    KeywordEntry { name: "Database", w0: 0x0000002b, w1: 0x00000000, w2: 0x000f64a6 },
    KeywordEntry { name: "Date", w0: 0x0000002c, w1: 0x000062d0, w2: 0x000f7377 },
    KeywordEntry { name: "Date$", w0: 0x0000002d, w1: 0x000002d0, w2: 0x000f7386 },
    KeywordEntry { name: "Debug", w0: 0x0000002e, w1: 0x00000210, w2: 0x000f64fb },
    KeywordEntry { name: "Decimal", w0: 0x0000002f, w1: 0x00000010, w2: 0x0010c85c },
    KeywordEntry { name: "Declare", w0: 0x00000030, w1: 0x00000230, w2: 0x000f64fe },
    KeywordEntry { name: "DefBool", w0: 0x00000031, w1: 0x00000210, w2: 0x000f6ca4 },
    KeywordEntry { name: "DefByte", w0: 0x00000032, w1: 0x00000210, w2: 0x000f6490 },
    KeywordEntry { name: "DefCur", w0: 0x00000033, w1: 0x00000210, w2: 0x000f648e },
    KeywordEntry { name: "DefDate", w0: 0x00000034, w1: 0x00000210, w2: 0x000f6ca5 },
    KeywordEntry { name: "DefDec", w0: 0x00000035, w1: 0x00000010, w2: 0x0010d336 },
    KeywordEntry { name: "DefDbl", w0: 0x00000036, w1: 0x00000210, w2: 0x000f648f },
    KeywordEntry { name: "DefInt", w0: 0x00000037, w1: 0x00000210, w2: 0x000f6490 },
    KeywordEntry { name: "DefLng", w0: 0x00000038, w1: 0x00000210, w2: 0x000f6491 },
    KeywordEntry { name: "DefObj", w0: 0x00000039, w1: 0x00000210, w2: 0x000f6eca },
    KeywordEntry { name: "DefSng", w0: 0x0000003a, w1: 0x00000210, w2: 0x000f6492 },
    KeywordEntry { name: "DefStr", w0: 0x0000003b, w1: 0x00000210, w2: 0x000f6493 },
    KeywordEntry { name: "DefVar", w0: 0x0000003c, w1: 0x00000210, w2: 0x000f6494 },
    KeywordEntry { name: "Dim", w0: 0x0000003d, w1: 0x00000210, w2: 0x000f6501 },
    KeywordEntry { name: "Dir", w0: 0x0000003e, w1: 0x00000000, w2: 0x000f6502 },
    KeywordEntry { name: "Dir$", w0: 0x0000003f, w1: 0x00000000, w2: 0x000f6e93 },
    KeywordEntry { name: "Do", w0: 0x00000040, w1: 0x00000210, w2: 0x000f6503 },
    KeywordEntry { name: "DoEvents", w0: 0x00000041, w1: 0x00000010, w2: 0x000f7386 },
    KeywordEntry { name: "Double", w0: 0x00000042, w1: 0x00005810, w2: 0x000f6505 },
    KeywordEntry { name: "Each", w0: 0x00000043, w1: 0x00000010, w2: 0x000f64b6 },
    KeywordEntry { name: "Else", w0: 0x00000044, w1: 0x00000230, w2: 0x000f6506 },
    KeywordEntry { name: "ElseIf", w0: 0x00000045, w1: 0x00000230, w2: 0x000f6499 },
    KeywordEntry { name: "Empty", w0: 0x00000046, w1: 0x00000010, w2: 0x000f6507 },
    KeywordEntry { name: "End", w0: 0x00000047, w1: 0x00000210, w2: 0x000f6508 },
    KeywordEntry { name: "EndIf", w0: 0x00000048, w1: 0x00000210, w2: 0x000f6ecb },
    KeywordEntry { name: "Enum", w0: 0x00000049, w1: 0x00000210, w2: 0x0010d69a },
    KeywordEntry { name: "Eqv", w0: 0x0020004a, w1: 0x00000012, w2: 0x000f650d },
    KeywordEntry { name: "Erase", w0: 0x0000004b, w1: 0x00000210, w2: 0x000f650e },
    KeywordEntry { name: "Error", w0: 0x0000004c, w1: 0x00000200, w2: 0x000f6c6d },
    KeywordEntry { name: "Error$", w0: 0x0000004d, w1: 0x00000000, w2: 0x000f6e94 },
    KeywordEntry { name: "Event", w0: 0x0000004e, w1: 0x00000230, w2: 0x0010d69b },
    KeywordEntry { name: "Exit", w0: 0x0000004f, w1: 0x00000210, w2: 0x000f6514 },
    KeywordEntry { name: "Explicit", w0: 0x00000050, w1: 0x00000000, w2: 0x000f64a8 },
    KeywordEntry { name: "F", w0: 0x00000051, w1: 0x00000000, w2: 0x000f6e90 },
    KeywordEntry { name: "False", w0: 0x00000052, w1: 0x00000010, w2: 0x000f6516 },
    KeywordEntry { name: "Fix", w0: 0x00000053, w1: 0x00000050, w2: 0x000f649b },
    KeywordEntry { name: "For", w0: 0x00000054, w1: 0x00000210, w2: 0x000f6d60 },
    KeywordEntry { name: "Format", w0: 0x00000055, w1: 0x00000000, w2: 0x000f651d },
    KeywordEntry { name: "Format$", w0: 0x00000056, w1: 0x00000000, w2: 0x000f6e95 },
    KeywordEntry { name: "FreeFile", w0: 0x00000057, w1: 0x00000000, w2: 0x000f651e },
    KeywordEntry { name: "Friend", w0: 0x00000058, w1: 0x00000210, w2: 0x0010d743 },
    KeywordEntry { name: "Function", w0: 0x00000059, w1: 0x00000230, w2: 0x000f651f },
    KeywordEntry { name: "Get", w0: 0x0000005a, w1: 0x00000210, w2: 0x000f6786 },
    KeywordEntry { name: "Global", w0: 0x0000005b, w1: 0x00000210, w2: 0x000f64bf },
    KeywordEntry { name: "Go", w0: 0x0000005c, w1: 0x00000200, w2: 0x000f6498 },
    KeywordEntry { name: "GoSub", w0: 0x0000005d, w1: 0x00000210, w2: 0x000f6526 },
    KeywordEntry { name: "GoTo", w0: 0x0000005e, w1: 0x00000210, w2: 0x000f6527 },
    KeywordEntry { name: "If", w0: 0x0000005f, w1: 0x00000210, w2: 0x000f652c },
    KeywordEntry { name: "Imp", w0: 0x001f0060, w1: 0x00000011, w2: 0x000f652d },
    KeywordEntry { name: "Implements", w0: 0x00000061, w1: 0x00000210, w2: 0x0010d69d },
    KeywordEntry { name: "In", w0: 0x00000062, w1: 0x00000010, w2: 0x000f667b },
    KeywordEntry { name: "Input", w0: 0x00000063, w1: 0x00000290, w2: 0x000f735f },
    KeywordEntry { name: "Input$", w0: 0x00000064, w1: 0x00000090, w2: 0x000f6e96 },
    KeywordEntry { name: "InputB", w0: 0x00000065, w1: 0x00000090, w2: 0x000f6d7a },
    KeywordEntry { name: "InputB$", w0: 0x00000066, w1: 0x00000090, w2: 0x000f6e97 },
    KeywordEntry { name: "InStr", w0: 0x00000067, w1: 0x00000080, w2: 0x000f6532 },
    KeywordEntry { name: "InStrB", w0: 0x00000068, w1: 0x00000080, w2: 0x000f6d7d },
    KeywordEntry { name: "Int", w0: 0x00000069, w1: 0x00000050, w2: 0x000f6533 },
    KeywordEntry { name: "Integer", w0: 0x0000006a, w1: 0x00003010, w2: 0x000f6534 },
    KeywordEntry { name: "Is", w0: 0x0036006b, w1: 0x00000016, w2: 0x000f6535 },
    KeywordEntry { name: "LBound", w0: 0x0000006c, w1: 0x00000090, w2: 0x000f653c },
    KeywordEntry { name: "Left", w0: 0x0000006d, w1: 0x00000000, w2: 0x000f7386 },
    KeywordEntry { name: "Len", w0: 0x0000006e, w1: 0x00000050, w2: 0x000f653f },
    KeywordEntry { name: "LenB", w0: 0x0000006f, w1: 0x00000050, w2: 0x000f6d79 },
    KeywordEntry { name: "Let", w0: 0x00000070, w1: 0x00000210, w2: 0x000f6785 },
    KeywordEntry { name: "Lib", w0: 0x00000071, w1: 0x00000000, w2: 0x000f648d },
    KeywordEntry { name: "Like", w0: 0x00250072, w1: 0x00000016, w2: 0x000f6541 },
    KeywordEntry { name: "Line", w0: 0x00000073, w1: 0x00000300, w2: 0x000f6542 },
    KeywordEntry { name: "LINEINPUT", w0: 0x00000074, w1: 0x00000310, w2: 0x000f7e21 },
    KeywordEntry { name: "Load", w0: 0x00000075, w1: 0x00000000, w2: 0x000f7386 },
    KeywordEntry { name: "Local", w0: 0x00000076, w1: 0x00000010, w2: 0x000f649f },
    KeywordEntry { name: "Lock", w0: 0x00000077, w1: 0x00000210, w2: 0x000f6e98 },
    KeywordEntry { name: "Long", w0: 0x00000078, w1: 0x00004010, w2: 0x000f6548 },
    KeywordEntry { name: "Loop", w0: 0x00000079, w1: 0x00000210, w2: 0x000f6495 },
    KeywordEntry { name: "LSet", w0: 0x0000007a, w1: 0x00000210, w2: 0x000f6549 },
    KeywordEntry { name: "Me", w0: 0x0000007b, w1: 0x00000010, w2: 0x000f64e4 },
    KeywordEntry { name: "Mid", w0: 0x0000007c, w1: 0x00000200, w2: 0x000f6e9a },
    KeywordEntry { name: "Mid$", w0: 0x0000007d, w1: 0x00000200, w2: 0x000f6e99 },
    KeywordEntry { name: "MidB", w0: 0x0000007e, w1: 0x00000200, w2: 0x000f6d7e },
    KeywordEntry { name: "MidB$", w0: 0x0000007f, w1: 0x00000200, w2: 0x000f6e9b },
    KeywordEntry { name: "Mod", w0: 0x001d0080, w1: 0x00000019, w2: 0x000f6550 },
    KeywordEntry { name: "Module", w0: 0x00000081, w1: 0x00000000, w2: 0x000f6d75 },
    KeywordEntry { name: "Name", w0: 0x00000082, w1: 0x00000200, w2: 0x000f6553 },
    KeywordEntry { name: "New", w0: 0x00000083, w1: 0x00000010, w2: 0x000f64e5 },
    KeywordEntry { name: "Next", w0: 0x00000084, w1: 0x00000210, w2: 0x000f6554 },
    KeywordEntry { name: "Not", w0: 0x00060085, w1: 0x00000015, w2: 0x000f6555 },
    KeywordEntry { name: "Nothing", w0: 0x00000086, w1: 0x00000010, w2: 0x000f6ecd },
    KeywordEntry { name: "Null", w0: 0x00000087, w1: 0x00000010, w2: 0x000f6c6c },
    KeywordEntry { name: "Object", w0: 0x00000088, w1: 0x0000b000, w2: 0x000f64bd },
    KeywordEntry { name: "On", w0: 0x00000089, w1: 0x00000210, w2: 0x000f6558 },
    KeywordEntry { name: "Open", w0: 0x0000008a, w1: 0x00000210, w2: 0x000f655b },
    KeywordEntry { name: "Option", w0: 0x0000008b, w1: 0x00000210, w2: 0x000f655d },
    KeywordEntry { name: "Optional", w0: 0x0000008c, w1: 0x00000010, w2: 0x000f6d13 },
    KeywordEntry { name: "Or", w0: 0x0021008d, w1: 0x00000014, w2: 0x000f6561 },
    KeywordEntry { name: "Output", w0: 0x0000008e, w1: 0x00000000, w2: 0x000f64a2 },
    KeywordEntry { name: "ParamArray", w0: 0x0000008f, w1: 0x00000010, w2: 0x000f6d76 },
    KeywordEntry { name: "Preserve", w0: 0x00000090, w1: 0x00000010, w2: 0x000f64a9 },
    KeywordEntry { name: "Print", w0: 0x00000091, w1: 0x00000310, w2: 0x000f6e9c },
    KeywordEntry { name: "Private", w0: 0x00000092, w1: 0x00000210, w2: 0x000f6564 },
    KeywordEntry { name: "Property", w0: 0x00000093, w1: 0x00000220, w2: 0x000f64be },
    KeywordEntry { name: "PSet", w0: 0x00000094, w1: 0x00000310, w2: 0x000f7386 },
    KeywordEntry { name: "Public", w0: 0x00000095, w1: 0x00000210, w2: 0x000f6781 },
    KeywordEntry { name: "Put", w0: 0x00000096, w1: 0x00000210, w2: 0x000f6565 },
    KeywordEntry { name: "RaiseEvent", w0: 0x00000097, w1: 0x00000210, w2: 0x0010d69c },
    KeywordEntry { name: "Random", w0: 0x00000098, w1: 0x00000000, w2: 0x000f64a3 },
    KeywordEntry { name: "Randomize", w0: 0x00000099, w1: 0x00000000, w2: 0x000f7386 },
    KeywordEntry { name: "Read", w0: 0x0000009a, w1: 0x00000000, w2: 0x000f6ecc },
    KeywordEntry { name: "ReDim", w0: 0x0000009b, w1: 0x00000210, w2: 0x000f6567 },
    KeywordEntry { name: "Rem", w0: 0x0000009c, w1: 0x00000230, w2: 0x000f6568 },
    KeywordEntry { name: "Resume", w0: 0x0000009d, w1: 0x00000210, w2: 0x000f6e9d },
    KeywordEntry { name: "Return", w0: 0x0000009e, w1: 0x00000210, w2: 0x000f6497 },
    KeywordEntry { name: "RGB", w0: 0x0000009f, w1: 0x00000000, w2: 0x000f7386 },
    KeywordEntry { name: "RSet", w0: 0x000000a0, w1: 0x00000210, w2: 0x000f6571 },
    KeywordEntry { name: "Scale", w0: 0x000000a1, w1: 0x00000310, w2: 0x000f7386 },
    KeywordEntry { name: "Seek", w0: 0x000000a2, w1: 0x00000210, w2: 0x000f6d63 },
    KeywordEntry { name: "Select", w0: 0x000000a3, w1: 0x00000210, w2: 0x000f6576 },
    KeywordEntry { name: "Set", w0: 0x000000a4, w1: 0x00000210, w2: 0x000f6787 },
    KeywordEntry { name: "Sgn", w0: 0x000000a5, w1: 0x00000050, w2: 0x000f657d },
    KeywordEntry { name: "Shared", w0: 0x000000a6, w1: 0x00000010, w2: 0x000f6e71 },
    KeywordEntry { name: "Single", w0: 0x000000a7, w1: 0x00005010, w2: 0x000f6581 },
    KeywordEntry { name: "Spc", w0: 0x000000a8, w1: 0x00000010, w2: 0x000f6583 },
    KeywordEntry { name: "Static", w0: 0x000000a9, w1: 0x00000210, w2: 0x000f6e74 },
    KeywordEntry { name: "Step", w0: 0x000000aa, w1: 0x00000000, w2: 0x000f6e7c },
    KeywordEntry { name: "Stop", w0: 0x000000ab, w1: 0x00000210, w2: 0x000f6589 },
    KeywordEntry { name: "StrComp", w0: 0x000000ac, w1: 0x00000080, w2: 0x000f658b },
    KeywordEntry { name: "String", w0: 0x000000ad, w1: 0x00008090, w2: 0x000f7378 },
    KeywordEntry { name: "String$", w0: 0x000000ae, w1: 0x00000090, w2: 0x000f6e9e },
    KeywordEntry { name: "Sub", w0: 0x000000af, w1: 0x00000230, w2: 0x000f658e },
    KeywordEntry { name: "Tab", w0: 0x000000b0, w1: 0x00000010, w2: 0x000f658f },
    KeywordEntry { name: "Text", w0: 0x000000b1, w1: 0x00000000, w2: 0x000f64a7 },
    KeywordEntry { name: "Then", w0: 0x000000b2, w1: 0x00000010, w2: 0x000f649a },
    KeywordEntry { name: "To", w0: 0x000000b3, w1: 0x00000010, w2: 0x000f6596 },
    KeywordEntry { name: "True", w0: 0x000000b4, w1: 0x00000010, w2: 0x000f6598 },
    KeywordEntry { name: "Type", w0: 0x000000b5, w1: 0x00000210, w2: 0x000f6599 },
    KeywordEntry { name: "TypeOf", w0: 0x000000b6, w1: 0x00000010, w2: 0x000f8dfa },
    KeywordEntry { name: "UBound", w0: 0x000000b7, w1: 0x00000090, w2: 0x000f659a },
    KeywordEntry { name: "Unload", w0: 0x000000b8, w1: 0x00000000, w2: 0x000f7386 },
    KeywordEntry { name: "Unlock", w0: 0x000000b9, w1: 0x00000210, w2: 0x000f6544 },
    KeywordEntry { name: "Unknown", w0: 0x000000ba, w1: 0x00000000, w2: 0x000f6ece },
    KeywordEntry { name: "Until", w0: 0x000000bb, w1: 0x00000010, w2: 0x000f6496 },
    KeywordEntry { name: "Variant", w0: 0x000000bc, w1: 0x00007810, w2: 0x000f65a0 },
    KeywordEntry { name: "Wend", w0: 0x000000bd, w1: 0x00000210, w2: 0x000f64ab },
    KeywordEntry { name: "While", w0: 0x000000be, w1: 0x00000210, w2: 0x000f65a3 },
    KeywordEntry { name: "Width", w0: 0x000000bf, w1: 0x00000200, w2: 0x000f65a4 },
    KeywordEntry { name: "With", w0: 0x000000c0, w1: 0x00000210, w2: 0x000f6793 },
    KeywordEntry { name: "WithEvents", w0: 0x000000c1, w1: 0x00000210, w2: 0x0010d697 },
    KeywordEntry { name: "Write", w0: 0x000000c2, w1: 0x00000210, w2: 0x000f65a5 },
    KeywordEntry { name: "Xor", w0: 0x002200c3, w1: 0x00000013, w2: 0x000f65a6 },
    KeywordEntry { name: "#Const", w0: 0x000000c4, w1: 0x00000000, w2: 0x000f7abc },
    KeywordEntry { name: "#Else", w0: 0x000000c5, w1: 0x00000000, w2: 0x000f7abd },
    KeywordEntry { name: "#ElseIf", w0: 0x000000c6, w1: 0x00000000, w2: 0x000f7abd },
    KeywordEntry { name: "#End", w0: 0x000000c7, w1: 0x00000000, w2: 0x000f7abd },
    KeywordEntry { name: "#If", w0: 0x000000c8, w1: 0x00000000, w2: 0x000f7abd },
    KeywordEntry { name: "Attribute", w0: 0x000000c9, w1: 0x00000210, w2: 0x00000000 },
    KeywordEntry { name: "VB_Base", w0: 0x000000ca, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Control", w0: 0x000000cb, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Creatable", w0: 0x000000cc, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Customizable", w0: 0x000000cd, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Description", w0: 0x000000ce, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Exposed", w0: 0x000000cf, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Ext_KEY", w0: 0x000000d0, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_HelpID", w0: 0x000000d1, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Invoke_Func", w0: 0x000000d2, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Invoke_Property", w0: 0x000000d3, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Invoke_PropertyPut", w0: 0x000000d4, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Invoke_PropertyPutRef", w0: 0x000000d5, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_MemberFlags", w0: 0x000000d6, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_Name", w0: 0x000000d7, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_PredeclaredId", w0: 0x000000d8, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_ProcData", w0: 0x000000d9, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_TemplateDerived", w0: 0x000000da, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_VarDescription", w0: 0x000000db, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_VarHelpID", w0: 0x000000dc, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_VarMemberFlags", w0: 0x000000dd, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_VarProcData", w0: 0x000000de, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_UserMemId", w0: 0x000000df, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_VarUserMemId", w0: 0x000000e0, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: "VB_GlobalNameSpace", w0: 0x000000e1, w1: 0x00000010, w2: 0x00000000 },
    KeywordEntry { name: ",", w0: 0x000000e2, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: ".", w0: 0x000000e3, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "\"", w0: 0x000000e4, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "_", w0: 0x000000e5, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000e6, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000e7, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000e8, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000e9, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000ea, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000eb, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000ec, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000ed, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000ee, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000ef, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000f0, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000f1, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000f2, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "0", w0: 0x000000f3, w1: 0x00000000, w2: 0x00000000 },
    KeywordEntry { name: "!", w0: 0x000000f4, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: "#", w0: 0x000000f5, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: "&", w0: 0x002400f6, w1: 0x00000407, w2: 0x00000000 },
    KeywordEntry { name: "'", w0: 0x000000f7, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: "(", w0: 0x000000f8, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: ")", w0: 0x000000f9, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: "*", w0: 0x001800fa, w1: 0x0000040b, w2: 0x00000000 },
    KeywordEntry { name: "+", w0: 0x001600fb, w1: 0x00000408, w2: 0x00000000 },
    KeywordEntry { name: "-", w0: 0x001700fc, w1: 0x00000408, w2: 0x00000000 },
    KeywordEntry { name: ".", w0: 0x000000fd, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: "/", w0: 0x001900fe, w1: 0x0000040b, w2: 0x00000000 },
    KeywordEntry { name: ":", w0: 0x000000ff, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: ";", w0: 0x00000100, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: "<", w0: 0x002a0101, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "<=", w0: 0x00280102, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "<>", w0: 0x00270103, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "=", w0: 0x00260104, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "=<", w0: 0x00280105, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "=>", w0: 0x00290106, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: ">", w0: 0x002b0107, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "><", w0: 0x00270108, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: ">=", w0: 0x00290109, w1: 0x00000406, w2: 0x00000000 },
    KeywordEntry { name: "?", w0: 0x0000010a, w1: 0x00000500, w2: 0x00000000 },
    KeywordEntry { name: "\\", w0: 0x001e010b, w1: 0x0000040a, w2: 0x00000000 },
    KeywordEntry { name: "^", w0: 0x001a010c, w1: 0x0000040c, w2: 0x00000000 },
    KeywordEntry { name: ":=", w0: 0x0000010d, w1: 0x00000400, w2: 0x00000000 },
    KeywordEntry { name: ",", w0: 0x0000010e, w1: 0x00000400, w2: 0x00000000 },
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_length_matches_declared_size() {
        assert_eq!(KEYWORD_TABLE.len(), 271);
    }

    #[test]
    fn token_id_equals_array_index_for_every_entry() {
        // The core invariant relied on by `intern_keywords`: the low 16 bits of
        // w0 (== `token()`) are exactly the entry's index. A regeneration bug
        // that desynchronised the two tables would trip this.
        for (i, e) in KEYWORD_TABLE.iter().enumerate() {
            assert_eq!(e.token() as usize, i, "index {i} (name {:?}) has token {:#x}", e.name, e.token());
        }
    }

    #[test]
    fn known_keywords_sit_at_their_vb6_token_ids() {
        assert_eq!(KEYWORD_TABLE[0x13].name, "Call");
        assert_eq!(KEYWORD_TABLE[0x5f].name, "If");
        assert_eq!(KEYWORD_TABLE[0xaf].name, "Sub");
    }

    #[test]
    fn reserved_word_marker_is_w1_bit_4() {
        // `intern_keywords` reads bit 4 of w1 as the reserved-word marker.
        let sub = KEYWORD_TABLE.iter().find(|e| e.name == "Sub").unwrap();
        let access = KEYWORD_TABLE.iter().find(|e| e.name == "Access").unwrap();
        assert_eq!(sub.w1 >> 4 & 1, 1, "Sub is a reserved word");
        assert_eq!(access.w1 >> 4 & 1, 0, "Access is contextual, not reserved");
    }
}
