/// Opcode bytes for an explicit type-conversion intrinsic, keyed by
/// (destination type tag, source type tag). Distinct from the implicit
/// assignment-coercion family: a same-type conversion is usually a no-op, and the
/// floating-point destinations use the dedicated `0xfc 0x39..0x41` block. Returns
/// an empty slice for a no-op; an empty slice is also returned (caller treats as a
/// gate) for unsupported pairs (Byte, Boolean/Date/Variant destinations).
/// Type tags: Integer=6, Long=8, Single=0xa, Double=0xb, Currency=0xd, String=0x10.
pub(super) fn explicit_conversion_bytes(dest: i32, src: i32) -> &'static [u8] {
    match (dest, src) {
        // → Integer
        (6, 6) => &[],
        (6, 8) => &[0xe4],
        (6, 0xa) | (6, 0xb) => &[0xe5],
        (6, 0xd) => &[0xe6],
        // → Long
        (8, 6) => &[0xe7],
        (8, 8) => &[],
        (8, 0xa) | (8, 0xb) => &[0xe8],
        (8, 0xd) => &[0xe9],
        // → Single
        (0xa, 6) => &[0xeb],
        (0xa, 8) => &[0xfc, 0x3e],
        (0xa, 0xa) => &[],
        (0xa, 0xb) => &[0xfc, 0x40],
        (0xa, 0xd) => &[0xfc, 0x41],
        // → Double
        (0xb, 6) => &[0xeb],
        (0xb, 8) => &[0xec],
        (0xb, 0xa) => &[0xfc, 0x3a],
        (0xb, 0xb) => &[0xfc, 0x3b],
        (0xb, 0xd) => &[0xfc, 0x39],
        // → Currency
        (0xd, 6) => &[0xef],
        (0xd, 8) => &[0xf0],
        (0xd, 0xa) | (0xd, 0xb) => &[0xf1],
        (0xd, 0xd) => &[],
        // → String
        (0x10, 6) => &[0xfb, 0xfd],
        (0x10, 8) => &[0xfb, 0xfe],
        (0x10, 0xa) => &[0xfb, 0xff],
        (0x10, 0xb) => &[0xfc, 0x00],
        (0x10, 0xd) => &[0xfc, 0x01],
        _ => &[],
    }
}

/// Opcode bytes for a dedicated-opcode unary intrinsic, keyed by the kind (Len=0,
/// Abs=1, Sgn=2, Int=3, Fix=4) and the argument type tag. `Int`/`Fix` of an
/// already-integral argument is a no-op (empty). Type tags: Integer=6, Long=8,
/// Single=0xa, Double=0xb, Currency=0xd.
pub(super) fn unary_intrinsic_bytes(kind: u32, arg: i32) -> &'static [u8] {
    match (kind, arg) {
        // Len → Long: a single opcode regardless of (String) argument.
        (0, _) => &[0x4a],
        // Abs → argument type.
        (1, 6) => &[0xbb],
        (1, 8) => &[0xbc],
        (1, 0xa) | (1, 0xb) => &[0xbd],
        (1, 0xd) => &[0xbe],
        // Sgn → Integer.
        (2, 6) => &[0xfb, 0xf3],
        (2, 8) => &[0xfb, 0xf4],
        (2, 0xa) => &[0xfb, 0xf5],
        (2, 0xb) => &[0xfb, 0xf6],
        (2, 0xd) => &[0xfb, 0xf7],
        // Int → argument type (no-op for integral arguments).
        (3, 6) | (3, 8) => &[],
        (3, 0xa) => &[0xfb, 0xe6],
        (3, 0xb) => &[0xfb, 0xe7],
        (3, 0xd) => &[0xfb, 0xe8],
        // Fix → argument type (no-op for integral arguments).
        (4, 6) | (4, 8) => &[],
        (4, 0xa) => &[0xfb, 0xde],
        (4, 0xb) => &[0xfb, 0xdf],
        (4, 0xd) => &[0xfb, 0xe0],
        _ => &[],
    }
}
