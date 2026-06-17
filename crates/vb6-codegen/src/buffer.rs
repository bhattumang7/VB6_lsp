//! The main P-code output stream.
//!
//! The runtime P-code byte stream is a dense byte sequence: each instruction
//! is a 1-byte opcode optionally followed by 1-byte, 2-byte, 4-byte, or
//! 8-byte operands at their natural position — the stream is not padded to any
//! word boundary.  `PcodeStream` backs the stream with a growable `Vec<u8>`
//! and exposes typed emit helpers so call sites match the exact layout each
//! instruction requires.

/// The P-code output stream: a byte-addressed, little-endian byte buffer.
#[derive(Debug, Default, Clone)]
pub struct PcodeStream {
    bytes: Vec<u8>,
}

/// A byte-position in the output stream. Returned by emit methods so a later
/// pass can backpatch a previously-written value (e.g. a jump target emitted
/// as a placeholder and fixed up once the destination is known).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BytePos(pub usize);

impl PcodeStream {
    pub fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    /// Current write position in bytes.
    pub fn pos(&self) -> BytePos {
        BytePos(self.bytes.len())
    }

    /// The P-code bytes written so far.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Consume the stream, yielding the exact P-code bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    // ── Core writers ─────────────────────────────────────────────────────────

    /// Write one byte and return the position it was written at.
    pub fn emit_byte(&mut self, b: u8) -> BytePos {
        let pos = self.pos();
        self.bytes.push(b);
        pos
    }

    /// Write a 16-bit little-endian unsigned value and return the position of
    /// its first byte.  Used both for 2-byte operands (frame offsets expressed
    /// as u16 bit patterns) and for multi-byte data fields.
    pub fn emit_word(&mut self, w: u16) -> BytePos {
        let pos = self.pos();
        self.bytes.extend_from_slice(&w.to_le_bytes());
        pos
    }

    /// Write a 16-bit little-endian signed value (frame offsets are signed —
    /// locals have negative offsets from the proc frame pointer).
    pub fn emit_i16(&mut self, v: i16) {
        self.bytes.extend_from_slice(&v.to_le_bytes());
    }

    /// Overwrite 2 bytes at a previously-returned position (backpatch a jump
    /// target or other forward-reference operand).
    pub fn patch_word(&mut self, at: BytePos, w: u16) {
        let off = at.0;
        self.bytes[off..off + 2].copy_from_slice(&w.to_le_bytes());
    }

    /// Emit all bytes in `src`.
    pub fn emit_bytes(&mut self, src: &[u8]) {
        self.bytes.extend_from_slice(src);
    }

    // ── Compound emit helpers ─────────────────────────────────────────────────

    /// Emit a 1-byte opcode followed by a 2-byte signed frame offset.  This is
    /// the canonical encoding for typed local-variable loads and stores in the
    /// runtime P-code stream.
    pub fn emit_load_store(&mut self, opcode: u8, frame_offset: i16) {
        self.emit_byte(opcode);
        self.emit_i16(frame_offset);
    }

    /// Emit two consecutive 16-bit words.
    pub fn emit_word4(&mut self, a: u16, b: u16) {
        self.emit_word(a);
        self.emit_word(b);
    }

    /// Emit three 16-bit words (opcode + 2 operands).
    pub fn emit_pcode3(&mut self, opcode: u16, operand1: u16, operand2: u16) {
        self.emit_word(opcode);
        self.emit_word(operand1);
        self.emit_word(operand2);
    }

    /// Emit a 2-byte opcode word followed by a 4-byte value (e.g. a Single
    /// literal whose 4-byte IEEE-754 encoding follows the opcode).
    pub fn emit_opcode4(&mut self, opcode: u16, value: [u8; 4]) {
        self.emit_word(opcode);
        self.bytes.extend_from_slice(&value);
    }

    /// Emit a 2-byte opcode word followed by an 8-byte literal payload (e.g.
    /// Currency `0xa9`, Date `0xb4`, Variant `0xaa`).
    pub fn emit_literal8(&mut self, opcode: u16, literal: [u8; 8]) {
        self.emit_word(opcode);
        self.bytes.extend_from_slice(&literal);
    }

    /// Emit a 2-byte opcode word, then a 2-byte logical-length count `len`,
    /// then `(len+1) & !1` bytes from `src` (the source must provide exactly
    /// that many bytes — the even-round-up byte count is a caller concern).
    pub fn emit_word_and_data(&mut self, opcode: u16, len: u16, src: &[u8]) {
        let copy = ((len as usize) + 1) & !1;
        assert_eq!(
            src.len(),
            copy,
            "emit_word_and_data: src must hold (len+1)&!1 = {copy} bytes"
        );
        self.emit_word(opcode);
        self.emit_word(len);
        self.bytes.extend_from_slice(src);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn emit_word_is_little_endian_and_advances() {
        let mut s = PcodeStream::new();
        let p = s.emit_word(0x1234);
        assert_eq!(p, BytePos(0));
        assert_eq!(s.emit_word(0x00ff), BytePos(2));
        assert_eq!(s.bytes(), &[0x34, 0x12, 0xff, 0x00]);
    }

    #[test]
    fn emit_byte_advances_by_one() {
        let mut s = PcodeStream::new();
        let p0 = s.emit_byte(0xab);
        assert_eq!(p0, BytePos(0));
        let p1 = s.emit_byte(0xcd);
        assert_eq!(p1, BytePos(1));
        assert_eq!(s.bytes(), &[0xab, 0xcd]);
    }

    #[test]
    fn emit_i16_writes_signed_le() {
        let mut s = PcodeStream::new();
        s.emit_i16(-140); // 0xff74 as i16
        assert_eq!(s.bytes(), &[0x74, 0xff]);
    }

    #[test]
    fn emit_load_store_writes_three_bytes() {
        // Double-load opcode 0x6f at frame offset 0xff74 (-140 as i16).
        let mut s = PcodeStream::new();
        s.emit_load_store(0x6f, -140);
        assert_eq!(s.bytes(), &[0x6f, 0x74, 0xff]);
    }

    #[test]
    fn emit_word4_writes_two_words_in_order() {
        let mut s = PcodeStream::new();
        s.emit_word4(0x1234, 0x5678);
        assert_eq!(s.bytes(), &[0x34, 0x12, 0x78, 0x56]);
    }

    #[test]
    fn emit_pcode3_writes_opcode_then_two_operands() {
        let mut s = PcodeStream::new();
        s.emit_pcode3(0x00e0, 0x0001, 0x0002);
        assert_eq!(s.bytes(), &[0xe0, 0x00, 0x01, 0x00, 0x02, 0x00]);
    }

    #[test]
    fn emit_literal8_writes_opcode_then_eight_bytes() {
        let mut s = PcodeStream::new();
        let payload = 10_000_i64.to_le_bytes();
        s.emit_literal8(0x00a9, payload);
        let mut expect = vec![0xa9, 0x00];
        expect.extend_from_slice(&payload);
        assert_eq!(s.bytes(), expect.as_slice());
    }

    #[test]
    fn emit_word_and_data_even_length_string() {
        let mut s = PcodeStream::new();
        s.emit_word_and_data(0x00b6, 2, b"AB");
        assert_eq!(s.bytes(), &[0xb6, 0x00, 0x02, 0x00, 0x41, 0x42]);
    }

    #[test]
    fn emit_word_and_data_records_logical_len_but_copies_rounded() {
        let mut s = PcodeStream::new();
        s.emit_word_and_data(0x00b6, 3, b"ABC\0");
        assert_eq!(s.bytes(), &[0xb6, 0x00, 0x03, 0x00, 0x41, 0x42, 0x43, 0x00]);
    }

    #[test]
    fn patch_word_backpatches_in_place() {
        let mut s = PcodeStream::new();
        s.emit_word(0x00e0);
        let target = s.emit_word(0xffff);
        s.emit_word(0x0001);
        s.patch_word(target, 0x0042);
        assert_eq!(s.bytes(), &[0xe0, 0x00, 0x42, 0x00, 0x01, 0x00]);
    }

    #[test]
    fn byte_and_word_interleave_without_alignment_constraint() {
        // Runtime stream mixes 1-byte opcodes and 2-byte operands freely.
        let mut s = PcodeStream::new();
        s.emit_byte(0xf4); // push-small-int opcode
        s.emit_byte(0x02); // immediate 2
        s.emit_byte(0xeb); // coerce opcode
        s.emit_load_store(0x74, -140i16); // store-Double at 0xff74
        assert_eq!(s.bytes(), &[0xf4, 0x02, 0xeb, 0x74, 0x74, 0xff]);
    }
}
