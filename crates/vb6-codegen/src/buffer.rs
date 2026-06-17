//! The main P-code output stream.
//!
//! P-code is emitted as little-endian 16-bit words onto a byte-addressed buffer
//! advanced by a word cursor, with periodic capacity checks that grow the backing
//! store. Those capacity checks are pure buffer management: they never change the
//! emitted bytes, only where the backing memory lives. We back the stream with a
//! growable `Vec<u8>`, so the checks become no-ops and the emitted byte sequence
//! is identical.
//!
//! Each method below emits one P-code primitive; the byte output is exactly what
//! the VB6 P-code format prescribes for that primitive.

/// The P-code output stream: a byte-addressed, little-endian, 2-byte-word stream.
#[derive(Debug, Default, Clone)]
pub struct PcodeStream {
    bytes: Vec<u8>,
}

/// A position in the stream, in 16-bit words from the start. Returned by the
/// word emitters so a later pass can backpatch a previously written word (jump
/// targets are emitted as placeholders and fixed up once known).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WordPos(pub usize);

impl PcodeStream {
    pub fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    /// Number of 16-bit words written so far (the word cursor offset).
    pub fn word_len(&self) -> usize {
        debug_assert_eq!(self.bytes.len() % 2, 0, "stream must stay 2-byte aligned");
        self.bytes.len() / 2
    }

    /// The current write position, in words.
    pub fn pos(&self) -> WordPos {
        WordPos(self.word_len())
    }

    /// Consume the stream, yielding the exact P-code bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// The P-code bytes written so far.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    // ── Core writers ─────────────────────────────────────────────────────────

    /// Write one 16-bit word (little-endian) and advance the cursor by one word.
    /// Returns the position of the word just written, for backpatching. This is
    /// the fundamental write-and-advance step shared by every emit primitive.
    pub fn emit_word(&mut self, w: u16) -> WordPos {
        let pos = self.pos();
        self.bytes.extend_from_slice(&w.to_le_bytes());
        pos
    }

    /// Overwrite a previously written word (backpatch), e.g. a jump-target fixup.
    pub fn patch_word(&mut self, at: WordPos, w: u16) {
        let off = at.0 * 2;
        self.bytes[off..off + 2].copy_from_slice(&w.to_le_bytes());
    }

    // ── Emit primitives ──────────────────────────────────────────────────────

    /// Write two consecutive 16-bit words.
    pub fn emit_word4(&mut self, a: u16, b: u16) {
        self.emit_word(a);
        self.emit_word(b);
    }

    /// Write three 16-bit words (opcode + 2 operands).
    pub fn emit_pcode3(&mut self, opcode: u16, operand1: u16, operand2: u16) {
        self.emit_word(opcode);
        self.emit_word(operand1);
        self.emit_word(operand2);
    }

    /// Write a 2-byte opcode followed by a 4-byte value (used for the Single
    /// literal opcode 0xb3).
    pub fn emit_opcode4(&mut self, opcode: u16, value: [u8; 4]) {
        self.emit_word(opcode);
        self.bytes.extend_from_slice(&value);
    }

    /// Write a 2-byte opcode followed by an 8-byte literal value (used for
    /// Currency `0xa9`, Date `0xb4`, Variant `0xaa`, …). The 8 bytes are written
    /// exactly as stored (little-endian payload).
    pub fn emit_literal8(&mut self, opcode: u16, literal: [u8; 8]) {
        self.emit_word(opcode);
        self.bytes.extend_from_slice(&literal);
    }

    /// Write a 2-byte opcode, then a 2-byte count holding the *logical* length
    /// `len`, then copy `(len+1) & !1` bytes from `src` (even-aligned). The count
    /// word records the real `len`, but the copy is the rounded-up byte count
    /// taken straight from the source buffer, so for odd `len` the trailing
    /// alignment byte is simply the source's next byte — a caller concern, not
    /// invented here. `src` must therefore provide exactly the rounded number of
    /// bytes to copy.
    pub fn emit_word_and_data(&mut self, opcode: u16, len: u16, src: &[u8]) {
        let copy = ((len as usize) + 1) & !1;
        assert_eq!(
            src.len(),
            copy,
            "emit_word_and_data: src must hold the even-rounded byte count \
             ((len+1)&!1 = {copy}); exactly that many bytes are copied",
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
        assert_eq!(p, WordPos(0));
        assert_eq!(s.emit_word(0x00ff), WordPos(1));
        assert_eq!(s.bytes(), &[0x34, 0x12, 0xff, 0x00]);
        assert_eq!(s.word_len(), 2);
    }

    #[test]
    fn emit_word4_writes_two_words_in_order() {
        let mut s = PcodeStream::new();
        s.emit_word4(0x1234, 0x5678);
        assert_eq!(s.bytes(), &[0x34, 0x12, 0x78, 0x56]);
    }

    #[test]
    fn emit_pcode3_writes_opcode_then_two_operands() {
        // The shape of a jump: opcode 0xe0, then a value and a target.
        let mut s = PcodeStream::new();
        s.emit_pcode3(0x00e0, 0x0001, 0x0002);
        assert_eq!(s.bytes(), &[0xe0, 0x00, 0x01, 0x00, 0x02, 0x00]);
    }

    #[test]
    fn emit_literal8_writes_opcode_then_eight_bytes() {
        // Currency literal opcode 0xa9 + 8-byte payload (1.0000@ = 10000 scaled,
        // stored as i64 LE).
        let mut s = PcodeStream::new();
        let payload = 10_000_i64.to_le_bytes();
        s.emit_literal8(0x00a9, payload);
        let mut expect = vec![0xa9, 0x00];
        expect.extend_from_slice(&payload);
        assert_eq!(s.bytes(), expect.as_slice());
        assert_eq!(s.word_len(), 5); // 1 opcode word + 4 payload words
    }

    #[test]
    fn emit_word_and_data_even_length_string() {
        // String literal opcode 0xb6, count = real length, then the bytes.
        let mut s = PcodeStream::new();
        s.emit_word_and_data(0x00b6, 2, b"AB");
        assert_eq!(s.bytes(), &[0xb6, 0x00, 0x02, 0x00, 0x41, 0x42]);
    }

    #[test]
    fn emit_word_and_data_records_logical_len_but_copies_rounded() {
        // Odd logical length 3: count word holds 3, but 4 bytes are copied from
        // the caller-supplied source (here the trailing NUL the source carries).
        let mut s = PcodeStream::new();
        s.emit_word_and_data(0x00b6, 3, b"ABC\0");
        assert_eq!(s.bytes(), &[0xb6, 0x00, 0x03, 0x00, 0x41, 0x42, 0x43, 0x00]);
    }

    #[test]
    fn patch_word_backpatches_in_place() {
        let mut s = PcodeStream::new();
        s.emit_word(0x00e0);
        let target = s.emit_word(0xffff); // placeholder jump target
        s.emit_word(0x0001);
        s.patch_word(target, 0x0042);
        assert_eq!(s.bytes(), &[0xe0, 0x00, 0x42, 0x00, 0x01, 0x00]);
    }
}
