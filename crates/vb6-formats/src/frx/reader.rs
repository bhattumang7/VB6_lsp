/// Errors returned by [`FrxReader`].
#[derive(Debug, PartialEq, Eq)]
pub enum FrxError {
    /// A read would extend past the end of the data slice.
    UnexpectedEof { pos: usize, needed: usize, available: usize },
    /// A seek target is out of range for the data slice.
    SeekOutOfRange { target: usize, len: usize },
    /// A record magic did not match the expected value.
    BadMagic { pos: usize, expected: u16, got: u16 },
    /// A declared record length is not plausible (e.g. exceeds remaining data).
    LengthOverflow { pos: usize, declared: u32, remaining: usize },
}

impl core::fmt::Display for FrxError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            FrxError::UnexpectedEof { pos, needed, available } =>
                write!(f, "unexpected EOF at offset {pos:#010x}: need {needed}B, have {available}B"),
            FrxError::SeekOutOfRange { target, len } =>
                write!(f, "seek to {target:#010x} out of range (file length {len:#010x})"),
            FrxError::BadMagic { pos, expected, got } =>
                write!(f, "bad magic at {pos:#010x}: expected {expected:#06x}, got {got:#06x}"),
            FrxError::LengthOverflow { pos, declared, remaining } =>
                write!(f, "declared length {declared} at {pos:#010x} exceeds remaining {remaining}B"),
        }
    }
}

/// Byte cursor over an FRX file (or any sub-slice thereof).
///
/// Constructed with [`FrxReader::new`], then seeked to the record offset
/// recorded in the `.frm` file.  All reads advance the internal position.
pub struct FrxReader<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> FrxReader<'a> {
    /// Create a new reader over `data`, positioned at byte 0.
    pub fn new(data: &'a [u8]) -> Self {
        FrxReader { data, pos: 0 }
    }

    /// Current byte offset within the data slice.
    pub fn pos(&self) -> usize {
        self.pos
    }

    /// Number of bytes remaining from the current position.
    pub fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.pos)
    }

    /// Move the cursor to an absolute byte `offset`.
    pub fn seek(&mut self, offset: usize) -> Result<(), FrxError> {
        if offset > self.data.len() {
            return Err(FrxError::SeekOutOfRange { target: offset, len: self.data.len() });
        }
        self.pos = offset;
        Ok(())
    }

    // --- primitive reads ----------------------------------------------------

    /// Read one byte.
    pub fn read_u8(&mut self) -> Result<u8, FrxError> {
        self.need(1)?;
        let v = self.data[self.pos];
        self.pos += 1;
        Ok(v)
    }

    /// Read a little-endian `u16`.
    pub fn read_u16_le(&mut self) -> Result<u16, FrxError> {
        self.need(2)?;
        let v = u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]);
        self.pos += 2;
        Ok(v)
    }

    /// Read a little-endian `u32`.
    pub fn read_u32_le(&mut self) -> Result<u32, FrxError> {
        self.need(4)?;
        let v = u32::from_le_bytes(self.data[self.pos..self.pos + 4].try_into().unwrap());
        self.pos += 4;
        Ok(v)
    }

    /// Read a little-endian `i32`.
    pub fn read_i32_le(&mut self) -> Result<i32, FrxError> {
        Ok(self.read_u32_le()? as i32)
    }

    /// Read exactly `n` bytes and return a slice into the backing data.
    pub fn read_bytes(&mut self, n: usize) -> Result<&'a [u8], FrxError> {
        self.need(n)?;
        let slice = &self.data[self.pos..self.pos + n];
        self.pos += n;
        Ok(slice)
    }

    /// Peek at the next `u16` without advancing the cursor.
    pub fn peek_u16_le(&self) -> Result<u16, FrxError> {
        if self.remaining() < 2 {
            return Err(FrxError::UnexpectedEof {
                pos: self.pos,
                needed: 2,
                available: self.remaining(),
            });
        }
        Ok(u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]))
    }

    /// Peek at the next `u32` without advancing the cursor.
    pub fn peek_u32_le(&self) -> Result<u32, FrxError> {
        if self.remaining() < 4 {
            return Err(FrxError::UnexpectedEof {
                pos: self.pos,
                needed: 4,
                available: self.remaining(),
            });
        }
        Ok(u32::from_le_bytes(self.data[self.pos..self.pos + 4].try_into().unwrap()))
    }

    // --- compound reads -----------------------------------------------------

    /// Read a `u32`-length-prefixed byte string.
    ///
    /// The four-byte length field gives the number of data bytes that follow.
    /// Returns a slice into the backing data (zero-copy).
    pub fn read_len_prefixed_bytes(&mut self) -> Result<&'a [u8], FrxError> {
        let at = self.pos;
        let len = self.read_u32_le()? as usize;
        if len > self.remaining() {
            return Err(FrxError::LengthOverflow {
                pos: at,
                declared: len as u32,
                remaining: self.remaining(),
            });
        }
        self.read_bytes(len)
    }

    /// Require that a specific `u16` magic follows at the current position.
    ///
    /// Advances past the magic on success; leaves the cursor unchanged on error.
    pub fn require_magic(&mut self, expected: u16) -> Result<(), FrxError> {
        let at = self.pos;
        let got = self.read_u16_le()?;
        if got != expected {
            self.pos = at; // rewind so the caller can inspect
            return Err(FrxError::BadMagic { pos: at, expected, got });
        }
        Ok(())
    }

    // --- internal -----------------------------------------------------------

    fn need(&self, n: usize) -> Result<(), FrxError> {
        let avail = self.remaining();
        if avail < n {
            Err(FrxError::UnexpectedEof { pos: self.pos, needed: n, available: avail })
        } else {
            Ok(())
        }
    }
}
