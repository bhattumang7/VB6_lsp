use super::*;

impl<'a> Emitter<'a> {
    // ── Typed local / argument / global load and store ───────────────────────

    /// Dispatch a synthetic ByRef parameter load node (opcode 0x75): the bound
    /// symbol child carries the frame offset in `type_info()`; `word[5]` is the
    /// type context.
    pub(super) fn emit_byref_param_node(&mut self, n: &RawNode) {
        let type_ctx = n.word(5) as usize;
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.emit_byref_load(type_ctx, frame_offset);
    }

    /// Dispatch a synthetic module-global load node (opcode 0x77): `word[4]` low
    /// u16 = module descriptor, high u16 = field offset; `word[5]` = type context.
    pub(super) fn emit_global_node_load(&mut self, n: &RawNode) {
        let type_ctx = n.word(5) as usize;
        let packed = n.word(4);
        let module_desc = (packed & 0xffff) as u16;
        let field_offset = (packed >> 16) as u16;
        self.emit_global_load(type_ctx, module_desc, field_offset);
    }

    /// Emit a typed local-variable load. The synthetic node carries the frame
    /// offset in the bound symbol child's `type_info()` and the type context in
    /// `word[5]`. Node types `0x74` and `0x76` both route here.
    pub(super) fn emit_var_load(&mut self, n: &RawNode, _context: u32) {
        let type_ctx = n.word(5) as usize;
        let sym = self.arena.get(n.lhs());
        let frame_offset = sym.type_info() as i16;
        self.emit_typed_load(type_ctx, frame_offset);
    }

    /// Emit a typed local-variable load from its type context and frame offset.
    /// The opcode comes from [`RT_LOAD_BY_CTX`]; the frame offset follows as a
    /// 2-byte signed little-endian value. Mirror of [`Self::emit_var_store`].
    pub fn emit_typed_load(&mut self, type_ctx: usize, frame_offset: i16) {
        // Byte (ctx 7) has a 2-byte escape-paged load opcode (`fc e0`) that the
        // single-byte RT_LOAD_BY_CTX shortcut cannot hold; emit via the value-
        // emitter load index 0x1e0 (RT_OPCODE_BYTE[0x1e0] = 0xfc → escape).
        if type_ctx == 7 {
            self.emit_value2(0x1e0);
            self.stream.emit_i16(frame_offset);
            return;
        }
        // String (ctx 8): the BSTR-pointer load (0x6c) via the value-emitter load
        // index 0x1e7 (string value-class).
        if type_ctx == 8 {
            self.emit_value2(0x1e7);
            self.stream.emit_i16(frame_offset);
            return;
        }
        let opcode = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if opcode == 0 {
            unimplemented!("no load opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a typed local-variable store. Mirror of [`Self::emit_var_load`] using
    /// [`RT_STORE_BY_CTX`]. The caller must have emitted the value to store first.
    pub fn emit_var_store(&mut self, type_ctx: usize, frame_offset: i16) {
        // Byte (ctx 7): 2-byte escape store opcode (`fc f0`) via value-emitter
        // store index 0x1f0 (RT_OPCODE_BYTE[0x1f0] = 0xfc → escape).
        if type_ctx == 7 {
            self.emit_value2(0x1f0);
            self.stream.emit_i16(frame_offset);
            return;
        }
        // String (ctx 8): the refcounted BSTR assign store (0x43) via index 0x201.
        if type_ctx == 8 {
            self.emit_value2(0x201);
            self.stream.emit_i16(frame_offset);
            return;
        }
        // String move-store (ctx 9): store a freshly-produced string temp (e.g. a
        // concat result) without addref — opcode 0x31 via index 0x1f7.
        if type_ctx == 9 {
            self.emit_value2(0x1f7);
            self.stream.emit_i16(frame_offset);
            return;
        }
        let opcode = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if opcode == 0 {
            unimplemented!("no store opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(opcode);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a ByRef parameter load at `frame_offset`. The ByRef load opcode is
    /// `RT_LOAD_BY_CTX[type_ctx] + 0x14`; the offset is positive (parameters sit
    /// above the frame pointer).
    pub fn emit_byref_load(&mut self, type_ctx: usize, frame_offset: i16) {
        let base = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no load opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x14);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a ByRef parameter store at `frame_offset`. The ByRef store opcode is
    /// `RT_STORE_BY_CTX[type_ctx] + 0x14`.
    pub fn emit_byref_store(&mut self, type_ctx: usize, frame_offset: i16) {
        let base = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no store opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x14);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a module-level global variable load. The opcode is
    /// `RT_LOAD_BY_CTX[type_ctx] + 0x28`; the 4-byte operand encodes the module
    /// descriptor in bytes 0–1 and the field offset (byte offset within the
    /// module's global data block) in bytes 2–3.
    pub fn emit_global_load(
        &mut self,
        type_ctx: usize,
        module_desc: u16,
        field_offset: u16,
    ) {
        let base = RT_LOAD_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no load opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x28);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(field_offset);
    }

    // ── Store-opcode selection for `=` assignment ────────────────────────────

    /// Store-opcode base for a destination type tag: the entry of
    /// [`RT_ASSIGN_STORE_OPCODE`] at the destination's type-offset class.
    /// Type tags whose class falls outside the store-opcode table are not valid
    /// assignment destinations on this path.
    pub(super) fn assign_store_base(dest_tag: i32) -> i32 {
        use crate::tables::RT_ASSIGN_STORE_OPCODE;
        let class = RT_TYPE_OFFSET[dest_tag as usize] as usize;
        if class >= RT_ASSIGN_STORE_OPCODE.len() {
            unimplemented!(
                "assignment store for type-offset class {class}: outside the \
                 store-opcode table; Phase 4"
            );
        }
        RT_ASSIGN_STORE_OPCODE[class]
    }

    /// Source-class adjustment added to the store base: the source's type-offset
    /// class, with `10 -> 4` and `9 -> 1` applied.
    pub(super) fn assign_source_adjust(src_tag: i32) -> i32 {
        match RT_TYPE_OFFSET[src_tag as usize] {
            10 => 4,
            9 => 1,
            c => c,
        }
    }

    /// Emit the store opcode for a simple `=` assignment after the value has
    /// already been pushed. `n` is the assignment node; the source is `word[4]`.
    ///
    /// The general store opcode is `assign_store_base(dest) + assign_source_adjust(src)`,
    /// with direct opcodes for specific Variant / Currency / Boolean / object
    /// type pairs.
    /// Dispatch a synthetic operand-coercion node (opcode 0x78): emit the child
    /// operand, then the conversion opcode that widens it to the node's (target)
    /// type. The conversion opcode index is
    /// `assign_store_base(target) + assign_source_adjust(src)` — the same
    /// store/coerce opcode family the `=` store uses (RT_ASSIGN_STORE_OPCODE
    /// indexed by RT_TYPE_OFFSET[target], plus the source-class adjust). E.g.
    /// Integer→Long → 0x11c+1 = 0x11d (byte 0xe7); Long→Double → 0x12c+2 = 0x12e
    /// (byte 0xec).
    pub(super) fn emit_coerce_node(&mut self, n: &RawNode) {
        self.emit_expr(NodeRef(n.w[4]), 2);
        let target_tag = n.type_tag();
        let src_tag = self.arena.get(NodeRef(n.w[4])).type_tag();
        self.emit_conversion(target_tag, src_tag);
    }

    /// Emit the runtime conversion opcode that converts a value of `src_tag` to
    /// `target_tag`. A Date destination uses dedicated opcodes (a Date carries an
    /// OLE serial with its own range/validity conversion); a Single source has
    /// already been widened to the common float representation by its load, so it
    /// converts as Double. Every other pair uses the base+adjust store family.
    pub fn emit_conversion(&mut self, target_tag: i32, src_tag: i32) {
        if target_tag == 0xc {
            match src_tag {
                0xa | 0xb => {
                    self.emit_value2(0x147);
                    return;
                }
                0xf => {
                    self.emit_value2(0x14f);
                    return;
                }
                0x10 => {
                    self.emit_value2(0x3c9);
                    return;
                }
                _ => {}
            }
        }
        let opcode = Self::assign_store_base(target_tag) + Self::assign_source_adjust(src_tag);
        self.emit_value2(opcode as usize);
    }

    /// Emit a load-address (`0x04`) of a frame slot at `frame_offset`.
    pub fn emit_ldaddr(&mut self, frame_offset: i16) {
        self.stream.emit_byte(0x04);
        self.stream.emit_i16(frame_offset);
    }

    /// Emit a Static-local load (synthetic node 0x7b): `0x5f <module_desc> 0x0004
    /// <load-op> <static offset>`. The load opcode follows from the type tag.
    pub(super) fn emit_static_load(&mut self, n: &RawNode) {
        let module_desc = (n.w[4] & 0xffff) as u16;
        let static_off = (n.w[4] >> 16) as u16;
        self.stream.emit_byte(0x5f);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(0x0004);
        match n.type_tag() {
            6 => { self.stream.emit_byte(0x89); }       // Integer / Boolean
            8 | 0x10 => { self.stream.emit_byte(0x8a); } // Long / String pointer
            0xd => { self.stream.emit_byte(0x8b); }     // Currency
            10 => { self.stream.emit_byte(0x8c); }      // Single
            11 | 0xc => { self.stream.emit_byte(0x8d); } // Double / Date
            5 => {
                self.stream.emit_byte(0xfd);
                self.stream.emit_byte(0x70);
            }
            _ => {}
        }
        self.stream.emit_word(static_off);
    }

    pub(super) fn emit_assign_op(&mut self, n: &RawNode) {
        let source = *self.arena.get(NodeRef(n.w[4]));
        let dest_hi = n.w[0] & 0xffff_0000;
        let dest_tag = (n.w[0] as i32) >> 16;
        let src_tag = (source.w[0] as i32) >> 16;

        // Object destination: a sized store. Sources whose type tag is in
        // [3, 0x17] go through a per-source-type sub-dispatch that emits size
        // operand words (needs the type-descriptor model); any other source
        // uses the store table with a trailing size operand.
        if dest_hi == 0xf0000 {
            let size = self.emit_get_type_size3(n.w[6]);
            if ((src_tag - 3) as u32) < 0x15 {
                unimplemented!(
                    "sized object/UDT store (per-source-type sub-dispatch emitting \
                     size operand words); needs the type-descriptor model; Phase 4"
                );
            }
            let opcode = Self::assign_source_adjust(src_tag) + Self::assign_store_base(dest_tag);
            self.emit_opcode2(opcode as usize, size as u16);
            return;
        }

        // Currency destination: direct opcodes for specific source kinds.
        if dest_hi == 0xc0000 {
            match src_tag {
                0xb => { self.emit_value2(0x147); return; }
                0xf => { self.emit_value2(0x14f); return; }
                0x10 => { self.emit_value2(0x3c9); return; }
                _ => {}
            }
        }

        // Both sides Variant / ByRef-Variant / Currency: guarded table store.
        if matches!(dest_tag, 10 | 0xb | 0xc) && matches!(src_tag, 10 | 0xb | 0xc) {
            // Flag-byte bit 0x80 clear → handled elsewhere (no-op here).
            if (n.w[1] >> 8) & 0x80 == 0 {
                return;
            }
            let base = Self::assign_store_base(dest_tag);
            if RT_TYPE_OFFSET[src_tag as usize] == 10 {
                self.emit_value2((base + 4) as usize);
                return;
            }
            self.emit_value2((Self::assign_source_adjust(src_tag) + base) as usize);
            return;
        }

        // Otherwise inspect the source's type region for special stores.
        let src_hi = source.w[0] & 0xffff_0000;
        if src_hi == 0xc0000 {
            // Currency source into a non-Currency destination.
            match dest_tag {
                0xf => { self.emit_value2(0x2fb); return; }
                0x10 => { self.emit_value2(0x3c8); return; }
                _ => {}
            }
        }
        if src_hi == 0x30000 {
            match dest_tag {
                5 => { self.emit_value2(0x138); return; }
                6 => { return; }
                0x10 => { self.emit_value2(0x3c7); return; }
                _ => {}
            }
        }
        if src_hi == 0x110000 {
            // Fixed-length string source: needs the type-length lookup.
            unimplemented!(
                "fixed-length string source store: needs the type-length lookup; Phase 4"
            );
        }
        if src_hi == 0xf0000 {
            // Object source into a Variant / TypeOf destination.
            if dest_hi == 0x140000 {
                unimplemented!(
                    "object source into a Variant target: needs the object-reference \
                     resolution path; Phase 4"
                );
            }
            if dest_hi == 0x120000 {
                unimplemented!(
                    "object source into a TypeOf target: needs the object-reference \
                     resolution path; Phase 4"
                );
            }
        }

        // Generic store: base from destination, adjustment from source.
        let base = Self::assign_store_base(dest_tag);
        if RT_TYPE_OFFSET[src_tag as usize] == 10 {
            self.emit_value2((base + 4) as usize);
            return;
        }
        self.emit_value2((Self::assign_source_adjust(src_tag) + base) as usize);
    }

    /// Emit a module-level global variable store. The opcode is
    /// `RT_STORE_BY_CTX[type_ctx] + 0x28`.
    pub fn emit_global_store(
        &mut self,
        type_ctx: usize,
        module_desc: u16,
        field_offset: u16,
    ) {
        let base = RT_STORE_BY_CTX.get(type_ctx).copied().unwrap_or(0);
        if base == 0 {
            unimplemented!("no store opcode for type context {}", type_ctx);
        }
        self.stream.emit_byte(base + 0x28);
        self.stream.emit_word(module_desc);
        self.stream.emit_word(field_offset);
    }
}
