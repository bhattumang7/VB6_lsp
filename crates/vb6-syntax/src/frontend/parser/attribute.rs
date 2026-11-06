//! `Attribute` lines.
//!
//! A line is read by position, one token at a time: the member before the dot (a plain name),
//! the attribute word, and what that word allows after it. A word that is not an attribute word,
//! or a first token that is not a name, ends the reading: the rest of the line is passed over and
//! nothing is wrong with it. What follows the value is never looked at.

use super::*;
use crate::frontend::module_name;

/// The attribute words (`VB_Base` through `VB_GlobalNameSpace`).
fn attribute_word(kind: &TokenKind) -> Option<Kw> {
    match kind {
        TokenKind::Kw(k) if (Kw::VbBase as u16..=Kw::VbGlobalNameSpace as u16).contains(&(*k as u16)) => Some(*k),
        _ => None,
    }
}

/// The four words that follow `VB_ProcData`; as the attribute word itself they are passed over.
fn is_invoke_word(k: Kw) -> bool {
    matches!(k, Kw::VbInvokeFunc | Kw::VbInvokeProperty | Kw::VbInvokePropertyPut | Kw::VbInvokePropertyPutRef)
}

/// What a token is to an `Attribute` line.
enum Word {
    /// A name: an identifier, a bracketed word, or a word that is not reserved.
    Name,
    Attribute(Kw),
    Other,
}

/// The value an `Attribute` line gives, by position in the line's tokens.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Value {
    /// Nothing the line's reader takes as a value follows the `=`; the index is that of the
    /// token after it.
    None(usize),
    /// A literal at this index.
    Literal(usize),
    /// A minus sign at this index and the integer literal after it.
    Negated(usize),
    /// `VB_Ext_KEY`: two strings, at these indexes.
    Pair(usize, usize),
}

/// How reading an `Attribute` line ended; indexes are into the tokens after `Attribute`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum AttributeRead {
    /// `[member.]word = value` was read.
    Applied { member: Option<usize>, word: usize, value: Value },
    /// The line is passed over: it does not begin with a name and an attribute word.
    Passed,
    /// The token at this index is not what the line needs there.
    Wrong(usize),
    /// `As` follows `Attribute`: it is read as the start of a type clause.
    TypeClause,
}

/// What `take_body_attribute_line` found.
pub(super) enum BodyAttribute {
    /// The line is not an attribute line.
    None,
    /// The line was an attribute line and has been read.
    Taken,
    /// An attribute line follows the member on this line, after a colon: the member is read for
    /// its errors and dropped.
    AfterMember,
}

impl<'src> Parser<'src> {
    /// What the token is to an `Attribute` line.
    fn attribute_line_word(&self, token: &Token) -> Word {
        let bracketed = self.src.get(token.span.start as usize) == Some(&b'[');
        if let (TokenKind::Ident, _) | (TokenKind::Kw(_), Some(_)) = (&token.kind, token.sym) {
            if bracketed || token.kind == TokenKind::Ident {
                return Word::Name;
            }
        }
        if let Some(k) = attribute_word(&token.kind) {
            return Word::Attribute(k);
        }
        match (&token.kind, token.sym) {
            (TokenKind::Kw(Kw::Attribute), _) => Word::Other,
            (TokenKind::Kw(_), Some(sym)) if !self.scanner.is_reserved(sym as u32) => Word::Name,
            _ => Word::Other,
        }
    }

    /// True when a blank (not a line continuation) lies between the two tokens.
    fn blank_between(&self, first: &Token, second: &Token) -> bool {
        let end = (first.span.start + first.span.len) as usize;
        let start = second.span.start as usize;
        start > end && start <= self.src.len() && !self.src[end..start].iter().any(|b| *b == b'\n' || *b == b'\r')
    }

    /// Reads the tokens of an `Attribute` line that follow the word `Attribute`; `toks` ends with
    /// the terminator of the line.
    fn read_attribute(&self, toks: &[Token]) -> AttributeRead {
        let last = toks.len() - 1;
        let at = |i: usize| &toks[i.min(last)];
        let ends = |i: usize| matches!(at(i).kind, TokenKind::Eol | TokenKind::Eof | TokenKind::Kw(Kw::Apos));
        if at(0).kind == TokenKind::Kw(Kw::As) {
            return AttributeRead::TypeClause;
        }
        let (member, word_at) = match self.attribute_line_word(at(0)) {
            Word::Attribute(_) => (None, 0),
            Word::Name => {
                // The dot joins the name to the word only when it follows the name directly.
                if at(1).kind != TokenKind::Kw(Kw::Dot) || self.blank_between(at(0), at(1)) {
                    return AttributeRead::Wrong(1.min(last));
                }
                (Some(0), 2)
            }
            Word::Other => return AttributeRead::Passed,
        };
        let Word::Attribute(word) = self.attribute_line_word(at(word_at)) else {
            return AttributeRead::Passed;
        };
        if is_invoke_word(word) {
            return AttributeRead::Passed;
        }
        let mut i = word_at + 1;
        if word == Kw::VbExtKey {
            return self.read_extended_key(toks, member, word_at);
        }
        if word == Kw::VbProcData {
            // One token is passed over without being looked at, and then one of the four
            // invoke words has to follow.
            if ends(i) {
                return AttributeRead::Wrong(i.min(last));
            }
            i += 1;
            match attribute_word(&at(i).kind) {
                Some(k) if is_invoke_word(k) => i += 1,
                _ => return AttributeRead::Wrong(i.min(last)),
            }
        }
        if at(i).kind != TokenKind::Kw(Kw::Eq) {
            return AttributeRead::Wrong(i.min(last));
        }
        i += 1;
        let value = match at(i).kind {
            // A minus sign is followed by an integer literal.
            TokenKind::Kw(Kw::Minus) => {
                if !matches!(at(i + 1).kind, TokenKind::IntLit | TokenKind::LongLit) {
                    return AttributeRead::Wrong((i + 1).min(last));
                }
                Value::Negated(i)
            }
            TokenKind::StrLit
            | TokenKind::IntLit
            | TokenKind::LongLit
            | TokenKind::SngLit
            | TokenKind::DblLit
            | TokenKind::CurLit
            | TokenKind::DateLit
            | TokenKind::Kw(Kw::True)
            | TokenKind::Kw(Kw::False) => Value::Literal(i),
            _ => Value::None(i),
        };
        AttributeRead::Applied { member, word: word_at, value }
    }

    /// `VB_Ext_KEY = "key", "value"`: two strings and the end of the line.
    fn read_extended_key(&self, toks: &[Token], member: Option<usize>, word_at: usize) -> AttributeRead {
        let last = toks.len() - 1;
        let kind = |i: usize| &toks[i.min(last)].kind;
        let wrong = |i: usize| AttributeRead::Wrong(i.min(last));
        let mut i = word_at + 1;
        for expected in [TokenKind::Kw(Kw::Eq), TokenKind::StrLit, TokenKind::Kw(Kw::Comma), TokenKind::StrLit] {
            if *kind(i) != expected {
                return wrong(i);
            }
            i += 1;
        }
        if !matches!(kind(i), TokenKind::Eol | TokenKind::Eof) {
            return wrong(i);
        }
        let value = Value::Pair(word_at + 2, word_at + 4);
        AttributeRead::Applied { member, word: word_at, value }
    }

    /// The tokens from the next one to the end of the logical line (the terminator last),
    /// without consuming them.
    pub(super) fn line_tokens(&mut self) -> Vec<Token> {
        let mut tokens: Vec<Token> = Vec::new();
        tokens.extend(self.peeked.take());
        tokens.extend(self.peeked2.take());
        loop {
            if tokens.last().is_some_and(|t| matches!(t.kind, TokenKind::Eol | TokenKind::Eof)) {
                break;
            }
            let next = match self.pushback.pop_front() {
                Some(token) => token,
                None => self.fetch_logical(),
            };
            tokens.push(next);
        }
        for token in tokens.iter().rev() {
            self.pushback.push_front(token.clone());
        }
        tokens
    }

    /// True when the tokens from `toks[at]` begin an `Attribute` line the reader acts on or
    /// finds wrong, as opposed to one it passes over.
    fn is_read_attribute_line(&self, toks: &[Token], at: usize) -> bool {
        toks.get(at).is_some_and(|t| t.kind == TokenKind::Kw(Kw::Attribute))
            && self.read_attribute(&toks[at + 1..]) != AttributeRead::Passed
    }

    /// True at `Else` or a line number that is followed, directly or after a colon, by an
    /// `Attribute` line that is read (see [`Self::is_read_attribute_line`]). Such a line is an
    /// attribute line whatever precedes it.
    pub(super) fn at_prefixed_attribute_line(&mut self) -> bool {
        if !matches!(self.peek().kind, TokenKind::Kw(Kw::Else) | TokenKind::IntLit | TokenKind::LongLit)
            || !self.logical_line_mentions_attribute()
        {
            return false;
        }
        let toks = self.line_tokens();
        let at = if toks.get(1).is_some_and(|t| t.kind == TokenKind::Kw(Kw::Colon)) { 2 } else { 1 };
        self.is_read_attribute_line(&toks, at)
    }

    /// The number of tokens before the `Attribute` word of the first `Attribute` line that is read
    /// and follows a colon in the rest of the line.
    pub(super) fn trailing_attribute_line_at(&mut self) -> Option<usize> {
        if !self.logical_line_mentions_attribute() {
            return None;
        }
        let toks = self.line_tokens();
        (0..toks.len())
            .find(|&i| toks[i].kind == TokenKind::Kw(Kw::Colon) && self.is_read_attribute_line(&toks, i + 1))
            .map(|colon| colon + 1)
    }

    /// True when the word `Attribute` is written somewhere in the logical line that starts at the
    /// next token; a cheap test of the text that lets the lines without it skip reading ahead.
    fn logical_line_mentions_attribute(&mut self) -> bool {
        let mut pos = self.peek().span.start as usize;
        loop {
            let rest = &self.src[pos.min(self.src.len())..];
            let len = rest.iter().position(|b| *b == b'\n' || *b == b'\r').unwrap_or(rest.len());
            let line = &rest[..len];
            if line.windows(9).any(|w| w.eq_ignore_ascii_case(b"attribute")) {
                return true;
            }
            // A line that ends in a continuation mark goes on in the next one.
            let end = pos + len;
            if !line.trim_ascii_end().ends_with(b"_") || end >= self.src.len() {
                return false;
            }
            pos = end + if self.src[end..].starts_with(b"\r\n") { 2 } else { 1 };
        }
    }

    /// True when the next tokens are the word `Attribute` that begins an attribute line, not a
    /// name followed by `As`.
    pub(super) fn at_attribute_word_start(&mut self) -> bool {
        self.peek().kind == TokenKind::Kw(Kw::Attribute) && self.peek2().kind != TokenKind::Kw(Kw::As)
    }

    /// True when the line that starts at the next token is an attribute line: it begins with the
    /// word `Attribute`, or an `Else`, a line number, a label or a colon comes before it.
    pub(super) fn at_attribute_line(&mut self) -> bool {
        self.at_attribute_word_start() || self.at_prefixed_attribute_line() || self.trailing_attribute_line_at().is_some()
    }

    /// In the body of a `Type` or `Enum`: reads the attribute line that starts at the next token,
    /// when the line is one that starts with the word, a line number, `Else` or a label. Returns
    /// whether it did, and, when the line is not one of those but holds an attribute line after a
    /// colon, whether the member before it is to be dropped.
    pub(super) fn take_body_attribute_line(&mut self, arena: &mut ExprArena) -> BodyAttribute {
        if self.at_attribute_word_start() {
            self.parse_attribute_stmt(arena);
            return BodyAttribute::Taken;
        }
        if self.at_prefixed_attribute_line() {
            self.advance();
            self.eat(TokenKind::Kw(Kw::Colon));
            self.parse_attribute_stmt(arena);
            return BodyAttribute::Taken;
        }
        match self.trailing_attribute_line_at() {
            // A label before the attribute line.
            Some(2) if self.peek2().kind == TokenKind::Kw(Kw::Colon) => {
                self.advance();
                self.advance();
                self.parse_attribute_stmt(arena);
                BodyAttribute::Taken
            }
            Some(_) => BodyAttribute::AfterMember,
            None => BodyAttribute::None,
        }
    }

    /// True when the `Attribute` word that starts at `start` is at the start of its line and is
    /// followed by a blank: such a line is taken out of the module's code before it is read.
    fn is_set_apart_attribute_line(&self, start: u32) -> bool {
        let start = start as usize;
        (start == 0 || matches!(self.src.get(start - 1), Some(b'\n' | b'\r'))) && starts_attribute_line(&self.src[start..])
    }

    /// The place an error in the set-apart attribute line at `start` is reported: the line that
    /// follows it and the attribute lines set apart with it.
    fn line_after_set_apart_attribute_lines(&self, start: u32) -> Span {
        let src = self.src;
        let mut pos = start as usize;
        loop {
            // The start of the line after this one.
            while pos < src.len() && src[pos] != b'\n' && src[pos] != b'\r' {
                pos += 1;
            }
            if src.get(pos) == Some(&b'\r') && src.get(pos + 1) == Some(&b'\n') {
                pos += 1;
            }
            if pos >= src.len() {
                return Span { start: src.len() as u32, len: 0 };
            }
            pos += 1;
            if !starts_attribute_line(&src[pos..]) {
                return Span { start: pos as u32, len: 0 };
            }
        }
    }

    /// Reports the syntax error of an `Attribute` line that starts at `start` and is wrong at `at`.
    fn report_attribute_error(&mut self, start: u32, at: Span) {
        let (span, set_apart) = if self.is_set_apart_attribute_line(start) {
            (self.line_after_set_apart_attribute_lines(start), true)
        } else {
            (at, false)
        };
        // Attribute lines set apart together are reported once, on the line that follows them.
        if set_apart && self.diagnostics.items().last().is_some_and(|d| d.code == ERR_SYNTAX_ERROR && d.span == span) {
            return;
        }
        self.diagnostics.push(ERR_SYNTAX_ERROR, span);
    }

    /// The literal node for the value token `toks[at]`.
    fn attribute_literal(&self, arena: &mut ExprArena, token: &Token, unescape: bool) -> NodeId {
        let lit = match token.kind {
            TokenKind::Kw(Kw::True) => AstLit::Bool(true),
            TokenKind::Kw(Kw::False) => AstLit::Bool(false),
            _ => match ast_lit_from_token(token) {
                AstLit::Str(s) if unescape => AstLit::Str(module_name::unescape(&s).into_boxed_str()),
                other => other,
            },
        };
        arena.alloc(ExprNode::Literal { lit })
    }

    /// `Attribute [member.]word = value` metadata line.
    pub(super) fn parse_attribute_stmt(&mut self, arena: &mut ExprArena) -> NodeId {
        let start = self.advance().span.start; // consume Attribute
        let toks = self.line_tokens();
        let read = self.read_attribute(&toks);
        for _ in 0..toks.len() {
            self.advance();
        }
        let empty = |arena: &mut ExprArena| arena.alloc(ExprNode::Block { stmts: Vec::new() });
        match read {
            AttributeRead::Passed => empty(arena),
            AttributeRead::TypeClause => {
                self.diagnostics.push_labeled(0x9c6f, toks[0].span, "New or type name");
                empty(arena)
            }
            AttributeRead::Wrong(at) => {
                self.report_attribute_error(start, toks[at].span);
                empty(arena)
            }
            AttributeRead::Applied { member, word, value } => {
                let target = member.map(|m| toks[m].sym.map(|s| s as u32).unwrap_or(0));
                let name = toks[word].sym.map(|s| s as u32).unwrap_or(0);
                let (Value::None(i) | Value::Literal(i) | Value::Negated(i) | Value::Pair(i, _)) = value;
                let value_span = toks[i].span;
                let values = match value {
                    Value::None(_) => Vec::new(),
                    Value::Literal(i) => vec![self.attribute_literal(arena, &toks[i], true)],
                    Value::Negated(i) => {
                        let operand = self.attribute_literal(arena, &toks[i + 1], false);
                        vec![arena.alloc(ExprNode::UnOp { op: UnOpKind::Neg, operand })]
                    }
                    Value::Pair(first, second) => vec![
                        self.attribute_literal(arena, &toks[first], false),
                        self.attribute_literal(arena, &toks[second], false),
                    ],
                };
                if target.is_none()
                    && toks[word].kind == TokenKind::Kw(Kw::VbExposed)
                    && values.first().is_some_and(|v| matches!(arena.get(*v), ExprNode::Literal { lit: AstLit::Bool(true) }))
                {
                    self.class_exposed = true;
                }
                let id = arena.alloc(ExprNode::Attribute { target, name, values });
                self.set_span(id, value_span);
                id
            }
        }
    }
}

/// True when `text` begins with `Attribute` and a blank, in any letter case.
fn starts_attribute_line(text: &[u8]) -> bool {
    text.len() >= 10 && text[..9].eq_ignore_ascii_case(b"Attribute") && text[9] == b' '
}
