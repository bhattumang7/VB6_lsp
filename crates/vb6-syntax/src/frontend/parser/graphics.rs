//! The graphics method statements: `Print`, `Circle`, `Line`, `PSet` and `Scale` (and `?`, the
//! shorthand for `Print`), as bare statements and as members of an object.
//!
//! A statement is one of these when its word stands where a method is called on an object: after a
//! dot that follows something (so not a leading `With` dot), written without brackets, at the
//! start of a statement or after `Call`. The word then hands the rest of the statement to the
//! reader of its own operand list, which takes exactly one shape. Any other spelling of the word
//! (`o!Print`, `o.[Print]`, `.Print` inside `With`, the target of `Let`) is an ordinary member and
//! an ordinary call.

use super::*;

/// The words that name the graphics methods.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum GraphicsWord {
    Print,
    Circle,
    Line,
    PSet,
    Scale,
    /// `LineInput` written as one word: it belongs to the same class as the others, but no
    /// operands make a statement of it.
    LineInput,
}

impl GraphicsWord {
    /// The graphics word a token spells, if it spells one.
    pub(super) fn of(kind: &TokenKind) -> Option<Self> {
        match kind {
            TokenKind::Kw(Kw::Print) | TokenKind::Kw(Kw::Question) => Some(Self::Print),
            TokenKind::Kw(Kw::Circle) => Some(Self::Circle),
            TokenKind::Kw(Kw::Line) => Some(Self::Line),
            TokenKind::Kw(Kw::PSet) => Some(Self::PSet),
            TokenKind::Kw(Kw::Scale) => Some(Self::Scale),
            TokenKind::Kw(Kw::LineInput) => Some(Self::LineInput),
            _ => None,
        }
    }

    /// True for the words that draw: `Circle`, `Line`, `PSet` and `Scale`.
    pub(super) fn draws(self) -> bool {
        matches!(self, Self::Circle | Self::Line | Self::PSet | Self::Scale)
    }
}

/// What the chain being read at the start of a statement is the target of.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum ChainStart {
    /// A statement that begins with a name.
    Statement,
    /// The target of a `Call` statement.
    Call,
    /// The target of a `Let` or `Set` statement: a graphics word is no member here.
    Assignment,
}

/// Whether a diagnostic says that text is not well formed (an expected token that is missing, or a
/// syntax error) rather than that a well formed construct is not allowed.
fn is_syntax_diagnostic(d: &Diagnostic) -> bool {
    matches!(d.code, ERR_UNEXPECTED_TOKEN | ERR_SYNTAX_ERROR)
}

/// How a graphics statement is written.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Form {
    /// The word alone starts the statement.
    Bare,
    /// A member of an object, at the start of a statement.
    Method,
    /// A member of an object after `Call`.
    CalledMethod,
}

impl<'src> Parser<'src> {
    /// Called by the member chain of a statement once it has read a dot: the graphics word that
    /// the next token spells when the chain stops at it, or `None` when the token is an ordinary
    /// member name.
    pub(super) fn graphics_word_after_dot(&mut self, arena: &ExprArena, base: NodeId) -> Option<GraphicsWord> {
        // Something must stand before the dot: a leading `With` dot names a member of the
        // object of the block, an ordinary call.
        if matches!(arena.get(base), ExprNode::WithContext) {
            return None;
        }
        let word = GraphicsWord::of(&self.peek().kind)?;
        let next = self.peek2().kind.clone();
        match word {
            // Whatever follows `Print` is the list it prints.
            GraphicsWord::Print => {}
            // `o.Line = 1` assigns to a member that happens to be called Line.
            GraphicsWord::Line if next == TokenKind::Kw(Kw::Eq) => return None,
            // `o.Circle.x`: the word is a link of a longer chain.
            _ if next == TokenKind::Kw(Kw::Dot) => return None,
            _ => {}
        }
        Some(word)
    }

    /// True when the next token is the keyword `kw`, or that word written in brackets.
    fn at_keyword_word(&mut self, kw: Kw) -> bool {
        let (kind, sym) = {
            let tok = self.peek();
            (tok.kind.clone(), tok.sym)
        };
        match kind {
            TokenKind::Kw(k) => k == kw,
            TokenKind::Ident => sym.is_some_and(|s| self.scanner.keyword_of(s as u32) == Some(kw)),
            _ => false,
        }
    }


    /// At the start of a statement a graphics word written in brackets (`[Print] "x"`) is the
    /// statement of that word: the next token is turned into the keyword.
    pub(super) fn promote_bracketed_graphics_word(&mut self) {
        let tok = self.peek().clone();
        if tok.kind != TokenKind::Ident || !self.is_bracketed(&tok) {
            return;
        }
        let Some(kw) = tok.sym.and_then(|s| self.scanner.keyword_of(s as u32)) else { return };
        if kw == Kw::Question
            || GraphicsWord::of(&TokenKind::Kw(kw)).is_none_or(|word| word == GraphicsWord::LineInput)
        {
            return;
        }
        // Followed by `=`, `.`, `!` or `:` the word is a name or a label.
        if matches!(
            self.peek2().kind,
            TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang) | TokenKind::Kw(Kw::Colon)
        ) {
            return;
        }
        if let Some(peeked) = self.peeked.as_mut() {
            peeked.kind = TokenKind::Kw(kw);
        }
    }

    /// The symbol a member word is stored under: the question mark is the `Print` member.
    pub(super) fn graphics_member_symbol(&mut self, tok: &Token) -> u32 {
        match tok.sym {
            Some(sym) if tok.kind != TokenKind::Kw(Kw::Question) => sym as u32,
            _ => self.scanner.symbol_for("Print"),
        }
    }

    /// The statement a member chain ended in a graphics word starts: `callee` is the chain up to
    /// and including the word, `start` where the statement began.
    pub(super) fn finish_graphics_method(
        &mut self,
        arena: &mut ExprArena,
        callee: NodeId,
        word: GraphicsWord,
        chain: ChainStart,
        start: Span,
    ) -> NodeId {
        match chain {
            ChainStart::Assignment => {
                // The word cannot be assigned to or set.
                self.diagnostics.push(ERR_SYNTAX_ERROR, start);
                self.skip_line();
                callee
            }
            ChainStart::Statement => {
                let node = self.read_graphics_statement(arena, callee, word, Form::Method, start);
                self.graphics_stmts.insert(node.0, (word, true));
                node
            }
            ChainStart::Call => self.read_graphics_statement(arena, callee, word, Form::CalledMethod, start),
        }
    }

    /// A bare `Print`, `Circle`, `Line`, `PSet` or `Scale` statement: `word` has been read.
    pub(super) fn parse_bare_graphics_stmt(&mut self, arena: &mut ExprArena, word_tok: Token) -> NodeId {
        let word = GraphicsWord::of(&word_tok.kind).expect("a graphics word");
        let sym = self.graphics_member_symbol(&word_tok);
        let callee = arena.alloc(ExprNode::NameRef { sym, suffix: word_tok.type_suffix });
        self.set_span(callee, word_tok.span);
        let node = self.read_graphics_statement(arena, callee, word, Form::Bare, word_tok.span);
        self.graphics_stmts.insert(node.0, (word, false));
        node
    }

    /// Whether the token that follows a bare `Circle`, `Line`, `PSet` or `Scale` (the word has
    /// been read) makes the word a name (`Scale = 2`, `Line.x`) rather than the statement.
    pub(super) fn graphics_word_is_name(&mut self) -> bool {
        matches!(
            self.peek().kind,
            TokenKind::Kw(Kw::Eq) | TokenKind::Kw(Kw::Dot) | TokenKind::Kw(Kw::Bang)
        )
    }

    /// Reads the operands of the statement of `word` and builds its node. A statement that does
    /// not have the exact shape is a syntax error at `start`, and the rest of the line is skipped.
    fn read_graphics_statement(
        &mut self,
        arena: &mut ExprArena,
        callee: NodeId,
        word: GraphicsWord,
        form: Form,
        start: Span,
    ) -> NodeId {
        let mark = self.diagnostics.len();
        let operands = match word {
            GraphicsWord::Print => Some(self.read_print_operands(arena)),
            GraphicsWord::PSet => self.read_pset_operands(arena),
            GraphicsWord::Circle => self.read_circle_operands(arena),
            GraphicsWord::Line => self.read_line_operands(arena),
            GraphicsWord::Scale => self.read_scale_operands(arena),
            // No operands make a statement of it.
            GraphicsWord::LineInput => None,
        };
        let Some(args) = operands else {
            return self.fail_graphics_statement(callee, form, mark, start);
        };
        // A `Scale` without points has been checked to end the statement already.
        let ended = (word == GraphicsWord::Scale && args.is_empty()) || self.at_graphics_statement_end();
        if !ended || self.found_syntax_error_since(mark) {
            return self.fail_graphics_statement(callee, form, mark, start);
        }
        self.end_stmt();
        if form == Form::Bare && self.module_kind == ModuleKind::Standard && self.diagnostics.len() == mark {
            // A standard module has no surface to draw on or print to.
            self.diagnostics.push(ERR_METHOD_NOT_VALID_WITHOUT_OBJECT, start);
        }
        if args.is_empty() && form != Form::CalledMethod {
            return callee;
        }
        let args = arena.alloc(ExprNode::ArgList { args });
        arena.alloc(ExprNode::CallStmt { callee, args })
    }

    /// Reports a graphics statement that is not of the exact shape and gives up on the line. The
    /// error is a plain syntax error where the operands stopped making sense: at the word for a
    /// bare statement, at the first token an operand expression could not take or the reader
    /// could not place otherwise.
    fn fail_graphics_statement(&mut self, callee: NodeId, form: Form, mark: usize, start: Span) -> NodeId {
        let first = self.diagnostics.items()[mark..].iter().find(|d| is_syntax_diagnostic(d)).map(|d| d.span);
        let at = match first {
            Some(span) => span,
            None if form == Form::Bare => start,
            None => self.current_span(),
        };
        self.diagnostics.truncate(mark);
        self.diagnostics.push(ERR_SYNTAX_ERROR, if form == Form::Bare { start } else { at });
        self.skip_line();
        callee
    }

    /// Whether an operand expression read since `mark` was not well formed.
    fn found_syntax_error_since(&self, mark: usize) -> bool {
        self.diagnostics.items()[mark..].iter().any(is_syntax_diagnostic)
    }

    /// At the end of the line, or at the colon, the comment or the `Else` that follows a
    /// complete statement.
    fn at_graphics_statement_end(&mut self) -> bool {
        self.at_plain_statement_end() || (self.in_inline_if && self.peek().kind == TokenKind::Kw(Kw::Else))
    }

    /// At the end of the line, the colon or the comment.
    fn at_plain_statement_end(&mut self) -> bool {
        let kind = self.peek().kind.clone();
        self.peek().is_stmt_end()
            || matches!(kind, TokenKind::Kw(Kw::Colon) | TokenKind::Eof)
            || (self.rem_ends_statement() && kind == TokenKind::Kw(Kw::Rem))
    }

    /// The items of a `Print`, separated by `;` or `,`.
    fn read_print_operands(&mut self, arena: &mut ExprArena) -> Vec<NodeId> {
        let mut args = Vec::new();
        self.parse_print_method_items(arena, &mut args);
        args
    }

    /// `[Step] (x, y)`: one point. `Step` is read only where `step_ok`.
    fn read_graphics_point(&mut self, arena: &mut ExprArena, step_ok: bool) -> Option<NodeId> {
        let step = self.at_keyword_word(Kw::Step);
        if step {
            if !step_ok {
                return None;
            }
            self.advance();
        }
        if !self.eat(TokenKind::Kw(Kw::LParen)) {
            return None;
        }
        let x = self.try_parse_expr(arena)?;
        if !self.eat(TokenKind::Kw(Kw::Comma)) {
            return None;
        }
        let y = self.try_parse_expr(arena)?;
        if !self.eat(TokenKind::Kw(Kw::RParen)) {
            return None;
        }
        Some(arena.alloc(ExprNode::GraphicsPoint { step, x, y }))
    }

    /// One operand after a comma: an expression, or nothing where the next comma follows at once.
    /// `None` when the operand is missing and the statement ends instead.
    fn read_optional_operand(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        if self.peek().kind == TokenKind::Kw(Kw::Comma) {
            return Some(arena.alloc(ExprNode::MissingArg));
        }
        self.try_parse_expr(arena)
    }

    /// `PSet [Step] (x, y) [, color]`
    fn read_pset_operands(&mut self, arena: &mut ExprArena) -> Option<Vec<NodeId>> {
        let mut args = vec![self.read_graphics_point(arena, true)?];
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            args.push(self.try_parse_expr(arena)?);
        }
        Some(args)
    }

    /// `Circle [Step] (x, y), radius [, [color] [, [start] [, [end] [, aspect]]]]`
    fn read_circle_operands(&mut self, arena: &mut ExprArena) -> Option<Vec<NodeId>> {
        let mut args = vec![self.read_graphics_point(arena, true)?];
        if !self.eat(TokenKind::Kw(Kw::Comma)) {
            return None;
        }
        args.push(self.try_parse_expr(arena)?);
        // Color, start, end and aspect may each be left out.
        for _ in 0..4 {
            if !self.eat(TokenKind::Kw(Kw::Comma)) {
                break;
            }
            args.push(self.read_optional_operand(arena)?);
        }
        Some(args)
    }

    /// `Line [[Step] (x, y)] - [Step] (x, y) [, [color] [, B | BF | B F]]`
    fn read_line_operands(&mut self, arena: &mut ExprArena) -> Option<Vec<NodeId>> {
        let mut args = Vec::new();
        if self.peek().kind == TokenKind::Kw(Kw::Minus) {
            // The first point is left out.
            args.push(arena.alloc(ExprNode::MissingArg));
        } else {
            args.push(self.read_graphics_point(arena, true)?);
        }
        if !self.eat(TokenKind::Kw(Kw::Minus)) {
            return None;
        }
        args.push(self.read_graphics_point(arena, true)?);
        if self.eat(TokenKind::Kw(Kw::Comma)) {
            args.push(self.read_optional_operand(arena)?);
            if self.eat(TokenKind::Kw(Kw::Comma)) {
                args.push(self.read_line_box(arena)?);
            }
        }
        Some(args)
    }

    /// The `B`, `BF` or `B F` that makes a line a box.
    fn read_line_box(&mut self, arena: &mut ExprArena) -> Option<NodeId> {
        let span = self.current_span();
        let filled = if self.at_keyword_word(Kw::BF) {
            self.advance();
            true
        } else if self.at_keyword_word(Kw::B) {
            self.advance();
            if self.at_keyword_word(Kw::F) {
                self.advance();
                true
            } else {
                false
            }
        } else {
            return None;
        };
        let id = arena.alloc(ExprNode::LineBox { filled });
        self.set_span(id, span);
        Some(id)
    }

    /// `Scale` alone, or `Scale (x, y) - (x, y)`; `Step` has no place in either point.
    fn read_scale_operands(&mut self, arena: &mut ExprArena) -> Option<Vec<NodeId>> {
        // Without points the statement ends where the line, a colon or a comment mark does:
        // not at a `Rem`, and not at an `Else`.
        let kind = self.peek().kind.clone();
        if self.peek().is_stmt_end() || matches!(kind, TokenKind::Kw(Kw::Colon) | TokenKind::Eof) {
            return Some(Vec::new());
        }
        let mut args = vec![self.read_graphics_point(arena, false)?];
        if !self.eat(TokenKind::Kw(Kw::Minus)) {
            return None;
        }
        args.push(self.read_graphics_point(arena, false)?);
        Some(args)
    }
}
