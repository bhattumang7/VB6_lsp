//! VB6 Lexer/Tokenizer
//!
//! Tokenizes VB6 source code for parsing.

#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Keywords
    Keyword(String),

    // Identifiers
    Identifier(String),

    // Literals
    StringLiteral(String),
    NumberLiteral(String),
    BooleanLiteral(bool),

    // Operators
    Plus,
    Minus,
    Star,
    Slash,
    Backslash,
    Mod,
    Power,
    Ampersand,
    Equals,
    NotEquals,
    LessThan,
    GreaterThan,
    LessThanOrEqual,
    GreaterThanOrEqual,
    And,
    Or,
    Not,
    Xor,

    // Delimiters
    LeftParen,
    RightParen,
    Comma,
    Dot,
    Colon,
    Semicolon,

    // Special
    Comment(String),
    Newline,
    LineContinuation,
    Eof,
}

pub struct Lexer {
    source: Vec<char>,
    position: usize,
    current_char: Option<char>,
}

impl Lexer {
    pub fn new(source: &str) -> Self {
        let source: Vec<char> = source.chars().collect();
        let current_char = source.get(0).copied();

        Self {
            source,
            position: 0,
            current_char,
        }
    }

    fn advance(&mut self) {
        self.position += 1;
        self.current_char = self.source.get(self.position).copied();
    }

    fn peek(&self, offset: usize) -> Option<char> {
        self.source.get(self.position + offset).copied()
    }

    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.current_char {
            if ch == ' ' || ch == '\t' {
                self.advance();
            } else {
                break;
            }
        }
    }

    fn read_string(&mut self, quote: char) -> String {
        let mut result = String::new();
        self.advance(); // Skip opening quote

        while let Some(ch) = self.current_char {
            if ch == quote {
                // Check for doubled quote (escape sequence in VB6)
                if self.peek(1) == Some(quote) {
                    result.push(quote);
                    self.advance();
                    self.advance();
                } else {
                    self.advance(); // Skip closing quote
                    break;
                }
            } else {
                result.push(ch);
                self.advance();
            }
        }

        result
    }

    fn read_number(&mut self) -> String {
        let mut result = String::new();

        while let Some(ch) = self.current_char {
            if ch.is_ascii_digit() || ch == '.' {
                result.push(ch);
                self.advance();
            } else if ch == 'E' || ch == 'e' {
                // Scientific notation
                result.push(ch);
                self.advance();
                if let Some(sign @ ('+' | '-')) = self.current_char {
                    result.push(sign);
                    self.advance();
                }
            } else {
                break;
            }
        }

        result
    }

    fn read_identifier(&mut self) -> String {
        let mut result = String::new();

        while let Some(ch) = self.current_char {
            if ch.is_alphanumeric() || ch == '_' {
                result.push(ch);
                self.advance();
            } else {
                break;
            }
        }

        result
    }

    fn read_comment(&mut self) -> String {
        let mut result = String::new();

        while let Some(ch) = self.current_char {
            if ch == '\n' || ch == '\r' {
                break;
            }
            result.push(ch);
            self.advance();
        }

        result
    }

    pub fn next_token(&mut self) -> Token {
        self.skip_whitespace();

        match self.current_char {
            None => Token::Eof,

            Some('\n') | Some('\r') => {
                self.advance();
                Token::Newline
            }

            Some('\'') => {
                let comment = self.read_comment();
                Token::Comment(comment)
            }

            Some('"') => {
                let string = self.read_string('"');
                Token::StringLiteral(string)
            }

            Some(ch) if ch.is_ascii_digit() => {
                let number = self.read_number();
                Token::NumberLiteral(number)
            }

            Some(ch) if ch.is_alphabetic() || ch == '_' => {
                let identifier = self.read_identifier();
                let upper = identifier.to_uppercase();

                // Check for keywords
                if is_keyword(&upper) {
                    Token::Keyword(upper)
                } else if upper == "TRUE" {
                    Token::BooleanLiteral(true)
                } else if upper == "FALSE" {
                    Token::BooleanLiteral(false)
                } else {
                    Token::Identifier(identifier)
                }
            }

            Some('+') => {
                self.advance();
                Token::Plus
            }

            Some('-') => {
                self.advance();
                Token::Minus
            }

            Some('*') => {
                self.advance();
                Token::Star
            }

            Some('/') => {
                self.advance();
                Token::Slash
            }

            Some('\\') => {
                self.advance();
                Token::Backslash
            }

            Some('^') => {
                self.advance();
                Token::Power
            }

            Some('&') => {
                self.advance();
                Token::Ampersand
            }

            Some('=') => {
                self.advance();
                Token::Equals
            }

            Some('<') => {
                self.advance();
                if self.current_char == Some('>') {
                    self.advance();
                    Token::NotEquals
                } else if self.current_char == Some('=') {
                    self.advance();
                    Token::LessThanOrEqual
                } else {
                    Token::LessThan
                }
            }

            Some('>') => {
                self.advance();
                if self.current_char == Some('=') {
                    self.advance();
                    Token::GreaterThanOrEqual
                } else {
                    Token::GreaterThan
                }
            }

            Some('(') => {
                self.advance();
                Token::LeftParen
            }

            Some(')') => {
                self.advance();
                Token::RightParen
            }

            Some(',') => {
                self.advance();
                Token::Comma
            }

            Some('.') => {
                self.advance();
                Token::Dot
            }

            Some(':') => {
                self.advance();
                Token::Colon
            }

            Some('_') if self.peek(1).map(|c| c.is_whitespace()).unwrap_or(false) => {
                self.advance();
                Token::LineContinuation
            }

            Some(ch) => {
                // Unknown character, skip it
                self.advance();
                self.next_token()
            }
        }
    }

    pub fn tokenize(&mut self) -> Vec<Token> {
        let mut tokens = Vec::new();

        loop {
            let token = self.next_token();
            if token == Token::Eof {
                tokens.push(token);
                break;
            }
            tokens.push(token);
        }

        tokens
    }
}

fn is_keyword(word: &str) -> bool {
    matches!(
        word,
        "AS"
            | "AND"
            | "BOOLEAN"
            | "BYREF"
            | "BYTE"
            | "BYVAL"
            | "CALL"
            | "CASE"
            | "CLASS"
            | "CONST"
            | "CURRENCY"
            | "DATE"
            | "DECLARE"
            | "DIM"
            | "DO"
            | "DOUBLE"
            | "EACH"
            | "ELSE"
            | "ELSEIF"
            | "END"
            | "ENUM"
            | "EVENT"
            | "EXIT"
            | "FOR"
            | "FRIEND"
            | "FUNCTION"
            | "GET"
            | "GLOBAL"
            | "GOSUB"
            | "GOTO"
            | "IF"
            | "IMPLEMENTS"
            | "IN"
            | "INTEGER"
            | "IS"
            | "LET"
            | "LIB"
            | "LIKE"
            | "LONG"
            | "LOOP"
            | "MOD"
            | "MODULE"
            | "NEW"
            | "NEXT"
            | "NOT"
            | "NOTHING"
            | "NULL"
            | "OBJECT"
            | "ON"
            | "OPTION"
            | "OPTIONAL"
            | "OR"
            | "PRESERVE"
            | "PRIVATE"
            | "PROPERTY"
            | "PUBLIC"
            | "RAISEEVENT"
            | "REDIM"
            | "REM"
            | "RESUME"
            | "RETURN"
            | "SELECT"
            | "SET"
            | "SINGLE"
            | "STATIC"
            | "STEP"
            | "STOP"
            | "STRING"
            | "SUB"
            | "THEN"
            | "TO"
            | "TYPE"
            | "TYPEOF"
            | "UNTIL"
            | "VARIANT"
            | "WEND"
            | "WHILE"
            | "WITH"
            | "WITHEVENTS"
            | "XOR"
    )
}
