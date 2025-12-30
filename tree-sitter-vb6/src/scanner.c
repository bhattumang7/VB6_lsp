/**
 * External scanner for VB6/VBA tree-sitter grammar
 *
 * Handles tokens that cannot be expressed in the grammar DSL:
 * - Line continuations (underscore at end of line)
 * - Date literals (#date/time#)
 * - GUID literals ({xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx})
 * - Callable identifiers (identifiers that are NOT reserved keywords)
 */

#include "tree_sitter/parser.h"
#include <wctype.h>
#include <string.h>
#include <stdbool.h>
#include <ctype.h>

enum TokenType {
    LINE_CONTINUATION,
    DATE_LITERAL_TOKEN,
    GUID_LITERAL,
    FILE_NUMBER,
    CALLABLE_IDENTIFIER,
    LABEL_IDENTIFIER,
};

// Reserved keywords that cannot be used as callable names (case-insensitive)
// These are keywords that would conflict with call statements
static const char *RESERVED_KEYWORDS[] = {
    "public", "private", "friend", "global",  // Visibility modifiers
    "dim", "static", "const", "withevents",   // Declaration keywords
    "type", "enum", "class",                  // Type definition keywords
    "sub", "function", "property", "event",   // Procedure keywords
    "declare", "implements",                  // Other declaration keywords
    "if", "then", "else", "elseif", "end",   // Control flow
    "for", "to", "step", "next", "each", "in", // For loops
    "do", "loop", "while", "until", "wend",   // While/Do loops
    "select", "case",                          // Select case
    "with", "new",                             // With/New
    "exit", "return", "goto", "gosub", "on",  // Jump statements
    "set", "let",                              // Assignment keywords
    "call",                                    // Call keyword
    "redim", "preserve", "erase",             // Array keywords
    "option", "attribute",                    // Module keywords
    "true", "false", "nothing", "null", "empty", // Literals
    "and", "or", "not", "xor", "eqv", "imp", "is", "like", "mod", // Operators
    "as", "byval", "byref", "optional", "paramarray", // Parameter keywords
    "resume", "error",                         // Error handling
    "raiseevent",                              // Event raising
    "version", "begin",                        // Module header/config
    "open", "close", "input", "line", "print", "write", // File I/O
    "get", "put", "seek", "lock", "unlock", "width",    // File I/O
    "appactivate", "beep", "chdir", "chdrive",          // System statements
    "mkdir", "rmdir", "kill", "name", "filecopy",
    "load", "unload", "date", "time", "randomize",
    "lset", "rset", "mid", "stop", "sendkeys",
    "savepicture", "savesetting", "deletesetting",
    "setattr", "reset",
    "rem",                                      // Comment keyword
    NULL  // Sentinel
};

// Forward declarations
static bool scan_line_continuation(TSLexer *lexer);
static bool scan_date_literal(TSLexer *lexer);
static bool scan_guid_literal(TSLexer *lexer);
static bool scan_file_number(TSLexer *lexer);
static bool scan_hash_literal(TSLexer *lexer, const bool *valid_symbols);
static bool scan_callable_identifier(TSLexer *lexer);
static bool scan_label_identifier(TSLexer *lexer);
static void advance(TSLexer *lexer);
static void skip(TSLexer *lexer);
static bool is_hex_digit(int32_t c);
static bool is_identifier_start(int32_t c);
static bool is_identifier_char(int32_t c);
static bool is_reserved_keyword(const char *word, size_t len);
static bool is_preproc_keyword(const char *word, size_t len);
static int to_lower(int c);

/**
 * Create scanner state (none needed for VB6)
 */
void *tree_sitter_vb6_external_scanner_create(void) {
    return NULL;
}

/**
 * Destroy scanner state
 */
void tree_sitter_vb6_external_scanner_destroy(void *payload) {
    // Nothing to free
}

/**
 * Serialize scanner state (none needed)
 */
unsigned tree_sitter_vb6_external_scanner_serialize(void *payload, char *buffer) {
    return 0;
}

/**
 * Deserialize scanner state (none needed)
 */
void tree_sitter_vb6_external_scanner_deserialize(void *payload, const char *buffer, unsigned length) {
    // Nothing to restore
}

/**
 * Main scan function - called by tree-sitter when it needs an external token
 */
bool tree_sitter_vb6_external_scanner_scan(
    void *payload,
    TSLexer *lexer,
    const bool *valid_symbols
) {
    // Skip whitespace (but not newlines - those are significant)
    while (lexer->lookahead == ' ' || lexer->lookahead == '\t') {
        skip(lexer);
    }

    // Try to match each token type if it's valid in current context
    if (valid_symbols[LINE_CONTINUATION] && scan_line_continuation(lexer)) {
        lexer->result_symbol = LINE_CONTINUATION;
        return true;
    }

    if ((valid_symbols[DATE_LITERAL_TOKEN] || valid_symbols[FILE_NUMBER]) && lexer->lookahead == '#') {
        if (scan_hash_literal(lexer, valid_symbols)) {
            return true;
        }
    }

    if (valid_symbols[GUID_LITERAL] && lexer->lookahead == '{') {
        if (scan_guid_literal(lexer)) {
            lexer->result_symbol = GUID_LITERAL;
            return true;
        }
    }

    if (valid_symbols[LABEL_IDENTIFIER] && is_identifier_start(lexer->lookahead)) {
        if (scan_label_identifier(lexer)) {
            lexer->result_symbol = LABEL_IDENTIFIER;
            return true;
        }
    }

    // Try callable_identifier - an identifier that is NOT a reserved keyword
    if (valid_symbols[CALLABLE_IDENTIFIER] && is_identifier_start(lexer->lookahead)) {
        if (scan_callable_identifier(lexer)) {
            lexer->result_symbol = CALLABLE_IDENTIFIER;
            return true;
        }
    }

    return false;
}

/**
 * Advance the lexer and include character in token
 */
static void advance(TSLexer *lexer) {
    lexer->advance(lexer, false);
}

/**
 * Advance the lexer but don't include character in token
 */
static void skip(TSLexer *lexer) {
    lexer->advance(lexer, true);
}

/**
 * Check if character is a hex digit
 */
static bool is_hex_digit(int32_t c) {
    return (c >= '0' && c <= '9') ||
           (c >= 'a' && c <= 'f') ||
           (c >= 'A' && c <= 'F');
}

/**
 * Check if character can start an identifier
 */
static bool is_identifier_start(int32_t c) {
    return (c >= 'a' && c <= 'z') ||
           (c >= 'A' && c <= 'Z') ||
           c == '_';
}

/**
 * Check if character can be part of an identifier
 */
static bool is_identifier_char(int32_t c) {
    return is_identifier_start(c) ||
           (c >= '0' && c <= '9');
}

/**
 * Convert character to lowercase
 */
static int to_lower(int c) {
    if (c >= 'A' && c <= 'Z') {
        return c + ('a' - 'A');
    }
    return c;
}

/**
 * Check if word is a reserved keyword (case-insensitive)
 */
static bool is_reserved_keyword(const char *word, size_t len) {
    for (int i = 0; RESERVED_KEYWORDS[i] != NULL; i++) {
        const char *keyword = RESERVED_KEYWORDS[i];
        size_t kw_len = strlen(keyword);

        if (len != kw_len) {
            continue;
        }

        bool match = true;
        for (size_t j = 0; j < len; j++) {
            if (to_lower(word[j]) != keyword[j]) {
                match = false;
                break;
            }
        }

        if (match) {
            return true;
        }
    }

    return false;
}

static bool is_preproc_keyword(const char *word, size_t len) {
    static const char *PREPROC_KEYWORDS[] = {
        "if", "elseif", "else", "end", "const",
        NULL
    };

    for (int i = 0; PREPROC_KEYWORDS[i] != NULL; i++) {
        const char *keyword = PREPROC_KEYWORDS[i];
        size_t kw_len = strlen(keyword);

        if (len != kw_len) {
            continue;
        }

        bool match = true;
        for (size_t j = 0; j < len; j++) {
            if (to_lower(word[j]) != keyword[j]) {
                match = false;
                break;
            }
        }

        if (match) {
            return true;
        }
    }

    return false;
}

/**
 * Scan line continuation: underscore followed by optional whitespace and newline
 *
 * VB6 uses underscore at end of line to continue a logical line:
 *   Dim x As Long _
 *       , y As String
 */
static bool scan_line_continuation(TSLexer *lexer) {
    if (lexer->lookahead != '_') {
        return false;
    }

    // Mark potential start of token
    lexer->mark_end(lexer);
    advance(lexer);

    // After underscore, only whitespace and newline are allowed
    while (lexer->lookahead == ' ' || lexer->lookahead == '\t') {
        advance(lexer);
    }

    // Must be followed by newline (or EOF for edge case)
    if (lexer->lookahead == '\r') {
        advance(lexer);
        if (lexer->lookahead == '\n') {
            advance(lexer);
        }
        lexer->mark_end(lexer);
        return true;
    }

    if (lexer->lookahead == '\n') {
        advance(lexer);
        lexer->mark_end(lexer);
        return true;
    }

    // Not a line continuation - underscore is part of identifier
    return false;
}

/**
 * Scan date literal: #date# or #time# or #date time#
 *
 * Examples:
 *   #1/1/2024#
 *   #January 1, 2024#
 *   #12:30:00 PM#
 *   #1/1/2024 12:30:00 PM#
 *
 * The content between # symbols is quite flexible in VB6.
 */
static bool scan_date_literal(TSLexer *lexer) {
    if (lexer->lookahead != '#') {
        return false;
    }

    lexer->mark_end(lexer);
    advance(lexer);  // consume opening #

    // Scan until we find closing # or newline (invalid)
    bool has_content = false;

    while (lexer->lookahead != 0) {
        if (lexer->lookahead == '#') {
            if (!has_content) {
                // Empty date literal ## is invalid
                return false;
            }
            advance(lexer);  // consume closing #
            lexer->mark_end(lexer);
            return true;
        }

        if (lexer->lookahead == '\n' || lexer->lookahead == '\r') {
            // Newline before closing # - invalid
            return false;
        }

        // Date literals can contain various characters:
        // digits, letters, spaces, slashes, colons, commas, dashes
        int32_t c = lexer->lookahead;
        if (iswdigit(c) || iswspace(c) || iswalpha(c) ||
            c == '/' || c == ':' || c == ',' || c == '-' || c == '.') {
            has_content = true;
            advance(lexer);
        } else {
            // Unexpected character
            return false;
        }
    }

    // EOF without closing #
    return false;
}

/**
 * Scan GUID literal: {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
 *
 * Used in class modules and type libraries.
 */
static bool scan_guid_literal(TSLexer *lexer) {
    if (lexer->lookahead != '{') {
        return false;
    }

    advance(lexer);  // consume opening {

    // GUID format: 8-4-4-4-12 hex digits separated by dashes
    // Total: 36 characters (32 hex + 4 dashes)
    int hex_counts[] = {8, 4, 4, 4, 12};
    int num_groups = 5;

    for (int group = 0; group < num_groups; group++) {
        // Scan hex digits for this group
        for (int i = 0; i < hex_counts[group]; i++) {
            if (!is_hex_digit(lexer->lookahead)) {
                return false;
            }
            advance(lexer);
        }

        // Expect dash between groups (except after last group)
        if (group < num_groups - 1) {
            if (lexer->lookahead != '-') {
                return false;
            }
            advance(lexer);
        }
    }

    // Expect closing brace
    if (lexer->lookahead != '}') {
        return false;
    }
    advance(lexer);

    lexer->mark_end(lexer);
    return true;
}

/**
 * Scan file number: #number or #identifier, but not preprocessor keywords.
 */
static bool scan_file_number(TSLexer *lexer) {
    if (lexer->lookahead != '#') {
        return false;
    }

    lexer->mark_end(lexer);
    advance(lexer);

    char buffer[256];
    size_t len = 0;

    if (iswdigit(lexer->lookahead)) {
        while (iswdigit(lexer->lookahead) && len < sizeof(buffer) - 1) {
            buffer[len++] = (char)lexer->lookahead;
            advance(lexer);
        }
    } else if (is_identifier_start(lexer->lookahead)) {
        while (is_identifier_char(lexer->lookahead) && len < sizeof(buffer) - 1) {
            buffer[len++] = (char)lexer->lookahead;
            advance(lexer);
        }
    } else {
        return false;
    }

    if (len == 0) {
        return false;
    }

    buffer[len] = '\0';

    if (is_preproc_keyword(buffer, len)) {
        return false;
    }

    switch (lexer->lookahead) {
        case '/':
        case ':':
        case '-':
        case '.':
        case '#':
            return false;
        default:
            break;
    }

    lexer->mark_end(lexer);
    return true;
}

/**
 * Scan either a date literal (#...#) or a file number (#n or #name).
 */
static bool scan_hash_literal(TSLexer *lexer, const bool *valid_symbols) {
    if (lexer->lookahead != '#') {
        return false;
    }

    bool want_date = valid_symbols[DATE_LITERAL_TOKEN];
    bool want_file = valid_symbols[FILE_NUMBER];

    if (!want_date && !want_file) {
        return false;
    }

    lexer->mark_end(lexer);
    advance(lexer);  // consume opening #

    char buffer[256];
    size_t len = 0;
    bool file_candidate = false;
    bool file_valid = false;

    if (iswdigit(lexer->lookahead)) {
        file_candidate = true;
        while (iswdigit(lexer->lookahead) && len < sizeof(buffer) - 1) {
            buffer[len++] = (char)lexer->lookahead;
            advance(lexer);
        }
    } else if (is_identifier_start(lexer->lookahead)) {
        file_candidate = true;
        while (is_identifier_char(lexer->lookahead) && len < sizeof(buffer) - 1) {
            buffer[len++] = (char)lexer->lookahead;
            advance(lexer);
        }
    }

    if (file_candidate && len > 0) {
        buffer[len] = '\0';
        if (!is_preproc_keyword(buffer, len)) {
            file_valid = true;
            lexer->mark_end(lexer);
        }
    }

    if (want_date) {
        bool has_content = file_candidate && len > 0;

        while (lexer->lookahead != 0) {
            if (lexer->lookahead == '#') {
                if (!has_content) {
                    break;
                }
                advance(lexer);
                lexer->mark_end(lexer);
                lexer->result_symbol = DATE_LITERAL_TOKEN;
                return true;
            }

            if (lexer->lookahead == '\n' || lexer->lookahead == '\r') {
                break;
            }

            int32_t c = lexer->lookahead;
            if (!(iswdigit(c) || iswspace(c) || iswalpha(c) ||
                  c == '/' || c == ':' || c == ',' || c == '-' || c == '.')) {
                break;
            }

            has_content = true;
            advance(lexer);
        }
    }

    if (want_file && file_valid) {
        lexer->result_symbol = FILE_NUMBER;
        return true;
    }

    return false;
}

/**
 * Scan callable identifier: an identifier that is NOT a reserved keyword
 *
 * This is used for implicit call statements to prevent keywords like
 * "Public", "Private", etc. from being parsed as procedure names.
 *
 * IMPORTANT: We must mark_end at the START before advancing, so that if we
 * reject the token, the lexer can reset to the original position.
 */
static bool scan_callable_identifier(TSLexer *lexer) {
    if (!is_identifier_start(lexer->lookahead)) {
        return false;
    }

    // Mark start position - if we reject, lexer resets to here
    lexer->mark_end(lexer);

    // Buffer to store the identifier (max 256 chars should be enough)
    char buffer[256];
    size_t len = 0;

    // Scan the identifier
    while (is_identifier_char(lexer->lookahead) && len < sizeof(buffer) - 1) {
        buffer[len++] = (char)lexer->lookahead;
        advance(lexer);
    }
    buffer[len] = '\0';

    // Check if it's a reserved keyword
    if (is_reserved_keyword(buffer, len)) {
        // It's a keyword, not a callable identifier
        // Don't mark_end again - leave it at start so lexer resets
        return false;
    }

    // Mark the end of the identifier before peeking ahead
    lexer->mark_end(lexer);

    // Peek ahead to avoid stealing identifiers that are part of assignments,
    // labels, or member/index expressions.
    while (lexer->lookahead == ' ' || lexer->lookahead == '\t') {
        advance(lexer);
    }

    switch (lexer->lookahead) {
        case '=':
        case ':':
        case '.':
        case '!':
        case '(':
            return false;
        case '+':
        case '-':
            advance(lexer);
            if (lexer->lookahead == '=') {
                return false;
            }
            break;
        default:
            break;
    }

    // It's a valid callable identifier - mark end at current position
    return true;
}

/**
 * Scan label identifier: same as callable identifier but allows colon afterward.
 */
static bool scan_label_identifier(TSLexer *lexer) {
    if (!is_identifier_start(lexer->lookahead)) {
        return false;
    }

    lexer->mark_end(lexer);

    char buffer[256];
    size_t len = 0;

    while (is_identifier_char(lexer->lookahead) && len < sizeof(buffer) - 1) {
        buffer[len++] = (char)lexer->lookahead;
        advance(lexer);
    }
    buffer[len] = '\0';

    if (is_reserved_keyword(buffer, len)) {
        return false;
    }

    if (lexer->lookahead != ':') {
        return false;
    }

    lexer->mark_end(lexer);
    return true;
}
