; VB6/VBA Syntax Highlighting Queries
; =====================================

; Keywords - Control Flow
; NOTE: keyword tokens are aliased to lowercase by the grammar's ci() helper,
; so query strings must be lowercase to match (e.g. "if", not "If").
[
  "if" "then" "else" "elseif" "end"
  "select" "case"
  "for" "to" "step" "next" "each" "in"
  "do" "loop" "while" "until" "wend"
  "with"
  "goto" "gosub" "return"
  "exit" "stop"
  "on" "error" "resume"
] @keyword.control

; Keywords - Declaration
[
  "dim" "private" "public" "friend" "global" "static"
  "const" "type" "enum" "event"
  "declare" "lib" "alias"
  "sub" "function" "property" "get" "let" "set"
  "byval" "byref" "optional" "paramarray"
  "as" "new" "withevents"
  "implements"
] @keyword

; Keywords - Operators
[
  "and" "or" "not" "xor" "eqv" "imp"
  "mod" "like" "is"
  "typeof" "addressof"
] @keyword.operator

; Keywords - File I/O
[
  "open" "close" "reset"
  "input" "output" "append" "binary" "random"
  "access" "read" "write" "shared" "lock"
  "print" "line" "get" "put" "seek"
  "width" "name"
] @keyword

; Keywords - DefType
[
  "defbool" "defbyte" "defint" "deflng"
  "defcur" "defsng" "defdbl" "defdec"
  "defdate" "defstr" "defobj" "defvar"
] @keyword.directive

; Keywords - System Statements
[
  "appactivate" "beep" "chdir" "chdrive"
  "mkdir" "rmdir" "kill" "filecopy"
  "load" "unload" "date" "time"
  "randomize" "error" "sendkeys"
  "savepicture" "savesetting" "deletesetting" "setattr"
] @keyword

; Keywords - Other
[
  "call" "raiseevent"
  "redim" "preserve" "erase"
  "lset" "rset" "mid"
  "option" "explicit" "compare" "base" "module"
  "attribute" "version" "class" "begin"
] @keyword

; Built-in Types
[
  "boolean" "byte" "currency" "date" "double"
  "integer" "long" "longlong" "longptr"
  "object" "single" "string" "variant" "any"
] @type.builtin

; Preprocessor - hash symbol
"#" @keyword.directive

; Preprocessor directives
(preproc_const) @keyword.directive
(preproc_if) @keyword.directive
(preproc_elseif) @keyword.directive
(preproc_else) @keyword.directive
(preproc_if_statement) @keyword.directive
(preproc_elseif_statement) @keyword.directive
(preproc_else_statement) @keyword.directive

(preproc_const
  name: (identifier) @constant.definition)

; Procedures
(sub_declaration
  name: (identifier) @function.definition)

(function_declaration
  name: (identifier) @function.definition)

(property_declaration
  name: (identifier) @function.definition)

(declare_statement
  name: (identifier) @function.definition)

(event_statement
  name: (identifier) @function.definition)

; Function/Sub calls
(call_expression
  function: (identifier) @function.call)

(call_expression
  function: (member_expression
    member: (identifier) @function.call))

; Variables
(variable_declarator
  name: (identifier) @variable.definition)

(parameter
  name: (identifier) @variable.parameter)

; Constants
(constant_declarator
  name: (identifier) @constant.definition)

; Types
(type_declaration
  name: (identifier) @type.definition)

(type_member
  name: (identifier) @variable.field)

; Enums
(enum_declaration
  name: (identifier) @type.definition)

(enum_member
  name: (identifier) @constant.definition)

; Member access
(member_expression
  member: (identifier) @variable.member)

; Literals
(integer_literal) @number
(float_literal) @number

(string_literal) @string

(boolean_literal) @constant.builtin

(nothing_literal) @constant.builtin

(date_literal) @string.special

(color_literal) @number

(file_number) @number

; Comments
(comment) @comment

; Operators
[
  "+"
  "-"
  "*"
  "/"
  "\\"
  "^"
  "&"
  "="
  "<>"
  "<"
  ">"
  "<="
  ">="
  ":="
] @operator

; Punctuation
[
  "("
  ")"
] @punctuation.bracket

[
  ","
  ";"
  ":"
] @punctuation.delimiter

[
  "."
  "!"
] @punctuation.delimiter

; Labels
(label
  (identifier) @label)

(label
  (integer_literal) @label)

; Attributes
(attribute_statement
  (dotted_name) @attribute)

; Identifiers (fallback)
(identifier) @variable
