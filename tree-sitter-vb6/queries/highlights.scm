; VB6/VBA Syntax Highlighting Queries
; =====================================

; Keywords - Control Flow
[
  "If" "Then" "Else" "ElseIf" "End"
  "Select" "Case"
  "For" "To" "Step" "Next" "Each" "In"
  "Do" "Loop" "While" "Until" "Wend"
  "With"
  "GoTo" "GoSub" "Return"
  "Exit" "Stop"
  "On" "Error" "Resume"
] @keyword.control

; Keywords - Declaration
[
  "Dim" "Private" "Public" "Friend" "Global" "Static"
  "Const" "Type" "Enum" "Event"
  "Declare" "Lib" "Alias"
  "Sub" "Function" "Property" "Get" "Let" "Set"
  "ByVal" "ByRef" "Optional" "ParamArray"
  "As" "New" "WithEvents"
  "Implements"
  "PtrSafe" "CDecl"
] @keyword

; Keywords - Operators
[
  "And" "Or" "Not" "Xor" "Eqv" "Imp"
  "Mod" "Like" "Is"
  "TypeOf" "AddressOf"
] @keyword.operator

; Keywords - File I/O
[
  "Open" "Close" "Reset"
  "Input" "Output" "Append" "Binary" "Random"
  "Access" "Read" "Write" "Shared" "Lock"
  "Print" "Line" "Get" "Put" "Seek"
  "Width" "Name"
] @keyword

; Keywords - DefType
[
  "DefBool" "DefByte" "DefInt" "DefLng"
  "DefCur" "DefSng" "DefDbl" "DefDec"
  "DefDate" "DefStr" "DefObj" "DefVar"
] @keyword.directive

; Keywords - System Statements
[
  "AppActivate" "Beep" "ChDir" "ChDrive"
  "MkDir" "RmDir" "Kill" "FileCopy"
  "Load" "Unload" "Date" "Time"
  "Randomize" "Error" "SendKeys"
  "SavePicture" "SaveSetting" "DeleteSetting" "SetAttr"
] @keyword

; Keywords - Other
[
  "Call" "RaiseEvent"
  "ReDim" "Preserve" "Erase"
  "LSet" "RSet" "Mid" "Mid$" "MidB" "MidB$"
  "Option" "Explicit" "Compare" "Base" "Module"
  "Attribute" "Version" "Class" "Begin"
  "Debug" "Assert"
] @keyword

; Built-in Types
[
  "Boolean" "Byte" "Currency" "Date" "Double"
  "Integer" "Long" "LongLong" "LongPtr"
  "Object" "Single" "String" "Variant" "Any"
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
