; VB6/VBA Code Navigation Tags
; =============================
; Used for symbol navigation (go to definition, outline, etc.)

; Procedures - Sub
(sub_declaration
  name: (identifier) @name) @definition.function

; Procedures - Function
(function_declaration
  name: (identifier) @name) @definition.function

; Procedures - Property Get
(property_get_declaration
  name: (identifier) @name) @definition.function

; Procedures - Property Let
(property_let_declaration
  name: (identifier) @name) @definition.function

; Procedures - Property Set
(property_set_declaration
  name: (identifier) @name) @definition.function

; External API declarations
(declare_statement
  name: (identifier) @name) @definition.function

; Events
(event_declaration
  name: (identifier) @name) @definition.function

; Types (UDT)
(type_declaration
  name: (identifier) @name) @definition.type

; Enums
(enum_declaration
  name: (identifier) @name) @definition.type

; Module-level variables (Public/Private)
(variable_declaration
  (variable_declarator
    name: (identifier) @name)) @definition.variable

; Constants
(constant_declaration
  (constant_declarator
    name: (identifier) @name)) @definition.constant

; Type members
(type_member
  name: (identifier) @name) @definition.field

; Enum members
(enum_member
  name: (identifier) @name) @definition.constant

; Function/Sub calls (references)
(call_expression
  function: (identifier) @name) @reference.call

(call_expression
  function: (member_access_expression
    member: (identifier) @name)) @reference.call

; Implements (interface reference)
(implements_statement
  interface: (qualified_identifier) @name) @reference.type
