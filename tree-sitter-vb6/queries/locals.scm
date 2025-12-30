; VB6/VBA Local Variable Scoping Queries
; =======================================

; Scopes - procedures create new scopes
(sub_declaration) @scope
(function_declaration) @scope
(property_get_declaration) @scope
(property_let_declaration) @scope
(property_set_declaration) @scope

; Scopes - control flow blocks (for variable shadowing)
(for_statement) @scope
(for_each_statement) @scope
(with_statement) @scope

; Definitions - module level variables
(variable_declaration
  (variable_declarator
    name: (identifier) @definition.var))

; Definitions - constants
(constant_declaration
  (constant_declarator
    name: (identifier) @definition.constant))

; Definitions - parameters
(parameter
  name: (identifier) @definition.parameter)

; Definitions - for loop counter
(for_statement
  counter: (identifier) @definition.var)

; Definitions - for each element
(for_each_statement
  element: (identifier) @definition.var)

; Definitions - types
(type_declaration
  name: (identifier) @definition.type)

; Definitions - type members
(type_member
  name: (identifier) @definition.field)

; Definitions - enums
(enum_declaration
  name: (identifier) @definition.type)

; Definitions - enum members
(enum_member
  name: (identifier) @definition.constant)

; Definitions - procedures
(sub_declaration
  name: (identifier) @definition.function)

(function_declaration
  name: (identifier) @definition.function)

(property_get_declaration
  name: (identifier) @definition.function)

(property_let_declaration
  name: (identifier) @definition.function)

(property_set_declaration
  name: (identifier) @definition.function)

; Definitions - events
(event_declaration
  name: (identifier) @definition.function)

; Definitions - external declarations
(declare_statement
  name: (identifier) @definition.function)

; References - all identifier usages
(identifier) @reference
