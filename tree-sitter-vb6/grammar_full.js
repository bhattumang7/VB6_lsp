/**
 * @file Complete VB6/VBA grammar for tree-sitter
 * @author VB6 LSP Team
 * @license MIT
 *
 * Full VB6 grammar with comprehensive language support
 */

/// <reference types="tree-sitter-cli/dsl" />
// @ts-check

// Operator precedence (lowest to highest)
const PREC = {
  IMP: 0,
  EQV: 1,
  XOR: 2,
  OR: 3,
  AND: 4,
  NOT: 5,
  COMPARE: 6,
  CONCAT: 7,
  ADD: 8,
  MOD: 9,
  IDIV: 10,
  MUL: 11,
  UNARY: 12,
  POW: 13,
  CALL: 14,
};

// Case-insensitive keyword helper
const ci = (word) => alias(new RegExp(word, 'i'), word);

module.exports = grammar({
  name: 'vb6',

  extras: $ => [
    /[ \t]+/,
    /[ \t]*_[ \t]*\r?\n/,  // Line continuation
  ],

  externals: $ => [
    $.line_continuation,
    $.date_literal_token,
    $.guid_literal,
  ],

  word: $ => $.identifier,

  conflicts: $ => [
    [$._expression, $.call_expression],
    [$._expression, $._lvalue],
    [$._expression, $.call_expression, $._lvalue],
    [$._expression, $.call_expression, $.dotted_name],
    [$.implicit_call_stmt, $.call_expression],
  ],

  rules: {
    // ============================================
    // ROOT
    // ============================================
    source_file: $ => repeat($._module_element),

    _module_element: $ => choice(
      $.module_header,
      $.module_config,
      $.option_statement,
      $.attribute_statement,
      $.variable_declaration,
      $.constant_declaration,
      $.type_declaration,
      $.enum_declaration,
      $.declare_statement,
      $.event_statement,
      $.deftype_statement,
      $.sub_declaration,
      $.function_declaration,
      $.property_declaration,
      $.preproc_const,
      $.preproc_if,
      $.implements_statement,
      $.comment,
      $._newline,
    ),

    // ============================================
    // MODULE STRUCTURE
    // ============================================
    module_header: $ => seq(
      ci('version'),
      $._expression,
      optional(seq(ci('class'))),
      $._terminator,
    ),

    module_config: $ => seq(
      ci('begin'),
      $._terminator,
      repeat($.module_config_element),
      ci('end'),
      $._terminator,
    ),

    module_config_element: $ => seq(
      $.dotted_name,
      '=',
      $._expression,
      $._terminator,
    ),

    // ============================================
    // CONDITIONAL COMPILATION
    // ============================================
    preproc_const: $ => seq(
      '#',
      ci('const'),
      field('name', $.identifier),
      '=',
      field('value', $._preproc_expression),
      $._terminator,
    ),

    preproc_if: $ => seq(
      '#',
      ci('if'),
      field('condition', $._preproc_expression),
      ci('then'),
      $._terminator,
      optional($._preproc_body),
      repeat($.preproc_elseif),
      optional($.preproc_else),
      '#',
      ci('end'),
      ci('if'),
      $._terminator,
    ),

    preproc_elseif: $ => seq(
      '#',
      ci('elseif'),
      field('condition', $._preproc_expression),
      ci('then'),
      $._terminator,
      optional($._preproc_body),
    ),

    preproc_else: $ => seq(
      '#',
      ci('else'),
      $._terminator,
      optional($._preproc_body),
    ),

    _preproc_body: $ => repeat1($._module_element),

    _preproc_expression: $ => choice(
      $.preproc_binary_expression,
      $.preproc_unary_expression,
      $.preproc_parenthesized,
      $.identifier,
      $.integer_literal,
      $.boolean_literal,
    ),

    preproc_binary_expression: $ => choice(
      prec.left(PREC.OR, seq($._preproc_expression, ci('or'), $._preproc_expression)),
      prec.left(PREC.AND, seq($._preproc_expression, ci('and'), $._preproc_expression)),
      prec.left(PREC.COMPARE, seq($._preproc_expression, choice('=', '<>', '<', '>', '<=', '>='), $._preproc_expression)),
    ),

    preproc_unary_expression: $ => prec(PREC.NOT, seq(ci('not'), $._preproc_expression)),

    preproc_parenthesized: $ => seq('(', $._preproc_expression, ')'),

    // ============================================
    // OPTION STATEMENTS
    // ============================================
    option_statement: $ => seq(
      ci('option'),
      choice(
        ci('explicit'),
        seq(ci('base'), choice('0', '1')),
        seq(ci('compare'), choice(ci('binary'), ci('text'), ci('database'))),
        seq(ci('private'), ci('module')),
      ),
      $._terminator,
    ),

    // ============================================
    // ATTRIBUTE STATEMENTS
    // ============================================
    attribute_statement: $ => seq(
      ci('attribute'),
      $.dotted_name,
      '=',
      $._expression,
      repeat(seq(',', $._expression)),
      $._terminator,
    ),

    // ============================================
    // DEFTYPE STATEMENTS
    // ============================================
    deftype_statement: $ => seq(
      choice(
        ci('defbool'), ci('defbyte'), ci('defint'), ci('deflng'),
        ci('defcur'), ci('defsng'), ci('defdbl'), ci('defdec'),
        ci('defdate'), ci('defstr'), ci('defobj'), ci('defvar')
      ),
      $.letter_range,
      repeat(seq(',', $.letter_range)),
      $._terminator,
    ),

    letter_range: $ => seq(
      $.identifier,
      optional(seq('-', $.identifier)),
    ),

    // ============================================
    // DECLARATIONS
    // ============================================
    variable_declaration: $ => seq(
      optional($._visibility),
      optional(ci('static')),
      choice(ci('dim'), $._visibility_only),
      optional(ci('withevents')),
      $.variable_list,
      $._terminator,
    ),

    _visibility_only: $ => choice(), // placeholder, visibility already captured

    variable_list: $ => seq(
      $.variable_declarator,
      repeat(seq(',', $.variable_declarator)),
    ),

    variable_declarator: $ => seq(
      field('name', $.identifier),
      optional($.type_hint),
      optional($.array_bounds),
      optional($.as_clause),
    ),

    array_bounds: $ => seq(
      '(',
      optional(seq(
        $.subscript,
        repeat(seq(',', $.subscript)),
      )),
      ')',
    ),

    subscript: $ => seq(
      optional(seq($._expression, ci('to'))),
      $._expression,
    ),

    as_clause: $ => seq(
      ci('as'),
      optional(ci('new')),
      field('type', $._type),
      optional($.field_length),
    ),

    field_length: $ => seq(
      '*',
      choice($.integer_literal, $.identifier),
    ),

    _type: $ => choice(
      $.builtin_type,
      $.dotted_name,
      seq($.dotted_name, '(', ')'),  // Array type
    ),

    builtin_type: $ => choice(
      ci('boolean'), ci('byte'), ci('currency'), ci('date'),
      ci('double'), ci('integer'), ci('long'), ci('longlong'),
      ci('longptr'), ci('object'), ci('single'), ci('string'),
      ci('variant'), ci('any'), ci('collection'),
    ),

    constant_declaration: $ => seq(
      optional($._visibility),
      ci('const'),
      $.constant_declarator,
      repeat(seq(',', $.constant_declarator)),
      $._terminator,
    ),

    constant_declarator: $ => seq(
      field('name', $.identifier),
      optional($.type_hint),
      optional($.as_clause),
      '=',
      field('value', $._expression),
    ),

    type_declaration: $ => seq(
      optional($._visibility),
      ci('type'),
      field('name', $.identifier),
      $._terminator,
      repeat($.type_member),
      ci('end'),
      ci('type'),
      $._terminator,
    ),

    type_member: $ => seq(
      field('name', $.identifier),
      optional($.array_bounds),
      $.as_clause,
      $._terminator,
    ),

    enum_declaration: $ => seq(
      optional($._visibility),
      ci('enum'),
      field('name', $.identifier),
      $._terminator,
      repeat($.enum_member),
      ci('end'),
      ci('enum'),
      $._terminator,
    ),

    enum_member: $ => seq(
      field('name', $.identifier),
      optional(seq('=', field('value', $._expression))),
      $._terminator,
    ),

    declare_statement: $ => seq(
      optional($._visibility),
      ci('declare'),
      choice(
        seq(ci('sub'), field('name', $.identifier)),
        seq(ci('function'), field('name', $.identifier), optional($.type_hint)),
      ),
      ci('lib'),
      $.string_literal,
      optional(seq(ci('alias'), $.string_literal)),
      optional($.parameter_list),
      optional($.as_clause),
      $._terminator,
    ),

    event_statement: $ => seq(
      optional($._visibility),
      ci('event'),
      field('name', $.identifier),
      $.parameter_list,
      $._terminator,
    ),

    implements_statement: $ => seq(
      ci('implements'),
      $.dotted_name,
      $._terminator,
    ),

    // ============================================
    // PROCEDURES
    // ============================================
    sub_declaration: $ => seq(
      optional($._visibility),
      optional(ci('static')),
      ci('sub'),
      field('name', $.identifier),
      optional($.parameter_list),
      $._terminator,
      optional($.block),
      ci('end'),
      ci('sub'),
      $._terminator,
    ),

    function_declaration: $ => seq(
      optional($._visibility),
      optional(ci('static')),
      ci('function'),
      field('name', $.identifier),
      optional($.type_hint),
      optional($.parameter_list),
      optional($.as_clause),
      $._terminator,
      optional($.block),
      ci('end'),
      ci('function'),
      $._terminator,
    ),

    property_declaration: $ => seq(
      optional($._visibility),
      optional(ci('static')),
      ci('property'),
      field('accessor', choice(ci('get'), ci('let'), ci('set'))),
      field('name', $.identifier),
      optional($.type_hint),
      optional($.parameter_list),
      optional($.as_clause),
      $._terminator,
      optional($.block),
      ci('end'),
      ci('property'),
      $._terminator,
    ),

    parameter_list: $ => seq(
      '(',
      optional(seq(
        $.parameter,
        repeat(seq(',', $.parameter)),
      )),
      ')',
    ),

    parameter: $ => seq(
      optional(ci('optional')),
      optional(choice(ci('byval'), ci('byref'), ci('paramarray'))),
      field('name', $.identifier),
      optional($.type_hint),
      optional(seq('(', ')')),  // Array param
      optional($.as_clause),
      optional(seq('=', field('default', $._expression))),
    ),

    // ============================================
    // STATEMENTS
    // ============================================
    block: $ => repeat1($._statement),

    _statement: $ => choice(
      $.assignment_statement,
      $.set_statement,
      $.if_statement,
      $.for_statement,
      $.for_each_statement,
      $.do_statement,
      $.while_statement,
      $.with_statement,
      $.select_statement,
      $.exit_statement,
      $.return_statement,
      $.goto_statement,
      $.gosub_statement,
      $.on_error_statement,
      $.on_goto_statement,
      $.on_gosub_statement,
      $.call_statement,
      $.redim_statement,
      $.erase_statement,
      $.raiseevent_statement,
      // File I/O
      $.open_statement,
      $.close_statement,
      $.input_statement,
      $.line_input_statement,
      $.print_statement,
      $.write_statement,
      $.get_statement,
      $.put_statement,
      $.seek_statement,
      $.lock_statement,
      $.unlock_statement,
      $.width_statement,
      // System statements
      $.app_activate_statement,
      $.beep_statement,
      $.chdir_statement,
      $.chdrive_statement,
      $.mkdir_statement,
      $.rmdir_statement,
      $.kill_statement,
      $.name_statement,
      $.filecopy_statement,
      $.load_statement,
      $.unload_statement,
      $.date_statement,
      $.time_statement,
      $.randomize_statement,
      $.lset_statement,
      $.rset_statement,
      $.mid_statement,
      $.error_statement,
      $.resume_statement,
      $.stop_statement,
      $.end_statement,
      $.sendkeys_statement,
      $.savepicture_statement,
      $.savesetting_statement,
      $.deletesetting_statement,
      $.setattr_statement,
      $.reset_statement,
      // Preprocessor in statements
      $.preproc_if_statement,
      $.label,
      $.comment,
      $._newline,
    ),

    // Conditional compilation within statements
    preproc_if_statement: $ => seq(
      '#',
      ci('if'),
      field('condition', $._preproc_expression),
      ci('then'),
      $._terminator,
      optional($.block),
      repeat($.preproc_elseif_statement),
      optional($.preproc_else_statement),
      '#',
      ci('end'),
      ci('if'),
      $._terminator,
    ),

    preproc_elseif_statement: $ => seq(
      '#',
      ci('elseif'),
      field('condition', $._preproc_expression),
      ci('then'),
      $._terminator,
      optional($.block),
    ),

    preproc_else_statement: $ => seq(
      '#',
      ci('else'),
      $._terminator,
      optional($.block),
    ),

    label: $ => seq(
      choice($.identifier, $.integer_literal),
      ':',
    ),

    assignment_statement: $ => seq(
      optional(ci('let')),
      field('target', $._lvalue),
      choice('=', '+=', '-='),
      field('value', $._expression),
      $._terminator,
    ),

    set_statement: $ => seq(
      ci('set'),
      field('target', $._lvalue),
      '=',
      optional(ci('new')),
      field('value', $._expression),
      $._terminator,
    ),

    call_statement: $ => seq(
      optional(ci('call')),
      choice($.call_expression, $.implicit_call_stmt),
      $._terminator,
    ),

    implicit_call_stmt: $ => seq(
      $.identifier,
      optional($.argument_list_no_parens),
    ),

    argument_list_no_parens: $ => seq(
      $._argument,
      repeat(seq(',', optional($._argument))),
    ),

    if_statement: $ => seq(
      ci('if'),
      field('condition', $._expression),
      ci('then'),
      choice(
        // Block If
        seq(
          $._terminator,
          optional($.block),
          repeat($.elseif_clause),
          optional($.else_clause),
          ci('end'),
          ci('if'),
          $._terminator,
        ),
        // Single-line If
        seq(
          $._inline_statement,
          optional(seq(ci('else'), $._inline_statement)),
          $._terminator,
        ),
      ),
    ),

    elseif_clause: $ => seq(
      ci('elseif'),
      field('condition', $._expression),
      ci('then'),
      $._terminator,
      optional($.block),
    ),

    else_clause: $ => seq(
      ci('else'),
      $._terminator,
      optional($.block),
    ),

    _inline_statement: $ => choice(
      seq(optional(ci('let')), $._lvalue, '=', $._expression),
      seq(ci('set'), $._lvalue, '=', optional(ci('new')), $._expression),
      seq(ci('goto'), choice($.identifier, $.integer_literal)),
      $.call_expression,
    ),

    for_statement: $ => seq(
      ci('for'),
      field('counter', $.identifier),
      '=',
      field('start', $._expression),
      ci('to'),
      field('end', $._expression),
      optional(seq(ci('step'), field('step', $._expression))),
      $._terminator,
      optional($.block),
      ci('next'),
      optional($.identifier),
      $._terminator,
    ),

    for_each_statement: $ => seq(
      ci('for'),
      ci('each'),
      field('element', $.identifier),
      ci('in'),
      field('collection', $._expression),
      $._terminator,
      optional($.block),
      ci('next'),
      optional($.identifier),
      $._terminator,
    ),

    do_statement: $ => choice(
      // Do...Loop
      seq(
        ci('do'),
        optional(seq(choice(ci('while'), ci('until')), $._expression)),
        $._terminator,
        optional($.block),
        ci('loop'),
        optional(seq(choice(ci('while'), ci('until')), $._expression)),
        $._terminator,
      ),
    ),

    while_statement: $ => seq(
      ci('while'),
      field('condition', $._expression),
      $._terminator,
      optional($.block),
      ci('wend'),
      $._terminator,
    ),

    with_statement: $ => seq(
      ci('with'),
      optional(ci('new')),
      field('object', $._expression),
      $._terminator,
      optional($.block),
      ci('end'),
      ci('with'),
      $._terminator,
    ),

    select_statement: $ => seq(
      ci('select'),
      ci('case'),
      field('test', $._expression),
      $._terminator,
      repeat($.case_clause),
      optional($.case_else_clause),
      ci('end'),
      ci('select'),
      $._terminator,
    ),

    case_clause: $ => seq(
      ci('case'),
      $.case_values,
      $._terminator,
      optional($.block),
    ),

    case_else_clause: $ => seq(
      ci('case'),
      ci('else'),
      $._terminator,
      optional($.block),
    ),

    case_values: $ => seq(
      $._case_value,
      repeat(seq(',', $._case_value)),
    ),

    _case_value: $ => choice(
      seq($._expression, ci('to'), $._expression),
      seq(ci('is'), $._compare_op, $._expression),
      $._expression,
    ),

    exit_statement: $ => seq(
      ci('exit'),
      choice(ci('sub'), ci('function'), ci('property'), ci('do'), ci('for')),
      $._terminator,
    ),

    return_statement: $ => seq(
      ci('return'),
      $._terminator,
    ),

    goto_statement: $ => seq(
      ci('goto'),
      choice($.identifier, $.integer_literal),
      $._terminator,
    ),

    gosub_statement: $ => seq(
      ci('gosub'),
      choice($.identifier, $.integer_literal),
      $._terminator,
    ),

    on_error_statement: $ => seq(
      choice(ci('on error'), ci('on local error')),
      choice(
        seq(ci('goto'), choice($.identifier, $.integer_literal, '0')),
        seq(ci('resume'), ci('next')),
      ),
      $._terminator,
    ),

    on_goto_statement: $ => seq(
      ci('on'),
      $._expression,
      ci('goto'),
      choice($.identifier, $.integer_literal),
      repeat(seq(',', choice($.identifier, $.integer_literal))),
      $._terminator,
    ),

    on_gosub_statement: $ => seq(
      ci('on'),
      $._expression,
      ci('gosub'),
      choice($.identifier, $.integer_literal),
      repeat(seq(',', choice($.identifier, $.integer_literal))),
      $._terminator,
    ),

    redim_statement: $ => seq(
      ci('redim'),
      optional(ci('preserve')),
      $.redim_variable,
      repeat(seq(',', $.redim_variable)),
      $._terminator,
    ),

    redim_variable: $ => seq(
      $._lvalue,
      '(',
      $._expression,
      repeat(seq(',', $._expression)),
      ')',
      optional($.as_clause),
    ),

    erase_statement: $ => seq(
      ci('erase'),
      $.identifier,
      repeat(seq(',', $.identifier)),
      $._terminator,
    ),

    raiseevent_statement: $ => seq(
      ci('raiseevent'),
      $.identifier,
      optional($.argument_list),
      $._terminator,
    ),

    // FILE I/O STATEMENTS
    open_statement: $ => seq(
      ci('open'),
      $._expression,
      ci('for'),
      choice(ci('input'), ci('output'), ci('append'), ci('binary'), ci('random')),
      optional(seq(ci('access'), choice(ci('read'), ci('write'), ci('read write')))),
      optional(choice(ci('shared'), ci('lock read'), ci('lock write'), ci('lock read write'))),
      ci('as'),
      $._expression,
      optional(seq(ci('len'), '=', $._expression)),
      $._terminator,
    ),

    close_statement: $ => seq(
      ci('close'),
      optional(seq(
        $._expression,
        repeat(seq(',', $._expression)),
      )),
      $._terminator,
    ),

    input_statement: $ => seq(
      ci('input'),
      $._expression,
      ',',
      $._expression,
      repeat(seq(',', $._expression)),
      $._terminator,
    ),

    line_input_statement: $ => seq(
      ci('line input'),
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    print_statement: $ => seq(
      ci('print'),
      $._expression,
      ',',
      optional($.output_list),
      $._terminator,
    ),

    write_statement: $ => seq(
      ci('write'),
      $._expression,
      ',',
      optional($.output_list),
      $._terminator,
    ),

    output_list: $ => seq(
      $.output_item,
      repeat(seq(choice(',', ';'), optional($.output_item))),
    ),

    output_item: $ => choice(
      seq(choice(ci('spc'), ci('tab')), optional($.argument_list)),
      $._expression,
    ),

    get_statement: $ => seq(
      ci('get'),
      $._expression,
      ',',
      optional($._expression),
      ',',
      $._expression,
      $._terminator,
    ),

    put_statement: $ => seq(
      ci('put'),
      $._expression,
      ',',
      optional($._expression),
      ',',
      $._expression,
      $._terminator,
    ),

    seek_statement: $ => seq(
      ci('seek'),
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    lock_statement: $ => seq(
      ci('lock'),
      $._expression,
      optional(seq(',', $._expression, optional(seq(ci('to'), $._expression)))),
      $._terminator,
    ),

    unlock_statement: $ => seq(
      ci('unlock'),
      $._expression,
      optional(seq(',', $._expression, optional(seq(ci('to'), $._expression)))),
      $._terminator,
    ),

    width_statement: $ => seq(
      ci('width'),
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    // SYSTEM STATEMENTS
    app_activate_statement: $ => seq(
      ci('appactivate'),
      $._expression,
      optional(seq(',', $._expression)),
      $._terminator,
    ),

    beep_statement: $ => seq(
      ci('beep'),
      $._terminator,
    ),

    chdir_statement: $ => seq(
      ci('chdir'),
      $._expression,
      $._terminator,
    ),

    chdrive_statement: $ => seq(
      ci('chdrive'),
      $._expression,
      $._terminator,
    ),

    mkdir_statement: $ => seq(
      ci('mkdir'),
      $._expression,
      $._terminator,
    ),

    rmdir_statement: $ => seq(
      ci('rmdir'),
      $._expression,
      $._terminator,
    ),

    kill_statement: $ => seq(
      ci('kill'),
      $._expression,
      $._terminator,
    ),

    name_statement: $ => seq(
      ci('name'),
      $._expression,
      ci('as'),
      $._expression,
      $._terminator,
    ),

    filecopy_statement: $ => seq(
      ci('filecopy'),
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    load_statement: $ => seq(
      ci('load'),
      $._expression,
      $._terminator,
    ),

    unload_statement: $ => seq(
      ci('unload'),
      $._expression,
      $._terminator,
    ),

    date_statement: $ => seq(
      ci('date'),
      '=',
      $._expression,
      $._terminator,
    ),

    time_statement: $ => seq(
      ci('time'),
      '=',
      $._expression,
      $._terminator,
    ),

    randomize_statement: $ => seq(
      ci('randomize'),
      optional($._expression),
      $._terminator,
    ),

    lset_statement: $ => seq(
      ci('lset'),
      $._lvalue,
      '=',
      $._expression,
      $._terminator,
    ),

    rset_statement: $ => seq(
      ci('rset'),
      $._lvalue,
      '=',
      $._expression,
      $._terminator,
    ),

    mid_statement: $ => seq(
      ci('mid'),
      '(',
      $._expression,
      ',',
      $._expression,
      optional(seq(',', $._expression)),
      ')',
      '=',
      $._expression,
      $._terminator,
    ),

    error_statement: $ => seq(
      ci('error'),
      $._expression,
      $._terminator,
    ),

    resume_statement: $ => seq(
      ci('resume'),
      optional(choice(ci('next'), $.identifier, $.integer_literal)),
      $._terminator,
    ),

    stop_statement: $ => seq(
      ci('stop'),
      $._terminator,
    ),

    end_statement: $ => seq(
      ci('end'),
      $._terminator,
    ),

    sendkeys_statement: $ => seq(
      ci('sendkeys'),
      $._expression,
      optional(seq(',', $._expression)),
      $._terminator,
    ),

    savepicture_statement: $ => seq(
      ci('savepicture'),
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    savesetting_statement: $ => seq(
      ci('savesetting'),
      $._expression,
      ',',
      $._expression,
      ',',
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    deletesetting_statement: $ => seq(
      ci('deletesetting'),
      $._expression,
      ',',
      $._expression,
      optional(seq(',', $._expression)),
      $._terminator,
    ),

    setattr_statement: $ => seq(
      ci('setattr'),
      $._expression,
      ',',
      $._expression,
      $._terminator,
    ),

    reset_statement: $ => seq(
      ci('reset'),
      $._terminator,
    ),

    // ============================================
    // EXPRESSIONS
    // ============================================
    _expression: $ => choice(
      $.literal,
      $.identifier,
      $.parenthesized_expression,
      $.unary_expression,
      $.binary_expression,
      $.new_expression,
      $.typeof_expression,
      $.addressof_expression,
      $.member_expression,
      $.index_expression,
      $.dictionary_access,
      $.call_expression,
    ),

    literal: $ => choice(
      $.integer_literal,
      $.float_literal,
      $.string_literal,
      $.boolean_literal,
      $.nothing_literal,
      $.date_literal,
      $.color_literal,
    ),

    integer_literal: $ => choice(
      /\d+/,
      /&[hH][0-9a-fA-F]+/,
      /&[oO][0-7]+/,
    ),

    float_literal: $ => /\d*\.\d+([eE][+-]?\d+)?|\d+[eE][+-]?\d+/,

    string_literal: $ => /"([^"]|"")*"/,

    boolean_literal: $ => choice(ci('true'), ci('false')),

    nothing_literal: $ => choice(ci('nothing'), ci('null'), ci('empty')),

    date_literal: $ => $.date_literal_token,

    color_literal: $ => /&[hH][0-9a-fA-F]+&/,

    parenthesized_expression: $ => seq('(', $._expression, ')'),

    unary_expression: $ => choice(
      prec(PREC.UNARY, seq('-', $._expression)),
      prec(PREC.UNARY, seq('+', $._expression)),
      prec(PREC.NOT, seq(ci('not'), $._expression)),
    ),

    binary_expression: $ => choice(
      prec.left(PREC.POW, seq($._expression, '^', $._expression)),
      prec.left(PREC.MUL, seq($._expression, choice('*', '/'), $._expression)),
      prec.left(PREC.IDIV, seq($._expression, '\\', $._expression)),
      prec.left(PREC.MOD, seq($._expression, ci('mod'), $._expression)),
      prec.left(PREC.ADD, seq($._expression, choice('+', '-'), $._expression)),
      prec.left(PREC.CONCAT, seq($._expression, '&', $._expression)),
      prec.left(PREC.COMPARE, seq($._expression, $._compare_op, $._expression)),
      prec.left(PREC.COMPARE, seq($._expression, ci('like'), $._expression)),
      prec.left(PREC.COMPARE, seq($._expression, ci('is'), $._expression)),
      prec.left(PREC.NOT, seq($._expression, ci('not'), $._expression)),
      prec.left(PREC.AND, seq($._expression, ci('and'), $._expression)),
      prec.left(PREC.OR, seq($._expression, ci('or'), $._expression)),
      prec.left(PREC.XOR, seq($._expression, ci('xor'), $._expression)),
      prec.left(PREC.EQV, seq($._expression, ci('eqv'), $._expression)),
      prec.left(PREC.IMP, seq($._expression, ci('imp'), $._expression)),
    ),

    _compare_op: $ => choice('=', '<>', '<', '>', '<=', '>='),

    new_expression: $ => prec(PREC.CALL, seq(ci('new'), $._type)),

    typeof_expression: $ => prec(PREC.CALL, seq(
      ci('typeof'),
      $._expression,
      ci('is'),
      $._type,
    )),

    addressof_expression: $ => prec(PREC.CALL, seq(ci('addressof'), $._expression)),

    member_expression: $ => prec.left(PREC.CALL, seq(
      field('object', $._expression),
      '.',
      field('member', $.identifier),
      optional($.type_hint),
    )),

    dictionary_access: $ => prec.left(PREC.CALL, seq(
      field('object', $._expression),
      '!',
      field('key', $.identifier),
      optional($.type_hint),
    )),

    index_expression: $ => prec(PREC.CALL, seq(
      field('object', $._expression),
      '(',
      optional($.argument_list_inner),
      ')',
    )),

    call_expression: $ => prec(PREC.CALL, seq(
      field('function', choice($.identifier, $.member_expression, $.dictionary_access)),
      optional($.type_hint),
      $.argument_list,
    )),

    argument_list: $ => seq('(', optional($.argument_list_inner), ')'),

    argument_list_inner: $ => seq(
      optional($._argument),
      repeat(seq(',', optional($._argument))),
    ),

    _argument: $ => choice(
      seq($.identifier, ':=', $._expression),  // Named
      $._expression,  // Positional
    ),

    _lvalue: $ => choice(
      $.identifier,
      $.member_expression,
      $.dictionary_access,
      $.index_expression,
    ),

    // ============================================
    // IDENTIFIERS
    // ============================================
    identifier: $ => /[a-zA-Z_][a-zA-Z0-9_]*/,

    dotted_name: $ => prec.left(seq(
      $.identifier,
      repeat(seq('.', $.identifier)),
    )),

    type_hint: $ => choice('%', '&', '!', '#', '@', '$'),

    // ============================================
    // VISIBILITY
    // ============================================
    _visibility: $ => choice(
      ci('public'),
      ci('private'),
      ci('friend'),
      ci('global'),
    ),

    // ============================================
    // COMMENTS AND TERMINATORS
    // ============================================
    comment: $ => token(choice(
      seq("'", /[^\r\n]*/),
      seq(/[rR][eE][mM]/, /[ \t]/, /[^\r\n]*/),
    )),

    _terminator: $ => choice($._newline, ':'),
    _newline: $ => /\r?\n/,
  },
});
