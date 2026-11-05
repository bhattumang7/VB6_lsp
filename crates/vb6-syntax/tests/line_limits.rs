//! A physical line is read up to its 1023rd byte, and a logical line joins at most 24
//! continued lines.

mod common;
use common::parse;

#[test]
fn a_line_is_cut_after_1023_bytes() {
    // `s = "` and `"` take 6 bytes plus 4 of indent: a 1013-byte body makes 1023 in all.
    let ok = format!("Sub S()\n    s = \"{}\"\nEnd Sub\n", "a".repeat(1013));
    assert_eq!(parse(&ok).found, vec![]);
    for extra in [1usize, 2] {
        let long = format!("Sub S()\n    s = \"{}\"\nEnd Sub\n", "a".repeat(1013 + extra));
        assert!(!parse(&long).found.is_empty(), "{extra}");
    }
    // Comment text past the cut does no harm.
    let comment = format!("Sub S()\n    x = 1 ' {}\nEnd Sub\n", "c".repeat(1100));
    assert_eq!(parse(&comment).found, vec![]);
}

#[test]
fn twenty_four_continuations_are_joined_and_a_twenty_fifth_is_an_error() {
    let joined = |n: usize| format!("Sub S()\n    x = 1{}\nEnd Sub\n", " + _\n 1".repeat(n));
    assert_eq!(parse(&joined(24)).found, vec![]);
    let p = parse(&joined(25));
    assert_eq!(p.found.first().map(|f| f.0), Some(0x9c00));
}

#[test]
fn a_full_symbol_table_is_out_of_memory() {
    // 32767 names fill the table of names; the next new one cannot be kept.
    let mut src = String::from("Sub S()\n");
    for chunk in (0..33000).collect::<Vec<usize>>().chunks(60) {
        let names: Vec<String> = chunk.iter().map(|i| format!("q{i}")).collect();
        src.push_str(&format!(" x = {}\n", names.join(" + ")));
    }
    src.push_str("End Sub\n");
    let p = parse(&src);
    assert!(p.found.iter().any(|f| f.0 == 7), "{:?}", p.found.first());
}

fn types_with_members(members: usize) -> String {
    let mut src = String::new();
    let (mut i, mut t) = (0, 0);
    while i < members {
        src.push_str(&format!("Type T{t}\n"));
        for k in i..members.min(i + 500) {
            src.push_str(&format!(" m{k} As Long\n"));
        }
        src.push_str("End Type\n");
        i += 500;
        t += 1;
    }
    src.push_str("Sub Main()\nEnd Sub\n");
    src
}

#[test]
fn the_symbol_table_holds_as_many_names_as_the_compiler_does() {
    // 32436 members fit next to the keyword symbols; one more does not.
    assert_eq!(parse(&types_with_members(32436)).found, vec![]);
    let p = parse(&types_with_members(32437));
    assert!(p.found.iter().any(|f| f.0 == 7), "{:?}", p.found.first());
}

/// `x = t + t + ...` over continued lines of `per` terms each.
fn chain(term: &str, op: &str, n: usize, per: usize) -> String {
    let terms = vec![term; n];
    let lines: Vec<String> = terms.chunks(per).map(|c| c.join(&format!(" {op} "))).collect();
    format!("Sub S()\nDim x As Long\nx = {}\nEnd Sub\n", lines.join(&format!(" {op} _\n")))
}

#[test]
fn a_statement_whose_compiled_line_passes_15360_bytes_is_too_complex() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    // (term, operator, terms per physical line, most terms that fit)
    for (term, op, per, fits) in [
        ("1", "+", 250, 2552),
        ("y", "+", 250, 2552),
        ("40000", "+", 110, 1910),
        ("1.5!", "+", 120, 1911),
        ("1.5", "+", 150, 1276),
        ("1.5@", "+", 120, 1276),
        ("#1/1/2000#", "+", 60, 1272),
        ("\"ab\"", "&", 120, 1911),
        ("\"abcdefghij\"", "&", 60, 955),
    ] {
        assert_eq!(parse(&chain(term, op, fits, per)).found, vec![], "{term} x {fits}");
        let found = parse(&chain(term, op, fits + 1, per)).found;
        assert_eq!(found.first().map(|f| f.0), Some(TOO_COMPLEX), "{term} x {}", fits + 1);
    }
}

#[test]
fn the_size_rule_holds_for_every_binary_operator_sign_not_and_brackets() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    // (term, operator, terms per physical line, most terms that fit)
    let mut cases: Vec<(&str, &str, usize, usize)> = Vec::new();
    for op in ["-", "*", "/", "\\", "Or", "^", "=", "<"] {
        cases.push(("1", op, 200, 2550));
    }
    for op in ["Mod", "And", "Xor", "Eqv", "Imp", ">=", "<>", ">", "<="] {
        cases.push(("1", op, 150, 2548));
    }
    cases.extend([
        ("-1", "+", 150, 1913),
        ("-1.5", "+", 120, 1094),
        ("-y", "+", 100, 1909),
        ("Not y", "+", 80, 1907),
        ("(y)", "+", 150, 1913),
    ]);
    for (term, op, per, fits) in cases {
        assert_eq!(parse(&chain(term, op, fits, per)).found, vec![], "{term} {op} x {fits}");
        let found = parse(&chain(term, op, fits + 1, per)).found;
        assert_eq!(found.first().map(|f| f.0), Some(TOO_COMPLEX), "{term} {op} x {}", fits + 1);
    }
}

#[test]
fn deeply_nested_brackets_and_calls_the_compiler_accepts_are_parsed() {
    let parens = format!("Sub S()\nx = {}1{}\nEnd Sub\n", "(".repeat(480), ")".repeat(480));
    assert_eq!(parse(&parens).found, vec![]);
    let calls = format!("Sub S()\nx = {}1{}\nEnd Sub\n", "Abs(".repeat(100), ")".repeat(100));
    assert_eq!(parse(&calls).found, vec![]);
}

/// `x = a op a op ...` with the terms and operators taken in turn.
fn mixed_chain(terms: &[&str], ops: &[&str], n: usize, per: usize) -> String {
    let mut parts = Vec::new();
    for k in 0..n {
        parts.push(if k == 0 { terms[0].to_string() } else { format!("{} {}", ops[(k - 1) % ops.len()], terms[k % terms.len()]) });
    }
    let lines: Vec<String> = parts.chunks(per).map(|c| c.join(" ")).collect();
    format!("Sub S()\nDim x As Long\nx = {}\nEnd Sub\n", lines.join(" _\n"))
}

#[test]
fn operators_and_operand_kinds_may_be_mixed_in_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let ok = |terms: &[&str], ops: &[&str], per: usize, n: usize| parse(&mixed_chain(terms, ops, n, per)).found.is_empty();
    // (terms, operators, terms per line, most terms that fit)
    for (terms, ops, per, fits) in [
        (vec!["1"], vec!["+", "-"], 150, 2548),
        (vec!["1"], vec!["+", "*", "And"], 150, 2548),
        (vec!["1"], vec!["+", "/", "^", "="], 150, 2548),
        (vec!["1"], vec!["<", "Or", "Mod", "\\"], 150, 2548),
        (vec!["1", "40000"], vec!["+"], 150, 2185),
        (vec!["1", "1.5"], vec!["+"], 100, 1699),
    ] {
        assert!(ok(&terms, &ops, per, fits), "{terms:?} {ops:?} x {fits}");
        let found = parse(&mixed_chain(&terms, &ops, fits + 1, per)).found;
        assert_eq!(found.first().map(|f| f.0), Some(TOO_COMPLEX), "{terms:?} {ops:?} x {}", fits + 1);
    }
}

#[test]
fn a_block_statement_line_follows_the_same_size_rule_as_an_assignment() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize| {
        let terms = vec!["1"; n];
        terms.chunks(150).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    for (head, tail) in [
        ("If ", " Then\nEnd If"),
        ("While ", "\nWend"),
        ("Do While ", "\nLoop"),
        ("Do Until ", "\nLoop"),
        ("Select Case ", "\nCase 1\nEnd Select"),
    ] {
        let src = |n| format!("Sub S()\n{head}{}{tail}\nEnd Sub\n", expr(n));
        assert_eq!(parse(&src(2548)).found, vec![], "{head}");
        assert_eq!(parse(&src(2549)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{head}");
    }
}

#[test]
fn debug_print_of_one_expression_costs_two_bytes_more_than_an_assignment() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (terms per line, most terms that fit)
    for (per, fits) in [(120, 2544), (135, 2546), (150, 2547), (170, 2549)] {
        let src = |n| format!("Sub S()\nDebug.Print {}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "per {per}");
    }
}

#[test]
fn assignments_before_one_on_the_same_line_add_to_its_compiled_size() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statements before, terms per line, most terms that fit)
    for (before, per, fits) in [
        ("y = 1: ", 120, 2542),
        ("y = 1: ", 135, 2544),
        ("y = 1: ", 150, 2546),
        ("y = 1: ", 170, 2547),
        ("y = 40000: ", 120, 2542),
        ("y = 40000: ", 150, 2545),
        ("y = 1: y = 1: ", 150, 2544),
    ] {
        let src = |n| format!("Sub S()\n{before}x = {}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before} per {per}");
    }
}

#[test]
fn signs_not_and_brackets_cost_two_on_literals_and_like_costs_what_an_operator_does() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    // (term, operator, terms per physical line, most terms that fit)
    for (term, op, per, fits) in [
        ("(1)", "+", 150, 1913),
        ("Not 1", "+", 100, 1909),
        ("(1.5)", "+", 100, 1093),
        ("-40000", "+", 100, 1529),
        ("(\"a\" Like \"b\")", "&", 60, 849),
    ] {
        assert_eq!(parse(&chain(term, op, fits, per)).found, vec![], "{term} x {fits}");
        let found = parse(&chain(term, op, fits + 1, per)).found;
        assert_eq!(found.first().map(|f| f.0), Some(TOO_COMPLEX), "{term} x {}", fits + 1);
    }
}

#[test]
fn a_continuation_mark_past_the_cut_is_not_part_of_the_text() {
    // 1023 bytes of text and then ` _`: the line is not cut and the mark joins the next line.
    let head = "    x = 1";
    let filler = " + 1".repeat((1023 - head.len()) / 4);
    let line = format!("{head}{filler}");
    let pad = " ".repeat(1023 - line.len());
    let ok = format!("Sub S()\n{line}{pad} _\n + 1\nEnd Sub\n");
    assert_eq!(parse(&ok).found, vec![]);
}

#[test]
fn wrapped_single_currency_and_date_literals_and_file_output_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    for (term, per, fits) in [("(1.5!)", 100, 1529), ("-1.5!", 100, 1529), ("(1.5@)", 100, 1093), ("(#1/1/2000#)", 60, 1091)] {
        assert_eq!(parse(&chain(term, "+", fits, per)).found, vec![], "{term} x {fits}");
        let found = parse(&chain(term, "+", fits + 1, per)).found;
        assert_eq!(found.first().map(|f| f.0), Some(TOO_COMPLEX), "{term} x {}", fits + 1);
    }
    // Print #n and Write #n of one expression cost 6 more than assigning it.
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    for head in ["Print #1, ", "Write #1, "] {
        for (per, fits) in [(120, 2543), (150, 2547), (170, 2548)] {
            let src = |n| format!("Sub S()\n{head}{}\nEnd Sub\n", expr(n, per));
            assert_eq!(parse(&src(fits)).found, vec![], "{head} per {per}");
            assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{head} per {per}");
        }
    }
}

#[test]
fn output_lists_cost_a_two_byte_separator_between_items() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let list = |term: &str, sep: &str, n: usize, per: usize| {
        let terms = vec![term; n];
        terms.chunks(per).map(|c| c.join(&format!(" {sep} "))).collect::<Vec<_>>().join(&format!(" {sep} _\n"))
    };
    // (statement start, term, separator, terms per line, most terms that fit)
    for (head, term, sep, per, fits) in [
        ("Print #1, ", "1", ";", 200, 2549),
        ("Print #1, ", "1", ",", 200, 2549),
        ("Write #1, ", "1", ",", 200, 2549),
        ("Print #1, ", "\"ab\"", ";", 100, 1908),
    ] {
        let src = |n| format!("Sub S()\n{head}{}\nEnd Sub\n", list(term, sep, n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{head}{term} {sep}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{head}{term} {sep}");
    }
}

#[test]
fn is_costs_what_an_operator_does() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    assert_eq!(parse(&chain("(o Is o)", "+", 1092, 80)).found, vec![]);
    let found = parse(&chain("(o Is o)", "+", 1093, 80)).found;
    assert_eq!(found.first().map(|f| f.0), Some(TOO_COMPLEX));
}

#[test]
fn open_costs_eight_bytes_more_than_assigning_its_path() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["s"; n];
        terms.chunks(per).map(|c| c.join(" & ")).collect::<Vec<_>>().join(" & _\n")
    };
    for clauses in [" For Output As #1", " For Input As #1", " For Binary As #1", " As #1", " For Output Access Write As #1"] {
        for (per, fits) in [(150, 2546), (170, 2548), (200, 2549)] {
            let src = |n| format!("Sub S()\nOpen {}{clauses}\nEnd Sub\n", expr(n, per));
            assert_eq!(parse(&src(fits)).found, vec![], "{clauses} per {per}");
            assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{clauses} per {per}");
        }
    }
}

#[test]
fn spc_tab_and_the_length_clause_of_open_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let list = |term: &str, sep: &str, n: usize, per: usize| {
        let terms = vec![term; n];
        terms.chunks(per).map(|c| c.join(&format!(" {sep} "))).collect::<Vec<_>>().join(&format!(" {sep} _\n"))
    };
    for term in ["Spc(1)", "Tab(1)"] {
        let src = |n| format!("Sub S()\nPrint #1, {}\nEnd Sub\n", list(term, ";", n, 100));
        assert_eq!(parse(&src(1908)).found, vec![], "{term}");
        assert_eq!(parse(&src(1909)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{term}");
    }
    for (per, fits) in [(120, 2543), (150, 2546), (170, 2547)] {
        let src = |n| format!("Sub S()\nOpen s For Output As #1 Len = {}\nEnd Sub\n", list("1", "+", n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "Len per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "Len per {per}");
    }
}

#[test]
fn the_channel_of_print_is_an_expression_that_adds_its_own_bytes() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (channel, terms per line, most terms that fit)
    for (channel, per, fits) in [
        ("y", 150, 2547),
        ("y", 170, 2548),
        ("1 + 1", 150, 2546),
        ("1 + 1", 170, 2547),
        ("40000", 150, 2546),
        ("40000", 170, 2548),
        ("(1)", 150, 2546),
        ("(1)", 170, 2548),
    ] {
        let src = |n| format!("Sub S()\nPrint #{channel}, {}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "#{channel} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "#{channel} per {per}");
    }
}

#[test]
fn for_elseif_and_case_lines_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statement before the expression, after it, terms per line, most terms that fit)
    for (before, after, per, fits) in [
        ("For i = ", " To 5\nNext i", 150, 2546),
        ("For i = ", " To 5\nNext i", 170, 2547),
        ("For i = 1 To ", "\nNext i", 150, 2546),
        ("For i = 1 To ", "\nNext i", 170, 2547),
        ("If y Then\nElseIf ", " Then\nEnd If", 150, 2548),
        ("If y Then\nElseIf ", " Then\nEnd If", 170, 2549),
        ("Select Case y\nCase ", "\nEnd Select", 150, 2548),
        ("Select Case y\nCase ", "\nEnd Select", 170, 2549),
    ] {
        let src = |n| format!("Sub S()\n{before}{}{after}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
    }
}

#[test]
fn step_and_the_other_forms_of_case_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statement before the expression, after it, terms per line, most terms that fit)
    for (before, after, per, fits) in [
        ("For i = ", " To 5 Step 1\nNext i", 150, 2545),
        ("For i = ", " To 5 Step 1\nNext i", 170, 2547),
        ("For i = 1 To 5 Step ", "\nNext i", 150, 2545),
        ("For i = 1 To 5 Step ", "\nNext i", 170, 2547),
        ("Select Case y\nCase Is > ", "\nEnd Select", 150, 2548),
        ("Select Case y\nCase Is > ", "\nEnd Select", 170, 2549),
        ("Select Case y\nCase 1 To ", "\nEnd Select", 150, 2547),
        ("Select Case y\nCase 1 To ", "\nEnd Select", 170, 2548),
        ("Select Case y\nCase 2, ", "\nEnd Select", 150, 2547),
        ("Select Case y\nCase ", ", 2\nEnd Select", 170, 2548),
    ] {
        let src = |n| format!("Sub S()\n{before}{}{after}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
    }
}

#[test]
fn loop_conditions_seek_and_local_const_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statement before the expression, after it, terms per line, most terms that fit)
    for (before, after, per, fits) in [
        ("Do\nLoop While ", "", 150, 2548),
        ("Do\nLoop While ", "", 170, 2549),
        ("Do\nLoop Until ", "", 150, 2548),
        ("Seek #1, ", "", 120, 2544),
        ("Seek #1, ", "", 165, 2548),
        ("Const c = ", "", 120, 2544),
        ("Const c = ", "", 165, 2548),
    ] {
        let src = |n| format!("Sub S()\n{before}{}{after}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
    }
}

#[test]
fn randomize_redim_mid_and_put_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statement before the expression, after it, terms per line, most terms that fit)
    for (before, after, per, fits) in [
        ("Randomize ", "", 150, 2547),
        ("Randomize ", "", 170, 2549),
        ("ReDim a(", ")", 150, 2546),
        ("ReDim a(", ")", 170, 2548),
        ("Mid$(s, 1) = ", "", 150, 2546),
        ("Mid$(s, 1) = ", "", 170, 2548),
        ("Put #1, , ", "", 120, 2543),
        ("Put #1, , ", "", 165, 2547),
    ] {
        let src = |n| format!("Sub S()\n{before}{}{after}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
    }
}

#[test]
fn every_mid_spelling_and_put_or_get_with_a_record_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    for (before, after) in [
        ("Mid(s, 1) = ", ""),
        ("MidB(s, 1) = ", ""),
        ("MidB$(s, 1) = ", ""),
        ("Mid$(s, 1) = ", ""),
        ("Put #1, ", ", y"),
        ("Get #1, ", ", y"),
    ] {
        for (per, fits) in [(120, 2543), (165, 2547)] {
            let src = |n| format!("Sub S()\n{before}{}{after}\nEnd Sub\n", expr(n, per));
            assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
            assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
        }
    }
}

#[test]
fn lset_rset_width_lock_and_unlock_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statement before the expression, terms per line, most terms that fit)
    for (before, per, fits) in [
        ("LSet s = ", 120, 2544),
        ("LSet s = ", 165, 2548),
        ("RSet s = ", 120, 2544),
        ("Width #1, ", 120, 2543),
        ("Width #1, ", 150, 2546),
        ("Width #1, ", 165, 2547),
        ("Lock #1, ", 120, 2543),
        ("Lock #1, ", 150, 2547),
        ("Unlock #1, ", 150, 2547),
    ] {
        let src = |n| format!("Sub S()\n{before}{}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
    }
}

#[test]
fn module_level_and_listed_const_values_follow_the_size_rule() {
    const TOO_COMPLEX: u32 = 0x9c6e;
    let expr = |n: usize, per: usize| {
        let terms = vec!["1"; n];
        terms.chunks(per).map(|c| c.join(" + ")).collect::<Vec<_>>().join(" + _\n")
    };
    // (statement before the expression, terms per line, most terms that fit)
    for (before, per, fits) in [
        ("Const c = ", 120, 2544),
        ("Const c = ", 150, 2547),
        ("Public Const c = ", 165, 2548),
        ("Private Const c = ", 120, 2544),
    ] {
        let src = |n| format!("{before}{}\nSub S()\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "{before:?} per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "{before:?} per {per}");
    }
    for (per, fits) in [(120, 2542), (150, 2545), (165, 2546)] {
        let src = |n| format!("Sub S()\nConst a = 1, c = {}\nEnd Sub\n", expr(n, per));
        assert_eq!(parse(&src(fits)).found, vec![], "list per {per}");
        assert_eq!(parse(&src(fits + 1)).found.first().map(|f| f.0), Some(TOO_COMPLEX), "list per {per}");
    }
}
