//! Type-declaration characters after keywords: `#` after the file statements is
//! the channel mark, `Date$` stands wherever `Date` does, `String$` needs its
//! arguments, and `Dir$` in a declaration with `As` is a suffixed name.

mod common;
use common::{offset_of, parse};

fn in_sub(stmt: &str) -> String {
    format!("Sub S(o As Object)\n    {stmt}\nEnd Sub\n")
}

#[test]
fn a_hash_after_a_file_word_is_the_channel_mark() {
    for word in ["Close", "Get", "Input", "Lock", "Open", "Print", "Put", "Seek", "Unlock", "Write"] {
        for text in [format!("x = {word}#"), format!("x = o.{word}#"), format!("{word}#")] {
            let src = in_sub(&text);
            assert!(!parse(&src).found.is_empty(), "{text}");
        }
    }
}

#[test]
fn a_hash_after_other_words_is_still_a_suffix() {
    for text in ["x = o.Name#", "x = o.Close%", "x = o.Print%"] {
        assert_eq!(parse(&in_sub(text)).found, vec![], "{text}");
    }
    assert_eq!(parse(&in_sub("Print #1, 2")).found, vec![]);
    assert_eq!(parse(&in_sub("Close #1")).found, vec![]);
}

#[test]
fn date_dollar_stands_where_date_stands() {
    for text in ["Date$ = \"1/1/2000\"", "Date$", "Date$ 1", "x = Date$", "x = o.Date$"] {
        assert_eq!(parse(&in_sub(text)).found, vec![], "{text}");
    }
}

#[test]
fn string_dollar_needs_its_two_arguments() {
    for text in ["x = String$", "x = String$(3)", "x = String$(3, \"a\", 1)"] {
        assert!(!parse(&in_sub(text)).found.is_empty(), "{text}");
    }
    assert_eq!(parse(&in_sub("x = String$(3, \"a\")")).found, vec![]);
}

#[test]
fn dir_dollar_with_as_in_a_declaration_is_a_syntax_error_at_the_as() {
    for text in ["Dim Dir$ As String", "Static dir$ As Long", "Dim a, Dir$ As String"] {
        let src = in_sub(text);
        let p = parse(&src);
        assert!(p.found.iter().any(|d| d.1 == offset_of(&src, "As", 0) || d.1 == offset_of(&src, "As", 1)), "{text}: {:?}", p.found);
    }
    let src = "Type T\n    Dir$ As Long\nEnd Type\n";
    assert!(!parse(src).found.is_empty());
}

#[test]
fn dir_dollar_elsewhere_and_other_dollar_words_stay_names() {
    for text in ["Dim Dir$", "Dim CurDir$ As Long", "Dim Format$ As Long", "Dim Mid$ As String", "Const Dir$ As Long = 1"] {
        assert_eq!(parse(&in_sub(text)).found, vec![], "{text}");
    }
    assert_eq!(parse("Sub S(Dir$ As Long)\nEnd Sub\n").found, vec![]);
}
