//! VB6 built-in function and statement registry.
//!
//! Used by the binder to recognise names that refer to VB6 intrinsics
//! (standard VBA functions and statements) so they are not erroneously
//! flagged as unresolved.
//!
//! The table lists the standard VB6/VBA built-in functions and statements —
//! the set the compiler resolves before reporting "Sub or Function not
//! defined". Host/object methods (e.g. `Form.Print`) are member accesses, not
//! global-name resolution, and are out of scope for this table.

use std::collections::HashSet;
use std::sync::OnceLock;

/// Returns true if `name` (lowercase) is a known VB6 built-in.
///
/// Backed by a lazily-built set so membership does not depend on the source
/// list staying sorted (a mis-ordered entry would otherwise silently
/// false-negative under a binary search and produce a spurious "not defined").
pub fn is_builtin(name_lower: &str) -> bool {
    static SET: OnceLock<HashSet<&'static str>> = OnceLock::new();
    SET.get_or_init(|| BUILTINS.iter().copied().collect())
        .contains(name_lower)
}

/// All known built-in names (lowercase).
pub fn builtin_names() -> &'static [&'static str] {
    BUILTINS
}

/// Lowercase built-in names. Kept roughly alphabetical for readability; exact
/// ordering is not required for correctness (see [`is_builtin`]).
static BUILTINS: &[&str] = &[
    "abs",
    "appactivate",
    "array",
    "asc",
    "ascb",
    "ascw",
    "atn",
    "beep",
    "callbyname",
    "cbool",
    "cbyte",
    "ccur",
    "cdate",
    "cdbl",
    "cdec",
    "chdir",
    "chdrive",
    "choose",
    "chr",
    "chr$",
    "chrb",
    "chrb$",
    "chrw",
    "chrw$",
    "cint",
    "clng",
    "close",
    "command",
    "command$",
    "cos",
    "createobject",
    "csng",
    "cstr",
    "curdir",
    "curdir$",
    "cvar",
    "cvdate",
    "cverr",
    "date",
    "date$",
    "dateadd",
    "datediff",
    "datepart",
    "dateserial",
    "datevalue",
    "day",
    "ddb",
    "deletesetting",
    "dir",
    "dir$",
    "doevents",
    "empty",
    "environ",
    "environ$",
    "eof",
    "erase",
    "erl",
    "err",
    "error",
    "error$",
    "eval",
    "exp",
    "false",
    "fileattr",
    "filecopy",
    "filedatetime",
    "filelen",
    "filter",
    "fix",
    "format",
    "format$",
    "formatcurrency",
    "formatdatetime",
    "formatnumber",
    "formatpercent",
    "freefile",
    "fv",
    "get",
    "getallsettings",
    "getattr",
    "getobject",
    "getref",
    "getsetting",
    "hex",
    "hex$",
    "hour",
    "iif",
    "imestatus",
    "input",
    "input$",
    "inputb",
    "inputb$",
    "inputbox",
    "inputbox$",
    "instr",
    "instrb",
    "instrrev",
    "int",
    "ipmt",
    "irr",
    "isarray",
    "isdate",
    "isempty",
    "iserror",
    "ismissing",
    "isnull",
    "isnumeric",
    "isobject",
    "join",
    "kill",
    "lbound",
    "lcase",
    "lcase$",
    "left",
    "left$",
    "leftb",
    "leftb$",
    "len",
    "lenb",
    "load",
    "loc",
    "lock",
    "lof",
    "log",
    "lset",
    "ltrim",
    "ltrim$",
    "mid",
    "mid$",
    "midb",
    "midb$",
    "minute",
    "mirr",
    "mkdir",
    "month",
    "monthname",
    "msgbox",
    "nothing",
    "now",
    "nper",
    "npv",
    "null",
    "nz",
    "oct",
    "oct$",
    "open",
    "partition",
    "pmt",
    "ppmt",
    "print",
    "put",
    "pv",
    "qbcolor",
    "randomize",
    "rate",
    "replace",
    "reset",
    "rgb",
    "right",
    "right$",
    "rightb",
    "rightb$",
    "rnd",
    "round",
    "rmdir",
    "rset",
    "rtrim",
    "rtrim$",
    "savesetting",
    "second",
    "seek",
    "sendkeys",
    "setattr",
    "sgn",
    "shell",
    "sin",
    "sln",
    "space",
    "space$",
    "split",
    "spc",
    "sqr",
    "str",
    "str$",
    "strcomp",
    "strconv",
    "string",
    "string$",
    "strreverse",
    "switch",
    "syd",
    "tab",
    "tan",
    "time",
    "time$",
    "timer",
    "timeserial",
    "timevalue",
    "trim",
    "trim$",
    "true",
    "typename",
    "ubound",
    "ucase",
    "ucase$",
    "unload",
    "unlock",
    "val",
    "vartype",
    "weekday",
    "weekdayname",
    "width",
    "with",
    "write",
    "year",
];

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;

    #[test]
    fn no_duplicate_or_empty_entries() {
        let mut seen = HashSet::new();
        for &name in BUILTINS {
            assert!(!name.is_empty(), "empty builtin name");
            assert_eq!(name, name.to_ascii_lowercase(), "builtin `{name}` must be lowercase");
            assert!(seen.insert(name), "duplicate builtin entry: {name}");
        }
    }

    #[test]
    fn recognises_runtime_intrinsics_previously_missing() {
        // Standard VB6 intrinsics that were absent before and would have
        // produced a spurious "Sub or Function not defined".
        for name in [
            "dir", "dir$", "randomize", "rgb", "qbcolor", "imestatus", "cvdate",
            "pmt", "ppmt", "ipmt", "fv", "pv", "nper", "npv", "irr", "mirr",
            "ddb", "sln", "syd", "rate",
            "formatnumber", "formatcurrency", "formatpercent", "formatdatetime",
        ] {
            assert!(is_builtin(name), "expected `{name}` to be a recognised builtin");
        }
    }

    #[test]
    fn lookup_is_order_independent() {
        // `spc` sits after `split` in source order; a binary search would miss
        // it. The set-based lookup must still find it.
        assert!(is_builtin("spc"));
        assert!(is_builtin("split"));
        assert!(!is_builtin("definitelynotabuiltin"));
    }
}
