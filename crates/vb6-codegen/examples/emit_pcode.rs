//! Emit the P-code byte stream for one proc of a VB6 module, for live
//! cross-checking against the original compiler.
//!
//! Usage: `emit_pcode <source.bas> [module_desc_hex] [proc_index]`
//!
//! Prints the lowered proc's bytes as space-separated lowercase hex on one line.
//! `module_desc` defaults to `0008` (the primary module in a single-module
//! project); `proc_index` defaults to `0`.

use vb6_codegen::{lower_module, lower_proc};
use vb6_sema::frontend::ast::ExprArena;
use vb6_sema::frontend::parser::Parser;
use vb6_sema::frontend::scanner::ScannerContext;
use vb6_sema::sema::bind;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("usage: emit_pcode <source.bas> [module_desc_hex] [proc_index]");
        std::process::exit(2);
    }
    let src = std::fs::read_to_string(&args[1]).expect("read source file");
    let module_desc = u16::from_str_radix(args.get(2).map(|s| s.as_str()).unwrap_or("0008"), 16)
        .expect("module_desc hex");
    let proc_index: usize = args.get(3).and_then(|s| s.parse().ok()).unwrap_or(0);

    let mut ctx = ScannerContext::new(1, 1, 0x0409);
    ctx.intern_keywords();
    let mut arena = ExprArena::new();
    let mut parser = Parser::new(&mut ctx, src.as_bytes());
    let top = parser.parse_module(&mut arena);
    let spans = std::mem::take(&mut parser.node_spans);
    let vis = std::mem::take(&mut parser.decl_public);
    drop(parser);
    let module = bind(&ctx, &arena, &top, &spans, &vis);

    // `proc_index` of "all" emits every procedure with one module-global string
    // pool, one line of hex per procedure (in declaration order).
    if args.get(3).map(|s| s.as_str()) == Some("all") {
        let procs = lower_module(&module, &arena, module_desc)
            .unwrap_or_else(|e| panic!("lower_module failed: {e:?}"));
        for bytes in procs {
            let hex: Vec<String> = bytes.iter().map(|b| format!("{b:02x}")).collect();
            println!("{}", hex.join(" "));
        }
        return;
    }

    let bytes = lower_proc(&module, proc_index, &arena, module_desc)
        .unwrap_or_else(|e| panic!("lower_proc failed: {e:?}"));

    let hex: Vec<String> = bytes.iter().map(|b| format!("{b:02x}")).collect();
    println!("{}", hex.join(" "));
}
