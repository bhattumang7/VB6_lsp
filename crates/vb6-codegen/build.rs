use std::env;
use std::fs;
use std::path::{Path, PathBuf};

fn sanitize(name: &str) -> String {
    name.chars()
        .map(|c| if c.is_ascii_alphanumeric() || c == '_' { c } else { '_' })
        .collect()
}

fn main() {
    let manifest_dir = env::var("CARGO_MANIFEST_DIR").unwrap();
    let fixtures_dir = Path::new(&manifest_dir).join("tests").join("fixtures");
    println!("cargo:rerun-if-changed={}", fixtures_dir.display());

    let out_dir = env::var("OUT_DIR").unwrap();
    let dest = Path::new(&out_dir).join("fixture_tests.rs");

    let mut generated = String::new();
    let mut count = 0usize;

    if fixtures_dir.is_dir() {
        let mut entries: Vec<PathBuf> = fs::read_dir(&fixtures_dir)
            .unwrap()
            .filter_map(|e| e.ok())
            .map(|e| e.path())
            .filter(|p| p.is_dir())
            .collect();
        entries.sort();

        for dir in entries {
            let input = dir.join("input.bas");
            let expected = dir.join("expected.pcode");
            if !input.is_file() || !expected.is_file() {
                continue;
            }
            println!("cargo:rerun-if-changed={}", input.display());
            println!("cargo:rerun-if-changed={}", expected.display());
            let cls = dir.join("input.cls");
            if cls.is_file() {
                println!("cargo:rerun-if-changed={}", cls.display());
            }
            let case_name = dir.file_name().unwrap().to_string_lossy().to_string();
            let fn_name = sanitize(&case_name);
            count += 1;
            generated.push_str(&format!(
                "#[test]\nfn fixture_{fn_name}() {{\n    run_fixture({case_name:?});\n}}\n\n"
            ));
        }
    }

    generated.push_str(&format!("#[allow(dead_code)]\npub const FIXTURE_COUNT: usize = {count};\n"));

    fs::write(&dest, generated).unwrap();
}
