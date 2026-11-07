//! Changing many files at once links the project once.

use vb6_engine::session::Session;

fn modules(count: usize) -> Vec<(String, Vec<u8>)> {
    (0..count)
        .map(|i| (format!("m{i}.bas"), format!("Option Explicit\nPublic Sub P{i}()\nEnd Sub\n").into_bytes()))
        .collect()
}

#[test]
fn updating_many_files_binds_each_module_once() {
    let mut s = Session::from_sources(modules(10));
    let before = s.bind_count();

    let changed: Vec<(String, Vec<u8>)> = (0..10)
        .map(|i| (format!("m{i}.bas"), format!("Option Explicit\nPublic Sub Q{i}()\nEnd Sub\n").into_bytes()))
        .collect();
    s.update_files(changed);

    assert_eq!(s.bind_count() - before, 10);
    assert_eq!(s.module_count(), 10);
    assert_eq!(s.document_symbols(0).len(), 1);
    assert_eq!(s.document_symbols(0)[0].name, "Q0");
}

#[test]
fn updating_files_one_by_one_links_the_project_each_time() {
    let mut s = Session::from_sources(modules(10));
    let before = s.bind_count();

    for i in 0..10 {
        s.update_file(&format!("m{i}.bas"), format!("Option Explicit\nPublic Sub Q{i}()\nEnd Sub\n").into_bytes());
    }

    assert_eq!(s.bind_count() - before, 100);
}

#[test]
fn a_batch_adds_new_files_and_replaces_existing_ones() {
    let mut s = Session::from_sources(modules(2));

    s.update_files(vec![
        ("m1.bas".to_string(), b"Option Explicit\nPublic Sub Changed()\nEnd Sub\n".to_vec()),
        ("extra.bas".to_string(), b"Option Explicit\nSub Run()\n    Changed\nEnd Sub\n".to_vec()),
    ]);

    assert_eq!(s.module_count(), 3);
    assert_eq!(s.module_of("m0.bas"), Some(0));
    assert_eq!(s.module_of("m1.bas"), Some(1));
    assert_eq!(s.module_of("extra.bas"), Some(2));
    // The new file's call resolves to the procedure the same batch introduced.
    assert_eq!(s.diagnostics(2).len(), 0);
}

#[test]
fn a_path_named_twice_in_a_batch_takes_the_last_source() {
    let mut s = Session::from_sources(modules(1));

    s.update_files(vec![
        ("m0.bas".to_string(), b"Option Explicit\nPublic Sub First()\nEnd Sub\n".to_vec()),
        ("m0.bas".to_string(), b"Option Explicit\nPublic Sub Second()\nEnd Sub\n".to_vec()),
    ]);

    assert_eq!(s.module_count(), 1);
    assert_eq!(s.document_symbols(0)[0].name, "Second");
}

#[test]
fn an_empty_batch_changes_nothing() {
    let mut s = Session::from_sources(modules(3));
    let before = s.bind_count();

    s.update_files(Vec::new());

    assert_eq!(s.bind_count(), before);
}

#[test]
fn removing_many_files_links_once() {
    let mut s = Session::from_sources(modules(6));
    let before = s.bind_count();

    let removed = s.remove_files(&["m1.bas".to_string(), "m4.bas".to_string(), "absent.bas".to_string()]);

    assert_eq!(removed, 2);
    assert_eq!(s.bind_count() - before, 4);
    assert_eq!(s.module_count(), 4);
    assert_eq!(s.module_of("m1.bas"), None);
    assert_eq!(s.module_of("m5.bas"), Some(3));
}

#[test]
fn removing_nothing_that_is_present_changes_nothing() {
    let mut s = Session::from_sources(modules(3));
    let before = s.bind_count();

    assert_eq!(s.remove_files(&["absent.bas".to_string()]), 0);
    assert_eq!(s.bind_count(), before);
}
