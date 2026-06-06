//! End-to-end Tier-3 check on the bundled VB6 sample (`tests/fixtures/vb6_sample`):
//! an MSChart `OleObjectBlob` (proprietary bag) + an MSHFlexGrid `MouseIcon`
//! (standard StdPicture) in one `Form1.frx`.
//!
//! The standard-type + coverage assertions are control-independent and always run.
//! The live COM-decode assertions require a registered, licensed MSChart (e.g. a
//! machine with VB6 installed) and a 32-bit PowerShell bridge; when those are
//! absent the bag resolves to an error/opaque value and those assertions are
//! skipped rather than failed — so the test is portable.

#![cfg(windows)]

use std::path::Path;
use std::process::Command;

fn fixture_frm() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("vb6_sample")
        .join("Form1.frm")
}

#[test]
fn sample_form_resolves_picture_and_decodes_mschart_bag() {
    let frm = fixture_frm();
    if !frm.exists() {
        eprintln!("SKIP: fixture not present at {}", frm.display());
        return;
    }

    let out = Command::new(env!("CARGO_BIN_EXE_vb6-lsp"))
        .args(["read-form", frm.to_str().unwrap(), "--com-decode"])
        .output()
        .expect("run read-form");
    let stdout = String::from_utf8_lossy(&out.stdout);
    let json: serde_json::Value =
        serde_json::from_str(&stdout).unwrap_or_else(|e| panic!("read-form stdout not JSON: {e}\n{stdout}"));

    let resources = json["resources"].as_array().expect("resources array");

    // (1) Control-independent: the MSHFlexGrid MouseIcon is a standard StdPicture
    //     (a 10-image .ico) and must decode losslessly.
    let mouse = resources
        .iter()
        .find(|r| r["property"] == "MouseIcon")
        .expect("MouseIcon resource");
    assert_eq!(mouse["value"]["type"], "Picture", "MouseIcon should decode");
    assert_eq!(mouse["value"]["format"], "Ico", "MouseIcon is an .ico");

    // (2) Control-independent: every byte of the .frx is attributed.
    let cov = json["coverage"]
        .as_array()
        .and_then(|c| c.first())
        .expect("coverage entry");
    assert_eq!(cov["complete"], true, "frx coverage must be complete");
    assert_eq!(cov["unexplained_bytes"], 0);
    assert_eq!(cov["overlaps"], 0);

    // (3) Tier-3: the MSChart OleObjectBlob proprietary bag. Decoded live when the
    //     control is available; otherwise skipped.
    let ole = resources
        .iter()
        .find(|r| r["property"] == "OleObjectBlob")
        .expect("OleObjectBlob resource");
    let val = &ole["value"];
    match val["type"].as_str() {
        Some("DecodedBag") => {
            let props = val["properties"].as_array().expect("decoded properties");
            let get = |k: &str| {
                props
                    .iter()
                    .find(|p| p[0] == k)
                    .and_then(|p| p[1].as_str())
                    .map(str::to_string)
            };
            // These are the recovered design-time values that lived only as binary
            // in the bag (MSChart's default 5x4 random-data grid).
            assert_eq!(get("RowCount").as_deref(), Some("5"), "props: {props:?}");
            assert_eq!(get("ColumnCount").as_deref(), Some("4"));
            assert_eq!(get("chartType").as_deref(), Some("1"));
            assert_eq!(
                val["clsid"], "{3A2B370C-BA0A-11D1-B137-0000F8753F5D}",
                "resolved MSChart coclass CLSID"
            );
        }
        other => eprintln!(
            "SKIP Tier-3 COM assertions: OleObjectBlob resolved as {other:?} \
             (registered+licensed MSChart / 32-bit bridge unavailable here)"
        ),
    }
}
