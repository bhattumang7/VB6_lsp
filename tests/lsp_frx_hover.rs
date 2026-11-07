//! Hovering a `"Form.frx":NNNN` property in a designer file summarises the companion blob.

mod common;

use common::{Harness, TempDir};
use serde_json::{json, Value};

fn picture(image: &[u8]) -> Vec<u8> {
    let mut b = (image.len() as u32 + 8).to_le_bytes().to_vec();
    b.extend_from_slice(b"lt\0\0");
    b.extend_from_slice(&(image.len() as u32).to_le_bytes());
    b.extend_from_slice(image);
    b
}

fn font(flags: u8, weight: u16, name: &[u8]) -> Vec<u8> {
    let mut b = vec![0x01, 0x00, 0x00, flags];
    b.extend_from_slice(&weight.to_le_bytes());
    b.extend_from_slice(&82500u32.to_le_bytes());
    b.push(name.len() as u8);
    b.extend_from_slice(name);
    b
}

/// Companion bytes and the offset each blob starts at.
fn companion() -> (Vec<u8>, [usize; 5]) {
    let mut bytes = Vec::new();
    let mut at = [0usize; 5];

    at[0] = bytes.len();
    bytes.extend(picture(&[0x42, 0x4D, 0x00, 0x00]));
    at[1] = bytes.len();
    bytes.extend(font(0x0E, 700, b"Arial"));
    at[2] = bytes.len();
    bytes.push(5);
    bytes.extend_from_slice(b"Hello");
    at[3] = bytes.len();
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0x000Bu16.to_le_bytes());
    for item in [&b"one"[..], &b"two"[..]] {
        bytes.extend_from_slice(&(item.len() as u16).to_le_bytes());
        bytes.extend_from_slice(item);
    }
    at[4] = bytes.len();
    bytes.extend(picture(&[]));
    (bytes, at)
}

fn form(at: [usize; 5]) -> String {
    format!(
        "VERSION 5.00\r\nBegin VB.Form Form1\r\n   Icon = \"Form1.frx\":{:04X}\r\n   Font = \"Form1.frx\":{:04X}\r\n   Tag = \"Form1.frx\":{:04X}\r\n   List = \"Form1.frx\":{:04X}\r\n   Picture = \"Form1.frx\":{:04X}\r\n   Caption = \"plain\"\r\nEnd\r\nAttribute VB_Name = \"Form1\"\r\n",
        at[0], at[1], at[2], at[3], at[4]
    )
}

async fn hover(h: &mut Harness, uri: &str, line: u32) -> Value {
    h.request(
        "textDocument/hover",
        json!({ "textDocument": { "uri": uri }, "position": { "line": line, "character": 5 } }),
    )
    .await
}

fn text(resp: &Value) -> String {
    resp["result"]["contents"]["value"].as_str().unwrap_or("<none>").to_string()
}

#[tokio::test]
async fn each_kind_of_blob_is_summarised() {
    let dir = TempDir::new("frxhover");
    let (bytes, at) = companion();
    std::fs::write(dir.path("Form1.frx"), &bytes).unwrap();
    let source = form(at);
    dir.write("Form1.frm", &source);
    let uri = dir.uri("Form1.frm");

    let (mut h, _) = Harness::start().await;
    h.open(&uri, &source).await;

    assert_eq!(
        text(&hover(&mut h, &uri, 2).await),
        format!("**Icon** → `Form1.frx`:0x{:04X}\n\n🖼 Bmp image · 4 bytes", at[0])
    );
    assert_eq!(
        text(&hover(&mut h, &uri, 3).await),
        format!(
            "**Font** → `Form1.frx`:0x{:04X}\n\n🔤 Font `Arial` · 8.25pt bold italic underline strikethrough",
            at[1]
        )
    );
    assert_eq!(
        text(&hover(&mut h, &uri, 4).await),
        format!("**Tag** → `Form1.frx`:0x{:04X}\n\n📝 \"Hello\"", at[2])
    );
    assert_eq!(
        text(&hover(&mut h, &uri, 5).await),
        format!("**List** → `Form1.frx`:0x{:04X}\n\n📋 2 list item(s): one, two", at[3])
    );
    assert_eq!(
        text(&hover(&mut h, &uri, 6).await),
        format!("**Picture** → `Form1.frx`:0x{:04X}\n\n∅ empty resource", at[4])
    );
}

#[tokio::test]
async fn a_missing_companion_file_is_reported() {
    let dir = TempDir::new("frxhover-missing");
    let (_, at) = companion();
    let source = form(at);
    dir.write("Form1.frm", &source);
    let uri = dir.uri("Form1.frm");

    let (mut h, _) = Harness::start().await;
    h.open(&uri, &source).await;

    assert_eq!(
        text(&hover(&mut h, &uri, 2).await),
        format!(
            "**Icon** → `Form1.frx`:0x{:04X}\n\n⚠️ companion file `Form1.frx` not found next to the form",
            at[0]
        )
    );
}

#[tokio::test]
async fn a_blob_that_does_not_decode_is_reported() {
    let dir = TempDir::new("frxhover-bad");
    let (_, at) = companion();
    std::fs::write(dir.path("Form1.frx"), [0xFFu8; 8]).unwrap();
    let source = form(at);
    dir.write("Form1.frm", &source);
    let uri = dir.uri("Form1.frm");

    let (mut h, _) = Harness::start().await;
    h.open(&uri, &source).await;

    // Offset 0 holds eight 0xFF bytes: not a picture.
    let resp = hover(&mut h, &uri, 2).await;
    assert_eq!(
        text(&resp),
        format!(
            "**Icon** → `Form1.frx`:0x{:04X}\n\n⚠️ could not decode: bad/unknown picture framing at offset 0x0",
            at[0]
        )
    );
}
