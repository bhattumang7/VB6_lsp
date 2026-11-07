//! Opening a project in the live language server: members are analysed whether or not the
//! editor has them open, closing a file goes back to the disk, and a change that reaches other
//! files refreshes their diagnostics.

mod common;

use common::{published_for, Harness, TempDir};
use serde_json::{json, Value};

const MAIN: &str = "Option Explicit\nSub Run()\n    Greet\nEnd Sub\n";
const UTIL: &str = "Option Explicit\nPublic Sub Greet()\nEnd Sub\n";
const UTIL_WITHOUT_GREET: &str = "Option Explicit\nPublic Sub Other()\nEnd Sub\n";

fn app(dir: &TempDir) -> std::path::PathBuf {
    dir.write("modMain.bas", MAIN);
    dir.write("modUtil.bas", UTIL);
    dir.write(
        "App.vbp",
        "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\nModule=modUtil; modUtil.bas\n",
    )
}

async fn start_with_project(dir: &TempDir, vbp: &std::path::Path) -> Harness {
    let _ = dir;
    let (mut h, _) = Harness::start_with(json!({
        "capabilities": {},
        "processId": null,
        "initializationOptions": { "project": vbp.to_string_lossy() },
    }))
    .await;
    h.wait_status("loaded").await;
    h
}

/// The diagnostic a call to the undefined `Greet` produces in `MAIN`.
fn undeclared_greet(d: &Value) {
    assert_eq!(d["range"], json!({ "start": { "line": 2, "character": 4 }, "end": { "line": 2, "character": 9 } }));
    assert_eq!(d["severity"], json!(1));
    assert_eq!(d["source"], json!("vb6-lsp"));
    assert_eq!(d["message"], json!("Sub or Function not defined"));
}

#[tokio::test]
async fn without_a_project_a_name_from_another_file_is_undefined() {
    let dir = TempDir::new("noproj");
    app(&dir);
    let (mut h, _) = Harness::start().await;

    let diags = h.open(&dir.uri("modMain.bas"), MAIN).await;

    let items = diags["diagnostics"].as_array().unwrap();
    assert_eq!(items.len(), 1);
    undeclared_greet(&items[0]);
}

#[tokio::test]
async fn opening_one_file_analyses_the_whole_project() {
    let dir = TempDir::new("proj");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let main = dir.uri("modMain.bas");

    let diags = h.open(&main, MAIN).await;

    assert_eq!(diags["uri"], json!(main));
    assert_eq!(diags["version"], json!(1));
    assert_eq!(diags["diagnostics"], json!([]));
    // The definition is in a file the editor never opened.
    let loc = h.definition(&main, 2, 5).await;
    assert_eq!(loc["uri"], json!(dir.uri("modUtil.bas")));
    assert_eq!(loc["range"]["start"], json!({ "line": 1, "character": 11 }));
}

#[tokio::test]
async fn project_load_announces_its_progress_and_result() {
    let dir = TempDir::new("status");
    let vbp = app(&dir);
    let (mut h, _) = Harness::start_with(json!({
        "capabilities": {}, "processId": null,
        "initializationOptions": { "project": vbp.to_string_lossy() },
    }))
    .await;

    let loading = h.wait_notification("vb6/projectStatus").await;
    let loaded = h.wait_notification("vb6/projectStatus").await;

    assert_eq!(loading["state"], json!("loading"));
    assert_eq!(loading["path"], json!(vbp.display().to_string()));
    assert_eq!(
        loaded,
        json!({
            "state": "loaded", "path": vbp.display().to_string(), "name": "App", "isGroup": false,
            "projectCount": 1, "memberCount": 2, "missingCount": 0, "errorCount": 0,
            "warningCount": 0, "informationCount": 0, "message": null,
        })
    );
}

#[tokio::test]
async fn closed_members_get_their_diagnostics_when_the_project_loads() {
    let dir = TempDir::new("closedmembers");
    dir.write("modMain.bas", MAIN);
    dir.write("modUtil.bas", UTIL_WITHOUT_GREET);
    let vbp = dir.write("App.vbp", "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\nModule=modUtil; modUtil.bas\n");
    let (mut h, _) = Harness::start_with(json!({
        "capabilities": {}, "processId": null,
        "initializationOptions": { "project": vbp.to_string_lossy() },
    }))
    .await;

    // Nothing is open, yet the broken call in modMain is reported.
    let diags = h.diagnostics_of(&dir.uri("modMain.bas")).await;

    assert_eq!(diags.len(), 1);
    undeclared_greet(&diags[0]);
}

#[tokio::test]
async fn an_edit_that_breaks_another_file_refreshes_that_files_diagnostics() {
    let dir = TempDir::new("republish");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let util = dir.uri("modUtil.bas");
    h.open(&util, UTIL).await;

    h.change_full(&util, 2, UTIL_WITHOUT_GREET).await;

    let main_diags = h.diagnostics_of(&dir.uri("modMain.bas")).await;
    // modMain was never opened: its error appears because modUtil lost `Greet`.
    assert_eq!(main_diags.len(), 1, "{main_diags:?}");
    undeclared_greet(&main_diags[0]);
}

#[tokio::test]
async fn closing_a_project_file_goes_back_to_the_disk() {
    let dir = TempDir::new("close");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let (main, util) = (dir.uri("modMain.bas"), dir.uri("modUtil.bas"));
    h.open(&main, MAIN).await;
    h.open(&util, UTIL).await;
    h.change_full(&util, 2, UTIL_WITHOUT_GREET).await;
    let broken = h.diagnostics_of(&main).await;
    assert_eq!(broken.len(), 1);

    // The unsaved edit is abandoned: Greet is back, and the file stays part of the project.
    h.close(&util).await;
    let (_, notes) = h.execute_settled("vb6.projectInfo", None).await;

    let main_after = published_for(&notes, &main);
    assert_eq!(main_after.last(), Some(&Vec::<Value>::new()), "{notes:#?}");
    let loc = h.definition(&main, 2, 5).await;
    assert_eq!(loc["uri"], json!(util));
}

#[tokio::test]
async fn closing_a_file_outside_the_project_removes_it_and_clears_its_diagnostics() {
    let dir = TempDir::new("loose");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let stray = dir.uri("stray.bas");
    let first = h.open(&stray, "Option Explicit\nSub A()\n    ghost\nEnd Sub\n").await;
    assert_eq!(first["diagnostics"].as_array().unwrap().len(), 1);

    h.close(&stray).await;

    let after = h.diagnostics_of(&stray).await;
    assert_eq!(after, Vec::<Value>::new());
}

#[cfg(windows)]
#[tokio::test]
async fn the_editors_encoded_uri_names_the_same_file_as_the_project() {
    let dir = TempDir::new("uri");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    // `file:///C:/...` as the project knows it; the editor writes `file:///c%3A/...`.
    let encoded = dir.uri("modMain.bas").replacen("file:///C:", "file:///c%3A", 1);
    assert_ne!(encoded, dir.uri("modMain.bas"));

    let diags = h.open(&encoded, MAIN).await;

    assert_eq!(diags["diagnostics"], json!([]));
    let loc = h.definition(&encoded, 2, 5).await;
    assert_eq!(loc["uri"], json!(dir.uri("modUtil.bas")));
}

#[tokio::test]
async fn a_member_changed_on_disk_is_reloaded() {
    let dir = TempDir::new("disk");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let (main, util) = (dir.uri("modMain.bas"), dir.uri("modUtil.bas"));
    h.open(&main, MAIN).await;

    dir.write("modUtil.bas", UTIL_WITHOUT_GREET);
    h.notify(
        "workspace/didChangeWatchedFiles",
        json!({ "changes": [{ "uri": util, "type": 2 }] }),
    )
    .await;

    let diags = h.diagnostics_of(&main).await;
    assert_eq!(diags.len(), 1);
    undeclared_greet(&diags[0]);
}

#[tokio::test]
async fn a_member_deleted_on_disk_leaves_the_analysis() {
    let dir = TempDir::new("deleted");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let (main, util) = (dir.uri("modMain.bas"), dir.uri("modUtil.bas"));
    h.open(&main, MAIN).await;
    std::fs::remove_file(dir.path("modUtil.bas")).unwrap();

    h.notify(
        "workspace/didChangeWatchedFiles",
        json!({ "changes": [{ "uri": util, "type": 3 }] }),
    )
    .await;

    let diags = h.diagnostics_of(&main).await;
    assert_eq!(diags.len(), 1);
    undeclared_greet(&diags[0]);
    let info = h.execute("vb6.projectInfo", None).await["result"].clone();
    let util_member = &info["projects"][0]["members"][1];
    assert_eq!(util_member["missing"], json!(true));
}

#[tokio::test]
async fn a_file_open_in_the_editor_is_not_replaced_by_a_disk_change() {
    let dir = TempDir::new("openwins");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let (main, util) = (dir.uri("modMain.bas"), dir.uri("modUtil.bas"));
    h.open(&main, MAIN).await;
    h.open(&util, UTIL).await;

    dir.write("modUtil.bas", UTIL_WITHOUT_GREET);
    h.notify("workspace/didChangeWatchedFiles", json!({ "changes": [{ "uri": util, "type": 2 }] })).await;
    let (_, notes) = h.execute_settled("vb6.projectInfo", None).await;

    assert_eq!(published_for(&notes, &main), Vec::<Vec<Value>>::new());
    let loc = h.definition(&main, 2, 5).await;
    assert_eq!(loc["uri"], json!(util));
}

#[tokio::test]
async fn a_changed_project_file_reloads_the_project() {
    let dir = TempDir::new("reloadvbp");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let main = dir.uri("modMain.bas");
    h.open(&main, MAIN).await;

    // modUtil is dropped from the project.
    dir.write("App.vbp", "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\n");
    h.notify(
        "workspace/didChangeWatchedFiles",
        json!({ "changes": [{ "uri": dir.uri("App.vbp"), "type": 2 }] }),
    )
    .await;

    let diags = h.diagnostics_of(&main).await;
    assert_eq!(diags.len(), 1);
    undeclared_greet(&diags[0]);
}

#[tokio::test]
async fn the_open_project_command_describes_the_project() {
    let dir = TempDir::new("cmd");
    let vbp = app(&dir);
    let (mut h, init) = Harness::start().await;
    assert_eq!(
        init["capabilities"]["executeCommandProvider"]["commands"],
        json!(["vb6.openProject", "vb6.closeProject", "vb6.projectInfo", "vb6.projectPreview"])
    );
    assert_eq!(h.execute("vb6.projectInfo", None).await["result"], Value::Null);

    let resp = h.execute("vb6.openProject", Some(&vbp.to_string_lossy())).await;

    let main_path = dir.path("modMain.bas").display().to_string();
    let util_path = dir.path("modUtil.bas").display().to_string();
    assert_eq!(
        resp["result"],
        json!({
            "path": vbp.display().to_string(),
            "uri": dir.uri("App.vbp"),
            "name": "App",
            "isGroup": false,
            "projects": [{
                "path": vbp.display().to_string(),
                "uri": dir.uri("App.vbp"),
                "name": "App",
                "kind": "exe",
                "listed": true,
                "startup": true,
                "members": [
                    { "name": "modMain", "extension": "bas", "path": main_path, "uri": dir.uri("modMain.bas"), "line": 3, "missing": false },
                    { "name": "modUtil", "extension": "bas", "path": util_path, "uri": dir.uri("modUtil.bas"), "line": 4, "missing": false },
                ],
                "references": [],
            }],
            "problems": [],
        })
    );
    assert_eq!(h.execute("vb6.projectInfo", None).await["result"], resp["result"]);
}

#[tokio::test]
async fn the_open_project_command_takes_a_file_uri() {
    let dir = TempDir::new("cmduri");
    let vbp = app(&dir);
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.openProject", Some(&dir.uri("App.vbp"))).await;

    assert_eq!(resp["result"]["path"], json!(vbp.display().to_string()));
}

#[tokio::test]
async fn opening_something_that_is_not_a_project_reports_why() {
    let dir = TempDir::new("badopen");
    let bas = dir.write("a.bas", "Option Explicit\n");
    let (mut h, _) = Harness::start().await;

    let (resp, notes) = h
        .request_collect(
            "workspace/executeCommand",
            json!({ "command": "vb6.openProject", "arguments": [bas.to_string_lossy()] }),
        )
        .await;

    assert_eq!(resp["error"]["code"], json!(1));
    let expected = format!("{}: Not a VB6 project (.vbp) or project group (.vbg)", bas.display());
    assert_eq!(resp["error"]["message"], json!(expected));
    let message = notes.iter().find(|n| n["method"] == json!("window/showMessage")).unwrap();
    assert_eq!(message["params"], json!({ "type": 1, "message": format!("Cannot open project: {expected}") }));
    let failed = notes.iter().rfind(|n| n["method"] == json!("vb6/projectStatus")).unwrap();
    assert_eq!(failed["params"]["state"], json!("error"));
    assert_eq!(failed["params"]["message"], json!(expected));
}

#[tokio::test]
async fn a_failed_open_keeps_the_project_that_was_open() {
    let dir = TempDir::new("keep");
    let vbp = app(&dir);
    let bad = dir.write("bad.vbg", "no header\n");
    let mut h = start_with_project(&dir, &vbp).await;

    let resp = h.execute("vb6.openProject", Some(&bad.to_string_lossy())).await;

    assert!(resp.get("error").is_some());
    let info = h.execute("vb6.projectInfo", None).await["result"].clone();
    assert_eq!(info["name"], json!("App"));
}

#[tokio::test]
async fn the_open_project_command_needs_a_path() {
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.openProject", None).await;

    assert_eq!(resp["error"]["code"], json!(-32602));
    assert_eq!(resp["error"]["message"], json!("vb6.openProject takes the path of a .vbp or .vbg file"));
}

#[tokio::test]
async fn an_unknown_command_is_rejected() {
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.nothing", None).await;

    assert_eq!(resp["error"]["code"], json!(-32602));
}

#[tokio::test]
async fn closing_the_project_clears_its_diagnostics_and_analyses_open_files_alone() {
    let dir = TempDir::new("closeproj");
    let vbp = app(&dir);
    let mut h = start_with_project(&dir, &vbp).await;
    let main = dir.uri("modMain.bas");
    assert_eq!(h.open(&main, MAIN).await["diagnostics"], json!([]));

    let (resp, notes) = h.execute_settled("vb6.closeProject", None).await;

    assert_eq!(resp["result"], Value::Null);
    // modMain stands alone again: Greet is undefined.
    let main_after = published_for(&notes, &main);
    assert_eq!(main_after.last().map(Vec::len), Some(1), "{notes:#?}");
    let status = notes.iter().rfind(|n| n["method"] == json!("vb6/projectStatus")).unwrap();
    assert_eq!(status["params"]["state"], json!("none"));
    assert_eq!(h.execute("vb6.projectInfo", None).await["result"], Value::Null);
}

#[tokio::test]
async fn opening_a_project_takes_over_files_that_were_open_on_their_own() {
    let dir = TempDir::new("takeover");
    let vbp = app(&dir);
    let (mut h, _) = Harness::start().await;
    let main = dir.uri("modMain.bas");
    assert_eq!(h.open(&main, MAIN).await["diagnostics"].as_array().unwrap().len(), 1);

    let (_, notes) = h.execute_settled("vb6.openProject", Some(&vbp.to_string_lossy())).await;

    assert_eq!(published_for(&notes, &main).last(), Some(&Vec::<Value>::new()), "{notes:#?}");
    let loc = h.definition(&main, 2, 5).await;
    assert_eq!(loc["uri"], json!(dir.uri("modUtil.bas")));
}

#[tokio::test]
async fn unsaved_text_of_an_open_file_is_used_when_the_project_loads() {
    let dir = TempDir::new("unsaved");
    let vbp = app(&dir);
    let (mut h, _) = Harness::start().await;
    let util = dir.uri("modUtil.bas");
    h.open(&util, UTIL_WITHOUT_GREET).await;

    let (_, notes) = h.execute_settled("vb6.openProject", Some(&vbp.to_string_lossy())).await;

    let diags = published_for(&notes, &dir.uri("modMain.bas")).pop().unwrap();
    assert_eq!(diags.len(), 1);
    undeclared_greet(&diags[0]);
}

#[tokio::test]
async fn problems_in_the_project_file_are_shown_on_their_lines() {
    let dir = TempDir::new("vbpproblems");
    dir.write("modMain.bas", MAIN);
    dir.write("modUtil.bas", UTIL);
    let vbp = dir.write(
        "App.vbp",
        "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\nModule=modGone; modGone.bas\nModule=modUtil; modUtil.bas\n\
Reference=*\\G{DEADBEEF-0000-4000-8000-000000000001}#2.5#0#C:\\msado25.tlb#ADO\n",
    );
    let (mut h, _) = Harness::start_with(json!({
        "capabilities": {}, "processId": null,
        "initializationOptions": { "project": vbp.to_string_lossy() },
    }))
    .await;

    let messages = h.until_status("loaded").await;
    let loaded = messages.last().unwrap()["params"].clone();
    let diags = published_for(&messages, &dir.uri("App.vbp")).pop().unwrap();

    assert_eq!(loaded["missingCount"], json!(1));
    assert_eq!(loaded["errorCount"], json!(1));
    assert_eq!(loaded["warningCount"], json!(1));
    assert_eq!(loaded["informationCount"], json!(0));
    assert_eq!(
        diags,
        vec![
            json!({
                "range": { "start": { "line": 3, "character": 0 }, "end": { "line": 3, "character": 27 } },
                "severity": 1,
                "source": "vb6-lsp",
                "message": format!("File not found: {}", dir.path("modGone.bas").display()),
            }),
            json!({
                "range": { "start": { "line": 5, "character": 0 }, "end": { "line": 5, "character": 76 } },
                "severity": 2,
                "source": "vb6-lsp",
                "message": common::missing_library_message("ADO {DEADBEEF-0000-4000-8000-000000000001} 2.5", r"C:\msado25.tlb"),
            }),
        ]
    );
}

#[tokio::test]
async fn a_group_is_opened_and_its_projects_see_each_others_global_classes() {
    let dir = TempDir::new("grp");
    let class = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"clsThing\"\n\
Attribute VB_GlobalNameSpace = True\nAttribute VB_Creatable = True\nAttribute VB_PredeclaredId = False\n\
Attribute VB_Exposed = True\nPublic Sub Hello()\nEnd Sub\n";
    dir.write("Lib/clsThing.cls", class);
    dir.write("Lib/Lib.vbp", "Type=OleDll\nName=\"Lib\"\nClass=clsThing; clsThing.cls\n");
    let main = "Option Explicit
Sub Run()
    Hello
End Sub
";
    dir.write("App/modMain.bas", main);
    dir.write("App/App.vbp", "Type=Exe\nReference=*\\A..\\Lib\\Lib.vbp\nModule=modMain; modMain.bas\nName=\"App\"\n");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nStartupProject=App\\App.vbp\nProject=App\\App.vbp\nProject=Lib\\Lib.vbp\n");
    let (mut h, _) = Harness::start().await;
    let resp = h.execute("vb6.openProject", Some(&vbg.to_string_lossy())).await;
    assert_eq!(resp["result"]["isGroup"], json!(true));
    assert_eq!(resp["result"]["projects"].as_array().unwrap().len(), 2);
    let main_uri = dir.uri("App/modMain.bas");
    let diags = h.open(&main_uri, main).await;
    assert_eq!(diags["diagnostics"], json!([]));

    // The library stops exposing its global class: the application's call to `Hello` no longer resolves.
    let lib_uri = dir.uri("Lib/clsThing.cls");
    h.open(&lib_uri, class).await;
    h.change_full(&lib_uri, 2, &class.replace("Attribute VB_Exposed = True", "Attribute VB_Exposed = False")).await;
    let (_, notes) = h.execute_settled("vb6.projectInfo", None).await;

    let published = published_for(&notes, &main_uri);
    assert_eq!(published.len(), 1, "the application's diagnostics are published again: {notes:#?}");
    assert_eq!(published[0].len(), 1, "{published:?}");
    assert_eq!(published[0][0]["message"], json!("Sub or Function not defined"));
}

#[tokio::test]
async fn previewing_a_project_describes_it_without_opening_it() {
    let dir = TempDir::new("preview");
    dir.write("modMain.bas", MAIN);
    let vbp = dir.write(
        "App.vbp",
        "Type=Exe\nName=\"App\"\nModule=modMain; modMain.bas\nModule=modGone; modGone.bas\n\
Reference=*\\G{DEADBEEF-0000-4000-8000-000000000001}#2.5#0#C:\\msado25.tlb#ADO\n",
    );
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.projectPreview", Some(&vbp.to_string_lossy())).await;

    let result = &resp["result"];
    assert_eq!(result["name"], json!("App"));
    assert_eq!(result["isGroup"], json!(false));
    let members = &result["projects"][0]["members"];
    assert_eq!(members[0]["missing"], json!(false));
    assert_eq!(members[1]["missing"], json!(true));
    assert_eq!(members[1]["line"], json!(4));
    assert_eq!(
        result["projects"][0]["references"],
        json!([{
            "label": "ADO", "kind": "typeLibrary", "status": "notLoaded", "line": 5,
            "detail": "{DEADBEEF-0000-4000-8000-000000000001}",
        }])
    );
    let messages: Vec<&str> = result["problems"].as_array().unwrap().iter().map(|p| p["message"].as_str().unwrap()).collect();
    assert_eq!(
        messages,
        vec![
            format!("File not found: {}", dir.path("modGone.bas").display()).as_str(),
            common::missing_library_message("ADO {DEADBEEF-0000-4000-8000-000000000001} 2.5", r"C:\msado25.tlb").as_str(),
        ]
    );
    // Nothing was opened.
    assert_eq!(h.execute("vb6.projectInfo", None).await["result"], Value::Null);
}

#[tokio::test]
async fn previewing_a_group_lists_every_project() {
    let dir = TempDir::new("previewgrp");
    dir.write("A/A.vbp", "Type=Exe\nName=\"A\"\n");
    dir.write("B/B.vbp", "Type=OleDll\nName=\"B\"\n");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nProject=A\\A.vbp\nProject=B\\B.vbp\n");
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.projectPreview", Some(&vbg.to_string_lossy())).await;

    let names: Vec<(&str, &str)> = resp["result"]["projects"]
        .as_array()
        .unwrap()
        .iter()
        .map(|p| (p["name"].as_str().unwrap(), p["kind"].as_str().unwrap()))
        .collect();
    assert_eq!(names, vec![("A", "exe"), ("B", "oledll")]);
    assert_eq!(resp["result"]["isGroup"], json!(true));
}

#[tokio::test]
async fn previewing_something_that_is_not_a_project_reports_why() {
    let dir = TempDir::new("previewbad");
    let bas = dir.write("a.bas", "Option Explicit\n");
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.projectPreview", Some(&bas.to_string_lossy())).await;

    assert_eq!(resp["error"]["code"], json!(1));
    assert_eq!(
        resp["error"]["message"],
        json!(format!("{}: Not a VB6 project (.vbp) or project group (.vbg)", bas.display()))
    );
}

#[tokio::test]
async fn previewing_needs_a_path() {
    let (mut h, _) = Harness::start().await;

    let resp = h.execute("vb6.projectPreview", None).await;

    assert_eq!(resp["error"]["code"], json!(-32602));
    assert_eq!(resp["error"]["message"], json!("vb6.projectPreview takes the path of a .vbp or .vbg file"));
}

const LIB_CLASS: &str = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"clsThing\"\n\
Attribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\nAttribute VB_PredeclaredId = False\n\
Attribute VB_Exposed = True\nPublic Sub Hello()\nEnd Sub\nPublic Property Get Size() As Long\nEnd Property\n\
Public Property Let Size(ByVal v As Long)\nEnd Property\nPublic Total As Long\n";

const LIB_MODULE_WITH_ENUM: &str = "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"clsKinds\"\n\
Attribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = True\nAttribute VB_PredeclaredId = False\n\
Attribute VB_Exposed = True\nPublic Enum Colour\n    Red = 1\n    Blue = 2\nEnd Enum\nPublic Type Pair\n    Left As Long\n    Right As Long\nEnd Type\n";

fn range(start_line: u64, start_character: u64, end_line: u64, end_character: u64) -> Value {
    json!({
        "start": { "line": start_line, "character": start_character },
        "end": { "line": end_line, "character": end_character },
    })
}

async fn open_group_with_library(dir: &TempDir, app_source: &str) -> (Harness, String) {
    dir.write("Lib/clsThing.cls", LIB_CLASS);
    dir.write("Lib/clsKinds.cls", LIB_MODULE_WITH_ENUM);
    dir.write("Lib/Lib.vbp", "Type=OleDll\nName=\"Lib\"\nClass=clsThing; clsThing.cls\nClass=clsKinds; clsKinds.cls\n");
    dir.write("App/modMain.bas", app_source);
    dir.write("App/App.vbp", "Type=Exe\nReference=*\\A..\\Lib\\Lib.vbp\nModule=modMain; modMain.bas\nName=\"App\"\n");
    let vbg = dir.write("All.vbg", "VBGROUP 5.0\nStartupProject=App\\App.vbp\nProject=App\\App.vbp\nProject=Lib\\Lib.vbp\n");
    let (mut h, _) = Harness::start().await;
    h.execute("vb6.openProject", Some(&vbg.to_string_lossy())).await;
    let uri = dir.uri("App/modMain.bas");
    h.open(&uri, app_source).await;
    (h, uri)
}

#[tokio::test]
async fn go_to_definition_of_a_method_of_a_referenced_project_opens_its_declaration() {
    let dir = TempDir::new("defmember");
    let source = "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Hello\nEnd Sub\n";
    let (mut h, uri) = open_group_with_library(&dir, source).await;

    let found = h.definition(&uri, 3, 6).await;

    assert_eq!(found, json!({ "uri": dir.uri("Lib/clsThing.cls"), "range": range(9, 11, 9, 16) }));
}

#[tokio::test]
async fn go_to_definition_of_a_property_prefers_the_accessor_that_is_used() {
    let dir = TempDir::new("defprop");
    let source = "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Size = 3\n    Debug.Print t.Size\nEnd Sub\n";
    let (mut h, uri) = open_group_with_library(&dir, source).await;

    let assigned = h.definition(&uri, 3, 6).await;
    let read = h.definition(&uri, 4, 18).await;

    assert_eq!(assigned, json!({ "uri": dir.uri("Lib/clsThing.cls"), "range": range(13, 20, 13, 24) }));
    assert_eq!(read, json!({ "uri": dir.uri("Lib/clsThing.cls"), "range": range(11, 20, 11, 24) }));
}

#[tokio::test]
async fn go_to_definition_of_a_public_field_of_a_referenced_class_opens_its_declaration() {
    let dir = TempDir::new("deffield");
    let source = "Option Explicit\nSub Run()\n    Dim t As Lib.clsThing\n    t.Total = 1\nEnd Sub\n";
    let (mut h, uri) = open_group_with_library(&dir, source).await;

    let found = h.definition(&uri, 3, 6).await;

    assert_eq!(found, json!({ "uri": dir.uri("Lib/clsThing.cls"), "range": range(15, 7, 15, 12) }));
}

#[tokio::test]
async fn go_to_definition_of_a_referenced_enumeration_member_opens_its_declaration() {
    let dir = TempDir::new("defenum");
    let source = "Option Explicit\nSub Run()\n    Dim c As Long\n    c = Lib.Blue\nEnd Sub\n";
    let (mut h, uri) = open_group_with_library(&dir, source).await;

    let found = h.definition(&uri, 3, 12).await;

    assert_eq!(found, json!({ "uri": dir.uri("Lib/clsKinds.cls"), "range": range(11, 4, 11, 8) }));
}

#[tokio::test]
async fn go_to_definition_of_a_name_from_a_runtime_library_has_no_source_to_open() {
    let dir = TempDir::new("defruntime");
    let source = "Option Explicit\nSub Run()\n    Dim s As String\n    s = Left$(\"abc\", 1)\nEnd Sub\n";
    let (mut h, uri) = open_group_with_library(&dir, source).await;

    assert_eq!(h.definition(&uri, 3, 8).await, Value::Null);
}

#[tokio::test]
async fn go_to_definition_of_a_referenced_class_name_opens_the_start_of_its_file() {
    let dir = TempDir::new("deftype");
    let source = "Option Explicit
Sub Run()
    Dim t As Lib.clsThing
    Dim u As clsThing
End Sub
";
    let (mut h, uri) = open_group_with_library(&dir, source).await;

    let qualified = h.definition(&uri, 2, 19).await;
    let bare = h.definition(&uri, 3, 14).await;

    let start = json!({ "uri": dir.uri("Lib/clsThing.cls"), "range": range(0, 0, 0, 0) });
    assert_eq!(qualified, start);
    assert_eq!(bare, start);
}
