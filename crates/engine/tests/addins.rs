//! Add-ins and macros through a real add-in process (protocol 1.1): the add-in reads and writes
//! the presentation while it works (`host/sync`), keeps its code in the presentation's
//! `vbaProject.bin`, and every command or macro is one undo step. The add-in is this test binary
//! started again with `--serve` (no libtest harness, so standard output is the protocol's alone).

use std::io::{BufRead, Write};

use deckcraft_engine::Session;
use serde_json::{Value as Json, json};

const PLUGIN: &str = "org.deckcraft.test-addin";

fn main() {
    if std::env::args().any(|a| a == "--serve") {
        serve();
        return;
    }
    let tests: [(&str, fn()); 5] = [
        ("commands_answer_engine_commands_as_one_step", commands_answer_engine_commands_as_one_step),
        ("macros_drive_the_presentation_as_one_undo_step", macros_drive_the_presentation_as_one_undo_step),
        ("a_failing_macro_changes_nothing", a_failing_macro_changes_nothing),
        ("macro_security_gates_running", macro_security_gates_running),
        ("modules_live_in_the_presentations_vba_project", modules_live_in_the_presentations_vba_project),
    ];
    for (name, t) in tests {
        t();
        println!("test {name} ... ok");
    }
}

/// The add-in. Commands: `title` adds a slide titled after the context's slide count; `broken`
/// adds a slide and then names a command that doesn't exist. Macros: `Module1.Stamp` adds a text
/// box on the first slide through `host/sync` and says so; `Module1.Broken` adds a slide, then
/// asks for a slide that isn't there. Modules are the vbaProject's bytes taken as UTF-8 text.
fn serve() {
    let stdin = std::io::stdin();
    let mut lines = stdin.lock().lines();
    let mut out = std::io::stdout();
    let mut next = 0u32;
    // Sends a host request and waits for its answer.
    let mut host = |out: &mut std::io::Stdout, lines: &mut std::io::Lines<std::io::StdinLock<'_>>, method: &str, params: Json| -> Json {
        next += 1;
        let id = format!("h{next}");
        writeln!(out, "{}", json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params})).unwrap();
        out.flush().unwrap();
        for line in lines.by_ref() {
            let msg: Json = serde_json::from_str(&line.unwrap()).unwrap();
            if msg["id"] == json!(id) {
                return msg;
            }
        }
        Json::Null
    };
    let slides = json!([{"p": "presentation"}, {"p": "slides"}]);
    let first_shapes = json!([{"p": "presentation"}, {"p": "slides"}, {"m": "getItemAt", "a": [0]}, {"p": "shapes"}]);
    while let Some(Ok(line)) = lines.next() {
        let Ok(msg) = serde_json::from_str::<Json>(&line) else { continue };
        let Some(id) = msg.get("id").cloned() else { break };
        let p = msg["params"].clone();
        let answer = match msg["method"].as_str() {
            Some("initialize") => json!({"result": {"protocolVersion": 1}}),
            Some("commands/run") if p["command"] == "title" => {
                let n = p["context"]["slides"].as_u64().unwrap_or(0);
                json!({"result": {"commands": [{"command": "slide.new", "params": {"layout": "title", "title": format!("After {n}")}}], "message": "added"}})
            }
            Some("commands/run") if p["command"] == "broken" => {
                json!({"result": {"commands": [{"command": "slide.new", "params": {"layout": "blank"}}, {"command": "no.suchCommand", "params": {}}]}})
            }
            Some("macros/list") => {
                json!({"result": {"macros": [{"name": "Module1.Stamp", "module": "Module1", "procedure": "Stamp"}, {"name": "Module1.Broken", "module": "Module1", "procedure": "Broken"}]}})
            }
            Some("macros/run") if p["name"] == "Module1.Stamp" || p["name"] == "Stamp" => {
                let r = host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "load", "path": slides, "props": ["count"]}]}));
                let n = r["result"]["results"][0]["count"].as_u64().unwrap_or(0);
                let tag = p["control"]["tag"].as_str().unwrap_or("").to_string();
                host(
                    &mut out,
                    &mut lines,
                    "host/sync",
                    json!({"actions": [{"op": "call", "path": first_shapes, "method": "addTextBox", "args": [format!("{n} slides{tag}"), {"left": 10, "top": 10, "width": 200, "height": 40}]}]}),
                );
                host(&mut out, &mut lines, "host/message", json!({"text": format!("stamped {n}")}));
                json!({"result": {}})
            }
            Some("macros/run") if p["name"] == "Module1.Broken" => {
                host(&mut out, &mut lines, "host/sync", json!({"actions": [{"op": "call", "path": slides, "method": "add", "args": [{}]}]}));
                let bad = host(
                    &mut out,
                    &mut lines,
                    "host/sync",
                    json!({"actions": [{"op": "load", "path": [{"p": "presentation"}, {"p": "slides"}, {"m": "getItemAt", "a": [999]}], "props": ["id"]}]}),
                );
                json!({"error": {"code": 1, "message": format!("Run-time error '9': Subscript out of range ({})", bad["error"]["data"]["code"].as_str().unwrap_or(""))}})
            }
            Some("macros/modules") => {
                let r = host(&mut out, &mut lines, "host/vbaProject/get", json!({}));
                let text = r["result"]["base64"]
                    .as_str()
                    .map(|b| String::from_utf8(deckcraft_engine::cmd::base64_decode(b).unwrap()).unwrap())
                    .unwrap_or_default();
                json!({"result": {"modules": if text.is_empty() { json!([]) } else { json!([{"name": "Module1", "kind": "standard", "source": text}]) }}})
            }
            Some("macros/saveModule") => {
                let b64 = deckcraft_engine::cmd::base64_encode(p["source"].as_str().unwrap_or("").as_bytes());
                host(&mut out, &mut lines, "host/vbaProject/set", json!({"base64": b64}));
                json!({"result": {}})
            }
            _ => json!({"result": {}}),
        };
        let mut reply = json!({"jsonrpc": "2.0", "id": id});
        if let (Some(r), Json::Object(a)) = (reply.as_object_mut(), answer) {
            r.extend(a);
        }
        if writeln!(out, "{reply}").and_then(|_| out.flush()).is_err() {
            break;
        }
    }
}

fn session() -> Session {
    let mut s = Session::with_new();
    let dir = std::env::temp_dir().join(format!("deckcraft-engine-addins-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let exe = std::env::current_exe().unwrap();
    let manifest = json!({
        "id": PLUGIN, "name": "Test add-in", "macros": {"language": "TestBasic"}, "run": {"command": exe, "args": ["--serve"]},
        "commands": [{"id": "title", "label": "Title"}, {"id": "broken", "label": "Broken"}]
    });
    std::fs::write(dir.join("plugin.json"), manifest.to_string()).unwrap();
    s.execute("plugins.load", &json!({"path": dir})).unwrap();
    s
}

fn slides(s: &Session) -> usize {
    s.active().unwrap().doc.slides.len()
}

fn commands_answer_engine_commands_as_one_step() {
    let mut s = session();
    let (n, undo) = (slides(&s), s.active().unwrap().history.undo.len());
    let r = s.execute("plugins.run", &json!({"plugin": PLUGIN, "command": "title"})).unwrap();
    assert_eq!(r["ran"], 1);
    assert_eq!(slides(&s), n + 1);
    assert_eq!(s.active().unwrap().doc.slides.last().unwrap().title(), format!("After {n}"));
    assert_eq!(s.active().unwrap().history.undo.len(), undo + 1);
    let doc = s.active().unwrap().doc.clone();
    let e = s.execute("plugins.run", &json!({"plugin": PLUGIN, "command": "broken"})).unwrap_err().to_string();
    assert!(e.contains("nothing was changed"), "{e}");
    assert!(std::sync::Arc::ptr_eq(&doc, &s.active().unwrap().doc));
    assert!(s.execute("plugins.run", &json!({"plugin": PLUGIN, "command": "nope"})).is_err());
    assert!(deckcraft_engine::plugin::unload(PLUGIN));
}

fn macros_drive_the_presentation_as_one_undo_step() {
    let mut s = session();
    let list = s.execute("developer.macros", &json!({})).unwrap();
    assert_eq!(list["macros"][0]["name"], "Module1.Stamp");
    assert_eq!(list["macros"][0]["language"], "TestBasic");
    assert_eq!(list["enabled"], false, "a presentation's macros are off until enabled");
    s.execute("developer.enableMacros", &json!({})).unwrap();
    let undo = s.active().unwrap().history.undo.len();
    let shapes = s.active().unwrap().doc.slides[0].shapes.len();
    s.execute("developer.runMacro", &json!({"name": "Module1.Stamp", "control": {"id": "b1", "tag": "!"}})).unwrap();
    let first = &s.active().unwrap().doc.slides[0];
    assert_eq!(first.shapes.len(), shapes + 1);
    let n = slides(&s);
    assert_eq!(first.shapes.last().unwrap().text.as_ref().unwrap().text(), format!("{n} slides!"));
    assert_eq!(s.active().unwrap().history.undo.len(), undo + 1, "one undo step for the whole macro");
    assert!(s.ui_requests.iter().any(|q| matches!(q, deckcraft_engine::UiRequest::Message { text } if text == &format!("stamped {n}"))));
    s.execute("edit.undo", &json!({})).unwrap();
    assert_eq!(s.active().unwrap().doc.slides[0].shapes.len(), shapes);
    assert!(deckcraft_engine::plugin::unload(PLUGIN));
}

fn a_failing_macro_changes_nothing() {
    let mut s = session();
    s.execute("developer.enableMacros", &json!({})).unwrap();
    let (n, undo) = (slides(&s), s.active().unwrap().history.undo.len());
    let e = s.execute("developer.runMacro", &json!({"name": "Module1.Broken"})).unwrap_err().to_string();
    assert!(e.contains("Subscript out of range") && e.contains("ItemNotFound"), "{e}");
    assert_eq!(slides(&s), n, "the slide added before the failure was taken back");
    assert_eq!(s.active().unwrap().history.undo.len(), undo);
    assert!(deckcraft_engine::plugin::unload(PLUGIN));
}

fn macro_security_gates_running() {
    let mut s = session();
    assert!(s.execute("developer.runMacro", &json!({"name": "Stamp"})).unwrap_err().to_string().contains("disabled"));
    s.execute("developer.macroSecurity", &json!({"level": "enable"})).unwrap();
    assert!(s.execute("developer.runMacro", &json!({"name": "Stamp"})).is_ok());
    s.execute("developer.macroSecurity", &json!({"level": "disable"})).unwrap();
    s.execute("developer.enableMacros", &json!({})).unwrap();
    assert!(s.execute("developer.runMacro", &json!({"name": "Stamp"})).is_err(), "Disable wins over a presentation's Enable Content");
    assert!(s.execute("developer.macroSecurity", &json!({"level": "maybe"})).is_err());
    assert!(deckcraft_engine::plugin::unload(PLUGIN));
    assert!(s.execute("developer.macros", &json!({})).unwrap()["macros"].as_array().unwrap().is_empty());
    assert!(s.execute("developer.runMacro", &json!({"name": "Stamp"})).is_err());
}

fn modules_live_in_the_presentations_vba_project() {
    let mut s = session();
    assert_eq!(s.execute("developer.modules", &json!({})).unwrap()["modules"], json!([]));
    s.execute("developer.saveModule", &json!({"name": "Module1", "source": "Sub Hello()\r\nEnd Sub"})).unwrap();
    assert_eq!(s.active().unwrap().doc.vba_project.as_deref().map(Vec::as_slice), Some(&b"Sub Hello()\r\nEnd Sub"[..]));
    assert_eq!(s.execute("developer.modules", &json!({})).unwrap()["modules"][0]["source"], "Sub Hello()\r\nEnd Sub");
    // Saving a module is undoable, like any change to the presentation.
    s.execute("edit.undo", &json!({})).unwrap();
    assert!(s.active().unwrap().doc.vba_project.is_none());
    assert!(deckcraft_engine::plugin::unload(PLUGIN));
}
