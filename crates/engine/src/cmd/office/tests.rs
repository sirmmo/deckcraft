//! `office.sync`: what an Office add-in's task pane does to the presentation through the
//! PowerPoint JavaScript API's objects.

use serde_json::{Value as Json, json};

use crate::Session;

/// An add-in manifest written for these tests from the public schema.
fn manifest(id: &str, permission: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>{id}</Id><Version>1.0.0.0</Version><ProviderName>Tests</ProviderName><DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Test pane"/><Description DefaultValue="Tests"/>
  <Hosts><Host Name="Presentation"/></Hosts>
  <DefaultSettings><SourceLocation DefaultValue="https://tests.example/pane.html"/></DefaultSettings>
  <Permissions>{permission}</Permissions>
</OfficeApp>"#
    )
}

/// One Office add-in loaded per test (ids differ, tests run in parallel), unloaded on drop.
struct Addin(String);

impl Addin {
    fn new(id: &str, permission: &str) -> Addin {
        let m = deckcraft_plugin::office::parse_manifest(&manifest(id, permission), deckcraft_plugin::HOST_ID).unwrap();
        let plugin = m.id.clone();
        deckcraft_plugin::install(m, std::env::temp_dir()).unwrap();
        Addin(plugin)
    }
}

impl Drop for Addin {
    fn drop(&mut self) {
        deckcraft_plugin::unload(&self.0);
    }
}

fn sync(s: &mut Session, addin: &Addin, actions: Json) -> Json {
    s.execute("office.sync", &json!({"addin": addin.0, "actions": actions})).unwrap()
}

fn results(s: &mut Session, addin: &Addin, actions: Json) -> Vec<Json> {
    let r = sync(s, addin, actions);
    assert!(r.get("error").is_none(), "{r}");
    r["results"].as_array().cloned().unwrap()
}

fn slides() -> Json {
    json!([{"p": "presentation"}, {"p": "slides"}])
}

/// `presentation.slides` followed by `extra` steps.
fn path(extra: Json) -> Json {
    let mut p = slides().as_array().cloned().unwrap();
    p.extend(extra.as_array().cloned().unwrap());
    Json::Array(p)
}

/// `base` followed by `extra` steps.
fn then(base: &Json, extra: Json) -> Json {
    let mut p = base.as_array().cloned().unwrap();
    p.extend(extra.as_array().cloned().unwrap());
    Json::Array(p)
}

#[test]
fn slides_and_shapes_are_added_read_changed_and_deleted() {
    let a = Addin::new("0d5c7a1e-0000-4bcd-9eaf-000000000001", "ReadWriteDocument");
    let mut s = Session::with_new();
    let undo = s.active().unwrap().history.undo.len();
    let r = results(
        &mut s,
        &a,
        json!([{"op": "load", "path": slides(), "props": ["count"]}, {"op": "call", "path": slides(), "method": "add", "args": [{}]}]),
    );
    let before = r[0]["count"].as_u64().unwrap();
    let slide = path(json!([{"m": "getItem", "a": [r[1]["id"].as_str().unwrap()]}]));
    let shapes = then(&slide, json!([{"p": "shapes"}]));
    let r = results(
        &mut s,
        &a,
        json!([
            {"op": "load", "path": slides(), "props": ["count"]},
            {"op": "call", "path": shapes, "method": "addTextBox", "args": ["Hello", {"left": 50, "top": 60, "width": 300, "height": 80}]},
            {"op": "call", "path": shapes, "method": "addGeometricShape", "args": ["ellipse", {"left": 400, "top": 100, "width": 120, "height": 120}]},
        ]),
    );
    assert_eq!(r[0]["count"].as_u64().unwrap(), before + 1);
    let text_box = r[1]["id"].as_str().unwrap().to_string();
    let ellipse = r[2]["id"].as_str().unwrap().to_string();
    let shape = |id: &str, rest: Json| then(&then(&shapes, json!([{"m": "getItem", "a": [id]}])), rest);
    let text = shape(&text_box, json!([{"p": "textFrame"}, {"p": "textRange"}]));
    let font = then(&text, json!([{"p": "font"}]));
    let r = results(
        &mut s,
        &a,
        json!([
            {"op": "set", "path": text, "prop": "text", "value": "Hello, world"},
            {"op": "set", "path": font, "prop": "bold", "value": true},
            {"op": "set", "path": font, "prop": "color", "value": "#FF0000"},
            {"op": "set", "path": font, "prop": "size", "value": 32},
            {"op": "set", "path": shape(&text_box, json!([])), "prop": "left", "value": 72},
            {"op": "call", "path": shape(&ellipse, json!([{"p": "fill"}])), "method": "setSolidColor", "args": ["#00FF00"]},
            {"op": "set", "path": shape(&ellipse, json!([])), "prop": "name", "value": "Green dot"},
            {"op": "load", "path": text, "props": ["text", "length"]},
            {"op": "load", "path": font, "props": ["bold", "color", "size", "italic"]},
            {"op": "load", "path": shape(&text_box, json!([])), "props": ["left", "top", "type"]},
            {"op": "load", "path": shape(&ellipse, json!([{"p": "fill"}])), "props": ["type", "foregroundColor"]},
            {"op": "load", "path": shapes, "props": ["items", "name"]},
            {"op": "load", "path": then(&text, json!([{"m": "getSubstring", "a": [7, 5]}])), "props": ["text"]},
        ]),
    );
    assert_eq!(r[7], json!({"text": "Hello, world", "length": 12}));
    assert_eq!(r[8], json!({"bold": true, "color": "#FF0000", "size": 32.0, "italic": false}));
    assert_eq!((r[9]["left"].as_f64(), r[9]["top"].as_f64(), r[9]["type"].as_str()), (Some(72.0), Some(60.0), Some("textBox")));
    assert_eq!(r[10], json!({"type": "solid", "foregroundColor": "#00FF00"}));
    let items = r[11]["items"].as_array().unwrap();
    assert!(items.iter().any(|i| i["id"] == json!(ellipse) && i["name"] == "Green dot"), "{items:?}");
    assert_eq!(r[12]["text"], "world");
    // A shape added in the same sync is named by that action's result.
    let r = results(
        &mut s,
        &a,
        json!([
            {"op": "call", "path": shapes, "method": "addGeometricShape", "args": ["star5", {}]},
            {"op": "set", "path": then(&shapes, json!([{"m": "getItem", "a": [{"result": 0}]}])), "prop": "name", "value": "Star"},
            {"op": "load", "path": then(&shapes, json!([{"m": "getItem", "a": [{"result": 0}]}])), "props": ["name", "id"]},
        ]),
    );
    assert_eq!(r[2]["name"], "Star");
    assert_eq!(r[2]["id"], r[0]["id"]);
    s.execute("edit.undo", &json!({})).unwrap();
    // Each sync that changed something is one undo step.
    assert_eq!(s.active().unwrap().history.undo.len(), undo + 3);
    s.execute("edit.undo", &json!({})).unwrap();
    let r = results(&mut s, &a, json!([{"op": "load", "path": text, "props": ["text"]}]));
    assert_eq!(r[0]["text"], "Hello");
    // Deleting a shape, then the slide.
    results(
        &mut s,
        &a,
        json!([{"op": "call", "path": shape(&ellipse, json!([])), "method": "delete"}, {"op": "call", "path": slide, "method": "delete"}]),
    );
    let r = results(&mut s, &a, json!([{"op": "load", "path": slides(), "props": ["count"]}]));
    assert_eq!(r[0]["count"].as_u64().unwrap(), before);
}

#[test]
fn a_failing_action_changes_nothing_and_says_which() {
    let a = Addin::new("0d5c7a1e-0000-4bcd-9eaf-000000000002", "ReadWriteDocument");
    let mut s = Session::with_new();
    let doc = s.active().unwrap().doc.clone();
    let r = sync(
        &mut s,
        &a,
        json!([
            {"op": "call", "path": slides(), "method": "add", "args": [{}]},
            {"op": "load", "path": path(json!([{"m": "getItem", "a": ["999999"]}])), "props": ["id"]},
        ]),
    );
    assert_eq!(r["error"]["code"], "ItemNotFound", "{r}");
    assert_eq!(r["error"]["action"], 1);
    assert!(std::sync::Arc::ptr_eq(&doc, &s.active().unwrap().doc), "the added slide was taken back");
    for bad in [
        json!({"op": "load", "path": path(json!([{"m": "getItemAt", "a": [-1]}])), "props": ["id"]}),
        json!({"op": "load", "path": path(json!([{"m": "getItemAt", "a": [0]}])), "props": ["nope"]}),
        json!({"op": "load", "path": [{"p": "nowhere"}], "props": ["id"]}),
        json!({"op": "call", "path": path(json!([{"m": "getItemAt", "a": [0]}, {"p": "shapes"}])), "method": "addGeometricShape", "args": ["notAShape"]}),
        json!({"op": "set", "path": path(json!([{"m": "getItemAt", "a": [0]}])), "prop": "id", "value": 3}),
    ] {
        let r = sync(&mut s, &a, json!([bad]));
        assert!(r["error"]["code"].is_string(), "{bad} → {r}");
    }
}

#[test]
fn permissions_bound_what_an_add_in_may_do() {
    let a = Addin::new("0d5c7a1e-0000-4bcd-9eaf-000000000003", "ReadDocument");
    let mut s = Session::with_new();
    let r = results(&mut s, &a, json!([{"op": "load", "path": slides(), "props": ["count"]}]));
    assert!(r[0]["count"].as_u64().is_some());
    let r = sync(&mut s, &a, json!([{"op": "call", "path": slides(), "method": "add", "args": [{}]}]));
    assert_eq!(r["error"]["code"], "AccessDenied");
    let none = Addin::new("0d5c7a1e-0000-4bcd-9eaf-000000000004", "Restricted");
    let r = sync(&mut s, &none, json!([{"op": "load", "path": slides(), "props": ["count"]}]));
    assert_eq!(r["error"]["code"], "AccessDenied");
    // Settings need no permission.
    let settings = json!([{"p": "settings"}]);
    results(&mut s, &none, json!([{"op": "call", "path": settings, "method": "save", "args": [{"k": 1}]}]));
    let r = results(&mut s, &none, json!([{"op": "load", "path": settings, "props": []}]));
    assert_eq!(r[0], json!({"k": 1}));
    assert!(s.execute("office.sync", &json!({"addin": "nope", "actions": []})).is_err());
}

#[test]
fn the_selection_and_the_selected_text() {
    let a = Addin::new("0d5c7a1e-0000-4bcd-9eaf-000000000005", "ReadWriteDocument");
    let mut s = Session::with_new();
    s.execute("slide.new", &json!({"layout": "blank"})).unwrap();
    let id = s.execute("insert.textBox", &json!({"rect": [10, 10, 200, 50], "text": "Pick me"})).unwrap()["id"].as_u64().unwrap();
    s.execute("selection.select", &json!({"ids": [id]})).unwrap();
    let r = results(
        &mut s,
        &a,
        json!([
            {"op": "load", "path": [{"p": "presentation"}, {"m": "getSelectedShapes"}], "props": ["items", "name"]},
            {"op": "load", "path": [{"p": "presentation"}, {"m": "getSelectedSlides"}], "props": ["count"]},
            {"op": "call", "path": [{"p": "document"}], "method": "getSelectedData", "args": ["text"]},
        ]),
    );
    assert_eq!(r[0]["items"][0]["id"], json!(id.to_string()));
    assert_eq!(r[1]["count"], 1);
    assert_eq!(r[2], json!("Pick me"));
    // Layouts by name, to add a slide with one (as VBA's Slides.Add(…, ppLayoutBlank) does).
    let layouts = json!([{"p": "presentation"}, {"p": "slideMasters"}, {"m": "getItemAt", "a": [0]}, {"p": "layouts"}]);
    let r = results(&mut s, &a, json!([{"op": "load", "path": layouts, "props": ["items", "name"]}]));
    let blank = r[0]["items"].as_array().unwrap().iter().find(|l| l["name"] == "Blank").map(|l| l["id"].clone()).unwrap();
    let r = results(&mut s, &a, json!([{"op": "call", "path": slides(), "method": "add", "args": [{"layoutId": blank, "index": 0}]}]));
    let new = r[0]["id"].as_str().unwrap().parse::<u32>().unwrap();
    assert_eq!(s.active().unwrap().doc.slides[0].id.0, new, "added where asked");
    // Text frames are for shapes that hold text.
    let line = results(
        &mut s,
        &a,
        json!([{"op": "call", "path": path(json!([{"m": "getItemAt", "a": [0]}, {"p": "shapes"}])), "method": "addLine", "args": ["straight", {}]}]),
    )[0]["id"]
        .clone();
    let r = sync(
        &mut s,
        &a,
        json!([{"op": "load", "path": path(json!([{"m": "getItemAt", "a": [0]}, {"p": "shapes"}, {"m": "getItem", "a": [line]}, {"p": "textFrame"}])), "props": ["hasText"]}]),
    );
    assert_eq!(r["error"]["code"], "InvalidArgument", "{r}");
    let first = results(&mut s, &a, json!([{"op": "load", "path": slides(), "props": ["items"]}]))[0]["items"][0]["id"].clone();
    results(&mut s, &a, json!([{"op": "call", "path": [{"p": "presentation"}], "method": "setSelectedSlides", "args": [[first]]}]));
    assert_eq!(s.active().unwrap().selection.slide, 0);
}
