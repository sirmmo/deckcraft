//! Add-ins in the window: their ribbon groups and tabs, the Developer tab with its macros bar,
//! the Add-ins window, and an Office add-in's task pane talking to the presentation.

use std::sync::{Arc, Mutex};

use deckcraft_engine::Session;
use egui_kittest::Harness;
use egui_kittest::kittest::Queryable;
use serde_json::{Value as Json, json};

use crate::office_pane::OfficeHost;
use crate::{Services, SlideApp};

fn harness(s: Session, services: Services) -> Harness<'static, SlideApp> {
    let app = SlideApp::new(s, services);
    let mut h = Harness::builder().with_size(egui::vec2(1600.0, 900.0)).with_step_dt(1.0 / 60.0).build_ui_state(
        |ui, app: &mut SlideApp| {
            app.logic(ui.ctx());
            app.ui(ui);
        },
        app,
    );
    h.run_steps(3);
    h
}

/// Add-ins are loaded process-wide: the tests that load one take turns.
static ADDINS: Mutex<()> = Mutex::new(());

/// Unloads the add-in however the test ends, so other tests never see it.
struct Unload(&'static str);

impl Drop for Unload {
    fn drop(&mut self) {
        deckcraft_engine::plugin::unload(self.0);
    }
}

#[test]
fn an_add_ins_buttons_and_the_developer_tab_with_its_macros_bar() {
    let _one_at_a_time = ADDINS.lock().unwrap_or_else(|e| e.into_inner());
    let m = deckcraft_engine::plugin::Manifest::parse(
        &json!({
            "id": "t.ui-addin", "name": "UI Test", "macros": {"language": "VBA"}, "run": {"command": "./not-started"},
            "commands": [{"id": "go", "label": "Go Now"}],
            "ui": {"ribbon": [{"tab": "Example", "group": "Mine", "buttons": [{"command": "go"}]}, {"tab": "Insert", "group": "Mine", "buttons": [{"command": "go", "label": "Go Here", "size": "small"}]}]}
        })
        .to_string(),
    )
    .unwrap();
    deckcraft_engine::plugin::install(m, std::env::temp_dir()).unwrap();
    let _unload = Unload("t.ui-addin");
    assert!(super::developer_visible());
    assert!(super::extra_tabs(&["Insert"]).contains(&"Example".to_string()));
    let mut s = Session::with_new();
    s.edit(|doc, _| {
        doc.vba_project = Some(Arc::new(b"project".to_vec()));
        Ok(())
    })
    .unwrap();
    let mut h = harness(s, Services::default());
    // The macros bar, until Enable Content.
    assert!(h.query_by_label("Enable Content").is_some(), "a presentation with macros opens with them disabled");
    h.get_by_label("Enable Content").click();
    h.run_steps(3);
    assert!(deckcraft_engine::cmd::developer::macros_allowed(&h.state().session));
    assert!(h.query_by_label("Enable Content").is_none());
    // The add-in's buttons: on Insert, and on its own tab.
    h.state_mut().ui.tab = "Insert".into();
    h.run_steps(3);
    assert!(h.query_by_label("Go Here").is_some());
    assert!(h.query_by_label("Add-ins").is_some());
    h.state_mut().ui.tab = "Example".into();
    h.run_steps(3);
    assert!(h.query_by_label("Go Now").is_some());
    // Pressing it runs the add-in's command (which can't start here) and says why.
    h.get_by_label("Go Now").click();
    h.run_steps(3);
    let said = h.state().status.as_ref().map(|(m, _)| m.clone()).unwrap_or_default();
    assert!(said.contains("not-started") || said.contains("start"), "{said}");
    // Developer: Macro Security, and the Add-ins window listing the add-in.
    h.state_mut().ui.tab = "Developer".into();
    h.run_steps(3);
    h.get_by_label("Macro Security").click();
    h.run_steps(3);
    assert!(h.state().addins.security);
    h.state_mut().addins.security = false;
    h.get_by_label("Add-ins").click();
    h.run_steps(3);
    assert!(h.query_by_label("UI Test").is_some(), "the Add-ins window lists it");
    h.get_by_label("Deactivate").click();
    h.run_steps(3);
    assert!(!deckcraft_engine::plugin::get("t.ui-addin").unwrap().enabled());
    h.state_mut().ui.tab = "Example".into();
    h.run_steps(3);
    assert!(h.query_by_label("Go Now").is_none(), "an add-in turned off takes its buttons away");
}

#[test]
fn a_presentations_ribbon_buttons_show_on_their_tabs_and_run_their_macros() {
    use deckcraft_model::{Ribbon, RibbonButton, RibbonGroup, RibbonTab};
    // No macro add-in loaded meanwhile: the buttons show without one.
    let _one_at_a_time = ADDINS.lock().unwrap_or_else(|e| e.into_inner());
    let button = |id: &str, label: &str| RibbonButton {
        id: id.into(),
        label: label.into(),
        on_action: format!("Module1.{label}"),
        large: true,
        ..Default::default()
    };
    let group = |b| vec![RibbonGroup { id: "g".into(), label: "G".into(), buttons: vec![b] }];
    let mut s = Session::with_new();
    s.edit(|doc, _| {
        doc.ribbon = Some(Ribbon {
            tabs: vec![
                RibbonTab { id: "tTools".into(), label: "Tools".into(), groups: group(button("b1", "Go")), ..Default::default() },
                RibbonTab { builtin: Some("TabHome".into()), groups: group(button("b2", "Hello")), ..Default::default() },
            ],
        });
        Ok(())
    })
    .unwrap();
    let mut h = harness(s, Services::default());
    // On Home, after Home's own groups.
    assert!(h.query_by_label("Hello").is_some(), "the presentation's button on Home");
    assert!(h.query_by_label("Go").is_none());
    h.state_mut().ui.tab = "Tools".into();
    h.run_steps(3);
    assert_eq!(h.state().ui.tab, "Tools", "the presentation's own tab joins the tab row (an unknown one falls back to Home)");
    h.get_by_label("Go").click();
    h.run_steps(3);
    // Macros start disabled: the click says so instead of running.
    let said = h.state().status.as_ref().map(|(m, _)| m.clone()).unwrap_or_default();
    assert!(said.contains("disabled"), "{said}");
    // Developer › Ribbon Buttons adds one more, saved in the presentation.
    super::open_buttons(h.state_mut(), Some("Module1.Report"));
    h.run_steps(3);
    h.get_by_label("Add").click();
    h.run_steps(3);
    let ribbon = h.state().session.active().unwrap().doc.ribbon.clone().unwrap();
    let added = ribbon.tabs.iter().find(|t| t.label == "Macros").and_then(|t| t.groups.first()).and_then(|g| g.buttons.first()).cloned();
    assert_eq!(added.map(|b| (b.label, b.on_action, b.image)), Some(("Report".into(), "Module1.Report".into(), Some("MacroPlay".into()))));
    h.state_mut().addins.buttons = None;
    h.state_mut().ui.tab = "Macros".into();
    h.run_steps(3);
    assert!(h.query_by_label("Report").is_some());
}

#[test]
fn record_macro_toggles_on_the_developer_tab() {
    let _one_at_a_time = ADDINS.lock().unwrap_or_else(|e| e.into_inner());
    let m = deckcraft_engine::plugin::Manifest::parse(
        &json!({"id": "t.ui-recorder", "name": "Recorder Test", "macros": {"language": "VBA"}, "run": {"command": "./not-started"}}).to_string(),
    )
    .unwrap();
    deckcraft_engine::plugin::install(m, std::env::temp_dir()).unwrap();
    let _unload = Unload("t.ui-recorder");
    let mut h = harness(Session::with_new(), Services::default());
    h.state_mut().ui.tab = "Developer".into();
    h.run_steps(3);
    assert!(h.query_by_label("Ribbon Buttons").is_some());
    h.get_by_label("Record Macro").click();
    h.run_steps(3);
    assert!(h.state().session.recording.is_some());
    // Stopping asks the add-in for the code; this one can't start, which the status says.
    h.get_by_label("Stop Recording").click();
    h.run_steps(3);
    assert!(h.state().session.recording.is_none());
    assert!(h.query_by_label("Record Macro").is_some());
    let said = h.state().status.as_ref().map(|(m, _)| m.clone()).unwrap_or_default();
    assert!(said.contains("not-started") || said.contains("start"), "{said}");
}

/// A web view host that records what it was asked and lets the test post a page's messages.
#[derive(Clone, Default)]
struct FakeWeb(Arc<Mutex<Web>>);

#[derive(Default)]
struct Web {
    shown: Vec<(String, String)>,
    inbox: Vec<(String, String)>,
    replies: Vec<(String, String)>,
    hidden: Vec<String>,
}

impl OfficeHost for FakeWeb {
    fn show(&mut self, key: &str, url: &str, _rect: egui::Rect, _ppp: f32) {
        let mut w = self.0.lock().unwrap();
        if !w.shown.iter().any(|(k, _)| k == key) {
            w.shown.push((key.into(), url.into()));
        }
    }
    fn hide(&mut self, key: &str) {
        self.0.lock().unwrap().hidden.push(key.into());
    }
    fn close(&mut self, _key: &str) {}
    fn run_function(&mut self, _key: &str, _url: &str, _function: &str) {}
    fn take_messages(&mut self) -> Vec<(String, String)> {
        std::mem::take(&mut self.0.lock().unwrap().inbox)
    }
    fn reply(&mut self, key: &str, json: &str) {
        self.0.lock().unwrap().replies.push((key.into(), json.into()));
    }
}

#[test]
fn an_office_add_ins_task_pane_reads_and_writes_the_presentation() {
    let _one_at_a_time = ADDINS.lock().unwrap_or_else(|e| e.into_inner());
    let manifest = r#"<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>0d5c7a1e-0000-4bcd-9eaf-0000000000aa</Id><Version>1.0.0.0</Version><ProviderName>Tests</ProviderName><DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Pane Test"/><Description DefaultValue="Tests"/>
  <Hosts><Host Name="Presentation"/></Hosts>
  <DefaultSettings><SourceLocation DefaultValue="https://tests.example/pane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>"#;
    let m = deckcraft_engine::plugin::office::parse_manifest(manifest, deckcraft_engine::plugin::HOST_ID).unwrap();
    let id = m.id.clone();
    deckcraft_engine::plugin::install(m, std::env::temp_dir()).unwrap();
    struct UnloadId(String);
    impl Drop for UnloadId {
        fn drop(&mut self) {
            deckcraft_engine::plugin::unload(&self.0);
        }
    }
    let _unload = UnloadId(id.clone());
    let web = FakeWeb::default();
    let services = Services { office: Some(Box::new(web.clone())), ..Default::default() };
    let mut h = harness(Session::with_new(), services);
    let _ = h.state_mut().run("plugins.run", json!({"plugin": id, "command": "taskpane"}));
    h.run_steps(3);
    let key = web.0.lock().unwrap().shown.first().cloned().map(|(k, url)| {
        assert_eq!(url, "https://tests.example/pane.html");
        k
    });
    let key = key.expect("the pane's page was shown");
    let slides = h.state().session.active().unwrap().doc.slides.len();
    // One context.sync(): add a slide, then count.
    let msg = json!({"id": 7, "actions": [
        {"op": "call", "path": [{"p": "presentation"}, {"p": "slides"}], "method": "add", "args": [{}]},
        {"op": "load", "path": [{"p": "presentation"}, {"p": "slides"}], "props": ["count"]}
    ]});
    web.0.lock().unwrap().inbox.push((key.clone(), msg.to_string()));
    h.run_steps(2);
    let (to, reply) = web.0.lock().unwrap().replies.last().cloned().unwrap();
    assert_eq!(to, key);
    let reply: Json = serde_json::from_str(&reply).unwrap();
    assert_eq!(reply["id"], 7);
    assert_eq!(reply["results"][1]["count"], json!(slides + 1), "{reply}");
    assert_eq!(h.state().session.active().unwrap().doc.slides.len(), slides + 1);
    // The File tab hides the page (a native web view would cover it).
    h.state_mut().open_backstage(crate::backstage::Page::Home);
    h.run_steps(2);
    assert!(web.0.lock().unwrap().hidden.contains(&key));
}
