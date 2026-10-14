//! Developer tab: macros. DeckCraft doesn't run macros itself; an add-in that declares
//! `"macros": {"language": "VBA"}` does (protocol 1.1, `macros/*`), and reaches the presentation
//! back through the host requests answered here ([`host_call`]: `host/sync` on the same objects as
//! Office add-ins, the presentation's `vbaProject.bin`, messages). Running a macro is one undoable
//! step. PowerPoint has no macro recorder, and neither has DeckCraft.
//!
//! Macro security (`prefs.macro_security`): `disable` runs nothing; `prompt` (the default) runs a
//! presentation's macros once the user enabled them for it (`developer.enableMacros`, the message
//! bar's Enable Content); `enable` runs all.

use deckcraft_plugin::office::{Action, Permission};
use deckcraft_plugin::{HostError, call_macros, macro_plugins};
use serde_json::{Value as Json, json};

use super::{CommandSpec, always, bad, cmd, str_param};
use crate::{EngineError, Result, Session, UiRequest};

/// Largest `vbaProject.bin` a plugin may write.
const MAX_PROJECT: usize = 64 << 20;

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(query "developer.macros", "Macros", ["Developer", "Code"], None, "{} → {macros: [{name, module, procedure, plugin, language}], enabled, security}. The presentation's macros", always, macros),
        cmd!(
            "developer.runMacro",
            "Run Macro",
            ["Developer", "Code"],
            None,
            "{name, plugin?} → {ran}. Runs a macro (Module1.Hello) as one undoable step; needs macros enabled for the presentation",
            always,
            run_macro
        ),
        cmd!(query "developer.modules", "Visual Basic", ["Developer", "Code"], None, "{plugin?} → {modules: [{name, kind, source}]}. The presentation's code modules", always, modules),
        cmd!(
            "developer.saveModule",
            "Save Module",
            ["Developer", "Code"],
            None,
            "{name, source, plugin?} → {}. Adds or replaces a code module (saved with the presentation as .pptm)",
            always,
            save_module
        ),
        cmd!("developer.deleteModule", "Remove Module", ["Developer", "Code"], None, "{name, plugin?} → {}", always, delete_module),
        cmd!(noundo "developer.enableMacros", "Enable Content", ["Developer", "Code"], None, "{on?: bool} → {enabled}. Lets this presentation's macros run (until it closes)", always, enable_macros),
        cmd!(
            noundo "developer.macroSecurity",
            "Macro Security",
            ["Developer", "Code"],
            None,
            "{level?: disable|prompt|enable} → {level}. Which presentations' macros may run",
            always,
            macro_security
        ),
    ]
}

/// The plugin that runs macros: `plugin` when given, else the first one loaded.
fn macro_plugin(id: &str, p: &Json) -> Result<(String, String)> {
    let all = macro_plugins();
    match str_param(p, "plugin") {
        Some(want) => all.into_iter().find(|(pid, _)| pid == want).ok_or_else(|| bad(id, format!("no macro add-in `{want}` is loaded"))),
        None => all.into_iter().next().ok_or_else(|| bad(id, "no add-in runs macros: install one (VBA) in File › Options › Add-ins")),
    }
}

/// Whether the active presentation's macros may run.
pub fn macros_allowed(s: &Session) -> bool {
    match s.prefs.macro_security.as_str() {
        "enable" => true,
        "disable" => false,
        _ => s.active().is_some_and(|d| d.macros_enabled),
    }
}

/// Answers a plugin's request to the host (protocol 1.1) for the active presentation.
pub fn host_call(s: &mut Session, plugin: &str, method: &str, params: &Json) -> std::result::Result<Json, HostError> {
    let fail = |message: String| HostError { code: -32000, message, data: Json::Null };
    match method {
        "host/sync" => {
            let actions: Vec<Action> = serde_json::from_value(params.get("actions").cloned().unwrap_or(Json::Array(Vec::new())))
                .map_err(|e| fail(format!("`actions`: {e}")))?;
            if actions.len() > deckcraft_plugin::office::MAX_ACTIONS {
                return Err(fail("too many actions".into()));
            }
            // A plugin is a program the user installed: the presentation is open to it.
            match super::office::run_actions(s, plugin, Permission::ReadWriteDocument, &actions) {
                Ok(results) => Ok(json!({"results": results})),
                Err(error) => {
                    let message = error.get("message").and_then(Json::as_str).unwrap_or("failed").to_string();
                    Err(HostError { code: -32000, message, data: error })
                }
            }
        }
        "host/vbaProject/get" => {
            let bytes = s.active().and_then(|d| d.doc.vba_project.as_ref().map(|b| super::base64_encode(b)));
            Ok(json!({"base64": bytes}))
        }
        "host/vbaProject/set" => {
            let b64 = params.get("base64").and_then(Json::as_str).ok_or_else(|| fail("`base64` is required".into()))?;
            if b64.len() > MAX_PROJECT / 3 * 4 + 8 {
                return Err(fail("the project is larger than 64 MiB".into()));
            }
            let bytes = super::base64_decode(b64).ok_or_else(|| fail("invalid base64".into()))?;
            s.edit(|doc, _| {
                doc.vba_project = Some(bytes.into());
                // A signature over the old project doesn't hold for the new one.
                doc.vba_parts.retain(|p| !p.rel_type.to_ascii_lowercase().contains("signature"));
                Ok(())
            })
            .map_err(|e| fail(e.to_string()))?;
            Ok(json!({}))
        }
        "host/message" => {
            let text: String = params.get("text").and_then(Json::as_str).unwrap_or("").chars().take(4000).collect();
            s.ui_requests.push(UiRequest::Message { text });
            // vbOK: the message doesn't wait for the user yet.
            Ok(json!({"result": 1}))
        }
        "host/log" => {
            let text: String = params.get("text").and_then(Json::as_str).unwrap_or("").chars().take(4000).collect();
            log::info!("add-in {plugin}: {text}");
            Ok(json!({}))
        }
        _ => Err(HostError::unknown()),
    }
}

/// Sends `method` to the macro plugin, answering its host requests; what its host requests change
/// is part of the running command's one undo step, and a failure takes it back.
fn call(s: &mut Session, plugin: &str, method: &str, params: Json) -> Result<Json> {
    let before = super::plugins::snapshot(s);
    let r = call_macros(plugin, method, params, &mut |m, p| host_call(s, plugin, m, p));
    if r.is_err() {
        super::plugins::restore(s, before);
    }
    // The add-in's own words (a macro's `Run-time error '9': Subscript out of range`).
    r.map_err(EngineError::Other)
}

fn macros(s: &mut Session, p: &Json) -> Result<Json> {
    let mut out = Vec::new();
    let plugins: Vec<(String, String)> = match str_param(p, "plugin") {
        Some(_) => vec![macro_plugin("developer.macros", p)?],
        None => macro_plugins(),
    };
    for (plugin, language) in plugins {
        let r = call(s, &plugin, "macros/list", json!({}))?;
        for m in r.get("macros").and_then(Json::as_array).into_iter().flatten().take(10_000) {
            let mut m = m.clone();
            m["plugin"] = json!(plugin);
            m["language"] = json!(language);
            out.push(m);
        }
    }
    Ok(json!({"macros": out, "enabled": macros_allowed(s), "security": s.prefs.macro_security}))
}

fn run_macro(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "developer.runMacro";
    let name = str_param(p, "name").filter(|n| !n.trim().is_empty()).ok_or_else(|| bad(ID, "missing `name`"))?.to_string();
    if !macros_allowed(s) {
        return Err(EngineError::Disabled(
            ID.into(),
            "macros are disabled for this presentation: Enable Content, or change Developer › Macro Security".into(),
        ));
    }
    let (plugin, _) = macro_plugin(ID, p)?;
    let mut params = json!({"name": name});
    // A ribbon button's macro gets the button (`control.Id`, `control.Tag`).
    if let Some(c) = p.get("control").filter(|c| c.is_object()) {
        params["control"] = json!({"id": c.get("id").cloned().unwrap_or(json!("")), "tag": c.get("tag").cloned().unwrap_or(json!(""))});
    }
    call(s, &plugin, "macros/run", params)?;
    Ok(json!({"ran": name}))
}

fn modules(s: &mut Session, p: &Json) -> Result<Json> {
    let (plugin, _) = macro_plugin("developer.modules", p)?;
    let mut r = call(s, &plugin, "macros/modules", json!({}))?;
    r["plugin"] = json!(plugin);
    Ok(r)
}

fn save_module(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "developer.saveModule";
    let name = str_param(p, "name").filter(|n| !n.trim().is_empty()).ok_or_else(|| bad(ID, "missing `name`"))?;
    let source = str_param(p, "source").ok_or_else(|| bad(ID, "missing `source`"))?;
    if source.len() > 16 << 20 {
        return Err(bad(ID, "a module is limited to 16 MiB"));
    }
    let (plugin, _) = macro_plugin(ID, p)?;
    call(s, &plugin, "macros/saveModule", json!({"name": name, "source": source}))
}

fn delete_module(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "developer.deleteModule";
    let name = str_param(p, "name").filter(|n| !n.trim().is_empty()).ok_or_else(|| bad(ID, "missing `name`"))?;
    let (plugin, _) = macro_plugin(ID, p)?;
    call(s, &plugin, "macros/deleteModule", json!({"name": name}))
}

fn enable_macros(s: &mut Session, p: &Json) -> Result<Json> {
    let on = p.get("on").and_then(Json::as_bool).unwrap_or(true);
    let d = s.active_mut().ok_or_else(|| bad("developer.enableMacros", "no presentation is open"))?;
    d.macros_enabled = on;
    Ok(json!({"enabled": on}))
}

fn macro_security(s: &mut Session, p: &Json) -> Result<Json> {
    if let Some(level) = str_param(p, "level") {
        if !matches!(level, "disable" | "prompt" | "enable") {
            return Err(bad("developer.macroSecurity", "`level` is disable, prompt or enable"));
        }
        s.prefs.macro_security = level.to_string();
    }
    Ok(json!({"level": s.prefs.macro_security}))
}
