//! Add-ins (Home › Add-ins, Developer › Add-ins): load and unload out-of-process plugins, list
//! them, and run their commands (see `deckcraft_plugin`). A plugin command is one undoable step:
//! the engine commands it answers with, and whatever it changed through the host meanwhile
//! (protocol 1.1, [`super::developer::host_call`]), succeed or go back together.

use std::sync::Arc;

use serde_json::{Value as Json, json};

use super::{CommandSpec, always, bad, cmd, str_param};
use crate::{Result, Session, UiRequest};

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(query "plugins.list", "Add-ins", [], None, "{} → [{id, name, version, description, dir, status, commands, ui, kind}]", always, |_, _| Ok(list())),
        cmd!(query "plugins.load", "Load Add-in", [], None, "{path} → the add-in. `path`: a folder holding plugin.json (or an Office add-in's manifest.xml), or the manifest", always, load),
        cmd!(query "plugins.unload", "Unload Add-in", [], None, "{id}", always, unload),
        cmd!(
            query "plugins.install",
            "Install Add-in from GitHub",
            [],
            None,
            "{url, dir?} → the add-in. Downloads the add-in in a GitHub repository (the release zip for this platform, else its source) into `dir` (default: the app's add-ins folder) and loads it",
            always,
            install
        ),
        cmd!(
            query "plugins.setEnabled",
            "Turn Add-in On or Off",
            [],
            None,
            "{id, enabled} → the add-in. Off, an add-in stays listed but its commands, ribbon and menu items are gone; remembered in the preferences",
            always,
            set_enabled
        ),
        cmd!(query "plugins.openHomepage", "Add-in Documentation", [], None, "{id} → {url}. Opens the add-in's web page (its manifest's `homepage`) in the browser", always, open_homepage),
        cmd!(
            "plugins.run",
            "Run Add-in Command",
            [],
            None,
            "{plugin, command, params?} → {ran, message?, result?}. The add-in's commands run as one undoable step",
            always,
            run
        ),
    ]
}

/// `plugins.list`.
pub fn list() -> Json {
    Json::Array(deckcraft_plugin::list().iter().map(|p| p.info()).collect())
}

fn load(s: &mut Session, p: &Json) -> Result<Json> {
    let path = str_param(p, "path").filter(|x| !x.trim().is_empty()).ok_or_else(|| bad("plugins.load", "missing `path`"))?;
    load_from(s, "plugins.load", std::path::Path::new(path))
}

fn install(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "plugins.install";
    let url = str_param(p, "url").filter(|x| !x.trim().is_empty()).ok_or_else(|| bad(ID, "missing `url`"))?;
    let root = match str_param(p, "dir").filter(|x| !x.trim().is_empty()) {
        Some(d) => std::path::PathBuf::from(d),
        None => deckcraft_plugin::install::install_dir().ok_or_else(|| bad(ID, "no add-ins folder: pass `dir`"))?,
    };
    let dir = deckcraft_plugin::install::from_github(url, &root).map_err(|e| bad(ID, e))?;
    load_from(s, ID, &dir)
}

/// Loads the add-in in `path` (a folder or its manifest).
pub fn load_from(s: &mut Session, id: &str, path: &std::path::Path) -> Result<Json> {
    let plugin = deckcraft_plugin::load(path).map_err(|e| bad(id, e))?;
    // An add-in turned off stays off when it's loaded again.
    if s.prefs.disabled_plugins.contains(&plugin.manifest.id) {
        let _ = deckcraft_plugin::set_enabled(&plugin.manifest.id, false);
    }
    Ok(plugin.info())
}

fn unload(_s: &mut Session, p: &Json) -> Result<Json> {
    let id = str_param(p, "id").ok_or_else(|| bad("plugins.unload", "missing `id`"))?;
    if !deckcraft_plugin::unload(id) {
        return Err(bad("plugins.unload", format!("no add-in `{id}` is loaded")));
    }
    Ok(json!({"unloaded": id}))
}

fn set_enabled(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "plugins.setEnabled";
    let id = str_param(p, "id").ok_or_else(|| bad(ID, "missing `id`"))?;
    let on = p.get("enabled").and_then(Json::as_bool).ok_or_else(|| bad(ID, "missing `enabled` (true or false)"))?;
    let plugin = deckcraft_plugin::get(id).ok_or_else(|| bad(ID, format!("no add-in `{id}` is loaded")))?;
    deckcraft_plugin::set_enabled(id, on).map_err(|e| bad(ID, e))?;
    s.prefs.disabled_plugins.retain(|x| x != id);
    if !on {
        s.prefs.disabled_plugins.push(id.to_string());
    }
    Ok(plugin.info())
}

fn open_homepage(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "plugins.openHomepage";
    let id = str_param(p, "id").ok_or_else(|| bad(ID, "missing `id`"))?;
    let plugin = deckcraft_plugin::get(id).ok_or_else(|| bad(ID, format!("no add-in `{id}` is loaded")))?;
    let url = plugin.manifest.homepage.trim();
    if url.is_empty() {
        return Err(bad(ID, format!("add-in {id} has no web page")));
    }
    s.ui_requests.push(UiRequest::OpenUrl { url: url.to_string() });
    Ok(json!({"url": url}))
}

/// Turns off the loaded add-ins the preferences list as off (at startup, after loading them).
pub fn apply_prefs(prefs: &crate::Prefs) {
    for id in &prefs.disabled_plugins {
        if deckcraft_plugin::get(id).is_some()
            && let Err(e) = deckcraft_plugin::set_enabled(id, false)
        {
            log::warn!("add-in {id}: {e}");
        }
    }
}

/// What a plugin command sees of the presentation: its title, the current slide (0-based index
/// and id), the slides' titles, the selected shapes' ids and the selected or edited text.
fn context(s: &Session) -> Json {
    let Some(d) = s.active() else { return json!({}) };
    let titles: Vec<String> = d.doc.slides.iter().take(10_000).map(|sl| super::office::slide_title(sl)).collect();
    let text = d.selection.text.as_ref().and_then(|t| {
        let sh = d.shape(t.shape)?;
        let body = sh.text.as_ref()?;
        let ((p0, c0), (p1, c1)) = t.ordered();
        Some(super::office::text_between(body, (p0, c0), (p1, c1)))
    });
    json!({
        "presentation": d.title(),
        "slide": d.selection.slide,
        "slideId": d.current_slide().map(|sl| sl.id.0.to_string()),
        "slides": d.doc.slides.len(),
        "titles": titles,
        "selection": d.selection.shapes.iter().map(|id| id.0.to_string()).collect::<Vec<_>>(),
        "text": text,
    })
}

/// The active presentation and selection, to put back if a step fails.
pub(crate) fn snapshot(s: &Session) -> Option<(u64, Arc<deckcraft_model::Presentation>, crate::Selection)> {
    s.active().map(|d| (d.uid, d.doc.clone(), d.selection.clone()))
}

/// Puts back what [`snapshot`] took (a failed step changes nothing).
pub(crate) fn restore(s: &mut Session, snap: Option<(u64, Arc<deckcraft_model::Presentation>, crate::Selection)>) {
    let Some((uid, doc, sel)) = snap else { return };
    if let Some(d) = s.active_mut().filter(|d| d.uid == uid) {
        d.doc = doc;
        d.selection = sel;
        d.revision += 1;
    }
}

fn run(s: &mut Session, p: &Json) -> Result<Json> {
    const ID: &str = "plugins.run";
    let plugin = str_param(p, "plugin").ok_or_else(|| bad(ID, "missing `plugin`"))?;
    let command = str_param(p, "command").ok_or_else(|| bad(ID, "missing `command`"))?;
    let params = p.get("params").cloned().unwrap_or_else(|| json!({}));
    // Commands run inside this one are no undo steps or journal entries of their own: this
    // command's step covers the add-in's host requests and its answer's commands.
    let before = snapshot(s);
    let ctx = context(s);
    let reply = deckcraft_plugin::run_command_with(plugin, command, params, ctx, &mut |m, p| super::developer::host_call(s, plugin, m, p));
    let reply = match reply {
        Ok(r) => r,
        Err(e) => {
            restore(s, before);
            return Err(bad(ID, e));
        }
    };
    // An Office add-in's command opens its task pane or calls its function: nothing runs here.
    if let Some(office) = reply.office {
        let title = deckcraft_plugin::get(plugin).map(|p| p.manifest.name.clone()).unwrap_or_default();
        s.ui_requests.push(match &office {
            deckcraft_plugin::office::OfficeCommand::ShowTaskpane { url, taskpane_id } => {
                UiRequest::OfficePane { plugin: plugin.to_string(), title, url: url.clone(), pane: taskpane_id.clone() }
            }
            deckcraft_plugin::office::OfficeCommand::ExecuteFunction { function, url } => {
                UiRequest::OfficeFunction { plugin: plugin.to_string(), url: url.clone(), function: function.clone() }
            }
        });
        return Ok(json!({"ran": 0, "office": office}));
    }
    if let Some((id, _)) = reply.commands.iter().find(|(id, _)| id.starts_with("plugins.") || id.starts_with("developer.")) {
        restore(s, before);
        return Err(bad(ID, format!("an add-in command can't run `{id}`")));
    }
    let mut failure = None;
    for (id, params) in &reply.commands {
        if let Err(e) = s.execute(id, params) {
            failure = Some(format!("{id}: {e}"));
            break;
        }
    }
    if let Some(why) = failure {
        restore(s, before);
        return Err(bad(ID, format!("{plugin} › {command} failed, nothing was changed: {why}")));
    }
    if let Some(m) = &reply.message {
        s.ui_requests.push(UiRequest::Message { text: m.clone() });
    }
    let mut out = json!({"ran": reply.commands.len()});
    if let Some(m) = reply.message {
        out["message"] = json!(m);
    }
    if !reply.result.is_null() {
        out["result"] = reply.result;
    }
    Ok(out)
}
