use serde_json::json;

use crate::{Manifest, parse_reply};

#[test]
fn manifests_are_validated() {
    // A shared manifest's spreadsheet functions are checked, and ignored here.
    let with = |functions: serde_json::Value| json!({"id": "t.plugin", "name": "T", "run": {"command": "true"}, "functions": functions}).to_string();
    let m = Manifest::parse(&with(json!([{"name": "t.one", "maxArgs": 2}]))).unwrap();
    assert_eq!(m.functions[0].name, "T.ONE");
    assert!(Manifest::parse(&with(json!([{"name": "1BAD"}]))).is_err());
    assert!(Manifest::parse(r#"{"id": "bad id", "name": "x", "run": {"command": "x"}}"#).is_err());
    assert!(Manifest::parse(r#"{"id": "x", "name": "x", "run": {"command": ""}}"#).is_err());
    assert!(Manifest::parse("not json").is_err());
}

#[test]
fn one_manifest_serves_several_apps() {
    // GridCraft's row menu means nothing here; `uiFor.deckcraft` replaces `ui`.
    let all = json!({
        "id": "t.all", "name": "All", "run": {"command": "true"}, "apps": ["gridcraft", "wordcraft", "DeckCraft"],
        "functions": [{"name": "T.TWICE", "minArgs": 1, "maxArgs": 1}],
        "commands": [{"id": "tidy", "label": "Tidy"}],
        "ui": {"menus": [{"menu": "row", "command": "tidy"}]},
        "uiFor": {"deckcraft": {"menus": [{"menu": "slide", "command": "tidy"}]}, "wordcraft": {"menus": [{"menu": "picture", "command": "tidy"}]}}
    });
    let m = Manifest::parse(&all.to_string()).unwrap();
    assert_eq!(m.ui.menus.len(), 1);
    assert_eq!(m.ui.menus[0].menu, crate::MenuKind::Slide);
    assert!(m.for_this_app());
    let only_excel = json!({"id": "t.xl", "name": "Excel", "run": {"command": "true"}, "apps": ["gridcraft"]});
    assert!(Manifest::parse(&only_excel.to_string()).unwrap_err().contains("not DeckCraft"));
    // Without `uiFor.deckcraft`, `ui` is this app's: another app's menu names are an error.
    let plain = json!({"id": "t.p", "name": "P", "run": {"command": "true"}, "commands": [{"id": "go", "label": "Go"}], "ui": {"menus": [{"menu": "row", "command": "go"}]}});
    assert!(Manifest::parse(&plain.to_string()).is_err());
}

#[test]
fn ui_declarations_must_name_the_plugins_commands() {
    let with_ui = |ui: serde_json::Value| {
        json!({"id": "t.ui", "name": "T", "run": {"command": "true"},
            "commands": [{"id": "go", "label": "Go"}], "ui": ui})
        .to_string()
    };
    let m = Manifest::parse(&with_ui(json!({
        "ribbon": [{"tab": "Design", "group": "Mine", "buttons": [{"command": "go", "size": "small", "params": {"n": 1}}]}],
        "menus": [{"command": "go"}, {"menu": "text", "command": "go", "label": "Go here"}]
    })))
    .unwrap();
    assert_eq!(m.ui.ribbon[0].buttons[0].size, crate::ButtonSize::Small);
    assert_eq!(m.ui.menus[0].menu, crate::MenuKind::Shape);
    assert_eq!(m.ui.menus[1].menu, crate::MenuKind::Text);
    assert_eq!(m.command("go").map(|c| c.label.as_str()), Some("Go"));
    for bad in [
        json!({"ribbon": [{"tab": "Design", "group": "G", "buttons": [{"command": "nope"}]}]}),
        json!({"ribbon": [{"tab": "Design", "group": "G", "buttons": []}]}),
        json!({"ribbon": [{"tab": " ", "group": "G", "buttons": [{"command": "go"}]}]}),
        json!({"ribbon": [{"tab": "Design", "group": "G", "buttons": [{"command": "go", "size": "huge"}]}]}),
        json!({"menus": [{"command": "nope"}]}),
        json!({"menus": [{"menu": "statusBar", "command": "go"}]}),
        json!({"menus": [{"command": "go", "label": "x".repeat(65)}]}),
    ] {
        assert!(Manifest::parse(&with_ui(bad.clone())).is_err(), "{bad}");
    }
}

#[test]
fn homepages_must_be_web_links() {
    let with = |url: &str| json!({"id": "t.home", "name": "T", "run": {"command": "true"}, "homepage": url}).to_string();
    assert_eq!(Manifest::parse(&with("https://example.org/doc")).unwrap().homepage, "https://example.org/doc");
    assert!(Manifest::parse(&with("")).is_ok());
    for bad in ["javascript:alert(1)", "file:///etc/passwd", "example.org", "https://a b"] {
        assert!(Manifest::parse(&with(bad)).is_err(), "{bad}");
    }
}

#[test]
fn turned_off_plugins_lose_their_commands() {
    let m = Manifest::parse(
        &json!({"id": "t.toggle", "name": "T", "run": {"command": "./never-started"}, "commands": [{"id": "go", "label": "Go"}]}).to_string(),
    )
    .unwrap();
    crate::install(m, std::env::temp_dir()).unwrap();
    assert_eq!(crate::set_enabled("t.toggle", false), Ok(true));
    assert_eq!(crate::set_enabled("t.toggle", false), Ok(false));
    assert_eq!(crate::get("t.toggle").map(|p| p.status()), Some("off".into()));
    assert!(crate::run_command("t.toggle", "go", json!({}), json!({})).unwrap_err().contains("turned off"));
    assert_eq!(crate::set_enabled("t.toggle", true), Ok(true));
    assert!(crate::unload("t.toggle"));
    assert!(crate::set_enabled("t.toggle", true).is_err());
}

#[test]
fn command_replies_are_checked() {
    let r = parse_reply(json!({"commands": [{"command": "slide.new", "params": {"title": "Hi"}}], "message": "hi"})).unwrap();
    assert_eq!(r.commands.len(), 1);
    assert_eq!(r.message.as_deref(), Some("hi"));
    assert!(parse_reply(json!({"commands": {}})).is_err());
    assert!(parse_reply(json!({"commands": [{"params": {}}]})).is_err());
    assert!(parse_reply(json!(null)).unwrap().commands.is_empty());
}

mod install {
    use std::io::Write;

    use crate::install::{Repo, Version, parse_github_url, pick_asset, unpack};

    fn repo(owner: &str, name: &str, version: Version) -> Repo {
        Repo { owner: owner.into(), name: name.into(), version }
    }

    #[test]
    fn github_addresses_are_read() {
        for url in ["https://github.com/acme/fin-tools", "github.com/acme/fin-tools/", "http://www.github.com/acme/fin-tools.git", "acme/fin-tools"] {
            assert_eq!(parse_github_url(url), Ok(repo("acme", "fin-tools", Version::Latest)), "{url}");
        }
        assert_eq!(parse_github_url("https://github.com/acme/fin/releases/tag/v1.2.0"), Ok(repo("acme", "fin", Version::Release("v1.2.0".into()))));
        assert_eq!(parse_github_url("https://github.com/acme/fin/tree/feat/x"), Ok(repo("acme", "fin", Version::Tree("feat/x".into()))));
        assert_eq!(parse_github_url("https://github.com/acme/fin/releases/latest"), Ok(repo("acme", "fin", Version::Latest)));
        for bad in [
            "",
            "https://gitlab.com/acme/fin",
            "https://github.com/acme",
            "https://github.com/../fin",
            "https://github.com/acme/fin/blob/main/x",
            "https://github.com/acme/fin/tree/../../x",
            "ftp://github.com/acme/fin",
            "a/b/c",
        ] {
            assert!(parse_github_url(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn the_release_zip_for_this_platform_is_picked() {
        let names: Vec<String> = [
            "fin-1.0-linux-x86_64.zip",
            "fin-1.0-macos-x86_64.zip",
            "fin-1.0-macos-aarch64.zip",
            "fin-1.0-windows-x64.zip",
            "fin-1.0-macos-aarch64.tar.gz",
            "checksums.txt",
        ]
        .map(String::from)
        .to_vec();
        assert_eq!(pick_asset(&names, "macos", "aarch64"), Some(2));
        assert_eq!(pick_asset(&names, "macos", "x86_64"), Some(1));
        assert_eq!(pick_asset(&names, "windows", "x86_64"), Some(3));
        assert_eq!(pick_asset(&names, "linux", "aarch64"), None, "no build for it: install the source");
        let universal: Vec<String> = ["fin-darwin-universal.zip", "fin-linux-amd64.zip"].map(String::from).to_vec();
        assert_eq!(pick_asset(&universal, "macos", "aarch64"), Some(0));
        let any: Vec<String> = ["fin-plugin.zip", "fin-linux-arm64.zip"].map(String::from).to_vec();
        assert_eq!(pick_asset(&any, "macos", "aarch64"), Some(0), "a zip naming no platform suits every one");
        assert_eq!(pick_asset(&any, "linux", "aarch64"), Some(1), "this platform's build beats the generic zip");
        assert_eq!(pick_asset(&["darwin-tools-windows.zip".to_string()], "macos", "x86_64"), Some(0));
        assert_eq!(pick_asset(&[], "macos", "aarch64"), None);
    }

    fn zip(files: &[(&str, &[u8], u32)]) -> Vec<u8> {
        let mut out = std::io::Cursor::new(Vec::new());
        let mut w = zip::ZipWriter::new(&mut out);
        for (name, data, mode) in files {
            let opts = zip::write::SimpleFileOptions::default().unix_permissions(*mode);
            w.start_file(*name, opts).unwrap();
            w.write_all(data).unwrap();
        }
        w.finish().unwrap();
        out.into_inner()
    }

    fn manifest(id: &str) -> Vec<u8> {
        serde_json::json!({"id": id, "name": "Fin", "run": {"command": "./run.sh"}}).to_string().into_bytes()
    }

    #[test]
    fn a_source_zip_unpacks_into_a_folder_named_after_the_plugin() {
        let root = std::env::temp_dir().join(format!("deckcraft-unpack-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        let m = manifest("acme.fin");
        let bytes = zip(&[
            ("acme-fin-1a2b3c/README.md", b"readme", 0o644),
            ("acme-fin-1a2b3c/plugin.json", &m, 0o644),
            ("acme-fin-1a2b3c/run.sh", b"#!/bin/sh\n", 0o755),
            ("acme-fin-1a2b3c/lib/util.py", b"x = 1\n", 0o644),
            ("acme-fin-1a2b3c/examples/other/plugin.json", b"{}", 0o644),
            ("elsewhere/secret.txt", b"not part of the plugin", 0o644),
        ]);
        let dir = unpack(&bytes, &root).unwrap();
        assert_eq!(dir, root.join("acme.fin"));
        assert!(dir.join("plugin.json").is_file() && dir.join("lib/util.py").is_file() && dir.join("examples/other/plugin.json").is_file());
        assert!(!dir.join("elsewhere").exists() && !root.join("elsewhere").exists());
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            assert_eq!(std::fs::metadata(dir.join("run.sh")).unwrap().permissions().mode() & 0o777, 0o755);
            assert_eq!(std::fs::metadata(dir.join("README.md")).unwrap().permissions().mode() & 0o777, 0o644);
        }
        crate::load(&dir).unwrap();
        assert!(crate::unload("acme.fin"));

        // Installing again replaces the earlier copy, leftovers included.
        std::fs::write(dir.join("stale.txt"), "old").unwrap();
        let again = unpack(&zip(&[("plugin.json", &m, 0o644), ("run.sh", b"new", 0o755)]), &root).unwrap();
        assert_eq!(again, dir);
        assert!(!dir.join("stale.txt").exists() && !dir.join("lib").exists());
        assert_eq!(std::fs::read_dir(&root).unwrap().count(), 1, "no staging folders left behind");
        std::fs::remove_dir_all(&root).unwrap();
    }

    #[test]
    fn unsafe_or_foreign_zips_are_refused() {
        let root = std::env::temp_dir().join(format!("deckcraft-unpack-bad-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        assert!(unpack(b"not a zip", &root).unwrap_err().contains("not a zip"));
        assert!(unpack(&zip(&[("README.md", b"hi", 0o644)]), &root).unwrap_err().contains("plugin.json"));
        assert!(unpack(&zip(&[("plugin.json", &manifest(".."), 0o644)]), &root).is_err(), "an id can't climb out of the folder");
        assert!(unpack(&zip(&[("plugin.json", b"{", 0o644)]), &root).is_err());
        // Paths escaping the plugin folder are skipped, not written.
        let m = manifest("acme.safe");
        let dir = unpack(&zip(&[("plugin.json", &m, 0o644), ("../escaped.txt", b"x", 0o644), ("/abs.txt", b"x", 0o644)]), &root).unwrap();
        assert!(!root.join("escaped.txt").exists() && !std::env::temp_dir().join("escaped.txt").exists());
        assert!(dir.join("plugin.json").is_file());
        std::fs::remove_dir_all(&root).unwrap();
    }
}

mod office {
    use crate::office::{Action, OfficeCommand, Permission, Step, parse_manifest};
    use crate::{MenuKind, run_command};

    /// An add-in manifest written for these tests from the public schema: a default task pane, a
    /// button on the Home tab, a menu with a function command on a custom tab, and a text menu item.
    pub(super) const MANIFEST: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>0d5c7a1e-1111-4bcd-9eaf-0123456789ab</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Acme &amp; Sons</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Acme Prices"/>
  <Description DefaultValue="Live prices on your slides."/>
  <IconUrl DefaultValue="https://acme.example/icon-32.png"/>
  <SupportUrl DefaultValue="https://acme.example/help"/>
  <AppDomains><AppDomain>https://login.acme.example</AppDomain></AppDomains>
  <Hosts><Host Name="Presentation"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="PowerPointApi" MinVersion="1.3"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="https://acme.example/taskpane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Acme.Group"><Label resid="Group.Label"/>
                <Control xsi:type="Button" id="Acme.Taskpane">
                  <Label resid="Button.Label"/>
                  <Supertip><Title resid="Button.Label"/><Description resid="Button.Tip"/></Supertip>
                  <Action xsi:type="ShowTaskpane"><TaskpaneId>Main</TaskpaneId><SourceLocation resid="Taskpane.Url"/></Action>
                </Control>
              </Group>
            </OfficeTab>
            <CustomTab id="Acme.Tab"><Label resid="Tab.Label"/>
              <Group id="Acme.Tools"><Label resid="Tools.Label"/>
                <Control xsi:type="Menu" id="Acme.Menu"><Label resid="Menu.Label"/>
                  <Items>
                    <Item id="Acme.Refresh"><Label resid="Refresh.Label"/><Action xsi:type="ExecuteFunction"><FunctionName>refreshAll</FunctionName></Action></Item>
                    <Item id="Acme.Bad"><Label resid="Refresh.Label"/><Action xsi:type="ShowTaskpane"><SourceLocation resid="Http.Url"/></Action></Item>
                  </Items>
                </Control>
              </Group>
            </CustomTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="ContextMenu">
            <OfficeMenu id="ContextMenuText">
              <Control xsi:type="Button" id="Acme.Text"><Label resid="Text.Label"/>
                <Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/></Action>
              </Control>
            </OfficeMenu>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="https://acme.example/taskpane.html"/>
        <bt:Url id="Commands.Url" DefaultValue="https://acme.example/commands.html"/>
        <bt:Url id="Http.Url" DefaultValue="http://acme.example/insecure.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="Acme"/>
        <bt:String id="Button.Label" DefaultValue="Show Prices"/>
        <bt:String id="Tab.Label" DefaultValue="Acme"/>
        <bt:String id="Tools.Label" DefaultValue="Tools"/>
        <bt:String id="Menu.Label" DefaultValue="More"/>
        <bt:String id="Refresh.Label" DefaultValue="Refresh All"/>
        <bt:String id="Text.Label" DefaultValue="Price This Text"/>
      </bt:ShortStrings>
      <bt:LongStrings><bt:String id="Button.Tip" DefaultValue="Opens the prices pane."/></bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>"#;

    #[test]
    fn office_manifests_become_plugins() {
        let m = parse_manifest(MANIFEST, "deckcraft").unwrap();
        assert_eq!(m.id, "office.0d5c7a1e-1111-4bcd-9eaf-0123456789ab");
        assert_eq!((m.name.as_str(), m.author.as_str(), m.homepage.as_str()), ("Acme Prices", "Acme & Sons", "https://acme.example/help"));
        assert_eq!(m.apps, ["deckcraft"]);
        let o = m.office.as_ref().unwrap();
        assert_eq!(o.permissions, Permission::ReadWriteDocument);
        assert_eq!(o.requirements, [("PowerPointApi".to_string(), "1.3".to_string())]);
        assert_eq!(o.app_domains, ["https://login.acme.example"]);
        // The default pane, the Home button, the menu's function (the http:// item is dropped) and the text item.
        let ids: Vec<&str> = m.commands.iter().map(|c| c.id.as_str()).collect();
        assert_eq!(ids, ["taskpane", "Acme.Taskpane", "Acme.Refresh", "Acme.Text"]);
        assert_eq!(
            o.commands["Acme.Taskpane"],
            OfficeCommand::ShowTaskpane { url: "https://acme.example/taskpane.html".into(), taskpane_id: Some("Main".into()) }
        );
        assert_eq!(
            o.commands["Acme.Refresh"],
            OfficeCommand::ExecuteFunction { function: "refreshAll".into(), url: "https://acme.example/commands.html".into() }
        );
        assert_eq!(m.ui.ribbon.len(), 2);
        assert_eq!((m.ui.ribbon[0].tab.as_str(), m.ui.ribbon[0].group.as_str()), ("Home", "Acme"));
        assert_eq!(m.ui.ribbon[0].buttons[0].tooltip.as_deref(), Some("Opens the prices pane."));
        assert_eq!((m.ui.ribbon[1].tab.as_str(), m.ui.ribbon[1].buttons[0].label.as_deref()), ("Acme", Some("Refresh All")));
        assert_eq!(m.ui.menus.len(), 1);
        assert_eq!(m.ui.menus[0].menu, MenuKind::Text);
    }

    #[test]
    fn office_commands_open_panes_instead_of_processes() {
        let m = parse_manifest(MANIFEST, "deckcraft").unwrap();
        let id = m.id.clone();
        crate::install(m, std::env::temp_dir()).unwrap();
        let r = run_command(&id, "Acme.Text", serde_json::json!({}), serde_json::json!({})).unwrap();
        assert!(matches!(r.office, Some(OfficeCommand::ShowTaskpane { .. })));
        assert!(r.commands.is_empty());
        assert_eq!(crate::get(&id).unwrap().status(), "ready", "no process was started");
        assert_eq!(crate::get(&id).unwrap().info()["kind"], "office");
        assert!(crate::unload(&id));
    }

    #[test]
    fn foreign_or_broken_office_manifests_are_refused() {
        assert!(parse_manifest(MANIFEST, "gridcraft").unwrap_err().contains("not this app"));
        for bad in [
            "<OfficeApp/>",
            "not xml <<<",
            "<Other/>",
            &MANIFEST.replace("0d5c7a1e-1111-4bcd-9eaf-0123456789ab", "../../etc"),
            &MANIFEST.replace("https://acme.example/", "http://acme.example/"),
        ] {
            assert!(parse_manifest(bad, "deckcraft").is_err(), "{bad:.40}");
        }
        let deep = format!("<OfficeApp>{}{}</OfficeApp>", "<a>".repeat(100), "</a>".repeat(100));
        assert!(parse_manifest(&deep, "deckcraft").unwrap_err().contains("deeply"));
    }

    #[test]
    fn sync_actions_read_as_json() {
        let a: Vec<Action> = serde_json::from_str(
            r#"[{"op":"load","path":[{"p":"presentation"},{"m":"getSelectedShapes"}],"props":["items","name"]},
                {"op":"set","path":[{"p":"presentation"},{"p":"slides"},{"m":"getItemAt","a":[0]},{"p":"shapes"},{"m":"getItem","a":["257"]}],"prop":"left","value":72},
                {"op":"call","path":[{"p":"presentation"},{"m":"getSelectedShapes"}],"method":"delete","args":[]}]"#,
        )
        .unwrap();
        assert!(matches!(&a[1], Action::Set { prop, .. } if prop == "left"));
        assert!(matches!(&a[0], Action::Load { path, .. } if path[1] == Step::Method { m: "getSelectedShapes".into(), a: vec![] }));
    }
}
