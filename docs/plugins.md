# Add-ins, Office add-ins and macros

> **Last reviewed:** 2026-10-14 · **Last updated:** 2026-10-15 · **Change:** minor (ribbon buttons that run macros, saved as the presentation's `customUI14.xml`; Record Macro, a DeckCraft extra; previously major: first version: out-of-process add-ins, Office add-ins' task panes, VBA macros through a macro add-in) · **Target:** Microsoft PowerPoint (Microsoft 365)

DeckCraft runs three kinds of extension, all as **add-ins** listed in the Add-ins window (Insert ›
Add-ins, Developer › Add-ins):

1. **Add-ins** (`plugin.json`): programs in any language that DeckCraft starts on first use and
   talks to over JSON-RPC. They add commands, ribbon groups and right-click items. The manifest
   and protocol are the craft apps' shared ones (GridCraft, WordCraft): one add-in can serve
   several apps.
2. **Office add-ins** (`manifest.xml`): web add-ins written for PowerPoint with the Office
   JavaScript API. Their pages show in a task pane, running **DeckCraft's own Office.js**
   (`crates/office-js`); Microsoft's `office.js` is never loaded, read or shipped.
3. **Macros**: an add-in that declares `"macros": {"language": "VBA"}` (craft-vba,
   [sirmmo/craft-vba](https://github.com/sirmmo/craft-vba)) runs the presentation's VBA. The
   Developer tab appears once one is loaded.

An add-in is a program that runs with the user's permissions; the Add-ins window says so before
installing one from GitHub. The web app can't run add-ins.

## Installing

- **From GitHub:** Add-ins › Install from GitHub (`plugins.install {"url": "owner/repo"}`): the
  release zip for this platform, else the repository's source, goes into `<settings>/plugins/` and
  is loaded. `<settings>` is `~/Library/Application Support/DeckCraft` on macOS, `%APPDATA%\DeckCraft`
  on Windows, `~/.config/deckcraft` elsewhere.
- **From a folder:** Add-ins › Load from a folder, or `plugins.load {"path": …}` (a folder holding
  `plugin.json` or an Office add-in's `manifest.xml`).
- At start-up DeckCraft loads every add-in in `DECKCRAFT_PLUGIN_PATH` and `<settings>/plugins/`.
- **Activate / Deactivate** (`plugins.setEnabled`) is remembered in the preferences; **Unload**
  lasts until the next start.

## The manifest (`plugin.json`)

Shared with the other craft apps; see GridCraft's `docs/plugins.md` for every field. What DeckCraft
reads:

| Field | DeckCraft |
|---|---|
| `id`, `name`, `version`, `description`, `author`, `homepage` | Shown in the Add-ins window |
| `apps` | Must name `deckcraft` when present (`["gridcraft", "deckcraft"]`) |
| `run` | `{command, args}`, started in the add-in's folder |
| `commands` | `[{id, label, description, params}]`, run by `plugins.run` |
| `ui` / `uiFor.deckcraft` | Ribbon groups (`tab`: Home, Insert, Draw, Design, Transitions, Animations, Slide Show, Record, Review, View, or a new tab) and right-click items (`menu`: `shape`, `slide`, `text`) |
| `macros` | `{language}`: the add-in runs macros (protocol 1.1) |
| `functions` | GridCraft's spreadsheet functions: accepted, not used |

## The protocol

JSON-RPC 2.0, one message per line on the add-in's standard input and output.

- `initialize {protocolVersion: 1, host: {id: "deckcraft", name, version}, plugin: {id, dir}}`.
  The environment also has `CRAFT_HOST=deckcraft` and `CRAFT_PLUGIN_DIR`.
- `commands/run {command, params, context}` → `{commands: [{command, params}], message?, result?}`.
  `context` is `{presentation, slide (0-based), slideId, slides, titles: [slide titles], selection: [shape ids], text}`.
  The answered commands are DeckCraft's own (`slide.new`, `insert.textBox`, `shape.fill`…; not
  `plugins.*` or `developer.*`) and run as **one undoable step**; if one fails, nothing changes.

**Protocol 1.1: asking the host.** While a request of DeckCraft's is waiting, the add-in may send
requests of its own; DeckCraft answers them, then waits again:

| Request | Answer |
|---|---|
| `host/sync {actions}` | `{results}`: the same actions as an Office add-in's `context.sync()` (below), on the open presentation; an error's `data` is `{code, action}` |
| `host/vbaProject/get {}` | `{base64}`: the presentation's `vbaProject.bin`, or `null` |
| `host/vbaProject/set {base64}` | `{}`: replaces it (at most 64 MiB; signatures over the old one are dropped) |
| `host/message {text}` | `{result: 1}`: shown in the status bar |
| `host/log {text}` | `{}`: written to the log |

What the add-in changes this way is part of the running command's one undo step.

**Macros** (`"macros"` in the manifest): `macros/list`, `macros/run {name, control?: {id, tag}}`,
`macros/modules`, `macros/saveModule {name, source}`, `macros/deleteModule {name}`,
`macros/record {steps, name?}` → `{source}`, as in GridCraft. Commands: `developer.macros`,
`developer.runMacro {name}`, `developer.modules`, `developer.saveModule`, `developer.deleteModule`,
`developer.recordMacro {stop?, module?, name?}`, `developer.recordToCode {steps, name?}`,
`developer.enableMacros`, `developer.macroSecurity {level}`.

**Record Macro** (Developer › Code; a **DeckCraft extra**: PowerPoint dropped its macro recorder
in PowerPoint 2007, DeckCraft has one). Record Macro notes where the command journal stands; Stop
Recording sends what you did since to the add-in as `macros/record {steps: [{command, params}],
name?}`, with DeckCraft's own command ids and their params exactly as journaled (slide indexes
0-based, shape and slide ids the engine's numbers, colours `#RRGGBB`), and puts the `source` it
answers at the end of Module1, which opens in the editor. The Developer tab's own commands and what
add-ins do (`developer.*`, `plugins.*`, `office.*`) aren't recorded; Delete while editing text is
journaled as the text delete it is (`text.delete {dir: forward}`), not as a shape delete. craft-vba
writes PowerPoint VBA from the steps (0.1.5 and later); steps it can't express become comments.
`developer.recordToCode {steps}` turns steps given directly into code without a recording.

**Ribbon buttons** (Developer › Ribbon Buttons, or Add to Ribbon… in the Macros window): buttons
that run the presentation's macros, on a built-in tab or one of the presentation's own. They are
saved with the presentation as the Office ribbon part (`customUI/customUI14.xml`, related from the
package root), so PowerPoint shows them too, and a presentation made in PowerPoint brings its
buttons along (`customUI14.xml` or the Office 2007 `customUI.xml`). A ribbon that holds more than
tabs, groups and buttons (pictures of its own, menus, `onLoad`) is kept unchanged, pictures
included: its buttons show and run but can't be edited. A button runs its macro as PowerPoint does,
with the button: `macros/run` gets `control: {id, tag}`, which a VBA add-in passes as the
`IRibbonControl` argument (`Sub Go(control As IRibbonControl)`); a macro without parameters runs as
it is. The buttons show whether or not a macro add-in is loaded (without one, a click says so).
Commands: `developer.ribbon` → `{tabs: [{tab, id, builtin, label, groups: [{id, label, buttons:
[{id, label, macro, image, large, screentip, tag}]}]}], editable}`,
`developer.addRibbonButton {macro, label?, tab?: "Macros", group?: "Macros", image?, large?,
screentip?, tag?}` → `{id}`, `developer.editRibbonButton {id, …}`,
`developer.removeRibbonButton {id}`, `developer.pressRibbonButton {id}` → `{ran}`. Built-in tabs
by name: Home, Insert, Design, Transitions, Animations, Slide Show, Review, View, Developer
(`TabHome`…); any other name is a tab of the presentation's own.

**Macro security** (Developer › Macro Security): `prompt`, the default, runs a presentation's macros
once the user clicked **Enable Content** in the bar a presentation with macros opens with;
`disable` runs none; `enable` runs all. Agents, the CLI and MCP are bound by the same setting.

**Saving:** macros live in `.pptm` (`.potm`, `.ppsm`). A presentation with macros saved under a
`.pptx`, `.potx` or `.ppsx` name is saved as `.pptm`, `.potm` or `.ppsm` instead, and DeckCraft says
so; the asked file isn't written.

## Office add-ins

An Office add-in's `manifest.xml` (with `<Host Name="Presentation"/>` and `https://` pages) becomes
the add-in `office.<its Id>`: its default task pane is the command `taskpane`, and each add-in
command of `<VersionOverrides>` (buttons on Office tabs — `TabHome`, `TabInsert`, `TabDesign`,
`TabTransitions`, `TabAnimations`, `TabSlideShow`, `TabReview`, `TabView`, `TabDeveloper` — or
custom tabs, menus, the text right-click menu `ContextMenuText`) is a command of its own. The
manifest's `<Permissions>` bound what the page may do (`Restricted`, `ReadDocument`,
`ReadWriteDocument`).

The task pane is the system web view (WKWebView on macOS, WebView2 on Windows) laid over a
right-hand panel (`apps/deckcraft/src/office_web.rs`). Before the page's own scripts run, it gets
the bootstrap (`crates/office-js/bootstrap.js`, the one hand-written script, copied from GridCraft's)
and DeckCraft's Office.js built to wasm (`cargo xtask office-js`; the macOS package carries it in
`Resources/office-js`, or set `DECKCRAFT_OFFICE_JS`). Each `context.sync()` becomes one
`office.sync` command: one undo step, all or nothing.

**Objects** (PowerPoint JavaScript API, PowerPointApi 1.1–1.4 subset):

| Object | Properties | Methods |
|---|---|---|
| `presentation` | `title` | `slides`, `slideMasters`, `getSelectedSlides()`, `getSelectedShapes()`, `setSelectedSlides(ids)` |
| `SlideCollection` | `count`, `items` | `getItem(id)`, `getItemAt(i)`, `add({layoutId?, index?})` |
| `Slide` | `id`, `index` | `shapes`, `layout`, `slideMaster`, `delete()`, `setSelectedShapes(ids)` |
| `SlideMasterCollection` / `SlideMaster` / `SlideLayoutCollection` / `SlideLayout` | `id`, `name`, `count`, `items` | `getItem`, `getItemAt`, `layouts` |
| `ShapeCollection` | `count`, `items` | `getItem(id)`, `getItemAt(i)`, `addTextBox(text, box)`, `addGeometricShape(type, box)`, `addLine(type, box)` |
| `Shape` | `id`, `name`, `type`, `left`, `top`, `width`, `height` | `textFrame`, `fill`, `lineFormat`, `delete()` |
| `TextFrame` | `hasText`, `wordWrap`, `autoSizeSetting` | `textRange` |
| `TextRange` | `text`, `start`, `length` | `font`, `getSubstring(start, length)` |
| `ShapeFont` | `bold`, `italic`, `underline`, `size`, `name`, `color` (`null` when mixed) | |
| `ShapeFill` | `type`, `foregroundColor`, `transparency` | `setSolidColor(color)`, `clear()` |
| `ShapeLineFormat` | `color`, `weight`, `visible`, `dashStyle` | |

The common API: `Office.onReady`, `Office.context.document.getSelectedDataAsync`/`setSelectedDataAsync`
(text), `settings` (kept per open presentation until it closes), `requirements.isSetSupported`.
A shape added earlier in the same sync is named by that action's result (`getItem({"result": i})`).

**Not yet:** events, dialogs (`displayDialogAsync`), tags, `insertSlidesFromBase64`, hyperlinks,
settings saved in the file, the unified (JSON) manifest, and a web view on Linux.

## VBA (craft-vba)

craft-vba picks PowerPoint's object model when `host.id` is `deckcraft`: `ActivePresentation`,
`Slides` (`Count`, `Item`, `Add(index, ppLayout…)`, `For Each`), `Slide` (`Shapes`, `SlideIndex`,
`SlideID`, `Delete`), `Shapes` (`AddTextbox`, `AddShape` with the `msoShape…` constants, `AddLine`),
`Shape` (`Name`, `Left`, `Top`, `Width`, `Height`, `TextFrame.TextRange.Text`, `Font.Bold/Italic/
Size/Name/Color.RGB`, `Fill.ForeColor.RGB`, `Line…`, `Delete`), `ActiveWindow.Selection`, `MsgBox`.
A macro is one undo step; a run-time error takes everything it did back
(`Run-time error '9': Subscript out of range`). See craft-vba's README for the full list.

## Revision history

| Date | Change | Summary |
|---|---|---|
| 2026-10-15 | minor | Ribbon buttons that run macros (`customUI14.xml`, `macros/run` gets `control`); Record Macro and `developer.recordToCode` (`macros/record`), a DeckCraft extra |
| 2026-10-14 | major | First version: add-ins (shared manifest and protocol 1 and 1.1), Office add-ins' task panes with DeckCraft's Office.js, VBA macros through craft-vba, the Developer tab |
