# Live sessions (co-editing)

> **Last reviewed:** 2026-10-15 · **Last updated:** 2026-10-15 · **Change:** major (created: plugin protocol 2 in DeckCraft, schema 1, Share in the title bar) · **Target:** Microsoft PowerPoint (Microsoft 365)

Several people edit one presentation at the same time, each in their own DeckCraft, and see each
other's selections, carets and edits as they happen: DeckCraft's equivalent of PowerPoint's
co-authoring. It's **native**: DeckCraft merges the edits itself (a CRDT, Yjs through the
[`yrs`](https://crates.io/crates/yrs) crate), shows who's there and keeps a per-person undo.
Moving the bytes between people is a **transport add-in**'s job (MQTT, a local relay, a
y-websocket server, Nextcloud…), which never needs to understand a presentation. With a room key
the content is encrypted end to end, so the transport only ever sees ciphertext.

The protocol is shared with GridCraft and WordCraft:
[craft plugin protocol 2](https://github.com/sirmmo/craft-collab/blob/main/docs/protocol-2.md)
(push channel, rooms, `collab/*` messages, encryption, limits). The shared document is
**DeckCraft schema 1**, specified (normatively) in craft-collab's
[`docs/schemas/deckcraft-1.md`](https://github.com/sirmmo/craft-collab/blob/main/docs/schemas/deckcraft-1.md);
this page summarises it and says how the app behaves.

## Using it

**Share**, at the right end of the title bar. Not in a session, it opens the Share window: pick
the transport add-in, your name and colour, then **Start live session** or **Join a session…**.
A transport with its own `share` and `join` commands (craft-collab's craft-mqtt and loopback have
them) takes over from there: it asks for what it needs (a broker, an invitation link) in its own
dialog and answers with the room to join. For a transport without them, type the room and its
key (Encrypt end to end makes a new key when starting). With no transport add-in loaded, the
window says one is needed and opens the Add-ins window.

In a session, the button shows the session's state with a dot (green: **Live**; amber:
**Connecting…** or **Joining…**; grey: **Offline**, editing goes on here; red: another app's
room) and opens the session popover: the state, whether it's encrypted, the room id and
transport, the transport's message, everyone in it (circle, name, "(you)"; click a name to go to
their selection), **Copy invitation** when the transport's Share or Join gave a link
(`result.link`), and **Leave** (the presentation stays open with every edit, as an ordinary
presentation; it isn't saved automatically).

To the left of Share, the people in the session are overlapping circles with their initials on
their colour (black or white text, whichever reads), you first, at most four and then "+N"
(hovering it lists the rest; clicking it opens the popover). Hover a circle for the name ("(you)"
for yourself); click it to go to the slide that person is on and select what they selected.

On the slide, others' selections are outlined in their colours with a name tag, their caret is
drawn where they're typing (it follows the text as you both type), and shapes they change are
outlined in amber for two seconds. The status bar adds the session's state after the slide
counter ("Slide 2 of 9 · Joining…"). While the room's document hasn't arrived yet (**Joining…**)
and for viewers (`role: view`), the presentation is read-only.

Undo and Redo, while in a session, undo **your** changes only; others' edits never enter your
undo history. Saving works as usual: each person saves their own copy.

## Commands

| Command | Params | Does |
|---|---|---|
| `collab.share` | `{plugin, room, user?, role?, key?, title?, link?}` | Starts a session from the active presentation (no params: the Share window) |
| `collab.join` | `{plugin, room, mode: share\|join, user?: {id, name, color}, role?: edit\|view, key?, path?, title?, link?}` | Joins or starts a session (no params: the Share window). `share` seeds the room from the active presentation (or `path`); `join` opens a new presentation (or `path`) filled from the room |
| `collab.leave` | `{room?}` | Leaves (default: the active presentation's room) |
| `collab.status` | `{}` | `{rooms: [{room, plugin, state, peers, listed, role, encrypted, clientId, user, title, document, message, link, canUndo, canRedo}], transports, user}` |
| `collab.setUser` | `{name?, color?}` | Your name (the author name comments carry) and colour (`#RRGGBB`) |

A plugin command's result may carry `collab: {join: {…}}` (with protocol 1.2's `open`, `dialog`
and `message`): `plugins.run` returns it to the caller, and the interface runs `collab.join` with
it (the plugin's id as `plugin`) when a button, menu item, plugin dialog or the Share window ran
the command, keeping `result.link` as the room's invitation. Programmatic calls never open
dialogs or windows.

Rust API (for embedders, tests and a future web build): `Session::collab_poll(now_ms)` runs the
sessions between frames; `Session::collab_serve_in_memory(id)` / `collab_join_in_memory`,
`collab_take_outgoing` and `collab_poll_pushed` connect a session to a transport the caller
implements; `collab_take_notices` gives the status-bar messages. The core, without the engine, is
the `deckcraft-collab` crate (L2, no egui, builds for the web).

## Schema 1 in brief

`meta` `{app: "deckcraft", schema: 1, title}`; `presentation` (presentation-wide settings, one
JSON-text entry each, and `masters`, the master ids in order); `slides` (slide ids in order);
`slide:<id>` per slide: its properties as JSON-text entries (`layout`, `transition`,
`animations`, `background`…), `shapes` (its top-level shape ids, back to front), a map
`shape:<id>` per shape (groups' children included, listed in the group's `children`), `notes`
(Yjs text) and `notesBody`; a shape's properties are JSON-text entries (`xfrm`, `geom`, `kind`,
`fill`, `line`…, last writer wins each) and its text is Yjs text; `masters`, `layouts` and
`theme` hold whole JSON values (they change rarely and together); `media` holds pictures and
other files by id (immutable).

- **Ids** are 8 characters of `0-9a-z`, the base-36 form of the object's 32-bit id in DeckCraft's
  model, so references inside values (connector ends, animation targets, links to slides,
  pictures' media, a slide's layout) are the same numbers on every peer. In a session DeckCraft
  hands out new ids from a random block (a random start from 2^24 to 2^31 with 65,536 free ids
  after it, renewed after 32,768), so two people adding at once don't meet.
- **Text** is Yjs text: each paragraph's characters then `\n`; a line break is U+000B; formatting
  is Yjs attributes, one per run property (`bold: true`, `size: 24`, `font`…), with fields
  (`field`) and equations (`math`) marked the same way; the `\n` carries the paragraph's `level`,
  its properties (`p:<name>`), its default run properties (`defRPr`) and the end mark's run
  properties. Typing is character inserts, so two people typing in one text box both keep their
  text, and formatting different properties of the same words keeps both.
- **Awareness** is `{user, selection: {slide, shapes, text?: {shape | notes, anchor, head}}}`,
  `anchor` and `head` Yjs relative positions (base64), so a caret stays put while others type.

## Behaviour

- **Push channel.** Every running add-in's standard output is read all the time; what an add-in
  sends on its own (notifications, `collab/snapshot` and `host/log` requests) waits in a queue that
  the app's main loop takes each frame, never in the middle of a command. `host/…` requests
  during one of DeckCraft's requests are still answered by it (protocol 1.1). Lines may be 64 MiB;
  an add-in sending more than 1,000 messages a second for 5 s is stopped and marked failed.
  `initialize` gets `host.features` `push` and `collab` and `host.collab = {formats: ["yjs-v1"],
  app: "deckcraft", schema: 1, encryption: ["aes-256-gcm"]}`.
- **Sharing**: the room is seeded from the presentation in one transaction (origin `seed`, not
  undoable), then `collab/joined` and `collab/sync1`. **Joining** (`mode: join`): a new
  presentation opens, read-only and showing "Joining…", until a state for DeckCraft schema 1
  arrives. A room without masters (a peer that shared only slides) is shown with DeckCraft's
  default master.
- **Local edits.** Each committed command (the outermost: a plugin command's steps are one), each
  finished drag and each committed input-method composition is **one** transaction and one
  `collab/update`, computed by comparing the presentation before and after: changed properties,
  shapes and slides added, removed and moved (the longest run already in order stays), and text
  as the characters typed, deleted or formatted. Commands that change only view state (selection,
  zoom, the current slide) send nothing; the selection goes to awareness.
- **Remote changes** are applied between frames, never during a drag or a composition (they wait
  for it to end), and only the slides they touched are read again. The caret stays where it was
  in the text, wherever others typed. A message that can't be decoded or applied is dropped and
  shown in the status bar; an update that writes another app's `meta` is dropped and says "This
  room is for <app> <schema>". Nothing a peer sends can crash DeckCraft (panics in the Yjs
  library are caught too; values a reader can't take are skipped one by one).
- **Sync.** A received `sync1` is answered with `sync2` addressed to the asker (`to` = its
  `from`); when its state vector shows changes DeckCraft lacks, DeckCraft asks back with its own
  `sync1`. `sync1` also goes out when the transport reports `online` after another state. An
  update that skips some of a peer's changes waits until they arrive (and DeckCraft asks for
  them).
- **Encryption** (a `key`, 32 bytes): AES-256-GCM per protocol 2 section 4; `sync2` with
  `snapshot: true` and the answer to `collab/snapshot` use the method name `collab/snapshot`.
  After 10 messages that don't decrypt, the status bar asks "Wrong room key?".
- **`collab/snapshot`** is answered with the whole state (encrypted when keyed) and, with
  `file: true`, the presentation as `.pptx` bytes, except in an encrypted room: there DeckCraft
  answers `fileRefused` instead, so the plaintext never reaches the transport.
- **Offline.** When the transport reports `offline` or exits, editing goes on; an exited
  transport is started again after 3 s (a crash counts as a plugin failure) and the room rejoins
  with `sync1`, which merges both sides' edits.
- **Recovery.** Every 30 s the desktop app writes each room's whole state into the AutoRecover
  folder (`rooms/room-<hash>.json`); joining that room again (`mode: join`) merges it first, so
  edits made before a crash aren't lost. Leaving deletes the file.
- **Limits.** An update is at most 16 MiB before base64 (a larger local change is announced with
  `sync1` so peers fetch it; a larger remote one is refused), awareness 64 KiB and 20 a second, 16
  rooms per transport, one room per presentation.

## Tests

- `crates/collab/src/tests.rs`: two and three replicas with random concurrent edits (slides
  added, removed and moved; shapes added, removed, moved, recoloured, grouped and reordered;
  typing, deleting and formatting text in shared text boxes; notes) and updates delivered out of
  order and twice converge, and each model equals its own document; typing in one text box
  merges and undo is per user; moves of different shapes both stay; relative positions follow
  others' typing; hostile updates are refused. `DECKCRAFT_FUZZ_SEEDS` soaks it.
- `crates/engine/tests/collab_convergence.rs`: two and three sessions run random command sequences
  (slides, shapes, typing keystroke by keystroke with others' updates arriving in between,
  formatting, grouping, notes, transitions, per-user undo and redo) with updates exchanged in
  random order and duplicated; they end identical, equal to a newcomer.
  `DECKCRAFT_COLLAB_SEEDS`, `DECKCRAFT_COLLAB_STEPS`, `DECKCRAFT_COLLAB_SEED` soak and replay it.
- `crates/engine/tests/collab_conformance.rs`: replays craft-collab's DeckCraft schema-1
  transcripts (`CRAFT_COLLAB_DIR`, or a `craft-collab-deckcraft` or `craft-collab` checkout next to
  this one).
- `crates/engine/tests/collab_push.rs`: a fake transport (the test binary) relays between two
  sessions through the push channel: joining from a plugin command's result, sync, edits both
  ways, `collab/snapshot` with the `.pptx`, awareness, per-user undo, leaving.
- `crates/engine/tests/collab_loopback.rs`: two DeckCraft processes through craft-collab's
  loopback transport (`CRAFT_COLLAB_LOOPBACK` = the built plugin).
- `crates/ui-egui/src/collab_tests.rs`: the Share button, its state in a session, the people's
  circles and "+N", going to a person's selection, the popover's Leave, the Share window.

## Not yet

- Masters, layouts and the theme are shared as whole JSON values: two people editing the same
  master at once keep the later. Tables (their cells' text) and charts are whole values inside
  their shape's `kind`; SmartArt's outline is one value (`diagram`).
- The web build has no add-ins; it will speak the y-websocket binding natively.
- yrs garbage collection is off (GridCraft's fuzzer crashed with it and the undo manager):
  deleted content stays in the document's history.
- A text position in a table cell or in a cell being edited isn't sent to others (only the
  table's selection is).

## Revision history

| Date | Change | Summary |
|---|---|---|
| 2026-10-15 | major | Created: plugin protocol 2 (2.0.1) in DeckCraft: push channel, rooms, schema 1 (texts as Yjs text with formatting attributes, ids that are the model's), encryption, per-user undo, awareness with remote selections and carets, recovery; Share and the people in the title bar |
