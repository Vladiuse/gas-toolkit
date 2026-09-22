# gas-toolkit

Utilities for Google Apps Script projects, written once and reused across
spreadsheets.

Plain ES6+ JavaScript for the V8 runtime. No build step, no npm dependency:
files are copied into the script editor as they are. npm is used here only to
run ESLint.

## orm

A sheet read as records instead of cell ranges. Columns are found by their
header text, so inserting a column breaks nothing.

```js
const PLAYERS_META = new SheetMeta({
    model: Player,
    titles: { partnerId: "Partner ID", hold: "Hold", comment: "Comment" },
    markerAttribute: "partnerId",
});

const players = new Manager(sheet, PLAYERS_META);

players.get({ partnerId: 42 }).hold;                     // one record, or an error
players.filter({ hold: "yes" }).update({ comment: "" }); // a selection, written back
players.add([{ "Partner ID": 43, Hold: "yes" }]);        // appended below the last record
```

`SheetHeaders` holds the header, `SheetMeta` says how the sheet is laid out,
`Manager` reads and writes it, `RecordSet` is a selection of records, `Model`
is one row.


## edit_router

One `onEdit` split into handlers, each with the filters that decide when it
runs. A handler that throws is logged and does not stop the others.

```js
const router = new EditRouter(logger).register({
    name: "hold comment",
    sheet: (name) => name.endsWith("players"),
    column: (sheet) => new SheetHeaders(sheet).get("Hold"),
    valueChange: "checked",
    run: (context) => {
        context.sheet.getRange(context.row, 5).setValue("held by " + Session.getActiveUser().getEmail());
    },
});

function onEdit(event) {
    router.handle(event);
}
```

Filters: `sheet` (name, sheet or predicate), `column` (number or a function of
the sheet), `firstDataRow` (edits above it are ignored, default 2),
`valueChange` (`any`, `filled`, `cleared`, `checked`, `unchecked`) and
`customFilter` for anything else. An omitted filter accepts everything.


## logging

Writes to the execution log and to a `_LOG` sheet, colored by level. Never
throws: a failed log must not break the code that called it.

```js
const logger = new SheetLogger();

logger.info("Imported 120 players");
logger.warning("No previous sheet found");
logger.error("API is down: " + error, { toast: true }); // also shown to the user
```

## storage

Key-value storage in a `_STORAGE` sheet, in the spirit of localStorage. Writes
are wrapped in a script lock, so two triggers cannot append the same key twice.

```js
const storage = new SheetStorage();

storage.set("last_import", "2026-09-22");
storage.get("last_import"); // "2026-09-22"
storage.get("missing", "fallback"); // "fallback"
storage.has("last_import"); // true
storage.remove("last_import");
```