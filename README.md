# Bridge_Releases

Public feed of the Excel VBA macros used by the Bridge design workbooks: model builders and
result readers for MIDAS Civil NX, and sheet, report-list and window tools. The workbooks'
built-in updater downloads them from here.

Everything in this repository is generated from a private source repository by
`scripts/make_manifest.py --publish`. Do not edit files here: the next publish overwrites them.

## Files

| File | Holds |
|---|---|
| `manifest.txt` | One line per module: `id\|version\|changelog\|needs\|module`. `#` lines are comments. |
| `<id>.bas` | The full module text for that id, with `Attribute VB_Name` so it can also be imported by hand. |
| `.gitattributes` | `*.bas -text`: the modules keep their CRLF line ends, which VBA import needs. |

Manifest fields:

- `id`: the module's `SCRIPT_ID`, also the file name (`midas-core` -> `midas-core.bas`).
- `version`: `YYYY-MM-DD` plus one letter (`2026-11-19a`). The updater compares these labels
  as text, so every changed module gets a new label.
- `changelog`: short notes on the changes, newest first; may be empty.
- `needs`: ids this module calls, comma separated, may be empty.
- `module`: the module's name in the VBA project.

Every line has all five fields; the updater reads them by position.

## How a workbook updates

1. Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings > tick
   **Trust access to the VBA project object model**. It is off by default and set per
   machine. Without it the updater can only report what is out of date.
2. In the workbook: the **MAKROLAR** button > Yardım > **Makro güncellemesi**, or Alt+F8 >
   `CheckBridgeMacroUpdates`.
3. The updater reads `manifest.txt`, compares it with the `SCRIPT_VERSION` of each module in
   the workbook, lists what is older, and replaces those modules only after a Yes.

What it does and does not do:

- It replaces the modules a workbook already has. It adds a module only when an installed one
  needs it (the `needs` field, for example the shared `midas-core`). Any other new module is
  imported once by hand (VBA editor: File > Import File).
- A module newer than the feed copy is shown as ahead and never replaced.
- A download without a `SCRIPT_ID`, or with a version other than the manifest's, is refused.
  GitHub's raw files can lag a few minutes after a publish, so a refused download right after a
  release only means: try again later.
- It replaces module code only. UserForm layouts are not shipped; the macro window builds its
  forms at run time.

## Not here

This feed holds macro code and the manifest only. Workbooks, models and API keys are never
published here.
