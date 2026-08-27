# MIRRORZ Sheets

An offline-first spreadsheet application that opens, edits and saves Excel files.
Built from scratch, Apache-2.0, no cloud, no telemetry, no account.

The goal in order: work properly offline and free for personal use first, then be
worth paying for.

> **Naming note.** The repository is called `MIRRORZ---Excel`, but the product is
> `MIRRORZ Sheets`. Shipping a commercial product with "Excel" in its name is a
> trademark problem regardless of how the software is built. The safe and
> accurate description is "opens and edits `.xlsx` files" or "compatible with
> Excel formats" - never anything implying affiliation with Microsoft. See
> `docs/naming-and-trademark.md`.

---

## What works today

Everything below is implemented and covered by tests that run against real
spreadsheet files, not mocks.

### Formats

| Format | Read | Write | Notes |
| --- | :---: | :---: | --- |
| `.xlsx` | yes | yes | Full round trip; unmodelled parts preserved byte-for-byte |
| `.xlsm` | yes | yes | The VBA project is carried through untouched, so macros survive a save |
| `.xltx` / `.xltm` | yes | yes | Template variants |
| `.xls` | yes | - | BIFF8 (Excel 97-2003). Old files open; save them as `.xlsx` |
| `.csv` / `.tsv` | yes | yes | Delimiter and encoding detection, non-destructive type inference |
| `.ods` | yes | - | OpenDocument, with OpenFormula translated to the A1 dialect |

The single most important property is **preservation**. The document model is a
projection over the file, never a replacement for it. Charts, pivot tables,
drawings, slicers and every other feature MIRRORZ cannot yet edit survive being
opened and saved untouched. Round-trip tests assert this byte-for-byte on real
workbooks, because a spreadsheet tool that silently drops your charts is worse
than one that refuses to open the file.

### Calculation

- A hand-written tokenizer and Pratt parser covering Excel's full expression
  grammar: references, ranges, 3-D references, structured table references,
  array constants, the union and intersection operators, and defined names.
- A dependency graph with incremental recalculation, iterative Tarjan ordering
  (chains tens of thousands deep do not overflow the stack), range vertices
  rather than per-cell edges, and change-based propagation.
- Circular-reference detection and Excel's iterative calculation mode.
- A growing worksheet function library, verified against a reference oracle.

### Correctness

Formula results are checked against a **differential oracle**: fixtures are
authored with openpyxl, recalculated by LibreOffice headless, and our engine must
reproduce the values a completely independent implementation computed. Where
LibreOffice and Excel genuinely disagree, we follow Excel and record the
divergence in `docs/oracle-divergences.md`.

The numeric layer reproduces Excel's deliberate departures from IEEE-754, each
derived from probe formulas rather than folklore:

- `2^3^2` is `64` - exponentiation is left-associative.
- `-2^2` is `4` - unary minus binds tighter than the exponent operator.
- `0.1+0.2` displays as `0.3` and compares equal to it.
- `(0.5-0.4-0.1)` is exactly `0`, and stays zero when multiplied by `1E20`.
- `1/3` is `0.333333333333333` - fifteen significant digits, not seventeen.
- The 1900 leap-year bug, including its effect on weekday numbering.

---

## What makes it different

Five improvements people have asked Microsoft for repeatedly, for years. Two of
them cannot be retrofitted, so they are built into the core rather than planned.

### 1. Undo that survives macros

Excel's VBA wipes the undo stack: if an automation goes wrong, you cannot take it
back. Here every mutation is a command carrying its own inverse, and a script or
macro runs inside a transaction that collapses to **one** undo entry. Undoing a
thousand-cell macro run is the same operation as undoing a keystroke.

Undo is also per-document (Excel shares one stack across every open workbook),
unlimited, and a **tree** rather than a stack - undoing and then doing something
else does not destroy the abandoned branch, and any point in the history can be
jumped to directly.

### 2. Data entry that does not corrupt data

Typing `SEPT1` gives you `SEPT1`. Pasting `007` gives you `007`. A nineteen-digit
order number keeps all nineteen digits. A genetics working group renamed human
genes because Excel kept turning them into dates; that is a statement about the
software, not about the genes.

Inference is not disabled - it is made **reversible**. The literal text is stored
alongside the inferred value, known false-positive shapes stay text and say why,
and an ambiguous date like `3/4/2024` is flagged rather than silently resolved.
Imports show what would change meaning *before* committing.

### 3. Formula debugging that names the culprit

Excel makes you chase tracer arrows one hop at a time to find which cell broke a
formula. The dependency graph already knows, so `explain()` walks precedents and
names the cell that actually **originated** an error, distinguishing it from the
cells merely carrying it onward.

### 4. Nothing proportional to the grid

A spreadsheet is 1,048,576 by 16,384 addresses and almost entirely empty. Cells
are stored sparsely on packed integer keys, styles are interned so a uniformly
formatted million-row sheet costs one format record, whole-column references are
clipped to the used range before materialising, and range dependencies are one
graph vertex rather than a hundred thousand edges.

### 5. Honest failure

A damaged sheet degrades to a warning and the rest of the workbook opens. An
encrypted file says so instead of producing garbage. A pre-BIFF8 file is named as
unsupported rather than misparsed. Corrupt archive data is caught by CRC rather
than returned as cell values.

---

## Repository layout

```
packages/
  core/        value model, A1 addressing, date serials, sparse sheet store,
               style interning, the command log, non-destructive entry
  formats/     zip, XML, OPC packaging, xlsx read/write, xls (BIFF8),
               CSV/TSV, CFB (OLE2), number formatting
  formula/     lexer, parser, AST, evaluator, dependency graph, engine,
               worksheet function library
  vba/         MS-OVBA extraction - reads macro source, does not execute it
  grid/        canvas grid renderer
  licensing/   offline licence verification
apps/
  desktop/     Electron shell and React UI
fixtures/      generated test workbooks in every supported format
tools/         fixture generator
docs/          design notes, oracle divergences, pricing, trademark guidance
```

Nothing in `core`, `formats` or `formula` has a runtime npm dependency. ZIP, XML,
CFB and the compression codecs are all implemented directly. That is deliberate:
this is the code path every opened file crosses, and it should carry no
supply-chain surface and no licence entanglement in a product intended to be sold.

---

## Building and testing

```bash
npm install
npm test           # the full suite
npm run typecheck
npm run fixtures   # regenerate test workbooks (needs Python and LibreOffice)
```

The fixture generator authors feature-rich workbooks with openpyxl, then uses
LibreOffice headless to convert them to every other format and to recalculate the
`.xlsx` so cached values become a correctness oracle. LibreOffice is used only as
an offline build and test tool, invoked as a subprocess - it is never linked into
or shipped with MIRRORZ, so its LGPL and MPL terms do not touch our Apache-2.0
code.

---

## Licence and provenance

Apache-2.0. Every dependency is permissively licensed; nothing GPL, LGPL or AGPL
is used, linked or vendored. In particular HyperFormula is deliberately avoided:
it is GPLv3-or-commercial and writes `#LIC!` errors into cells at runtime without
a key, which would be doubly disqualifying here.

The Excel file formats are implemented from the published ECMA-376 standard and
Microsoft's Open Specifications, which is the same basis LibreOffice, Gnumeric,
Apache POI and SheetJS work from.

Fonts: Calibri and Cambria cannot be redistributed. Carlito and Caladea are the
metric-compatible substitutes, so a workbook laid out in Calibri still measures
correctly.

---

## Status

Under active development. The formats, calculation engine, command log and entry
layer are working and tested. The desktop shell, grid renderer and licensing are
in progress. See `docs/` for the design notes behind each decision.
