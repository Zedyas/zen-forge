# Zendo

Simplified office suite for macOS because I hate using Numbers and needed functionality of a PDF editor without the cost. More to come.

![Zendo Home: the applications, quick actions and recent files](docs/media/home.png)

## How it works

One window with tabs. Home is the first tab, and spreadsheets, PDFs, documents and presentations open as tabs beside it. Right-click a tab to move it to its own window, or press ⇧⌘N for a new one.

![Tabs: open a spreadsheet and a PDF from Home; a second spreadsheet joins the spreadsheet group](docs/media/tabs.gif)

## Applications

| Application | For | Files | Status |
|---|---|---|---|
| **Ledger** | Spreadsheets | Excel (`.xlsx`, `.xlsm`, `.xls`), `.csv`, `.tsv`, OpenDocument (`.ods`), Numbers (`.numbers`) | Available |
| **Hanko** | PDF editing | `.pdf` | Available |
| **Sumi** | Documents | Word (`.docx`), Markdown (`.md`) | Available |
| **Slides** | Presentations | PowerPoint (`.pptx`) | Available |

### Ledger

![Ledger: bold headers, currency format, AutoSum, recalculation and the inspector](docs/media/ledger.gif)

- Excel formulas and formatting
- Sort, AutoSum, and find and replace
- Print, and export as PDF
- Opens and saves Excel, CSV and TSV files, and opens Numbers, OpenDocument and older Excel files (saved as `.xlsx`). If a file has something Ledger can't keep, such as macros, Ledger tells you when it opens and saves a copy instead of overwriting the original.

### Hanko

![Hanko: highlight, redact, draw a signature, place it and type the date](docs/media/hanko.gif)

- Add text, highlights and your signature
- Fill in forms
- Redact: removes what's under the box from the file, like Adobe Acrobat Pro. Search for text and redact every match at once.
- Rotate, delete and reorder pages, combine PDFs, turn photos into a PDF, and print

### Sumi

- Word documents and Markdown: styles, fonts, lists, tables and images
- Page setup: paper size, margins and page numbers
- Find and replace, spellcheck, print, and export as PDF
- If a Word file has something Sumi can't keep, such as comments or tracked changes, Sumi tells you when it opens and offers to save a copy.

### Slides

- Text, shapes, pictures and tables, with layouts and colour themes
- Speaker notes, and a full-screen slideshow (⌥⌘P)
- Print, and export as PDF
- If a PowerPoint file has something Slides can't keep, such as charts or animations, Slides tells you when it opens and offers to save a copy.

## Roadmap

In order, with the next item first:

1. Templates for documents and presentations
2. More Ledger tools: borders, filters, conditional formatting, dropdown lists and charts
3. iPhone photos (`.heic`) in PDF from images
4. Notarized builds, so Zendo opens without the Open Anyway step and installs its own updates
5. Windows support: the same app, adjusted to run and fit on Windows

Not planned: `.xlsb` and other macro-enabled formats, and Apple Pages or Keynote files.

## Install

**Download:** get the latest `.dmg` from [Releases](https://github.com/Zedyas/zen-forge/releases). It runs on Macs with Apple silicon. Previews aren't notarized by Apple yet, so the first time you open Zendo, click **Done**, then go to **System Settings → Privacy & Security** and click **Open Anyway**.

**Updates:** when it opens, Zendo checks GitHub for a new version, at most once a day; you can turn this off in the Zendo menu. When there is one, click **Download** on Home. Zendo downloads the `.dmg`, checks it against the release's checksum and opens it. Quit Zendo, drag the new Zendo into Applications and choose **Replace**.

**Build from source:** requires [Node.js](https://nodejs.org) 22.13 or newer and [pnpm](https://pnpm.io) 10.

```sh
git clone https://github.com/Zedyas/zen-forge.git
cd zen-forge
pnpm install
pnpm package:mac
```

The app is written to `dist/mac-arm64/Zendo.app`. Drag it into `/Applications`. `pnpm release:mac` builds the `.dmg` instead.

**Make Zendo the default app for a file type:** in Finder, select a `.xlsx`, `.pdf`, `.docx` or `.pptx` file, press ⌘I, choose Zendo under **Open with**, then click **Change All…**.

## Keyboard shortcuts

Every command is also in the menu bar and in the command palette (⌘K). Hover any toolbar button to see its name and shortcut.

| Tabs and windows | |
|---|---|
| ⇧⌘H | Home |
| ⌘1 … ⌘8 | Go to tab 1 to 8 (Home is tab 1) |
| ⌘9 | Go to the last tab |
| ⌃Tab / ⌃⇧Tab, or ⇧⌘] / ⇧⌘[ | Next / previous tab |
| ⌘W | Close tab (on Home, closes the window) |
| ⇧⌘N / ⇧⌘W | New window / close window |
| ⌘K | Command palette |

| Files and editing | |
|---|---|
| ⌘N | New spreadsheet |
| ⌘O | Open |
| ⌘S / ⇧⌘S | Save / Save As |
| ⌘P | Print |
| ⌘Z / ⇧⌘Z | Undo / Redo |
| ⌥⌘T | Show or hide the toolbar |
| ⌥⌘I | Show or hide the inspector |

| Ledger | |
|---|---|
| ⌘B / ⌘I / ⌘U | Bold / Italic / Underline |
| ⌘F | Find and replace |

| Hanko | |
|---|---|
| V · T · H · D · R · W | Select, Text, Highlight, Draw, Rectangle, White-out |
| S · X | Signature, Redact |
| ⌘L / ⌘R | Rotate page left / right |
| ⌘⌫ | Delete page |
| ⌘F | Find, and redact every match |
| ⌘= / ⌘- / ⌘0 | Zoom in / Zoom out / Actual size |

| Sumi | |
|---|---|
| ⌘B / ⌘I / ⌘U / ⇧⌘X | Bold / Italic / Underline / Strikethrough |
| Tab / ⇧Tab | Indent / outdent a list item |
| ⌘↩ | Page break |
| ⌘F | Find and replace |
| ⌘= / ⌘- / ⌘0 | Zoom in / Zoom out / Actual size |

| Slides | |
|---|---|
| ⌥⌘N | New slide |
| ⌥⌘P | Play the slideshow; → and ← move, Escape stops |
| ⌘B / ⌘I / ⌘U | Bold / Italic / Underline |
| ⌘C / ⌘V / ⌘D | Copy / Paste / Duplicate the selected objects |
| Arrow keys | Nudge the selection by 1 pt, or 10 pt with Shift |

## Development

```sh
pnpm dev          # run the app with hot reload
pnpm test         # unit tests (Vitest)
pnpm typecheck    # TypeScript, strict mode
pnpm lint         # ESLint, zero warnings allowed
pnpm build        # production build into out/
pnpm package:mac  # build and package Zendo.app into dist/
pnpm release:mac  # build the downloadable .dmg into dist/
```

**Releasing:** update `version` in `package.json` and `.github/release-notes.md`, then push a tag such as `v0.2.0`. The Release workflow builds the `.dmg` on GitHub, records where it came from (checksum and build attestation) and creates a draft release to publish.

### How it is built

- **Electron** runs the app. The main process (`electron/main`) owns windows, the native menu bar, file dialogs, file reads and writes, and the saved session. Every window loads the same React page; a window is a row of tabs that starts on Home.
- **The renderer** (`src/renderer/src`) is React 19 with Zustand stores. One store holds the window's tabs; each tab records which application edits it. The renderer reaches the file system only through the `window.desktop` bridge defined in `electron/preload`, and a lint rule keeps that bridge inside `services/file`.
- **Security.** The page loads from a private `app://` address under a strict Content Security Policy (no network, no eval), cannot navigate away or open windows, and gets no permissions beyond writing to the clipboard. The main process answers only that page. Electron fuses lock the packaged app to its own integrity-checked code. See `electron/main/security.ts`.
- **Updates.** Zendo's only network use is the main process checking this repository's GitHub releases and, when you click Download, fetching the release's `.dmg` and `SHA256SUMS.txt` (`electron/main/updates.ts`). The `.dmg` is checked against the checksum before it opens.
- **Printing.** A module builds a printable copy of the document (`services/print`), and the main process prints it or turns it into a PDF. Hanko prints page images with redaction boxes painted in, so no text under a box reaches the printer or a PDF saved from the print dialog.
- **Commands** are defined once in `src/shared/commands.ts`. The native menu, the command palette, keyboard shortcuts and toolbar tooltips all read from that list, and each command says which kind of tab it applies to.
- **Editors** (`modules/sheets`, `modules/pdf`, `modules/docs`, `modules/slides`) each export a handler with `run`, `save` and `release`. The shell finds the handler by the tab's kind, so saving or closing a tab works whether or not its editor is on screen.
- **Ledger** uses [HyperFormula](https://hyperformula.handsontable.com) for formulas, [Glide Data Grid](https://grid.glideapps.com) to draw the grid, [ExcelJS](https://github.com/exceljs/exceljs) for `.xlsx`, [Papa Parse](https://www.papaparse.com) for `.csv` and `.tsv`, and [SheetJS](https://sheetjs.com) for `.xls`, `.ods` and `.numbers`.
- **Hanko** uses [pdf.js](https://mozilla.github.io/pdf.js/) to draw pages, [pdf-lib](https://github.com/cantoo-scribe/pdf-lib) to write every change back into the file, and [MuPDF](https://mupdf.com) to apply redactions. MuPDF's WebAssembly module loads only when a redaction is saved.
- **Sumi** uses [TipTap](https://tiptap.dev) (ProseMirror) for editing, its own `.docx` reader (`modules/docs/io/docx-read.ts`), [docx](https://docx.js.org) to write `.docx`, and TipTap's Markdown support. A `.docx` that Sumi saves reopens exactly as saved, and before saving over a file Sumi reads back what it wrote and asks first if anything would change.
- **Slides** keeps each presentation as plain data (slides, text boxes, shapes, pictures and tables, in points), reads `.pptx` with [pptxtojson](https://github.com/pipipi-pikachu/pptxtojson) and writes it with [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

```text
electron/
  main/        windows, menus, file access, settings and session
  preload/     the window.desktop bridge
src/
  shared/      applications, commands and types used by both processes
  renderer/src/
    app/       the window, tabs, Home, commands, document actions
    ui/        shared components: tabs, toolbar, inspector, tooltips, icons
    services/  file access, recent files, previews
    modules/
      sheets/  Ledger: workbook model, file formats, grid, toolbar
      pdf/     Hanko: PDF engine, pages, markup, signatures, forms
      docs/    Sumi: editor, styles and page setup, .docx and Markdown files
      slides/  Slides: slide model, canvas, slideshow, .pptx files
scripts/       the app icon generator (pnpm icon:generate) and the license notices for releases
.github/       checks and release workflows, Dependabot, issue forms
patches/       fixes for Glide Data Grid, applied by pnpm install
vendor/        SheetJS, installed from a copy in the repository as its authors recommend
build/         the app icon used for packaging
```

The patch in `patches/` changes two things in Glide Data Grid. It loads the cell editor up front instead of on first use; without that, the first key typed into a cell is sometimes dropped. It also lets row numbers skip hidden rows.

## Contributing

Suggestions are welcome as [issues](https://github.com/Zedyas/zen-forge/issues); code changes are made by the maintainer. See [CONTRIBUTING.md](CONTRIBUTING.md), and report security problems privately as described in [SECURITY.md](SECURITY.md).

## License

Copyright © 2026 Zed Y.

Zendo is licensed under the [GNU General Public License v3.0](LICENSE) (`GPL-3.0-only`). You can use, study, change and share it; if you distribute Zendo or a modified version, you must share its source under the same license.

It uses GPL-3.0 because the formula engine, HyperFormula, is GPL-3.0, and a packaged `Zendo.app` includes it. Using one license for both keeps the source and the app under the same terms.

Redaction uses MuPDF, which is licensed under the GNU Affero General Public License v3.0 (AGPL-3.0). AGPL-3.0 and GPL-3.0 can be combined; MuPDF keeps its own license inside the app. AGPL adds one rule beyond GPL: if you run a modified MuPDF for users over a network, you must offer them its source. A desktop app like Zendo is not affected. The other dependencies use permissive licenses (MIT, and Apache-2.0 for pdf.js and SheetJS) that are compatible with GPL-3.0.
