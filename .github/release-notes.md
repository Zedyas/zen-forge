Zendo is a simplified office suite for macOS: spreadsheets (Ledger), PDF editing (Hanko), documents (Sumi) and presentations (Slides). This is a preview, so expect rough edges, and please report problems in [Issues](https://github.com/Zedyas/zen-forge/issues).

## What's new

- **Sumi, for documents.** Write and edit Word documents and Markdown: styles, fonts, lists, tables, pictures, page setup and page numbers, with printing and PDF export.
- **Slides, for presentations.** Edit PowerPoint files: text, shapes, pictures, tables, themes and speaker notes, and play them full screen.
- **Print and Export as PDF** (⌘P) in Ledger and Hanko. Hanko prints your unsaved edits, and redacted areas print as solid black.
- **More spreadsheet files.** Open Numbers, OpenDocument and older Excel (.xls) files, and open and save .tsv.
- **Search in PDFs** (⌘F), and redact every match at once.
- **Stronger redaction.** "Also remove hidden information" now removes hidden layers too, and Zendo refuses to save a redaction that wouldn't land exactly where you drew it.
- **Updates from inside the app.** Zendo checks for new versions and downloads them for you (you can turn this off in the Zendo menu).
- **Clearer import notes.** When a file has something Zendo can't keep, each note says what happens to it.
- **Fixes and polish.** Row numbers skip hidden rows, formula suggestions appear right after =, the title bar lines up, and menus and popups have a frosted-glass look.

## Install

1. Download the `.dmg` below. It runs on Macs with Apple silicon (M1 or newer).
2. Open it and drag **Zendo** into **Applications**.
3. Open Zendo. macOS says it can't verify the app, because previews aren't notarized by Apple yet. Click **Done**.
4. Open **System Settings → Privacy & Security**, scroll down, click **Open Anyway** next to Zendo, and confirm.

You only do this once per version.

## Verify the download (optional)

- `SHA256SUMS.txt` has the file's checksum: `shasum -a 256 -c SHA256SUMS.txt`
- Proof it was built by this repository: `gh attestation verify <file>.dmg --repo Zedyas/zen-forge`

Zendo is licensed under GPL-3.0. Third-party licenses are inside the app, in `Zendo.app/Contents/Resources/licenses`.
