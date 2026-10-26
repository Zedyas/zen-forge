Zen Suzu is a simplified office suite for macOS: spreadsheets (Ledger), PDF editing (Hanko), documents (Sumi) and presentations (Slides). This is a preview, so expect rough edges, and please report problems in [Issues](https://github.com/Zedyas/zen-forge/issues).

## What's new

**Saving is safer**
- If another app changes a file while it is open, saving asks before replacing those changes.
- Ledger and Sumi ask before saving over a file whenever the save would change or lose something, every time a new loss appears.
- Save As won't replace a file that is open in another tab or window.
- Saving keeps the file's permissions and works through aliases (symlinks). Settings and open tabs are written so a crash can't leave them half-written.

**Ledger**
- Copy, cut and paste keep formulas and styles, and relative references move with the paste.
- The fill handle continues series: 1, 2, 3 · Jan, Feb · Monday, Tuesday · dates · Item 1, Item 2.
- Wrap text, row heights, font family and size, strikethrough (⇧⌘X).
- Merge cells, custom number formats, cell notes and links (⌘-click opens a web link). All are kept in `.xlsx` files.
- Text such as IDs and ZIP codes (`00123`) stays text, CSV files save full values and ISO dates, and `YYYY-MM-DD` dates no longer show a day early in the Americas.
- Clearing, formatting and summing a selection skips hidden rows, and sheet names follow Excel's rules.

**Sumi**
- Headers and footers with page numbers, and a different first page.
- Headings 1 to 6, and a table of contents that keeps up with them and becomes a real table of contents in Word.

**Slides**
- Find and replace across text boxes, tables and speaker notes (⌘F).
- Every insert, arrange, align and table action is in the Slide menu and the command palette.

**Hanko**
- PDFs with permission restrictions (no password needed to open) are edited as plain copies, so saving no longer damages them.
- Photos turned into a PDF come out upright and fit the paper.

**Everywhere**
- File > Open Recent lists the last 10 files, also in the command palette.

## Install

1. Download `Zen-Suzu-0.3.0-arm64.dmg` below. It runs on Macs with Apple silicon (M1 or newer).
2. Open it and drag **Zen Suzu** into **Applications**.
3. Open Zen Suzu. macOS says it can't verify the app, because previews aren't notarized by Apple yet. Click **Done**.
4. Open **System Settings → Privacy & Security**, scroll down, click **Open Anyway** next to Zen Suzu, and confirm.

You only do this once per version.

## Verify the download (optional)

- `SHA256SUMS.txt` has each file's checksum: `shasum -a 256 -c --ignore-missing SHA256SUMS.txt`
- Proof it was built by this repository: `gh attestation verify <file>.dmg --repo Zedyas/zen-forge`

Zen Suzu is licensed under GPL-3.0. Third-party licenses are inside the app, in `Zen Suzu.app/Contents/Resources/licenses`.
