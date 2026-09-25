First preview of Zendo: spreadsheets (Ledger) and PDF editing (Hanko) for macOS. Expect rough edges, and please report problems in [Issues](https://github.com/Zedyas/zen-forge/issues).

## Install

1. Download **Zendo-0.1.0-arm64.dmg** below. It runs on Macs with Apple silicon (M1 or newer).
2. Open it and drag **Zendo** into **Applications**.
3. Open Zendo. macOS says it can't verify the app, because this preview isn't notarized by Apple yet. Click **Done**.
4. Open **System Settings → Privacy & Security**, scroll down, click **Open Anyway** next to Zendo, and confirm.

You only do this once.

## What's in it

- **Ledger:** Excel formulas and formatting, sort, AutoSum, find and replace. Opens and saves Excel and CSV files.
- **Hanko:** add text, highlights and signatures, fill in forms, redact, rotate and reorder pages, combine PDFs, turn photos into a PDF.
- One window with tabs, and a Home tab for new and recent files.

## Verify the download (optional)

- `SHA256SUMS.txt` has the file's checksum: `shasum -a 256 Zendo-0.1.0-arm64.dmg`
- Proof it was built by this repository: `gh attestation verify Zendo-0.1.0-arm64.dmg --repo Zedyas/zen-forge`

Zendo is licensed under GPL-3.0. Third-party licenses are inside the app, in `Zendo.app/Contents/Resources/licenses`.
