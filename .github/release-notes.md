Zen Suzu is a simplified office suite for macOS: spreadsheets (Ledger), PDF editing (Hanko), documents (Sumi) and presentations (Slides). This is a preview, so expect rough edges, and please report problems in [Issues](https://github.com/Zedyas/zen-forge/issues).

## What's new

- **Zendo is now Zen Suzu.** Only the name changes in this version.

## Updating from Zendo

1. In Zendo, click **Download** on Home, or download the `.dmg` below.
2. Quit Zendo, then drag **Zen Suzu** into **Applications**. It goes beside Zendo instead of replacing it.
3. Open Zen Suzu. The first time, it copies your settings, open tabs, recent files and saved signatures from Zendo.
4. Move **Zendo** from Applications to the Bin. Its data stays in `~/Library/Application Support/Zendo`, which you can delete once Zen Suzu shows your files.
5. If Zendo was the default app for a file type, set Zen Suzu instead: select the file in Finder, press ⌘I, choose Zen Suzu under **Open with**, then click **Change All…**.

## Install

1. Download `Zen-Suzu-0.2.1-arm64.dmg` below. It runs on Macs with Apple silicon (M1 or newer). `Zendo-0.2.1-arm64.dmg` is the same file under the old name, for Zendo's updater.
2. Open it and drag **Zen Suzu** into **Applications**.
3. Open Zen Suzu. macOS says it can't verify the app, because previews aren't notarized by Apple yet. Click **Done**.
4. Open **System Settings → Privacy & Security**, scroll down, click **Open Anyway** next to Zen Suzu, and confirm.

You only do this once per version.

## Verify the download (optional)

- `SHA256SUMS.txt` has each file's checksum: `shasum -a 256 -c --ignore-missing SHA256SUMS.txt`
- Proof it was built by this repository: `gh attestation verify <file>.dmg --repo Zedyas/zen-forge`

Zen Suzu is licensed under GPL-3.0. Third-party licenses are inside the app, in `Zen Suzu.app/Contents/Resources/licenses`.
