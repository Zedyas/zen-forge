Zendo is a simplified office suite for macOS: spreadsheets (Ledger) and PDF editing (Hanko). This is a preview, so expect rough edges, and please report problems in [Issues](https://github.com/Zedyas/zen-forge/issues).

## What's new

- **Security hardening.** A file you open can no longer reach the network, open windows or ask for permissions, and the app checks its own code for tampering before it runs.
- **Updated a library** used by the spreadsheet grid, with its security fixes.
- After updating, the recent files list on Home starts empty once.

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
