# Contributing to Zendo

Zendo is maintained by its owner. Anyone can suggest changes; the maintainer reviews every suggestion and makes the final call.

## Ideas and bugs

Open an [issue](https://github.com/Zedyas/zen-forge/issues/new/choose). For a bug, say what you did, what you expected and what happened, and attach the file if you can share it.

## Code

Pull requests are welcome as suggestions. For anything large, open an issue first so the approach can be agreed before you spend time on it.

1. Fork the repository and create a branch.
2. Make the change. Add a focused test when you fix a bug in a model, a file format or the PDF engine.
3. Run the checks: `pnpm typecheck`, `pnpm lint`, `pnpm test` and `pnpm build`.
4. Open a pull request that says what changed and why.

The same checks run automatically on every pull request. A pull request is merged only after they pass and the maintainer approves it.

## Security problems

Don't open a public issue. Report them privately, as described in [SECURITY.md](SECURITY.md).
