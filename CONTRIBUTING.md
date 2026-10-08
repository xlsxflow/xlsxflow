# Contributing

Bug reports and pull requests are welcome.

## Reporting a bug

Open an issue with:

- the version of `@xlsxflow/core` and your runtime (Node, Bun, browser or Workers, with versions);
- the smallest file or code that shows the problem. If the file is private, describe how it was made (Excel, LibreOffice, another library);
- what you expected and what happened.

For security issues, follow [SECURITY.md](SECURITY.md) instead.

## Making a change

```bash
pnpm install
cd packages/core
pnpm test        # vitest
npx tsc --noEmit # type check
pnpm build
```

- Keep the library free of runtime dependencies. It only uses Web APIs (`ReadableStream`, `CompressionStream`, `Blob`, Web Crypto), so it runs in browsers and Workers as well as Node.
- Add a test that fails without your change. Files that need an Excel-made fixture go in `packages/core/test/fixtures`.
- Keep memory flat: rows should stream through, not be collected.
- Describe user-visible changes in `packages/core/CHANGELOG.md`.

By contributing, you agree that your contribution is licensed under the MIT License.
