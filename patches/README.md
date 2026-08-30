# Erstan Duke Sheets patch

`duke-sheets-0.1.23-erstan.1.patch` is the source provenance for the patched
`@dukelib/sheets-wasm@0.1.23` package in `../vendor/`.

## Inputs

- Upstream repository: <https://github.com/guseggert/duke-sheets.git>
- Upstream commit: `b30fff867622703f71edea6be16f20a2d4f96193`
- Rust: `rustc 1.98.0 (88d9e12ae 2026-08-18)`
- Cargo: `cargo 1.98.0 (797e8a9bc 2026-08-05)`
- wasm-pack: `0.15.0`
- Node.js: `v24.15.0`
- npm: `11.12.1`

## Rebuild

From a clean checkout of the upstream commit:

```powershell
git apply C:\dev\erstan\react-xlsx\patches\duke-sheets-0.1.23-erstan.1.patch
wasm-pack test --node bindings\wasm
wasm-pack build bindings\wasm --target web --release --scope dukelib --out-name duke_sheets_wasm --out-dir <staging-directory>
```

wasm-pack derives the scoped name `@dukelib/duke-sheets-wasm` from the Rust
crate name. Before packing, change only the generated package's `name` field to
`@dukelib/sheets-wasm`. Keep its package version at `0.1.23` so the exact
dependency declared by `@extend-ai/react-xlsx` deduplicates to this package.
Then run:

```powershell
npm pack --pack-destination <artifact-directory>
```

The checked-in artifact is named
`dukelib-sheets-wasm-0.1.23-erstan.1.tgz`; the `erstan.1` suffix identifies the
artifact while the package metadata inside remains version `0.1.23`.

## Integrity

- Source patch SHA-256: `68F75269BE86A004583CC2BBFF72F8488BC84DCF3100A9575A3E6AF87FB51EA3`
- Package SHA-256: `FB8139843D38D02AE8B4D2165BAFDB1F5ED34E98DB216212EDE070F2C955C442`
- Packed npm SHA-1: `98b7c1bdb6e6c480a53aefd531ef3acb4978f3c1`

The combined React package is
`../vendor/extend-ai-react-xlsx-0.16.2-erstan.4.tgz`:

- Package SHA-256: `678C713F59E6549FA53012D09993B574D9DA09311B03B4DD996B69B40F057C4F`
- Packed npm SHA-1: `a30c16ce13e8e5c3045fb2720c680cefe9f04962`

Install both tarballs at the application root. The Duke package intentionally
keeps version `0.1.23`, so npm resolves the React package's exact dependency to
the same patched instance. After installing into a clean directory, validate
that resolution and the runtime round trip with:

```powershell
npm ls @dukelib/sheets-wasm @extend-ai/react-xlsx
node scripts\smoke-duke-worksheet-api.mjs <install-directory>
```

The patch adds workbook-indexed worksheet rename and move bindings, faithful
reference rewriting for the modeled workbook structures, sheet-owned metadata
remapping, and save/reopen tests. Opaque or unmodeled rich OOXML remains outside
the eligible editing path because its references cannot be rewritten safely.
