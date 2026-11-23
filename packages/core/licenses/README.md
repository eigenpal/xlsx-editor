# HarfBuzz component notices

`harfbuzz-manifest.json` records the versions, source locations, and hashes for the bundled HarfBuzz distribution. `THIRD_PARTY_NOTICES.md` lists each notice file in this directory.

The HarfBuzz source notices retain copyright and permission comments from the pinned upstream source tree. They include optional components, so they can cover more code than the WebAssembly binary contains. The Universal Shaping Engine (USE) notice covers the override data used by the generated USE shaping table. The Unicode notice covers the Unicode data tables.

The upstream release workflow specifies Emscripten 4.0.13. The Emscripten, musl, and compiler runtime notices come from that version. These notices also cover optional runtime components. The binary hash identifies the npm artifact. It does not prove a reproducible upstream build.

When upgrading HarfBuzz:

1. Check the `harfbuzzjs` release, its HarfBuzz submodule, and its build workflow.
2. Compare all component licenses and copyright notices with their pinned sources.
3. Update the notice files and manifest. Preserve the original legal text.
4. Update the exact dependency version and binary hash.
5. Build packages, generate notices, and inspect the packed core archive.

`scripts/harfbuzz-notices.mjs` rejects changed versions, binary hashes, or notice hashes during notice generation. The source-notice file reproduces distinct copyright and permission comment blocks from upstream `src/**/*.cc`, `src/**/*.hh`, and `src/**/*.h`. Each block lists its source locations.

These files are copied from `pptx-editor/packages/core/licenses/` at commit `d4502552`. The notice texts are unchanged. There are two differences. First, the Universal Shaping Engine notice file is `HarfBuzz-USE-MIT.txt`, `harfbuzz-manifest.json` names its component "Universal Shaping Engine data", and the second paragraph of this README uses the same name. Second, the first paragraph says that `THIRD_PARTY_NOTICES.md` lists each notice file, where the source says that it links to each file, because the generated file names each path in code font. The new names follow the naming rule of `AGENTS.md`, which keeps the vendor name only in the copyright line that the license requires.
