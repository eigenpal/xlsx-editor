# @xlsx-editor.dev/i18n

The user-facing strings of the XLSX Editor chrome. The English catalog is the source of truth. A locale catalog mirrors its shape, and a `null` entry falls back to English.

## Install

The engine and each adapter install this package for you. Install it yourself only when you import a catalog, a type, or the translator from it directly:

```bash
npm install @xlsx-editor.dev/i18n
```

## Use

`createT` builds a translator from a full catalog. `deepMerge` fills a partial catalog with English:

<!-- sample: i18n-use -->

```ts
import { createT, deepMerge, en, type Translations } from '@xlsx-editor.dev/i18n';

const t = createT(en);
t('formattingBar.bold'); // "Bold"
t('grid.label', { sheet: 'Data' }); // "Worksheet Data"

const custom: Translations = { formattingBar: { bold: 'Fett' } };
const tCustom = createT(deepMerge(en, custom) as typeof en);
```

`@xlsx-editor.dev/i18n/en` is the English catalog as its own entry. In a React app, pass a catalog to the `locale` prop of `XlsxEditor` and read the translator with `useTranslation()`. See [Root props](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/layout-and-props.md#root-props) and [Hooks](https://github.com/eigenpal/xlsx-editor/blob/main/packages/react/docs/layout-and-props.md#hooks) in the React guide.

## Contribute

### Add a string

1. Add the key and its English text to `packages/i18n/en.json` in this repository.
2. Use the key in the component or the engine that shows it.
3. Run `bun run i18n:validate`. It fails when a locale file is out of step with `en.json`, or when no shipping code uses a key.

### Add a locale

To add a locale, follow these steps:

1. Run `bun run i18n:new <lang>` with a BCP 47 tag, for example `bun run i18n:new de`. The command writes `<lang>.json` with every key set to `null`, writes the entry `src/<lang>.ts`, and adds the locale to `src/index.ts` and `package.json`.
2. Open `packages/i18n/<lang>.json`, and replace the `null` values with translations. A key that stays `null` falls back to English.
3. Add the locale code to the `entries` of `@xlsx-editor.dev/i18n` in `scripts/lib-packages.mjs`, so that the API report and the consumer install check cover the new entry.
4. Run `bun run build:packages`.
5. Run `node scripts/api-extractor.mjs --local --record-forgotten-exports`.
6. Run `bun run i18n:status` to see the coverage of each locale.
7. Commit the locale files, `scripts/lib-packages.mjs`, `scripts/api-forgotten-exports.baseline.json`, and the reports in `docs/api/i18n/`, and open a pull request.

The validator and its helpers are copied from DOCX Editor. Each file header names its source.

### Refusal messages

Each refusal message is one key that holds the whole message, as in DOCX Editor. The editor shows the text of the key as it is, with its placeholders filled in. It never joins keys or cuts translated text. Translate each message as one unit. The editor picks the key from six families:

- `editing.refused.<code>` is the message of a refused command that no other family covers, one for each refusal code.
- `editing.lockedRefused.<key>` is the message of a `locked` refusal of a format, clear, cut, paste, resize, structure, merge, or sheet command. The editor picks the key from the protection rule (`lock`), then from the command:
  - The `structure` rule, or a structure command, gives `structure`.
  - The `size` rule, a column width or a row height, gives `size`.
  - The `workbook` rule, a sheet add, rename, delete, or move, gives `workbook`.
  - The `merge` rule, a merge or an unmerge, gives `merge`.
  - The `format` rule gives `cutFormats` for a cut and `pasteFormats` for a paste. It gives `clearFormats` for a clear of formats, `clearAll` for a clear of values and formats, and `format` for a format command.
  - The `cell` rule gives `cut` for a cut and `paste` for a paste. It gives `clear` for a clear of values or a clear of values and formats. A cell edit gives `editing.refused.locked`. The store gives the `cell` rule to no format or structure command; if one has it, the key is `format` or `structure`.
- `editing.notWorksheetRefused.named` and `editing.notWorksheetRefused.unnamed` are the message of a `structure-blocked` refusal whose `about` kind is `not-worksheet`. That is a move or a delete of a sheet that is not a worksheet.
- `editing.wrappedFormulaRefused.named` and `editing.wrappedFormulaRefused.unnamed` are the message of a `structure-blocked` or `sheet-range-blocked` refusal whose `about` kind is `wrapped-formula`. That is a row, column, or sheet edit that can change a formula that a cell keeps inside `mc:AlternateContent`.
- `editing.saveRefused.<code>` is the message of a refused save, one for each code that a save can return. `editing.saveRefused.other` is the message of a save that another code refused.
- `editing.refusedUnnamed.<code>` is the message of a refused command that names nothing, for a code whose message has `{subject}`. It holds no placeholder.
- `editing.saveRefusedUnnamed.<code>` is the message of such a refused save. It holds no placeholder.

`{subject}` is a name from the workbook or the words of a content kind, such as "the defined name Rates". `{limit}` is a number of cells, grouped with the group separator of the number locale, such as `16,384` or `16.384`.

## License

Apache-2.0. See [LICENSE](LICENSE).
