// After pptx-editor/packages/vue/test/parity.test.ts at d4502552 (EigenPal). Differences: each
// case edits an XLSX Editor part and prop, and the cases for compound parts, named slots, and
// teleported parts are left out, because the XLSX adapters have none of them.
//
// The parity contract against the real adapters, and each kind of drift it must catch. Runs in
// the DOM test run, because the Vue side is read by mounting the components.

import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parityErrors } from '../../../scripts/check-parity-contract.mjs';
import { vueAdapterShape, type VueAdapterShape } from './parity-shape.ts';

const root = join(import.meta.dir, '../../..');
const read = (path: string): string => readFileSync(join(root, path), 'utf8');

let shape: VueAdapterShape;
beforeAll(async () => {
  shape = await vueAdapterShape(JSON.parse(read('scripts/parity/parity.contract.json')));
});

const inputs = () => ({
  reactApi: read('docs/api/react/index.api.md'),
  vueApi: read('docs/api/vue/index.api.md'),
  vue: shape,
  contract: JSON.parse(read('scripts/parity/parity.contract.json')),
});

describe('check:parity-contract', () => {
  test('the committed adapters match the contract', () => {
    expect(parityErrors(inputs())).toEqual([]);
  });

  test('the probes see the class and the slots of the paired components', () => {
    expect(shape.XlsxEditor.classLands).toBe(true);
    expect(shape.XlsxEditorRoot.slots).toContain('default');
    expect(shape.XlsxEditorViewport.slots).toContain('default');
    expect(shape.XlsxEditorViewport.classNameLands).toBe(true);
    expect(shape.XlsxEditorContent.classNameLands).toBe(true);
    expect(shape.XlsxEditorToolbar.classLands).toBe(true);
    expect(shape.XlsxEditorToolbar.classNameLands).toBe(true);
    expect(shape.XlsxEditorToolbarSlot.slots).toContain('default');
    expect(shape.XlsxEditorToolbarGroup.slots).toContain('default');
    expect(shape.XlsxEditorToolbarGroup.classNameLands).toBe(true);
  });

  test('the value probes see the paired props that show in a part alone', () => {
    expect(shape.XlsxEditorViewport.values).toEqual({ label: true });
    expect(shape.XlsxEditorToolbar.values).toEqual({ hidden: true, preset: true });
    expect(shape.XlsxEditorToolbarSlot.values).toEqual({ slotId: true, hidden: true });
    expect(shape.XlsxEditorToolbarGroup.values).toEqual({
      id: true,
      label: true,
      labelKey: true,
      hidden: true,
    });
  });

  test('a prop inherited through extends must be classified', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /export interface XlsxEditorRootProps \{/,
      'export interface XlsxEditorRootProps {\n    readonly zoom?: number;'
    );
    const errors = parityErrors(input);
    expect(errors).toContain(
      'XlsxEditorRoot (XlsxEditorRootProps): React prop "zoom" is not classified in the contract.'
    );
    expect(errors).toContain(
      'XlsxEditor (XlsxEditorProps): React prop "zoom" is not classified in the contract.'
    );
  });

  test('a prop removed with Omit<> is not a React prop', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      "export interface XlsxEditorProps extends Omit<XlsxEditorRootProps, 'children'> {",
      'export interface XlsxEditorProps extends XlsxEditorRootProps {'
    );
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): React prop "children" is not classified in the contract.'
    );
  });

  test('a base that is not in the snapshot fails', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      "export interface XlsxEditorProps extends Omit<XlsxEditorRootProps, 'children'> {",
      'export interface XlsxEditorProps extends HiddenBase {'
    );
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): XlsxEditorProps extends HiddenBase, which is not an interface of the snapshot; its members cannot be checked.'
    );
  });

  test('a paired prop whose type differs between the adapters fails', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /(export interface XlsxEditorViewportProps \{[^}]*?readonly label\?: )string \| false;/,
      '$1string;'
    );
    expect(parityErrors(input)).toContain(
      'XlsxEditorViewport (XlsxEditorViewportProps): paired "label" type drift: React `string`, Vue `false | string`.'
    );
  });

  test('a paired prop of the namespace-typed XlsxEditor whose Vue type drifts fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(
      /(\nconst xlsxEditorProps: \{[\s\S]*?readonly opening: \{\s*readonly type: )BooleanConstructor;/,
      '$1StringConstructor;'
    );
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): paired "opening" type drift: React `boolean`, Vue `string`.'
    );
  });

  test('a component whose Vue prop types cannot be read fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(
      'export type XlsxEditorProps = ExtractPublicPropTypes<typeof xlsxEditorProps>;',
      ''
    );
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): the Vue snapshot has no runtime prop types for XlsxEditor.'
    );
  });

  test('a Vue emit that the contract does not map fails', () => {
    const input = inputs();
    const toolbar = shape.XlsxEditorToolbar;
    input.vue = {
      ...shape,
      XlsxEditorToolbar: { ...toolbar, emits: [...toolbar.emits, 'resize'] },
    };
    expect(parityErrors(input)).toContain(
      'XlsxEditorToolbar (XlsxEditorToolbarProps): Vue emit "resize" is not declared in reactCallbacksAsVueEmits.'
    );
  });

  test('a default slot that does not render fails', () => {
    const input = inputs();
    input.vue = { ...shape, XlsxEditorViewport: { ...shape.XlsxEditorViewport, slots: [] } };
    expect(parityErrors(input)).toContain(
      'XlsxEditorViewport (XlsxEditorViewportProps): Vue XlsxEditorViewport does not render its "default" slot.'
    );
  });

  test('a className that does not reach an element fails', () => {
    const input = inputs();
    const content = shape.XlsxEditorContent;
    input.vue = { ...shape, XlsxEditorContent: { ...content, classNameLands: false } };
    expect(parityErrors(input)).toContain(
      'XlsxEditorContent (XlsxEditorContentProps): the className prop of Vue XlsxEditorContent does not reach a rendered element.'
    );
  });

  test('a Vue part that ignores a paired prop value fails', () => {
    const input = inputs();
    const toolbar = shape.XlsxEditorToolbar;
    input.vue = { ...shape, XlsxEditorToolbar: { ...toolbar, values: { hidden: false } } };
    expect(parityErrors(input)).toContain(
      'XlsxEditorToolbar (XlsxEditorToolbarProps): Vue XlsxEditorToolbar ignores a value of its paired "hidden" prop.'
    );
  });

  test('stale entries fail', () => {
    const input = inputs();
    input.contract.components.XlsxEditorToolbar.vueExclusive = { gone: 'Removed.' };
    input.contract.components.XlsxEditorToolbar.deferredInVue = { gone2: 'Removed.' };
    input.contract.exports.vueExclusive.NoSuchExport = 'Removed.';
    const errors = parityErrors(input);
    expect(errors).toContain(
      'XlsxEditorToolbar (XlsxEditorToolbarProps): vueExclusive "gone" is missing in Vue (contract stale).'
    );
    expect(errors).toContain(
      'XlsxEditorToolbar (XlsxEditorToolbarProps): deferredInVue "gone2" is missing in React (contract stale).'
    );
    expect(errors).toContain(
      'exports: vueExclusive "NoSuchExport" is missing in Vue (contract stale).'
    );
  });

  test('a React component with a props type needs an entry or a propsUnchecked reason', () => {
    const input = inputs();
    input.reactApi +=
      '\nexport function NewPart(input: NewPartProps): react.JSX.Element;\n' +
      'export interface NewPartProps {\n    readonly className?: string;\n}\n';
    expect(parityErrors(input)).toContain(
      'components: React component "NewPart" has no components entry and no propsUnchecked reason.'
    );
    input.contract.propsUnchecked.NewPart = 'Probe-only part.';
    expect(parityErrors(input).some((e) => e.includes('component "NewPart"'))).toBe(false);
  });

  test('the top-level XlsxEditor takes callbacks as emits, never as props', () => {
    const input = inputs();
    const editor = shape.XlsxEditor;
    input.vue = { ...shape, XlsxEditor: { ...editor, props: [...editor.props, 'onSave'] } };
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): "onSave" must be a Vue emit, not a prop.'
    );
  });

  test('the top-level XlsxEditor takes className as a fallthrough class, never as a prop', () => {
    const input = inputs();
    const editor = shape.XlsxEditor;
    input.vue = { ...shape, XlsxEditor: { ...editor, props: [...editor.props, 'className'] } };
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): "className" must fall through as a Vue class, not be a prop.'
    );
  });

  test('a new React prop with no contract entry fails', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /export interface XlsxEditorToolbarProps \{/,
      'export interface XlsxEditorToolbarProps {\n    readonly zoom?: number;'
    );
    expect(parityErrors(input)).toContain(
      'XlsxEditorToolbar (XlsxEditorToolbarProps): React prop "zoom" is not classified in the contract.'
    );
  });

  test('a renamed export missing in Vue fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(/^export const xlsxEditorKey\b.*$/m, '');
    expect(parityErrors(input)).toContain(
      'exports: Vue has no "xlsxEditorKey" (pair of "XlsxEditorContext").'
    );
  });

  test('a class that does not reach an element fails', () => {
    const input = inputs();
    input.vue = { ...shape, XlsxEditor: { ...shape.XlsxEditor, classLands: false } };
    expect(parityErrors(input)).toContain(
      'XlsxEditor (XlsxEditorProps): a class on Vue XlsxEditor does not reach a rendered element.'
    );
  });

  test('a deferred prop that ships in Vue must move to paired', () => {
    const input = inputs();
    input.reactApi = input.reactApi.replace(
      /export interface XlsxEditorToolbarProps \{/,
      'export interface XlsxEditorToolbarProps {\n    readonly zoom?: number;'
    );
    input.contract.components.XlsxEditorToolbar.deferredInVue = { zoom: 'Not built in Vue yet.' };
    const toolbar = shape.XlsxEditorToolbar;
    input.vue = { ...shape, XlsxEditorToolbar: { ...toolbar, props: [...toolbar.props, 'zoom'] } };
    expect(parityErrors(input)).toContain(
      'XlsxEditorToolbar (XlsxEditorToolbarProps): deferredInVue "zoom" has shipped in Vue; move it to paired.'
    );
  });

  test('a React callback without a Vue emit fails', () => {
    const input = inputs();
    input.vue = { ...shape, XlsxEditorRoot: { ...shape.XlsxEditorRoot, emits: ['ready'] } };
    expect(parityErrors(input)).toContain(
      'XlsxEditorRoot (XlsxEditorRootProps): Vue XlsxEditorRoot does not emit "change".'
    );
  });

  test('a React-only export that appears in Vue must be reclassified', () => {
    const input = inputs();
    input.contract.exports.reactExclusive.reactOnlyHelper = 'A React helper.';
    input.reactApi += '\nexport function reactOnlyHelper(): void;\n';
    input.vueApi += '\nexport function reactOnlyHelper(): void;\n';
    expect(parityErrors(input)).toContain(
      'exports: reactExclusive "reactOnlyHelper" appeared in Vue; reclassify it.'
    );
  });

  test('a pair with a different release tag in the two adapters fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(
      /\/\/ @public\r?\nexport function translatorFor\b/,
      '// @internal\nexport function translatorFor'
    );
    expect(parityErrors(input)).toContain(
      'exports: "translatorFor" is @public in React but "translatorFor" is @internal in Vue.'
    );
  });
});

describe('check:parity-contract ref gate', () => {
  const REF = [
    '    focus(): void;',
    '    getEditor(): XlsxEditorInstance | null;',
    '    load(bytes: Uint8Array | ArrayBuffer): Promise<EditorSnapshot | null>;',
    '    save(): Uint8Array<ArrayBuffer> | null;',
    '    saveInTasks(): Promise<SavedWorkbook | null>;',
    '    snapshot(): EditorSnapshot;',
  ];
  // API Extractor writes CRLF line endings.
  const HEADER = /\r?\nexport interface XlsxEditorRef \{[\s\S]*?\r?\n\}\r?\n/;
  /** The snapshot with its `XlsxEditorRef` block replaced by one with `members`. */
  const withRef = (api: string, members: readonly string[]): string =>
    `${api.replace(HEADER, '\n')}\n// @public\nexport interface XlsxEditorRef {\n${members.join('\n')}\n}\n`;
  const refErrors = (react: readonly string[], vue: readonly string[]): string[] => {
    const input = inputs();
    input.reactApi = withRef(input.reactApi, react);
    input.vueApi = withRef(input.vueApi, vue);
    return parityErrors(input).filter((e) => e.startsWith('ref:'));
  };

  test('the same members on both refs pass', () => {
    expect(refErrors(REF, REF)).toEqual([]);
  });

  test('a member on one ref only fails', () => {
    expect(refErrors([...REF, '    print(): void;'], REF)).toContain(
      'ref: React XlsxEditorRef member "print" is not classified in the contract.'
    );
    expect(refErrors(REF, REF.slice(1))).toContain('ref: paired "focus" is missing in Vue.');
  });

  test('a paired member with another type fails', () => {
    const vue = REF.map((m) => (m.includes('save()') ? '    save(): Uint8Array;' : m));
    expect(refErrors(REF, vue)).toContain(
      'ref: paired "save" type drift: React `() => Uint8Array<ArrayBuffer> | null`, Vue `() => Uint8Array`.'
    );
  });

  test('a snapshot without the interface fails', () => {
    const input = inputs();
    input.vueApi = input.vueApi.replace(HEADER, '\n');
    expect(parityErrors(input)).toContain('ref: Vue snapshot has no XlsxEditorRef interface.');
  });
});
