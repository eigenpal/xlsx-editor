// After pptx-editor/packages/vue/test/parity-shape.ts at d4502552 (EigenPal). Differences: the
// probes mount the XLSX Editor parts (the editor, its root, viewport, content, toolbar, and the
// toolbar's slot, group, action, and separator parts), and the value probes check the paired props that show in a
// part alone (`label` of the viewport, `hidden` and `preset` of the toolbar, and the props of the
// slot and group parts that render without a bar). The
// probe helpers and the shape that `check:parity-contract` reads are the same.
//
// The Vue adapter's runtime shape for the parity check: each paired component's declared props
// and emits (the runtime truth; the public prop types derive from them), and mount probes that
// prove a host `class` and the `className` prop land on an element, every mapped slot renders,
// and that a paired prop given a value shows in the part. Static reads cannot prove these: a
// root that renders only its slot drops `class` with a Vue warning.
//
// Registers happy-dom, so it runs in its own process (`bun run check:parity-contract`) or in
// the DOM test run (`bun run test:dom`), never in the engine test run.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

export interface VueComponentShape {
  readonly props: readonly string[];
  readonly emits: readonly string[];
  /** A host `class` on the component lands on a rendered element. */
  readonly classLands: boolean;
  /** A `className` prop value lands on a rendered element (false when there is no such prop). */
  readonly classNameLands: boolean;
  /** The slots whose content renders: `default` and each named slot the contract maps. */
  readonly slots: readonly string[];
  /** Value probes: a paired prop, given a value, shows in the rendered part (name → passed). */
  readonly values: Readonly<Record<string, boolean>>;
}

/** One value probe: mount the part with `props` and check that `shows` finds their effect. */
interface ValueProbe {
  readonly props: Record<string, unknown>;
  readonly shows: (scope: ParentNode) => boolean;
}

export type VueAdapterShape = Readonly<Record<string, VueComponentShape>>;

const PROBE_CLASS = 'xlsx-parity-probe';
const PROBE_CLASS_NAME = 'xlsx-parity-class-name';
const slotId = (slot: string): string => `xlsx-parity-slot-${slot}`;

/** The part of `scripts/parity/parity.contract.json` the probes read. */
interface ParityContract {
  readonly components: Record<
    string,
    { readonly reactRenderPropsAsVueSlots?: Record<string, string> } | string
  >;
}

interface ComponentDefinition {
  readonly props?: Record<string, unknown> | readonly string[];
  readonly emits?: Record<string, unknown> | readonly string[];
}

const keys = (value: ComponentDefinition['props']): string[] =>
  Array.isArray(value) ? [...value] : Object.keys(value ?? {});

/** The slots a probe fills: `default`, and every Vue slot the contract maps for the part. */
function probedSlots(contract: ParityContract | undefined, name: string): string[] {
  const rules = contract?.components[name];
  const mapped =
    typeof rules === 'object' ? Object.values(rules.reactRenderPropsAsVueSlots ?? {}) : [];
  return [...new Set(['default', ...mapped])];
}

/**
 * Mount the paired Vue components and read their shape. With the contract, each probe also
 * fills every named slot the contract maps for the part.
 */
export async function vueAdapterShape(contract?: ParityContract): Promise<VueAdapterShape> {
  if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
  const { createApp, h, nextTick } = await import('vue');
  const vue = await import('../src/index.ts');

  type Wrap = (node: ReturnType<typeof h>) => ReturnType<typeof h>;
  const inRoot: Wrap = (node) => h(vue.XlsxEditorRoot, null, () => [node]);
  const inViewport: Wrap = (node) => inRoot(h(vue.XlsxEditorViewport, null, () => [node]));

  const probes: Record<
    string,
    {
      component: unknown;
      props?: Record<string, unknown>;
      /** Renders the probed vnode inside another part (a toolbar inside a root). */
      wrap?: Wrap;
      /** Paired props that must show in the rendered part, by the React prop name. */
      values?: Record<string, ValueProbe>;
      /** The slots that each value probe fills (a part that renders nothing when empty). */
      valueSlots?: readonly string[];
    }
  > = {
    XlsxEditor: { component: vue.XlsxEditor },
    // The root renders no element of its own: only its default slot.
    XlsxEditorRoot: { component: vue.XlsxEditorRoot },
    XlsxEditorViewport: {
      component: vue.XlsxEditorViewport,
      wrap: inRoot,
      values: {
        label: {
          props: { label: 'Probe workbook' },
          shows: (s) => s.querySelector('[role="region"][aria-label="Probe workbook"]') !== null,
        },
      },
    },
    XlsxEditorContent: { component: vue.XlsxEditorContent, wrap: inViewport },
    XlsxEditorToolbar: {
      component: vue.XlsxEditorToolbar,
      wrap: inViewport,
      values: {
        hidden: {
          props: { hidden: ['font.bold'] },
          shows: (s) =>
            s.querySelector('[data-slot="font.italic"]') !== null &&
            s.querySelector('[data-slot="font.bold"]') === null,
        },
        preset: {
          props: { preset: false },
          shows: (s) =>
            s.querySelector('[role="toolbar"]') !== null &&
            s.querySelector('[data-group]') === null,
        },
      },
    },
    // A slot alone renders its built-in control, so the control needs the editor's context.
    XlsxEditorToolbarSlot: {
      component: vue.XlsxEditorToolbarSlot,
      props: { slotId: 'font.bold' },
      wrap: inViewport,
      values: {
        slotId: {
          props: { slotId: 'font.italic' },
          shows: (s) => s.querySelector('[data-slot="font.italic"]') !== null,
        },
        hidden: {
          props: { hidden: true },
          shows: (s) => s.querySelector('[data-slot="font.bold"]') === null,
        },
        icon: {
          props: { icon: h('b', { id: 'xlsx-parity-icon' }) },
          shows: (s) => s.querySelector('[data-slot="font.bold"] #xlsx-parity-icon') !== null,
        },
      },
    },
    XlsxEditorToolbarGroup: {
      component: vue.XlsxEditorToolbarGroup,
      props: { id: 'probe' },
      wrap: inViewport,
      // A group with no content renders nothing, so each value probe gives it a child.
      valueSlots: ['default'],
      values: {
        id: {
          props: { id: 'probe-id' },
          shows: (s) => s.querySelector('[data-host-group="probe-id"]') !== null,
        },
        label: {
          props: { label: 'Probe group' },
          shows: (s) => s.querySelector('[role="group"][aria-label="Probe group"]') !== null,
        },
        labelKey: {
          props: { labelKey: 'formattingBar.groups.file' },
          shows: (s) => s.querySelector('[role="group"][aria-label="File"]') !== null,
        },
        hidden: {
          props: { hidden: true },
          shows: (s) => s.querySelector('[data-host-group]') === null,
        },
      },
    },
    ToolbarAction: {
      component: vue.ToolbarAction,
      props: { label: 'Probe action' },
      wrap: inViewport,
      values: {
        label: {
          props: { label: 'Probe label' },
          shows: (s) => s.querySelector('button[aria-label="Probe label"]') !== null,
        },
        icon: {
          props: { icon: h('b', { id: 'xlsx-parity-icon' }) },
          shows: (s) => s.querySelector('button #xlsx-parity-icon') !== null,
        },
        active: {
          props: { active: true },
          shows: (s) => s.querySelector('button[aria-pressed="true"][data-active]') !== null,
        },
        disabled: {
          props: { disabled: true },
          shows: (s) => s.querySelector('button[aria-disabled="true"]') !== null,
        },
        disabledReason: {
          props: { disabled: true, disabledReason: 'Probe reason' },
          shows: (s) => s.querySelector('button[title="Probe reason"]') !== null,
        },
      },
    },
    ToolbarSeparator: { component: vue.ToolbarSeparator, wrap: inViewport },
  };

  type Probe = (typeof probes)[string];
  /** Mount one probe with the given slots filled, and read what renders. */
  const mount = async (
    probe: Probe,
    filled: readonly string[],
    value?: ValueProbe
  ): Promise<{
    classLands: boolean;
    classNameLands: boolean;
    slots: string[];
    shows: boolean;
  }> => {
    const props = keys((probe.component as ComponentDefinition).props);
    const host = document.createElement('div');
    document.body.append(host);
    const warn = console.warn;
    // Vue warns about the dropped `class` this probe is looking for; the probe reports it.
    console.warn = () => undefined;
    const slotContent = (slot: string) => () => h('i', { id: slotId(slot) });
    const node = () =>
      h(
        probe.component as Parameters<typeof h>[0],
        {
          class: PROBE_CLASS,
          ...(props.includes('className') ? { className: PROBE_CLASS_NAME } : {}),
          ...probe.props,
          ...value?.props,
        },
        Object.fromEntries(filled.map((slot) => [slot, slotContent(slot)]))
      );
    const app = createApp({ render: () => (probe.wrap ? probe.wrap(node()) : node()) });
    try {
      app.mount(host);
      await nextTick();
      return {
        classLands: host.querySelector(`.${PROBE_CLASS}`) !== null,
        classNameLands: host.querySelector(`.${PROBE_CLASS_NAME}`) !== null,
        slots: filled.filter((slot) => host.querySelector(`#${slotId(slot)}`) !== null),
        shows: value ? value.shows(host) : false,
      };
    } finally {
      app.unmount();
      host.remove();
      console.warn = warn;
    }
  };

  const out: Record<string, VueComponentShape> = {};
  for (const [name, probe] of Object.entries(probes)) {
    const definition = probe.component as ComponentDefinition;
    const slotNames = probedSlots(contract, name);
    // Every slot filled first. A part can render its class only without host content, so a
    // missing class gets a mount of its own with no slot.
    const all = await mount(probe, slotNames);
    let { classLands, classNameLands } = all;
    if (!classLands || !classNameLands) {
      const bare = await mount(probe, []);
      classLands ||= bare.classLands;
      classNameLands ||= bare.classNameLands;
    }
    const slots = new Set(all.slots);
    for (const slot of slotNames)
      if (!slots.has(slot) && (await mount(probe, [slot])).slots.includes(slot)) slots.add(slot);
    const values: Record<string, boolean> = {};
    for (const [prop, value] of Object.entries(probe.values ?? {}))
      values[prop] = (await mount(probe, probe.valueSlots ?? [], value)).shows;
    out[name] = {
      props: keys(definition.props),
      emits: keys(definition.emits),
      classLands,
      classNameLands,
      slots: slotNames.filter((slot) => slots.has(slot)),
      values,
    };
  }
  return out;
}
