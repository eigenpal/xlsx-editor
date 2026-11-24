<script setup lang="ts">
// The Vue demo: the Vue twin of the React demo's page (examples/vite/src/App.tsx), in a smaller
// form. It opens the sample workbook of the React demo at start, or a workbook that you pick,
// and downloads the saved copy. The shell is the React demo's (examples/vite/src/demo.css): the
// top bar with the brand lockup, the React | Vue switcher, the name of the open workbook with
// its muted extension, the editor's menu bar, Show all headings, View only, the theme toggle, Open a file, and Save;
// then the toolbar, and the formula bar, grid, and footer in one card. The demo's own text is
// English and stays out of `packages/i18n` (AGENTS.md, "i18n").
//
// The open, headings, and save logic is in `session.ts`, as the React demo does it: an open
// checks the file before it asks about unsaved edits, a newer pick stops an older open, the
// headings option is on at start, and Save puts the file's own heading flags back.
//
// It leaves out these parts of the React demo: the file save picker (Save always downloads),
// the status of a long save, the F6 stop in the top bar, the refusal of keys during an open,
// the sheet zooms of a reopen, and the notices after a save or a headings switch.
import {
  computed,
  onBeforeUnmount,
  onMounted,
  reactive,
  shallowRef,
  useTemplateRef,
  watch,
} from 'vue';
import {
  FormulaBar,
  SheetTabs,
  StatusMessage,
  XlsxEditorContent,
  XlsxEditorMenu,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  ZoomControl,
  type EditorSnapshot,
  type SavedWorkbook,
  type XlsxEditorRef,
} from '@xlsx-editor.dev/vue';
import { SHOW_ADAPTER_SWITCHER } from '../../vite/src/adapters.ts';
import { downloadBlob } from '../../vite/src/download.ts';
import { chromeLocale } from '../../vite/src/chrome-locale.ts';
import { initialTheme, rememberTheme, type Theme } from '../../vite/src/theme.ts';
import { shownName } from '../../vite/src/workbook-names.ts';
import { SAMPLE_URL } from '../../vite/src/workbooks.ts';
import AdapterSwitcher from './AdapterSwitcher.vue';
import BrandLogo from './BrandLogo.vue';
import Icon from './Icon.vue';
import ThemeToggle from './ThemeToggle.vue';
import { DemoSession, initialSessionState } from './session.ts';

const editor = useTemplateRef<XlsxEditorRef>('editor');
// The file input of Open a file, which File > Open of the menu bar also opens.
const openInput = useTemplateRef<HTMLInputElement>('openInput');
// The status of the editor, so the grid shows idle until a workbook is ready.
const status = shallowRef<EditorSnapshot['status']>('empty');
const viewOnly = shallowRef(false);
// The shell's light or dark theme, as in the React demo: `?theme=`, else the theme picked
// before, else light. A pick is kept for the next visit.
const theme = shallowRef<Theme>(initialTheme(window.location.search));
// The chrome catalog from `?lang=`, such as `?lang=ar` for a right-to-left chrome. Read once.
const locale = chromeLocale(window.location.search);
watch(theme, rememberTheme);
let attached = (): void => {};
const ready = new Promise<void>((resolve) => (attached = resolve));

const session = new DemoSession(
  {
    editor: () => editor.value,
    ready,
    confirm: (text) => window.confirm(text),
    download: downloadBlob,
    // After its dialog, a file input keeps the focus or leaves none. So the grid takes the keys.
    focusGrid: () => {
      const active = document.activeElement;
      const fileInput = active instanceof HTMLInputElement && active.type === 'file';
      if (active === document.body || fileInput) editor.value?.focus();
    },
  },
  reactive(initialSessionState())
);
const state = session.state;
// The name of the open workbook for the top bar, or null before one is ready.
const shown = computed(() => (state.label === null ? null : shownName(state.label)));

function pick(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  // Empty the field, so a second pick of the same file opens it again.
  input.value = '';
  if (file) void session.openFile(file);
}

/** Switches the headings option. The box shows the option in force after the switch. */
async function onHeadings(event: Event) {
  const box = event.target as HTMLInputElement;
  await session.setHeadings(box.checked);
  box.checked = state.headings;
}

/** Closes the message and gives the focus back to the grid. */
function closeMessage() {
  session.clearMessage();
  editor.value?.focus();
}

const toolbarSave = (bytes: Uint8Array<ArrayBuffer>, result: SavedWorkbook) =>
  void session.save(bytes, result);
const toolbarSaveStart = (started: Promise<SavedWorkbook | null>) =>
  session.onToolbarSaveStart(started);

/** A held Enter repeats the click, and each click would build the file again. */
function onSaveKey(e: KeyboardEvent) {
  if (e.key === 'Enter' && e.repeat) e.preventDefault();
}

// Leaving the page or reloading it over unsaved edits asks first.
const warn = (event: BeforeUnloadEvent) => {
  if (session.hasUnsavedEdits()) event.preventDefault();
};
onMounted(() => window.addEventListener('beforeunload', warn));
onBeforeUnmount(() => window.removeEventListener('beforeunload', warn));

void session.openSample(async () => {
  const response = await fetch(SAMPLE_URL);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return response.arrayBuffer();
});
</script>

<template>
  <XlsxEditorRoot
    ref="editor"
    :mode="viewOnly ? 'view' : 'edit'"
    :color-mode="theme"
    :locale="locale"
    @ready="attached"
    @change="(s: EditorSnapshot) => (status = s.status)"
  >
    <!-- The page has its own landmarks, such as the top bar's banner, so the viewport renders
         no region of its own. -->
    <XlsxEditorViewport class="demo-root" :label="false">
      <header class="demo-topbar">
        <BrandLogo />
        <AdapterSwitcher v-if="SHOW_ADAPTER_SWITCHER" current="vue" :theme="theme" />
        <span class="demo-topbar__divider" />
        <div class="demo-file">
          <template v-if="shown">
            <span class="demo-file__name" :title="shown.name + shown.ext">{{ shown.name }}</span>
            <span v-if="shown.ext" class="demo-file__ext">{{ shown.ext }}</span>
          </template>
        </div>
        <!-- The menu bar after the name, as in the React demo: File > Open runs the demo's own
             Open, and File > Save its save. -->
        <XlsxEditorMenu
          class-name="demo-menubar"
          :opening="state.opening || state.switching"
          :on-open="() => openInput?.click()"
          :on-save="toolbarSave"
          :on-save-start="toolbarSaveStart"
        />
        <div class="demo-topbar__actions">
          <span class="demo-busy" role="status">{{
            state.switching
              ? 'Switching the headings.'
              : state.opening
                ? 'Opening the workbook.'
                : ''
          }}</span>
          <label class="demo-pill" title="Show the row and column headings of every sheet">
            <input
              type="checkbox"
              class="demo-sr-only"
              aria-label="Show all headings"
              :checked="state.headings"
              @change="onHeadings"
            />
            <Icon name="table_chart" />
            <span class="demo-pill__label">Show all headings</span>
          </label>
          <label class="demo-pill" title="Refuse every edit">
            <input v-model="viewOnly" type="checkbox" class="demo-sr-only" aria-label="View only" />
            <Icon name="visibility" />
            <span class="demo-pill__label">View only</span>
          </label>
          <ThemeToggle v-model="theme" />
          <span class="demo-topbar__divider" />
          <!-- A styled button for the file input. The input keeps its name when a narrow window
               hides the label text. -->
          <label class="demo-btn demo-btn--quiet demo-btn--sm" title="Open a file">
            <Icon name="folder_open" />
            <span class="demo-btn__label">Open a file</span>
            <input
              ref="openInput"
              class="demo-sr-only"
              type="file"
              aria-label="Open a file"
              accept=".xlsx,.xltx"
              @change="pick"
            />
          </label>
          <button
            type="button"
            class="demo-btn demo-btn--primary demo-btn--sm"
            title="Save"
            @mousedown.prevent
            @keydown="onSaveKey"
            @click="session.saveNow()"
          >
            <Icon name="file_download" />
            <span class="demo-btn__label">Save</span>
          </button>
        </div>
      </header>
      <main class="demo-main">
        <h1 class="demo-sr-only">XLSX Editor</h1>
        <div class="demo-dock">
          <XlsxEditorToolbar
            :opening="state.opening || state.switching"
            :on-save="toolbarSave"
            :on-save-start="toolbarSaveStart"
          />
        </div>
        <div class="demo-workspace">
          <div class="demo-canvas">
            <FormulaBar />
            <StatusMessage />
            <XlsxEditorContent :class-name="status === 'ready' ? undefined : 'xlsx-canvas-idle'" />
            <div class="xlsx-footer">
              <SheetTabs />
              <ZoomControl />
            </div>
          </div>
          <div v-if="state.error" class="demo-message demo-message--error">
            <span class="demo-message__text" role="alert">{{ state.error }}</span>
            <button
              type="button"
              class="demo-icon-btn demo-close"
              aria-label="Close the message"
              @mousedown.prevent
              @click="closeMessage"
            >
              <Icon name="close" />
            </button>
          </div>
        </div>
      </main>
    </XlsxEditorViewport>
  </XlsxEditorRoot>
</template>
