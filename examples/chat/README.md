# XLSX Editor chat panel example

Open a workbook from a chat message in a panel beside the conversation. The assistant is scripted, and the workbook stays in your browser.

## Run the example

From the repository root, run:

```bash
bun install
bun run dev:chat
```

Open `http://localhost:5182`.

## Try it

1. Select **Open** on the workbook card. The editor slides in beside the chat, at half the window.
2. Type `set B5 to 1200`, `go to Formats`, `select B5:E8`, `bold`, `add a sheet`, or `undo`. The editor does each action, and the card in the chat shows the change.
3. Select a cell in the panel and type. Press Escape in a cell edit to cancel it. Press Escape in the chat to close the panel.
4. Type `view only`, then `set B5 to 9`. The editor refuses the edit, and the chat shows the editor's reason.
5. Select **Download .xlsx**, or press Control+S (Command+S on macOS) in the grid.
6. Drag the divider, or select **Expand**. On a narrow window, the panel covers the window, and **Back to chat** closes it.

## How it works

One `XlsxEditor.Root` wraps the chat and the panel, so the chat reads and drives the workbook through `useXlsxEditor()`. The panel is `XlsxEditor.Viewport` with the toolbar, the formula bar, the status message, the grid, and the footer. Its header buttons are `XlsxEditor.Toolbar.Action` parts, with icons drawn by `Icon`. The card in the chat shows the sheet with `XlsxEditor.SheetPreview`.

| Task in the chat | Call |
| --- | --- |
| Show a sheet | `editor.setActiveSheet(index)` |
| Select a range | `editor.selectReference('B5:E8')` |
| Write a cell | `editor.selectReference('B5')`, then `editor.enterValue('1200')` |
| Toggle bold | `editor.toggleFormat('bold')` |
| Say why an edit was refused | `refusalText(result.refusal, t, snapshot.numberLocale)` |
| View only | the root's `mode` prop |
| Download | `editor.saveInTasks()` |
| Show the sheet on the card | `<XlsxEditor.SheetPreview className="card__grid" />` |
| Show the first sheet after each open | `useEditorEvent('change', …)` with `source === 'load'` |

The chat's colors and shapes reach the editor chrome through the `--doc-*` tokens on `.xlsx-editor` (`src/styles.css`, "EDITOR THEME"). The workbook canvas keeps the workbook's own colors.

`scripts/verify-chat.ts` (`bun run test:e2e:chat`) runs the happy path in Chromium.
