// The chat example's own English text, in one place, as the PPTX Editor chat example keeps its
// text in pptx-editor/examples/chat/src/text.ts at e0ac8351 (EigenPal). The editor's controls
// use the published locale catalog. The example's text stays out of `packages/i18n`
// (AGENTS.md, "i18n").

export const TEXT = {
  appName: 'Ledger',
  you: 'You',
  youInitial: 'Y',
  newChat: 'New chat',
  greeting: 'What do you want to work out?',
  greetingLead: 'Describe a table, and Ledger drafts the workbook. You edit the cells right here.',
  composerPlaceholder: 'Message Ledger, or try “go to Sales”',
  composerLabel: 'Message Ledger',
  composerHint: 'Enter to send, Shift+Enter for a new line',
  send: 'Send',
  thinking: 'Building your workbook',
  working: 'Working',
  disclaimer:
    'This assistant is a scripted mock. The workbook and every edit stay in your browser.',
  suggestions: [
    'Make a workbook that shows what the editor can do',
    'Put my quarterly sales into a table',
    'Draft a tracker for a project schedule',
  ],
  seedPrompt: 'Make a workbook that shows what the editor can do',
  seedTitle: 'Editor showcase workbook',
  workbookTitle: 'Showcase.xlsx',
  workbookName: 'Showcase',
  workbook: 'Workbook',
  sheets: (n: number) => `${n} ${n === 1 ? 'sheet' : 'sheets'}`,
  edited: 'Edited',
  viewOnly: 'View only',
  openWorkbook: 'Open',
  openWorkbookLabel: (title: string) => `Open ${title} in the editor`,
  workbookOpen: 'Editing',
  closeWorkbook: 'Close',
  expand: 'Expand',
  collapse: 'Show chat',
  back: 'Back to chat',
  download: 'Download .xlsx',
  lock: 'View only',
  resize: 'Resize the workbook panel',
  notReady: 'The workbook is not open yet.',
  moreSheets: (n: number) => `+${n}`,
  refused: (reason: string) => `I could not do that. ${reason}`,
} as const;
