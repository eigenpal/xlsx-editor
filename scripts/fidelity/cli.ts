// The entry of the fidelity commands: `bun run fidelity`, `bun run fidelity:check-gaps`, and
// `bun run fidelity:record-reference` run `cli.ts <command> [arguments...]`.
//
// The entry loads the module of its command with a dynamic import. So a module that does not
// load, such as a `thresholds.ts` that `--update-gaps` or an edit left broken, prints one
// `fidelity:` line and exits with code 2, never a stack trace or exit code 1. Exit code 1 keeps
// its meaning: a failed gate, or a looser list. The entry imports nothing else, so it loads also
// when every other module is broken.
//
// No PPTX Editor file has this role: its commands import their modules statically.

/** The module of each command, next to this file. */
const COMMANDS: ReadonlyMap<string, string> = new Map([
  ['run', './run.ts'],
  ['check-gaps', './check-known-gaps.ts'],
  ['record-reference', './record-reference.ts'],
]);

/** A command module: `command` runs it and gives its exit code. */
interface CommandModule {
  command(): Promise<number>;
}

/** The first line of the message of `error`. */
function firstLine(error: unknown): string {
  const text = error instanceof Error ? error.message : String(error);
  return text.split('\n')[0] ?? '';
}

async function main(): Promise<number> {
  const name = process.argv[2];
  const commands = [...COMMANDS.keys()].join(', ');
  if (name === undefined) {
    console.error(`fidelity: missing command: use ${commands}`);
    return 2;
  }
  const path = COMMANDS.get(name);
  if (!path) {
    console.error(`fidelity: unknown command "${name}": use ${commands}`);
    return 2;
  }
  // The command reads its own arguments from `process.argv`.
  process.argv.splice(2, 1);
  let module: CommandModule;
  try {
    module = (await import(path)) as CommandModule;
  } catch (error) {
    console.error(`fidelity: ${path.slice(2)} does not load: ${firstLine(error)}`);
    return 2;
  }
  return module.command();
}

// The entry is a module (top-level await), with no import that can fail.
export {};

process.exit(await main());
