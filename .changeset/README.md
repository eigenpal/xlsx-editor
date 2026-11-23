# Changesets

After the first publish, every code change to a published package needs a changeset. Until then, no pull request has one (`AGENTS.md`, "Releasing"). Run `bun changeset`, or write `.changeset/<name>.md` by hand. Published packages are one fixed group: declare one bump and the others follow. `config.json` lists the group (`fixed`). Add each new publishable package to it.
