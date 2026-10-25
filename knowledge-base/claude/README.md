# Claude Code on this machine

Personal reference for using Claude Code (CLI, v2.1.x) on macOS: where history lives, how to clean it, MCP / skills, and CLAUDE.md.

Last verified: 2026-10-24 against `claude --help` and https://code.claude.com/docs.

| Topic | Jump to |
| --- | --- |
| Where conversations are stored, resume / rewind | [1](#1-conversation-history) |
| What every folder in `~/.claude` is, and what is safe to delete | [2](#2-the-claude-folder-what-is-safe-to-delete) |
| Freeing disk space | [3](#3-cleaning-up-stale-history) |
| MCP and skills | [4](#4-mcp-and-skills) |
| Two features worth using daily | [5](#5-two-features-worth-using-daily) |
| CLAUDE.md and a template | [6](#6-claudemd) |
| What memory is, and how to use `/memory` | [7](#7-memory-and-memory) |

---

## 1. Conversation history

### Where it lives

Every session is one JSONL file:

```
~/.claude/projects/<encoded-working-dir>/<session-id>.jsonl 
```

`<encoded-working-dir>` is the directory you launched `claude` from, with every non-alphanumeric character replaced by `-`. Example from this machine:

```
/Users/yu-shenglee/Desktop/dev/dev-toolbox
  -> ~/.claude/projects/-Users-yu-shenglee-Desktop-dev-dev-toolbox/
```

Launching `claude` from a different folder creates a different project folder, so history is per-directory. Set `CLAUDE_CONFIG_DIR` to move the whole `~/.claude` tree elsewhere.

> Transcripts contain full prompts, file contents and command output, including any secrets that appeared on screen. Treat the folder as sensitive and never commit or share it.

### Managing it from inside Claude

| Command | What it does |
| --- | --- |
| `claude -c` / `--continue` | Resume the most recent session in the current directory |
| `claude -r` / `--resume [id or name]` | Pick a past session (or resume a specific one) |
| `claude -n "name"` | Name a session so it is easy to find later |
| `claude --resume <id> --fork-session` | Resume into a new session ID, leaving the original untouched |
| `claude --from-pr <n>` | Resume the session linked to a pull request |
| `/resume` | Same picker, from inside a running session |
| `/rewind` | Roll code and/or conversation back to an earlier point (uses `file-history/`) |
| `/compact` | Summarize the conversation to free context (same session continues) |
| `/clear` | Start a fresh conversation (old one stays on disk and can be resumed) |
| `claude --no-session-persistence` | (with `-p`) do not write a transcript at all |

---

## 2. The `~/.claude` folder: what is safe to delete

Quit all Claude sessions before deleting anything by hand.

Legend: 🟢 safe to delete, 🟡 deletable but you lose something, 🔴 do not touch.

| Path | What it is for | Tier |
| --- | --- | --- |
| `projects/<dir>/*.jsonl` | Conversation transcripts. What `--resume` reads. | 🟡 Deleting a file permanently loses that conversation. Fine for stale ones. |
| `projects/<dir>/memory/` | Auto memory: notes Claude saved about you and the project, loaded into future sessions. **Not history.** | 🔴 Keep unless you want Claude to forget. Do not delete a whole project folder blindly; this lives inside it. |
| `history.jsonl` | Up-arrow prompt recall across projects. | 🟡 Only loses prompt recall. |
| `file-history/<session-id>/` | Snapshots of files before Claude edited them, used by `/rewind`. | 🟢 You lose rewind for those sessions only. |
| `plans/` | Plan files written in plan mode. | 🟢 Old plans are just notes. |
| `session-env/`, `shell-snapshots/` | Per-session environment and shell state for the Bash tool. | 🟢 Regenerated. Only while no session is running. |
| `cache/` | Changelog, model catalog, misc fetched data. | 🟢 Re-downloaded. |
| `downloads/` | Temp area for files fetched at startup. | 🟢 |
| `backups/` | Timestamped copies of `~/.claude.json` taken before config changes. | 🟢 Rotates itself (at most 5 kept). Leave it, or keep the newest one or two. |
| `plugins/` | Installed plugins and marketplace clones (largest item here, ~22 MB). | 🟡 Reinstall needed. Prefer `claude plugin ...` over `rm`. |
| `skills/` | Your user-level skills (`synced/` comes from claude.ai). | 🟡 Your own skills are lost if not in git. |
| `settings.json` | Permissions allowlist, theme, model, hooks. | 🔴 Hand-curated. Back it up, do not delete. |
| `sessions/` (`*.json`, `*.key`) | Registry of running sessions; `.key` files are credentials. | 🔴 |
| `jobs/` | State for background sessions (`claude --bg`). | 🔴 Use `claude rm <id>` instead. |
| `ide/*.lock` | IDE integration locks. | 🔴 Managed automatically. |
| `config.json` | Approved API-key fingerprints. | 🔴 Auth related. |
| `.last-cleanup`, `.last-update-result.json` | Internal bookkeeping for cleanup and auto-update. | 🔴 |
| `~/.claude.json` (outside the folder) | Main app state: login, per-project trust, local-scope MCP servers. | 🔴 The most important file. Never edit or delete by hand. |
| `CLAUDE.md` (if you create one here) | Your personal global instructions. | 🔴 Keep. See section 6. |

Rule of thumb: prefer built-in commands (`claude purge`, `claude plugin`, `claude mcp remove`, `claude rm`, `cleanupPeriodDays`) over `rm`. They know about the cross-references between folders.

---

## 3. Cleaning up stale history

### Automatic: `cleanupPeriodDays`

Claude Code already sweeps old data on its own. The default is **30 days** (minimum 1; `0` is rejected). It removes transcripts, `file-history/`, `plans/`, `debug/`, `paste-cache/` and similar session data. It does not delete auto memory. Change it in any settings file:

```json
{
  "cleanupPeriodDays": 14
}
```

Put this in `~/.claude/settings.json` to apply everywhere.

### Check what is using space

```bash
du -sh ~/.claude/* 2>/dev/null | sort -h | tail
du -sh ~/.claude/projects/* | sort -h            # per project
find ~/.claude/projects -name '*.jsonl' | wc -l  # number of sessions
```

### Delete one whole project's state: `claude purge`

```bash
claude purge /path/to/project --dry-run   # preview first, always
claude purge /path/to/project             # asks for confirmation
claude purge /path/to/project -i          # confirm item by item
claude purge --all --dry-run              # every project (be careful)
```

Purging a path also matches projects nested under it. On this machine, purging `dev-toolbox` would also hit `dev-toolbox/invlog_utils`, so read the dry-run list. It removes that project's transcripts **and auto memory**, plus matching `tasks/`, `debug/`, `file-history/` entries, `history.jsonl` lines and the entry in `~/.claude.json`. It does not touch settings, CLAUDE.md, `backups/` or `shell-snapshots/`. If you want to keep what Claude learned about a project, use the manual recipe below instead.

### Manual: delete only old transcripts

```bash
# 1. preview sessions older than 30 days
find ~/.claude/projects -maxdepth 2 -name '*.jsonl' -mtime +30 -print

# 2. delete them (memory/ folders are untouched)
find ~/.claude/projects -maxdepth 2 -name '*.jsonl' -mtime +30 -delete

# 3. optional: drop snapshots for sessions that no longer exist
for d in ~/.claude/file-history/*/; do
  id=$(basename "$d")
  ls ~/.claude/projects/*/"$id.jsonl" >/dev/null 2>&1 || rm -rf "$d"
done
```

### Never

- Do not `rm -rf ~/.claude`: it wipes settings, skills, plugins and memory in one go.
- Do not delete anything while a session is open.
- Do not delete `sessions/`, `jobs/` or `~/.claude.json` by hand.

---

## 4. MCP and skills

### MCP (Model Context Protocol)

MCP servers give Claude **tools and data from external systems** (a database, GitHub, Google Drive, a browser). Claude calls them as tools, and they run as a separate process or remote service.

```bash
# remote HTTP server
claude mcp add --transport http sentry https://mcp.sentry.dev/mcp

# local stdio server with env vars (everything after -- is the command)
claude mcp add my-server -e API_KEY=xxx -- npx my-mcp-server

claude mcp list            # configured servers and health
claude mcp get <name>      # details
claude mcp remove <name>
```

Scopes (`--scope` / `-s`):

| Scope | Visible to | Stored in |
| --- | --- | --- |
| `local` (default) | You, this project only | `~/.claude.json` |
| `project` | Everyone via git | `.mcp.json` in the repo root (approval prompt on first use) |
| `user` | You, all projects | `~/.claude.json` |

Inside a session `/mcp` shows status, handles OAuth login and lets you reconnect or toggle servers. Connectors enabled on claude.ai (Google Drive, Claude Docs) show up automatically when you are logged in with that account.

Each connected server adds tool descriptions to the context, so only enable what you use.

### Skills

A skill is a **folder of instructions** (a `SKILL.md`, optionally scripts and reference files) that teaches Claude a repeatable workflow. It loads only when relevant, so it is cheap on context.

```
~/.claude/skills/<name>/SKILL.md     # personal, all projects
<repo>/.claude/skills/<name>/SKILL.md  # project, commit to share
```

```markdown
---
name: sql-review
description: Review T-SQL for this repo's conventions. Use when the user asks to review or write SQL Server queries.
---
1. Check for SELECT * and missing schema prefixes.
2. Flag implicit conversions in JOIN / WHERE.
3. Output findings as a table: line, issue, fix.
```

Use it by typing `/sql-review`, or let Claude pick it automatically from the `description`. Useful optional frontmatter: `disable-model-invocation: true` (manual only), `allowed-tools`, `argument-hint`; arguments arrive as `$ARGUMENTS`. Plugins (`claude plugin`) bundle skills, commands and MCP servers for sharing.

### Which one to use

| You want | Use |
| --- | --- |
| Claude to reach an external system or live data | MCP server |
| Claude to follow your procedure or checklist | Skill |
| A fact Claude should always know about the repo | CLAUDE.md |
| Something to happen every time, with no exceptions | Hook (below) |

---

## 5. Two features worth using daily

### a) Background sessions: `claude --bg` and `claude agents`

Start a long task and get your terminal back:

```bash
claude --bg "upgrade dependencies in invlog_utils and run the tests"
claude agents          # dashboard: list, peek, attach, stop
claude attach <id>     # jump into a session
claude logs <id>       # recent output
claude stop <id>       # stop, conversation is kept
claude rm <id>         # delete session (and its worktree when safe)
```

By default background sessions edit an isolated git worktree, so they do not collide with what you are doing. Note: this folder is not a git repo right now, so run `git init` first if you want that isolation. Agent view is a research preview.

### b) Hooks: deterministic automation

CLAUDE.md is advice that Claude may or may not follow. A **hook** is a shell command the harness always runs at a lifecycle point. Example in `~/.claude/settings.json` that formats Python after every edit:

```json
{
  "hooks": {
    "PostToolUse": [
      {
        "matcher": "Edit|Write",
        "hooks": [
          { "type": "command", "command": "jq -r '.tool_input.file_path' | grep '\\.py$' | xargs -r uv run ruff format" }
        ]
      }
    ]
  }
}
```

Related time-savers:

- `/fewer-permission-prompts` scans your past sessions and proposes an allowlist for the `permissions.allow` list in `settings.json` (your current list already has AWS and terraform read commands).
- Plan mode (Shift+Tab) lets Claude research and propose before touching files.
- `/loop` and `/schedule` repeat a prompt on an interval or on a cron.

---

## 6. CLAUDE.md

`CLAUDE.md` is a markdown file Claude reads at the start of every session. It is how you give it standing context: commands, conventions, gotchas.

### Where files go

| Location | Scope | Commit? |
| --- | --- | --- |
| `~/.claude/CLAUDE.md` | You, every project | n/a |
| `./CLAUDE.md` or `./.claude/CLAUDE.md` | Whole team, this repo | Yes |
| `./CLAUDE.local.md` | Just you, this repo | No (gitignore it) |
| `<subdir>/CLAUDE.md` | Loaded on demand when Claude works in that folder | Yes |
| `.claude/rules/*.md` | Topic rules; add `paths:` frontmatter to load only for matching files | Yes |
| `/Library/Application Support/ClaudeCode/CLAUDE.md` | Organization policy (macOS) | n/a |

Files from the working directory and all parent folders load at launch, concatenated from the root down. Import other files with `@path`, e.g. `See @README.md for the layout.` (not parsed inside code blocks).

### Creating and editing

- `/init` generates a starting file from the codebase (and suggests improvements if one exists). Treat the output as a draft and trim it.
- `/memory` opens CLAUDE.md and auto-memory files for editing and toggles auto memory.
- In conversation: say "add this to CLAUDE.md: ..." or "remember that ...". (There is no `#` shortcut.)

### Writing rules

- Keep each file **under ~200 lines**. Move long or path-specific material into `.claude/rules/` or a skill.
- Write what Claude cannot infer from code: exact commands, non-obvious conventions, traps. Skip things it can read from the files.
- Be specific and verifiable: "Run `uv run pytest tools/pdf-merger`" beats "make sure tests pass".
- Put commands first. Use short bullets and headings.
- Do not store secrets or account IDs here. It is committed.
- Review it when behavior drifts; stale instructions are worse than none.

### Recommended template

```markdown
# <Project name>

<One or two sentences: what this repo is and who uses it.>

## Commands
- Install: `uv sync`
- Test: `uv run pytest <path>`
- Lint / format: `uv run ruff check . && uv run ruff format .`
- Run <main tool>: `uv run python -m <module> --help`

## Layout
- `tools/` - uv workspace members and standalone projects
- `sql/` - standalone SQL (T-SQL, SQL Server)
- `knowledge-base/` - notes; markdown only, keep headings flat
- `_archive/` - git-ignored parking spot; never edit

## Conventions
- Python 3.x, type hints on public functions, no new dependencies without asking.
- SQL: T-SQL, schema-qualified names, no `SELECT *`.
- Terraform: run `terraform fmt` and `terraform validate` before proposing a change.

## Boundaries
- Never run `terraform apply`, `terraform destroy` or write operations against AWS without asking.
- AWS: use the profile named in my environment; do not hard-code account IDs or keys.
- Do not edit generated files or `uv.lock` by hand.

## Gotchas
- <Thing that has bitten you before, e.g. Windows .bat launchers need CRLF.>

## Further reading
@README.md
```

And a short personal `~/.claude/CLAUDE.md` for habits that apply everywhere:

```markdown
# Personal preferences
- Explain the why in one or two sentences, then show the change.
- Ask before deleting files or running anything that changes cloud resources.
- Prefer small, reviewable diffs over large refactors.
```

---

## 7. Memory and `/memory`

"Memory" is what Claude carries from one session into the next. A conversation starts blank, so there are two sources of carried-over context:

| | CLAUDE.md | Auto memory |
| --- | --- | --- |
| Written by | You | Claude, as it works |
| Holds | Rules and instructions you want followed | Things Claude learned: your preferences, corrections, project facts |
| Location | `./CLAUDE.md`, `~/.claude/CLAUDE.md`, ... (section 6) | `~/.claude/projects/<encoded-working-dir>/memory/` |
| Shared via git | Yes (project file) | No, local to this machine |
| Scope | Project, user or org | Per working directory |

### How auto memory works

- When you say "remember that ..." or correct Claude on something worth keeping, it saves a short note as a markdown file in the project's `memory/` folder.
- `MEMORY.md` in that folder is an index, one line per note. It is loaded into context at the start of every session in that directory; the individual note files are read when relevant.
- It is per directory: memory saved while working in `dev-toolbox` is not visible from another project. Memory is not conversation history either, so deleting a transcript does not delete it (and `claude purge` does, see section 3).

Real example on this machine (`projects/-Users-yu-shenglee-Desktop-Nutritionell/memory/`):

```
MEMORY.md                           # index
feedback_tf_comment_dont_delete.md  # "comment out obsolete Terraform blocks, don't delete them"
```

### Using `/memory`

Type `/memory` inside a session. It lets you:

1. See which CLAUDE.md files and auto memory files are loaded for this session.
2. Open any of them in your editor to read or fix it.
3. Turn auto memory on or off.

Day-to-day use:

| You want to | Do this |
| --- | --- |
| Save a fact or preference | Say "remember that the API tests need a local Redis" |
| Put a rule in the team file | Say "add this to CLAUDE.md: ..." |
| Check what Claude believes about you | `/memory`, open the memory folder, read the notes |
| Fix a wrong or stale memory | `/memory`, open the note and edit or delete it (also delete its line in `MEMORY.md`) |
| Make Claude forget something | Say "forget that ..." or delete the note file by hand |
| Stop Claude saving notes | `/memory`, toggle auto memory off |

Tips: memories are plain markdown, so you can edit them directly in `~/.claude/projects/<dir>/memory/`. Keep them short, and prune ones that no longer apply, because Claude acts on what it reads there. Never ask it to remember secrets, since the files are unencrypted. Prefer CLAUDE.md for rules the whole team needs, and memory for personal, per-machine preferences.

---

## Cheat sheet

```bash
claude                      # start in current dir
claude -c                   # continue last session
claude -r                   # pick a session to resume
claude --bg "task"          # background session
claude agents               # manage background sessions
claude mcp list             # MCP servers
claude purge <path> --dry-run
claude doctor               # health check
du -sh ~/.claude/* | sort -h
```

Slash commands: `/resume` `/rewind` `/compact` `/clear` `/init` `/memory` `/mcp` `/fewer-permission-prompts`.
