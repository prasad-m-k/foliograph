# /foliograph Claude Code Skill

A Claude Code slash command that builds or refreshes a Foliograph knowledge
graph for the current project.

## Installation

Copy this file to your Claude Code skills directory:

```bash
cp foliograph.md ~/.claude/commands/foliograph.md
```

Then use `/foliograph` inside any Claude Code session.

## The command

```
/foliograph [build|check|fetch] [args]
```

### /foliograph build

Scans the current directory for supported documents and builds the graph.

```bash
foliograph build . --output . --name "$(basename $PWD)"
```

Expected output:
- `FOLIO_GRAPH.md`
- `FOLIO_INDEX.md`
- `FOLIO_RELATIONS.json`
- `FOLIO_SESSION.md`
- `CLAUDE.md`
- `.claude/settings.json`

### /foliograph check

Checks whether the graph is current relative to the source documents.

```bash
foliograph check --graph FOLIO_GRAPH.md
```

Returns a drift report. If drift is detected, suggests running `/foliograph build`.

### /foliograph fetch

Fetches the full text of a specific section from a source document.

```bash
foliograph fetch "chapter4.docx § The Swarm Model"
```

Prints only that section to stdout, not the whole document.

## Prompt injected by the PreToolUse hook

When `.claude/settings.json` is present, Claude Code automatically reads
the first 80 lines of `FOLIO_GRAPH.md` before every tool call, giving the
model structural context at zero manual cost.

## Supported file types

.docx  .pdf  .pptx  .md  .txt

## Full CLI reference

```
foliograph build  <sources...> [-o DIR] [-n NAME] [--no-session]
foliograph check  [--graph FOLIO_GRAPH.md] [-v]
foliograph stats  <FOLIO_GRAPH.md>
foliograph fetch  "<file> § <Section Title>"
```
