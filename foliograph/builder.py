"""
foliograph.builder
~~~~~~~~~~~~~~~~~~
Assemble FOLIO_GRAPH.md, FOLIO_INDEX.md, FOLIO_RELATIONS.json,
and FOLIO_SESSION.md from one or more DocumentRecords.
"""

from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from .extractor import DocumentRecord, Section, extract
from .relationships import build_relationships, render_relationships_block
from .drift import detect_drift, render_drift_block, DriftReport
from .stats_html import generate_stats_html


# ---------------------------------------------------------------------------
# Graph rendering
# ---------------------------------------------------------------------------

_INDENT = "  "


def _render_section_tree(sections: list[Section]) -> list[str]:
    lines: list[str] = []
    for sec in sections:
        if sec.level == 0:
            continue
        indent = _INDENT * max(0, sec.level - 1)
        loc = f" [{sec.page_hint}]" if sec.page_hint else ""
        wc = f" ({sec.word_count}w)" if sec.word_count > 20 else ""
        lines.append(f"{indent}- **{sec.title}**{loc}{wc}")
        if sec.summary:
            sum_indent = _INDENT * sec.level
            lines.append(f"{sum_indent}  > {sec.summary}")
    return lines


def _render_document_block(rec: DocumentRecord) -> list[str]:
    lines: list[str] = []
    fname = rec.path.name
    ftype = rec.file_type.upper()
    words = f"{rec.total_words:,}"
    pages = f" | {rec.total_pages} pages" if rec.total_pages else ""

    lines.append(f"### `{fname}` [{ftype}]")
    lines.append(f"**Title:** {rec.title}  ")
    lines.append(f"**Words:** {words}{pages}  ")
    lines.append("")

    if rec.sections:
        lines.append("**Structure:**")
        lines.extend(_render_section_tree(rec.sections))
        lines.append("")

    if rec.tables:
        lines.append("**Tables:**")
        for t in rec.tables:
            lines.append(f"- {t}")
        lines.append("")

    if rec.figures:
        lines.append("**Figures/Diagrams:**")
        for f in rec.figures:
            lines.append(f"- {f}")
        lines.append("")

    if rec.named_entities:
        terms = ", ".join(rec.named_entities[:30])
        lines.append(f"**Key Terms:** {terms}")
        lines.append("")

    lines.append("---")
    lines.append("")
    return lines


# ---------------------------------------------------------------------------
# Index rendering
# ---------------------------------------------------------------------------

def _build_index(records: list[DocumentRecord]) -> list[tuple[str, str, str]]:
    entries: list[tuple[str, str, str]] = []
    seen: set[str] = set()

    for rec in records:
        fname = rec.path.name

        for sec in rec.sections:
            if sec.level == 0 or not sec.title:
                continue
            key = sec.title.lower()
            if key in seen:
                continue
            seen.add(key)
            loc = sec.page_hint or f"> {sec.title}"
            entries.append((sec.title, fname, loc))

        for ent in rec.named_entities:
            key = ent.lower()
            if key in seen:
                continue
            seen.add(key)
            location = fname
            for sec in rec.sections:
                if ent.lower() in sec.summary.lower() or ent.lower() in sec.title.lower():
                    location = f"{fname} > {sec.title}"
                    break
            entries.append((ent, fname, location))

        for t in rec.tables:
            clean = re.sub(r"\*", "", t).strip()
            if " | " in clean or clean.startswith("#"):
                continue
            key = clean.lower()
            if key not in seen and len(clean) > 8:
                seen.add(key)
                entries.append((clean[:80], fname, fname))

    return sorted(entries, key=lambda x: x[0].lower())


def _render_index(entries: list[tuple[str, str, str]]) -> list[str]:
    lines: list[str] = []
    current_letter = ""
    for concept, fname, location in entries:
        first = concept[0].upper() if concept else "?"
        if first != current_letter:
            current_letter = first
            lines.append(f"\n### {current_letter}\n")
        lines.append(f"- **{concept}** → `{location}`")
    return lines


# ---------------------------------------------------------------------------
# Session starter
# ---------------------------------------------------------------------------

SESSION_STARTER_TEMPLATE = """\
# Foliograph Session Starter

Paste this block at the start of every new LLM session to orient the model
with minimal token cost.

---

```
You have access to a Foliograph knowledge graph for this document set.
Graph files available:

  FOLIO_GRAPH.md      - structural map of every document + relationships
  FOLIO_INDEX.md      - concept → location index
  FOLIO_RELATIONS.json- machine-readable relationship graph

Rules for this session:
1. Read FOLIO_GRAPH.md first for orientation (do not load source files yet).
2. Use FOLIO_INDEX.md to locate any concept before loading a full section.
3. Load sections on demand only: "Load [filename] > [Section Title]"
   Do not load entire source files unless explicitly asked.
4. Never re-read a section you have already processed this session.
5. If FOLIO_GRAPH.md shows a drift WARNING, note it before answering
   questions that touch the flagged section.

Documents in scope:
{file_list}
```
"""


# ---------------------------------------------------------------------------
# Hook injection
# ---------------------------------------------------------------------------

CLAUDE_MD_TEMPLATE = """\
# Foliograph Context (auto-injected)

This project uses Foliograph for token-efficient document navigation.

## On every session start

1. Read `FOLIO_GRAPH.md` for the structural map of all documents.
2. Use `FOLIO_INDEX.md` to locate concepts before loading source files.
3. Load sections on demand: respond to "Load [file] > [Section]" by reading
   only that section from the source document.
4. Run `foliograph check` before any session where you plan to edit source
   documents, to verify the graph is current.

## Documents indexed

{file_list}

## Quick reference

- Rebuild graph:  `foliograph build {source_args} -o {output_dir}`
- Check for drift: `foliograph check --graph {graph_path}`
- Fetch a section: `foliograph fetch "[file] > [Section Title]"`
"""

CLAUDE_CODE_HOOK_TEMPLATE = """\
{
  "hooks": {
    "PreToolUse": [
      {
        "matcher": ".*",
        "hooks": [
          {
            "type": "command",
            "command": "cat FOLIO_GRAPH.md 2>/dev/null | head -80 || echo 'No FOLIO_GRAPH.md found. Run: foliograph build <sources>'"
          }
        ]
      }
    ]
  }
}
"""


def write_hook_files(
    output_dir: Path,
    records: list[DocumentRecord],
    source_args: str,
) -> dict[str, Path]:
    """
    Write CLAUDE.md (Project context) and .claude/settings.json (hook injection).
    """
    file_list = "\n".join(f"  - {r.path.name}" for r in records)
    graph_path = output_dir / "FOLIO_GRAPH.md"

    # CLAUDE.md: used by Claude Code as persistent project context
    claude_md = CLAUDE_MD_TEMPLATE.format(
        file_list=file_list,
        source_args=source_args,
        output_dir=str(output_dir),
        graph_path=str(graph_path),
    )
    claude_md_path = output_dir / "CLAUDE.md"
    claude_md_path.write_text(claude_md, encoding="utf-8")

    # .claude/settings.json: PreToolUse hook
    claude_dir = output_dir / ".claude"
    claude_dir.mkdir(exist_ok=True)
    hook_path = claude_dir / "settings.json"
    hook_path.write_text(CLAUDE_CODE_HOOK_TEMPLATE, encoding="utf-8")

    return {"claude_md": claude_md_path, "hook": hook_path}


# ---------------------------------------------------------------------------
# /foliograph Claude Code skill
# ---------------------------------------------------------------------------

FOLIOGRAPH_SKILL_TEMPLATE = """\
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
foliograph fetch "chapter4.docx > The Swarm Model"
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
foliograph fetch  "<file> > <Section Title>"
```
"""


def write_skill_file(output_dir: Path) -> Path:
    """Write the /foliograph Claude Code skill markdown file."""
    skill_path = output_dir / "foliograph.skill.md"
    skill_path.write_text(FOLIOGRAPH_SKILL_TEMPLATE, encoding="utf-8")
    return skill_path


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build(
    sources: list[Path | str],
    output_dir: Path | str = ".",
    project_name: str = "My Project",
    include_session_starter: bool = True,
    check_drift: bool = True,
    write_hooks: bool = True,
    write_stats: bool = True,
) -> dict[str, Path]:
    """
    Build FOLIO_GRAPH.md, FOLIO_INDEX.md, FOLIO_RELATIONS.json,
    and supporting files from a list of documents or directories.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Resolve sources
    file_paths: list[Path] = []
    from .extractor import SUPPORTED
    for src in sources:
        src = Path(src)
        if src.is_dir():
            for p in sorted(src.rglob("*")):
                if p.suffix.lower() in SUPPORTED and p.is_file():
                    file_paths.append(p)
        elif src.is_file():
            file_paths.append(src)

    if not file_paths:
        raise ValueError("No supported documents found in the provided sources.")

    # Extract all documents
    records: list[DocumentRecord] = []
    errors: list[str] = []
    for fp in file_paths:
        try:
            print(f"  Extracting: {fp.name} ...", end=" ", flush=True)
            rec = extract(fp)
            records.append(rec)
            print(f"OK  ({rec.total_words:,} words)")
        except Exception as exc:
            print(f"SKIP ({exc})")
            errors.append(f"{fp.name}: {exc}")

    if not records:
        raise RuntimeError("All documents failed to extract.")

    # Drift detection (compare against previous graph if it exists)
    drift_report: DriftReport | None = None
    prev_graph = output_dir / "FOLIO_GRAPH.md"
    if check_drift and prev_graph.exists():
        print("  Checking for drift ...", end=" ", flush=True)
        drift_report = detect_drift(records, prev_graph)
        if drift_report.is_clean:
            print("clean")
        else:
            n_warn = len(drift_report.warnings) + len(drift_report.breaking)
            print(f"{n_warn} change(s) detected")

    # Relationship mapping
    print("  Mapping relationships ...", end=" ", flush=True)
    rel_graph = build_relationships(records)
    print(f"{len(rel_graph.relationships)} edge(s)")

    # Build metadata
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    total_words = sum(r.total_words for r in records)
    file_list_str = "\n".join(f"  - {r.path.name}" for r in records)

    # -----------------------------------------------------------------------
    # FOLIO_GRAPH.md
    # -----------------------------------------------------------------------
    graph_lines: list[str] = [
        "# Foliograph Knowledge Graph",
        "",
        f"**Project:** {project_name}  ",
        f"**Built:** {now}  ",
        f"**Documents:** {len(records)}  ",
        f"**Total words:** {total_words:,}  ",
        f"**Relationships:** {len(rel_graph.relationships)}  ",
        "",
        "> Structural map of your document set.",
        "> Use FOLIO_INDEX.md for targeted lookups.",
        "> Use FOLIO_RELATIONS.json for programmatic graph traversal.",
        "> Never load full source files when a section lookup will do.",
        "",
        "---",
        "",
    ]

    # Drift status block (if previous graph existed)
    if drift_report is not None:
        graph_lines.extend(render_drift_block(drift_report))

    graph_lines += ["## Documents", ""]

    for rec in records:
        graph_lines.extend(_render_document_block(rec))

    # Relationships block
    rel_block = render_relationships_block(rel_graph)
    if rel_block:
        graph_lines.extend(rel_block)

    if errors:
        graph_lines += ["## Extraction Errors", ""]
        for e in errors:
            graph_lines.append(f"- {e}")
        graph_lines.append("")

    graph_path = output_dir / "FOLIO_GRAPH.md"
    graph_path.write_text("\n".join(graph_lines), encoding="utf-8")

    # -----------------------------------------------------------------------
    # FOLIO_INDEX.md
    # -----------------------------------------------------------------------
    entries = _build_index(records)
    index_lines: list[str] = [
        "# Foliograph Concept Index",
        "",
        f"**Project:** {project_name}  ",
        f"**Built:** {now}  ",
        f"**Entries:** {len(entries)}  ",
        "",
        "> Locate any concept, term, or section across all documents.",
        '> Format: **Concept** → `file.ext > Section Title`',
        "> To load: ask your LLM 'Load [file] > [Section]'",
        "",
        "---",
        "",
    ]
    index_lines.extend(_render_index(entries))
    index_lines.append("")

    index_path = output_dir / "FOLIO_INDEX.md"
    index_path.write_text("\n".join(index_lines), encoding="utf-8")

    # -----------------------------------------------------------------------
    # FOLIO_RELATIONS.json
    # -----------------------------------------------------------------------
    relations_path = output_dir / "FOLIO_RELATIONS.json"
    relations_path.write_text(rel_graph.to_json(), encoding="utf-8")

    outputs = {"graph": graph_path, "index": index_path, "relations": relations_path}

    # -----------------------------------------------------------------------
    # FOLIO_SESSION.md
    # -----------------------------------------------------------------------
    if include_session_starter:
        starter = SESSION_STARTER_TEMPLATE.format(file_list=file_list_str)
        session_path = output_dir / "FOLIO_SESSION.md"
        session_path.write_text(starter, encoding="utf-8")
        outputs["session"] = session_path

    # -----------------------------------------------------------------------
    # Hook files (CLAUDE.md + .claude/settings.json)
    # -----------------------------------------------------------------------
    if write_hooks:
        source_args = " ".join(str(Path(s)) for s in sources)
        hook_outputs = write_hook_files(output_dir, records, source_args)
        outputs.update(hook_outputs)

        # /foliograph skill file
        skill_path = write_skill_file(output_dir)
        outputs["skill"] = skill_path

    # -----------------------------------------------------------------------
    # FOLIO_STATS.html
    # -----------------------------------------------------------------------
    if write_stats:
        stats_path = output_dir / "FOLIO_STATS.html"
        try:
            generate_stats_html(graph_path, index_path, stats_path)
            outputs["stats"] = stats_path
        except Exception as exc:
            print(f"  Warning: could not generate stats dashboard: {exc}")

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    g_size = graph_path.stat().st_size
    i_size = index_path.stat().st_size
    total_size = g_size + i_size
    print(f"\nFoliograph build complete.")
    print(f"  {graph_path}  ({g_size:,} bytes)")
    print(f"  {index_path}  ({i_size:,} bytes)")
    print(f"  {relations_path}  ({relations_path.stat().st_size:,} bytes)")
    if include_session_starter:
        print(f"  {outputs['session']}")
    if write_hooks:
        print(f"  {outputs['claude_md']}")
        print(f"  {outputs['hook']}")
        print(f"  {outputs['skill']}")
    if "stats" in outputs:
        print(f"  {outputs['stats']}  (open in browser)")
    print(f"\n  Graph tokens (est.): ~{total_size // 4:,}  "
          f"vs full-text tokens (est.): ~{total_words // 1:.0f}")

    return outputs
