"""
foliograph.cli
~~~~~~~~~~~~~~
Command-line interface for Foliograph.

Commands:
  build   - Extract documents and write graph files
  check   - Detect drift between source docs and existing graph
  fetch   - Extract and print a specific section to stdout
  stats   - Show token-cost comparison for an existing graph
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------

def cmd_build(args: argparse.Namespace) -> int:
    from .builder import build

    sources = [Path(s) for s in args.sources]
    output_dir = Path(args.output)
    project_name = args.name or _infer_project_name(sources)

    try:
        build(
            sources=sources,
            output_dir=output_dir,
            project_name=project_name,
            include_session_starter=not args.no_session,
            check_drift=not args.no_drift,
            write_hooks=not args.no_hooks,
        )
    except (ValueError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    """Check whether source documents have drifted from the existing graph."""
    from .extractor import extract, SUPPORTED
    from .drift import detect_drift, render_drift_report

    graph_path = Path(args.graph)
    if not graph_path.exists():
        print(f"Graph not found: {graph_path}", file=sys.stderr)
        print("Run 'foliograph build' first.", file=sys.stderr)
        return 1

    # Resolve source documents
    sources = [Path(s) for s in args.sources] if args.sources else [graph_path.parent]
    file_paths: list[Path] = []
    for src in sources:
        if src.is_dir():
            for p in sorted(src.rglob("*")):
                if p.suffix.lower() in SUPPORTED and p.is_file():
                    # Skip the graph output files themselves
                    if p.name.startswith("FOLIO_"):
                        continue
                    file_paths.append(p)
        elif src.is_file():
            file_paths.append(src)

    if not file_paths:
        print("No supported documents found to check.", file=sys.stderr)
        return 1

    print(f"Checking {len(file_paths)} document(s) against {graph_path.name} ...")
    records = []
    for fp in file_paths:
        try:
            records.append(extract(fp))
        except Exception as exc:
            print(f"  Warning: could not extract {fp.name}: {exc}")

    report = detect_drift(records, graph_path)
    print(render_drift_report(report, verbose=args.verbose))

    # Exit code: 0 = clean, 1 = warnings/breaking
    return 0 if report.is_clean else 1


def cmd_fetch(args: argparse.Namespace) -> int:
    """
    Fetch and print the full text of a specific section from a source document.

    Usage:
      foliograph fetch "chapter4.docx § The Swarm Model"
      foliograph fetch chapter4.docx --section "The Swarm Model"
    """
    from .extractor import extract

    # Parse "file § Section" syntax or separate args
    ref = args.ref
    section_title: str | None = args.section or None

    if "§" in ref:
        parts = ref.split("§", 1)
        file_arg = parts[0].strip()
        section_title = parts[1].strip()
    else:
        file_arg = ref.strip()

    # Resolve file path
    file_path = Path(file_arg)
    if not file_path.exists():
        # Try relative to sources if provided
        if args.sources:
            for src in args.sources:
                candidate = Path(src) / file_arg
                if candidate.exists():
                    file_path = candidate
                    break

    if not file_path.exists():
        print(f"File not found: {file_arg}", file=sys.stderr)
        return 1

    try:
        rec = extract(file_path)
    except Exception as exc:
        print(f"Error extracting {file_path.name}: {exc}", file=sys.stderr)
        return 1

    if not section_title:
        # No section specified: print full document summary
        print(f"# {rec.title}\n")
        print(f"Words: {rec.total_words:,}\n")
        print("Sections:")
        for sec in rec.sections:
            if sec.level > 0:
                indent = "  " * (sec.level - 1)
                print(f"{indent}- {sec.title} ({sec.word_count}w)")
        return 0

    # Find the matching section
    section_title_lower = section_title.lower().strip()
    matched_idx: int | None = None
    for i, sec in enumerate(rec.sections):
        if sec.title.lower().strip() == section_title_lower:
            matched_idx = i
            break
        # Fuzzy: section title contains the search term
        if section_title_lower in sec.title.lower():
            matched_idx = i
            break

    if matched_idx is None:
        print(f"Section not found: '{section_title}' in {file_path.name}", file=sys.stderr)
        print("\nAvailable sections:")
        for sec in rec.sections:
            if sec.level > 0:
                print(f"  - {sec.title}")
        return 1

    # Extract the text between this section heading and the next same-or-higher level heading
    # by scanning the raw text for the section title
    raw = rec.raw_text
    lines = raw.splitlines()
    target_sec = rec.sections[matched_idx]

    # Find start line
    start_line: int | None = None
    for i, line in enumerate(lines):
        if target_sec.title.lower() in line.lower():
            start_line = i
            break

    if start_line is None:
        # Fallback: print summary
        print(f"## {target_sec.title}\n")
        print(target_sec.summary)
        return 0

    # Find end line (next heading at same or higher level)
    end_line = len(lines)
    heading_pattern = re.compile(r"^#{1,6}\s+")
    current_level_pattern = re.compile(r"^#{1," + str(target_sec.level) + r"}\s+")

    for i in range(start_line + 1, len(lines)):
        if heading_pattern.match(lines[i]) and current_level_pattern.match(lines[i]):
            end_line = i
            break

    section_text = "\n".join(lines[start_line:end_line])
    print(section_text)
    return 0


def cmd_stats(args: argparse.Namespace) -> int:
    """Print token-cost comparison for a built graph."""
    graph = Path(args.graph)
    index = Path(args.index) if args.index else graph.parent / "FOLIO_INDEX.md"
    relations = graph.parent / "FOLIO_RELATIONS.json"

    if not graph.exists():
        print(f"Graph file not found: {graph}", file=sys.stderr)
        return 1

    g_size = graph.stat().st_size
    i_size = index.stat().st_size if index.exists() else 0
    r_size = relations.stat().st_size if relations.exists() else 0
    total = g_size + i_size
    est_tokens = total // 4

    print(f"\nFoliograph graph stats")
    print(f"  {graph.name:35s} {g_size:>10,} bytes")
    if index.exists():
        print(f"  {index.name:35s} {i_size:>10,} bytes")
    if relations.exists():
        print(f"  {relations.name:35s} {r_size:>10,} bytes")
    print(f"  {'─' * 47}")
    print(f"  {'Graph + Index total':35s} {total:>10,} bytes")
    print(f"  {'Est. tokens (graph + index)':35s} {est_tokens:>10,}")
    print()
    return 0


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def cmd_stats_html(args: argparse.Namespace) -> int:
    """Generate a self-contained HTML savings dashboard from an existing graph."""
    from .stats_html import generate_stats_html

    graph = Path(args.graph)
    index = Path(args.index) if args.index else graph.parent / "FOLIO_INDEX.md"
    output = Path(args.output) if args.output else graph.parent / "FOLIO_STATS.html"

    if not graph.exists():
        print(f"Graph not found: {graph}", file=sys.stderr)
        print("Run foliograph build first.", file=sys.stderr)
        return 1

    try:
        result = generate_stats_html(graph, index, output)
        print(f"Stats dashboard written to: {result}")
        print(f"Open in any browser. No internet connection required.")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


def _infer_project_name(sources: list[Path]) -> str:
    if len(sources) == 1:
        p = sources[0]
        return (
            p.stem.replace("_", " ").replace("-", " ").title()
            if p.is_file()
            else p.name
        )
    try:
        parents = {p.parent if p.is_file() else p for p in sources}
        if len(parents) == 1:
            return list(parents)[0].name
    except Exception:
        pass
    return "Foliograph Project"


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="foliograph",
        description="Build a compact knowledge graph from office documents.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # -- build ---------------------------------------------------------------
    p_build = sub.add_parser(
        "build",
        help="Extract documents and write FOLIO_GRAPH.md + FOLIO_INDEX.md",
    )
    p_build.add_argument(
        "sources", nargs="+", metavar="SOURCE",
        help="Files (.docx .pdf .pptx .md .txt) or directories to scan",
    )
    p_build.add_argument(
        "-o", "--output", default=".", metavar="DIR",
        help="Output directory (default: current directory)",
    )
    p_build.add_argument(
        "-n", "--name", default="", metavar="NAME",
        help="Project name shown in the graph header",
    )
    p_build.add_argument(
        "--no-session", action="store_true",
        help="Skip writing FOLIO_SESSION.md",
    )
    p_build.add_argument(
        "--no-drift", action="store_true",
        help="Skip drift check against existing graph",
    )
    p_build.add_argument(
        "--no-hooks", action="store_true",
        help="Skip writing CLAUDE.md and .claude/settings.json",
    )
    p_build.set_defaults(func=cmd_build)

    # -- check ---------------------------------------------------------------
    p_check = sub.add_parser(
        "check",
        help="Detect drift between source documents and existing graph",
    )
    p_check.add_argument(
        "--graph", default="FOLIO_GRAPH.md", metavar="FOLIO_GRAPH.md",
        help="Path to existing FOLIO_GRAPH.md (default: ./FOLIO_GRAPH.md)",
    )
    p_check.add_argument(
        "sources", nargs="*", metavar="SOURCE",
        help="Documents or directories to check (default: graph's directory)",
    )
    p_check.add_argument(
        "-v", "--verbose", action="store_true",
        help="Show INFO-level changes in addition to warnings",
    )
    p_check.set_defaults(func=cmd_check)

    # -- fetch ---------------------------------------------------------------
    p_fetch = sub.add_parser(
        "fetch",
        help="Print the full text of a specific section to stdout",
    )
    p_fetch.add_argument(
        "ref", metavar="REF",
        help=(
            'File and optional section. Examples:\n'
            '  "report.docx § Executive Summary"\n'
            '  report.docx  (lists all sections)'
        ),
    )
    p_fetch.add_argument(
        "--section", metavar="TITLE",
        help="Section title (alternative to § syntax)",
    )
    p_fetch.add_argument(
        "--sources", nargs="*", metavar="DIR",
        help="Directories to search for the file",
    )
    p_fetch.set_defaults(func=cmd_fetch)

    # -- stats ---------------------------------------------------------------
    p_stats = sub.add_parser(
        "stats",
        help="Show token-cost comparison for an existing graph",
    )
    p_stats.add_argument(
        "graph", metavar="FOLIO_GRAPH.md",
        help="Path to an existing FOLIO_GRAPH.md",
    )
    p_stats.add_argument(
        "--index", metavar="FOLIO_INDEX.md",
        help="Path to FOLIO_INDEX.md (default: same directory as graph)",
    )
    p_stats.set_defaults(func=cmd_stats)


    # -- stats-html ----------------------------------------------------------
    p_stats_html = sub.add_parser(
        "stats-html",
        help="Generate a self-contained HTML token savings dashboard",
    )
    p_stats_html.add_argument(
        "graph",
        metavar="FOLIO_GRAPH.md",
        help="Path to an existing FOLIO_GRAPH.md",
    )
    p_stats_html.add_argument(
        "--index", metavar="FOLIO_INDEX.md",
        help="Path to FOLIO_INDEX.md (default: same directory as graph)",
    )
    p_stats_html.add_argument(
        "-o", "--output", metavar="FILE",
        help="Output path (default: FOLIO_STATS.html in graph directory)",
    )
    p_stats_html.set_defaults(func=cmd_stats_html)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
