"""
foliograph.drift
~~~~~~~~~~~~~~~~
Detect graph drift: compare a freshly-extracted DocumentRecord set
against a previously built FOLIO_GRAPH.md and report what has changed.

Drift types:
  NEW_FILE       - file present now, not in previous graph
  REMOVED_FILE   - file in previous graph, no longer present
  SECTION_ADDED  - section exists now, not previously
  SECTION_REMOVED- section in previous graph, no longer present
  WORD_COUNT     - section word count changed by more than the threshold
  ENTITY_DRIFT   - named entity set changed significantly

Usage:
  from foliograph.drift import detect_drift, render_drift_report
  report = detect_drift(records, Path("FOLIO_GRAPH.md"))
  print(render_drift_report(report))
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from .extractor import DocumentRecord


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class DriftItem:
    kind: str           # NEW_FILE | REMOVED_FILE | SECTION_ADDED |
                        # SECTION_REMOVED | WORD_COUNT | ENTITY_DRIFT
    filename: str
    detail: str
    severity: str       # INFO | WARNING | BREAKING


@dataclass
class DriftReport:
    items: list[DriftItem] = field(default_factory=list)
    previous_build: Optional[str] = None   # timestamp from old graph
    is_clean: bool = True

    def add(self, item: DriftItem) -> None:
        self.items.append(item)
        if item.severity in ("WARNING", "BREAKING"):
            self.is_clean = False

    @property
    def breaking(self) -> list[DriftItem]:
        return [i for i in self.items if i.severity == "BREAKING"]

    @property
    def warnings(self) -> list[DriftItem]:
        return [i for i in self.items if i.severity == "WARNING"]

    @property
    def info(self) -> list[DriftItem]:
        return [i for i in self.items if i.severity == "INFO"]


# ---------------------------------------------------------------------------
# Graph parser (reads FOLIO_GRAPH.md to extract previous state)
# ---------------------------------------------------------------------------

def _parse_previous_graph(graph_path: Path) -> dict:
    """
    Parse an existing FOLIO_GRAPH.md into a dict:
    {
      "built": "2026-01-01 ...",
      "files": {
        "chapter1.docx": {
          "words": 4200,
          "sections": ["Section Title", ...],
          "entities": ["Term A", "Term B", ...],
        }
      }
    }
    """
    state: dict = {"built": None, "files": {}}
    if not graph_path.exists():
        return state

    text = graph_path.read_text(encoding="utf-8", errors="replace")

    # Extract build timestamp
    m = re.search(r"\*\*Built:\*\*\s+(.+)", text)
    if m:
        state["built"] = m.group(1).strip()

    # Parse per-file blocks
    # Each block starts with: ### `filename.ext` [TYPE]
    file_blocks = re.split(r"\n### `([^`]+)`", text)
    # file_blocks: [preamble, filename1, block1, filename2, block2, ...]
    i = 1
    while i < len(file_blocks) - 1:
        fname = file_blocks[i].strip()
        block = file_blocks[i + 1]
        i += 2

        file_state: dict = {"words": 0, "sections": [], "entities": []}

        # Word count
        wm = re.search(r"\*\*Words:\*\*\s+([\d,]+)", block)
        if wm:
            file_state["words"] = int(wm.group(1).replace(",", ""))

        # Section titles (lines starting with "- **Title**")
        for sm in re.finditer(r"-\s+\*\*([^*]+)\*\*", block):
            title = sm.group(1).strip()
            if title and title not in ("Structure:", "Tables:", "Figures:", "Key Terms:"):
                file_state["sections"].append(title)

        # Key terms
        km = re.search(r"\*\*Key Terms:\*\*\s+(.+)", block)
        if km:
            file_state["entities"] = [
                t.strip() for t in km.group(1).split(",") if t.strip()
            ]

        state["files"][fname] = file_state

    return state


# ---------------------------------------------------------------------------
# Drift detection
# ---------------------------------------------------------------------------

_WORD_COUNT_THRESHOLD = 0.15   # 15% change triggers WARNING
_ENTITY_DRIFT_THRESHOLD = 0.30  # 30% entity set change triggers WARNING


def detect_drift(
    records: list["DocumentRecord"],
    previous_graph: Path,
    word_threshold: float = _WORD_COUNT_THRESHOLD,
    entity_threshold: float = _ENTITY_DRIFT_THRESHOLD,
) -> DriftReport:
    """
    Compare current DocumentRecords against a previously built FOLIO_GRAPH.md.

    Parameters
    ----------
    records : list[DocumentRecord]
        Freshly extracted documents.
    previous_graph : Path
        Path to the existing FOLIO_GRAPH.md to compare against.
    word_threshold : float
        Fraction of word-count change that triggers a WARNING (default 15%).
    entity_threshold : float
        Fraction of entity-set change that triggers a WARNING (default 30%).

    Returns
    -------
    DriftReport
        All detected changes with severity labels.
    """
    report = DriftReport()
    prev = _parse_previous_graph(previous_graph)
    report.previous_build = prev.get("built")

    prev_files = prev.get("files", {})
    current_files = {rec.path.name: rec for rec in records}

    # NEW files
    for fname in current_files:
        if fname not in prev_files:
            report.add(DriftItem(
                kind="NEW_FILE",
                filename=fname,
                detail=f"New document not in previous graph.",
                severity="INFO",
            ))

    # REMOVED files
    for fname in prev_files:
        if fname not in current_files:
            report.add(DriftItem(
                kind="REMOVED_FILE",
                filename=fname,
                detail=f"Document was in previous graph but is no longer present.",
                severity="BREAKING",
            ))

    # Per-file comparison
    for fname, rec in current_files.items():
        if fname not in prev_files:
            continue  # already reported as NEW
        prev_state = prev_files[fname]

        # Word count drift
        prev_words = prev_state.get("words", 0)
        curr_words = rec.total_words
        if prev_words > 0:
            delta = abs(curr_words - prev_words) / prev_words
            if delta > word_threshold:
                direction = "grew" if curr_words > prev_words else "shrank"
                report.add(DriftItem(
                    kind="WORD_COUNT",
                    filename=fname,
                    detail=(
                        f"Word count {direction} by {delta:.0%}: "
                        f"{prev_words:,} -> {curr_words:,}. "
                        f"Graph may be stale."
                    ),
                    severity="WARNING",
                ))

        # Section drift
        prev_sections = set(prev_state.get("sections", []))
        curr_sections = {
            sec.title for sec in rec.sections if sec.level > 0 and sec.title
        }
        added = curr_sections - prev_sections
        removed = prev_sections - curr_sections

        for title in sorted(added):
            report.add(DriftItem(
                kind="SECTION_ADDED",
                filename=fname,
                detail=f'New section: "{title}"',
                severity="INFO",
            ))
        for title in sorted(removed):
            report.add(DriftItem(
                kind="SECTION_REMOVED",
                filename=fname,
                detail=f'Section removed: "{title}" - index entries may be stale.',
                severity="WARNING",
            ))

        # Entity drift
        prev_ents = set(e.lower() for e in prev_state.get("entities", []))
        curr_ents = set(e.lower() for e in rec.named_entities)
        if prev_ents:
            union = prev_ents | curr_ents
            intersection = prev_ents & curr_ents
            jaccard = len(intersection) / len(union) if union else 1.0
            drift = 1.0 - jaccard
            if drift > entity_threshold:
                report.add(DriftItem(
                    kind="ENTITY_DRIFT",
                    filename=fname,
                    detail=(
                        f"Key terms changed by {drift:.0%} (Jaccard similarity: "
                        f"{jaccard:.0%}). Consider rebuilding the index."
                    ),
                    severity="WARNING",
                ))

    return report


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def render_drift_report(report: DriftReport, verbose: bool = False) -> str:
    """Render a DriftReport as a human-readable string for CLI output."""
    lines: list[str] = []

    if report.previous_build:
        lines.append(f"Comparing against graph built: {report.previous_build}")
    lines.append("")

    if report.is_clean and not report.items:
        lines.append("Graph is up to date. No drift detected.")
        return "\n".join(lines)

    if report.breaking:
        lines.append(f"BREAKING ({len(report.breaking)})")
        for item in report.breaking:
            lines.append(f"  [{item.filename}] {item.detail}")
        lines.append("")

    if report.warnings:
        lines.append(f"WARNINGS ({len(report.warnings)})")
        for item in report.warnings:
            lines.append(f"  [{item.filename}] {item.detail}")
        lines.append("")

    if verbose and report.info:
        lines.append(f"INFO ({len(report.info)})")
        for item in report.info:
            lines.append(f"  [{item.filename}] {item.detail}")
        lines.append("")

    if not report.is_clean:
        lines.append("Run 'foliograph build' to rebuild the graph.")

    return "\n".join(lines)


def render_drift_block(report: DriftReport) -> list[str]:
    """Render drift status as a markdown block for embedding in FOLIO_GRAPH.md."""
    if report.is_clean and not report.items:
        return [
            "## Graph Status",
            "",
            "Graph is current. No drift detected.",
            "",
        ]

    lines = [
        "## Graph Status",
        "",
    ]

    if report.breaking:
        lines.append(f"> **BREAKING** - {len(report.breaking)} removed file(s).")
        for item in report.breaking:
            lines.append(f"> - `{item.filename}`: {item.detail}")
        lines.append(">")

    if report.warnings:
        lines.append(f"> **WARNINGS** - {len(report.warnings)} stale section(s) or changed content.")
        for item in report.warnings:
            lines.append(f"> - `{item.filename}`: {item.detail}")
        lines.append(">")

    if report.info:
        for item in report.info:
            lines.append(f"> INFO: `{item.filename}`: {item.detail}")

    lines.append("")
    lines.append("Run `foliograph build` to refresh.")
    lines.append("")
    return lines
