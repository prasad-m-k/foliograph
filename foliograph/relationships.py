"""
foliograph.relationships
~~~~~~~~~~~~~~~~~~~~~~~~
Map relationships between documents and between sections.

Three relationship types:
  SHARES_CONCEPT  - two documents both mention the same named entity
  REFERENCES      - one section explicitly names another document or section title
  SEQUENCE        - heuristic ordering (Chapter N -> Chapter N+1, Slide N -> Slide N+1)

Output is written into FOLIO_GRAPH.md as a relationship block and into
FOLIO_RELATIONS.json for programmatic use.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .extractor import DocumentRecord


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Relationship:
    kind: str           # SHARES_CONCEPT | REFERENCES | SEQUENCE
    source_file: str
    source_section: str
    target_file: str
    target_section: str
    label: str          # human-readable reason


@dataclass
class RelationshipGraph:
    relationships: list[Relationship] = field(default_factory=list)

    def add(self, r: Relationship) -> None:
        self.relationships.append(r)

    def edges_from(self, filename: str) -> list[Relationship]:
        return [r for r in self.relationships if r.source_file == filename]

    def edges_to(self, filename: str) -> list[Relationship]:
        return [r for r in self.relationships if r.target_file == filename]

    def to_json(self) -> str:
        return json.dumps(
            [asdict(r) for r in self.relationships],
            indent=2,
        )


# ---------------------------------------------------------------------------
# Relationship detection
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.lower().strip())


def _find_shared_concepts(
    records: list["DocumentRecord"],
) -> list[Relationship]:
    """Documents that share named entities get SHARES_CONCEPT edges."""
    rels: list[Relationship] = []

    # Build entity -> [files] map
    entity_files: dict[str, list[str]] = {}
    for rec in records:
        for ent in rec.named_entities:
            key = _normalize(ent)
            if len(key) < 4:
                continue
            entity_files.setdefault(key, [])
            fname = rec.path.name
            if fname not in entity_files[key]:
                entity_files[key].append(fname)

    # Emit one edge per pair per shared concept (cap at 3 concepts per pair)
    pair_concepts: dict[tuple[str, str], list[str]] = {}
    for entity, files in entity_files.items():
        if len(files) < 2:
            continue
        for i, fa in enumerate(files):
            for fb in files[i + 1 :]:
                pair = (fa, fb) if fa < fb else (fb, fa)
                pair_concepts.setdefault(pair, [])
                if len(pair_concepts[pair]) < 3:
                    pair_concepts[pair].append(entity)

    for (fa, fb), concepts in pair_concepts.items():
        label = "Shared concepts: " + ", ".join(concepts)
        rels.append(Relationship(
            kind="SHARES_CONCEPT",
            source_file=fa, source_section="",
            target_file=fb, target_section="",
            label=label,
        ))

    return rels


def _find_references(
    records: list["DocumentRecord"],
) -> list[Relationship]:
    """
    Section text that mentions another document's filename (stem) or
    another section's exact title gets a REFERENCES edge.
    """
    rels: list[Relationship] = []

    # Build lookup: normalized title -> (file, title)
    section_titles: dict[str, tuple[str, str]] = {}
    file_stems: dict[str, str] = {}  # normalized stem -> filename

    for rec in records:
        stem = _normalize(rec.path.stem.replace("_", " ").replace("-", " "))
        file_stems[stem] = rec.path.name
        for sec in rec.sections:
            if sec.level > 0 and sec.title:
                key = _normalize(sec.title)
                if len(key) > 5:
                    section_titles[key] = (rec.path.name, sec.title)

    # Scan each section's summary for references to other titles / file names
    for rec in records:
        for sec in rec.sections:
            if not sec.summary:
                continue
            text_lower = _normalize(sec.summary + " " + sec.title)

            # Check against other section titles
            for norm_title, (target_file, orig_title) in section_titles.items():
                if target_file == rec.path.name:
                    continue  # skip self-references
                if len(norm_title) > 8 and norm_title in text_lower:
                    rels.append(Relationship(
                        kind="REFERENCES",
                        source_file=rec.path.name,
                        source_section=sec.title,
                        target_file=target_file,
                        target_section=orig_title,
                        label=f'"{sec.title}" mentions "{orig_title}"',
                    ))

            # Check against file stems
            for stem, target_file in file_stems.items():
                if target_file == rec.path.name:
                    continue
                if len(stem) > 4 and stem in text_lower:
                    rels.append(Relationship(
                        kind="REFERENCES",
                        source_file=rec.path.name,
                        source_section=sec.title,
                        target_file=target_file,
                        target_section="",
                        label=f'"{sec.title}" references file "{target_file}"',
                    ))

    return rels


def _find_sequences(
    records: list["DocumentRecord"],
) -> list[Relationship]:
    """
    Detect natural ordering: Chapter 1 -> Chapter 2, Slide 3 -> Slide 4, etc.
    Also handles 'Part I / Part II' and 'Appendix A / Appendix B' patterns.
    """
    rels: list[Relationship] = []

    numbered_patterns = [
        (r"chapter\s+(\d+)", "Chapter"),
        (r"slide\s+(\d+)", "Slide"),
        (r"part\s+(\d+)", "Part"),
        (r"section\s+(\d+)", "Section"),
        (r"appendix\s+([a-z])", "Appendix"),
        (r"module\s+(\d+)", "Module"),
    ]

    for rec in records:
        for sec in rec.sections:
            if sec.level > 2:
                continue
            title_lower = sec.title.lower()
            for pattern, kind in numbered_patterns:
                m = re.search(pattern, title_lower)
                if not m:
                    continue
                raw_num = m.group(1)
                # Find the next number
                try:
                    n = int(raw_num)
                    next_raw = str(n + 1)
                except ValueError:
                    # alphabetic (appendix a -> b)
                    next_raw = chr(ord(raw_num) + 1)

                # Search other sections in all records for the successor
                for other_rec in records:
                    for other_sec in other_rec.sections:
                        if other_sec.level > 2:
                            continue
                        other_lower = other_sec.title.lower()
                        successor_pattern = pattern.replace(
                            r"(\d+)", re.escape(next_raw)
                        ).replace(r"([a-z])", re.escape(next_raw))
                        if re.search(successor_pattern, other_lower):
                            rels.append(Relationship(
                                kind="SEQUENCE",
                                source_file=rec.path.name,
                                source_section=sec.title,
                                target_file=other_rec.path.name,
                                target_section=other_sec.title,
                                label=f"{kind} sequence: {sec.title} -> {other_sec.title}",
                            ))

    return rels


def build_relationships(
    records: list["DocumentRecord"],
) -> RelationshipGraph:
    """
    Build the full relationship graph for a set of DocumentRecords.

    Returns a RelationshipGraph containing SHARES_CONCEPT, REFERENCES,
    and SEQUENCE edges.
    """
    graph = RelationshipGraph()

    for rel in _find_shared_concepts(records):
        graph.add(rel)
    for rel in _find_references(records):
        graph.add(rel)
    for rel in _find_sequences(records):
        graph.add(rel)

    return graph


# ---------------------------------------------------------------------------
# Rendering helpers (used by builder.py)
# ---------------------------------------------------------------------------

def render_relationships_block(graph: RelationshipGraph) -> list[str]:
    """Render the relationship graph as a markdown section for FOLIO_GRAPH.md."""
    if not graph.relationships:
        return []

    lines = [
        "## Relationships",
        "",
        "Cross-document connections detected automatically.",
        "Format: `[KIND] source → target: reason`",
        "",
    ]

    by_kind = {"SHARES_CONCEPT": [], "REFERENCES": [], "SEQUENCE": []}
    for r in graph.relationships:
        by_kind.get(r.kind, by_kind["REFERENCES"]).append(r)

    kind_labels = {
        "SEQUENCE":      "Sequence (natural reading order)",
        "REFERENCES":    "References (explicit mentions)",
        "SHARES_CONCEPT": "Shared Concepts (common terminology)",
    }

    for kind, label in kind_labels.items():
        rels = by_kind[kind]
        if not rels:
            continue
        lines.append(f"### {label}")
        lines.append("")
        for r in rels:
            src = f"`{r.source_file}`" + (f" > *{r.source_section}*" if r.source_section else "")
            tgt = f"`{r.target_file}`" + (f" > *{r.target_section}*" if r.target_section else "")
            lines.append(f"- {src} → {tgt}")
            lines.append(f"  _{r.label}_")
        lines.append("")

    return lines
