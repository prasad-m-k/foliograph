# Changelog

All notable changes to Foliograph will be documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.3.0] - 2026-05-27

### Changed

- SKILL.md token footprint reduced from ~8,500 to ~8,000 tokens (extractors refactored into shared helpers)
- Extractors consolidated: `_sec()` and `_rec()` helpers eliminate duplicated flush logic across all four formats
- Entity extractor merged into extractors section, removing a standalone section
- HTML dashboard now generated at runtime from a prompt instruction rather than carried as a 931-line template
- Default output changed to FOLIO_TIPS.md only; dashboard and cumulative chart generated on request
- Concealed mode: all processing runs inside a collapsible details block by default
- Graphify parity: confidence tags (EXTRACTED / INFERRED / AMBIGUOUS) on all relationships
- Graphify parity: god nodes, surprising connections, and suggested questions added to analysis
- README: "Get started in 3 steps" import instructions added as first section

### Added

- `analyse_graph()` function with god nodes, surprising connections, confidence tagging, and suggested questions
- Cumulative savings chart (on demand): 12-month SVG/canvas with three usage scenarios
- Drift check severity levels: INFO, WARNING, BREAKING

[0.3.0]: https://github.com/prasad-m-k/foliograph/releases/tag/v0.3.0

## [0.1.0] - 2026-05-26

### Added

- Core extractor for `.md`, `.txt`, `.docx`, `.pdf`, `.pptx`
- `DocumentRecord` data structure with sections, tables, figures, named entities
- `FOLIO_GRAPH.md` builder: structural skeleton with heading tree and summaries
- `FOLIO_INDEX.md` builder: alphabetical concept-to-location index
- `FOLIO_SESSION.md` session starter template
- CLI: `foliograph build` and `foliograph stats`
- Python API: `extract()` and `build()`
- Directory scanning with recursive file discovery
- Named entity extraction (acronyms, Model/Framework patterns, TitleCase phrases)
- Token cost estimate printed at build time
- 18 tests covering extractor, builder, and CLI paths
- Example: Scientific Discovery presentation (public domain PPTX)

[0.1.0]: https://github.com/prasad-m-k/foliograph/releases/tag/v0.1.0
