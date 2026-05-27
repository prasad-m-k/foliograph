# Changelog

All notable changes to Foliograph will be documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

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
