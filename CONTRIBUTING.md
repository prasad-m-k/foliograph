# Contributing to Foliograph

Thank you for considering a contribution.

## Setup

```bash
git clone https://github.com/prasad-m-k/foliograph
cd foliograph
pip install -e ".[dev]"
pytest tests/         # should be 18 passed
```

## What to work on

Issues labeled `good first issue` are a safe starting point.
Open an issue before starting anything large. Some features are already in progress.

### High-value areas

- **Better named-entity extraction**: the current regex approach misses
  domain-specific terms. A lightweight NER model (or spaCy integration as
  an optional dependency) would improve index quality significantly.

- **Excel / `.xlsx` support**: extract sheet names, column headers,
  named ranges, and key cell values.

- **Incremental rebuild**: `foliograph update` that re-extracts only
  files modified since the last build, rather than rebuilding everything.

- **Google Docs / Notion support**: accept a Google Docs URL or a
  Notion export folder as a source.

- **`foliograph fetch`**: given a `FOLIO_INDEX.md` entry, extract and
  print only that section's full text to stdout, so it can be piped
  directly into an LLM prompt.

## Code style

- Black formatting, no line-length enforcement
- Type annotations on all public functions
- Docstrings on public API only
- No new mandatory dependencies without discussion

## Tests

All new features need tests. Run with `pytest tests/ -v`.
If you add a new extractor for a new format, add a fixture and at minimum
four tests: title extraction, section count, word count, and round-trip
via the public `extract()` API.

## Pull requests

- One feature or fix per PR
- Update `CHANGELOG.md` under `[Unreleased]`
- Tests must pass before review
