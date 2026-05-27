# Foliograph

**Pre-compile office documents into compact knowledge graphs for LLM sessions.**

Inspired by [Graphify](https://github.com/safishamsi/graphify) for code, Foliograph does the same for office documents: `.docx`, `.pdf`, `.pptx`, `.md`, and `.txt`.

Instead of loading entire documents into every session, you build the graph once and navigate by index. Token costs drop by 60-90% on document-heavy projects.

> **The skill generates the graph. You keep the graph. The only thing you install is `SKILL.md`.**

---

## The Problem

Every new LLM session on a large document project starts blind. You paste the whole chapter, the whole spec, the whole report, because you don't know what the model will need. By message three you've burned most of your context window on content the model never touched.

Foliograph fixes this structurally:

```
Without Foliograph:
  Session start → paste Chapter 4 (8,000 tokens) → ask one question → done
  Next session   → paste Chapter 4 again (8,000 tokens) → ...

With Foliograph:
  Session start → load FOLIO_GRAPH.md (~400 tokens) → "load Chapter 4 § The Swarm Model"
               → fetch only that section (~600 tokens) → done
```

The graph is built once. Every subsequent session pays only the index cost.

---

## Quickstart

```bash
pip install foliograph
foliograph build my_project/ --name "My Project"
```

This produces three files in your working directory:

| File | Purpose |
|------|---------|
| `FOLIO_GRAPH.md` | Structural skeleton of every document: headings, summaries, word counts, figures, tables |
| `FOLIO_INDEX.md` | Concept → location index (168+ entries for a typical book) |
| `FOLIO_SESSION.md` | Copy-paste session starter prompt for any LLM |

---

## Installation

```bash
# Core (no heavy dependencies; uses CLI tools already on most systems)
pip install foliograph

# With Python library support for each format
pip install "foliograph[docx]"   # python-docx
pip install "foliograph[pdf]"    # pdfminer.six
pip install "foliograph[pptx]"   # python-pptx
pip install "foliograph[all]"    # everything
```

**System tools used (if available):** `extract-text`, `pdftotext`, `pdfinfo`
These are pre-installed in most document-processing environments.

---

## Usage

### CLI

```bash
# Single file
foliograph build report.docx --name "Q3 Report"

# Multiple files
foliograph build chapter1.docx chapter2.docx appendix.pdf --name "My Book"

# Entire directory (recursive)
foliograph build ./manuscript/ --output ./graph/ --name "My Book"

# Check token savings on an existing graph
foliograph stats FOLIO_GRAPH.md
```

### Python API

```python
from foliograph.builder import build
from foliograph.extractor import extract

# Build graph from a list of files or directories
outputs = build(
    sources=["chapter1.docx", "appendix.pdf", "./slides/"],
    output_dir="./graph/",
    project_name="My Project",
)
# outputs: {"graph": Path, "index": Path, "session": Path}

# Extract a single document
from foliograph.extractor import extract
rec = extract("report.docx")
print(rec.title)
print(rec.total_words)
for section in rec.sections:
    print(f"  {'  ' * section.level}{section.title} ({section.word_count}w)")
    print(f"    {section.summary}")
```

---

## How to Use the Graph in an LLM Session

1. **Start every session** by pasting the content of `FOLIO_SESSION.md`, or simply load `FOLIO_GRAPH.md` as context.

2. **Ask questions by concept:** "What does the book say about Channel Siloing?"
   The LLM checks `FOLIO_INDEX.md` → finds `§ Mistake 2: Channel Siloing` → fetches only that section.

3. **Load sections on demand:** "Load escalation_intelligence.md § The Swarm Model"
   You paste only that section (typically 400-800 tokens) rather than the whole document.

4. **Never reload** a section you've already discussed in the session.

---

## Supported Formats

| Format | Extension | Extraction Method |
|--------|-----------|------------------|
| Word Document | `.docx` | `extract-text` / `python-docx` |
| PDF | `.pdf` | `pdftotext` / `pdfminer.six` |
| PowerPoint | `.pptx` | `extract-text` / `python-pptx` |
| Markdown | `.md` | Native parser |
| Plain Text | `.txt` | Native parser |

---

## Output Format

### FOLIO_GRAPH.md (structure map)

```markdown
### `chapter4.docx` [DOCX]
**Title:** The Swarm Model
**Words:** 2,847

**Structure:**
- **The Swarm Model**
  > Replacing the Hierarchy with Parallel Expert Engagement.
  - **Why Sequential Escalation Fails at Scale** (187w)
    > The sequential model has a structural bottleneck at every tier boundary.
  - **How AI Assembles the Swarm** (312w)
    > Swarm assembly uses four criteria evaluated simultaneously.
    - **Criterion 1: Skill Match** (94w)
        > AI matches issue class taxonomy against each specialist's resolution history.

**Tables:**
- Table 4.1 The Three Swarm Roles
- Table 4.2 Swarm Lifecycle Phases and Target Timeframes

**Key Terms:** Algorithmic Friction, Agent Churn, Escalation Debt, Feedback Loop, ...
```

### FOLIO_INDEX.md (concept index)

```markdown
### S

- **Sentiment Drift** → `chapter2.docx § Signal 1: Sentiment Drift`
- **Signal Maturity Matrix** → `chapter3.docx § Signal Maturity Matrix`
- **Swarm Model** → `chapter4.docx § The Swarm Model`
- **Swarm Brief Generator** → `appendix_c.docx`
```

---

## Real-World Example

The `examples/scientific_discovery/` directory contains a full worked example using the Scientific Discovery presentation.

```
Input:  escalation_intelligence.md  (31,669 bytes, ~8,000 tokens)
Output: FOLIO_GRAPH.md              (13,162 bytes)
        FOLIO_INDEX.md              (11,091 bytes)
        FOLIO_SESSION.md            (400 bytes)

Combined graph tokens (est.): ~6,063
vs. loading full text each session: ~8,000+ tokens per session
```

For multi-document projects (full manuscript, multiple appendices, slide decks), the savings compound significantly, typically 70-90% per session.

---

## Architecture

```
foliograph/
├── extractor.py     # Per-format extraction → DocumentRecord
├── builder.py       # DocumentRecord[] → FOLIO_GRAPH.md + FOLIO_INDEX.md
└── cli.py           # foliograph build / foliograph stats
```

**`DocumentRecord`** is the central data structure:
- `sections: list[Section]`: heading tree with summaries and word counts
- `tables: list[str]`: table captions
- `figures: list[str]`: figure captions
- `named_entities: list[str]`: key terms, frameworks, acronyms
- `raw_text: str`: full text (used during build, not written to graph)

Extraction is **format-aware but dependency-light**: each extractor tries the CLI tool first, falls back to the Python library, and degrades gracefully if neither is available.

---

## Contributing

Contributions welcome. The most valuable additions are:

- Better named-entity extraction (the current regex approach misses domain-specific terms)
- `.xlsx` support (extract sheet names, column headers, and key cell ranges)
- Google Docs / Notion export support
- A `foliograph update` command that rebuilds only changed files

Please open an issue before starting a large feature. Some of these are already in progress.

```bash
git clone https://github.com/prasad-m-k/foliograph
cd foliograph
pip install -e ".[dev]"
pytest tests/
```

---

## License

MIT. See [LICENSE](LICENSE).

---

## Acknowledgements

Foliograph is directly inspired by [Graphify](https://github.com/safishamsi/graphify) by Safi Shamsi, which demonstrated the same approach for codebases. The core insight is to pay the indexing cost once, query from the graph every session, and that belongs to that project. Foliograph extends it to office documents and to claude.ai chat environments where no terminal or IDE is available.

---

## Author

**Prasad MK**
Research: [ssrn.com/author=10270516](https://ssrn.com/author=10270516)
