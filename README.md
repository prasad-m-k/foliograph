# Foliograph

**Pre-compile office documents into compact knowledge graphs for LLM sessions.**

Inspired by [Graphify](https://github.com/safishamsi/graphify) for code, Foliograph does the same for office documents: `.docx`, `.pdf`, `.pptx`, `.md`, and `.txt`.

Instead of loading entire documents into every session, you build the graph once and navigate by index. Token costs drop by 60-90% on document-heavy projects.

> **The skill generates the graph. You keep the graph. The only thing you install is `SKILL.md`.**

---

## Get started in 3 steps (claude.ai)

No terminal. No installation. Works entirely in your browser.

**Step 1: Download SKILL.md**

Download `SKILL.md` from this repository (click the file, then the download button).

**Step 2: Add it to your Claude Project**

1. Go to [claude.ai](https://claude.ai) and open or create a Project
2. Click the project name at the top of the left sidebar
3. Click **Add content** (or the **+** icon next to Files)
4. Upload `SKILL.md`
5. That is it. The skill is now active for every conversation in this Project.

**Step 3: Use it**

Upload any `.docx`, `.pdf`, `.pptx`, `.md`, or `.txt` file into a conversation and say:

```
foliograph this
```

You will get `FOLIO_TIPS.md` with your document map, key concepts, token savings, and ready-made commands. Ask for a visual dashboard with:

```
Show me an executive dashboard
```

**What you need:**
- A claude.ai account (Free, Pro, or Team)
- A Project (available on all plans)
- The `SKILL.md` file from this repo

---

## The Problem

Every new LLM session on a large document project starts blind. You paste the whole chapter, the whole spec, the whole report, because you don't know what the model will need. By message three you've burned most of your context window on content the model never touched.

Foliograph fixes this structurally:

```
Without Foliograph:
  Session start -> paste Chapter 4 (8,000 tokens) -> ask one question -> done
  Next session  -> paste Chapter 4 again (8,000 tokens) -> ...

With Foliograph:
  Session start -> load FOLIO_GRAPH.md (~400 tokens) -> "load Chapter 4 § The Swarm Model"
               -> fetch only that section (~600 tokens) -> done
```

The graph is built once. Every subsequent session pays only the index cost.

---

## Quickstart (Python CLI)

```bash
pip install foliograph
foliograph build my_project/ --name "My Project"
```

This produces three files in your working directory:

| File | Purpose |
|------|---------|
| `FOLIO_GRAPH.md` | Structural skeleton of every document: headings, summaries, word counts, figures, tables |
| `FOLIO_INDEX.md` | Concept to location index (168+ entries for a typical book) |
| `FOLIO_SESSION.md` | Copy-paste session starter prompt for any LLM |

---

## Installation

```bash
# Core (no heavy dependencies)
pip install foliograph

# With Python library support for each format
pip install "foliograph[docx]"
pip install "foliograph[pdf]"
pip install "foliograph[pptx]"
pip install "foliograph[all]"
```

---

## CLI usage

```bash
# Single file
foliograph build report.docx --name "Q3 Report"

# Multiple files
foliograph build chapter1.docx chapter2.docx appendix.pdf --name "My Book"

# Entire directory (recursive)
foliograph build ./manuscript/ --output ./graph/ --name "My Book"

# Check for drift against existing graph
foliograph check --graph FOLIO_GRAPH.md

# Fetch a specific section to stdout
foliograph fetch "chapter4.docx § The Swarm Model"

# Token savings stats
foliograph stats FOLIO_GRAPH.md

# Generate HTML savings dashboard
foliograph stats-html FOLIO_GRAPH.md
```

---

## Python API

```python
from foliograph.builder import build
from foliograph.extractor import extract

# Build graph from a list of files or directories
outputs = build(
    sources=["chapter1.docx", "appendix.pdf", "./slides/"],
    output_dir="./graph/",
    project_name="My Project",
)

# Extract a single document
rec = extract("report.docx")
print(rec.title)
print(rec.total_words)
for section in rec.sections:
    print(f"  {'  ' * section.level}{section.title} ({section.word_count}w)")
```

---

## How to use the graph in a session

1. **Start every session** by pasting the content of `FOLIO_SESSION.md`
2. **Ask questions by concept:** "What does the book say about Channel Siloing?"
3. **Load sections on demand:** "Load escalation_intelligence.md § The Swarm Model"
4. **Never reload** a section you've already discussed in the session

---

## Supported formats

| Format | Extension | Extraction method |
|--------|-----------|------------------|
| Word Document | `.docx` | `extract-text` / `python-docx` |
| PDF | `.pdf` | `pdftotext` / `pdfminer.six` |
| PowerPoint | `.pptx` | `extract-text` / `python-pptx` |
| Markdown | `.md` | Native parser |
| Plain Text | `.txt` | Native parser |

---

## Output format

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

**Key Terms:** Algorithmic Friction, Agent Churn, Escalation Debt, Feedback Loop
```

### FOLIO_INDEX.md (concept index)

```markdown
### S

- **Sentiment Drift** -> `chapter2.docx § Signal 1: Sentiment Drift`
- **Swarm Model** -> `chapter4.docx § The Swarm Model`
```

---

## Real-world example

The `examples/sample/` directory contains a worked example showing Foliograph output on a plain markdown document. Open `FOLIO_GRAPH.md` and `FOLIO_INDEX.md` to see the structure.

---

## Architecture

```
foliograph/
├── extractor.py     # Per-format extraction -> DocumentRecord
├── builder.py       # DocumentRecord[] -> FOLIO_GRAPH.md + FOLIO_INDEX.md
├── relationships.py # Cross-document relationship mapping
├── drift.py         # Graph drift detection
├── stats_html.py    # Token savings HTML dashboard
└── cli.py           # foliograph build / check / fetch / stats
```

---

## Contributing

Contributions welcome. The most valuable additions are:

- Better named-entity extraction
- `.xlsx` support (sheet names, column headers, key cell ranges)
- Google Docs / Notion export support
- `foliograph update` command for incremental rebuilds

Open an issue before starting a large feature. Some of these are already in progress.

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

## Author

**Prasad MK**
Research: [ssrn.com/author=10270516](https://ssrn.com/author=10270516)

---

## Acknowledgements

Foliograph is directly inspired by [Graphify](https://github.com/safishamsi/graphify) by Safi Shamsi, which demonstrated the same approach for codebases. The core insight is to pay the indexing cost once, query from the graph every session, and that insight belongs to that project. Foliograph extends it to office documents and to claude.ai chat environments where no terminal or IDE is available.
