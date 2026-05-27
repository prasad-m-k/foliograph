---
name: foliograph
description: >
  Use this skill whenever the user wants to build a knowledge graph from
  uploaded documents, or says "foliograph this", "build the graph", "map this
  document", "index this for the session", "reduce my token cost", "split my
  document", "set up project instructions", "save my session summary", or
  "check for drift". Also triggers when the user uploads one or more .docx,
  .pdf, .pptx, .md, or .txt files and wants to navigate them efficiently across
  sessions. Produces FOLIO_GRAPH.md, FOLIO_INDEX.md, FOLIO_RELATIONS.json,
  FOLIO_SESSION.md, FOLIO_PROJECT_INSTRUCTIONS.md, and FOLIO_TIPS.md as
  downloadable files. Non-technical users should be guided through each step
  with plain English explanations. Always show the token savings estimate.
---

# Foliograph Skill
## Token-Efficient Document Navigation for claude.ai

**Created by:** Prasad MK
**Research:** https://ssrn.com/author=10270516

Foliograph pre-compiles your uploaded documents into a compact knowledge
graph. Instead of re-reading entire files every session, you load the graph
once and fetch only the sections you need.

Inspired by [Graphify](https://github.com/safishamsi/graphify) for IDE
environments. Foliograph brings the same approach to claude.ai chat with no terminal, no installation, and no setup beyond uploading this skill file.

> The skill generates the graph. You keep the graph.
> The only thing you install is this SKILL.md file.

---

## What to say to trigger each feature

| You say | What happens |
|---|---|
| "Foliograph this" / upload files | Builds the full graph |
| "Split my document for Foliograph" | Splits one large doc into chapters |
| "Set up my Project instructions" | Generates text to paste into Project settings |
| "Save my session summary" | Compresses today's work into a reusable note |
| "Has the graph changed?" | Checks for drift against uploaded source |
| "Load [file] § [Section Title]" | Fetches only that section |
| "What does the index say about [topic]?" | Answers from index without loading files |

---

## Seven ways to reduce token cost further

Before building the graph, Claude should always brief the user on these:

1. **Split large documents into chapters**: load only the chapter you need
2. **Put FOLIO_GRAPH.md in Project instructions**: zero cost per session
3. **Ask the index first**: ask "What does the index say about X?" before loading
4. **Build micro-graphs for deep sessions**: graph one chapter at a time
5. **Save session summaries**: start next session from a 200-word note
6. **Rebuild only when content changes**: not every session
7. **Never paste full documents**: always load by section name

---

## Environment Setup (always run first)

```python
import subprocess, sys, os, re, json
from datetime import datetime, timezone

for pkg in ["python-docx", "python-pptx", "pdfminer.six"]:
    subprocess.run(
        [sys.executable, "-m", "pip", "install", pkg,
         "--break-system-packages", "-q"],
        capture_output=True
    )
```

---

## Feature 1: Build the full graph

### Detect uploaded files

```python
upload_dir = "/mnt/user-data/uploads"
output_dir = "/mnt/user-data/outputs"
os.makedirs(output_dir, exist_ok=True)

SUPPORTED = {".docx", ".pdf", ".pptx", ".md", ".txt"}

uploaded = [
    f for f in os.listdir(upload_dir)
    if os.path.splitext(f)[1].lower() in SUPPORTED
]
```

### Extract .pptx

```python
from pptx import Presentation

def extract_pptx(path):
    prs = Presentation(path)
    sections, all_text = [], []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.append(t)
        if texts:
            title = texts[0]
            body = " ".join(texts[1:])
            summary = body[:160] + ("…" if len(body) > 160 else "")
            sections.append({"level": 1, "title": title, "summary": summary,
                              "word_count": len(body.split()), "page_hint": f"slide {i}"})
            all_text.extend(texts)
    return {
        "filename": os.path.basename(path), "file_type": "PPTX",
        "title": sections[0]["title"] if sections else os.path.basename(path),
        "total_words": len(" ".join(all_text).split()),
        "total_pages": len(prs.slides), "sections": sections,
        "tables": [], "figures": [], "raw_text": "\n".join(all_text)
    }
```

### Extract .docx

```python
import docx as dx

def extract_docx(path):
    doc = dx.Document(path)
    sections, current_body, current_heading, all_text = [], [], None, []

    def flush(heading, body):
        text = " ".join(body).strip()
        lvl, ttl = heading if heading else (0, "(preamble)")
        first = re.split(r'[.!?]', text)[0][:160] if text else ""
        sections.append({"level": lvl, "title": ttl, "summary": first,
                          "word_count": len(text.split()), "page_hint": None})

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t:
            continue
        all_text.append(t)
        style = para.style.name if para.style else ""
        m = re.match(r'Heading (\d)', style)
        if m:
            flush(current_heading, current_body)
            current_heading = (int(m.group(1)), t)
            current_body = []
        else:
            current_body.append(t)
    flush(current_heading, current_body)

    tables = []
    for tbl in doc.tables:
        if tbl.rows:
            header = " | ".join(c.text.strip() for c in tbl.rows[0].cells)
            if header.strip() and " | " not in header[:5]:
                tables.append(header[:120])

    return {
        "filename": os.path.basename(path), "file_type": "DOCX",
        "title": next((s["title"] for s in sections if s["level"] == 1),
                      os.path.basename(path)),
        "total_words": len(" ".join(all_text).split()),
        "total_pages": None, "sections": sections,
        "tables": tables[:20], "figures": [], "raw_text": "\n".join(all_text)
    }
```

### Extract .pdf

```python
from pdfminer.high_level import extract_text as pdf_extract

def extract_pdf(path):
    raw = pdf_extract(path) or ""
    lines = raw.splitlines()
    sections, current_body, current_heading = [], [], None

    def is_heading(line):
        line = line.strip()
        if not line or len(line) > 120:
            return False
        if re.match(r'^(Chapter|Section|CHAPTER|SECTION|\d+\.)\s+\S', line):
            return True
        if re.match(r'^#{1,4}\s+', line):
            return True
        if len(line) < 80 and line == line.title() and not line.endswith('.'):
            return True
        return False

    def flush(heading, body):
        text = " ".join(body).strip()
        lvl, ttl = heading if heading else (0, "(preamble)")
        first = re.split(r'[.!?]', text)[0][:160] if text else ""
        sections.append({"level": lvl, "title": ttl, "summary": first,
                          "word_count": len(text.split()), "page_hint": None})

    for line in lines:
        if is_heading(line):
            flush(current_heading, current_body)
            current_heading = (1, line.strip())
            current_body = []
        else:
            current_body.append(line)
    flush(current_heading, current_body)

    return {
        "filename": os.path.basename(path), "file_type": "PDF",
        "title": sections[0]["title"] if sections else os.path.basename(path),
        "total_words": len(raw.split()), "total_pages": None,
        "sections": sections, "tables": [], "figures": [], "raw_text": raw
    }
```

### Extract .md / .txt

```python
def extract_md(path):
    raw = open(path, encoding="utf-8", errors="replace").read()
    lines = raw.splitlines()
    sections, current_body, current_heading = [], [], None

    def flush(heading, body):
        text = " ".join(body).strip()
        lvl, ttl = heading if heading else (0, "(preamble)")
        first = re.split(r'[.!?]', text)[0][:160] if text else ""
        sections.append({"level": lvl, "title": ttl, "summary": first,
                          "word_count": len(text.split()), "page_hint": None})

    for line in lines:
        m = re.match(r'^(#{1,4})\s+(.+)', line)
        if m:
            flush(current_heading, current_body)
            current_heading = (len(m.group(1)), m.group(2).strip())
            current_body = []
        else:
            current_body.append(line)
    flush(current_heading, current_body)

    return {
        "filename": os.path.basename(path),
        "file_type": os.path.splitext(path)[1].upper().lstrip("."),
        "title": sections[0]["title"] if sections else os.path.basename(path),
        "total_words": len(raw.split()), "total_pages": None,
        "sections": sections, "tables": [], "figures": [], "raw_text": raw
    }

EXTRACTORS = {".docx": extract_docx, ".pptx": extract_pptx,
              ".pdf": extract_pdf, ".md": extract_md, ".txt": extract_md}
```

### Extract named entities

```python
def extract_entities(text):
    text = re.sub(r'\s+', ' ', text)
    entities = set()
    for m in re.finditer(r'\b([A-Z]{2,6})\b', text):
        entities.add(m.group(1))
    for m in re.finditer(
        r'\b((?:[A-Z][a-z]+ ){1,4}(?:Model|Framework|System|Method|'
        r'Score|Rate|Loop|Matrix|Protocol|Index|Threshold|Spectrum))\b', text):
        entities.add(m.group(1).strip())
    for m in re.finditer(r'\b((?:[A-Z][a-z]+\s){1,3}[A-Z][a-z]+)\b', text):
        phrase = m.group(1).strip()
        if 2 <= len(phrase.split()) <= 4:
            entities.add(phrase)
    stopwords = {"The Book", "This Chapter", "In Practice", "For Example",
                 "As A", "In My", "At The", "Of The"}
    return sorted(
        e for e in entities
        if e not in stopwords and "|" not in e
        and not re.match(r'^\d', e) and len(e) > 2
    )[:40]
```

### Build relationships

```python
def build_relationships(records):
    rels = []
    entity_files = {}
    for rec in records:
        for ent in extract_entities(rec["raw_text"]):
            key = ent.lower()
            entity_files.setdefault(key, [])
            if rec["filename"] not in entity_files[key]:
                entity_files[key].append(rec["filename"])
    pair_concepts = {}
    for entity, files in entity_files.items():
        if len(files) < 2:
            continue
        for i, fa in enumerate(files):
            for fb in files[i+1:]:
                pair = (fa, fb) if fa < fb else (fb, fa)
                pair_concepts.setdefault(pair, [])
                if len(pair_concepts[pair]) < 3:
                    pair_concepts[pair].append(entity)
    for (fa, fb), concepts in pair_concepts.items():
        rels.append({"kind": "SHARES_CONCEPT", "source_file": fa,
                     "source_section": "", "target_file": fb,
                     "target_section": "", "label": "Shared: " + ", ".join(concepts)})
    patterns = [(r'chapter\s+(\d+)', 'Chapter'), (r'slide\s+(\d+)', 'Slide'),
                (r'part\s+(\d+)', 'Part'), (r'section\s+(\d+)', 'Section')]
    for rec in records:
        for sec in rec["sections"]:
            for pattern, kind in patterns:
                m = re.search(pattern, sec["title"].lower())
                if not m:
                    continue
                try:
                    next_n = str(int(m.group(1)) + 1)
                except ValueError:
                    continue
                for other_rec in records:
                    for other_sec in other_rec["sections"]:
                        if re.search(pattern.replace(r'(\d+)', re.escape(next_n)),
                                     other_sec["title"].lower()):
                            rels.append({"kind": "SEQUENCE",
                                         "source_file": rec["filename"],
                                         "source_section": sec["title"],
                                         "target_file": other_rec["filename"],
                                         "target_section": other_sec["title"],
                                         "label": f"{kind} sequence"})
    return rels
```

### Build all graph files

```python
GRAPH_TOKENS = 1300
TOKENS_PER_WORD = 1.3

def build_graph(records, project_name="My Project"):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    total_words = sum(r["total_words"] for r in records)
    rels = build_relationships(records)

    # ── FOLIO_GRAPH.md ───────────────────────────────────────────────────────
    lines = [
        "# Foliograph Knowledge Graph", "",
        f"**Project:** {project_name}  ",
        f"**Built:** {now}  ",
        f"**Documents:** {len(records)}  ",
        f"**Total words:** {total_words:,}  ",
        f"**Relationships:** {len(rels)}  ", "",
        "> Structural map of your document set.",
        "> Use FOLIO_INDEX.md to locate concepts.",
        "> Load sections on demand. Never entire files.",
        "> See FOLIO_TIPS.md for how to get the most from this graph.",
        "", "---", "", "## Documents", ""
    ]
    for rec in records:
        pages = f" | {rec['total_pages']} pages" if rec.get("total_pages") else ""
        lines += [
            f"### `{rec['filename']}` [{rec['file_type']}]",
            f"**Title:** {rec['title']}  ",
            f"**Words:** {rec['total_words']:,}{pages}  ", ""
        ]
        titled = [s for s in rec["sections"] if s["level"] > 0]
        if titled:
            lines.append("**Structure:**")
            for sec in titled:
                indent = "  " * (sec["level"] - 1)
                loc = f" [{sec['page_hint']}]" if sec.get("page_hint") else ""
                wc = f" ({sec['word_count']}w)" if sec["word_count"] > 20 else ""
                lines.append(f"{indent}- **{sec['title']}**{loc}{wc}")
                if sec["summary"]:
                    lines.append(f"{'  ' * sec['level']}  > {sec['summary']}")
            lines.append("")
        if rec["tables"]:
            lines.append("**Tables:**")
            for t in rec["tables"]:
                lines.append(f"- {t}")
            lines.append("")
        entities = extract_entities(rec["raw_text"])
        if entities:
            lines.append(f"**Key Terms:** {', '.join(entities[:25])}")
            lines.append("")
        lines += ["---", ""]
    if rels:
        lines += ["## Relationships", ""]
        by_kind = {}
        for r in rels:
            by_kind.setdefault(r["kind"], []).append(r)
        for kind, label in [("SEQUENCE", "Sequence (reading order)"),
                             ("REFERENCES", "References"),
                             ("SHARES_CONCEPT", "Shared Concepts")]:
            if kind not in by_kind:
                continue
            lines += [f"### {label}", ""]
            for r in by_kind[kind]:
                src = f"`{r['source_file']}`" + (f" § *{r['source_section']}*" if r["source_section"] else "")
                tgt = f"`{r['target_file']}`" + (f" § *{r['target_section']}*" if r["target_section"] else "")
                lines += [f"- {src} → {tgt}", f"  _{r['label']}_"]
            lines.append("")
    graph_md = "\n".join(lines)

    # ── FOLIO_INDEX.md ───────────────────────────────────────────────────────
    entries = []
    seen = set()
    for rec in records:
        for sec in rec["sections"]:
            if sec["level"] == 0 or not sec["title"]:
                continue
            key = sec["title"].lower()
            if key in seen:
                continue
            seen.add(key)
            loc = sec.get("page_hint") or f"§ {sec['title']}"
            entries.append((sec["title"], rec["filename"], loc))
        for ent in extract_entities(rec["raw_text"]):
            key = ent.lower()
            if key in seen:
                continue
            seen.add(key)
            location = rec["filename"]
            for sec in rec["sections"]:
                if ent.lower() in sec["title"].lower() or ent.lower() in sec["summary"].lower():
                    location = f"{rec['filename']} § {sec['title']}"
                    break
            entries.append((ent, rec["filename"], location))
    entries.sort(key=lambda x: x[0].lower())

    idx_lines = [
        "# Foliograph Concept Index", "",
        f"**Project:** {project_name}  ",
        f"**Built:** {now}  ",
        f"**Entries:** {len(entries)}  ", "",
        "> Find any concept, section, or term across all your documents.",
        '> Format: **Concept** → `file § Section`',
        "> How to use: tell Claude 'Load [file] § [Section]' to read just that part.",
        "", "---", ""
    ]
    current_letter = ""
    for concept, fname, location in entries:
        first = concept[0].upper() if concept else "?"
        if first != current_letter:
            current_letter = first
            idx_lines += ["", f"### {first}", ""]
        idx_lines.append(f"- **{concept}** → `{location}`")
    index_md = "\n".join(idx_lines)

    # ── FOLIO_RELATIONS.json ─────────────────────────────────────────────────
    relations_json = json.dumps(rels, indent=2)

    # ── FOLIO_SESSION.md ─────────────────────────────────────────────────────
    file_list = "\n".join(f"  - {r['filename']}" for r in records)
    session_md = f"""# Foliograph Session Starter
## Copy and paste this at the start of every new Claude session

---

You have a Foliograph knowledge graph for this project.

**How to use it, in plain English:**

1. I already have the graph loaded. You do NOT need to upload your documents again.
2. To find something, say: "What does the index say about [topic]?"
3. To read a section, say: "Load [filename] § [Section Title]"
4. To check if anything has changed, say: "Has the graph changed?" and upload the source file.
5. At the end of this session, say: "Save my session summary" to get a compact note for next time.

**Documents in this project:**
{file_list}

**Token-saving rules (applied automatically):**
- Check the index before loading any section
- Load one section at a time, not the whole document
- Never reload a section already discussed today
- Flag any drift warning before answering questions about that section

---
"""

    # ── FOLIO_PROJECT_INSTRUCTIONS.md ────────────────────────────────────────
    project_instructions_md = f"""# Foliograph Project Instructions
## Paste the text below into your Claude Project's custom instructions

---

PASTE THIS INTO PROJECT SETTINGS → INSTRUCTIONS:

────────────────────────────────────────────────────
This project uses Foliograph for token-efficient document navigation.

I have a pre-built knowledge graph for this document set. Rules:

1. Always check FOLIO_INDEX.md before loading any source file.
2. Load sections on demand only. Never entire documents.
   Format: "Load [filename] § [Section Title]"
3. Never re-read a section already discussed in this session.
4. If you see a drift warning in FOLIO_GRAPH.md, flag it before
   answering questions about the affected section.
5. When asked to save a session summary, write a 200-word note
   covering decisions made, sections edited, and next steps.

Documents in scope:
{file_list}
────────────────────────────────────────────────────

HOW TO ADD THIS TO YOUR PROJECT:
1. Go to claude.ai
2. Open your Project
3. Click the project name at the top
4. Find "Instructions" or "Custom instructions"
5. Paste the text between the lines above
6. Save

Once added, Claude will follow these rules automatically in every
session. You will not need to paste the session starter each time.
"""

    # ── FOLIO_TIPS.md ─────────────────────────────────────────────────────────
    full_tokens = int(total_words * TOKENS_PER_WORD)
    graph_tokens_est = (len(graph_md) + len(index_md)) // 4
    saved_per_session = max(0, full_tokens - graph_tokens_est - 800)
    saved_per_month_10 = saved_per_session * 10
    saved_per_year = saved_per_month_10 * 12

    tips_md = f"""# Foliograph Tips
## How to get the most from your knowledge graph

**Project:** {project_name}
**Built:** {now}

---

## Your token savings at a glance

| | Tokens |
|---|---|
| Loading all documents in full | ~{full_tokens:,} per session |
| Using Foliograph (graph + 3 sections) | ~{graph_tokens_est + 800:,} per session |
| Saved per session | ~{saved_per_session:,} |
| Saved per month (10 sessions) | ~{saved_per_month_10:,} |
| Saved per year | ~{saved_per_year:,} |

---

## The 7 habits of a token-efficient user

### Habit 1: Ask the index first
Before loading any section, ask:
> "What does the index say about [topic]?"

Claude will check FOLIO_INDEX.md (already loaded) and often
answer from the summary alone. You only load a section when the summary
is not enough.

### Habit 2: Load by section name, not by file
Instead of:
> "Here is my whole manuscript, please review chapter 4"

Say:
> "Load manuscript.docx § Chapter 4: The Swarm Model"

Claude reads only that section, not the 50,000-word file.

### Habit 3: Put the graph in your Project instructions
Go to your Claude Project → Instructions and paste the contents of
FOLIO_PROJECT_INSTRUCTIONS.md. Claude will then have the graph rules
loaded automatically at zero cost in every session.

### Habit 4: Save a session summary at the end
At the end of any productive session, say:
> "Save my session summary"

Claude writes a 200-word note covering what was decided, what was
changed, and what comes next. Start your next session with that note
instead of reloading the graph.

### Habit 5: Split large documents into chapters
If you have one large document (50,000+ words), split it into chapter
files before running Foliograph. Say:
> "Split my document into chapters for Foliograph"

Claude will separate it and you can upload the individual files.
Each chapter fetch is then 10x smaller than loading the whole document.

### Habit 6: Build a micro-graph for deep work sessions
When spending a whole session on one chapter, say:
> "Build a micro-graph for chapter 4 only"

Claude builds a graph just for that file. Startup cost drops from
~{graph_tokens_est:,} tokens to under 400 tokens.

### Habit 7: Rebuild only when content changes significantly
You do not need to rebuild the graph every session. Rebuild when:
- You add a new chapter or section
- You restructure the document significantly
- Claude flags a drift warning about word count changes
- More than 2-3 weeks have passed since the last build

---

## What each output file does

| File | What it is | When to use it |
|---|---|---|
| FOLIO_GRAPH.md | Structural map of all documents | Load once per session for orientation |
| FOLIO_INDEX.md | Concept → location lookup | Check before every section load |
| FOLIO_RELATIONS.json | Machine-readable connections | Used by advanced tools and visualizers |
| FOLIO_SESSION.md | Session starter prompt | Paste at start of each new session |
| FOLIO_PROJECT_INSTRUCTIONS.md | Permanent project rules | Paste into Project settings once |
| FOLIO_TIPS.md | This file | Keep as a reference |

---

## Quick reference: what to say

| Goal | Say this |
|---|---|
| Find a topic | "What does the index say about [topic]?" |
| Read a section | "Load [file] § [Section Title]" |
| Check for changes | "Has the graph changed?" + upload source file |
| Split a large doc | "Split my document into chapters for Foliograph" |
| Deep work session | "Build a micro-graph for [filename] only" |
| End of session | "Save my session summary" |
| Rebuild everything | Upload files + "Foliograph this" |

---

*Foliograph, created by Prasad MK*
*Research: https://ssrn.com/author=10270516*
*Inspired by Graphify: https://github.com/safishamsi/graphify*
"""

    graph_tokens_out = (len(graph_md) + len(index_md)) // 4
    return {
        "FOLIO_GRAPH.md": graph_md,
        "FOLIO_INDEX.md": index_md,
        "FOLIO_RELATIONS.json": relations_json,
        "FOLIO_SESSION.md": session_md,
        "FOLIO_PROJECT_INSTRUCTIONS.md": project_instructions_md,
        "FOLIO_TIPS.md": tips_md,
        "_graph_tokens": graph_tokens_out,
        "_full_tokens": full_tokens,
        "_saved_per_session": saved_per_session,
    }
```

---

## Feature 2: Split a large document into chapters

Trigger: user says "split my document" or uploads a large .docx/.md and asks
to prepare it for Foliograph.

```python
def split_document(path, output_dir):
    """Split a .docx or .md into one file per top-level section."""
    ext = os.path.splitext(path)[1].lower()
    os.makedirs(output_dir, exist_ok=True)
    created = []

    if ext == ".docx":
        import docx as dx
        from docx import Document
        doc = dx.Document(path)
        current_heading = None
        current_paras = []
        base = os.path.splitext(os.path.basename(path))[0]

        def save_chunk(heading, paras, idx):
            if not paras:
                return
            new_doc = Document()
            safe = re.sub(r'[^\w\s-]', '', heading or f"section_{idx}")[:40].strip()
            fname = f"{base}_{idx:02d}_{safe.replace(' ', '_')}.docx"
            fpath = os.path.join(output_dir, fname)
            if heading:
                new_doc.add_heading(heading, level=1)
            for p in paras:
                new_p = new_doc.add_paragraph(p.text)
                new_p.style = new_doc.styles['Normal']
            new_doc.save(fpath)
            created.append(fname)

        idx = 0
        for para in doc.paragraphs:
            style = para.style.name if para.style else ""
            if re.match(r'Heading 1', style) and para.text.strip():
                save_chunk(current_heading, current_paras, idx)
                current_heading = para.text.strip()
                current_paras = []
                idx += 1
            else:
                current_paras.append(para)
        save_chunk(current_heading, current_paras, idx)

    elif ext == ".md":
        raw = open(path, encoding="utf-8", errors="replace").read()
        chunks = re.split(r'\n(?=# )', raw)
        base = os.path.splitext(os.path.basename(path))[0]
        for i, chunk in enumerate(chunks):
            if not chunk.strip():
                continue
            first_line = chunk.splitlines()[0]
            title = re.sub(r'^#+\s*', '', first_line)[:40].strip()
            safe = re.sub(r'[^\w\s-]', '', title).strip().replace(' ', '_')
            fname = f"{base}_{i:02d}_{safe}.md"
            fpath = os.path.join(output_dir, fname)
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(chunk)
            created.append(fname)

    return created
```

After splitting, tell the user:

```
Splitting complete. Created [N] chapter files:
  - manuscript_01_Introduction.docx
  - manuscript_02_Chapter_One.docx
  ...

Next step: upload these individual chapter files and say
"Foliograph this" to build the graph.

Tip: When working on Chapter 2, upload only that file.
Claude will load the graph (~400 tokens) plus the section
you need (~600 tokens) instead of the whole manuscript
(~40,000 tokens).
```

---

## Feature 3: Save a session summary

Trigger: user says "save my session summary" or "compress today's work".

```python
def build_session_summary(project_name, conversation_summary, next_steps):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    summary = f"""# Session Summary
## {project_name} | {now}

### What we worked on
{conversation_summary}

### Decisions made
[Claude fills this from the conversation]

### Next steps
{next_steps}

### How to use this note
Paste this at the start of your next session instead of the full
session starter. It gives Claude just enough context to continue
without reloading the graph from scratch.

Estimated tokens: ~{len(conversation_summary.split()) * 2} vs graph startup ~1,300
"""
    return summary
```

When the user asks for a session summary, Claude should:
1. Write a 150-200 word summary of what was discussed and decided
2. List 3-5 specific next steps
3. Save it as FOLIO_SESSION_SUMMARY_[date].md
4. Tell the user: "Start your next session by pasting this note.
   It costs ~300 tokens vs ~1,300 for the full graph."

---

## Feature 4: Build a micro-graph (single file)

Trigger: user says "build a micro-graph for [filename]" or
"I'm only working on chapter 4 today".

Run the full build workflow on just that one file and tell the user:

```
Micro-graph built for [filename].

Startup cost: ~380 tokens (vs ~1,300 for the full project graph)
Sections available: [list the sections]

Load a section by saying:
  "Load [filename] § [Section Title]"
```

---

## Feature 5: Drift check

Trigger: user says "has the graph changed?" or "check for drift"
and uploads a source file alongside FOLIO_GRAPH.md.

```python
def check_drift(records, graph_path):
    if not os.path.exists(graph_path):
        return [{"kind": "NO_GRAPH", "severity": "INFO",
                 "detail": "No previous graph found. Run Foliograph to build one."}]
    prev_text = open(graph_path, encoding="utf-8").read()
    drift = []
    prev_words = {}
    for m in re.finditer(r'### `([^`]+)`.*?\*\*Words:\*\*\s*([\d,]+)',
                         prev_text, re.DOTALL):
        prev_words[m.group(1)] = int(m.group(2).replace(",", ""))
    current_files = {r["filename"] for r in records}
    for fname in current_files - set(prev_words.keys()):
        drift.append({"kind": "NEW_FILE", "filename": fname, "severity": "INFO",
                      "detail": "New document not in previous graph."})
    for fname in set(prev_words.keys()) - current_files:
        drift.append({"kind": "REMOVED_FILE", "filename": fname,
                      "severity": "BREAKING",
                      "detail": "Was in graph but no longer present."})
    for rec in records:
        if rec["filename"] not in prev_words:
            continue
        prev_w = prev_words[rec["filename"]]
        curr_w = rec["total_words"]
        if prev_w > 0 and abs(curr_w - prev_w) / prev_w > 0.15:
            direction = "grew" if curr_w > prev_w else "shrank"
            drift.append({"kind": "WORD_COUNT", "filename": rec["filename"],
                          "severity": "WARNING",
                          "detail": f"Word count {direction} by "
                                    f"{abs(curr_w-prev_w)/prev_w:.0%}: "
                                    f"{prev_w:,} → {curr_w:,}. Rebuild recommended."})
    return drift


def render_drift_plain(drift_items):
    """Plain English drift report for non-technical users."""
    if not drift_items:
        return ("Your graph is up to date. No changes detected.\n"
                "You do not need to rebuild.")
    lines = ["Here is what changed since your last graph was built:", ""]
    for item in drift_items:
        sev = {"INFO": "ℹ", "WARNING": "⚠", "BREAKING": "✗"}.get(item["severity"], "•")
        lines.append(f"{sev} {item.get('filename', '')}: {item['detail']}")
    has_warn = any(i["severity"] in ("WARNING", "BREAKING") for i in drift_items)
    lines.append("")
    if has_warn:
        lines.append("Recommendation: rebuild the graph by uploading your "
                     "files and saying 'Foliograph this'.")
    else:
        lines.append("No rebuild needed. Your graph is still accurate.")
    return "\n".join(lines)
```

---

## Feature 6: Fetch a specific section on demand

Trigger: user says "Load [file] § [Section Title]".

```python
def fetch_section(records, filename, section_title):
    rec = next((r for r in records if r["filename"] == filename), None)
    if not rec:
        return f"File not found: {filename}"
    title_lower = section_title.lower().strip()
    matched = None
    for i, sec in enumerate(rec["sections"]):
        if sec["title"].lower().strip() == title_lower:
            matched = i
            break
        if title_lower in sec["title"].lower():
            matched = i
            break
    if matched is None:
        available = [s["title"] for s in rec["sections"] if s["level"] > 0]
        return (f"Section not found: '{section_title}'\n\nAvailable sections:\n"
                + "\n".join(f"  - {t}" for t in available))
    raw = rec["raw_text"]
    lines = raw.splitlines()
    target = rec["sections"][matched]
    start = None
    for i, line in enumerate(lines):
        if target["title"].lower() in line.lower():
            start = i
            break
    if start is None:
        return f"## {target['title']}\n\n{target['summary']}"
    end = len(lines)
    for i in range(start + 1, len(lines)):
        if re.match(r'^#{1,' + str(target["level"]) + r'}\s+', lines[i]):
            end = i
            break
    return "\n".join(lines[start:end])
```

---

## Full workflow (runs on every "foliograph this" trigger)

```python
# 1. Setup
import os, sys, subprocess, re, json
from datetime import datetime, timezone
for pkg in ["python-docx", "python-pptx", "pdfminer.six"]:
    subprocess.run([sys.executable, "-m", "pip", "install", pkg,
                    "--break-system-packages", "-q"], capture_output=True)

upload_dir = "/mnt/user-data/uploads"
output_dir = "/mnt/user-data/outputs"
os.makedirs(output_dir, exist_ok=True)

SUPPORTED = {".docx", ".pdf", ".pptx", ".md", ".txt"}
EXTRACTORS = {".docx": extract_docx, ".pptx": extract_pptx,
              ".pdf": extract_pdf, ".md": extract_md, ".txt": extract_md}

uploaded = sorted([
    f for f in os.listdir(upload_dir)
    if os.path.splitext(f)[1].lower() in SUPPORTED
])

if not uploaded:
    print("No supported files found.")
    print("Upload .docx, .pdf, .pptx, .md, or .txt files and say 'Foliograph this'.")
else:
    # 2. Extract
    records = []
    for fname in uploaded:
        path = os.path.join(upload_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        try:
            rec = EXTRACTORS[ext](path)
            records.append(rec)
            print(f"  Extracted: {fname} ({rec['total_words']:,} words, "
                  f"{len(rec['sections'])} sections)")
        except Exception as e:
            print(f"  Skipped: {fname} ({e})")

    if records:
        # 3. Build
        project_name = "My Project"
        outputs = build_graph(records, project_name)

        # 4. Write files
        for filename, content in outputs.items():
            if filename.startswith("_"):
                continue
            with open(os.path.join(output_dir, filename), "w", encoding="utf-8") as f:
                f.write(content)

        # 5. Report savings in plain English
        saved = outputs["_saved_per_session"]
        full = outputs["_full_tokens"]
        graph = outputs["_graph_tokens"]
        ratio = round(full / max(graph, 1))

        print(f"""
Foliograph complete.

Your graph covers {len(records)} document(s) and {sum(r['total_words'] for r in records):,} words.

Token savings per session:
  Without Foliograph: ~{full:,} tokens (loading everything)
  With Foliograph:    ~{graph + 800:,} tokens (graph + 3 sections)
  Saved per session:  ~{saved:,} tokens ({ratio}x more efficient)

6 files are ready to download:
  FOLIO_GRAPH.md               : your document map
  FOLIO_INDEX.md               : concept lookup index
  FOLIO_RELATIONS.json         : document connections
  FOLIO_SESSION.md             : paste this to start each session
  FOLIO_PROJECT_INSTRUCTIONS.md: paste this into Project settings once
  FOLIO_TIPS.md                : plain English guide to saving more tokens

Recommended next step:
  Open FOLIO_PROJECT_INSTRUCTIONS.md and follow the 5-step instructions
  to add it to your Claude Project. After that, you will never need to
  paste the session starter again.
""")
```

---

## Output files reference

| File | Purpose | Who needs it |
|---|---|---|
| `FOLIO_GRAPH.md` | Structural map of all documents | Claude reads this |
| `FOLIO_INDEX.md` | Concept → location index | Claude reads this |
| `FOLIO_RELATIONS.json` | Document connections (machine-readable) | Advanced users / visualizers |
| `FOLIO_SESSION.md` | Session starter prompt | Paste at start of each session |
| `FOLIO_PROJECT_INSTRUCTIONS.md` | Permanent project rules | Paste into Project settings once |
| `FOLIO_TIPS.md` | Plain English savings guide | Read this first |

---

## Supported file types

| Format | Extension |
|---|---|
| Word Document | `.docx` |
| PDF | `.pdf` |
| PowerPoint | `.pptx` |
| Markdown | `.md` |
| Plain Text | `.txt` |

---

## Attribution

Foliograph skill for claude.ai, created by Prasad MK
Research: https://ssrn.com/author=10270516
Inspired by Graphify: https://github.com/safishamsi/graphify
Standalone Python package: https://github.com/prasad-m-k/foliograph
