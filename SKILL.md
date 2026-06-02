---
name: foliograph
description: >
  Use this skill when the user uploads documents (.docx, .pdf, .pptx, .xlsx, .md, .txt)
  and wants to work with them efficiently, or says any of: "foliograph this",
  "build the graph", "map this document", "index this", "reduce my token cost",
  "split my document", "save my session summary", "check for drift",
  "load [file] > [section]", "show me a dashboard", "executive report",
  "html report", "cumulative savings", or "what commands can I use".
  Designed for ALL users including executives and non-technical professionals.
  Default output is FOLIO_TIPS.md only. HTML dashboard is generated only when
  the user explicitly requests it. All processing runs in concealed mode by
  default. No technical knowledge required.
---

# Foliograph Skill
## Smart Document Navigation for claude.ai

**Created by:** Prasad MK
**Research:** https://ssrn.com/author=10270516
**GitHub:** https://github.com/prasad-m-k/foliograph

Inspired by Graphify (https://github.com/safishamsi/graphify) for IDE environments.
Foliograph brings the same approach to claude.ai chat with no terminal,
no installation, and no setup beyond uploading this skill file.

> Upload your documents. Say "foliograph this". Done.
> Ask for a dashboard if you want one.

---

## Concealed mode (default)

All extraction and processing runs silently inside a collapsed section.
Wrap all code execution output in a details block:

```
<details>
<summary>Foliograph processing log (click to expand)</summary>

[all print/status output goes here]

</details>
```

Only the final summary and output files appear in the main chat.
If the user says "show me what's happening" or "verbose mode",
run without concealment for that session.

---

## How this skill behaves

When a user uploads documents and triggers foliograph:

1. Run all extraction silently inside a concealed details block
2. Build the internal knowledge graph
3. Apply Graphify-parity analysis: god nodes, surprising connections,
   confidence tags, and suggested questions
4. Default: present FOLIO_TIPS.md only
5. If user asks for dashboard / report / html / executive view:
   generate FOLIO_EXPLORER.html at runtime using the prompt in the
   "HTML report on demand" section below
6. Never present internal working files (FOLIO_GRAPH.md, FOLIO_INDEX.md,
   FOLIO_RELATIONS.json) as downloads unless explicitly requested

The skill carries all navigation rules internally.
Users do NOT need to configure Project instructions separately.

---

## Built-in navigation rules (always active)

- Check the internal index before loading any source file
- Load sections on demand only: "Load [file] > [Section Title]"
- Never re-read a section already discussed this session
- When user says "save my session summary": write a 150-200 word note
  covering decisions made and next steps, present as a download
- When user says "has the graph changed?" or "check for drift":
  run the drift check
- When user says "split my document": run the document splitter
- When user asks about a topic: check index first, load section only
  if the index summary is insufficient

---

## Environment setup

```python
import subprocess, sys, os, re, json, math
from datetime import datetime, timezone

for pkg in ["python-docx", "python-pptx", "pdfminer.six", "openpyxl"]:
    subprocess.run(
        [sys.executable, "-m", "pip", "install", pkg,
         "--break-system-packages", "-q"],
        capture_output=True
    )
```

---

## Extractors and entity extraction

```python
from pptx import Presentation
import docx as dx
from pdfminer.high_level import extract_text as pdf_extract

def _sec(lvl, ttl, body, hint=None):
    text = " ".join(body).strip() if isinstance(body, list) else body
    first = re.split(r"[.!?]", text)[0][:160] if text else ""
    return {"level": lvl, "title": ttl, "summary": first,
            "word_count": len(text.split()), "page_hint": hint}

def _rec(path, ftype, title, words, pages, sections, tables, raw):
    return {"filename": os.path.basename(path), "file_type": ftype,
            "title": title, "total_words": words, "total_pages": pages,
            "sections": sections, "tables": tables, "figures": [],
            "raw_text": raw}

def extract_pptx(path):
    prs = Presentation(path)
    secs, all_text = [], []
    for i, slide in enumerate(prs.slides, 1):
        txts = [t for shape in slide.shapes if shape.has_text_frame
                for p in shape.text_frame.paragraphs
                for t in [p.text.strip()] if t]
        if txts:
            body = " ".join(txts[1:])
            secs.append(_sec(1, txts[0], body, f"slide {i}"))
            all_text.extend(txts)
    raw = "\n".join(all_text)
    return _rec(path, "PPTX", secs[0]["title"] if secs else path,
                len(raw.split()), len(prs.slides), secs, [], raw)

def extract_docx(path):
    doc = dx.Document(path)
    secs, body, hdg, all_text = [], [], None, []
    def flush():
        secs.append(_sec(*(hdg if hdg else (0, "(preamble)")), body))
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t: continue
        all_text.append(t)
        m = re.match(r"Heading (\d)", p.style.name if p.style else "")
        if m: flush(); hdg = (int(m.group(1)), t); body = []
        else: body.append(t)
    flush()
    tables = [" | ".join(c.text.strip() for c in tbl.rows[0].cells)[:120]
              for tbl in doc.tables if tbl.rows][:20]
    raw = "\n".join(all_text)
    title = next((s["title"] for s in secs if s["level"] == 1),
                 os.path.basename(path))
    return _rec(path, "DOCX", title, len(raw.split()), None, secs, tables, raw)

def extract_pdf(path):
    raw = pdf_extract(path) or ""
    def is_h(l):
        l = l.strip()
        return bool(l and len(l) < 120 and (
            re.match(r"^(Chapter|Section|CHAPTER|SECTION|\d+\.)\s", l) or
            re.match(r"^#{1,4}\s", l) or
            (len(l) < 80 and l == l.title() and not l.endswith("."))))
    secs, body, hdg = [], [], None
    def flush():
        secs.append(_sec(*(hdg if hdg else (0, "(preamble)")), body))
    for line in raw.splitlines():
        if is_h(line): flush(); hdg = (1, line.strip()); body = []
        else: body.append(line)
    flush()
    return _rec(path, "PDF", secs[0]["title"] if secs else path,
                len(raw.split()), None, secs, [], raw)

def extract_md(path):
    raw = open(path, encoding="utf-8", errors="replace").read()
    secs, body, hdg = [], [], None
    def flush():
        secs.append(_sec(*(hdg if hdg else (0, "(preamble)")), body))
    for line in raw.splitlines():
        m = re.match(r"^(#{1,4})\s+(.+)", line)
        if m: flush(); hdg = (len(m.group(1)), m.group(2).strip()); body = []
        else: body.append(line)
    flush()
    ext = os.path.splitext(path)[1].upper().lstrip(".")
    return _rec(path, ext, secs[0]["title"] if secs else path,
                len(raw.split()), None, secs, [], raw)

def extract_xml(path):
    """Extract from .xml including Office Open XML content files."""
    import xml.etree.ElementTree as ET
    raw_bytes = open(path, "rb").read()
    raw = raw_bytes.decode("utf-8", errors="replace")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    is_ooxml = "openxmlformats.org" in raw[:2000] or "schemas.microsoft.com" in raw[:2000]
    try:
        root = ET.fromstring(raw.encode("utf-8"))
    except ET.ParseError:
        raw = re.sub(r"<[^>]+>", " ", raw)
        return extract_md.__wrapped__(path) if hasattr(extract_md, "__wrapped__") else                {"filename": os.path.basename(path), "file_type": "XML",
                "title": os.path.basename(path), "total_words": len(raw.split()),
                "total_pages": None, "sections": [], "tables": [], "figures": [],
                "raw_text": re.sub(r"\s+", " ", raw).strip()}
    all_text, sections, body, heading = [], [], [], None
    def flush():
        text = " ".join(body).strip()
        if heading or text:
            lvl, ttl = heading if heading else (0, "(preamble)")
            first = re.split(r"[.!?]", text)[0][:160] if text else ""
            sections.append({"level": lvl, "title": ttl, "summary": first,
                              "word_count": len(text.split()), "page_hint": None})
    if is_ooxml:
        paras = list(root.iter(f"{{{W}}}p"))
        for para in paras:
            style_elem = para.find(f".//{{{W}}}pStyle")
            style_val = (style_elem.get(f"{{{W}}}val","") if style_elem is not None else "")
            hm = re.match(r"[Hh]eading\s*(\d)", style_val)
            lvl = int(hm.group(1)) if hm else 0
            runs = ["".join(t.text or "" for t in para.iter(f"{{{W}}}t"))]
            para_text = "".join(runs).strip()
            if para_text:
                all_text.append(para_text)
                if lvl > 0:
                    flush(); heading_box = [lvl, para_text]; body.clear()
                    heading = tuple(heading_box)
                else:
                    body.append(para_text)
    else:
        for elem in root.iter():
            t = (elem.text or "").strip()
            if t: all_text.append(t); body.append(t)
    flush()
    raw_out = " ".join(all_text)
    return {"filename": os.path.basename(path), "file_type": "XML",
            "title": sections[0]["title"] if sections and sections[0]["level"] > 0
                     else os.path.basename(path),
            "total_words": len(raw_out.split()), "total_pages": None,
            "sections": sections or [{"level":0,"title":"(content)",
                "summary": raw_out[:160], "word_count": len(raw_out.split()),
                "page_hint": None}],
            "tables": [], "figures": [], "raw_text": raw_out}

def extract_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    secs, tables, all_text = [], [], []
    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            cells = [str(c).strip() for c in row
                     if c is not None and str(c).strip() not in ("", "None")]
            if cells:
                rows.append(cells)
        if not rows:
            continue
        headers = rows[0]
        tables.append((f"{name}: " + " | ".join(headers[:10]))[:120])
        sheet_text = " ".join(c for r in rows for c in r)
        all_text.append(sheet_text)
        first = re.split(r"[.!?]", sheet_text)[0][:160] if sheet_text else ""
        secs.append(_sec(1, name, first or
                         f"{len(rows)} rows × {max(len(r) for r in rows)} cols",
                         f"sheet: {name}"))
    wb.close()
    raw = " ".join(all_text)
    return _rec(path, "XLSX",
                os.path.splitext(os.path.basename(path))[0].replace("_", " ").title(),
                len(raw.split()), None, secs or
                [_sec(0, "(empty workbook)", "")], tables, raw)

EXTRACTORS = {".docx": extract_docx, ".pptx": extract_pptx,
              ".pdf": extract_pdf, ".xlsx": extract_xlsx,
              ".md": extract_md, ".txt": extract_md,
              ".xml": extract_xml}

def extract_entities(text):
    text = re.sub(r"\s+", " ", text)
    ents = set()
    for m in re.finditer(r"\b([A-Z]{2,6})\b", text):
        ents.add(m.group(1))
    for m in re.finditer(
        r"\b((?:[A-Z][a-z]+ ){1,4}(?:Model|Framework|System|Method|"
        r"Score|Rate|Loop|Matrix|Protocol|Index|Threshold|Spectrum))\b", text):
        ents.add(m.group(1).strip())
    for m in re.finditer(r"\b((?:[A-Z][a-z]+\s){1,3}[A-Z][a-z]+)\b", text):
        p = m.group(1).strip()
        if 2 <= len(p.split()) <= 4: ents.add(p)
    stops = {"The Book", "This Chapter", "In Practice", "For Example",
             "As A", "In My", "At The", "Of The"}
    return sorted(e for e in ents if e not in stops and "|" not in e
                  and not re.match(r"^\d", e) and len(e) > 2)[:40]
```

## Graphify-parity analysis

Applies the same analytical depth as Graphify's GRAPH_REPORT.md.
Run this after extraction to enrich the internal graph.

```python
def analyse_graph(records):
    """
    Returns enriched graph data with:
    - god_nodes:             most-connected concepts (everything flows through these)
    - surprising_connections: cross-document links ranked by unexpectedness
    - confidence_tags:        EXTRACTED | INFERRED | AMBIGUOUS per relationship
    - suggested_questions:    4-5 questions the graph is uniquely positioned to answer
    """
    # Build concept frequency map across all documents
    concept_freq = {}
    concept_docs  = {}
    for rec in records:
        ents = extract_entities(rec["raw_text"])
        for ent in ents:
            concept_freq[ent] = concept_freq.get(ent, 0) + rec["raw_text"].lower().count(ent.lower())
            concept_docs.setdefault(ent, set()).add(rec["filename"])

    # God nodes: top concepts by frequency x document spread
    god_scores = {
        e: concept_freq[e] * len(concept_docs[e])
        for e in concept_freq
    }
    god_nodes = sorted(god_scores, key=lambda x: -god_scores[x])[:8]

    # Relationships with confidence tags
    rels = []
    entity_files = {}
    for rec in records:
        for ent in extract_entities(rec["raw_text"]):
            entity_files.setdefault(ent.lower(), [])
            if rec["filename"] not in entity_files[ent.lower()]:
                entity_files[ent.lower()].append(rec["filename"])

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
        # Confidence: EXTRACTED if concept appears in section title,
        # INFERRED if body only, AMBIGUOUS if very short concept name
        confidence = "EXTRACTED" if any(len(c) > 6 for c in concepts) else \
                     "INFERRED"  if any(len(c) > 3 for c in concepts) else \
                     "AMBIGUOUS"
        rels.append({
            "kind": "SHARES_CONCEPT",
            "source_file": fa, "target_file": fb,
            "label": "Shared: " + ", ".join(concepts),
            "confidence": confidence
        })

    # Surprising connections: cross-doc links where the shared concept
    # is domain-specific (longer, less common)
    surprising = sorted(
        [r for r in rels if r["confidence"] == "EXTRACTED"],
        key=lambda r: -max(len(c) for c in r["label"].replace("Shared: ", "").split(", "))
    )[:5]

    # Sequence relationships
    patterns = [(r'chapter\s+(\d+)', 'Chapter'), (r'slide\s+(\d+)', 'Slide'),
                (r'part\s+(\d+)', 'Part'), (r'section\s+(\d+)', 'Section')]
    for rec in records:
        for sec in rec["sections"]:
            for pattern, kind in patterns:
                m = re.search(pattern, sec["title"].lower())
                if not m: continue
                try: next_n = str(int(m.group(1)) + 1)
                except ValueError: continue
                for other_rec in records:
                    for other_sec in other_rec["sections"]:
                        if re.search(pattern.replace(r'(\d+)', re.escape(next_n)),
                                     other_sec["title"].lower()):
                            rels.append({
                                "kind": "SEQUENCE",
                                "source_file": rec["filename"],
                                "source_section": sec["title"],
                                "target_file": other_rec["filename"],
                                "target_section": other_sec["title"],
                                "label": f"{kind} sequence",
                                "confidence": "EXTRACTED"
                            })

    # Suggested questions: based on god nodes and document types
    doc_types = list({r["file_type"] for r in records})
    suggested_questions = []
    if god_nodes:
        suggested_questions.append(
            f"What is the relationship between {god_nodes[0]} and "
            f"{god_nodes[1] if len(god_nodes) > 1 else 'the main topic'}?"
        )
    if len(records) > 1:
        suggested_questions.append(
            f"Where do {records[0]['filename']} and {records[1]['filename']} overlap?"
        )
    suggested_questions.append(
        "Which sections should I read first to understand the core argument?"
    )
    suggested_questions.append(
        "What are the key terms I need to know before diving into the detail?"
    )
    if surprising:
        suggested_questions.append(
            f"Why does '{surprising[0]['label'].replace('Shared: ', '')}' appear across multiple documents?"
        )

    return {
        "god_nodes": god_nodes,
        "surprising_connections": surprising,
        "all_relationships": rels,
        "suggested_questions": suggested_questions[:5],
    }
```

---

## Internal graph builder

```python
def build_internal_graph(records, project_name):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    total_words = sum(r["total_words"] for r in records)
    analysis = analyse_graph(records)
    rels = analysis["all_relationships"]

    lines = [
        "# Foliograph Knowledge Graph", "",
        f"**Project:** {project_name}  ",
        f"**Built:** {now}  ",
        f"**Documents:** {len(records)}  ",
        f"**Total words:** {total_words:,}  ",
        f"**Relationships:** {len(rels)}  ", "",
        "---", "", "## Documents", ""
    ]

    for rec in records:
        pages = f" | {rec['total_pages']} pages" if rec.get("total_pages") else ""
        lines += [
            f"### `{rec['filename']}` [{rec['file_type']}]",
            f"**Title:** {rec['title']}  ",
            f"**Words:** {rec['total_words']:,}{pages}  ", ""
        ]
        for sec in [s for s in rec["sections"] if s["level"] > 0]:
            indent = "  " * (sec["level"] - 1)
            loc = f" [{sec['page_hint']}]" if sec.get("page_hint") else ""
            wc = f" ({sec['word_count']}w)" if sec["word_count"] > 20 else ""
            lines.append(f"{indent}- **{sec['title']}**{loc}{wc}")
            if sec["summary"]:
                lines.append(f"{'  ' * sec['level']}  > {sec['summary']}")
        lines.append("")
        entities = extract_entities(rec["raw_text"])
        if entities:
            lines.append(f"**Key Terms:** {', '.join(entities[:25])}")
            lines.append("")
        lines += ["---", ""]

    # Relationships with confidence tags
    if rels:
        lines += ["## Relationships", ""]
        by_kind = {}
        for r in rels:
            by_kind.setdefault(r["kind"], []).append(r)
        for kind, label in [
            ("SEQUENCE", "Sequence (reading order)"),
            ("SHARES_CONCEPT", "Shared Concepts")
        ]:
            if kind not in by_kind:
                continue
            lines += [f"### {label}", ""]
            for r in by_kind[kind]:
                src = f"`{r['source_file']}`" + (f" > *{r.get('source_section','')}*" if r.get("source_section") else "")
                tgt = f"`{r['target_file']}`" + (f" > *{r.get('target_section','')}*" if r.get("target_section") else "")
                conf = r.get("confidence", "EXTRACTED")
                lines.append(f"- {src} → {tgt}  `[{conf}]`")
                lines.append(f"  _{r['label']}_")
            lines.append("")

    graph_md = "\n".join(lines)

    # Build index entries
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
            loc = sec.get("page_hint") or f"> {sec['title']}"
            entries.append((sec["title"], rec["filename"], loc))
        for ent in extract_entities(rec["raw_text"]):
            key = ent.lower()
            if key in seen:
                continue
            seen.add(key)
            location = rec["filename"]
            for sec in rec["sections"]:
                if ent.lower() in sec["title"].lower() or ent.lower() in sec["summary"].lower():
                    location = f"{rec['filename']} > {sec['title']}"
                    break
            entries.append((ent, rec["filename"], location))
    entries.sort(key=lambda x: x[0].lower())

    return graph_md, entries, analysis, total_words
```

---

## FOLIO_TIPS.md builder (default output)

```python
def build_tips_md(records, project_name, total_words, analysis):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    full_tokens  = int(total_words * 1.3)
    graph_tokens = max(400, int(total_words * 0.15))
    saved        = max(0, full_tokens - graph_tokens - 800)
    file_list    = "\n".join(f"  - {r['filename']}" for r in records)

    load_cmds = "\n".join(
        f"  Load {r['filename']} > {s['title']}"
        for r in records
        for s in r["sections"]
        if s["level"] == 1 and s["title"]
    )[:8 * 60]  # cap at ~8 commands

    god_nodes_str = ", ".join(analysis["god_nodes"][:5]) if analysis["god_nodes"] else "none detected"

    surprising_str = ""
    for sc in analysis["surprising_connections"][:3]:
        surprising_str += f"\n  - {sc['label'].replace('Shared: ','')} ({sc['source_file']} + {sc['target_file']})"

    questions_str = "\n".join(
        f"  {i+1}. {q}"
        for i, q in enumerate(analysis["suggested_questions"])
    )

    return f"""# Foliograph Tips
## {project_name}
Built: {now}

---

## What just happened

Foliograph analysed your documents and built a knowledge map.
All working files are held internally. Nothing was sent outside this session.

Documents in this project:
{file_list}

---

## Token savings

| | Tokens |
|---|---|
| Loading all documents every session | ~{full_tokens:,} |
| With Foliograph (graph + typical sections) | ~{graph_tokens + 800:,} |
| Saved per session | ~{saved:,} |
| Saved over 10 sessions | ~{saved * 10:,} |

---

## Key concepts in your documents

**Central concepts (god nodes):** {god_nodes_str}
These appear most frequently across your content. Good starting points.

**Surprising connections:**{surprising_str if surprising_str else chr(10) + "  None detected (single document or sparse overlap)."}

**Suggested questions to ask:**
{questions_str}

---

## Commands for your documents

Load a specific section (paste into Claude):
{load_cmds}

Session commands:
  Save my session summary
  Has the graph changed?
  What does the index say about [topic]?
  Show me an executive dashboard
  Show me cumulative savings over 12 months

---

## 7 habits that reduce token cost

1. Ask the index first: "What does the index say about X?"
2. Load by section name, not the whole file
3. End sessions with: "Save my session summary"
4. For one-chapter sessions: "Build a micro-graph for [file] only"
5. Rebuild only when content changes significantly
6. Share FOLIO_TIPS.md with colleagues as a project map
7. Ask for a dashboard only when presenting to others

---

Foliograph, created by Prasad MK
Research: https://ssrn.com/author=10270516
GitHub: https://github.com/prasad-m-k/foliograph
"""
```

---

## HTML report on demand

When the user asks for a dashboard, report, executive view, or html output,
generate a self-contained HTML file at runtime using this prompt as the guide.
Do NOT use a pre-written template. Generate the HTML fresh each time.

**Prompt to follow when generating the HTML:**

Generate a single self-contained executive dashboard HTML file with:

Colors: Primary #002060 (headers, titles, nav), Accent #E07020 (borders,
highlights, active states), Body #1A1A1A on white, Callout background #F0F4FA,
Table rows alternating #F7F9FC and #FFFFFF.

Typography: system sans-serif for body, monospace for data labels and metrics.

Layout: sticky header with project name and live stats, three-column layout
(nav sidebar 200px, main content, right summary panel 260px), six navigation tabs
(Overview, Structure, Sections, Search, Savings, Commands).

Content to include:
- Overview: metric cards (documents, sections, index entries, efficiency ratio),
  word distribution bar chart per document
- Structure: expandable tree of document sections with word counts
- Sections: all sections as rows with word count bars, click-to-copy load command
- Search: live filter across all index entries, click-to-copy result
- Savings: slider controls (sessions/month, sections/session, words/section),
  breakdown table showing full-load vs Foliograph cost, results update live
- Commands: copyable command buttons for all document sections plus session commands

No emojis anywhere. Use solid diamond (filled triangle right) as nav markers.
Navy header with orange bottom border. Playfair Display for logo mark.
Click any section row or search result to copy the "Load file > Section" command.
Toast notification on copy. Animated orange pulse dot in header.

Inject the actual data: document list, section tree, index entries, token numbers,
project name, build date. Make sliders update the savings panel live with JavaScript.

---

## Cumulative savings chart (on demand)

When user asks for "cumulative savings", "savings over time", or "12-month chart",
generate an inline HTML snippet or canvas chart showing:

- X axis: 12 months
- Y axis: cumulative tokens saved
- Line: grows each month based on sessions-per-month x saved-per-session
- Annotate the inflection point where Foliograph pays back its one-time build cost
- Use orange (#E07020) line on white, navy (#002060) axis labels
- Include three scenarios: light use (5 sessions/month), typical (15), heavy (30)
- No chart library dependency: use SVG or inline canvas with vanilla JS

Present this as a short HTML file or an artifact the user can open.

---

## Drift check

Trigger: "has the graph changed?" or "check for drift"

```python
def check_drift(records, prev_graph_text):
    prev_words = {}
    for m in re.finditer(
        r'### `([^`]+)`.*?\*\*Words:\*\*\s*([\d,]+)',
        prev_graph_text, re.DOTALL
    ):
        prev_words[m.group(1)] = int(m.group(2).replace(",", ""))

    current_files = {r["filename"] for r in records}
    drift = []

    for fname in current_files - set(prev_words.keys()):
        drift.append({"severity": "INFO",   "detail": f"New document: {fname}"})
    for fname in set(prev_words.keys()) - current_files:
        drift.append({"severity": "BREAKING","detail": f"Removed: {fname}. Index entries may be stale."})

    for rec in records:
        if rec["filename"] not in prev_words:
            continue
        prev_w = prev_words[rec["filename"]]
        curr_w = rec["total_words"]
        if prev_w > 0 and abs(curr_w - prev_w) / prev_w > 0.15:
            direction = "grew" if curr_w > prev_w else "shrank"
            drift.append({
                "severity": "WARNING",
                "detail": (
                    f"{rec['filename']}: word count {direction} by "
                    f"{abs(curr_w-prev_w)/prev_w:.0%} "
                    f"({prev_w:,} to {curr_w:,}). Rebuild recommended."
                )
            })

    if not drift:
        return "Graph is current. No significant changes detected."

    lines = ["Changes since last build:", ""]
    for item in drift:
        sev = {"INFO": "i", "WARNING": "!", "BREAKING": "X"}.get(item["severity"], "-")
        lines.append(f"  [{sev}] {item['detail']}")
    lines += ["", "Run 'foliograph this' with your updated files to rebuild."]
    return "\n".join(lines)
```

---

## Document splitter (on demand)

Trigger: "split my document into chapters"

For .docx: split on Heading 1 boundaries, save each as a separate file.
For .md: split on # headings.
Present split files as downloads and tell the user:
"Upload individual chapter files next time instead of the whole document.
Each section fetch will be 10x smaller."

---

## Full workflow

```python
import os, sys, subprocess, re, json, math
from datetime import datetime, timezone

# Setup (concealed)
for pkg in ["python-docx", "python-pptx", "pdfminer.six"]:
    subprocess.run([sys.executable, "-m", "pip", "install", pkg,
                    "--break-system-packages", "-q"], capture_output=True)

upload_dir = "/mnt/user-data/uploads"
output_dir = "/mnt/user-data/outputs"
os.makedirs(output_dir, exist_ok=True)

SUPPORTED = {".docx", ".pdf", ".pptx", ".md", ".txt", ".xml"}

uploaded = sorted([
    f for f in os.listdir(upload_dir)
    if os.path.splitext(f)[1].lower() in SUPPORTED
    and not f.startswith("FOLIO_")
    and not f.startswith("SKILL")
])

if not uploaded:
    print("No supported files found.")
    print("Upload .docx, .pdf, .pptx, .md, or .txt files.")
else:
    records = []
    for fname in uploaded:
        path = os.path.join(upload_dir, fname)
        ext  = os.path.splitext(fname)[1].lower()
        try:
            rec = EXTRACTORS[ext](path)
            records.append(rec)
            print(f"  Extracted: {fname} ({rec['total_words']:,} words, "
                  f"{len(rec['sections'])} sections)")
        except Exception as e:
            print(f"  Skipped: {fname} ({e})")

    if records:
        project_name = (
            os.path.splitext(records[0]["filename"])[0]
            .replace("_", " ").replace("-", " ").title()
        ) if len(records) == 1 else "My Project"

        graph_md, entries, analysis, total_words = \
            build_internal_graph(records, project_name)

        full_tokens  = int(total_words * 1.3)
        graph_tokens = max(400, len(graph_md) // 4)
        saved        = max(0, full_tokens - graph_tokens - 800)
        ratio        = max(1, round(full_tokens / max(graph_tokens + 800, 1)))

        tips_md = build_tips_md(records, project_name, total_words, analysis)

        tips_path = os.path.join(output_dir, "FOLIO_TIPS.md")
        with open(tips_path, "w", encoding="utf-8") as f:
            f.write(tips_md)

        # Plain-English summary (shown in main chat, outside concealed block)
        god_preview = ", ".join(analysis["god_nodes"][:3]) if analysis["god_nodes"] else "none detected"

        print(f"""
Foliograph complete.

{len(records)} document(s), {total_words:,} words, {len(entries)} index entries.

Central concepts: {god_preview}
Relationships found: {len(analysis["all_relationships"])} ({sum(1 for r in analysis["all_relationships"] if r.get("confidence") == "EXTRACTED")} extracted, {sum(1 for r in analysis["all_relationships"] if r.get("confidence") == "INFERRED")} inferred)

Token efficiency: {ratio}x more efficient than loading full documents.
Saved per session: ~{saved:,} tokens.

FOLIO_TIPS.md is ready with your commands, key concepts, and savings breakdown.

To get a visual dashboard, say: "Show me an executive dashboard"
To see cumulative savings, say: "Show me cumulative savings over 12 months"
""")
```

---

## What users see vs what Claude uses internally

Default outputs (presented to user):
  FOLIO_TIPS.md              plain-English guide, commands, savings

On request only:
  FOLIO_EXPLORER.html        executive dashboard (generated at runtime)
  FOLIO_CUMULATIVE.html      12-month savings chart (generated at runtime)
  FOLIO_SESSION_SUMMARY.md   when user says "save my session summary"

Internal only (never presented as downloads unless explicitly asked):
  FOLIO_GRAPH.md             structural map Claude uses for navigation
  FOLIO_INDEX.md             concept lookup index
  FOLIO_RELATIONS.json       relationship edges with confidence tags

---

## Supported file types

| Format | Extension |
|---|---|
| Word Document | .docx |
| PDF | .pdf |
| PowerPoint | .pptx |
| Markdown | .md |
| Plain Text | .txt |
| XML | .xml |

---

## Attribution

Foliograph skill for claude.ai, created by Prasad MK
Research: https://ssrn.com/author=10270516
GitHub: https://github.com/prasad-m-k/foliograph
Inspired by Graphify: https://github.com/safishamsi/graphify
