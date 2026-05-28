"""
foliograph.extractor
~~~~~~~~~~~~~~~~~~~~
Extract structured content from office documents.
Supports: .docx, .pdf, .pptx, .md, .txt, .xml
"""

from __future__ import annotations

import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Section:
    """A logical section extracted from a document."""
    level: int          # heading depth (1=H1, 2=H2, etc.; 0=body)
    title: str          # heading text
    summary: str        # first sentence / first ~150 chars of body
    word_count: int
    page_hint: Optional[str] = None   # "p.12" or "slide 4" etc.
    tags: list[str] = field(default_factory=list)


@dataclass
class DocumentRecord:
    """All extracted metadata for one document."""
    path: Path
    file_type: str          # docx | pdf | pptx | md | txt
    title: str
    total_words: int
    total_pages: Optional[int]
    sections: list[Section]
    tables: list[str]       # table captions / first-row headers
    figures: list[str]      # figure captions
    named_entities: list[str]  # key terms, frameworks, proper nouns
    raw_text: str           # full text (used during build, not written to graph)


# ---------------------------------------------------------------------------
# Extraction helpers
# ---------------------------------------------------------------------------

def _run(cmd: list[str], input_text: str | None = None) -> str:
    """Run a shell command, return stdout."""
    result = subprocess.run(
        cmd,
        input=input_text,
        capture_output=True,
        text=True,
        timeout=60,
    )
    return result.stdout or ""


def _first_sentence(text: str, max_chars: int = 160) -> str:
    """Return the first meaningful sentence from a text block."""
    text = text.strip()
    if not text:
        return ""
    # Try sentence boundary
    m = re.search(r"[.!?]", text)
    if m and m.start() < max_chars:
        return text[: m.start() + 1].strip()
    return text[:max_chars].rstrip() + ("…" if len(text) > max_chars else "")


def _count_words(text: str) -> int:
    return len(text.split())


def _extract_named_entities(text: str) -> list[str]:
    """
    Lightweight NE extraction without spaCy dependency.
    Finds: TitleCase multi-word phrases, ALL-CAPS acronyms,
    quoted terms, and 'the X Model/Framework/Method' patterns.
    """
    # Normalize whitespace so newlines don't split phrases
    text = re.sub(r'\s+', ' ', text)
    entities: set[str] = set()

    # ALL-CAPS acronyms (2-6 chars)
    for m in re.finditer(r"\b([A-Z]{2,6})\b", text):
        entities.add(m.group(1))

    # "The X Model / Framework / System / Method / Score / Rate"
    for m in re.finditer(
        r"\b((?:[A-Z][a-z]+ ){1,4}(?:Model|Framework|System|Method|Score|"
        r"Rate|Loop|Matrix|Staircase|Protocol|Index|Threshold|Spectrum))\b",
        text,
    ):
        entities.add(m.group(1).strip())

    # TitleCase 2-4 word phrases (likely proper nouns / named concepts)
    for m in re.finditer(r"\b((?:[A-Z][a-z]+\s){1,3}[A-Z][a-z]+)\b", text):
        phrase = m.group(1).strip()
        if 2 <= len(phrase.split()) <= 4:
            entities.add(phrase)

    # Quoted terms
    for m in re.finditer(r'["\u201c\u201d]([^"\u201c\u201d]{3,40})["\u201c\u201d]', text):
        entities.add(m.group(1).strip())

    # Remove very common stop phrases
    stopwords = {
        "The Book", "This Chapter", "In Practice", "Key Insight",
        "For Example", "As A", "In My", "At The", "Of The",
    }
    return sorted(
        e for e in entities
        if e not in stopwords
        and "|" not in e
        and not re.match(r"^\d", e)
        and len(e) > 2
    )[:60]


# ---------------------------------------------------------------------------
# Per-format extractors
# ---------------------------------------------------------------------------

def _extract_md_txt(path: Path) -> DocumentRecord:
    raw = path.read_text(encoding="utf-8", errors="replace")
    lines = raw.splitlines()
    sections: list[Section] = []
    tables: list[str] = []
    figures: list[str] = []
    current_body: list[str] = []
    current_heading: tuple[int, str] | None = None

    def flush(heading, body_lines):
        body = " ".join(body_lines).strip()
        if heading:
            lvl, ttl = heading
        else:
            lvl, ttl = 0, "(preamble)"
        sections.append(Section(
            level=lvl,
            title=ttl,
            summary=_first_sentence(body),
            word_count=_count_words(body),
        ))

    for line in lines:
        hm = re.match(r"^(#{1,4})\s+(.+)", line)
        if hm:
            if current_body or current_heading:
                flush(current_heading, current_body)
            current_heading = (len(hm.group(1)), hm.group(2).strip())
            current_body = []
        else:
            current_body.append(line)
            # detect markdown tables
            if re.match(r"^\s*\|", line) and "---" not in line:
                cell = re.sub(r"\s*\|\s*", " | ", line).strip(" |")
                if cell and cell not in tables:
                    tables.append(cell[:120])
            # detect figure references
            if re.match(r"!\[", line):
                cap = re.search(r"!\[([^\]]+)\]", line)
                if cap:
                    figures.append(cap.group(1)[:120])

    flush(current_heading, current_body)

    title = sections[0].title if sections else path.stem
    return DocumentRecord(
        path=path,
        file_type=path.suffix.lstrip("."),
        title=title,
        total_words=_count_words(raw),
        total_pages=None,
        sections=sections,
        tables=tables[:30],
        figures=figures[:30],
        named_entities=_extract_named_entities(raw),
        raw_text=raw,
    )


def _extract_docx(path: Path) -> DocumentRecord:
    # extract-text is available in the Claude environment
    raw = _run(["extract-text", str(path)])
    if not raw.strip():
        # fallback: python-docx
        try:
            import docx as dx
            doc = dx.Document(str(path))
            raw = "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            raw = ""

    # Parse markdown output from extract-text (headings become # lines)
    return _extract_md_txt_from_string(raw, path, "docx")


def _extract_pdf(path: Path) -> DocumentRecord:
    raw = _run(["pdftotext", "-layout", str(path), "-"])
    if not raw.strip():
        raw = _run(["extract-text", str(path)])

    # Get page count
    info = _run(["pdfinfo", str(path)])
    pages = None
    m = re.search(r"Pages:\s+(\d+)", info)
    if m:
        pages = int(m.group(1))

    rec = _extract_md_txt_from_string(raw, path, "pdf")
    rec.total_pages = pages
    return rec


def _extract_pptx(path: Path) -> DocumentRecord:
    raw = _run(["extract-text", str(path)])
    if not raw.strip():
        try:
            from pptx import Presentation
            prs = Presentation(str(path))
            lines = []
            for i, slide in enumerate(prs.slides, 1):
                lines.append(f"## Slide {i}")
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        lines.append(shape.text_frame.text)
            raw = "\n".join(lines)
        except Exception:
            raw = ""

    rec = _extract_md_txt_from_string(raw, path, "pptx")

    # Re-label sections with slide numbers for pptx
    slide_n = 0
    for sec in rec.sections:
        if re.match(r"slide\s+\d+", sec.title.lower()):
            slide_n += 1
            sec.page_hint = f"slide {slide_n}"
        sec.level = max(sec.level, 1)

    return rec


def _extract_md_txt_from_string(
    raw: str, path: Path, file_type: str
) -> DocumentRecord:
    """Parse already-extracted text using the markdown extractor logic."""
    # Write to temp file and reuse _extract_md_txt
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".md", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(raw)
        tmp_path = Path(tmp.name)

    rec = _extract_md_txt(tmp_path)
    tmp_path.unlink(missing_ok=True)

    rec.path = path
    rec.file_type = file_type
    rec.title = path.stem.replace("_", " ").replace("-", " ").title()

    # Try to pull a real title from first H1
    for sec in rec.sections:
        if sec.level == 1 and sec.title not in ("(preamble)",):
            rec.title = sec.title
            break

    # Tables and figures from raw text
    tables = []
    figures = []
    for line in raw.splitlines():
        if re.match(r"\*?\*?Table\s+\d", line):
            tables.append(re.sub(r"\*", "", line).strip()[:120])
        if re.match(r"\*?\*?Figure\s+\d", line):
            figures.append(re.sub(r"\*", "", line).strip()[:120])

    rec.tables = tables[:30]
    rec.figures = figures[:30]
    rec.named_entities = _extract_named_entities(raw)
    rec.raw_text = raw
    return rec




def _extract_xml(path: Path) -> DocumentRecord:
    """
    Extract structured content from an XML file.

    Handles three cases:

    1. Office Open XML content files (word/document.xml, xl/worksheets/*.xml,
       ppt/slides/*.xml) - strips namespaced tags, extracts text runs.

    2. Structured XML with heading-like elements (h1-h6, title, section,
       chapter, heading) - treats them as section boundaries.

    3. Generic XML - strips all tags, treats text nodes as flat content,
       segments into pseudo-sections on blank lines or top-level elements.
    """
    import xml.etree.ElementTree as ET

    raw_bytes = path.read_bytes()
    try:
        raw_text = raw_bytes.decode("utf-8", errors="replace")
    except Exception:
        raw_text = raw_bytes.decode("latin-1", errors="replace")

    # --- Detect Office Open XML by namespace prefix ---------------------
    is_ooxml = any(ns in raw_text[:2000] for ns in [
        "http://schemas.openxmlformats.org/wordprocessingml",
        "http://schemas.openxmlformats.org/spreadsheetml",
        "http://schemas.openxmlformats.org/presentationml",
        "http://schemas.openxmlformats.org/drawingml",
        "schemas.microsoft.com/office",
    ])

    if is_ooxml:
        return _extract_ooxml_content(path, raw_text)

    # --- Try structured XML (HTML-like or DocBook-like) -----------------
    try:
        tree = ET.fromstring(raw_text.encode("utf-8"))
    except ET.ParseError:
        # Fall back to treating as plain text
        return _extract_md_txt_from_string(
            re.sub(r"<[^>]+>", " ", raw_text), path, "xml"
        )

    HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6",
                    "title", "section", "chapter", "heading",
                    "topic", "part", "article"}

    def tag_local(elem):
        """Strip namespace from tag name."""
        tag = elem.tag
        return tag.split("}")[-1].lower() if "}" in tag else tag.lower()

    def all_text(elem):
        """Collect all text content recursively."""
        parts = []
        if elem.text:
            parts.append(elem.text.strip())
        for child in elem:
            parts.append(all_text(child))
            if child.tail:
                parts.append(child.tail.strip())
        return " ".join(p for p in parts if p)

    sections: list[Section] = []
    current_body: list[str] = []
    current_heading: tuple[int, str] | None = None
    all_raw_text: list[str] = []

    def flush(heading, body):
        text = " ".join(body).strip()
        if heading or text:
            lvl, ttl = heading if heading else (0, "(preamble)")
            first = re.split(r"[.!?]", text)[0][:160] if text else ""
            sections.append(Section(
                level=lvl, title=ttl,
                summary=first,
                word_count=len(text.split()),
                page_hint=None,
            ))

    def walk(elem, depth=0):
        local = tag_local(elem)
        if local in HEADING_TAGS:
            flush(current_heading, current_body)
            level = int(local[1]) if local.startswith("h") and len(local) == 2 else 1
            title = all_text(elem).strip()
            current_heading_box[0] = (level, title)
            current_body_box.clear()
        else:
            text = (elem.text or "").strip()
            if text:
                current_body_box.append(text)
                all_raw_text.append(text)
            if elem.tail:
                t = elem.tail.strip()
                if t:
                    current_body_box.append(t)
                    all_raw_text.append(t)
        for child in elem:
            walk(child, depth + 1)

    # Use mutable containers to allow flush() to see updates
    current_heading_box: list[tuple[int, str] | None] = [None]
    current_body_box: list[str] = []

    # Simpler iterative approach for generic XML
    for elem in tree.iter():
        local = tag_local(elem)
        text = (elem.text or "").strip()
        if text:
            all_raw_text.append(text)
        if elem.tail:
            t = elem.tail.strip()
            if t:
                all_raw_text.append(t)

    # Build sections from top-level children
    for child in tree:
        local = tag_local(child)
        child_text = all_text(child)
        if local in HEADING_TAGS:
            flush(current_heading, current_body)
            current_heading = (1, child_text[:80])
            current_body = []
        else:
            if child_text:
                current_body.append(child_text)

    flush(current_heading, current_body)

    raw = " ".join(all_raw_text)
    title = path.stem.replace("_", " ").replace("-", " ").title()
    if sections and sections[0].level == 1:
        title = sections[0].title

    return DocumentRecord(
        path=path,
        file_type="xml",
        title=title,
        total_words=len(raw.split()),
        total_pages=None,
        sections=sections if sections else [
            Section(level=0, title="(content)", summary=raw[:160], word_count=len(raw.split()))
        ],
        tables=[],
        figures=[],
        named_entities=_extract_named_entities(raw),
        raw_text=raw,
    )


def _extract_ooxml_content(path: Path, raw_text: str) -> DocumentRecord:
    """
    Extract text from Office Open XML content files.

    Strips all namespace-qualified tags and extracts text runs,
    paragraph breaks, and table cell boundaries.

    Token savings rationale: a .docx unpacked to word/document.xml
    loses binary overhead and all formatting markup. Only text content
    remains, reducing tokens by 40-60% vs loading the binary .docx.
    """
    import xml.etree.ElementTree as ET

    # Register common OOXML namespaces to avoid ns0: prefixes
    namespaces = {
        "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "p":  "http://schemas.openxmlformats.org/presentationml/2006/main",
        "v":  "urn:schemas-microsoft-com:vml",
    }
    for prefix, uri in namespaces.items():
        try:
            ET.register_namespace(prefix, uri)
        except Exception:
            pass

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"

    try:
        root = ET.fromstring(raw_text.encode("utf-8"))
    except ET.ParseError:
        # Strip namespaces and retry
        cleaned = re.sub(r'\s+xmlns(?::\w+)?="[^"]*"', "", raw_text)
        cleaned = re.sub(r"<\w+:", "<", cleaned)
        cleaned = re.sub(r"</\w+:", "</", cleaned)
        try:
            root = ET.fromstring(cleaned.encode("utf-8"))
        except ET.ParseError:
            # Final fallback: strip all tags
            text = re.sub(r"<[^>]+>", " ", raw_text)
            text = re.sub(r"\s+", " ", text).strip()
            return _extract_md_txt_from_string(text, path, "xml")

    sections: list[Section] = []
    all_text_parts: list[str] = []

    def get_ns_text(elem, tag_local, ns_uri):
        """Find all elements with a given local name in a namespace."""
        return elem.iter(f"{{{ns_uri}}}{tag_local}")

    # WordprocessingML: extract paragraphs and heading styles
    paragraphs = list(root.iter(f"{{{W}}}p"))
    if paragraphs:
        current_heading: tuple[int, str] | None = None
        current_body: list[str] = []

        def flush_ooxml(heading, body):
            text = " ".join(body).strip()
            if heading or text:
                lvl, ttl = heading if heading else (0, "(preamble)")
                first = re.split(r"[.!?]", text)[0][:160] if text else ""
                sections.append(Section(
                    level=lvl, title=ttl, summary=first,
                    word_count=len(text.split()), page_hint=None,
                ))

        for para in paragraphs:
            # Get heading level from paragraph style
            style_elem = para.find(f".//{{{W}}}pStyle")
            style_val = style_elem.get(f"{{{W}}}val", "") if style_elem is not None else ""
            heading_match = re.match(r"[Hh]eading\s*(\d)", style_val)
            level = int(heading_match.group(1)) if heading_match else 0

            # Collect text runs
            runs = [
                t.text or ""
                for t in para.iter(f"{{{W}}}t")
                if t.text
            ]
            para_text = "".join(runs).strip()
            if para_text:
                all_text_parts.append(para_text)

            if level > 0 and para_text:
                flush_ooxml(current_heading, current_body)
                current_heading = (level, para_text)
                current_body = []
            elif para_text:
                current_body.append(para_text)

        flush_ooxml(current_heading, current_body)
    else:
        # Generic: collect all text nodes
        for elem in root.iter():
            if elem.text and elem.text.strip():
                all_text_parts.append(elem.text.strip())
            if elem.tail and elem.tail.strip():
                all_text_parts.append(elem.tail.strip())

    raw = " ".join(all_text_parts)
    title = path.stem.replace("_", " ").replace("-", " ").title()
    if sections and sections[0].level >= 1:
        title = sections[0].title

    return DocumentRecord(
        path=path,
        file_type="xml",
        title=title,
        total_words=len(raw.split()),
        total_pages=None,
        sections=sections or [
            Section(level=0, title="(content)", summary=raw[:160],
                    word_count=len(raw.split()))
        ],
        tables=[],
        figures=[],
        named_entities=_extract_named_entities(raw),
        raw_text=raw,
    )

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

SUPPORTED = {".docx", ".pdf", ".pptx", ".md", ".txt", ".xml"}

EXTRACTORS = {
    ".docx": _extract_docx,
    ".pdf":  _extract_pdf,
    ".pptx": _extract_pptx,
    ".md":   _extract_md_txt,
    ".txt":  _extract_md_txt,
    ".xml":  _extract_xml,
}


def extract(path: Path | str) -> DocumentRecord:
    """
    Extract structured content from a single document.

    Parameters
    ----------
    path : Path or str
        Path to a .docx, .pdf, .pptx, .md, .txt, or .xml file.

    Returns
    -------
    DocumentRecord
        Structured representation ready for graph building.

    Raises
    ------
    ValueError
        If the file type is not supported.
    FileNotFoundError
        If the file does not exist.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED:
        raise ValueError(
            f"Unsupported file type '{suffix}'. "
            f"Supported: {', '.join(sorted(SUPPORTED))}"
        )
    return EXTRACTORS[suffix](path)