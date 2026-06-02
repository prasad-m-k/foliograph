"""
tests/test_foliograph.py
------------------------
Basic tests for extractor and builder.
Run with: pytest tests/
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from foliograph.extractor import _extract_md_txt, _extract_named_entities, extract
from foliograph.builder import build


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

SAMPLE_MD = textwrap.dedent("""\
    # The State of Escalation

    What the numbers tell us.

    ## The Real Cost Nobody Is Measuring

    When organizations talk about escalation costs, they tend to measure
    what is easy to count. Time to resolution. Number of tickets.

    ## The Architecture of a Broken Handoff

    The traditional escalation model was designed for a world that no
    longer exists. The Swarm Model replaces it.

    ### A Sub-section

    Detail here. The Feedback Loop closes the cycle.

    ## Introducing Escalation Debt

    Every time an escalation is closed without capturing what caused it,
    you take on Process Debt.
""")


@pytest.fixture
def sample_md_file(tmp_path: Path) -> Path:
    p = tmp_path / "chapter1.md"
    p.write_text(SAMPLE_MD)
    return p


# ---------------------------------------------------------------------------
# Extractor tests
# ---------------------------------------------------------------------------

class TestMdExtractor:
    def test_title_extracted(self, sample_md_file):
        rec = _extract_md_txt(sample_md_file)
        assert rec.title == "The State of Escalation"

    def test_section_count(self, sample_md_file):
        rec = _extract_md_txt(sample_md_file)
        # H1 + 3 H2s + 1 H3 = 5 sections (preamble sections with level>0)
        titled = [s for s in rec.sections if s.level > 0]
        assert len(titled) >= 4

    def test_word_count_positive(self, sample_md_file):
        rec = _extract_md_txt(sample_md_file)
        assert rec.total_words > 50

    def test_summaries_not_empty(self, sample_md_file):
        rec = _extract_md_txt(sample_md_file)
        for sec in rec.sections:
            if sec.level > 0 and sec.word_count > 5:
                assert sec.summary, f"Empty summary for section '{sec.title}'"

    def test_named_entities_include_known_terms(self, sample_md_file):
        rec = _extract_md_txt(sample_md_file)
        all_entities = " ".join(rec.named_entities)
        # Swarm Model and Feedback Loop should be detected
        assert "Swarm Model" in all_entities or "Escalation Debt" in all_entities


class TestNamedEntityExtraction:
    def test_acronym_detection(self):
        text = "The NLP model uses SLA data and REST APIs."
        entities = _extract_named_entities(text)
        assert "NLP" in entities
        assert "SLA" in entities

    def test_model_framework_detection(self):
        text = "The Swarm Model replaces the Feedback Loop in most scenarios."
        entities = _extract_named_entities(text)
        assert any("Swarm Model" in e for e in entities)

    def test_no_crash_on_empty(self):
        assert _extract_named_entities("") == []


class TestExtractPublicApi:
    def test_unsupported_type_raises(self, tmp_path):
        p = tmp_path / "file.csv"
        p.write_text("a,b,c")
        with pytest.raises(ValueError, match="Unsupported"):
            extract(p)

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            extract(tmp_path / "nonexistent.md")

    def test_md_round_trip(self, sample_md_file):
        rec = extract(sample_md_file)
        assert rec.file_type == "md"
        assert rec.total_words > 0
        assert len(rec.sections) > 0


# ---------------------------------------------------------------------------
# Builder tests
# ---------------------------------------------------------------------------

class TestBuilder:
    def test_build_creates_files(self, tmp_path, sample_md_file):
        out = tmp_path / "graph_out"
        result = build(
            sources=[sample_md_file],
            output_dir=out,
            project_name="Test Project",
        )
        assert result["graph"].exists()
        assert result["index"].exists()
        assert result["session"].exists()

    def test_graph_contains_title(self, tmp_path, sample_md_file):
        out = tmp_path / "graph_out"
        build(sources=[sample_md_file], output_dir=out, project_name="Test")
        content = (out / "FOLIO_GRAPH.md").read_text()
        assert "The State of Escalation" in content

    def test_index_contains_entries(self, tmp_path, sample_md_file):
        out = tmp_path / "graph_out"
        build(sources=[sample_md_file], output_dir=out)
        content = (out / "FOLIO_INDEX.md").read_text()
        assert "→" in content  # index entries use →

    def test_no_session_flag(self, tmp_path, sample_md_file):
        out = tmp_path / "graph_out"
        result = build(
            sources=[sample_md_file],
            output_dir=out,
            include_session_starter=False,
        )
        assert "session" not in result
        assert not (out / "FOLIO_SESSION.md").exists()

    def test_directory_scan(self, tmp_path):
        docs = tmp_path / "docs"
        docs.mkdir()
        for i in range(3):
            (docs / f"doc{i}.md").write_text(
                f"# Chapter {i}\n\nContent for chapter {i}."
            )
        out = tmp_path / "out"
        result = build(sources=[docs], output_dir=out, project_name="Multi")
        content = (out / "FOLIO_GRAPH.md").read_text()
        assert "doc0.md" in content
        assert "doc2.md" in content

    def test_graph_token_estimate_printed(
        self, tmp_path, sample_md_file, capsys
    ):
        out = tmp_path / "out"
        build(sources=[sample_md_file], output_dir=out)
        captured = capsys.readouterr()
        assert "tokens" in captured.out.lower()

    def test_empty_sources_raises(self, tmp_path):
        with pytest.raises(ValueError):
            build(sources=[], output_dir=tmp_path)


class TestXlsxExtractor:
    pytest.importorskip("openpyxl")

    def _make_workbook(self, tmp_path, sheets: dict) -> Path:
        import openpyxl
        wb = openpyxl.Workbook()
        first = True
        for sheet_name, rows in sheets.items():
            ws = wb.active if first else wb.create_sheet(sheet_name)
            if first:
                ws.title = sheet_name
                first = False
            for row in rows:
                ws.append(row)
        p = tmp_path / "test.xlsx"
        wb.save(str(p))
        return p

    def test_sheets_become_sections(self, tmp_path):
        p = self._make_workbook(tmp_path, {
            "Revenue": [["Region", "Q1", "Q2"], ["North", 100, 200], ["South", 150, 250]],
            "Costs":   [["Category", "Amount"], ["Salaries", 500], ["Rent", 200]],
        })
        rec = extract(p)
        titles = [s.title for s in rec.sections]
        assert "Revenue" in titles
        assert "Costs" in titles

    def test_headers_registered_as_tables(self, tmp_path):
        p = self._make_workbook(tmp_path, {
            "Data": [["Name", "Value", "Date"], ["Alpha", 1, "2024-01"], ["Beta", 2, "2024-02"]],
        })
        rec = extract(p)
        assert any("Name" in t for t in rec.tables)

    def test_word_count_positive(self, tmp_path):
        p = self._make_workbook(tmp_path, {
            "Sheet1": [["Header"], ["Some text content here"]],
        })
        rec = extract(p)
        assert rec.total_words > 0

    def test_file_type(self, tmp_path):
        p = self._make_workbook(tmp_path, {"Sheet1": [["A", "B"]]})
        rec = extract(p)
        assert rec.file_type == "xlsx"

    def test_xlsx_in_supported(self):
        from foliograph.extractor import SUPPORTED
        assert ".xlsx" in SUPPORTED


class TestXmlExtractor:
    def test_generic_xml(self, tmp_path):
        xml = tmp_path / "doc.xml"
        xml.write_text("""<?xml version="1.0"?>
<document>
  <title>Test Document</title>
  <section>
    <heading>Introduction</heading>
    <para>This is the first paragraph of the document.</para>
  </section>
  <section>
    <heading>Conclusion</heading>
    <para>This is the conclusion section content.</para>
  </section>
</document>""")
        rec = extract(xml)
        assert rec.file_type == "xml"
        assert rec.total_words > 0

    def test_ooxml_wordprocessing(self, tmp_path):
        xml = tmp_path / "document.xml"
        xml.write_text("""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Chapter One</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">This is the body text of chapter one.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>A Subsection</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Subsection body text here.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>""")
        rec = extract(xml)
        assert rec.file_type == "xml"
        assert rec.total_words > 0
        titled = [s for s in rec.sections if s.level > 0]
        assert len(titled) >= 2

    def test_xml_in_supported(self):
        from foliograph.extractor import SUPPORTED
        assert ".xml" in SUPPORTED

    def test_malformed_xml_fallback(self, tmp_path):
        xml = tmp_path / "broken.xml"
        xml.write_text("<root><unclosed>Some text content here</root>")
        rec = extract(xml)
        assert rec.file_type == "xml"
        assert rec.total_words > 0
