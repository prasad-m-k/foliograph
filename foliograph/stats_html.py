"""
foliograph.stats_html
~~~~~~~~~~~~~~~~~~~~~
Generate a self-contained FOLIO_STATS.html dashboard from an existing
FOLIO_GRAPH.md and FOLIO_INDEX.md.

The dashboard shows:
  - Token savings estimate (per session, per month, per year)
  - Interactive sliders to model different usage patterns
  - Section-level breakdown
  - Cumulative savings trend chart

Human and analytical tone throughout. No em-dashes.
"""

from __future__ import annotations

import re
from pathlib import Path


# ---------------------------------------------------------------------------
# Graph parser
# ---------------------------------------------------------------------------

def _parse_graph(graph_path: Path) -> dict:
    """Extract key metrics from an existing FOLIO_GRAPH.md."""
    text = graph_path.read_text(encoding="utf-8", errors="replace")

    project = "My Project"
    m = re.search(r"\*\*Project:\*\*\s+(.+)", text)
    if m:
        project = m.group(1).strip().rstrip()

    built = ""
    m = re.search(r"\*\*Built:\*\*\s+(.+)", text)
    if m:
        built = m.group(1).strip().rstrip()

    total_words = 0
    m = re.search(r"\*\*Total words:\*\*\s+([\d,]+)", text)
    if m:
        total_words = int(m.group(1).replace(",", ""))

    doc_count = 0
    m = re.search(r"\*\*Documents:\*\*\s+(\d+)", text)
    if m:
        doc_count = int(m.group(1))

    rel_count = 0
    m = re.search(r"\*\*Relationships:\*\*\s+(\d+)", text)
    if m:
        rel_count = int(m.group(1))

    # Per-file word counts
    files = []
    for fm in re.finditer(
        r"### `([^`]+)` \[([A-Z]+)\].*?\*\*Words:\*\*\s*([\d,]+)",
        text, re.DOTALL
    ):
        files.append({
            "name": fm.group(1),
            "type": fm.group(2),
            "words": int(fm.group(3).replace(",", "")),
        })

    return {
        "project": project,
        "built": built,
        "total_words": total_words,
        "doc_count": doc_count,
        "rel_count": rel_count,
        "files": files,
    }


def _parse_index(index_path: Path) -> int:
    """Count entries in FOLIO_INDEX.md."""
    if not index_path.exists():
        return 0
    text = index_path.read_text(encoding="utf-8", errors="replace")
    m = re.search(r"\*\*Entries:\*\*\s+(\d+)", text)
    if m:
        return int(m.group(1))
    return len(re.findall(r"^- \*\*", text, re.MULTILINE))


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Foliograph Token Savings: {project}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    font-size: 15px; line-height: 1.6;
    background: #f8f9fb; color: #1A1A1A;
    padding: 2rem 1rem;
  }}
  .container {{ max-width: 860px; margin: 0 auto; }}
  h1 {{ font-size: 22px; font-weight: 500; margin-bottom: 4px; color: #002060; }}
  .subtitle {{ font-size: 13px; color: #002060; margin-bottom: 2rem; opacity: 0.7; }}
  .section {{ margin-bottom: 2rem; }}
  .section-title {{ font-size: 14px; font-weight: 500;
    color: #002060; text-transform: uppercase;
    letter-spacing: 0.05em; margin-bottom: 12px; }}
  .metric-grid {{ display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 10px; margin-bottom: 1.5rem; }}
  .metric {{
    background: #fff; border: 1px solid #E07020;
    border-radius: 8px; padding: 14px 16px; }}
  .metric-label {{ font-size: 12px; color: #002060; margin-bottom: 6px; opacity: 0.7; }}
  .metric-value {{ font-size: 22px; font-weight: 500; color: #002060; }}
  .metric-sub {{ font-size: 11px; color: #1A1A1A; margin-top: 2px; opacity: 0.5; }}
  .card {{ background: #fff; border: 1px solid #E07020;
    border-radius: 10px; padding: 1.25rem; margin-bottom: 1rem; }}
  .compare-grid {{ display: grid;
    grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 1.5rem; }}
  .compare-box {{ border-radius: 8px; padding: 16px; border: 1px solid #E07020; }}
  .compare-box.savings {{ background: #FFF4E6; border-color: #E07020; }}
  .compare-box.baseline {{ background: #F0F4FA; }}
  .compare-title {{ font-size: 12px; font-weight: 500; margin-bottom: 8px; color: #002060; }}
  .compare-num {{ font-size: 28px; font-weight: 500; }}
  .compare-num.green {{ color: #E07020; }}
  .compare-unit {{ font-size: 12px; color: #1A1A1A; margin-top: 2px; opacity: 0.6; }}
  .slider-row {{ display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }}
  .slider-label {{ font-size: 13px; color: #1A1A1A; min-width: 200px; }}
  .slider-val {{ font-size: 13px; font-weight: 500; min-width: 70px; text-align: right; }}
  input[type=range] {{
    flex: 1; height: 4px; cursor: pointer;
    accent-color: #E07020;
  }}
  .bar-row {{ margin-bottom: 14px; }}
  .bar-label-row {{ display: flex; justify-content: space-between;
    font-size: 12px; color: #1A1A1A; margin-bottom: 4px; opacity: 0.7; }}
  .bar-track {{ height: 10px; background: #F0F4FA;
    border-radius: 99px; overflow: hidden; }}
  .bar-fill {{ height: 100%; border-radius: 99px;
    transition: width 0.3s ease; }}
  .file-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .file-table th {{ text-align: left; padding: 8px 12px;
    font-weight: 500; color: #fff; background: #002060;
    border-bottom: 1px solid #E07020; }}
  .file-table td {{ padding: 8px 12px; border-bottom: 1px solid #F0F4FA; }}
  .file-table tr:last-child td {{ border-bottom: none; }}
  .file-table tbody tr:nth-child(odd) {{ background: #F7F9FC; }}
  .file-table tbody tr:nth-child(even) {{ background: #ffffff; }}
  .badge {{ display: inline-block; font-size: 11px; padding: 2px 8px;
    border-radius: 99px; font-weight: 500; }}
  .badge-green {{ background: #FFF4E6; color: #E07020; }}
  .badge-blue {{ background: #F0F4FA; color: #002060; }}
  .badge-amber {{ background: #FFF4E6; color: #E07020; }}
  footer {{ text-align: center; font-size: 12px; color: #002060;
    margin-top: 3rem; padding-top: 1rem;
    border-top: 1px solid #E07020; opacity: 0.7; }}
  @media (max-width: 500px) {{
    .compare-grid {{ grid-template-columns: 1fr; }}
    .slider-label {{ min-width: 140px; }}
  }}
</style>
</head>
<body>
<div class="container">

  <h1>Foliograph Token Savings</h1>
  <p class="subtitle">Project: {project} &nbsp;|&nbsp; Built: {built}
    &nbsp;|&nbsp; {doc_count} document(s) &nbsp;|&nbsp;
    {total_words} words &nbsp;|&nbsp;
    {rel_count} relationship(s) &nbsp;|&nbsp;
    {index_entries} index entries</p>

  <div class="section">
    <div class="section-title">Savings at a glance</div>
    <div class="compare-grid">
      <div class="compare-box baseline">
        <div class="compare-title">Without Foliograph</div>
        <div class="compare-num" id="cmp-without">--</div>
        <div class="compare-unit">tokens per session</div>
      </div>
      <div class="compare-box savings">
        <div class="compare-title">With Foliograph</div>
        <div class="compare-num green" id="cmp-with">--</div>
        <div class="compare-unit">tokens per session</div>
      </div>
    </div>
    <div class="metric-grid">
      <div class="metric">
        <div class="metric-label">Compression ratio</div>
        <div class="metric-value" id="m-ratio">--</div>
        <div class="metric-sub">graph vs full documents</div>
      </div>
      <div class="metric">
        <div class="metric-label">Saved per session</div>
        <div class="metric-value" id="m-session">--</div>
        <div class="metric-sub">tokens</div>
      </div>
      <div class="metric">
        <div class="metric-label">Saved per month</div>
        <div class="metric-value" id="m-month">--</div>
        <div class="metric-sub">tokens</div>
      </div>
      <div class="metric">
        <div class="metric-label">Saved per year</div>
        <div class="metric-value" id="m-year">--</div>
        <div class="metric-sub">tokens</div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Model your usage</div>
    <div class="card">
      <div class="slider-row">
        <span class="slider-label">Sessions per month</span>
        <input type="range" id="sl-sessions" min="1" max="60" value="10" step="1">
        <span class="slider-val" id="out-sessions">10</span>
      </div>
      <div class="slider-row">
        <span class="slider-label">Sections loaded per session</span>
        <input type="range" id="sl-sections" min="1" max="20" value="3" step="1">
        <span class="slider-val" id="out-sections">3</span>
      </div>
      <div class="slider-row">
        <span class="slider-label">Average words per section</span>
        <input type="range" id="sl-secwords" min="100" max="3000" value="600" step="100">
        <span class="slider-val" id="out-secwords">600</span>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Token breakdown per session</div>
    <div class="card">
      <div class="bar-row">
        <div class="bar-label-row">
          <span>Graph startup (FOLIO_GRAPH + INDEX)</span>
          <span id="bar-graph-val">--</span>
        </div>
        <div class="bar-track">
          <div class="bar-fill" id="bar-graph"
            style="background:#002060;width:2%;"></div>
        </div>
      </div>
      <div class="bar-row">
        <div class="bar-label-row">
          <span>Sections loaded on demand</span>
          <span id="bar-sec-val">--</span>
        </div>
        <div class="bar-track">
          <div class="bar-fill" id="bar-sec"
            style="background:#E07020;width:2%;"></div>
        </div>
      </div>
      <div class="bar-row">
        <div class="bar-label-row">
          <span>Would have loaded without Foliograph</span>
          <span id="bar-full-val">--</span>
        </div>
        <div class="bar-track">
          <div class="bar-fill" id="bar-full"
            style="background:#1A1A1A;width:100%;opacity:0.15;"></div>
        </div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Cumulative savings over 12 months</div>
    <div class="card">
      <div style="position:relative;width:100%;height:200px;">
        <canvas id="trend-chart"
          role="img"
          aria-label="Line chart showing cumulative token savings growing each month">
          Savings increase each month as more sessions use the graph.
        </canvas>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Documents indexed</div>
    <div class="card" style="padding:0;">
      <table class="file-table">
        <thead>
          <tr>
            <th>File</th>
            <th>Type</th>
            <th>Words</th>
            <th>Tokens (est.)</th>
            <th>Savings per load</th>
          </tr>
        </thead>
        <tbody id="file-tbody"></tbody>
      </table>
    </div>
  </div>

  <footer>
    Foliograph, created by Prasad MK &nbsp;|&nbsp;
    <a href="https://ssrn.com/author=10270516" target="_blank">
      ssrn.com/author=10270516</a> &nbsp;|&nbsp;
    <a href="https://github.com/prasad-m-k/foliograph" target="_blank">
      github.com/prasad-m-k/foliograph</a>
  </footer>

</div>

<script>
const GRAPH_TOKENS = {graph_tokens};
const FULL_TOKENS  = {full_tokens};
const FILES = {files_json};

function fmt(n) {{
  n = Math.round(n);
  if (n >= 1000000) return (n/1000000).toFixed(1) + 'M';
  if (n >= 1000)    return (n/1000).toFixed(0) + 'K';
  return n.toLocaleString();
}}

let trendChart = null;

function recalc() {{
  const sessions  = parseInt(document.getElementById('sl-sessions').value);
  const sections  = parseInt(document.getElementById('sl-sections').value);
  const secWords  = parseInt(document.getElementById('sl-secwords').value);

  document.getElementById('out-sessions').textContent = sessions;
  document.getElementById('out-sections').textContent = sections;
  document.getElementById('out-secwords').textContent = secWords.toLocaleString();

  const secTokens  = Math.round(sections * secWords * 1.3);
  const withTokens = GRAPH_TOKENS + secTokens;
  const savedSess  = Math.max(0, FULL_TOKENS - withTokens);
  const savedMonth = savedSess * sessions;
  const savedYear  = savedMonth * 12;
  const ratio      = Math.max(1, Math.round(FULL_TOKENS / withTokens));

  document.getElementById('cmp-without').textContent = fmt(FULL_TOKENS);
  document.getElementById('cmp-with').textContent    = fmt(withTokens);
  document.getElementById('m-ratio').textContent     = ratio + 'x';
  document.getElementById('m-session').textContent   = fmt(savedSess);
  document.getElementById('m-month').textContent     = fmt(savedMonth);
  document.getElementById('m-year').textContent      = fmt(savedYear);

  document.getElementById('bar-graph-val').textContent =
    GRAPH_TOKENS.toLocaleString() + ' tokens';
  document.getElementById('bar-sec-val').textContent   =
    secTokens.toLocaleString() + ' tokens';
  document.getElementById('bar-full-val').textContent  =
    FULL_TOKENS.toLocaleString() + ' tokens';

  const maxBar = FULL_TOKENS;
  document.getElementById('bar-graph').style.width =
    Math.max(2, Math.round(GRAPH_TOKENS / maxBar * 100)) + '%';
  document.getElementById('bar-sec').style.width   =
    Math.max(2, Math.round(secTokens / maxBar * 100)) + '%';

  const months   = ['Jan','Feb','Mar','Apr','May','Jun',
                    'Jul','Aug','Sep','Oct','Nov','Dec'];
  const cumData  = months.map((_, i) => Math.round(savedMonth * (i + 1)));

  if (trendChart) {{
    trendChart.data.datasets[0].data = cumData;
    trendChart.update('none');
  }}
}}

function buildFileTable() {{
  const tbody = document.getElementById('file-tbody');
  tbody.innerHTML = FILES.map(f => {{
    const ft   = Math.round(f.words * 1.3);
    const pct  = Math.round((1 - (GRAPH_TOKENS / ft)) * 100);
    const safe = Math.max(0, pct);
    const typeColors = {{
      DOCX: 'badge-blue', PPTX: 'badge-amber',
      PDF: 'badge-green', MD: 'badge-green', TXT: 'badge-green'
    }};
    const cls = typeColors[f.type] || 'badge-blue';
    return `<tr>
      <td>${{f.name}}</td>
      <td><span class="badge ${{cls}}">${{f.type}}</span></td>
      <td>${{f.words.toLocaleString()}}</td>
      <td>~${{fmt(ft)}}</td>
      <td>~${{safe}}% per session</td>
    </tr>`;
  }}).join('');
}}

window.addEventListener('load', function() {{
  buildFileTable();

  const ctx = document.getElementById('trend-chart').getContext('2d');
  const months = ['Jan','Feb','Mar','Apr','May','Jun',
                  'Jul','Aug','Sep','Oct','Nov','Dec'];
  trendChart = new Chart(ctx, {{
    type: 'line',
    data: {{
      labels: months,
      datasets: [{{
        label: 'Cumulative tokens saved',
        data: months.map((_, i) => Math.round(50000 * (i + 1))),
        borderColor: '#E07020',
        backgroundColor: 'rgba(224,112,32,0.08)',
        fill: true, tension: 0.3,
        pointRadius: 3, pointBackgroundColor: '#E07020',
      }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            label: (ctx) => ' ' + fmt(ctx.raw) + ' tokens saved'
          }}
        }}
      }},
      scales: {{
        x: {{ ticks: {{ font: {{ size: 11 }}, color: '#002060' }},
              grid: {{ display: false }} }},
        y: {{ ticks: {{ font: {{ size: 11 }}, color: '#002060',
                        callback: (v) => fmt(v) }},
              grid: {{ color: 'rgba(0,32,96,0.06)' }} }}
      }}
    }}
  }});

  ['sl-sessions','sl-sections','sl-secwords'].forEach(id => {{
    document.getElementById(id).addEventListener('input', recalc);
  }});

  recalc();
}});
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_stats_html(
    graph_path: Path,
    index_path: Path | None = None,
    output_path: Path | None = None,
) -> Path:
    """
    Generate a self-contained FOLIO_STATS.html dashboard.

    Parameters
    ----------
    graph_path : Path
        Path to an existing FOLIO_GRAPH.md.
    index_path : Path, optional
        Path to FOLIO_INDEX.md. Defaults to same directory as graph.
    output_path : Path, optional
        Where to write the HTML. Defaults to FOLIO_STATS.html in graph's directory.

    Returns
    -------
    Path
        Path to the written HTML file.
    """
    import json

    graph_path = Path(graph_path)
    if index_path is None:
        index_path = graph_path.parent / "FOLIO_INDEX.md"
    if output_path is None:
        output_path = graph_path.parent / "FOLIO_STATS.html"

    data = _parse_graph(graph_path)
    index_entries = _parse_index(Path(index_path))

    # Token estimates
    full_tokens = int(data["total_words"] * 1.3)
    graph_size = graph_path.stat().st_size
    idx_size = Path(index_path).stat().st_size if Path(index_path).exists() else 0
    graph_tokens = (graph_size + idx_size) // 4

    files_json = json.dumps(data["files"])

    html = _HTML_TEMPLATE.format(
        project=data["project"],
        built=data["built"] or "unknown",
        doc_count=data["doc_count"],
        total_words=f"{data['total_words']:,}",
        rel_count=data["rel_count"],
        index_entries=index_entries,
        graph_tokens=graph_tokens,
        full_tokens=full_tokens,
        files_json=files_json,
    )

    output_path = Path(output_path)
    output_path.write_text(html, encoding="utf-8")
    return output_path
