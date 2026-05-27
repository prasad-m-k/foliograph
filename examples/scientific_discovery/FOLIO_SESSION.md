# Foliograph Session Starter

Paste this block at the start of every new LLM session to orient the model
with minimal token cost.

---

```
You have access to a Foliograph knowledge graph for this document set.
Graph files available:

  FOLIO_GRAPH.md      - structural map of every document + relationships
  FOLIO_INDEX.md      - concept → location index
  FOLIO_RELATIONS.json- machine-readable relationship graph

Rules for this session:
1. Read FOLIO_GRAPH.md first for orientation (do not load source files yet).
2. Use FOLIO_INDEX.md to locate any concept before loading a full section.
3. Load sections on demand only: "Load [filename] § [Section Title]"
   Do not load entire source files unless explicitly asked.
4. Never re-read a section you have already processed this session.
5. If FOLIO_GRAPH.md shows a drift WARNING, note it before answering
   questions that touch the flagged section.

Documents in scope:
  - ScientificDiscovery.pptx
```
