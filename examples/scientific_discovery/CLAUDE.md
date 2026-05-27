# Foliograph Context (auto-injected)

This project uses Foliograph for token-efficient document navigation.

## On every session start

1. Read `FOLIO_GRAPH.md` for the structural map of all documents.
2. Use `FOLIO_INDEX.md` to locate concepts before loading source files.
3. Load sections on demand: respond to "Load [file] § [Section]" by reading
   only that section from the source document.
4. Run `foliograph check` before any session where you plan to edit source
   documents, to verify the graph is current.

## Documents indexed

  - ScientificDiscovery.pptx

## Quick reference

- Rebuild graph:  `foliograph build examples/scientific_discovery/ScientificDiscovery.pptx -o examples/scientific_discovery`
- Check for drift: `foliograph check --graph examples/scientific_discovery/FOLIO_GRAPH.md`
- Fetch a section: `foliograph fetch "[file] § [Section Title]"`
