# Changelog

All notable changes to this project are logged here — additions, removals, and
behavioral changes — so it's clear what moved between versions before anything
merges to `main`. Update this file in the same commit as the change it describes.

## [Unreleased] — v1.1 branch

### Added
- `quote_file_watcher.py` — folder-watcher that ingests vendor files dropped into
  a `RawFiles` folder and auto-maps them to the Quote Data Template schema.
  Unlike `transform_price_list.py`, it does not require a per-sheet manifest:
  it detects header rows, merges multi-row headers/continuation rows, maps
  columns via keyword + content-based heuristics, and falls back to Azure
  OpenAI for columns it can't otherwise classify or to disambiguate multiple
  price-like columns.
- PDF ingestion (`.pdf` support) — text-based PDFs via `pdfplumber`, with an
  AI-vision fallback (renders pages as images, asks the model to extract the
  table) for scanned/complex PDFs.
- Manufacturer-name resolution via filename, using an alias crosswalk
  (`ManufacturerAlias.xlsx`, not committed — see Security below) when no
  manufacturer column is present in the source file.
- `requirements.txt` — watcher dependencies (pandas, openpyxl, watchdog, openai,
  pdfplumber, pymupdf).
- `QP_WatcherInstall.docx` — setup/handoff instructions for running the watcher
  on a new machine.
- `QuoteTemplate.xlsx` — canonical empty output template (headers only),
  matching `TEMPLATE_COLUMNS` / the spec's output schema exactly.

### Security
- Azure OpenAI API key removed from source; the script now reads
  `AZURE_OPENAI_API_KEY` from the environment. Never hardcode a key again —
  this is a standing rule for any future secret in this repo.
- `.gitignore` expanded so real vendor/customer data (raw price lists,
  corrected outputs, the manufacturer alias crosswalk, source PDFs) can never
  enter git history. This repo tracks code, spec, and empty templates only.

## [1.0.0] — initial manifest-based transform

### Added
- `transform_price_list.py` + `TRANSFORMATION_SPEC.md` — manifest-driven
  transform for two known vendor layouts: Atos Medical (lattice) and Tracoe
  (tabular). Each source sheet is explicitly mapped to a manufacturer name and
  layout type via `SHEET_MANIFEST`.
