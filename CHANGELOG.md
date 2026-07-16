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
- Generic lattice-layout support: detects the repeated Ref#/Price
  column-block pattern (spec §2a) dynamically from the header row — no
  per-vendor hardcoded column positions required — builds the authoritative
  part+price universe from every block, then left-joins Description/UOM from
  any detailed section found lower on the sheet. Verified against the real
  Atos Medical price list: found the same 6 column blocks and 300 parts (290
  enriched) that `transform_price_list.py`'s hardcoded positions find for
  that same file.
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

### Fixed
- `_normalize_uom()` no longer logs a spurious "Unrecognized UOM value" warning
  when a vendor file already uses a valid ANSI code (e.g. `EA`, `BX`, `CA`) —
  it now recognizes already-normalized codes as well as the full words
  (`each`, `box`, `case`) it previously required. Found via end-to-end smoke
  test.
- Duplicate column headers that don't match the known lattice pattern now
  fail cleanly with an explanatory log message (pointing to
  `transform_price_list.py`) instead of crashing with an opaque
  `ValueError: The truth value of a Series is ambiguous` deep inside
  price-column disambiguation.

### Changed
- `_get_ai_client()` now authenticates via Azure AD (Entra ID) first, using
  `DefaultAzureCredential` and refreshing the bearer token as it nears
  expiry. Falls back to a static `AZURE_OPENAI_API_KEY` only if Azure AD
  auth isn't available, and to `None` (AI features disabled, no crash) if
  neither works. Driven by BroadJump IT disabling key-based auth on the
  `broadjump-foundry-dev` resource (`AuthenticationTypeDisabled` on every
  AI call). Requires the `azure-identity` package (added to
  `requirements-watcher.txt`) and a signed-in credential source (e.g.
  `az login`) with a role assigned on the Foundry resource.

### Known gap
- Manufacturer-name alias lookup can false-match on filename substrings
  (e.g. `AtosMedInc_PriceList.xlsx` matched the alias `"MEDINC"` and resolved
  to "Medtronic" instead of "Atos Medical"). Needs a fix before this is
  trusted on filenames with ambiguous substrings — verify ManufacturerName on
  every output until resolved.

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
