# QPMapper

Converts raw vendor/manufacturer price lists into a single normalized
"Quote Data Template" sheet.

- **Spec (source of truth):** [TRANSFORMATION_SPEC.md](TRANSFORMATION_SPEC.md)
- **Script:** [transform_price_list.py](transform_price_list.py)
- **Changelog:** [CHANGELOG.md](CHANGELOG.md) — update in the same commit as
  any added/removed/changed behavior, so version branches (v1.1, v1.2, ...)
  stay traceable before merging to `main`.

Keep the spec and script in sync — every change to parsing/derivation logic
should update both in the same commit, cross-referenced by section number.

## Data handling

This repo tracks code, spec, and empty templates only. Real vendor/customer
data (raw price lists, corrected outputs, the manufacturer alias crosswalk,
source PDFs) must never be committed — see `.gitignore`. Never hardcode API
keys or other secrets; read them from the environment.

## v1.1 — quote_file_watcher.py

A folder-watcher variant that auto-maps arbitrary vendor `.xlsx`/`.pdf` files
without a manifest, using column-keyword/content heuristics plus an Azure
OpenAI fallback. See [QP_WatcherInstall.docx](QP_WatcherInstall.docx) for setup
and [CHANGELOG.md](CHANGELOG.md) for what it adds relative to
`transform_price_list.py`. Requires `AZURE_OPENAI_API_KEY` set in the
environment to enable AI-assisted column mapping and scanned-PDF extraction.

## Usage

```
python transform_price_list.py <input_workbook.xlsx> <output_workbook.xlsx>
```

Edit `SHEET_MANIFEST` in `transform_price_list.py` (or pass `--manifest a.json`)
to map each source sheet to a manufacturer name + layout type.

## Adding a new vendor layout

See spec §6 — add a case to `LAYOUTS`, don't special-case the parsing loop.
