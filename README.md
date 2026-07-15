# QPMapper

Converts raw vendor/manufacturer price lists into a single normalized
"Quote Data Template" sheet.

- **Spec (source of truth):** [TRANSFORMATION_SPEC.md](TRANSFORMATION_SPEC.md)
- **Script:** [transform_price_list.py](transform_price_list.py)

Keep the spec and script in sync — every change to parsing/derivation logic
should update both in the same commit, cross-referenced by section number.

## Usage

```
python transform_price_list.py <input_workbook.xlsx> <output_workbook.xlsx>
```

Edit `SHEET_MANIFEST` in `transform_price_list.py` (or pass `--manifest a.json`)
to map each source sheet to a manufacturer name + layout type.

## Adding a new vendor layout

See spec §6 — add a case to `LAYOUTS`, don't special-case the parsing loop.
