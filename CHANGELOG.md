# Changelog

All notable changes to this project are logged here — additions, removals, and
behavioral changes — so it's clear what moved between versions before anything
merges to `main`. Update this file in the same commit as the change it describes.

## [1.0.0] — initial manifest-based transform

### Added
- `transform_price_list.py` + `TRANSFORMATION_SPEC.md` — manifest-driven
  transform for two known vendor layouts: Atos Medical (lattice) and Tracoe
  (tabular). Each source sheet is explicitly mapped to a manufacturer name and
  layout type via `SHEET_MANIFEST`.
