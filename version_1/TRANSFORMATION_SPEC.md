# Price List → Quote Data Template — Transformation Spec

**Version:** 1.0.0
**Status:** Working / authoritative
**Scope:** Converts raw vendor/manufacturer price lists into a single normalized
"Quote Data Template" sheet.

> **Merge note:** This version is derived solely from the Atos Medical / Tracoe
> handoff document. A second Python project's logic was referenced as a future
> merge target but has not yet been supplied. When it is, reconcile it into
> `version_2` and log the diff in the Changelog below — don't edit this file
> in place once a version ships.

---

## 1. Output Schema

One sheet, exact column order:

| # | Column header (exact)            | Type  | Notes                                |
|---|-----------------------------------|-------|---------------------------------------|
| 0 | ManufacturerName                 | str   | One value per source sheet            |
| 1 | Manufacturer Catalog Number       | str   | Ref# / part number                    |
| 2 | Manufacturer Catalog Description  | str   | May be blank if source lacks it       |
| 3 | Proposed UOM                      | str   | Normalized unit (PC/BOX/SET/KIT/EA)   |
| 4 | Proposed UOM Quantity             | int   | Units per UOM; blank if unknown       |
| 5 | Proposed UOM Price                | float | Numeric only                          |
| 6 | Proposed Purchase Quantity        | int   | Always blank (buyer fills)            |

Rows are stacked per source sheet — all rows of sheet A, then all rows of sheet B, etc.
Every row carries its own `ManufacturerName`.

---

## 2. Source Layout Types

Price lists are not clean tables. Each source sheet is classified into one of the
layout types below before parsing.

### 2a. Lattice layout (e.g. Atos Medical)

- Skip top metadata rows: title, "20XX Price List", "Pricing effective…", customer
  tier labels ("Direct User" / "Hospital"), internal codes.
- **Quick-reference lattice:** the same `(Ref#, Price)` pair repeated across several
  column blocks on the same rows. Observed block column pairs (0-indexed from col A):

  ```
  [(1,2), (5,6), (9,10), (12,14), (16,17), (22,24)]
  ```

  Walk every block for every data row → record `(ref, price)`.

- **Detailed section**, lower on the sheet, header repeats roughly every 77 rows.
  Columns (0-indexed from col B):

  | Field        | Index (from col B) | Excel col |
  |--------------|---------------------|-----------|
  | REF#         | 0                   | B         |
  | Description  | 2                   | D         |
  | Rx / !       | 12                  | N *(not mapped to output)* |
  | U/M          | 16                  | R         |
  | Price        | 19                  | U         |

- The lattice is the **authoritative part+price universe**. The detailed section adds
  Description + U/M but does not cover every part.
- **Strategy:** build rows from the lattice, then LEFT-JOIN Description + U/M from the
  detailed section, keyed on Ref#. Unmatched parts keep blank Description/UOM.

### 2b. Tabular layout (e.g. Tracoe)

- Skip metadata rows and section-header rows (a section header has exactly one filled
  cell and no price).
- Data columns (0-indexed from col A):

  | Field        | Index | Excel col |
  |--------------|-------|-----------|
  | Ref#         | 1     | B         |
  | Sizes        | 3     | D         |
  | Description  | 5     | F         |
  | UoM          | 8     | I         |
  | Pc/UoM       | 10    | K         |
  | Price        | 11    | L         |

---

## 3. Field Derivation Rules

- **ManufacturerName** — derived from sheet title / metadata / file name
  (e.g. `"AtosMedInc"` → `"Atos Medical"`, `"Tracoe"` → `"Tracoe"`).
  **Never** use the customer tier ("Direct User", "Hospital") as the manufacturer name.
- **Catalog Number** — strip a leading `"REF "` (case-insensitive); keep the rest
  verbatim (e.g. `301`, `888-306`, `501-X`).
- **Description**:
  - If a Sizes column exists (tabular layout): `f"{description} ({sizes})"` when
    sizes is present, else just `description`.
  - Otherwise: use the Description cell as-is.
- **Proposed UOM + Quantity**:
  - Lattice/detailed: parse the U/M string with `split_uom()`.
  - Tabular: `UOM = normalize_unit(UoM)`; `Quantity = Pc/UoM`.
- **Proposed UOM Price** — numeric price via `to_price()`.
- **Proposed Purchase Quantity** — always blank; buyer fills this in later.

---

## 4. Row Validity / Junk Filter

Keep a row only if it has **both** a Ref# **and** a numeric price.

```python
import re

NUM_RE = re.compile(r'^\$?\d[\d,]*\.?\d*$')

def is_price(p) -> bool:
    if isinstance(p, (int, float)):
        return True
    return bool(NUM_RE.match(str(p).strip()))

def to_price(p) -> float:
    if isinstance(p, (int, float)):
        return float(p)
    return float(str(p).replace('$', '').replace(',', ''))
```

A "price" containing slashes/letters (e.g. `06/05/026`, an internal code) fails
`is_price()` → the row is dropped.

---

## 5. UOM Parsing

```python
def normalize_unit(u: str) -> str:
    l = u.lower()
    if l.startswith("pc"):  return "PC"
    if l.startswith("set"): return "SET"
    if l.startswith("kit"): return "KIT"
    if l.startswith("box"): return "BOX"
    return u.upper()

def split_uom(uom_str: str):
    """'1 pc' -> (1, 'PC'); 'BOX' -> ('', 'BOX')"""
    m = re.match(r'^(\d+)\s*(.*)$', uom_str.strip())
    if m:
        qty = int(m.group(1))
        unit = m.group(2).strip() or uom_str
    else:
        qty, unit = "", uom_str
    return qty, normalize_unit(unit)
```

---

## 6. Adding a New Vendor / Layout

When a new price list doesn't fit Lattice or Tabular:

1. Identify metadata rows to skip and the manufacturer-name source.
2. Determine whether Ref#+Price appear once per row (tabular) or repeated in
   column blocks (lattice).
3. Record exact 0-indexed column positions in a new subsection under §2.
4. Add a case to `LAYOUTS` in the script (see `transform_price_list.py`) — do not
   special-case it inline in the row-building loop.
5. Bump the spec's minor version and add a Changelog entry.

---

## Changelog

- **1.0.0** — Initial consolidated spec from Atos Medical / Tracoe handoff doc.
  No second-project merge yet (pending).
