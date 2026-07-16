#!/usr/bin/env python3
"""
Watches RawFiles for incoming .xlsx files, maps columns to the QuoteTemplate
format, strips $ and commas from numeric fields, and saves to CorrectedFiles.

Usage:
    python quote_file_watcher.py
    python quote_file_watcher.py --process-existing   # also process files already in folder
"""

import argparse
import logging
import os
import re
import time
from pathlib import Path

import pandas as pd
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

try:
    from openai import OpenAI
    _OPENAI_AVAILABLE = True
except ImportError:
    _OPENAI_AVAILABLE = False
    logging.warning("openai package not installed — AI fallback (Pass 6) disabled. "
                    "Run: pip install openai")

try:
    import pdfplumber
    _PDFPLUMBER_AVAILABLE = True
except ImportError:
    _PDFPLUMBER_AVAILABLE = False

try:
    import fitz  # PyMuPDF — PDF-to-image conversion for AI vision fallback
    _PYMUPDF_AVAILABLE = True
except ImportError:
    _PYMUPDF_AVAILABLE = False

try:
    from azure.identity import DefaultAzureCredential
    _AZURE_IDENTITY_AVAILABLE = True
except ImportError:
    _AZURE_IDENTITY_AVAILABLE = False
    logging.warning("azure-identity package not installed — Azure AD auth unavailable. "
                    "Run: pip install azure-identity")

# ── Azure OpenAI config (dev POC) ─────────────────────────────────────────────
# This resource has key-based auth disabled server-side (BroadJump IT policy),
# so Azure AD (Entra ID) auth is the primary path: sign in locally (e.g.
# `az login`) with an account that has a role assigned on the Azure AI
# Foundry resource, and DefaultAzureCredential picks that up automatically.
# AZURE_OPENAI_API_KEY (static key) is kept only as a fallback in case a
# different environment has key auth enabled — never hardcode it.
AZURE_ENDPOINT   = "https://broadjump-foundry-dev.services.ai.azure.com/openai/v1"
AZURE_API_KEY    = os.environ.get("AZURE_OPENAI_API_KEY", "")
AZURE_DEPLOYMENT = "gpt-5.4"
AZURE_API_VER    = "2025-01-01-preview"
AZURE_AD_SCOPE   = "https://cognitiveservices.azure.com/.default"
_TOKEN_REFRESH_MARGIN_SEC = 60

_azure_client = None        # OpenAI client, created once; api_key refreshed in place for AD auth
_azure_credential = None    # DefaultAzureCredential instance, or False if setup failed
_azure_token_expiry = 0     # epoch seconds the current AD bearer token expires


def _refresh_azure_ad_token() -> bool:
    """Fetch/refresh an Azure AD bearer token onto the shared client's api_key.
    Returns True if the client now holds a fresh-enough token."""
    global _azure_credential, _azure_client, _azure_token_expiry

    if _azure_credential is None:
        try:
            _azure_credential = DefaultAzureCredential()
        except Exception as exc:
            logging.warning("Azure AD credential setup failed: %s", exc)
            _azure_credential = False
            return False

    if _azure_credential is False:
        return False

    if _azure_client is not None and time.time() < _azure_token_expiry - _TOKEN_REFRESH_MARGIN_SEC:
        return True  # current token still fresh enough

    try:
        token = _azure_credential.get_token(AZURE_AD_SCOPE)
    except Exception as exc:
        logging.warning(
            "Azure AD token fetch failed: %s. Sign in (e.g. 'az login') and "
            "confirm your account has a role assigned on the Azure AI Foundry "
            "resource.", exc
        )
        return _azure_client is not None  # keep using a stale client if we have one

    _azure_token_expiry = token.expires_on
    if _azure_client is None:
        _azure_client = OpenAI(base_url=AZURE_ENDPOINT, api_key=token.token)
    else:
        _azure_client.api_key = token.token
    return True


def _get_ai_client():
    """Return the OpenAI client pointed at Azure AI Foundry, creating it on first
    call. Prefers Azure AD (Entra ID) auth, refreshing the bearer token as it
    nears expiry; falls back to a static AZURE_OPENAI_API_KEY only if Azure AD
    auth isn't available. Returns None if neither works — AI features degrade
    gracefully rather than crashing."""
    global _azure_client

    if not _OPENAI_AVAILABLE:
        return None

    if _AZURE_IDENTITY_AVAILABLE and _refresh_azure_ad_token():
        return _azure_client

    if _azure_client is not None:
        return _azure_client  # already-created client (e.g. stale AD token) — better than nothing

    if AZURE_API_KEY:
        logging.info("Azure AD auth unavailable — falling back to static AZURE_OPENAI_API_KEY.")
        _azure_client = OpenAI(base_url=AZURE_ENDPOINT, api_key=AZURE_API_KEY)
        return _azure_client

    logging.warning(
        "No usable Azure OpenAI auth — AI fallback disabled. Sign in for Azure "
        "AD access (e.g. 'az login', with a role assigned on the Foundry "
        "resource) or set AZURE_OPENAI_API_KEY."
    )
    return None


def _ask_ai_to_classify_column(col_name: str, sample_values: list) -> tuple[str, str]:
    """Send an unrecognised column name + sample values to Azure OpenAI and ask
    it to return the matching template field name and a one-sentence rationale.
    Returns (field_name, rationale) — both empty strings on any error."""
    client = _get_ai_client()
    if client is None:
        return "", ""

    prompt = (
        "You are helping map vendor spreadsheet columns to a standard quote template.\n\n"
        "Template fields:\n"
        "- ManufacturerName\n"
        "- Manufacturer Catalog Number\n"
        "- Manufacturer Catalog Description\n"
        "- Proposed UOM (2-letter code like EA, BX, CA)\n"
        "- Proposed UOM Quantity (numeric — units per package)\n"
        "- Proposed UOM Price (dollar amount)\n"
        "- Proposed Purchase Quantity (estimated order volume)\n\n"
        f'Column name: "{col_name}"\n'
        f"Sample values: {sample_values}\n\n"
        "Reply in exactly this format on two lines:\n"
        "FIELD: <the matching template field name, or 'ignore'>\n"
        "REASON: <one sentence explaining why>"
    )
    try:
        response = client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            max_completion_tokens=100,
        )
        raw = response.choices[0].message.content.strip()
        field, reason = "", ""
        for line in raw.splitlines():
            if line.upper().startswith("FIELD:"):
                field = line.split(":", 1)[-1].strip()
            elif line.upper().startswith("REASON:"):
                reason = line.split(":", 1)[-1].strip()
        # Fallback: if model ignored the format, treat the whole response as field
        if not field:
            field = raw
        return field, reason
    except Exception as exc:
        logging.warning("AI classification failed for column '%s': %s", col_name, exc)
        return "", ""

INCOMING_DIR = Path(r"C:\Users\RileySasiain\Desktop\RawFiles")

CORRECTED_SUFFIX = "_corrected"
ALIAS_FILE_PREFIX = "manufactureralias"

# Skip the template, already-corrected files, and the alias crosswalk itself
def should_skip(filename: str) -> bool:
    stem = Path(filename).stem
    return (filename == "QuoteTemplate.xlsx"
            or stem.endswith(CORRECTED_SUFFIX)
            or stem.lower().startswith(ALIAS_FILE_PREFIX))

# Keyword patterns per template column — checked case-insensitively, first match wins.
# Catalog Number is listed before ManufacturerName so "Mfg Part Num" resolves correctly.
COLUMN_PATTERNS = {
    "Manufacturer Catalog Number": [
        "catalog_number", "catalog_num",
        "product code", "product_code", "prod code", "prod no", "prod num",
        "mfg part", "part number", "part num", "part no",
        "catalog number", "catalog num", "catalog no",
        "item number", "item num", "item no", "sku",
    ],
    "Manufacturer Catalog Description": ["description", "desc", "product name", "product"],
    # UOM Quantity must come before Proposed UOM so "Contract UOM Conv" doesn't
    # match the bare "uom" keyword and get misrouted to Proposed UOM.
    "Proposed UOM Quantity": [
        "uom qty", "uom quantity",
        "uom conv", "uom factor", "uom conversion",
        "conversion factor", "conv factor",
        "packaging qty", "packaging quantity", "packaging string",
        "pack qty", "pack quantity",
        "qoe",
    ],
    "Proposed UOM":                     ["puom", "uom", "unit of measure", "unit"],
    "Proposed UOM Price":               ["proposed pricing", "proposed price",
                                         "contract price", "contract pricing",
                                         "unit price", "unit pricing",
                                         "price", "pricing", "cost"],
    "ManufacturerName":                 ["vendor name", "manufacturer name", "mfg name",
                                         "supplier name"],
}

# Short single-word names that match ManufacturerName only when the ENTIRE column name matches
MANUFACTURER_EXACT_NAMES = {"vendor", "manufacturer", "mfg", "supplier"}

# Columns containing these words alongside "price/pricing" are delta/variance columns,
# not base price columns — exclude them from Proposed UOM Price matching
PRICE_EXCLUSION_KEYWORDS = {"increase", "decrease", "delta", "change", "variance", "difference"}

# Keywords that identify a quantity column as purchase/usage intent
PURCHASE_QTY_KEYWORDS = {"purchase", "usage", "order", "annual"}
# Keywords that identify any quantity-like column
QTY_IDENTIFIERS      = {"quant", "qty", "qoe", "usage", "annual", "conv", "factor"}


def _classify_content(series):
    """Inspect cell values to classify a column's likely template role.
    Returns one of: 'uom', 'uom_quantity', 'catalog_number', 'manufacturer_name', or None."""
    values = series.dropna().astype(str).str.strip()
    if values.empty:
        return None

    # Proposed UOM: 2–4 char purely-alphabetic codes (EA, BX, CA, PK, RL …)
    # Single-letter status flags like "V" are intentionally excluded.
    if all(re.match(r"^[A-Za-z]{2,4}$", v) for v in values):
        return "uom"

    # UOM Quantity is intentionally not detected from content — small numeric IDs
    # (contact keys, category keys, etc.) are indistinguishable from pack quantities.

    has_digit = sum(bool(re.search(r"\d", v)) for v in values) / len(values)
    has_space = sum(" " in v for v in values) / len(values)

    if has_digit >= 0.5:
        return "catalog_number"
    if has_digit < 0.2 or has_space >= 0.3:
        return "manufacturer_name"
    return None


def _ask_ai_price_preference(candidates_info: list) -> tuple[str, str]:
    """Ask the AI which of several candidate price columns is the per-unit/each price.
    candidates_info is a list of (col_name, sample_values) tuples.
    Returns (selected_col_name_or_'uncertain', rationale)."""
    client = _get_ai_client()
    if client is None:
        return "uncertain", "AI client not available"

    col_lines = "\n".join(
        f'  - "{col}": sample values {samples}'
        for col, samples in candidates_info
    )
    prompt = (
        "You are helping select the correct price column for a healthcare vendor quote template.\n\n"
        "The target field is 'Proposed UOM Price' — the contract price per unit of measure.\n"
        "If one column is a per-each / per-individual-unit price and another is per-UOM "
        "(case, box, pack), prefer the per-each price.\n\n"
        f"Candidate columns:\n{col_lines}\n\n"
        "Which column best represents the per-unit or per-each price?\n"
        "Reply in exactly this format:\n"
        "COLUMN: <exact column name from the list, or 'uncertain' if you cannot determine>\n"
        "REASON: <one sentence explaining why>"
    )
    try:
        response = client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            max_completion_tokens=150,
        )
        raw = response.choices[0].message.content.strip()
        col_pick, reason = "uncertain", ""
        for line in raw.splitlines():
            if line.upper().startswith("COLUMN:"):
                col_pick = line.split(":", 1)[-1].strip()
            elif line.upper().startswith("REASON:"):
                reason = line.split(":", 1)[-1].strip()
        return col_pick, reason
    except Exception as exc:
        logging.warning("AI price preference failed: %s", exc)
        return "uncertain", f"AI error: {exc}"


def _disambiguate_price_column(candidates: list, df) -> tuple[str, list]:
    """Select the best Proposed UOM Price column from multiple candidates.

    3-step process:
      1. Prefer the column that is fully populated (no blank cells).
      2. Ask the AI which is a per-unit/each price rather than a UOM/case price.
      3. Fall back to the column with the lowest median non-zero price.

    Returns (selected_column_name, log_lines)."""
    log = [f"    Candidates: {candidates}"]

    # Step 1 — population check
    fully_populated = [
        col for col in candidates
        if pd.to_numeric(df[col], errors="coerce").notna().all()
        or df[col].replace("", pd.NA).notna().all()
    ]
    if len(fully_populated) == 1:
        log.append(f"    Step 1 (population): '{fully_populated[0]}' is fully populated — selected")
        return fully_populated[0], log
    elif len(fully_populated) == 0:
        log.append("    Step 1 (population): no fully populated columns — all candidates continue")
        working = candidates
    else:
        log.append(f"    Step 1 (population): multiple fully populated {fully_populated} — proceeding to Step 2")
        working = fully_populated

    # Step 2 — AI unit/each preference
    if _OPENAI_AVAILABLE:
        samples_info = [(col, df[col].dropna().head(5).tolist()) for col in working]
        ai_pick, ai_reason = _ask_ai_price_preference(samples_info)
        matched = next((col for col in working if col == ai_pick), None)
        if matched:
            log.append(f"    Step 2 (AI):        selected '{matched}'")
            log.append(f"    AI rationale:       {ai_reason}")
            return matched, log
        else:
            log.append(f"    Step 2 (AI):        uncertain ('{ai_pick}') — {ai_reason}")
    else:
        log.append("    Step 2 (AI):        skipped — AI not available")

    # Step 3 — lowest median non-zero price
    best_col, best_median = None, float("inf")
    for col in working:
        vals = pd.to_numeric(df[col], errors="coerce")
        non_zero = vals[(vals > 0) & vals.notna()]
        if not non_zero.empty:
            med = non_zero.median()
            log.append(f"    Step 3 (low price): '{col}' median non-zero = {med:.4f}")
            if med < best_median:
                best_median, best_col = med, col
    if best_col:
        log.append(f"    Step 3 (low price): selected '{best_col}' (median {best_median:.4f})")
        return best_col, log

    # Ultimate fallback
    log.append(f"    Fallback:           selected first candidate '{working[0]}'")
    return working[0], log


def _map_columns(df):
    """Return a rename dict by matching each incoming column to a template column.
    Uses column-name keyword matching first, then content-based fallback."""
    columns = df.columns
    rename = {}
    used_targets = set()

    # Pass 1: keyword matching on column names.
    # Price columns are collected separately so we can scan ALL columns before
    # committing to one — multiple columns often contain "price" in their name.
    price_candidates = []   # list of (col_name, matched_keyword)

    for col in columns:
        col_lower = col.lower()
        for target, keywords in COLUMN_PATTERNS.items():
            if target in used_targets:
                continue
            if any(kw in col_lower for kw in keywords):
                # Extra guard for price: skip columns that look like delta/change columns
                if (target == "Proposed UOM Price"
                        and any(ex in col_lower for ex in PRICE_EXCLUSION_KEYWORDS)):
                    continue
                # Defer price assignment — collect all candidates first
                if target == "Proposed UOM Price":
                    matched_kw = next((kw for kw in keywords if kw in col_lower), "")
                    price_candidates.append((col, matched_kw))
                    break
                rename[col] = target
                used_targets.add(target)
                break

    # Price column selection — now that all columns have been scanned
    price_disambiguation_log = []
    if len(price_candidates) == 1:
        col, _ = price_candidates[0]
        rename[col] = "Proposed UOM Price"
        used_targets.add("Proposed UOM Price")
        price_disambiguation_log.append(f"    Single match: '{col}' — selected directly")
        logging.info("Price column: '%s' (single match)", col)
    elif len(price_candidates) > 1:
        # Group candidates by the keyword that matched them
        from collections import defaultdict
        by_kw: dict = defaultdict(list)
        for col, kw in price_candidates:
            by_kw[kw].append(col)

        same_reason_groups = {kw: cols for kw, cols in by_kw.items() if len(cols) > 1}

        if same_reason_groups:
            # Multiple columns matched the SAME keyword → run disambiguation
            ambiguous = [col for cols in same_reason_groups.values() for col in cols]
            unique    = [col for col, kw in price_candidates if kw not in same_reason_groups]
            logging.info("Price column: %d candidate(s) matched same keyword — running disambiguation on %s",
                         len(ambiguous), ambiguous)
            price_disambiguation_log.append(
                f"    {len(price_candidates)} candidate(s) found — same-keyword match on: "
                + ", ".join(f'"{k}"' for k in same_reason_groups)
            )
            if unique:
                # A column matched a more specific/unique keyword — prefer it
                selected = unique[0]
                price_disambiguation_log.append(
                    f"    '{selected}' matched a unique keyword — selected over ambiguous group"
                )
                logging.info("Price column: '%s' selected (unique keyword match)", selected)
            else:
                selected, dis_log = _disambiguate_price_column(ambiguous, df)
                price_disambiguation_log.extend(dis_log)
        else:
            # Every candidate matched a different keyword — prefer the one whose
            # keyword appears earliest in the COLUMN_PATTERNS list (most specific)
            price_keywords = COLUMN_PATTERNS["Proposed UOM Price"]
            def _kw_priority(col_kw_pair):
                col, kw = col_kw_pair
                return price_keywords.index(kw) if kw in price_keywords else len(price_keywords)
            selected = min(price_candidates, key=_kw_priority)[0]
            matched_kw = next(kw for c, kw in price_candidates if c == selected)
            price_disambiguation_log.append(
                f"    {len(price_candidates)} candidate(s), each with a unique keyword — "
                f"'{selected}' selected (keyword \"{matched_kw}\" has highest specificity)"
            )
            logging.info("Price column: '%s' selected (highest-specificity keyword match)", selected)

        rename[selected] = "Proposed UOM Price"
        used_targets.add("Proposed UOM Price")
        logging.info("Price column final selection: '%s'", selected)

    # Pass 2: quantity columns (keyword + purchase/usage intent)
    unmapped = [c for c in columns if c not in rename]
    qty_cols = [c for c in unmapped if any(kw in c.lower() for kw in QTY_IDENTIFIERS)]
    if len(qty_cols) == 1:
        col = qty_cols[0]
        target = ("Proposed Purchase Quantity"
                  if any(kw in col.lower() for kw in PURCHASE_QTY_KEYWORDS)
                  else "Proposed UOM Quantity")
        rename[col] = target
        used_targets.add(target)
    else:
        for col in qty_cols:
            target = ("Proposed Purchase Quantity"
                      if any(kw in col.lower() for kw in PURCHASE_QTY_KEYWORDS)
                      else "Proposed UOM Quantity")
            if target not in used_targets:
                rename[col] = target
                used_targets.add(target)

    # Pass 2b: unlabeled columns — if UOM Quantity still unmatched, look for a
    # placeholder-named column (col_N) whose values are all small positive integers.
    # Vendor files sometimes leave the quantity column header blank.
    if "Proposed UOM Quantity" not in used_targets:
        for col in [c for c in columns if c not in rename and re.match(r"^col_\d+$", c)]:
            vals = df[col].dropna()
            if len(vals) > 0:
                numeric_vals = pd.to_numeric(vals, errors="coerce").dropna()
                if (len(numeric_vals) / len(vals) >= 0.9
                        and (numeric_vals >= 0).all()
                        and (numeric_vals % 1 == 0).all()):
                    rename[col] = "Proposed UOM Quantity"
                    used_targets.add("Proposed UOM Quantity")
                    logging.info("Unlabeled column '%s' detected as Proposed UOM Quantity "
                                 "(positive integers, no header)", col)
                    break

    # Pass 3: exact single-word names for ManufacturerName
    if "ManufacturerName" not in used_targets:
        for col in columns:
            if col not in rename and col.strip().lower() in MANUFACTURER_EXACT_NAMES:
                rename[col] = "ManufacturerName"
                used_targets.add("ManufacturerName")
                break

    # Pass 4: content-based fallback for still-unmapped columns.
    # UOM: 2-4 char alpha codes are distinctive enough to detect reliably.
    # UOM Quantity: excluded — small numeric IDs can't be distinguished from pack quantities.
    CONTENT_TARGETS = {
        "uom":               "Proposed UOM",
        "catalog_number":    "Manufacturer Catalog Number",
        "manufacturer_name": "ManufacturerName",
    }
    still_unmapped = [c for c in columns if c not in rename]
    for col in still_unmapped:
        guess = _classify_content(df[col])
        target = CONTENT_TARGETS.get(guess)
        if target and target not in used_targets:
            rename[col] = target
            used_targets.add(target)
            logging.info("Content-based mapping: '%s' -> %s", col, target)

    # Pass 5: validate UOM vs UOM Quantity using the hard content rules —
    # Proposed UOM is always letters-only; Proposed UOM Quantity is always numeric.
    # Correct any swap that keyword matching may have introduced.
    uom_col      = next((c for c, t in rename.items() if t == "Proposed UOM"), None)
    uom_qty_col  = next((c for c, t in rename.items() if t == "Proposed UOM Quantity"), None)
    if uom_col and _classify_content(df[uom_col]) == "uom_quantity":
        rename[uom_col] = "Proposed UOM Quantity"
        logging.info("Content correction: '%s' reassigned to Proposed UOM Quantity", uom_col)
        if uom_qty_col and _classify_content(df[uom_qty_col]) == "uom":
            rename[uom_qty_col] = "Proposed UOM"
            logging.info("Content correction: '%s' reassigned to Proposed UOM", uom_qty_col)
    elif uom_qty_col and _classify_content(df[uom_qty_col]) == "uom":
        rename[uom_qty_col] = "Proposed UOM"
        logging.info("Content correction: '%s' reassigned to Proposed UOM", uom_qty_col)
        if uom_col and _classify_content(df[uom_col]) == "uom_quantity":
            rename[uom_col] = "Proposed UOM Quantity"
            logging.info("Content correction: '%s' reassigned to Proposed UOM Quantity", uom_col)

    # Pass 6: AI fallback — ask Azure OpenAI to classify anything still unmapped
    ai_mapped = []
    ai_calls  = []   # full record of every AI interaction for the summary
    still_unmapped = [c for c in columns if c not in rename]
    if still_unmapped and _OPENAI_AVAILABLE:
        logging.info("Sending %d unrecognised column(s) to AI for classification: %s",
                     len(still_unmapped), still_unmapped)
        for col in still_unmapped:
            sample = df[col].dropna().head(5).tolist()
            result, reason = _ask_ai_to_classify_column(col, sample)
            mapped = False
            if result and result.lower() != "ignore" and result in TEMPLATE_COLUMNS:
                if result not in used_targets:
                    rename[col] = result
                    used_targets.add(result)
                    ai_mapped.append(col)
                    mapped = True
                    logging.info("  [AI] '%s' -> %s", col, result)
                else:
                    logging.info("  [AI] '%s' -> %s (already filled — skipped)", col, result)
            else:
                logging.info("  [AI] '%s' -> ignored (%s)", col, result or "no response")
            ai_calls.append({
                "column":  col,
                "sample":  sample,
                "decision": result or "no response",
                "reason":  reason or "no rationale provided",
                "mapped":  mapped,
            })

    ignored = [c for c in columns if c not in rename]
    if ignored:
        logging.info("Columns not mapped (ignored): %s", ignored)
    return rename, {
        "ignored":                  ignored,
        "ai_mapped":                ai_mapped,
        "ai_calls":                 ai_calls,
        "price_disambiguation_log": price_disambiguation_log,
    }

# ---------------------------------------------------------------------------
# Layout normalisation — handles non-standard Excel files where data doesn't
# start in row 1 / column 1, headers span multiple rows, or data rows have
# continuation lines.
# ---------------------------------------------------------------------------

# Words likely to appear in header rows (used for header-row detection)
HEADER_SCAN_KEYWORDS = {
    "product", "item", "part", "catalog", "sku", "description", "desc",
    "uom", "unit", "price", "cost", "quantity", "qty", "volume", "size",
    "vendor", "manufacturer", "mfg", "supplier", "number", "num",
    "code", "name", "list", "contract", "pack", "ndc",
}


def _find_header_row(df, max_scan=20):
    """Return the index of the row with the most header-keyword matches.
    Defaults to row 0 if no row scores at least 2 hits."""
    best_idx, best_score = 0, 0
    for idx in range(min(max_scan, len(df))):
        score = sum(
            1 for cell in df.iloc[idx]
            if pd.notna(cell)
            and any(kw in str(cell).lower() for kw in HEADER_SCAN_KEYWORDS)
        )
        if score > best_score:
            best_score, best_idx = score, idx
    if best_score < 2:
        logging.info("No clear header row found — defaulting to row 0")
    return best_idx


def _merge_header_rows(df, header_idx):
    """Collapse sub-header rows (continuation rows that carry no numeric data
    and are mostly blank) into the main header row.
    Returns (data_start_index, list_of_header_strings)."""
    header = [str(v).strip() if pd.notna(v) else "" for v in df.iloc[header_idx]]
    data_start = header_idx + 1

    for i in range(header_idx + 1, min(header_idx + 4, len(df))):
        row = df.iloc[i]
        has_numbers = any(isinstance(c, (int, float)) and pd.notna(c) for c in row)
        non_null    = int(row.notna().sum())
        # A sub-header row is mostly blank and contains no numeric data values
        if not has_numbers and non_null <= max(2, int(len(row) * 0.4)):
            for col_idx, cell in enumerate(row):
                if pd.notna(cell) and str(cell).strip():
                    existing = header[col_idx]
                    header[col_idx] = (existing + " " + str(cell).strip()).strip()
            data_start = i + 1
        else:
            break

    # Give blank header slots a positional placeholder so column count is preserved
    header = [h if h else f"col_{i}" for i, h in enumerate(header)]
    return data_start, header


def _merge_continuation_rows(df):
    """A row is a continuation if it has fewer non-null values (≤ 2) than the
    row above it — its values are appended to the previous row's cells."""
    if df.empty or len(df) < 2:
        return df
    rows = []
    for _, row in df.iterrows():
        non_null      = int(row.notna().sum())
        prev_non_null = sum(1 for v in rows[-1].values() if pd.notna(v)) if rows else 0
        if rows and non_null <= 2 and non_null < prev_non_null:
            for col in df.columns:
                if pd.notna(row[col]):
                    prev = rows[-1][col]
                    rows[-1][col] = (
                        str(prev) + " " + str(row[col]) if pd.notna(prev) else row[col]
                    )
        else:
            rows.append(row.to_dict())
    return pd.DataFrame(rows, columns=df.columns).reset_index(drop=True)


def _normalize_layout(df_raw):
    """Pre-process a raw (header=None) DataFrame:
      1. Detect the true header row
      2. Merge multi-line header rows into one
      3. Drop entirely blank columns
      4. Merge data continuation rows
      5. Drop entirely blank rows
    Returns a clean DataFrame ready for column mapping."""
    header_idx = _find_header_row(df_raw)
    if header_idx > 0:
        logging.info("Header detected at row %d — skipping %d title row(s)",
                     header_idx, header_idx)

    data_start, headers = _merge_header_rows(df_raw, header_idx)
    logging.info("Data starts at row %d | Headers found: %s",
                 data_start,
                 [h for h in headers if not h.startswith("col_")])

    df = df_raw.iloc[data_start:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    df = df.dropna(axis=1, how="all")
    df = _merge_continuation_rows(df)
    df = df.dropna(how="all")
    return df


# Required output columns in order (must match QuoteTemplate.xlsx exactly)
TEMPLATE_COLUMNS = [
    "ManufacturerName",
    "Manufacturer Catalog Number",
    "Manufacturer Catalog Description",
    "Proposed UOM",
    "Proposed UOM Quantity",
    "Proposed UOM Price",
    "Proposed Purchase Quantity",
]

# These columns get $ and comma stripping, then numeric conversion
NUMERIC_COLUMNS = {"Proposed UOM Quantity", "Proposed UOM Price", "Proposed Purchase Quantity"}

# Unit of measure normalization — keys are lowercase, values are ANSI 2-char codes
UOM_MAP = {
    "case": "CA",
    "box":  "BX",
    "each": "EA",
    "pack": "PK",
    "roll": "RL",
}


def _normalize_uom(value):
    if pd.isna(value):
        return value
    raw = str(value).strip()
    normalized = UOM_MAP.get(raw.lower())
    if normalized:
        return normalized
    if raw.upper() in UOM_MAP.values():
        # Already an ANSI code (e.g. vendor file already says "EA") — pass through as-is
        return raw.upper()
    logging.warning("Unrecognized UOM value '%s' — left as-is", value)
    return value


def _clean_numeric(value):
    if pd.isna(value):
        return value
    cleaned = re.sub(r"[$,]", "", str(value).strip())
    try:
        return float(cleaned) if "." in cleaned else int(cleaned)
    except ValueError:
        return value


# ---------------------------------------------------------------------------
# Lattice layout detection & parsing — handles vendor files where the same
# Ref#/Price fields repeat across several side-by-side column blocks on the
# same rows (a "quick reference" price list), instead of one row per part.
# See TRANSFORMATION_SPEC.md §2a. Unlike transform_price_list.py's hardcoded
# per-vendor block positions, this detects the blocks dynamically from the
# header row so it isn't tied to one specific file's column layout.
# ---------------------------------------------------------------------------

REF_HEADER_RE = re.compile(r'\bref\s*#?\b', re.IGNORECASE)
LATTICE_NUM_RE = re.compile(r'^\$?\d[\d,]*\.?\d*$')


def _is_ref_header(cell) -> bool:
    if pd.isna(cell):
        return False
    return bool(REF_HEADER_RE.search(str(cell)))


def _is_price_header(cell) -> bool:
    if pd.isna(cell):
        return False
    c = str(cell).lower()
    if any(ex in c for ex in PRICE_EXCLUSION_KEYWORDS):
        return False
    return any(kw in c for kw in COLUMN_PATTERNS["Proposed UOM Price"])


def _is_description_header(cell) -> bool:
    if pd.isna(cell):
        return False
    c = str(cell).strip().lower()
    return "description" in c or c == "desc"


def _is_uom_header(cell) -> bool:
    if pd.isna(cell):
        return False
    c = str(cell).strip().lower()
    return c in ("u/m", "uom", "unit", "unit of measure") or "u/m" in c


def _lattice_is_price(p) -> bool:
    if pd.isna(p):
        return False
    if isinstance(p, (int, float)):
        return True
    return bool(LATTICE_NUM_RE.match(str(p).strip()))


def _lattice_clean_ref(raw) -> str:
    if pd.isna(raw):
        return ""
    s = str(raw).strip()
    return re.sub(r'^REF\s+', '', s, flags=re.IGNORECASE)


def _lattice_normalize_unit(u: str) -> str:
    if not u:
        return ""
    l = str(u).lower()
    if l.startswith("pc"):
        return "PC"
    if l.startswith("set"):
        return "SET"
    if l.startswith("kit"):
        return "KIT"
    if l.startswith("box"):
        return "BOX"
    return str(u).upper()


def _lattice_split_uom(uom_str):
    """'1 pc' -> (1, 'PC'); 'BOX' -> ('', 'BOX')"""
    if not uom_str:
        return "", ""
    m = re.match(r'^(\d+)\s*(.*)$', str(uom_str).strip())
    if m:
        qty = int(m.group(1))
        unit = m.group(2).strip() or uom_str
    else:
        qty, unit = "", uom_str
    return qty, _lattice_normalize_unit(unit)


def _detect_lattice_blocks(header_row: list) -> list:
    """Given one row's raw cell values, find repeated (ref_idx, price_idx)
    column blocks. Returns [] if this doesn't look like a lattice header —
    at least 2 repeated Ref#-like columns are required to call it a lattice
    rather than a coincidental single Ref#/Price pair."""
    ref_positions = [i for i, c in enumerate(header_row) if _is_ref_header(c)]
    if len(ref_positions) < 2:
        return []
    blocks = []
    for n, ref_idx in enumerate(ref_positions):
        next_ref = ref_positions[n + 1] if n + 1 < len(ref_positions) else len(header_row)
        price_idx = next(
            (i for i in range(ref_idx + 1, min(next_ref, ref_idx + 6))
             if i < len(header_row) and _is_price_header(header_row[i])),
            None,
        )
        if price_idx is not None:
            blocks.append((ref_idx, price_idx))
    return blocks if len(blocks) >= 2 else []


def _find_lattice_detail_sections(df_raw, start_idx: int) -> list:
    """Find header rows for a lower 'detailed' section that repeats
    periodically and carries Description/UOM, keyed on Ref# (spec §2a).
    Returns a list of dicts: {row_idx, ref_idx, description_idx, uom_idx}."""
    sections = []
    for idx in range(start_idx, len(df_raw)):
        row = df_raw.iloc[idx].tolist()
        ref_idx = next((i for i, c in enumerate(row) if _is_ref_header(c)), None)
        desc_idx = next((i for i, c in enumerate(row) if _is_description_header(c)), None)
        if ref_idx is not None and desc_idx is not None:
            uom_idx = next((i for i, c in enumerate(row) if _is_uom_header(c)), None)
            sections.append({"row_idx": idx, "ref_idx": ref_idx,
                              "description_idx": desc_idx, "uom_idx": uom_idx})
    return sections


def _parse_lattice_sheet(df_raw, header_idx: int, blocks: list):
    """Build the authoritative (ref -> price) universe from every lattice
    block, then left-join Description/UOM from any detailed section found
    lower on the sheet. Returns (df, info) — df already shaped to
    TEMPLATE_COLUMNS (ManufacturerName left blank, filled by the usual
    filename/alias fallback later in process_file)."""
    lattice_pairs = {}
    for _, row in df_raw.iloc[header_idx + 1:].iterrows():
        values = row.tolist()
        for ref_idx, price_idx in blocks:
            if ref_idx >= len(values) or price_idx >= len(values):
                continue
            ref, price = values[ref_idx], values[price_idx]
            if pd.isna(ref) or not _lattice_is_price(price):
                continue
            ref_clean = _lattice_clean_ref(ref)
            if ref_clean:
                lattice_pairs[ref_clean] = _clean_numeric(price)

    detail_map = {}
    sections = _find_lattice_detail_sections(df_raw, header_idx + 1)
    for i, section in enumerate(sections):
        end = sections[i + 1]["row_idx"] if i + 1 < len(sections) else len(df_raw)
        for r in range(section["row_idx"] + 1, end):
            row = df_raw.iloc[r].tolist()
            ref_idx = section["ref_idx"]
            if ref_idx >= len(row) or pd.isna(row[ref_idx]):
                continue
            ref_clean = _lattice_clean_ref(row[ref_idx])
            if not ref_clean or ref_clean not in lattice_pairs:
                continue  # detailed section only enriches parts already in the lattice
            desc_idx, uom_idx = section["description_idx"], section["uom_idx"]
            description = row[desc_idx] if desc_idx is not None and desc_idx < len(row) else None
            uom_raw = row[uom_idx] if uom_idx is not None and uom_idx < len(row) else None
            if not pd.isna(description) or not pd.isna(uom_raw):
                detail_map[ref_clean] = {"description": description, "uom": uom_raw}

    out_rows = []
    enriched = 0
    for ref_clean, price in lattice_pairs.items():
        detail = detail_map.get(ref_clean, {})
        if detail:
            enriched += 1
        uom_qty, uom_unit = _lattice_split_uom(detail.get("uom") or "")
        description = detail.get("description")
        out_rows.append({
            "ManufacturerName": None,
            "Manufacturer Catalog Number": ref_clean,
            "Manufacturer Catalog Description":
                str(description).strip() if not pd.isna(description) else "",
            "Proposed UOM": uom_unit,
            "Proposed UOM Quantity": uom_qty,
            "Proposed UOM Price": price,
            "Proposed Purchase Quantity": None,
        })
    df = pd.DataFrame(out_rows, columns=TEMPLATE_COLUMNS)
    info = {
        "blocks": blocks,
        "detail_sections_found": len(sections),
        "parts_enriched": enriched,
        "parts_total": len(out_rows),
    }
    return df, info


def _try_parse_lattice(df_raw):
    """If df_raw's detected header row shows the repeated Ref#/Price block
    pattern, parse it as a lattice layout. Returns (df, info), or None if
    this file doesn't look like a lattice."""
    header_idx = _find_header_row(df_raw)
    header_row = df_raw.iloc[header_idx].tolist()
    blocks = _detect_lattice_blocks(header_row)
    if not blocks:
        return None
    logging.info("Lattice layout detected: %d repeated Ref#/Price block(s) at row %d",
                 len(blocks), header_idx)
    return _parse_lattice_sheet(df_raw, header_idx, blocks)


def _load_manufacturer_aliases(folder: Path) -> list:
    """Read the manufacturer alias crosswalk from the watched folder.
    Column A = alias string to search in filename, Column B = clean name.
    Returns list of (alias, clean_name) sorted longest-first for specificity."""
    for f in folder.glob("manufactureralias*"):
        if f.suffix.lower() not in (".xlsx", ".xls", ".csv"):
            continue
        try:
            df = pd.read_csv(f) if f.suffix.lower() == ".csv" else pd.read_excel(f, header=0)
            pairs = [
                (str(row.iloc[0]).strip(), str(row.iloc[1]).strip())
                for _, row in df.iterrows()
                if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1])
            ]
            pairs.sort(key=lambda x: len(x[0]), reverse=True)
            logging.info("Loaded %d manufacturer aliases from %s", len(pairs), f.name)
            return pairs
        except Exception as exc:
            logging.warning("Could not load alias file %s: %s", f.name, exc)
    return []


def _lookup_manufacturer_from_filename(filename: str, aliases: list) -> str:
    """Search each alias (col A) as a substring of the filename.
    Returns the clean name (col B) for the first/longest match, or empty string."""
    name_lower = Path(filename).stem.lower()
    for alias, clean_name in aliases:
        if alias.lower() in name_lower:
            logging.info("Manufacturer alias match: '%s' in '%s' -> '%s'",
                         alias, filename, clean_name)
            return clean_name
    return ""


SUMMARY_SUFFIX = "_corrected_summary"

def _write_summary(out_path: str, lines: list[str]) -> None:
    """Write a plain-text processing summary alongside the corrected file."""
    summary_path = Path(out_path).with_suffix("").parent / (
        Path(out_path).stem.replace(CORRECTED_SUFFIX, "") + SUMMARY_SUFFIX + ".txt"
    )
    try:
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        logging.info("Summary written: %s", summary_path)
    except Exception as exc:
        logging.warning("Could not write summary file: %s", exc)


def _best_sheet(filepath: Path) -> str | int:
    """Return the sheet name of the visible sheet with the highest header-keyword score.
    Hidden and very-hidden sheets are skipped entirely.
    Falls back to sheet index 0 if nothing else can be determined."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        wb.close()

        if not visible:
            logging.warning("No visible sheets found — falling back to first sheet")
            return 0

        if len(visible) == 1:
            logging.info("Single visible sheet: %r", visible[0])
            return visible[0]

        best_name, best_score = visible[0], -1
        for name in visible:
            try:
                df = pd.read_excel(filepath, sheet_name=name, header=None, nrows=20)
                header_idx = _find_header_row(df)
                score = sum(
                    1 for cell in df.iloc[header_idx]
                    if pd.notna(cell)
                    and any(kw in str(cell).lower() for kw in HEADER_SCAN_KEYWORDS)
                )
                logging.info("Sheet %r (visible) header score: %d", name, score)
                if score > best_score:
                    best_score, best_name = score, name
            except Exception:
                pass

        logging.info("Selected sheet: %r", best_name)
        return best_name
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# PDF extraction
#   Layer 1 — pdfplumber: fast text extraction for digitally created PDFs
#   Layer 2 — AI vision:  renders pages as images and asks GPT to extract
#                          the table; handles scanned and complex layouts
# ---------------------------------------------------------------------------

def _extract_pdf_pdfplumber(filepath: Path) -> "pd.DataFrame | None":
    """Extract tables from a text-based PDF using pdfplumber.
    Collects the product table across all pages; returns a raw DataFrame
    (header=None style, ready for _normalize_layout), or None if no tables found."""
    if not _PDFPLUMBER_AVAILABLE:
        return None
    try:
        all_rows = []
        col_count = None
        with pdfplumber.open(str(filepath)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables:
                    continue
                # Take the widest table on each page
                page_table = max(tables, key=lambda t: len(t[0]) if t else 0)
                if not page_table:
                    continue
                ncols = len(page_table[0])
                if col_count is None:
                    col_count = ncols
                    all_rows.extend(page_table)
                elif ncols == col_count:
                    # Continuation page — skip repeated header row if present
                    first_row_text = " ".join(str(c or "").lower() for c in page_table[0])
                    if any(kw in first_row_text for kw in HEADER_SCAN_KEYWORDS):
                        all_rows.extend(page_table[1:])
                    else:
                        all_rows.extend(page_table)
        if not all_rows or len(all_rows) < 3:
            logging.info("pdfplumber: no usable tables found in %s", filepath.name)
            return None
        # A product pricing table needs at least 4 columns to hold catalog number,
        # description, UOM, and price. Narrower tables are likely TOCs, coverage
        # definitions, or other non-pricing content — skip them and let AI vision handle it.
        if col_count is not None and col_count < 4:
            logging.info(
                "pdfplumber: widest table has only %d column(s) — not a pricing table, "
                "falling back to AI vision", col_count
            )
            return None
        df = pd.DataFrame(all_rows)
        df = df.replace({None: pd.NA, "": pd.NA})
        logging.info("pdfplumber: extracted %d rows x %d cols", len(df), len(df.columns))
        return df
    except Exception as exc:
        logging.warning("pdfplumber extraction failed for %s: %s", filepath.name, exc)
        return None


def _extract_pdf_ai_vision(filepath: Path) -> "pd.DataFrame | None":
    """Convert PDF pages to images and ask GPT-5.4 to extract the table as CSV.
    Fallback for scanned PDFs and layouts where pdfplumber finds nothing."""
    client = _get_ai_client()
    if client is None:
        logging.warning("AI client not available — cannot use vision fallback for PDF")
        return None
    if not _PYMUPDF_AVAILABLE:
        logging.warning("PyMuPDF not installed — cannot convert PDF to images. "
                        "Run: pip install pymupdf")
        return None
    import base64
    import io
    try:
        doc = fitz.open(str(filepath))
        all_frames = []
        first_headers = None
        MAX_PAGES = 3
        logging.info("AI vision: rendering %d page(s) of %s (this may take a moment)",
                     min(MAX_PAGES, len(doc)), filepath.name)
        for page_num in range(min(MAX_PAGES, len(doc))):
            page = doc[page_num]
            pix = page.get_pixmap(dpi=150)
            b64 = base64.b64encode(pix.tobytes("png")).decode()
            if first_headers:
                prompt_text = (
                    f"Page {page_num + 1} of a vendor price list PDF.\n"
                    f"Extract product rows continuing from the previous page.\n"
                    f"Use these exact column headers: {first_headers}\n"
                    "Return CSV with these headers on row 1, then all data rows.\n"
                    "No markdown fences, no explanation. "
                    "If no table rows are on this page, reply: NO_TABLE"
                )
            else:
                prompt_text = (
                    "This is a vendor price list PDF page.\n"
                    "Extract the product/pricing table as plain CSV.\n"
                    "- First row: column headers exactly as shown in the table\n"
                    "- Remaining rows: all product data rows\n"
                    "- Comma delimiter; quote any field that contains a comma\n"
                    "- No markdown code fences, no explanation\n"
                    "- If no product table is visible on this page, reply: NO_TABLE"
                )
            try:
                response = client.chat.completions.create(
                    model=AZURE_DEPLOYMENT,
                    messages=[{
                        "role": "user",
                        "content": [
                            {"type": "image_url",
                             "image_url": {"url": f"data:image/png;base64,{b64}",
                                           "detail": "high"}},
                            {"type": "text", "text": prompt_text},
                        ]
                    }],
                    max_completion_tokens=4000,
                )
            except Exception as exc:
                logging.warning("AI vision call failed for page %d: %s", page_num + 1, exc)
                continue
            csv_text = response.choices[0].message.content.strip()
            # Strip markdown code fences if the model added them
            if csv_text.startswith("```"):
                csv_text = "\n".join(
                    ln for ln in csv_text.splitlines() if not ln.startswith("```")
                )
            if csv_text.strip() == "NO_TABLE":
                logging.info("AI vision: page %d has no table rows", page_num + 1)
                continue
            try:
                df_page = pd.read_csv(io.StringIO(csv_text))
                if first_headers is None:
                    first_headers = df_page.columns.tolist()
                all_frames.append(df_page)
                logging.info("AI vision: page %d -> %d rows", page_num + 1, len(df_page))
            except Exception as exc:
                logging.warning("Could not parse AI vision CSV (page %d): %s", page_num + 1, exc)
        doc.close()
        if not all_frames:
            return None
        combined = pd.concat(all_frames, ignore_index=True)
        logging.info("AI vision: total %d rows extracted from %s", len(combined), filepath.name)
        return combined
    except Exception as exc:
        logging.error("AI vision extraction failed for %s: %s", filepath.name, exc)
        return None


def _extract_pdf(filepath: Path) -> "tuple[pd.DataFrame | None, str]":
    """Orchestrate PDF extraction: pdfplumber first, AI vision fallback.

    Returns (df, method) where:
      'pdfplumber' — df is raw (header=None style), pass to _normalize_layout
      'ai_vision'  — df already has column headers, skip normalization
      'failed'     — df is None, file could not be extracted
    """
    if _PDFPLUMBER_AVAILABLE:
        df = _extract_pdf_pdfplumber(filepath)
        if df is not None:
            return df, "pdfplumber"
        logging.info("pdfplumber found no tables — trying AI vision fallback")
    else:
        logging.warning("pdfplumber not installed — trying AI vision. "
                        "Install with: pip install pdfplumber")
    if _OPENAI_AVAILABLE:
        df = _extract_pdf_ai_vision(filepath)
        if df is not None:
            return df, "ai_vision"
        logging.warning("AI vision extraction also failed for %s", filepath.name)
    else:
        logging.warning("AI not available and pdfplumber failed — cannot process %s", filepath.name)
    return None, "failed"


def process_file(filepath: Path) -> bool:
    """Map columns, clean numerics, and write corrected file. Returns True on success."""
    ext = filepath.suffix.lower()
    if ext not in (".xlsx", ".pdf"):
        return False
    if should_skip(filepath.name):
        logging.debug("Skipping: %s", filepath.name)
        return False

    logging.info("Processing: %s", filepath.name)
    try:
        # Small delay to ensure the file is fully written before reading
        time.sleep(0.5)

        pdf_extraction_method = None
        lattice_info = None

        if ext == ".pdf":
            df_raw, pdf_extraction_method = _extract_pdf(filepath)
            if df_raw is None:
                logging.error("Could not extract data from PDF: %s — file skipped", filepath.name)
                return False
            sheet = f"PDF — {pdf_extraction_method}"
            if pdf_extraction_method == "pdfplumber":
                # Raw rows (header=None style) — run through normal layout normalisation
                df = _normalize_layout(df_raw)
            else:
                # AI vision already produced a structured DataFrame with column headers
                df = df_raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        else:
            sheet = _best_sheet(filepath)
            df_raw = pd.read_excel(filepath, sheet_name=sheet, header=None)
            lattice_result = _try_parse_lattice(df_raw)
            if lattice_result is not None:
                df, lattice_info = lattice_result
            else:
                df = _normalize_layout(df_raw)

        if df.empty:
            logging.warning("Skipping empty file: %s", filepath.name)
            return False

        if lattice_info is None and df.columns.duplicated().any():
            dup_names = sorted(set(df.columns[df.columns.duplicated()]))
            logging.error(
                "Skipping %s: repeated column headers %s detected, but this doesn't "
                "match the known lattice pattern (repeated Ref#/Price column blocks — "
                "see TRANSFORMATION_SPEC.md §2a). Try transform_price_list.py with an "
                "explicit layout mapping instead.",
                filepath.name, dup_names,
            )
            return False

        if lattice_info is not None:
            logging.info(
                "Lattice parse: %d block(s), %d part(s) found, %d enriched from %d "
                "detailed section(s)",
                len(lattice_info["blocks"]), lattice_info["parts_total"],
                lattice_info["parts_enriched"], lattice_info["detail_sections_found"],
            )
            col_map, map_info = {}, {
                "ignored": [], "ai_mapped": [], "ai_calls": [],
                "price_disambiguation_log": [], "lattice_info": lattice_info,
            }
        else:
            # Rename incoming columns to template names
            col_map, map_info = _map_columns(df)
            logging.info("Column mapping: %s", col_map)
            df = df.rename(columns=col_map)

        # If no ManufacturerName column was found (or it's entirely blank),
        # try to derive it from the filename using the alias crosswalk
        mfr_source = "mapped from column"
        mfr_missing = ("ManufacturerName" not in df.columns
                        or df["ManufacturerName"].isna().all())
        if mfr_missing:
            aliases = _load_manufacturer_aliases(filepath.parent)
            mfr_name = _lookup_manufacturer_from_filename(filepath.name, aliases)
            if mfr_name:
                df["ManufacturerName"] = mfr_name
                mfr_source = f"alias lookup — populated with \"{mfr_name}\""
            else:
                mfr_source = "not found — left blank"
                logging.warning("No manufacturer name found for %s", filepath.name)

        # Normalize UOM to ANSI 2-char codes (skip for lattice rows — already
        # normalized to PC/SET/KIT/BOX by the lattice-specific parser, spec §5)
        uom_change_count = 0
        if lattice_info is None and "Proposed UOM" in df.columns:
            uom_before = df["Proposed UOM"].copy()
            df["Proposed UOM"] = df["Proposed UOM"].apply(_normalize_uom)
            uom_change_count = int((uom_before != df["Proposed UOM"]).sum())

        # Add any template columns that are absent in the incoming file
        for col in TEMPLATE_COLUMNS:
            if col not in df.columns:
                df[col] = None

        # Reorder columns to exactly match the template
        df = df[TEMPLATE_COLUMNS]

        # Strip $ and commas from numeric columns
        for col in NUMERIC_COLUMNS:
            df[col] = df[col].apply(_clean_numeric)

        # Replace any remaining non-numeric text in Proposed UOM Price with 0.
        # Values like "FREE", "N/A", "TBD", etc. would cause upload failures.
        price_col = "Proposed UOM Price"
        non_numeric_mask = (
            pd.to_numeric(df[price_col], errors="coerce").isna()
            & df[price_col].notna()
        )
        price_text_replaced = []
        if non_numeric_mask.any():
            price_text_replaced = df.loc[non_numeric_mask, price_col].unique().tolist()
            count = int(non_numeric_mask.sum())
            logging.info(
                "Proposed UOM Price: replaced %d non-numeric value(s) with 0 — %s",
                count, price_text_replaced
            )
            df.loc[non_numeric_mask, price_col] = 0

        # Drop fully empty rows, then fill blank numeric fields with 0 on populated rows
        rows_before_drop = len(df)
        df = df.dropna(how="all")
        rows_dropped = rows_before_drop - len(df)
        for col in NUMERIC_COLUMNS:
            df[col] = df[col].fillna(0)

        out_name = filepath.stem + CORRECTED_SUFFIX + ".xlsx"
        out_path = os.path.join(str(filepath.parent), out_name)
        df.to_excel(out_path, index=False, sheet_name="Data Template ")
        logging.info("Saved: %s", out_path)

        # ── Build and write summary file ──────────────────────────────────────
        from datetime import datetime
        divider = "=" * 60
        thin    = "-" * 60
        summary = [
            divider,
            "  Quote File Watcher — Processing Summary",
            divider,
            f"  File:      {filepath.name}",
            f"  Processed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"  Sheet:     {sheet}",
            f"  Output:    {out_name}",
            "",
        ]

        if pdf_extraction_method:
            summary += [
                thin,
                "  PDF EXTRACTION",
                thin,
                f"    Method:  {pdf_extraction_method}",
            ]
            if pdf_extraction_method == "ai_vision":
                summary.append(
                    "    Note:    AI vision was used — verify data accuracy before uploading"
                )
            summary.append("")

        if map_info.get("lattice_info"):
            li = map_info["lattice_info"]
            summary += [
                thin,
                "  LATTICE LAYOUT (positional, no column mapping — spec §2a)",
                thin,
                f"    Repeated Ref#/Price blocks: {len(li['blocks'])} at columns {li['blocks']}",
                f"    Parts found (lattice = authoritative universe): {li['parts_total']}",
                f"    Detailed section(s) found lower on the sheet: {li['detail_sections_found']}",
                f"    Parts enriched with Description/UOM: {li['parts_enriched']}",
            ]
        else:
            summary += [
                thin,
                "  COLUMN MAPPING",
                thin,
            ]
            if col_map:
                for src, tgt in col_map.items():
                    ai_tag = "  [AI]" if src in map_info["ai_mapped"] else ""
                    summary.append(f"    {src:<35} ->  {tgt}{ai_tag}")
            else:
                summary.append("    (no columns mapped)")

        if map_info["ignored"]:
            summary += ["", f"  Ignored (not mapped to any template field):"]
            for col in map_info["ignored"]:
                summary.append(f"    {col}")

        summary += [
            "",
            thin,
            "  MANUFACTURER NAME",
            thin,
            f"    {mfr_source}",
            "",
            thin,
            "  DATA TRANSFORMATIONS",
            thin,
        ]
        if uom_change_count:
            summary.append(f"    UOM normalization:    {uom_change_count} value(s) standardized to ANSI codes")
        else:
            summary.append( "    UOM normalization:    no changes needed")

        if price_text_replaced:
            summary.append(f"    Non-numeric prices:   {len(price_text_replaced)} unique value(s) replaced with 0 — {price_text_replaced}")
        else:
            summary.append( "    Non-numeric prices:   none found")

        if rows_dropped:
            summary.append(f"    Empty rows removed:   {rows_dropped} blank row(s) dropped")
        else:
            summary.append( "    Empty rows removed:   none")

        summary.append(    f"    Numeric zero-fill:    blank cells in price/qty columns set to 0")

        # Price disambiguation section — only included when multiple candidates were found
        if map_info.get("price_disambiguation_log"):
            summary += [
                "",
                thin,
                "  PROPOSED UOM PRICE — COLUMN SELECTION",
                thin,
            ]
            summary.extend(map_info["price_disambiguation_log"])

        # AI classification section — only included when the AI was consulted
        if map_info.get("ai_calls"):
            summary += [
                "",
                thin,
                "  AI COLUMN CLASSIFICATION",
                thin,
                "  The following columns were not recognised by the rule-based mapping",
                "  and were sent to the AI for classification.",
                "",
            ]
            for call in map_info["ai_calls"]:
                decision_line = (
                    f"  -> Mapped to: {call['decision']}"
                    if call["mapped"]
                    else f"  -> Decision:  {call['decision']} (not mapped)"
                )
                summary += [
                    f"  Column:       {call['column']}",
                    f"  Sample data:  {call['sample']}",
                    decision_line,
                    f"  AI rationale: {call['reason']}",
                    "",
                ]

        summary += [
            thin,
            "  OUTPUT",
            thin,
            f"    {len(df)} rows written to: {out_name}",
            divider,
        ]
        _write_summary(out_path, summary)

        return True

    except Exception as exc:
        logging.error("Failed to process %s: %s", filepath.name, exc)
        return False


class QuoteFileHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self._recently_processed = {}  # path -> last processed timestamp

    def _handle(self, path_str):
        lower = path_str.lower()
        if not (lower.endswith(".xlsx") or lower.endswith(".pdf")):
            return
        now = time.time()
        # Debounce: ignore repeat events for the same file within 3 seconds
        if now - self._recently_processed.get(path_str, 0) < 3:
            return
        self._recently_processed[path_str] = now
        process_file(Path(path_str))

    def on_created(self, event):
        logging.debug("Event on_created: %s", event.src_path)
        if not event.is_directory:
            self._handle(event.src_path)

    def on_modified(self, event):
        logging.debug("Event on_modified: %s", event.src_path)
        if not event.is_directory:
            self._handle(event.src_path)

    def on_moved(self, event):
        logging.debug("Event on_moved: %s -> %s", event.src_path, event.dest_path)
        if not event.is_directory:
            self._handle(event.dest_path)


def main():
    parser = argparse.ArgumentParser(description="Quote file format corrector watcher")
    parser.add_argument(
        "--process-existing",
        action="store_true",
        help="Process any .xlsx files already in the incoming folder on startup",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    if args.process_existing:
        logging.info("Processing existing files in %s ...", INCOMING_DIR)
        existing = sorted(
            list(INCOMING_DIR.glob("*.xlsx")) + list(INCOMING_DIR.glob("*.pdf"))
        )
        for f in existing:
            process_file(f)

    observer = Observer()
    observer.schedule(QuoteFileHandler(), str(INCOMING_DIR), recursive=False)
    observer.start()

    logging.info("Watching %s for new .xlsx and .pdf files ...", INCOMING_DIR)
    logging.info("Corrected files saved alongside originals as *_corrected.xlsx")
    logging.info("Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("Stopping watcher ...")
        observer.stop()
    observer.join()
    logging.info("Done.")


if __name__ == "__main__":
    main()
