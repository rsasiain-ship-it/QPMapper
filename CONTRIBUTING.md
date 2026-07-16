# Contributing / Process

This document is the checklist every change to this repo follows — written
down so the process doesn't depend on any one person remembering it.

## One-time setup after cloning

Enable the repo's tracked git hooks (see [Automated checks](#automated-checks)):

```
git config core.hooksPath hooks
```

## Rules

1. **Spec and script stay in sync.** Any change to parsing, derivation, or
   filtering logic updates both `TRANSFORMATION_SPEC.md` and the relevant
   script in the same commit, cross-referenced by section number
   (e.g. "see spec §2a").

2. **Every add/remove/behavior change is logged in `CHANGELOG.md`**, in the
   same commit that makes the change. Note *what* changed and *why* — not
   just "updated script."

3. **No real data, ever.** This repo tracks code, spec, and empty templates
   only. Real vendor/customer pricing data, alias crosswalks, source PDFs,
   and processed outputs must never be committed — see `.gitignore`.

4. **No hardcoded secrets.** API keys and credentials are read from
   environment variables, never written as literals in source.

5. **Version branches, not direct commits to `main`.** New work happens on
   a version branch (`v1.1`, `v1.2`, ...). A version only becomes "released"
   — merged to `main` and tagged (e.g. `v1.1.0`) — once it's actually ready
   to execute. Until then it stays in the changelog's `[Unreleased]` section.

6. **Test against real data before calling something done.** Type-checking
   or a clean run isn't enough — run the script against an actual sample
   file (or a realistic synthetic one) and check the output.

## Automated checks

`hooks/pre-commit` (tracked in this repo — activate it with the one-time
setup command above) blocks a commit if it:

- Modifies `transform_price_list.py` or `quote_file_watcher.py` without also
  touching `CHANGELOG.md` in the same commit.
- Contains an obvious hardcoded secret pattern (a long literal string
  assigned to a variable named like `*_KEY`, `*_TOKEN`, `*_SECRET`, or
  `*_PASSWORD`).
- Stages a file that looks like real data despite `.gitignore` (e.g. a
  `ManufacturerAlias*` workbook, anything under `RawFiles/`, a `.pdf`) —
  a backstop in case someone force-adds with `git add -f`.

These are heuristics, not a substitute for judgment — review what you're
committing regardless.
