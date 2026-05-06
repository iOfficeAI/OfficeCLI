# Branch Strategy

## Current State (2026-05-06)
- **upstream/main**: tracks `iOfficeAI/OfficeCLI` main. Pull-only, never push.
- **main**: local mirror of `upstream/main` plus compliance docs.
- **feat/hwpx**: active working branch on `origin`. Phase 36 HWP/HWPX evidence work (compatibility corpus, round-trip catalog, visual thresholds, provider matrix, release gate, safe-save) lives here.
- **docs/structure-init-2026-05-06**: short-lived doc branch that introduced `structure/`, fast-forwarded into `feat/hwpx`.

## Branch Roles
- **upstream/main**: read-only upstream reference.
- **main**: clean mirror; rebased from `upstream/main`.
- **feat/hwpx**: primary working branch. All cli-jaw HWP/HWPX changes land here.
- **feature/\*** / **docs/\***: optional short-lived branches. Merge back into `feat/hwpx` (typically `--ff-only`) and delete.

## Sync Procedure
```bash
git fetch upstream
git checkout main
git merge --ff-only upstream/main
git checkout feat/hwpx
git rebase main          # or: git merge main --no-ff, depending on history needs
```

## Working with `structure/`
- See [`structure/INDEX.md`](structure/INDEX.md) for the doc map.
- Update `structure/*.md` whenever the corresponding source/schema/test surface changes (sync checklist in `structure/INDEX.md`).
- Doc-only changes should not modify `.cs`, `.ts`, `.js`, `.py`, or other source files.

## Rules
- Never push to `upstream`.
- All cli-jaw changes go on `feat/hwpx` (or short-lived branches that merge into it).
- HWP/HWPX claims must be evidence-gated per `docs/qa/phase-36-release-gate.md` — no DOCX parity language.
- Tag releases as `cjk-v{version}` on `feat/hwpx`.
