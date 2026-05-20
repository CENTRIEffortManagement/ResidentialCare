# ResidentialCare

ResidentialCare Power Query workspace scaffold.

This repository is intentionally minimal. Product-specific workflows belong under `Workflows/`, while shared Power Query tooling is expected beside this repository at `../mcode-pq-tools`.

## Layout

- `pq.project.json` - project configuration consumed by shared PQ tools.
- `scripts/refresh-residentialcare.ps1` - validate-only-capable wrapper for the shared refresh entry point.
- `Workflows/` - future ResidentialCare workflow source.
- `docs/` - project notes and setup documentation.

Full refresh automation is not configured in this scaffold.

