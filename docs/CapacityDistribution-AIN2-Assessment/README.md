# AIN2 Capacity Distribution assessment bundle

This folder collects the assessment and handoff evidence relevant to this chat. These are copies for review; the original documents and evidence remain in their existing locations. The authoritative Power Query source remains under `Workflows/ResidentialCare/CapacityDistribution/`.

## Current AIN2 assessment

- `implementation-status.md` — source-refactor status, validation performed, and work still awaiting workbook approval.
- `expected-differences.md` — expected output differences between baseline and corrected designs.
- `source-gate.json` — fresh extraction and source-of-truth checks for the exact AIN2 workbook copies.
- `copy-manifest.json` — original AIN-to-AIN2 copy scope and hashes.
- `snapshot-manifest.json` — hashes for the attributable source snapshots. The snapshots themselves remain under `outputs/capacity-distribution-ain2/snapshots/`.

## Related capacity assessments and context

- `AIN-B-Performance-Changes.md` — preceding AIN B performance review and source change.
- `Unit1-AIN-Contracted-Hours-and-Availability.md` — AIN contract-cap implementation handoff and limitations.
- `DATExx-Whiddon-Capacity-Shift-Availability-Rules.md` — upstream shift-availability rules.
- `Capacity-Determination-Lineage.mmd` — capacity dependency diagram.

## Important current limitation

The current AIN2 A1/B workbooks were synchronized by the user, but their embedded definitions and blank results have not been independently verified. This bundle is source and historical assessment evidence, not proof that the live AIN2 workbook outputs pass. Before further merge acceptance, re-extract the closed AIN2 workbooks and trace the first empty import. Upstream role filters must select `AIN` and relabel working rows to `AIN2`.
