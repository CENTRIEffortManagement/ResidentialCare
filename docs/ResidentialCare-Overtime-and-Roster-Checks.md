# ResidentialCare overtime and roster checks

These system rules define the inputs, thresholds and results. The approved calculation location and data handoff are recorded in [the OvertimeRosterChecks implementation plan](OvertimeRosterChecks-Implementation-Plan.md). This version checks published rosters.

Default scope is employees covered by the Whiddon agreement in these roles and employment types:

- Roles: Assistant in Nursing (AIN), Enrolled Nurse (EN), Registered Nurse (RN), Care Support Worker (CS).
- Employment types: permanent full-time, permanent part-time, fixed-term and casual.

A dash in **Applies to** means the default scope applies. The column lists only conditions that change the result. Apply documented overtime exemptions as a shared eligibility check. Use payroll agreement stream or grade where it changes a threshold or eligibility.

| Rule | Applies to — differentiators only | Inputs | Test and result |
|---|---|---|---|
| **1A. Net work duration** | **Roles =** AIN, EN, RN | Net worked hours | Flag work **over 10 hours** in a shift. Excess hours = `max(0, net worked hours − 10)`. |
| **1B. Net work duration — CS** | **Roles =** CS. **Agreement stream** selects the limit. | Net worked hours, CS agreement stream, configured CS limit | Proposed limit is **8 hours** for Hospitality as a roster review threshold, or **10 hours** for Personal Care. Excess hours = `max(0, net worked hours − CS limit)`. |
| **2A. Minimum break** | — | Previous shift finish, next shift start | Flag a minimum-break issue if **gap hours < 8**. |
| **2B. Short-gap overtime** | **Employment type =** permanent full-time, permanent part-time, fixed-term. **Condition =** broken-shift exception does not apply. | Gap hours from 2A, broken-shift status | Flag the later shift for overtime review if **gap hours < 10**. Confirm the payable hours with payroll. |
| **3A. Fortnight days off — roster check** | — | Fortnight boundaries, complete duty records | Flag if **duty-free days < 4**. Days short = `max(0, 4 − duty-free days)`. |
| **3B. Fortnight days off — overtime review** | **Employment type =** permanent full-time, permanent part-time, fixed-term. | Shortfall from 3A | If **days short > 0**, flag the fortnight for overtime review. Confirm the affected shifts with payroll. |
| **4. Fortnight hours** | — | Fortnight boundaries from Settings Data, qualifying worked hours in chronological order | Flag work **over 76 hours**. Identify the portion of work above the threshold. |

## Measurement definitions

- **Net worked hours** cover the complete shift, including overnight work. Deduct unpaid meals once. Include meal periods during which Whiddon requires the employee to remain available to attend to client services.
- **Gap hours** are the elapsed hours from the previous worked shift's finish to the next worked shift's start. Use consecutive shifts for the same employee across units of the same employer.
- **Duty-free day** is provisionally defined as one whole local midnight-to-midnight calendar day entirely within the fortnight, with zero overlapping duty. A Monday 3 pm finish followed by a Wednesday 10 pm start gives **55 gap hours and one duty-free day**, Tuesday, assuming complete records and no intervening duty.
- **Fortnight start date** is the earliest `Date` in the configured BD, TE and JH-RY Settings Data `PermutationDimensions` tables. Use one common start date across units. Do not take the first worked shift of each employee as the anchor.
- **Fortnight excess hours** begin after 76 qualifying hours. If an employee starts a six-hour shift at 74 accumulated hours, two hours reach the threshold and four exceed it. Preserve every applicable overtime trigger and count overlapping intervals once.

## Decisions to confirm before costing

- Verify the payroll stream for CS. The agreement lists Care Service Employee under Hospitality. Confirm how that classification maps to the CS role in the data.
- Confirm what counts as duty and whether the proposed calendar-day definition matches payroll practice. Return an unknown result when duty records are incomplete.
- Confirm which hours or shifts attract overtime after a short gap or a days-off shortfall.
- The approved roster-review counter includes all net rostered hours once, including hours also identified as daily excess. Confirm monetary costing treatment separately with payroll.

Agreement reference: [Whiddon Group Agreement 2017 with undertakings](https://www.whiddon.com.au/wp-content/uploads/2025/08/Whiddon-Agreement.pdf), clauses 23.3–23.4, 25.1, 37.3, 37.5, 37.7 and 75–78.
