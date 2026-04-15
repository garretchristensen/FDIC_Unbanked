# FDIC Unbanked Survey — Table Automation

Stata code to reproduce tables from the [FDIC National Survey of Unbanked and Underbanked Households](https://www.fdic.gov/analysis/household-survey/) annual report.

Data: `hh2023_analys.csv` from the 2023 survey public use file  
Weight: `hhsupwgt` (household supplement weight)  
Output: one `.xlsx` file per table section, written to `data/`

---

## Code files

### Analysis scripts

**`fdic_tables_putexcel.do`**  
Main script to reproduce the report tables for a given survey year. Loads the data, declares the survey design, recodes variables, computes weighted estimates (`svy: mean`) by demographic subgroup, and writes results to Excel using `putexcel`. Saves a `estimates_YEAR.dta` file for use in the differences script. To adapt for a new year, update the `YEAR` global and CSV filename in Section 0.

**`fdic_add_diffs_putexcel.do`**  
Run after `fdic_tables_putexcel.do` for a new survey year. Merges current and prior year estimates, computes year-over-year differences with two-sample z-tests, and writes the Diff and significance columns back into the Excel table files. Significance conventions: `*` p < .05, `**` p < .01.

---

### Table templates

Three progressively more advanced versions of reusable templates covering four table layouts found in the annual report:

| Type | Layout | Report examples |
|------|--------|-----------------|
| A | Binary outcome × years as columns × demographics as rows | Tables 1.1, 2.2, 3.1, 4.1 |
| B | Multiple binary outcomes as columns × demographics as rows, one year | Tables 3.3, 5.1 |
| C | Categorical distribution: rows = categories, one column of percentages | Tables 1.2, 3.2 |
| D | Rows = survey years, columns = ordered response categories | Tables 1.3, 1.4, 2.1, 2.3 |

**`fdic_templates_putexcel.do`**  
Beginner-friendly version. All four templates use the same three-step pattern: configure, estimate via `svy: mean` loop, export matrix to Excel with `putexcel`. Easiest to follow and modify.

**`fdic_templates_collect.do`**  
Intermediate version. Types B, C, and D are rewritten using `table` + `collect`, which is more concise and avoids manual matrix indexing. Type A keeps the loop+matrix approach because the diff column requires storing standard errors for a post-hoc z-test, which the `collect` framework does not support natively.

**`fdic_templates_collect_reg.do`**  
Most advanced version. Type A is rewritten to use `svy: reg` (linear probability model) instead of separate `svy: mean` calls per year. A single regression per demographic subgroup yields all year estimates, the difference, and its significance in one model — no post-hoc z-test needed. Types B, C, D unchanged from the collect version.

---

## Requirements

- Stata 17 or later
- 2023 FDIC survey public use data (`hh2023/`)
