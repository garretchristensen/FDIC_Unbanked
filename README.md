# FDIC Unbanked Survey — Table Automation

Stata code to reproduce tables from the [FDIC National Survey of Unbanked and Underbanked Households](https://www.fdic.gov/analysis/household-survey/) annual report.

---

## Setup

### 1. Download the data

Go to the [FDIC Household Survey page](https://www.fdic.gov/analysis/household-survey/) and download the **multiyear public use file** (`hhmultiyears.zip`). Save it to the `data/` folder in this repo.

### 2. Extract the zip

From the repo root:

```bash
unzip data/hhmultiyears.zip -d data/
```

This creates `data/hhmultiyear/` containing the analysis CSV and replicate weight files.

### 3. Load and prepare the data in Stata

Run this once to build the Stata dataset you will use for all table scripts:

```stata
set type double

* --- Load main multiyear analysis file ---
* Stata treats "." in the CSV as missing automatically for numeric columns.
import delimited using "data/hhmultiyear/hh_multiyear_analys.csv", clear

* --- Keep only supplement respondents ---
keep if hsupresp == 1

save "data/hhmultiyear_analys.dta", replace
```

Note: do **not** call `svyset` here. The table scripts set up the full BRR survey design (including replicate weights) themselves.

This dataset covers survey years 2009, 2011, 2013, 2015, 2017, 2019, 2021, and 2023.
Use `hryear4` to subset to a specific year or set of years.

---

## Running the table scripts

The scripts in `code/` should be run from the repo root with the working directory set there. They read from `data/hhmultiyear_analys.dta` (once it has been built per the setup above) and write `.xlsx` output to `output/`.

**To reproduce the published tables:**

Run `01_fdic_tables_putexcel.do` — produces one `.xlsx` per table section in `output/`, exactly matching the report layout (three year columns + Difference).

**To experiment with template approaches**, see the three template files described below.

---

## Code files

### Analysis scripts

**`01_fdic_tables_putexcel.do`**  
Main script. Loads the multiyear dataset (2019, 2021, 2023), merges the 160 Census Bureau BRR replicate weights from `hhmultiyear/hhrep19_23.csv`, and writes each table to Excel. Layout exactly matches the published report: `Characteristic | 2019 | 2021 | 2023 | Difference (2023–2021)`. Significance (`*`) at the 10 percent level using BRR standard errors (Fay factor = 0.5). To update for a future survey year, change `Y1`/`Y2`/`Y3` in Section 0 and update the replicate weight file name.

---

### Table templates

Three progressively more advanced versions of reusable templates covering four table layouts found in the annual report:

| Type | Layout | Report examples |
|------|--------|-----------------|
| A | Binary outcome × years as columns × demographics as rows | Tables 1.1, 2.2, 3.1, 4.1 |
| B | Multiple binary outcomes as columns × demographics as rows, one year | Tables 3.3, 5.1 |
| C | Categorical distribution: rows = categories, one column of percentages | Tables 1.2, 3.2 |
| D | Rows = survey years, columns = ordered response categories | Tables 1.3, 1.4, 2.1, 2.3 |

**`03a_fdic_templates_putexcel.do`**  
Beginner-friendly version. All four templates use the same three-step pattern: configure, estimate via `svy: mean` loop, export matrix to Excel with `putexcel`. Easiest to follow and modify.

**`03b_fdic_templates_collect.do`**  
Intermediate version. Types B, C, and D are rewritten using `table` + `collect`, which is more concise and avoids manual matrix indexing. Type A keeps the loop+matrix approach because the diff column requires storing standard errors for a post-hoc z-test, which the `collect` framework does not support natively.

**`03c_fdic_templates_collect_reg.do`**  
Most advanced version. Type A is rewritten to use `svy: reg` (linear probability model) instead of separate `svy: mean` calls per year. A single regression per demographic subgroup yields all year estimates, the difference, and its significance in one model — no post-hoc z-test needed. Types B, C, D unchanged from the collect version.

---

## Requirements

- Stata 15 or later
- `hhmultiyears.zip` downloaded from the FDIC website and extracted to `data/` (see Setup above)
- `output/` directory must exist in the repo root before running `01_`
