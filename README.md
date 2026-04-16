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

**`01_fdic_tables_putexcel.do`** (Stata) / **`01_fdic_tables_r.R`** (R)  
Main scripts. Load the multiyear dataset (2019, 2021, 2023), merge the 160 Census Bureau BRR replicate weights from `hhmultiyear/hhrep19_23.csv`, and write each table to Excel. Layout exactly matches the published report: `Characteristic | 2019 | 2021 | 2023 | Difference (2023–2021)`. Significance (`*`) at the 10 percent level using BRR standard errors (Fay factor = 0.5, `combined.weights = TRUE` in R's `svrepdesign()`). To update for a future survey year, change `Y1`/`Y2`/`Y3` at the top and update the replicate weight file name.

**R packages required:** `haven`, `survey`, `dplyr`, `openxlsx`

---

### Table templates

Three progressively more advanced versions of reusable templates covering four table layouts found in the annual report:

| Type | Layout | Report examples |
|------|--------|-----------------|
| A | Binary outcome × years as columns × demographics as rows | Tables 1.1, 2.2, 3.1, 4.1 |
| B | Multiple binary outcomes as columns × demographics as rows, one year | Tables 3.3, 5.1 |
| C | Categorical distribution: rows = categories, one column of percentages | Tables 1.2, 3.2 |
| D | Rows = survey years, columns = ordered response categories | Tables 1.3, 1.4, 2.1, 2.3 |

**`03a_fdic_templates_putexcel.do`** (Stata) / **`03a_fdic_templates_r.R`** (R)  
Beginner-friendly templates for all four table types (A, B, C, D). Each type is a self-contained, heavily commented example showing the full pattern from data to Excel. The R version uses `survey` + `openxlsx` and is structured so a beginner can copy one function call to produce a new table with future data.

**`03b_fdic_templates_collect.do`** (Stata only)  
Intermediate version. Types B, C, and D are rewritten using `table` + `collect`, which is more concise and avoids manual matrix indexing. Type A keeps the loop+matrix approach because the diff column requires storing standard errors for a post-hoc z-test.

**`03c_fdic_templates_collect_reg.do`** (Stata only)  
Most advanced version. Type A is rewritten to use `svy: reg` (linear probability model). A single regression per demographic subgroup yields all year estimates, the difference, and its significance in one model — no post-hoc z-test needed.

---

## Requirements

**Stata scripts:** Stata 17 or later (the `collect` framework used in `03b` and `03c` requires Stata 17; `01` and `03a` require only Stata 15)

**R scripts:** R 4.0 or later with packages: `haven`, `survey`, `dplyr`, `openxlsx`

```r
install.packages(c("haven", "survey", "dplyr", "openxlsx"))
```

**Data:** `hhmultiyears.zip` downloaded from the FDIC website and extracted to `data/` (see Setup above)

**Output directory:** `output/` must exist in the repo root (already tracked in git via `.gitkeep`)
