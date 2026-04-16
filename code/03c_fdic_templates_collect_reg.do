/*============================================================================
  FDIC Unbanked Survey — Table Templates v3

  Changes from v2:
    - Template A now uses svy: reg (linear probability model) instead of
      separate svy: mean calls per year. One regression per demographic
      subgroup gives all year estimates, the diff, and its significance
      in a single model — no post-hoc z-test required.

  Why regression for Type A:
    - Setting diff_from as the base category (ib`diff_from'.year) makes
      _b[`diff_to'.year] exactly the diff we want, with its own SE and
      t-statistic from the model.
    - Year estimates for all other years come from _b[_cons] + _b[yr.year].
    - The t-stat from the regression implicitly pools variance across years
      within each subgroup, which is slightly more efficient than the
      independent-samples z-test used in v2.
    - This matches the way the team has been computing differences.

  Templates B, C, D: unchanged from v2.

  PREREQUISITES
  -------------
    * Stata 17 or later
    * Run the SETUP section once per Stata session
    * Output goes to output/ as .xlsx files (set via $output global)
============================================================================*/


*=============================================================================
* SETUP — Run once per session
*=============================================================================

* Drop all loaded data, programs, and matrices from any previous run
clear all
* Disable the "more" prompt so long output scrolls without pausing
set more off

global data   "data"     // input data directory (relative to repo root)
global output "output"   // output directory for .xlsx results

* Store new variables as double-precision floats to avoid rounding in proportions
set type double

* ---- Load data and set survey design --------------------------------------
* Keep all years needed for regression-based Type A diff estimates.
* Adjust the inlist() if you want a different year range.
use "${data}/hhmultiyear_analys.dta", clear
* Drop observations outside the analysis window; reduces memory and prevents
* spurious year coefficients in the regressions below
keep if inlist(hryear4, 2017, 2019, 2021, 2023)
* Declare the survey design using the household supplement weight; no explicit
* clustering/strata needed here because BRR replicate weights are handled separately
svyset [pw=hhsupwgt]

* ---- Recode 1=Yes / 2=No variables to 0/1 --------------------------------
* Each original variable codes 1=Yes, 2=No; this loop creates a _b version
* coded 1=Yes, 0=No, missing for any other value (skips, refusals, etc.)
foreach v of varlist                                                         ///
    huse12mo huse12cc huse12mt huse12rm huse12rmany                         ///
    husenowops husenowpp huse12afsc                                          ///
    huse12pdl huse12pwn huse12ral huse12atl huse12rto                       ///
    hcred12cc hcred12sc hcred12car hcred12hmln hcred12sl hcred12any         ///
    hcred12bnpl hcred12bnpldq huse12cryp                                    ///
    hbnkaccm1v2 hbnkaccm2v2 hbnkaccm3v2                                     ///
    hbnkaccm4v2 hbnkaccm5v2 hbnkaccm6v2 {
    * `capture` absorbs "variable already exists" errors on re-run
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* Create binary banking-status indicators from their source variables;
* `inlist` restricts to valid response codes, leaving skip/refusal as missing
gen byte unbanked          = (hunbnk == 1)          if inlist(hunbnk, 1, 2)
gen byte underbanked       = (hbankstatv6 == 2)     if inlist(hbankstatv6, 1, 2, 3)
gen byte cashonly_unbanked = (hunbnkcashonly == 1)  if inlist(hunbnkcashonly, 1, 2)

* ---- Year variable ---------------------------------------------------------
* hryear4 is already in the multiyear data; rename for consistency with templates.
* A short name also makes factor-variable notation cleaner (e.g. ib2021.year)
rename hryear4 year
label variable year "Survey year"

* ---- all_hh: constant for "All households" rows in collect ----------------
* Setting this to 1 for every observation gives collect/svy a grouping variable
* that produces a single aggregate row with no subsetting
gen byte all_hh = 1
label define all_hh_lbl 1 "All households"  // value label used in collect output
label values all_hh all_hh_lbl

* ---- Demographic variable globals -----------------------------------------
* nControls drives the forvalues loops that iterate over demographic cuts;
* Controls1–Controls7 list the variables in the order rows should appear in output
global nControls 7
global Controls1 praceeth
global Controls2 pagegrp
global Controls3 peducgrp
global Controls4 hhincome2
global Controls5 pdisabl_age25to64
global Controls6 hhtypev2
global Controls7 gtmetsta

* ---- build_demo_rows -------------------------------------------------------
* Defines every row in a standard demographic table: its display label (lab),
* the grouping variable (dv), and the value of that variable for this row (dval).
* Rows with dv=="header" are section dividers (no regression run for them).
* Rows with dv=="all" use the full sample (no group filter in the regression).
* All macros are set as globals so they're visible at the do-file top level.
* Drop any prior definition so the file can be re-sourced without error
capture program drop build_demo_rows
program define build_demo_rows
    local k 0   // row counter; incremented before each definition
    * #delimit ; makes each semicolon a command terminator, so `local ++k;`
    * executes and updates k BEFORE the next command's macros are expanded.
    * Without this, Stata expands `k` across the whole physical line before
    * running any command, so lab`k' would always resolve to lab0.
    #delimit ;
    * Each line: bump counter (one command), then set three globals (three commands).
    * globals are used instead of c_local because c_local only works when the caller
    * is another program — it silently does nothing when called from a do-file.
    local ++k; global lab`k' "All households";         global dv`k' "all";               global dval`k' 0 ;
    local ++k; global lab`k' "Race/Ethnicity";          global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  Black";                 global dv`k' "praceeth";          global dval`k' 1 ;
    local ++k; global lab`k' "  Hispanic";              global dv`k' "praceeth";          global dval`k' 2 ;
    local ++k; global lab`k' "  Asian";                 global dv`k' "praceeth";          global dval`k' 3 ;
    local ++k; global lab`k' "  AIAN";                  global dv`k' "praceeth";          global dval`k' 4 ;
    local ++k; global lab`k' "  NHOPI";                 global dv`k' "praceeth";          global dval`k' 5 ;
    local ++k; global lab`k' "  White";                 global dv`k' "praceeth";          global dval`k' 6 ;
    local ++k; global lab`k' "  Other/Multiracial";     global dv`k' "praceeth";          global dval`k' 7 ;
    local ++k; global lab`k' "Age";                     global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  15-24";                 global dv`k' "pagegrp";           global dval`k' 1 ;
    local ++k; global lab`k' "  25-34";                 global dv`k' "pagegrp";           global dval`k' 2 ;
    local ++k; global lab`k' "  35-44";                 global dv`k' "pagegrp";           global dval`k' 3 ;
    local ++k; global lab`k' "  45-54";                 global dv`k' "pagegrp";           global dval`k' 4 ;
    local ++k; global lab`k' "  55-64";                 global dv`k' "pagegrp";           global dval`k' 5 ;
    local ++k; global lab`k' "  65+";                   global dv`k' "pagegrp";           global dval`k' 6 ;
    local ++k; global lab`k' "Education";               global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  Less than HS";          global dv`k' "peducgrp";          global dval`k' 1 ;
    local ++k; global lab`k' "  HS diploma/GED";        global dv`k' "peducgrp";          global dval`k' 2 ;
    local ++k; global lab`k' "  Some college";          global dv`k' "peducgrp";          global dval`k' 3 ;
    local ++k; global lab`k' "  College degree";        global dv`k' "peducgrp";          global dval`k' 4 ;
    local ++k; global lab`k' "Annual Household Income"; global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  Less than $15,000";     global dv`k' "hhincome2";         global dval`k' 1 ;
    local ++k; global lab`k' "  $15,000-$30,000";       global dv`k' "hhincome2";         global dval`k' 2 ;
    local ++k; global lab`k' "  $30,000-$50,000";       global dv`k' "hhincome2";         global dval`k' 3 ;
    local ++k; global lab`k' "  $50,000-$75,000";       global dv`k' "hhincome2";         global dval`k' 4 ;
    local ++k; global lab`k' "  $75,000 or more";       global dv`k' "hhincome2";         global dval`k' 5 ;
    local ++k; global lab`k' "  Unknown";               global dv`k' "hhincome2";         global dval`k' 99 ;
    local ++k; global lab`k' "Disability (ages 25-64)"; global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  Disabled";              global dv`k' "pdisabl_age25to64"; global dval`k' 1 ;
    local ++k; global lab`k' "  Not disabled";          global dv`k' "pdisabl_age25to64"; global dval`k' 2 ;
    local ++k; global lab`k' "Household Type";          global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  Married couple";        global dv`k' "hhtypev2";          global dval`k' 1 ;
    local ++k; global lab`k' "  Single mother";         global dv`k' "hhtypev2";          global dval`k' 2 ;
    local ++k; global lab`k' "  Other female-HH";       global dv`k' "hhtypev2";          global dval`k' 3 ;
    local ++k; global lab`k' "  Single father";         global dv`k' "hhtypev2";          global dval`k' 4 ;
    local ++k; global lab`k' "  Other male-HH";         global dv`k' "hhtypev2";          global dval`k' 5 ;
    local ++k; global lab`k' "  Female nonfamily";      global dv`k' "hhtypev2";          global dval`k' 6 ;
    local ++k; global lab`k' "  Male nonfamily";        global dv`k' "hhtypev2";          global dval`k' 7 ;
    local ++k; global lab`k' "  Other";                 global dv`k' "hhtypev2";          global dval`k' 8 ;
    local ++k; global lab`k' "Metro Status";            global dv`k' "header";            global dval`k' 0 ;
    local ++k; global lab`k' "  Metropolitan";          global dv`k' "gtmetsta";          global dval`k' 1 ;
    local ++k; global lab`k' "  Nonmetropolitan";       global dv`k' "gtmetsta";          global dval`k' 2 ;
    local ++k; global lab`k' "  Not identified";        global dv`k' "gtmetsta";          global dval`k' 3 ;
    #delimit cr ;   // switch back to newline mode while still in ; mode (the ; terminates this command)
    // Pass the total row count back to the calling scope; now safely in cr mode
    global nrows = `k'
end

di "Setup complete."


*=============================================================================
* TEMPLATE A — Binary outcome × years as columns × demographics as rows
*
* Report examples: Table 1.1, 2.2, 3.1, 4.1
*
* Method: one svy: reg per demographic subgroup, regressing outcome on
* year dummies with diff_from as the base category. This gives:
*   _b[_cons]            = estimate for the base (diff_from) year
*   _b[yr.year]          = estimate for yr minus the base year
*   _b[diff_to.year]     = the diff directly, with its own SE
*   t = _b[.] / _se[.]   = significance test for the diff
*
* All year estimates, the diff, and the significance star come from a
* single regression per subgroup — no separate z-test step needed.
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local outcome    unbanked
local years      2019 2021 2023
local diff_from  2021           // base year: diff measures change FROM this year
local diff_to    2023           // to this year
local subpop     ""             // "" = all HH; e.g. "hunbnk==2" = banked only
local xlfile     "${output}/table_A_example.xlsx"
local sheet      "Table 1.1"
* Use "replace" for the first (or only) sheet; use "modify" when adding a second
* sheet to a file that already exists from an earlier putexcel call
local filemode   replace
* ---- END CONFIGURE ---------------------------------------------------------

* Count how many years are listed — used to size result matrices and locate columns
local nyears = wordcount("`years'")
* Master list of Excel column letters for converting a numeric index to a letter
local colnames "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"

* Populate global lab*, dv*, dval*, and nrows — globals are used because
* c_local only works when called from within a program, not from a do-file
build_demo_rows   // sets $lab1...$labN, $dv1...$dvN, $dval1...$dvalN, $nrows

* Restrict regression to only the years being shown (exclude e.g. year2017)
* subinstr converts space-separated "2019 2021 2023" → "2019,2021,2023" for inlist()
local year_comma = subinstr("`years'", " ", ",", .)

* Base universe: correct years + optional subpop restriction
* Build the if-condition string once so both regression branches share the same filter
if "`subpop'" == "" local universe "inlist(year, `year_comma')"
else                 local universe "inlist(year, `year_comma') & (`subpop')"

* Matrices: M = year estimates (%), DIFF = diff column (%)
* Initialize both with missing so any skipped row stays blank in the output
matrix M    = J($nrows, `nyears', .)
matrix DIFF = J($nrows, 1, .)

* ---- Estimation loop -------------------------------------------------------
* One regression per demographic subgroup.
* ib`diff_from'.year sets diff_from as the omitted (base) category so that
* _b[`diff_to'.year] is directly the diff we want to display and test.

forvalues i = 1/$nrows {
    * Header rows have no associated group variable — skip them entirely
    if "${dv`i'}" == "header" continue

    if "${dv`i'}" == "all" {
        * "All households" row: no group filter, just restrict to the analysis years
        * `capture quietly` suppresses output and catches failures without stopping the loop
        capture quietly svy, subpop(if `universe'): ///
            reg `outcome' ib`diff_from'.year
    }
    else {
        * Demographic subgroup row: restrict to the specific value of the grouping variable
        capture quietly svy, subpop(if `universe' & ${dv`i'}==${dval`i'}): ///
            reg `outcome' ib`diff_from'.year
    }

    * Skip storing results if the regression failed or the subpop is too small to report
    if _rc != 0 | e(N_sub) < 5 continue

    * Copy the coefficient row vector into a named matrix so we can look up
    * columns by name (colnumb) rather than by fragile numeric position
    matrix BETA = e(b)

    * The intercept (_cons) is the estimated mean for the base year (diff_from),
    * because diff_from is the omitted category in the ib. factor-variable notation
    local base = BETA[1, colnumb(BETA, "_cons")]

    * Fill the year-estimate columns of matrix M
    local c = 0
    foreach yr of local years {
        local ++c
        if `yr' == `diff_from' {
            * Base year: just the intercept, scaled from proportion to percentage
            matrix M[`i', `c'] = `base' * 100
        }
        else {
            * Other years: intercept + year coefficient gives the mean for that year
            * colnumb returns missing (.) if the coefficient was dropped (empty cell)
            local cn = colnumb(BETA, "`yr'.year")
            if `cn' < . {
                matrix M[`i', `c'] = (`base' + BETA[1, `cn']) * 100
            }
        }
    }

    * Diff column: _b[diff_to.year] is already the change relative to diff_from,
    * so no subtraction needed — the regression parameterization does the work
    local cn_diff = colnumb(BETA, "`diff_to'.year")
    if `cn_diff' < . {
        matrix DIFF[`i', 1] = BETA[1, `cn_diff'] * 100
        * t-statistic uses the regression SE, which already accounts for the
        * survey design via svy:; no manual sqrt(SE1²+SE2²) calculation needed
        local t = abs(BETA[1, `cn_diff']) / _se[`diff_to'.year]
        if      `t' > 2.576 local sig`i' "**"   // p < 0.01 (two-tailed)
        else if `t' > 1.96  local sig`i' "*"    // p < 0.05 (two-tailed)
        else                 local sig`i' ""     // not significant
    }
}

* ---- Export to Excel -------------------------------------------------------
* Open the workbook; `filemode` is "replace" for first sheet, "modify" for additional sheets
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

* Column headers: A = row labels, then one column per year, then diff and sig
putexcel A1 = "Group"
local c = 1
foreach yr of local years {
    local ++c   // advance column index before writing (B, C, D, ...)
    putexcel `=word("`colnames'", `c')'1 = "`yr'"
}
* Diff and significance columns sit immediately after the last year column
local c_diff = `nyears' + 2
local c_sig  = `nyears' + 3
putexcel `=word("`colnames'", `c_diff')'1 = "Diff (`diff_to'–`diff_from')"
putexcel `=word("`colnames'", `c_sig')'1  = "Sig"

* Row labels: data rows start at row 2 (row 1 holds headers)
forvalues i = 1/$nrows {
    putexcel A`=`i'+1' = "${lab`i'}"
}

* Dump both result matrices in a single call each; Stata fills the block automatically
putexcel B2 = matrix(M),    nformat("0.0")
putexcel `=word("`colnames'", `c_diff')'2 = matrix(DIFF), nformat("0.0")

* Write significance stars only where non-empty to avoid overwriting with blank strings
forvalues i = 1/$nrows {
    if "`sig`i''" != "" {
        putexcel `=word("`colnames'", `c_sig')'`=`i'+1' = "`sig`i''"
    }
}

di "Template A complete → `xlfile' [sheet: `sheet']"

/*
  TWO-OUTCOME TABLES (e.g., Table 2.2: bank teller + mobile app):
  Run this template twice. First run: local filemode replace, first outcome.
  Second run: local filemode modify, second outcome, different sheet name.
*/


*=============================================================================
* TEMPLATE B — Multiple binary outcomes as columns × demographics as rows
*
* Note: svy:table is not a supported svy estimation command in Stata 19 with
* vce(linearized). Restructured to use svy:mean with over(), which is
* supported and computes all subgroup means in a single svy call per
* demographic variable (more efficient than one call per category).
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local outcomes   "huse12mo_b huse12cc_b huse12mt_b"
local out_labels `""Money order" "Check cashing" "Money transfer""'
local year       2023
local subpop     ""
local xlfile     "${output}/table_B_example.xlsx"
local sheet      "Table B"
* ---- END CONFIGURE ---------------------------------------------------------

* year is a 4-digit integer in the 03c dataset (renamed from hryear4), not a 0/1
* indicator, so filter with == rather than looking up a binary year variable
if "`subpop'" == "" local sp "year == `year'"
else                 local sp "year == `year' & (`subpop')"

* Clear any results from a previous collect run before accumulating new chunks
collect clear

* "All households" row: mean of all outcomes across the full subpop;
* over(all_hh) works because all_hh == 1 for every observation
collect, tag(demo[all_hh]): ///
    svy, subpop(if `sp'): mean `outcomes', over(all_hh)

* One collect call per demographic variable; over() returns all categories in
* a single svy pass — no loop over individual category values needed
forvalues i = 1/$nControls {
    local dv ${Controls`i'}
    collect, tag(demo[`dv']): ///
        svy, subpop(if `sp'): mean `outcomes', over(`dv')
}

* Rename outcome variable labels in the var dimension to display headers;
* svy:mean places variable names in the var dimension (not result as table did)
local k = 0
foreach out of local outcomes {
    local ++k
    local lbl : word `k' of `out_labels'
    collect label levels var `out' "`lbl'", modify
}

* Build the row-dimension string: all_hh first, then each demographic variable
local row_dim "demo[all_hh]#all_hh"
forvalues i = 1/$nControls {
    local row_dim "`row_dim' demo[${Controls`i'}]#${Controls`i'}"
}

* Arrange collected results: rows = demographic groups, columns = outcome variables
* var dimension holds the outcome names; result[_r_b] is the point estimate
collect layout (`row_dim') (var)
* Scale proportions to percentages and format to one decimal place
collect style cell result[_r_b], nformat(%6.1f) transform(* 100)

* Write the formatted table to Excel
collect export "`xlfile'", sheet("`sheet'") replace
di "Template B complete → `xlfile' [sheet: `sheet']"


*=============================================================================
* TEMPLATE C — Categorical distribution (one column of percentages)
* (unchanged from v2)
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local catvar   hbankstatv6
local year     2023
local subpop   ""
local xlfile   "${output}/table_C_example.xlsx"
local sheet    "Table C"
* ---- END CONFIGURE ---------------------------------------------------------

* year is a 4-digit integer in this dataset — filter with == not a binary indicator
if "`subpop'" == "" local sp "year == `year'"
else                 local sp "year == `year' & (`subpop')"

* Clear any prior collect results before running the proportion estimate
collect clear
* `svy: proportion` estimates the weighted share of each category of catvar
svy, subpop(if `sp'): proportion `catvar'

* Layout: rows = category values, column = the point estimate (_r_b)
collect layout (`catvar') (result[_r_b])
* Multiply proportions by 100 and format to one decimal place for display
collect style cell result[_r_b], nformat(%6.1f) transform(* 100)

collect export "`xlfile'", sheet("`sheet'") replace
di "Template C complete → `xlfile' [sheet: `sheet']"


*=============================================================================
* TEMPLATE D — Rows = survey years, columns = ordered response categories
* (unchanged from v2)
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local catvar   hbankint
local years    "2019 2021 2023"
local subpop   "hunbnk==1"   // e.g. restrict to unbanked households only
local xlfile   "${output}/table_D_example.xlsx"
local sheet    "Table 1.4"
* ---- END CONFIGURE ---------------------------------------------------------

* Clear any prior collect results before accumulating year-by-year chunks
collect clear

* Run a separate proportion estimate for each year and tag it with the year value
* so the layout step can place each as its own row
foreach yr of local years {
    * year is a 4-digit integer — filter with == not a binary indicator
    if "`subpop'" == "" local sp "year == `yr'"
    else                 local sp "year == `yr' & (`subpop')"

    * Tag this chunk with year[yr] so the layout step can use it as a row dimension
    collect, tag(year[`yr']): ///
        svy, subpop(if `sp'): proportion `catvar'
}

* Arrange results: rows = survey years, columns = response categories of catvar
collect layout (year) (`catvar')
* Convert proportions to percentages and format to one decimal place
collect style cell result[_r_b], nformat(%6.1f) transform(* 100)

collect export "`xlfile'", sheet("`sheet'") replace
di "Template D complete → `xlfile' [sheet: `sheet']"

/*
  TABLE 2.3 NOTE: bank-access methods are "check all that apply" — row
  percentages sum above 100. Use Template B with year fixed and outcomes =
  hbnkaccm1v2_b through hbnkaccm6v2_b, rather than Template D.
*/
