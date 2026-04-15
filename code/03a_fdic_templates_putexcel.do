/*============================================================================
  FDIC Unbanked Survey — Simple Table Templates  (beginner-friendly)

  Four templates, one per common table layout in the annual report.
  Designed for colleagues new to Stata automation; adapt for 2025 data
  by editing only the CONFIGURE block at the top of each template.

  TABLE FORMAT TYPES
  ------------------
  A  Binary outcome × years as columns × demographics as rows
       Report examples: Table 1.1, 2.2, 3.1, 4.1

  B  Multiple binary outcomes as columns × demographics as rows, one year
       Report examples: Table 3.3, 5.1

  C  Categorical distribution: rows = categories, one column of percentages
       Report examples: Table 1.2, 3.2

  D  Rows = survey years, columns = ordered response categories
       Report examples: Table 1.4, 1.3, 2.1, 2.3

  HOW THESE TEMPLATES WORK
  -------------------------
  Each template follows the same three-step pattern:

    Step 1. CONFIGURE: Set the outcome(s), year(s), universe, and output
            file name. These are the only lines you normally need to change.

    Step 2. ESTIMATE: A loop runs svy: mean (or svy: proportion) for each
            demographic subgroup and stores the result — as a percentage —
            in a Stata matrix M.

    Step 3. EXPORT: One putexcel call writes the entire matrix M plus row
            labels to Excel.

  NOTE ON dtable
  --------------
  Stata's -dtable- is designed for sample-description tables (means/SDs
  by group in one year). It does not easily produce the layouts here, where
  years or multiple outcomes appear as columns and demographics run as rows.
  These templates use svy: mean loops and putexcel matrix export instead,
  which is explicit and easy to follow step by step.

  PREREQUISITES
  -------------
    * Stata 17 or later
    * Run the SETUP section once per Stata session before any template
    * Output goes to output/ as .xlsx files (set via $output global)
============================================================================*/


*=============================================================================
* SETUP — Run this block once per Stata session
*=============================================================================

clear all
set more off

global data   "data"     // input data directory (relative to repo root)
global output "output"   // output directory for .xlsx results

set type double   // avoids floating-point rounding surprises

* ---- Load data --------------------------------------------------------------
* Template A uses a single year; set YEAR to the survey year you want.
* Templates B, C, D also work on a single year; adjust CONFIGURE blocks as needed.
local YEAR 2023

use "${data}/hhmultiyear_analys.dta", clear
keep if hryear4 == `YEAR'

* ---- Survey design ----------------------------------------------------------
* The public-use file has probability weights only (no cluster or strata).
svyset [pw=hhsupwgt]

* ---- Recode 1=Yes / 2=No variables to 1/0 ----------------------------------
* svy: mean on a 0/1 variable gives a proportion (= percent / 100).
* Pattern: newvar_b = 1 if rawvar==1, = 0 if rawvar==2, missing otherwise.

foreach v of varlist                                                         ///
    huse12mo huse12cc huse12mt huse12rm huse12rmany                         ///
    husenowops husenowpp huse12afsc                                          ///
    huse12pdl huse12pwn huse12ral huse12atl huse12rto                       ///
    hcred12cc hcred12sc hcred12car hcred12hmln hcred12sl hcred12any         ///
    hcred12bnpl hcred12bnpldq huse12cryp                                    ///
    hbnkaccm1v2 hbnkaccm2v2 hbnkaccm3v2                                     ///
    hbnkaccm4v2 hbnkaccm5v2 hbnkaccm6v2 {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* Banking status as 0/1 convenience variables
gen byte unbanked          = (hunbnk == 1)          if inlist(hunbnk, 1, 2)
gen byte underbanked       = (hbankstatv6 == 2)     if inlist(hbankstatv6, 1, 2, 3)
gen byte cashonly_unbanked = (hunbnkcashonly == 1)  if inlist(hunbnkcashonly, 1, 2)

di "Setup complete."


*=============================================================================
* TEMPLATE A — Binary outcome × years as columns × demographics as rows
*
* Report examples: Table 1.1 (unbanked), 2.2 (bank access), 3.1 (OPS/prepaid),
*                  4.1 (transaction AFS)
*
* Output layout:
*   Group                | 2019  | 2021  | 2023  | Diff (2023–2021)
*   All households       |  5.4  |  4.5  |  4.2  |  -0.3
*   Race/Ethnicity       |       |       |       |
*     Black              | 13.8  | 11.3  |  9.3  |  -2.0
*     Hispanic           | ...
*
* For a table with TWO outcomes side by side (e.g., Table 2.2 with teller
* and mobile), run this template twice: first call uses "replace", second
* call uses "modify" and a different sheet name. See note at end.
*=============================================================================

* ---- CONFIGURE (edit these locals) -----------------------------------------
local outcome    unbanked       // 0/1 outcome variable (append _b for recoded vars)
local years      2019 2021 2023 // years shown as columns; must exist in hryear4
local diff_from  2021           // "prior" year for the Diff column
local diff_to    2023           // "current" year for the Diff column
local subpop     ""             // restrict universe: leave "" for all HH,
                                //   or e.g. "hunbnk==2" for banked HH only
local xlfile     "${output}/table_A_example.xlsx"
local sheet      "Table 1.1"
local filemode   replace        // "replace" to start a new workbook; "modify"
                                //   to add a sheet to an existing workbook
* ---- END CONFIGURE ----------------------------------------------------------

* Number of years and total columns (years + diff)
local nyears = wordcount("`years'")
local ncols  = `nyears' + 1

* ---- Define demographic rows as parallel lists ------------------------------
* Each row has three parallel locals: lab`r', dv`r', dval`r'.
*   lab  = text label for Excel
*   dv   = Stata variable name to condition on ("all" for national, "header"
*           for a section heading with no number)
*   dval = category value to condition on (ignored for "all" and "header")

local r = 0   // row counter

* National total
local ++r; local lab`r' "All households";       local dv`r' "all";             local dval`r' 0
* Race / Ethnicity
local ++r; local lab`r' "Race/Ethnicity";        local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Black";               local dv`r' "praceeth";        local dval`r' 1
local ++r; local lab`r' "  Hispanic";            local dv`r' "praceeth";        local dval`r' 2
local ++r; local lab`r' "  Asian";               local dv`r' "praceeth";        local dval`r' 3
local ++r; local lab`r' "  AIAN";                local dv`r' "praceeth";        local dval`r' 4
local ++r; local lab`r' "  NHOPI";               local dv`r' "praceeth";        local dval`r' 5
local ++r; local lab`r' "  White";               local dv`r' "praceeth";        local dval`r' 6
local ++r; local lab`r' "  Other/Multiracial";   local dv`r' "praceeth";        local dval`r' 7
* Age
local ++r; local lab`r' "Age";                   local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  15-24";               local dv`r' "pagegrp";         local dval`r' 1
local ++r; local lab`r' "  25-34";               local dv`r' "pagegrp";         local dval`r' 2
local ++r; local lab`r' "  35-44";               local dv`r' "pagegrp";         local dval`r' 3
local ++r; local lab`r' "  45-54";               local dv`r' "pagegrp";         local dval`r' 4
local ++r; local lab`r' "  55-64";               local dv`r' "pagegrp";         local dval`r' 5
local ++r; local lab`r' "  65+";                 local dv`r' "pagegrp";         local dval`r' 6
* Education
local ++r; local lab`r' "Education";             local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Less than HS";        local dv`r' "peducgrp";        local dval`r' 1
local ++r; local lab`r' "  HS diploma/GED";      local dv`r' "peducgrp";        local dval`r' 2
local ++r; local lab`r' "  Some college";        local dv`r' "peducgrp";        local dval`r' 3
local ++r; local lab`r' "  College degree";      local dv`r' "peducgrp";        local dval`r' 4
* Income
local ++r; local lab`r' "Annual Household Income"; local dv`r' "header";        local dval`r' 0
local ++r; local lab`r' "  Less than $15,000";   local dv`r' "hhincome2";       local dval`r' 1
local ++r; local lab`r' "  $15,000-$30,000";     local dv`r' "hhincome2";       local dval`r' 2
local ++r; local lab`r' "  $30,000-$50,000";     local dv`r' "hhincome2";       local dval`r' 3
local ++r; local lab`r' "  $50,000-$75,000";     local dv`r' "hhincome2";       local dval`r' 4
local ++r; local lab`r' "  $75,000 or more";     local dv`r' "hhincome2";       local dval`r' 5
local ++r; local lab`r' "  Unknown";             local dv`r' "hhincome2";       local dval`r' 99
* Disability
local ++r; local lab`r' "Disability Status (ages 25-64)"; local dv`r' "header"; local dval`r' 0
local ++r; local lab`r' "  Disabled";            local dv`r' "pdisabl_age25to64"; local dval`r' 1
local ++r; local lab`r' "  Not disabled";        local dv`r' "pdisabl_age25to64"; local dval`r' 2
* Household type
local ++r; local lab`r' "Household Type";        local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Married couple";      local dv`r' "hhtypev2";        local dval`r' 1
local ++r; local lab`r' "  Single mother";       local dv`r' "hhtypev2";        local dval`r' 2
local ++r; local lab`r' "  Other female-HH";     local dv`r' "hhtypev2";        local dval`r' 3
local ++r; local lab`r' "  Single father";       local dv`r' "hhtypev2";        local dval`r' 4
local ++r; local lab`r' "  Other male-HH";       local dv`r' "hhtypev2";        local dval`r' 5
local ++r; local lab`r' "  Female nonfamily";    local dv`r' "hhtypev2";        local dval`r' 6
local ++r; local lab`r' "  Male nonfamily";      local dv`r' "hhtypev2";        local dval`r' 7
local ++r; local lab`r' "  Other";               local dv`r' "hhtypev2";        local dval`r' 8
* Metro status
local ++r; local lab`r' "Metro Status";          local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Metropolitan";        local dv`r' "gtmetsta";        local dval`r' 1
local ++r; local lab`r' "  Nonmetropolitan";     local dv`r' "gtmetsta";        local dval`r' 2
local ++r; local lab`r' "  Not identified";      local dv`r' "gtmetsta";        local dval`r' 3

local nrows = `r'

* ---- Find which matrix columns correspond to diff_from and diff_to ---------
local col_from = 0
local col_to   = 0
local c = 0
foreach yr of local years {
    local ++c
    if "`yr'" == "`diff_from'" local col_from = `c'
    if "`yr'" == "`diff_to'"   local col_to   = `c'
}
if `col_from' == 0 | `col_to' == 0 {
    di as error "diff_from (`diff_from') or diff_to (`diff_to') not found in years list."
    exit 198
}

* ---- Initialize results matrix ---------------------------------------------
* Rows = demographic groups; Cols = one per year + one for diff.
* Missing (.) means either a header row or too few obs to estimate.
matrix M = J(`nrows', `ncols', .)

* ---- Estimation loop --------------------------------------------------------
* For each demographic row and each year:
*   - Build the subpopulation condition (year + optional universe restriction)
*   - Run svy: mean within that subpopulation
*   - Store estimate × 100 (percentage) in the matrix

local colnames "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"

forvalues i = 1/`nrows' {

    if "`dv`i''" == "header" continue   // header rows: skip estimation

    local c = 0
    foreach yr of local years {
        local ++c

        * Build subpop condition: restrict to year; add universe restriction if specified
        if "`subpop'" == "" {
            local sp_cond "hryear4==`yr'"
        }
        else {
            local sp_cond "hryear4==`yr' & (`subpop')"
        }

        * Add demographic restriction (not needed for the national "all" row)
        if "`dv`i''" == "all" {
            capture quietly svy, subpop(if `sp_cond'): mean `outcome'
        }
        else {
            capture quietly svy, subpop(if `sp_cond' & `dv`i''==`dval`i''): mean `outcome'
        }

        * Store as percentage if estimation succeeded and N_sub >= 5
        if _rc == 0 & e(N_sub) >= 5 {
            matrix M[`i', `c'] = _b[`outcome'] * 100
        }
    }   // end year loop

    * Difference column (last column of M)
    if M[`i', `col_to'] < . & M[`i', `col_from'] < . {
        matrix M[`i', `ncols'] = M[`i', `col_to'] - M[`i', `col_from']
    }

}   // end demographic row loop

* ---- Export to Excel -------------------------------------------------------
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

* Column headers in row 1 (columns B onward)
local c = 1
foreach yr of local years {
    local ++c
    local col = word("`colnames'", `c')
    putexcel `col'1 = "`yr'"
}
local c_diff = `nyears' + 2
local col_diff = word("`colnames'", `c_diff')
putexcel `col_diff'1 = "Diff (`diff_to'–`diff_from')"

* Row labels in column A (rows 2 onward)
forvalues i = 1/`nrows' {
    putexcel A`=`i'+1' = "`lab`i''"
}

* All numeric estimates in one call (matrix starts at B2)
putexcel B2 = matrix(M), nformat("0.0")

di "Template A complete. Output: `xlfile' [sheet: `sheet']"

/*
  TO PRODUCE TABLE 2.2 (two outcomes: teller and mobile app):
  Run Template A twice, changing outcome and sheet each time:

    First run:  local outcome = "hbnkaccm1v2_b"  (teller)
                local sheet   = "Teller"
                local filemode = replace
                (also set subpop to banked HH that accessed account)

    Second run: local outcome = "hbnkaccm5v2_b"  (mobile app)
                local sheet   = "Mobile"
                local filemode = modify            // <-- keep existing workbook
*/


*=============================================================================
* TEMPLATE B — Multiple binary outcomes as columns × demographics as rows
*
* Report examples: Table 3.3 (cash-only unbanked, OPS, prepaid),
*                  Table 5.1 (mainstream credit products)
*
* Output layout:
*   Group             | Money order | Check cash | Money transfer | ...
*   All households    |    14.2     |    3.1     |     7.0        | ...
*   Race/Ethnicity    |             |            |                |
*     Black           |    25.4     |    ...
*
* NOTE: This is the same as Template A but with outcomes-as-columns instead
* of years-as-columns. Change the CONFIGURE block and the row definitions
* (copied from Template A) to use it for a different table.
*=============================================================================

* ---- CONFIGURE (edit these) ------------------------------------------------
local outcomes  "huse12mo_b huse12cc_b huse12mt_b"   // space-separated list of 0/1 variables
local out_labels "Money order" "Check cashing" "Money transfer"  // one label per outcome
local year       2023          // single year
local subpop     ""            // universe restriction: "" = all HH
local xlfile     "${output}/table_B_example.xlsx"
local sheet      "Table B"
local filemode   replace
* ---- END CONFIGURE ----------------------------------------------------------

local nout   = wordcount("`outcomes'")
local ncols_B = `nout'    // one column per outcome (no diff in Type B)

* Demographic rows: same parallel list as Template A.
* (Copy/paste the local ++r block from Template A here and remove those lines
*  that are not needed. This block is repeated here so Template B is
*  self-contained.)

local r = 0
local ++r; local lab`r' "All households";       local dv`r' "all";             local dval`r' 0
local ++r; local lab`r' "Race/Ethnicity";        local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Black";               local dv`r' "praceeth";        local dval`r' 1
local ++r; local lab`r' "  Hispanic";            local dv`r' "praceeth";        local dval`r' 2
local ++r; local lab`r' "  Asian";               local dv`r' "praceeth";        local dval`r' 3
local ++r; local lab`r' "  AIAN";                local dv`r' "praceeth";        local dval`r' 4
local ++r; local lab`r' "  NHOPI";               local dv`r' "praceeth";        local dval`r' 5
local ++r; local lab`r' "  White";               local dv`r' "praceeth";        local dval`r' 6
local ++r; local lab`r' "  Other/Multiracial";   local dv`r' "praceeth";        local dval`r' 7
local ++r; local lab`r' "Age";                   local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  15-24";               local dv`r' "pagegrp";         local dval`r' 1
local ++r; local lab`r' "  25-34";               local dv`r' "pagegrp";         local dval`r' 2
local ++r; local lab`r' "  35-44";               local dv`r' "pagegrp";         local dval`r' 3
local ++r; local lab`r' "  45-54";               local dv`r' "pagegrp";         local dval`r' 4
local ++r; local lab`r' "  55-64";               local dv`r' "pagegrp";         local dval`r' 5
local ++r; local lab`r' "  65+";                 local dv`r' "pagegrp";         local dval`r' 6
local ++r; local lab`r' "Education";             local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Less than HS";        local dv`r' "peducgrp";        local dval`r' 1
local ++r; local lab`r' "  HS diploma/GED";      local dv`r' "peducgrp";        local dval`r' 2
local ++r; local lab`r' "  Some college";        local dv`r' "peducgrp";        local dval`r' 3
local ++r; local lab`r' "  College degree";      local dv`r' "peducgrp";        local dval`r' 4
local ++r; local lab`r' "Annual Household Income"; local dv`r' "header";        local dval`r' 0
local ++r; local lab`r' "  Less than $15,000";   local dv`r' "hhincome2";       local dval`r' 1
local ++r; local lab`r' "  $15,000-$30,000";     local dv`r' "hhincome2";       local dval`r' 2
local ++r; local lab`r' "  $30,000-$50,000";     local dv`r' "hhincome2";       local dval`r' 3
local ++r; local lab`r' "  $50,000-$75,000";     local dv`r' "hhincome2";       local dval`r' 4
local ++r; local lab`r' "  $75,000 or more";     local dv`r' "hhincome2";       local dval`r' 5
local ++r; local lab`r' "  Unknown";             local dv`r' "hhincome2";       local dval`r' 99
local ++r; local lab`r' "Disability Status (ages 25-64)"; local dv`r' "header"; local dval`r' 0
local ++r; local lab`r' "  Disabled";            local dv`r' "pdisabl_age25to64"; local dval`r' 1
local ++r; local lab`r' "  Not disabled";        local dv`r' "pdisabl_age25to64"; local dval`r' 2
local ++r; local lab`r' "Household Type";        local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Married couple";      local dv`r' "hhtypev2";        local dval`r' 1
local ++r; local lab`r' "  Single mother";       local dv`r' "hhtypev2";        local dval`r' 2
local ++r; local lab`r' "  Other female-HH";     local dv`r' "hhtypev2";        local dval`r' 3
local ++r; local lab`r' "  Single father";       local dv`r' "hhtypev2";        local dval`r' 4
local ++r; local lab`r' "  Other male-HH";       local dv`r' "hhtypev2";        local dval`r' 5
local ++r; local lab`r' "  Female nonfamily";    local dv`r' "hhtypev2";        local dval`r' 6
local ++r; local lab`r' "  Male nonfamily";      local dv`r' "hhtypev2";        local dval`r' 7
local ++r; local lab`r' "  Other";               local dv`r' "hhtypev2";        local dval`r' 8
local ++r; local lab`r' "Metro Status";          local dv`r' "header";          local dval`r' 0
local ++r; local lab`r' "  Metropolitan";        local dv`r' "gtmetsta";        local dval`r' 1
local ++r; local lab`r' "  Nonmetropolitan";     local dv`r' "gtmetsta";        local dval`r' 2
local ++r; local lab`r' "  Not identified";      local dv`r' "gtmetsta";        local dval`r' 3

local nrows = `r'

* ---- Initialize results matrix (rows = demographics, cols = outcomes) ------
matrix MB = J(`nrows', `ncols_B', .)

* ---- Estimation loop --------------------------------------------------------
forvalues i = 1/`nrows' {

    if "`dv`i''" == "header" continue

    local c = 0
    foreach out of local outcomes {
        local ++c

        if "`subpop'" == "" {
            local sp_cond "year`year'==1"
        }
        else {
            local sp_cond "year`year'==1 & (`subpop')"
        }

        if "`dv`i''" == "all" {
            capture quietly svy, subpop(if `sp_cond'): mean `out'
        }
        else {
            capture quietly svy, subpop(if `sp_cond' & `dv`i''==`dval`i''): mean `out'
        }

        if _rc == 0 & e(N_sub) >= 5 {
            matrix MB[`i', `c'] = _b[`out'] * 100
        }
    }
}

* ---- Export to Excel --------------------------------------------------------
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

* Column headers (B1, C1, ...)
local c = 1
foreach lbl of local out_labels {
    local ++c
    local col = word("`colnames'", `c')
    putexcel `col'1 = "`lbl'"
}

* Row labels
forvalues i = 1/`nrows' {
    putexcel A`=`i'+1' = "`lab`i''"
}

* All numeric values
putexcel B2 = matrix(MB), nformat("0.0")

di "Template B complete. Output: `xlfile' [sheet: `sheet']"


*=============================================================================
* TEMPLATE C — Categorical distribution: rows = categories, one column
*
* Report examples:
*   Table 1.2  Account-ownership transition status (4 categories, all HH)
*   Table 3.2  OPS/prepaid combination status (4 categories, unbanked HH)
*
* Output layout:
*   Category                     | Percent
*   Longer-term unbanked         |   3.7
*   Recently unbanked            |   0.5
*   Recently banked              |   5.9
*   Longer-term banked           |  89.9
*
* How it works: svy: proportion returns the share of each category.
* Percentages sum to 100 across categories.
*=============================================================================

* ---- CONFIGURE (edit these) ------------------------------------------------
local catvar  hbankstatv6       // categorical variable (integer-coded)
local cat_vals "1 2 3"          // category values (space-separated, in order)
local cat_labs "Unbanked" "Underbanked" "Fully banked"  // matching labels
local year     2023
local subpop   ""               // "" = all HH; or condition like "hunbnk==1"
local xlfile   "${output}/table_C_example.xlsx"
local sheet    "Table C"
local filemode replace
* ---- END CONFIGURE ----------------------------------------------------------

local ncat = wordcount("`cat_vals'")

* ---- Run svy: proportion ------------------------------------------------
if "`subpop'" == "" {
    svy, subpop(if year`year'==1): proportion `catvar'
}
else {
    svy, subpop(if year`year'==1 & (`subpop')): proportion `catvar'
}

* svy: proportion stores results in e(b): one coefficient per category.
* Coefficients are proportions (0-1). We multiply by 100 for percent.

* Store as a column vector (ncat rows × 1 column)
matrix MC = J(`ncat', 1, .)
forvalues k = 1/`ncat' {
    local val : word `k' of `cat_vals'
    * e(b) stores values as _prop_1, _prop_2, etc., corresponding to
    * the sorted values of catvar. Find which column matches val.
    * Easiest: just read from the e(b) matrix in order.
    matrix MC[`k', 1] = el(e(b), 1, `k') * 100
}

* ---- Export to Excel --------------------------------------------------------
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

putexcel A1 = "Category"
putexcel B1 = "`year'"

forvalues k = 1/`ncat' {
    local lbl : word `k' of `cat_labs'
    putexcel A`=`k'+1' = "`lbl'"
}

putexcel B2 = matrix(MC), nformat("0.0")

di "Template C complete. Output: `xlfile' [sheet: `sheet']"

/*
  CAUTION: svy: proportion labels categories by the sorted integer value
  of catvar. If your category values are not 1, 2, 3, ... in order, verify
  that el(e(b), 1, k) corresponds to the k-th value in cat_vals.
  You can check by running:  svy, subpop(if year20XX==1): proportion catvar
  and reading the e(b) matrix in the Results window.
*/


*=============================================================================
* TEMPLATE D — Rows = survey years, columns = ordered response categories
*
* Report examples:
*   Table 1.4  Interest in having bank account (4 levels, unbanked HH)
*   Table 1.3  Unbanked prior banking status (2 categories, unbanked HH)
*   Table 2.1  Primary account-access method (6 methods, banked HH)
*   Table 2.3  Account-access methods used (6 methods, multiple response)
*
* Output layout:
*   Year | Very interested | Somewhat | Not very | Not at all
*   2019 |      XX.X       |  XX.X    |  XX.X    |   XX.X
*   2021 |      XX.X       |  ...
*   2023 |      ...
*
* How it works: run svy: proportion once per year, store results as a row
* in matrix MD, then export the matrix with years as row labels.
*=============================================================================

* ---- CONFIGURE (edit these) ------------------------------------------------
local catvar   hbankint         // categorical response variable
local cat_vals "1 2 3 4"        // category values, in the column order you want
local cat_labs "Very interested" "Somewhat interested" "Not very interested" "Not at all interested"
local years    "2019 2021 2023" // years to show as rows
local subpop   "hunbnk==1"      // universe: unbanked HH (leave "" for all HH)
local xlfile   "${output}/table_D_example.xlsx"
local sheet    "Table 1.4"
local filemode replace
* ---- END CONFIGURE ----------------------------------------------------------

local ncat   = wordcount("`cat_vals'")
local nyears = wordcount("`years'")

* ---- Initialize results matrix (rows = years, cols = categories) -----------
matrix MD = J(`nyears', `ncat', .)

* ---- Estimation loop --------------------------------------------------------
local row = 0
foreach yr of local years {
    local ++row

    if "`subpop'" == "" {
        capture quietly svy, subpop(if year`yr'==1): proportion `catvar'
    }
    else {
        capture quietly svy, subpop(if year`yr'==1 & (`subpop')): proportion `catvar'
    }

    if _rc != 0 {
        di as error "  Warning: estimation failed for year `yr'. Row left missing."
        continue
    }

    * Store each category percentage in the corresponding column.
    * e(b) columns correspond to category values in sorted order.
    * We store in the order specified by cat_vals.
    forvalues k = 1/`ncat' {
        matrix MD[`row', `k'] = el(e(b), 1, `k') * 100
    }
}

* ---- Export to Excel --------------------------------------------------------
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

* Column headers
putexcel A1 = "Year"
local c = 1
foreach lbl of local cat_labs {
    local ++c
    local col = word("`colnames'", `c')
    putexcel `col'1 = "`lbl'"
}

* Row labels (year values in column A)
local row = 0
foreach yr of local years {
    local ++row
    putexcel A`=`row'+1' = "`yr'"
}

* Numeric values
putexcel B2 = matrix(MD), nformat("0.0")

di "Template D complete. Output: `xlfile' [sheet: `sheet']"

/*
  SAME CAUTION AS TEMPLATE C: svy: proportion stores categories in sorted
  integer order. Verify that el(e(b), 1, k) matches cat_vals[k] by
  inspecting the proportion output for one year before trusting the matrix.

  TABLE 2.3 NOTE: Bank-access methods are "check all that apply" (multiple
  response). Use svy: mean on each binary indicator (hbnkaccm1v2_b through
  hbnkaccm6v2_b) rather than svy: proportion. Row percentages will sum to
  more than 100. Use Template B instead, restricted to banked HH, with one
  year per run, if you want the Table 2.3 layout with methods as columns.
*/
