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
    * Output goes to ../data/ as .xlsx files
============================================================================*/


*=============================================================================
* SETUP — Run once per session
*=============================================================================

clear all
set more off
set type double

* ---- Load data and set survey design --------------------------------------
* Keep all years needed for regression-based Type A diff estimates.
* Adjust the inlist() if you want a different year range.
use "data/hhmultiyear_analys.dta", clear
keep if inlist(hryear4, 2017, 2019, 2021, 2023)
svyset [pw=hhsupwgt]

* ---- Recode 1=Yes / 2=No variables to 0/1 --------------------------------
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

gen byte unbanked          = (hunbnk == 1)          if inlist(hunbnk, 1, 2)
gen byte underbanked       = (hbankstatv6 == 2)     if inlist(hbankstatv6, 1, 2, 3)
gen byte cashonly_unbanked = (hunbnkcashonly == 1)  if inlist(hunbnkcashonly, 1, 2)

* ---- Year variable ---------------------------------------------------------
* hryear4 is already in the multiyear data; rename for consistency with templates.
rename hryear4 year
label variable year "Survey year"

* ---- all_hh: constant for "All households" rows in collect ----------------
gen byte all_hh = 1
label define all_hh_lbl 1 "All households"
label values all_hh all_hh_lbl

* ---- Demographic variable globals -----------------------------------------
global nControls 7
global Controls1 praceeth
global Controls2 pagegrp
global Controls3 peducgrp
global Controls4 hhincome2
global Controls5 pdisabl_age25to64
global Controls6 hhtypev2
global Controls7 gtmetsta

* ---- build_demo_rows -------------------------------------------------------
capture program drop build_demo_rows
program define build_demo_rows
    local k 0
    local ++k; c_local lab`k' "All households";         c_local dv`k' "all";               c_local dval`k' 0
    local ++k; c_local lab`k' "Race/Ethnicity";          c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  Black";                 c_local dv`k' "praceeth";          c_local dval`k' 1
    local ++k; c_local lab`k' "  Hispanic";              c_local dv`k' "praceeth";          c_local dval`k' 2
    local ++k; c_local lab`k' "  Asian";                 c_local dv`k' "praceeth";          c_local dval`k' 3
    local ++k; c_local lab`k' "  AIAN";                  c_local dv`k' "praceeth";          c_local dval`k' 4
    local ++k; c_local lab`k' "  NHOPI";                 c_local dv`k' "praceeth";          c_local dval`k' 5
    local ++k; c_local lab`k' "  White";                 c_local dv`k' "praceeth";          c_local dval`k' 6
    local ++k; c_local lab`k' "  Other/Multiracial";     c_local dv`k' "praceeth";          c_local dval`k' 7
    local ++k; c_local lab`k' "Age";                     c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  15-24";                 c_local dv`k' "pagegrp";           c_local dval`k' 1
    local ++k; c_local lab`k' "  25-34";                 c_local dv`k' "pagegrp";           c_local dval`k' 2
    local ++k; c_local lab`k' "  35-44";                 c_local dv`k' "pagegrp";           c_local dval`k' 3
    local ++k; c_local lab`k' "  45-54";                 c_local dv`k' "pagegrp";           c_local dval`k' 4
    local ++k; c_local lab`k' "  55-64";                 c_local dv`k' "pagegrp";           c_local dval`k' 5
    local ++k; c_local lab`k' "  65+";                   c_local dv`k' "pagegrp";           c_local dval`k' 6
    local ++k; c_local lab`k' "Education";               c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  Less than HS";          c_local dv`k' "peducgrp";          c_local dval`k' 1
    local ++k; c_local lab`k' "  HS diploma/GED";        c_local dv`k' "peducgrp";          c_local dval`k' 2
    local ++k; c_local lab`k' "  Some college";          c_local dv`k' "peducgrp";          c_local dval`k' 3
    local ++k; c_local lab`k' "  College degree";        c_local dv`k' "peducgrp";          c_local dval`k' 4
    local ++k; c_local lab`k' "Annual Household Income"; c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  Less than $15,000";     c_local dv`k' "hhincome2";         c_local dval`k' 1
    local ++k; c_local lab`k' "  $15,000-$30,000";       c_local dv`k' "hhincome2";         c_local dval`k' 2
    local ++k; c_local lab`k' "  $30,000-$50,000";       c_local dv`k' "hhincome2";         c_local dval`k' 3
    local ++k; c_local lab`k' "  $50,000-$75,000";       c_local dv`k' "hhincome2";         c_local dval`k' 4
    local ++k; c_local lab`k' "  $75,000 or more";       c_local dv`k' "hhincome2";         c_local dval`k' 5
    local ++k; c_local lab`k' "  Unknown";               c_local dv`k' "hhincome2";         c_local dval`k' 99
    local ++k; c_local lab`k' "Disability (ages 25-64)"; c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  Disabled";              c_local dv`k' "pdisabl_age25to64"; c_local dval`k' 1
    local ++k; c_local lab`k' "  Not disabled";          c_local dv`k' "pdisabl_age25to64"; c_local dval`k' 2
    local ++k; c_local lab`k' "Household Type";          c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  Married couple";        c_local dv`k' "hhtypev2";          c_local dval`k' 1
    local ++k; c_local lab`k' "  Single mother";         c_local dv`k' "hhtypev2";          c_local dval`k' 2
    local ++k; c_local lab`k' "  Other female-HH";       c_local dv`k' "hhtypev2";          c_local dval`k' 3
    local ++k; c_local lab`k' "  Single father";         c_local dv`k' "hhtypev2";          c_local dval`k' 4
    local ++k; c_local lab`k' "  Other male-HH";         c_local dv`k' "hhtypev2";          c_local dval`k' 5
    local ++k; c_local lab`k' "  Female nonfamily";      c_local dv`k' "hhtypev2";          c_local dval`k' 6
    local ++k; c_local lab`k' "  Male nonfamily";        c_local dv`k' "hhtypev2";          c_local dval`k' 7
    local ++k; c_local lab`k' "  Other";                 c_local dv`k' "hhtypev2";          c_local dval`k' 8
    local ++k; c_local lab`k' "Metro Status";            c_local dv`k' "header";            c_local dval`k' 0
    local ++k; c_local lab`k' "  Metropolitan";          c_local dv`k' "gtmetsta";          c_local dval`k' 1
    local ++k; c_local lab`k' "  Nonmetropolitan";       c_local dv`k' "gtmetsta";          c_local dval`k' 2
    local ++k; c_local lab`k' "  Not identified";        c_local dv`k' "gtmetsta";          c_local dval`k' 3
    c_local nrows `k'
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
local xlfile     "../data/table_A_example.xlsx"
local sheet      "Table 1.1"
local filemode   replace
* ---- END CONFIGURE ---------------------------------------------------------

local nyears = wordcount("`years'")
local colnames "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"

build_demo_rows   // sets lab1...labN, dv1...dvN, dval1...dvalN, nrows

* Restrict regression to only the years being shown (exclude e.g. year2017)
local year_comma = subinstr("`years'", " ", ",", .)

* Base universe: correct years + optional subpop restriction
if "`subpop'" == "" local universe "inlist(year, `year_comma')"
else                 local universe "inlist(year, `year_comma') & (`subpop')"

* Matrices: M = year estimates (%), DIFF = diff column (%)
matrix M    = J(`nrows', `nyears', .)
matrix DIFF = J(`nrows', 1, .)

* ---- Estimation loop -------------------------------------------------------
* One regression per demographic subgroup.
* ib`diff_from'.year sets diff_from as the omitted (base) category so that
* _b[`diff_to'.year] is directly the diff we want to display and test.

forvalues i = 1/`nrows' {
    if "`dv`i''" == "header" continue

    if "`dv`i''" == "all" {
        capture quietly svy, subpop(if `universe'): ///
            reg `outcome' ib`diff_from'.year
    }
    else {
        capture quietly svy, subpop(if `universe' & `dv`i''==`dval`i''): ///
            reg `outcome' ib`diff_from'.year
    }

    if _rc != 0 | e(N_sub) < 5 continue

    * Store coefficient vector for safe lookup by column name
    matrix BETA = e(b)

    * Base year estimate (= _b[_cons] since diff_from is the reference)
    local base = BETA[1, colnumb(BETA, "_cons")]

    * Fill year estimate columns
    local c = 0
    foreach yr of local years {
        local ++c
        if `yr' == `diff_from' {
            matrix M[`i', `c'] = `base' * 100
        }
        else {
            * colnumb returns missing (.) if the coefficient was dropped
            local cn = colnumb(BETA, "`yr'.year")
            if `cn' < . {
                matrix M[`i', `c'] = (`base' + BETA[1, `cn']) * 100
            }
        }
    }

    * Diff column: coefficient on diff_to year (already relative to diff_from)
    local cn_diff = colnumb(BETA, "`diff_to'.year")
    if `cn_diff' < . {
        matrix DIFF[`i', 1] = BETA[1, `cn_diff'] * 100
        * t-statistic from regression SE (no manual sqrt(SE1²+SE2²) needed)
        local t = abs(BETA[1, `cn_diff']) / _se[`diff_to'.year]
        if      `t' > 2.576 local sig`i' "**"
        else if `t' > 1.96  local sig`i' "*"
        else                 local sig`i' ""
    }
}

* ---- Export to Excel -------------------------------------------------------
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

* Column headers
putexcel A1 = "Group"
local c = 1
foreach yr of local years {
    local ++c
    putexcel `=word("`colnames'", `c')'1 = "`yr'"
}
local c_diff = `nyears' + 2
local c_sig  = `nyears' + 3
putexcel `=word("`colnames'", `c_diff')'1 = "Diff (`diff_to'–`diff_from')"
putexcel `=word("`colnames'", `c_sig')'1  = "Sig"

* Row labels
forvalues i = 1/`nrows' {
    putexcel A`=`i'+1' = "`lab`i''"
}

* Year estimates and diff (two matrix calls)
putexcel B2 = matrix(M),    nformat("0.0")
putexcel `=word("`colnames'", `c_diff')'2 = matrix(DIFF), nformat("0.0")

* Significance strings
forvalues i = 1/`nrows' {
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
* (unchanged from v2)
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local outcomes   "huse12mo_b huse12cc_b huse12mt_b"
local out_labels `""Money order" "Check cashing" "Money transfer""'
local year       2023
local subpop     ""
local xlfile     "../data/table_B_example.xlsx"
local sheet      "Table B"
* ---- END CONFIGURE ---------------------------------------------------------

local stat_string ""
foreach out of local outcomes {
    local stat_string "`stat_string' statistic(mean `out')"
}

if "`subpop'" == "" local sp "year`year'==1"
else                 local sp "year`year'==1 & (`subpop')"

collect clear

collect, tag(demo[all_hh]): ///
    svy, subpop(if `sp'): table all_hh, `stat_string'

forvalues i = 1/$nControls {
    local dv ${Controls`i'}
    collect, tag(demo[`dv']): ///
        svy, subpop(if `sp'): table `dv', `stat_string'
}

local k = 0
foreach out of local outcomes {
    local ++k
    local lbl : word `k' of `out_labels'
    collect label levels result "mean(`out')" "`lbl'", modify
}

local row_dim "demo[all_hh]#all_hh"
forvalues i = 1/$nControls {
    local row_dim "`row_dim' demo[${Controls`i'}]#${Controls`i'}"
}

collect layout (`row_dim') (result)
collect style cell, nformat(%6.1f)

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
local xlfile   "../data/table_C_example.xlsx"
local sheet    "Table C"
* ---- END CONFIGURE ---------------------------------------------------------

if "`subpop'" == "" local sp "year`year'==1"
else                 local sp "year`year'==1 & (`subpop')"

collect clear
svy, subpop(if `sp'): proportion `catvar'

collect layout (`catvar') (result[_r_b])
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
local subpop   "hunbnk==1"
local xlfile   "../data/table_D_example.xlsx"
local sheet    "Table 1.4"
* ---- END CONFIGURE ---------------------------------------------------------

collect clear

foreach yr of local years {
    if "`subpop'" == "" local sp "year`yr'==1"
    else                 local sp "year`yr'==1 & (`subpop')"

    collect, tag(year[`yr']): ///
        svy, subpop(if `sp'): proportion `catvar'
}

collect layout (year) (`catvar')
collect style cell result[_r_b], nformat(%6.1f) transform(* 100)

collect export "`xlfile'", sheet("`sheet'") replace
di "Template D complete → `xlfile' [sheet: `sheet']"

/*
  TABLE 2.3 NOTE: bank-access methods are "check all that apply" — row
  percentages sum above 100. Use Template B with year fixed and outcomes =
  hbnkaccm1v2_b through hbnkaccm6v2_b, rather than Template D.
*/
