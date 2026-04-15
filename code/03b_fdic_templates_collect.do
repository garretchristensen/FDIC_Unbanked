/*============================================================================
  FDIC Unbanked Survey — Table Templates v2

  Changes from v1 templates:
    - Types B, C, D rewritten using table + collect (fewer lines, no manual
      matrix indexing)
    - Type A keeps the loop+matrix approach: diff column needs a z-test on
      stored SEs, which table+collect does not support natively
    - Demographic characteristic variables defined once as $Controls globals
    - build_demo_rows program eliminates the duplicate row-definition block
      that appeared in both Templates A and B in v1

  TEMPLATES
  ---------
  A  Binary outcome × years as columns × demographics as rows
       Keeps: loop + matrix + putexcel
       Reason: diff column requires storing and computing on SEs; not
               supported in the collect framework without extra steps

  B  Multiple binary outcomes as columns × demographics as rows, one year
       Uses: table + collect

  C  Categorical distribution (one column of percentages)
       Uses: svy: proportion + collect

  D  Rows = survey years, columns = ordered response categories
       Uses: svy: proportion with per-year collect tags

  PREREQUISITES
  -------------
    * Stata 17 or later
    * Run the SETUP section once per Stata session
    * Output goes to output/ as .xlsx files (set via $output global)
============================================================================*/


*=============================================================================
* SETUP — Run once per session
*=============================================================================

clear all
set more off

global data   "data"     // input data directory (relative to repo root)
global output "output"   // output directory for .xlsx results

set type double

* ---- Load data and set survey design --------------------------------------
* Keep all years needed for multi-year templates (Types A and D).
* Adjust the inlist() if you want a different year range.
use "${data}/hhmultiyear_analys.dta", clear
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

* ---- Year variable (needed for Types A and D) ----------------------------
* hryear4 is already in the multiyear data; rename for consistency with templates.
rename hryear4 year
label variable year "Survey year"

* ---- all_hh: constant variable for "All households" rows in collect ------
* table+collect needs a row variable. Using a constant gives the national-
* total row without any subgroup restriction.
gen byte all_hh = 1
label define all_hh_lbl 1 "All households"
label values all_hh all_hh_lbl

* ---- Demographic variable globals ----------------------------------------
* $Controls1 through $Controls$nControls are the household characteristic
* variables used in every demographic breakdown table. Edit this list to
* add, remove, or reorder characteristics; all templates pick them up
* automatically.

global nControls 7
global Controls1 praceeth
global Controls2 pagegrp
global Controls3 peducgrp
global Controls4 hhincome2
global Controls5 pdisabl_age25to64
global Controls6 hhtypev2
global Controls7 gtmetsta

* ---- build_demo_rows: row definitions for Template A --------------------
* Sets lab1...labN, dv1...dvN, dval1...dvalN, and nrows in the CALLING
* environment via c_local. Template A calls this once instead of listing
* all rows inline. Edit here to change row labels or add/remove categories.

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
* WHY NOT table+collect: the diff column requires storing the SE from each
* year's estimate and computing z = diff / sqrt(SE1² + SE2²). The collect
* framework does not compute across-column expressions. We keep the
* loop+matrix approach and store both estimates and SEs so the diff and
* significance star can be computed here (rather than in a separate script).
*
* Output layout:
*   Group               | 2019  | 2021  | 2023  | Diff  | Sig
*   All households      |  5.4  |  4.5  |  4.2  | -0.3  |
*   Race/Ethnicity       |       |       |       |       |
*     Black             | 13.8  | 11.3  |  9.3  | -2.0  |  **
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local outcome    unbanked
local years      2019 2021 2023
local diff_from  2021
local diff_to    2023
local subpop     ""            // "" = all HH; e.g. "hunbnk==2" = banked only
local xlfile     "${output}/table_A_example.xlsx"
local sheet      "Table 1.1"
local filemode   replace       // "replace" for new workbook; "modify" to add sheet
* ---- END CONFIGURE ---------------------------------------------------------

local nyears = wordcount("`years'")
local ncols  = `nyears' + 2    // years + diff + sig (sig stored separately)

* Get demographic row definitions from build_demo_rows
build_demo_rows

* Find column indices for diff computation
local col_from = 0
local col_to   = 0
local c = 0
foreach yr of local years {
    local ++c
    if "`yr'" == "`diff_from'" local col_from = `c'
    if "`yr'" == "`diff_to'"   local col_to   = `c'
}
if `col_from'==0 | `col_to'==0 {
    di as error "diff_from/diff_to not found in years list"
    exit 198
}

* Matrices: M for estimates (%), SE for standard errors (%)
matrix M  = J(`nrows', `nyears', .)
matrix SE = J(`nrows', `nyears', .)

local colnames "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"

* ---- Estimation loop -------------------------------------------------------
forvalues i = 1/`nrows' {
    if "`dv`i''" == "header" continue

    local c = 0
    foreach yr of local years {
        local ++c
        if "`subpop'" == "" local sp_cond "year`yr'==1"
        else                 local sp_cond "year`yr'==1 & (`subpop')"

        if "`dv`i''" == "all" {
            capture quietly svy, subpop(if `sp_cond'): mean `outcome'
        }
        else {
            capture quietly svy, subpop(if `sp_cond' & `dv`i''==`dval`i''): mean `outcome'
        }

        if _rc == 0 & e(N_sub) >= 5 {
            matrix M[`i',  `c'] = _b[`outcome'] * 100
            matrix SE[`i', `c'] = _se[`outcome'] * 100
        }
    }
}

* ---- Diff and significance -------------------------------------------------
* diff = estimate_to - estimate_from (percentage points)
* z    = diff / sqrt(SE_to^2 + SE_from^2)  [independent surveys → add variances]
* sig  = "**" if |z|>2.576 (p<.01), "*" if |z|>1.96 (p<.05)

matrix DIFF = J(`nrows', 1, .)
local sig_strings ""   // will build a list of sig strings, one per row

forvalues i = 1/`nrows' {
    if "`dv`i''" == "header" {
        local sig_strings `"`sig_strings' """'
        continue
    }
    local est_to   = M[`i',  `col_to']
    local est_from = M[`i',  `col_from']
    local se_to    = SE[`i', `col_to']
    local se_from  = SE[`i', `col_from']

    if `est_to' < . & `est_from' < . {
        local d = `est_to' - `est_from'
        matrix DIFF[`i', 1] = `d'
        local sig_str ""
        if `se_to' < . & `se_from' < . & (`se_to'^2 + `se_from'^2) > 0 {
            local z = abs(`d') / sqrt(`se_to'^2 + `se_from'^2)
            if `z' > 2.576 local sig_str "**"
            else if `z' > 1.96 local sig_str "*"
        }
        local sig_strings `"`sig_strings' "`sig_str'""'
    }
    else {
        local sig_strings `"`sig_strings' """'
    }
}

* ---- Export to Excel -------------------------------------------------------
putexcel set "`xlfile'", sheet("`sheet'") `filemode'

* Column headers
putexcel A1 = "Group"
local c = 1
foreach yr of local years {
    local ++c
    putexcel `=word("`colnames'",`c')'1 = "`yr'"
}
local c_diff = `nyears' + 2
local c_sig  = `nyears' + 3
putexcel `=word("`colnames'",`c_diff')'1 = "Diff (`diff_to'–`diff_from')"
putexcel `=word("`colnames'",`c_sig')'1  = "Sig"

* Row labels
forvalues i = 1/`nrows' {
    putexcel A`=`i'+1' = "`lab`i''"
}

* Year estimates and diff (two matrix calls)
putexcel B2 = matrix(M),    nformat("0.0")
putexcel `=word("`colnames'",`c_diff')'2 = matrix(DIFF), nformat("0.0")

* Significance strings (individual cells — unavoidable for text)
local i = 0
foreach s of local sig_strings {
    local ++i
    if "`s'" != "" putexcel `=word("`colnames'",`c_sig')'`=`i'+1' = "`s'"
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
* Report examples: Table 3.3, 5.1
*
* Uses table + collect. The $Controls globals drive the demographic loop.
* One svy: table call per characteristic; collect stacks the results.
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local outcomes   "huse12mo_b huse12cc_b huse12mt_b"
local out_labels `""Money order" "Check cashing" "Money transfer""'
local year       2023
local subpop     ""
local xlfile     "${output}/table_B_example.xlsx"
local sheet      "Table B"
* ---- END CONFIGURE ---------------------------------------------------------

* Build the statistic() string: one statistic(mean X) clause per outcome
local stat_string ""
foreach out of local outcomes {
    local stat_string "`stat_string' statistic(mean `out')"
}

* Build subpop condition
if "`subpop'" == "" local sp "year`year'==1"
else                 local sp "year`year'==1 & (`subpop')"

* ---- Collect estimates -----------------------------------------------------
collect clear

* National total (all_hh is a constant = 1, created in SETUP)
collect, tag(demo[all_hh]): ///
    svy, subpop(if `sp'): table all_hh, `stat_string'

* One table call per demographic characteristic
forvalues i = 1/$nControls {
    local dv ${Controls`i'}
    collect, tag(demo[`dv']): ///
        svy, subpop(if `sp'): table `dv', `stat_string'
}

* ---- Layout and labels -----------------------------------------------------
* Relabel result dimension so columns show "Money order" etc., not "mean(var)"
local k = 0
foreach out of local outcomes {
    local ++k
    local lbl : word `k' of `out_labels'
    collect label levels result "mean(`out')" "`lbl'", modify
}

* Build row dimension: stack all_hh first, then each Controls variable
local row_dim "demo[all_hh]#all_hh"
forvalues i = 1/$nControls {
    local row_dim "`row_dim' demo[${Controls`i'}]#${Controls`i'}"
}

collect layout (`row_dim') (result)
collect style cell, nformat(%6.1f)

collect export "`xlfile'", sheet("`sheet'") replace
di "Template B complete → `xlfile' [sheet: `sheet']"


*=============================================================================
* TEMPLATE C — Categorical distribution (rows = categories, one column)
*
* Report examples: Table 1.2 (account-transition status), Table 3.2 (OPS/
*                  prepaid combination)
*
* Uses svy: proportion + collect. One command; collect handles the layout.
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local catvar   hbankstatv6
local year     2023
local subpop   ""              // "" = all HH; e.g. "hunbnk==1" = unbanked only
local xlfile   "${output}/table_C_example.xlsx"
local sheet    "Table C"
* ---- END CONFIGURE ---------------------------------------------------------

if "`subpop'" == "" local sp "year`year'==1"
else                 local sp "year`year'==1 & (`subpop')"

collect clear
svy, subpop(if `sp'): proportion `catvar'

* _r_b = point estimate; collect layout puts categories as rows, estimate as col
collect layout (`catvar') (result[_r_b])
collect style cell result[_r_b], nformat(%6.1f) transform(* 100)

collect export "`xlfile'", sheet("`sheet'") replace
di "Template C complete → `xlfile' [sheet: `sheet']"


*=============================================================================
* TEMPLATE D — Rows = survey years, columns = ordered response categories
*
* Report examples: Table 1.4 (interest in banking), Table 1.3 (prior banking
*                  status), Table 2.1 (primary access method), Table 2.3
*                  (access methods used)
*
* Uses svy: proportion with a collect tag per year. The year tag becomes
* the row dimension; category levels become columns.
*=============================================================================

* ---- CONFIGURE (edit these) -----------------------------------------------
local catvar   hbankint        // categorical response variable
local years    "2019 2021 2023"
local subpop   "hunbnk==1"     // universe: unbanked HH; "" for all HH
local xlfile   "${output}/table_D_example.xlsx"
local sheet    "Table 1.4"
* ---- END CONFIGURE ---------------------------------------------------------

collect clear

foreach yr of local years {
    if "`subpop'" == "" local sp "year`yr'==1"
    else                 local sp "year`yr'==1 & (`subpop')"

    collect, tag(year[`yr']): ///
        svy, subpop(if `sp'): proportion `catvar'
}

* year = custom tag dimension (rows); catvar = category levels (columns)
collect layout (year) (`catvar')
collect style cell result[_r_b], nformat(%6.1f) transform(* 100)

collect export "`xlfile'", sheet("`sheet'") replace
di "Template D complete → `xlfile' [sheet: `sheet']"

/*
  TABLE 2.3 NOTE: bank-access methods are "check all that apply" — row
  percentages sum above 100. Use Template B with year fixed and outcomes =
  hbnkaccm1v2_b through hbnkaccm6v2_b, rather than Template D.
*/
