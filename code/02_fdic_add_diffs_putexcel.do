/*============================================================================
  FDIC Unbanked Survey — Add Year-over-Year Differences to Excel Tables

  Run AFTER 01_fdic_tables_putexcel.do for the current year's data.
  Reads data/hhmultiyear_analys.dta directly — no pre-saved estimates files needed.

  What this script does:
    1. Loads the multiyear dataset and creates outcome variables
    2. Computes weighted estimates for CUR_YEAR and PREV_YEAR in a single pass
    3. Computes difference and z-test for significance
    4. Opens each Excel file and fills the Diff (col F) and Sig (col G) columns

  Significance convention (two-tailed):
    |z| > 1.96  → *    (p < .05)
    |z| > 2.576 → **   (p < .01)

  Run from the repo root directory.
  Stata version: 17+
============================================================================*/

clear all
set more off
set type double

global CUR_YEAR  2023    // year whose tables you want to update
global PREV_YEAR 2021    // year to compare against

global data "data"   // relative to repo root

/*----------------------------------------------------------------------------
  STEP 1 — Load data and create outcome variables
  Keep both survey years so we can estimate each year via svy subpop.
----------------------------------------------------------------------------*/
use "${data}/hhmultiyear_analys.dta", clear
keep if inlist(hryear4, ${CUR_YEAR}, ${PREV_YEAR})

svyset [pw=hhsupwgt]

* ── Banking status ───────────────────────────────────────────────────────────
gen byte unbanked    = (hunbnk == 1)       if inlist(hunbnk, 1, 2)
gen byte underbanked = (hbankstatv6 == 2)  if inlist(hbankstatv6, 1, 2, 3)

* ── AFS transaction products ─────────────────────────────────────────────────
foreach v in huse12mo huse12cc huse12mt huse12rm huse12rmany {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* Online payment services and prepaid cards (2021+)
foreach v in husenowops husenowpp {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* ── AFS credit products ───────────────────────────────────────────────────────
foreach v in huse12pdl huse12pwn huse12ral huse12atl huse12rto {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
capture gen byte anyafs_credit = (huse12AFSCv3 == 1) if inlist(huse12AFSCv3, 1, 2)
if _rc != 0 capture gen byte anyafs_credit = (huse12AFSC == 1) if inlist(huse12AFSC, 1, 2)

* ── New products ─────────────────────────────────────────────────────────────
foreach v in hcred12bnpl hcred12bnpldq huse12cryp {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* ── Bank account access methods ──────────────────────────────────────────────
foreach v in hbnkaccm1v2 hbnkaccm5v2 {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* ── Cash-only unbanked ───────────────────────────────────────────────────────
capture gen byte cashonly_unbanked = (hunbnkcashonly == 1) if inlist(hunbnkcashonly, 1, 2)

/*----------------------------------------------------------------------------
  STEP 2 — Compute estimates for both years
  One postfile row per (outcome, group_var, group_val, year).
  Then reshape wide and compute diff.
----------------------------------------------------------------------------*/
local all_outcomes ///
    unbanked underbanked cashonly_unbanked ///
    huse12mo_b huse12cc_b huse12mt_b huse12rmany_b huse12rm_b ///
    husenowops_b husenowpp_b ///
    huse12pdl_b huse12pwn_b huse12ral_b huse12atl_b huse12rto_b anyafs_credit ///
    hcred12bnpl_b hcred12bnpldq_b huse12cryp_b ///
    hbnkaccm1v2_b hbnkaccm5v2_b

local demo_gvars   praceeth  pagegrp  peducgrp  hhincome2  hhtypev2  gtmetsta  hbankstatv6
local demo_maxvals 7         6        4         99         8         3         3

tempfile est_long
postfile pf_est                   ///
    str40  outcome                ///
    str40  group_var              ///
    double group_val              ///
    double survey_year            ///
    double mean_pct               ///
    double se_pct                 ///
    using `est_long', replace

foreach o of local all_outcomes {
    foreach yr in ${CUR_YEAR} ${PREV_YEAR} {

        * National
        capture quietly svy, subpop(if hryear4==`yr'): mean `o'
        if _rc == 0 & e(N_sub) >= 5 {
            post pf_est ("`o'") ("national") (0) (`yr') ///
                (`=e(b)[1,1]*100') (`=sqrt(e(V)[1,1])*100')
        }

        * Demographic subgroups
        local gi = 0
        foreach gv of local demo_gvars {
            local ++gi
            local maxv : word `gi' of `demo_maxvals'
            local vals ""
            if "`gv'" == "hhincome2" {
                local vals "1 2 3 4 5 99"
            }
            else {
                forvalues v = 1/`maxv' {
                    local vals `vals' `v'
                }
            }
            foreach val of local vals {
                capture quietly svy, subpop(if hryear4==`yr' & `gv'==`val'): mean `o'
                if _rc == 0 & e(N_sub) >= 5 {
                    post pf_est ("`o'") ("`gv'") (`val') (`yr') ///
                        (`=e(b)[1,1]*100') (`=sqrt(e(V)[1,1])*100')
                }
            }
        }
    }
}
postclose pf_est

/*----------------------------------------------------------------------------
  STEP 3 — Reshape wide and compute differences
----------------------------------------------------------------------------*/
use `est_long', clear

* Reshape: one row per (outcome, group_var, group_val); columns for each year
reshape wide mean_pct se_pct, i(outcome group_var group_val) j(survey_year)

rename mean_pct${CUR_YEAR}  mean_cur
rename se_pct${CUR_YEAR}    se_cur
rename mean_pct${PREV_YEAR} mean_prev
rename se_pct${PREV_YEAR}   se_prev

gen double diff    = mean_cur - mean_prev
gen double se_diff = sqrt(se_cur^2 + se_prev^2)
gen double z_stat  = diff / se_diff

gen str3 sig = ""
replace sig = "**" if abs(z_stat) > 2.576 & !missing(z_stat)
replace sig = "*"  if abs(z_stat) > 1.96  & abs(z_stat) <= 2.576 & !missing(z_stat)

tempfile alldiffs
save `alldiffs'

/*----------------------------------------------------------------------------
  STEP 4 — Write diff + sig to each Excel table

  Row layout matches what 01_fdic_tables_putexcel.do writes starting at row 5:
    Row 5:  All households
    Rows 8-14:   Race/Ethnicity (praceeth 1-7)
    Rows 17-22:  Age (pagegrp 1-6)
    Rows 25-28:  Education (peducgrp 1-4)
    Rows 31-36:  Income (hhincome2 1,2,3,4,5,99)
    Rows 39-40:  Disability (pdisabl_age25to64 1-2)
    Rows 43-50:  HH type (hhtypev2 1-8)
    Rows 53-55:  Metro (gtmetsta 1-3)

  Col F = Diff, Col G = Sig  (matching the header written by script 01)
----------------------------------------------------------------------------*/

* Row-to-(group_var, group_val) lookup
quietly {
    tempfile lookup
    clear
    input str40 group_var double group_val double xlrow
    "national"           0    5
    "praceeth"           1    8
    "praceeth"           2    9
    "praceeth"           3   10
    "praceeth"           4   11
    "praceeth"           5   12
    "praceeth"           6   13
    "praceeth"           7   14
    "pagegrp"            1   17
    "pagegrp"            2   18
    "pagegrp"            3   19
    "pagegrp"            4   20
    "pagegrp"            5   21
    "pagegrp"            6   22
    "peducgrp"           1   25
    "peducgrp"           2   26
    "peducgrp"           3   27
    "peducgrp"           4   28
    "hhincome2"          1   31
    "hhincome2"          2   32
    "hhincome2"          3   33
    "hhincome2"          4   34
    "hhincome2"          5   35
    "hhincome2"         99   36
    "pdisabl_age25to64"  1   39
    "pdisabl_age25to64"  2   40
    "hhtypev2"           1   43
    "hhtypev2"           2   44
    "hhtypev2"           3   45
    "hhtypev2"           4   46
    "hhtypev2"           5   47
    "hhtypev2"           6   48
    "hhtypev2"           7   49
    "hhtypev2"           8   50
    "gtmetsta"           1   53
    "gtmetsta"           2   54
    "gtmetsta"           3   55
    end
    save `lookup'
}

capture program drop write_diffs
program define write_diffs
    syntax, outcome(string) xlfile(string) sheet(string) alldiffs(string)

    use "`alldiffs'", clear
    keep if outcome == "`outcome'"
    keep group_var group_val diff sig
    merge 1:1 group_var group_val using "`c(alldiffs_lookup)'", keep(match) nogen
    * Note: caller must set global alldiffs_lookup to the lookup tempfile path

    putexcel set "`xlfile'", sheet("`sheet'") modify
    forvalues i = 1/`=_N' {
        local r = xlrow[`i']
        local d = diff[`i']
        local s = sig[`i']
        if !missing(`d')          putexcel F`r' = (`d'), nformat("0.0")
        if "`s'" != "" & "`s'" != "." putexcel G`r' = "`s'"
    }
    di "  --> diffs written to `xlfile' [sheet: `sheet']"
end

* Pass lookup path via global so the program can access it
global alldiffs_lookup `lookup'

* Re-usable call block: merge diffs + lookup, then write
capture program drop apply_diffs
program define apply_diffs
    syntax, outcome(string) xlfile(string) sheet(string)

    use "${alldiffs_tempfile}", clear
    keep if outcome == "`outcome'"
    keep group_var group_val diff sig
    merge 1:1 group_var group_val using "${alldiffs_lookup}", keep(match) nogen

    putexcel set "`xlfile'", sheet("`sheet'") modify
    forvalues i = 1/`=_N' {
        local r = xlrow[`i']
        local d = diff[`i']
        local s = sig[`i']
        if !missing(`d')               putexcel F`r' = (`d'), nformat("0.0")
        if "`s'" != "" & "`s'" != "."  putexcel G`r' = "`s'"
    }
    di "  --> diffs written to `xlfile' [sheet: `sheet']"
end

global alldiffs_tempfile `alldiffs'

* Table 1.1 — Unbanked
apply_diffs, outcome("unbanked") ///
    xlfile("${data}/table_1_1_unbanked_rates.xlsx") ///
    sheet("Table 1.1")

* Table 2.2 — Teller and mobile banking
apply_diffs, outcome("hbnkaccm1v2_b") ///
    xlfile("${data}/table_2_bank_access.xlsx") ///
    sheet("Table 2.2a Teller")

apply_diffs, outcome("hbnkaccm5v2_b") ///
    xlfile("${data}/table_2_bank_access.xlsx") ///
    sheet("Table 2.2b Mobile")

* Table 3.1 — OPS and prepaid
apply_diffs, outcome("husenowops_b") ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.1a OPS")

apply_diffs, outcome("husenowpp_b") ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.1b Prepaid")

* Table 3.3 — Cash-only unbanked
apply_diffs, outcome("cashonly_unbanked") ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.3 Cash-only")

* Table 4.1 — Transaction AFS
apply_diffs, outcome("huse12mo_b") ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1a Money Order")

apply_diffs, outcome("huse12cc_b") ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1b Check Cash")

apply_diffs, outcome("huse12mt_b") ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1c Money Transfer")

* Table 5.3 — Any AFS credit
apply_diffs, outcome("anyafs_credit") ///
    xlfile("${data}/table_5_credit.xlsx") ///
    sheet("Table 5.3 Any AFS credit")

* Table 7.1 — Underbanked
apply_diffs, outcome("underbanked") ///
    xlfile("${data}/table_7_underbanked_rates.xlsx") ///
    sheet("Table 7.1")

di _n "All diffs written."
di "Review Diff (col F) and Sig (col G) columns in each Excel file."
di "Stars: * = p<.05, ** = p<.01  (two-tailed z-test, independent surveys)"
