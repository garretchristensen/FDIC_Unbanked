/*============================================================================
  FDIC Unbanked Survey — Reproduce Report Tables

  Replicates tables from the 2023 FDIC National Survey of Unbanked and
  Underbanked Households, matching the exact published layout:

    Characteristic | 2019 | 2021 | 2023 | Difference (2023–2021)

  Significance (* p<.10) uses 160 Census Bureau BRR replicate weights
  (Fay factor = 0.5), matching the report's variance methodology.
  Point estimates use hhsupwgt (household supplement weight).

  Run from the repo root directory. Stata 15+.
============================================================================*/

clear all
set more off
set type double

/*----------------------------------------------------------------------------
  SECTION 0 — GLOBALS
----------------------------------------------------------------------------*/
global Y1    2019    // first year column
global Y2    2021    // reference year for Difference column
global Y3    2023    // current (most recent) year

* Minimum unweighted subpop N to publish an estimate; below this → "NA"
global MIN_N 30

global data   "data"
global output "output"

/*----------------------------------------------------------------------------
  SECTION 1 — LOAD DATA AND DECLARE SURVEY DESIGN

  Point estimates: hhsupwgt (FDIC/Census supplement weight)
  Standard errors: 160 BRR replicate weights from hhrep19_23.csv,
                   Fay factor = 0.5 (Census Bureau standard for CPS)
----------------------------------------------------------------------------*/
use "${data}/hhmultiyear_analys.dta", clear
keep if inlist(hryear4, ${Y1}, ${Y2}, ${Y3})

* Import BRR replicate weights and merge on (hryear4, qstnum)
tempfile repwgts
preserve
    import delimited "${data}/hhmultiyear/hhrep19_23.csv", ///
        clear varnames(1)
    keep hryear4 qstnum repwgt1-repwgt160
    keep if inlist(hryear4, ${Y1}, ${Y2}, ${Y3})
    save `repwgts'
restore
merge 1:1 hryear4 qstnum using `repwgts', keep(master match) nogen

* Build brrweight varlist (repwgt1 … repwgt160)
local repvars ""
forvalues i = 1/160 {
    local repvars `repvars' repwgt`i'
}
svyset [pw=hhsupwgt], vce(brr) brrweight(`repvars') fay(0.5) mse

/*----------------------------------------------------------------------------
  SECTION 2 — RECODE OUTCOME VARIABLES
----------------------------------------------------------------------------*/

* ── Banking status ───────────────────────────────────────────────────────────
gen byte unbanked          = (hunbnk == 1)          if inlist(hunbnk, 1, 2)
gen byte underbanked       = (hbankstatv6 == 2)     if inlist(hbankstatv6, 1, 2, 3)
gen byte cashonly_unbanked = (hunbnkcashonly == 1)  if inlist(hunbnkcashonly, 1, 2)

* ── Transaction AFS products ─────────────────────────────────────────────────
foreach v in huse12mo huse12cc huse12mt huse12rm huse12rmany {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* Online payment services and prepaid cards (2021+; 2019 column will show NA)
foreach v in husenowops husenowpp {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* ── AFS credit products ───────────────────────────────────────────────────────
foreach v in huse12pdl huse12pwn huse12ral huse12atl huse12rto {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
* Variable name changed in 2023 multiyear file — try both versions
capture gen byte anyafs_credit = (huse12afscv3 == 1) if inlist(huse12afscv3, 1, 2)
if _rc != 0 capture gen byte anyafs_credit = (huse12afsc == 1) if inlist(huse12afsc, 1, 2)

* ── New 2023 products ────────────────────────────────────────────────────────
foreach v in hcred12bnpl huse12cryp {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

* ── Bank account access methods ──────────────────────────────────────────────
foreach v in hbnkaccm1v2 hbnkaccm5v2 {
    capture gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}

/*----------------------------------------------------------------------------
  SECTION 3 — ESTIMATION HELPER

  fdic_est_row: estimates one demographic cell for all three survey years
  and writes the result to the currently open putexcel workbook.

  Syntax:
    fdic_est_row, outcome(varname) gv(groupvar | "national") gval(#) row(#)
                  [skip_y2]

  skip_y2: leave the Y2 column and Difference column blank.
           Use for variables not collected in Y2 (e.g., monthly income
           volatility not available in 2021).

  Columns written in the current putexcel target:
    Col B = Y1 estimate    Col C = Y2 estimate    Col D = Y3 estimate
    Col E = Difference (Y3 − Y2) with * appended if |z| > 1.645

  "NA" is written when subpop N < MIN_N or estimation fails.
  Blank cell when skip_y2 is set (intentionally not collected).
----------------------------------------------------------------------------*/
* Drop any prior definition so the program can be re-run without "already defined" errors
capture program drop fdic_est_row
* Begin program definition
program define fdic_est_row
    * Required args: outcome variable name, grouping variable name (or "national"),
    * the value of that grouping variable for this row, and the Excel row number.
    * Optional: skip_y2 flag when the outcome was not collected in survey year 2 (2021).
    syntax, outcome(string) gv(string) gval(integer) row(integer) [skip_y2]

    * Short alias so the row number is easier to type throughout
    local r `row'

    * --- ESTIMATION LOOP: run svy: mean for each of the three survey years ---
    forvalues yidx = 1/3 {
        * If the caller set skip_y2 and we're on year 2, store missing and move on
        if `yidx' == 2 & "`skip_y2'" != "" {
            local m`yidx' = .   // no estimate to store
            local s`yidx' = .   // no SE to store
            continue             // skip to yidx == 3
        }
        * Look up the actual four-digit calendar year for this index (e.g. Y1=2019)
        local y = ${Y`yidx'}

        * Run the BRR-weighted mean; "capture quietly" prevents the loop from
        * aborting if the subpopulation is empty or the estimator fails
        if "`gv'" == "national" {
            * National row: no group filter — use the full sample for year y
            capture quietly svy, subpop(if hryear4 == `y'): mean `outcome'
        }
        else {
            * Demographic subgroup row: restrict to the specific group value
            capture quietly svy, subpop(if hryear4 == `y' & `gv' == `gval'): mean `outcome'
        }

        * Only store results when svy succeeded (_rc==0) AND the subpop is large
        * enough to report (at least MIN_N unweighted observations)
        if _rc == 0 & e(N_sub) >= ${MIN_N} {
            mat __B = e(b)                      // 1×1 coefficient matrix (the mean)
            mat __V = e(V)                      // 1×1 variance-covariance matrix
            local m`yidx' = __B[1,1] * 100     // convert proportion → percentage points
            local s`yidx' = sqrt(__V[1,1]) * 100  // convert variance → SE in pct-pts
        }
        else {
            * Estimation failed or subpop too small — flag both as missing
            local m`yidx' = .
            local s`yidx' = .
        }
    }

    * --- OUTPUT LOOP: write point estimates into Excel columns B, C, D ---
    forvalues yidx = 1/3 {
        * Map the loop index to the corresponding Excel column letter
        local col: word `yidx' of B C D
        if `yidx' == 2 & "`skip_y2'" != "" {
            * Leave cell blank — outcome was not collected in this survey year
        }
        else if !missing(`m`yidx'') {
            * Valid estimate: write as a number formatted to one decimal place
            putexcel `col'`r' = (`m`yidx''), nformat("0.0")
        }
        else {
            * Estimation failed or N too small: write the sentinel string "NA"
            putexcel `col'`r' = "NA"
        }
    }

    * --- DIFFERENCE COLUMN (col E): change from Y2 to Y3 in pct-pts ---
    if "`skip_y2'" != "" {
        * Y2 was not collected, so no meaningful before/after comparison; leave blank
    }
    else if !missing(`m3') & !missing(`m2') {
        * Both years estimated successfully — compute the raw change
        local diff    = `m3' - `m2'
        * SE of the difference under independence: sqrt(Var_Y3 + Var_Y2)
        local se_diff = sqrt(`s3'^2 + `s2'^2)
        * Format to one decimal place and strip leading/trailing spaces
        local dstr    = trim(string(`diff', "%5.1f"))
        if `se_diff' > 0 {
            * Append * when the z-statistic exceeds 1.645 (≈ 90% two-tailed significance)
            if abs(`diff') / `se_diff' > 1.645 local dstr "`dstr'*"
        }
        * Write the formatted string (e.g. "-2.3" or "1.4*") to column E
        putexcel E`r' = "`dstr'"
    }
    else {
        * At least one year unavailable — cannot compute difference
        putexcel E`r' = "NA"
    }
end

/*----------------------------------------------------------------------------
  SECTION 4 — STANDARD DEMOGRAPHIC TABLE PROGRAM

  fdic_std_table: writes a full table matching the layout of Table 1.1 and
  all other standard tables in the 2023 FDIC report.

  Row layout (all tables):
    Row 1:  Table title (bold)
    Row 2:  "All Households, Row Percent"
    Row 4:  Column headers
    Row 5:  All
    Row 6:  Family Income  (section header)
    Rows 7–11:  5 income groups (< $15k / $15–30k / $30–50k / $50–75k / ≥$75k)
    Row 12: Education  (section header)
    Rows 13–16: 4 education groups
    Row 17: Age Group  (section header)
    Rows 18–23: 6 age groups
    Row 24: Race/Ethnicity  (section header)
    Rows 25–31: 7 race/ethnicity groups
    Row 32: Disability Status  (section header)
    Rows 33–34: 2 disability groups (aged 25–64 only)
    Row 35: Monthly Income Volatility  (section header)
    Rows 36–38: 3 volatility groups  (Y2 column and Diff blank — not in 2021)
    Row 39: Note

  Syntax:
    fdic_std_table varname, xlfile(path) sheet(name) title(text) [REPLACE]
----------------------------------------------------------------------------*/
capture program drop fdic_std_table
program define fdic_std_table
    syntax varname, xlfile(string) sheet(string) title(string) [REPLACE]

    local outcome `varlist'
    local fmode modify
    if "`replace'" != "" local fmode replace

    putexcel set "`xlfile'", sheet("`sheet'") `fmode'

    * ── Headers ──────────────────────────────────────────────────────────────
    putexcel A1 = "`title'", bold
    putexcel A2 = "All Households, Row Percent"
    putexcel A4 = "Characteristic",           bold
    putexcel B4 = "${Y1}",                    bold
    putexcel C4 = "${Y2}",                    bold
    putexcel D4 = "${Y3}",                    bold
    putexcel E4 = "Difference (${Y3}–${Y2})", bold

    * ── Row 5: All households ────────────────────────────────────────────────
    putexcel A5 = "All", bold
    fdic_est_row, outcome(`outcome') gv(national) gval(0) row(5)

    * ── Rows 6–11: Family Income ─────────────────────────────────────────────
    putexcel A6 = "Family Income", bold
    local inc_lbl `""Less Than $15,000" "$15,000 to $30,000" "$30,000 to $50,000" "$50,000 to $75,000" "At Least $75,000""'
    forvalues i = 1/5 {
        local lbl: word `i' of `inc_lbl'
        putexcel A`=6+`i'' = "  `lbl'"
        fdic_est_row, outcome(`outcome') gv(hhincome2) gval(`i') row(`=6+`i'')
    }

    * ── Rows 12–16: Education ────────────────────────────────────────────────
    putexcel A12 = "Education", bold
    local educ_lbl `""No High School Diploma" "High School Diploma" "Some College" "College Degree""'
    forvalues i = 1/4 {
        local lbl: word `i' of `educ_lbl'
        putexcel A`=12+`i'' = "  `lbl'"
        fdic_est_row, outcome(`outcome') gv(peducgrp) gval(`i') row(`=12+`i'')
    }

    * ── Rows 17–23: Age Group ────────────────────────────────────────────────
    putexcel A17 = "Age Group", bold
    local age_lbl `""15 to 24 Years" "25 to 34 Years" "35 to 44 Years" "45 to 54 Years" "55 to 64 Years" "65 Years or More""'
    forvalues i = 1/6 {
        local lbl: word `i' of `age_lbl'
        putexcel A`=17+`i'' = "  `lbl'"
        fdic_est_row, outcome(`outcome') gv(pagegrp) gval(`i') row(`=17+`i'')
    }

    * ── Rows 24–31: Race/Ethnicity ───────────────────────────────────────────
    putexcel A24 = "Race/Ethnicity", bold
    local race_lbl `""Black" "Hispanic" "Asian" "American Indian or Alaska Native" "Native Hawaiian or Other Pacific Islander" "White" "Two or More Races""'
    forvalues i = 1/7 {
        local lbl: word `i' of `race_lbl'
        putexcel A`=24+`i'' = "  `lbl'"
        fdic_est_row, outcome(`outcome') gv(praceeth) gval(`i') row(`=24+`i'')
    }

    * ── Rows 32–34: Disability Status ────────────────────────────────────────
    putexcel A32 = "Disability Status", bold
    putexcel A33 = "  Disabled, Aged 25 to 64"
    fdic_est_row, outcome(`outcome') gv(pdisabl_age25to64) gval(1) row(33)
    putexcel A34 = "  Not Disabled, Aged 25 to 64"
    fdic_est_row, outcome(`outcome') gv(pdisabl_age25to64) gval(2) row(34)

    * ── Rows 35–38: Monthly Income Volatility ────────────────────────────────
    * hincvol not collected in ${Y2} (2021). Y2 column and Diff column left blank.
    putexcel A35 = "Monthly Income Volatility", bold
    putexcel A36 = "  Income Was About the Same Each Month"
    fdic_est_row, outcome(`outcome') gv(hincvol) gval(1) row(36) skip_y2
    putexcel A37 = "  Income Varied Somewhat From Month to Month"
    fdic_est_row, outcome(`outcome') gv(hincvol) gval(2) row(37) skip_y2
    putexcel A38 = "  Income Varied a Lot From Month to Month"
    fdic_est_row, outcome(`outcome') gv(hincvol) gval(3) row(38) skip_y2

    * ── Row 39: Note ──────────────────────────────────────────────────────────
    putexcel A39 = "Note: Monthly income volatility not available for ${Y2}. * = statistically significant at the 10 percent level. NA = sample too small to produce a reliable estimate. Standard errors from 160 Census Bureau BRR replicate weights (Fay factor = 0.5)."

    di "  --> `sheet' written to `xlfile'"
end

/*----------------------------------------------------------------------------
  SECTION 5 — WRITE ALL TABLES

  First call to each .xlsx file uses , replace to create it fresh.
  Subsequent sheets in the same file use the default , modify.
----------------------------------------------------------------------------*/

di _n "Writing tables..."

* ── Table 1.1: Unbanked rates ─────────────────────────────────────────────────
fdic_std_table unbanked, ///
    xlfile("${output}/table_1_1_unbanked_rates.xlsx") ///
    sheet("Table 1.1") ///
    title("TABLE 1.1 Unbanked Rates by Selected Household Characteristics, ${Y1}–${Y3}") ///
    replace

* ── Table 7.1: Underbanked rates ─────────────────────────────────────────────
fdic_std_table underbanked, ///
    xlfile("${output}/table_7_1_underbanked_rates.xlsx") ///
    sheet("Table 7.1") ///
    title("TABLE 7.1 Underbanked Rates by Selected Household Characteristics, ${Y1}–${Y3}") ///
    replace

* ── Table 2.2: Bank account access methods ───────────────────────────────────
fdic_std_table hbnkaccm1v2_b, ///
    xlfile("${output}/table_2_2_bank_access.xlsx") ///
    sheet("Table 2.2a Teller") ///
    title("TABLE 2.2a Bank Teller as Primary Method of Bank Account Access, ${Y1}–${Y3}") ///
    replace

fdic_std_table hbnkaccm5v2_b, ///
    xlfile("${output}/table_2_2_bank_access.xlsx") ///
    sheet("Table 2.2b Mobile") ///
    title("TABLE 2.2b Mobile Banking as Primary Method of Bank Account Access, ${Y1}–${Y3}")

* ── Table 3.1: Nonbank online payment services and prepaid cards ─────────────
* Note: husenowops and husenowpp first collected in 2021; Y1 column shows NA
fdic_std_table husenowops_b, ///
    xlfile("${output}/table_3_1_ops_prepaid.xlsx") ///
    sheet("Table 3.1a OPS") ///
    title("TABLE 3.1a Use of Nonbank Online Payment Services, ${Y1}–${Y3}") ///
    replace

fdic_std_table husenowpp_b, ///
    xlfile("${output}/table_3_1_ops_prepaid.xlsx") ///
    sheet("Table 3.1b Prepaid") ///
    title("TABLE 3.1b Use of Prepaid Cards, ${Y1}–${Y3}")

* ── Table 3.3: Cash-only unbanked ────────────────────────────────────────────
fdic_std_table cashonly_unbanked, ///
    xlfile("${output}/table_3_3_cashonly.xlsx") ///
    sheet("Table 3.3") ///
    title("TABLE 3.3 Cash-Only Unbanked Rates by Selected Household Characteristics, ${Y1}–${Y3}") ///
    replace

* ── Table 4.1: Nonbank transaction services ──────────────────────────────────
fdic_std_table huse12mo_b, ///
    xlfile("${output}/table_4_1_transaction_afs.xlsx") ///
    sheet("Table 4.1a Money Order") ///
    title("TABLE 4.1a Use of Nonbank Money Orders, ${Y1}–${Y3}") ///
    replace

fdic_std_table huse12cc_b, ///
    xlfile("${output}/table_4_1_transaction_afs.xlsx") ///
    sheet("Table 4.1b Check Cash") ///
    title("TABLE 4.1b Use of Nonbank Check Cashing, ${Y1}–${Y3}")

fdic_std_table huse12mt_b, ///
    xlfile("${output}/table_4_1_transaction_afs.xlsx") ///
    sheet("Table 4.1c Money Transfer") ///
    title("TABLE 4.1c Use of Nonbank Money Transfer Services, ${Y1}–${Y3}")

* ── Table 5.3: Any AFS credit product ───────────────────────────────────────
fdic_std_table anyafs_credit, ///
    xlfile("${output}/table_5_3_afs_credit.xlsx") ///
    sheet("Table 5.3 Any AFS credit") ///
    title("TABLE 5.3 Use of Any Alternative Financial Credit Products, ${Y1}–${Y3}") ///
    replace

di _n "Done. All tables written to ${output}/."
di "Layout: Characteristic | ${Y1} | ${Y2} | ${Y3} | Difference (${Y3}–${Y2})"
di "  * p<.10 (two-tailed)   NA = N < ${MIN_N}   ${Y2} blank for Monthly Income Volatility"
