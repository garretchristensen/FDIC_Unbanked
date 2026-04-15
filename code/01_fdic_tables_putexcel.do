/*============================================================================
  FDIC National Survey of Unbanked and Underbanked Households
  Reproduce tables from the Annual Report

  Author:  [Your name]
  Data:    data/hhmultiyear_analys.dta  (built from FDIC hhmultiyears.zip; see README)
  Weight:  hhsupwgt  (household supplement weight)
  Output:  data/  (one .xlsx per table section)

  Run from the repo root directory.

  To adapt for a new year:
    1. Update YEAR and PREV_YEAR globals in Section 0
    2. Verify variable names against the new year's metadata
    3. Run 02_fdic_add_diffs_putexcel.do after this script to fill diff columns

  Stata version: 17+  (uses putexcel, svy subpop; collect not required)
============================================================================*/

/*----------------------------------------------------------------------------
  SECTION 0 — GLOBALS AND SETUP
----------------------------------------------------------------------------*/
clear all
set more off
set type double

* ── User-configurable ────────────────────────────────────────────────────────
global YEAR       2023
global PREV_YEAR  2021

global data   "data"   // relative to repo root; run this script from repo root

/*----------------------------------------------------------------------------
  SECTION 1 — LOAD DATA AND DECLARE SURVEY DESIGN
----------------------------------------------------------------------------*/
use "${data}/hhmultiyear_analys.dta", clear
keep if hryear4 == ${YEAR}

* CPS supplement: probability weight only (no explicit cluster/strata released)
svyset [pw=hhsupwgt]

/*----------------------------------------------------------------------------
  SECTION 2 — RECODE AND CREATE VARIABLES
  AFS product variables use 1=YES / 2=NO / 99=Unknown in the raw data.
  We recode to 1=YES / 0=NO (99 and any other value → missing).
----------------------------------------------------------------------------*/

* ── Banking status ───────────────────────────────────────────────────────────
gen byte unbanked    = (hunbnk == 1)       if inlist(hunbnk, 1, 2)
gen byte underbanked = (hbankstatv6 == 2)  if inlist(hbankstatv6, 1, 2, 3)
gen byte fullybanked = (hbankstatv6 == 3)  if inlist(hbankstatv6, 1, 2, 3)
gen byte banked      = (hunbnk == 2)       if inlist(hunbnk, 1, 2)

label var unbanked    "Unbanked"
label var underbanked "Underbanked (among all HH; 0 = banked-fully or unbanked)"
label var fullybanked "Fully banked"
label var banked      "Has bank account"

* ── AFS transaction products ─────────────────────────────────────────────────
foreach v in huse12mo huse12cc huse12mt huse12rm {
    gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
label var huse12mo_b  "Used nonbank money order (past 12 mo)"
label var huse12cc_b  "Used nonbank check cashing (past 12 mo)"
label var huse12mt_b  "Used nonbank money transfer (past 12 mo)"
label var huse12rm_b  "Used intl remittance – nonbank (past 12 mo)"

* Online payment services and prepaid cards (2021+)
foreach v in husenowops husenowpp {
    gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
label var husenowops_b "Used nonbank online payment service"
label var husenowpp_b  "Used prepaid card"

* ── AFS credit products ───────────────────────────────────────────────────────
foreach v in huse12pdl huse12pwn huse12ral huse12atl huse12rto {
    gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
label var huse12pdl_b "Used payday loan (past 12 mo)"
label var huse12pwn_b "Used pawn shop loan (past 12 mo)"
label var huse12ral_b "Used refund anticipation loan (past 12 mo)"
label var huse12atl_b "Used auto title loan (past 12 mo)"
label var huse12rto_b "Used rent-to-own (past 12 mo)"

* Any nonbank credit — variable name changed across survey years in multiyear file:
*   huse12afsc   = 2013–2021  (lowercase in CSV)
*   huse12afscv3 = 2023+      (lowercase in CSV)
* Build from whichever version exists for the loaded year.
capture gen byte anyafs_credit = (huse12afscv3 == 1) if inlist(huse12afscv3, 1, 2)
if _rc != 0 capture gen byte anyafs_credit = (huse12afsc == 1) if inlist(huse12afsc, 1, 2)
label var anyafs_credit "Used any nonbank credit product (past 12 mo)"

* ── 2023-new products ────────────────────────────────────────────────────────
foreach v in hcred12bnpl hcred12bnpldq huse12cryp {
    gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
label var hcred12bnpl_b   "Used Buy Now Pay Later (past 12 mo)"
label var hcred12bnpldq_b "Missed/late BNPL payment (among BNPL users)"
label var huse12cryp_b    "Owned or used cryptocurrency (past 12 mo)"

* ── Mainstream credit ────────────────────────────────────────────────────────
foreach v in hcred12cc hcred12sc hcred12car hcred12hmln hcred12sl {
    gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
gen byte hcred12any_b = (hcred12any == 1) if inlist(hcred12any, 1, 2)
label var hcred12cc_b   "Had Visa/MC/AmEx/Discover credit card"
label var hcred12sc_b   "Had store credit card"
label var hcred12car_b  "Had auto loan"
label var hcred12hmln_b "Had mortgage or home equity loan/line"
label var hcred12sl_b   "Had student loan"
label var hcred12any_b  "Used any mainstream credit"

* ── Bank account access methods (2019+; universe: banked HHs that accessed) ──
* hbnkaccmNv2: 1=Yes 2=No -1=NIU (not a banked HH that accessed)
foreach v in hbnkaccm1v2 hbnkaccm2v2 hbnkaccm3v2 hbnkaccm4v2 hbnkaccm5v2 hbnkaccm6v2 {
    gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2)
}
label var hbnkaccm1v2_b "Used bank teller"
label var hbnkaccm2v2_b "Used ATM/kiosk"
label var hbnkaccm3v2_b "Used telephone banking"
label var hbnkaccm4v2_b "Used online banking"
label var hbnkaccm5v2_b "Used mobile banking"
label var hbnkaccm6v2_b "Used other method"

* ── Cash-only unbanked (pre-computed in data) ────────────────────────────────
* hunbnkcashonly: 1=Cash-only unbanked  2=Not cash-only (banked or not cash-only unbanked)
gen byte cashonly_unbanked = (hunbnkcashonly == 1) if inlist(hunbnkcashonly, 1, 2)
label var cashonly_unbanked "Cash-only unbanked (no OPS, no prepaid)"

* ── Remittance: any method ───────────────────────────────────────────────────
gen byte huse12rmany_b = (huse12rmany == 1) if inlist(huse12rmany, 1, 2)
label var huse12rmany_b "Sent intl remittance – any method (past 12 mo)"

* ── Value labels for demographic variables ───────────────────────────────────
label define race_lbl   1 "Black" 2 "Hispanic" 3 "Asian" 4 "AIAN" 5 "NHOPI" ///
                        6 "White" 7 "Other/Multiracial"
label values praceeth   race_lbl

label define age_lbl    1 "15-24" 2 "25-34" 3 "35-44" 4 "45-54" 5 "55-64" 6 "65+"
label values pagegrp    age_lbl

label define educ_lbl   1 "Less than high school" 2 "High school diploma" ///
                        3 "Some college" 4 "College degree"
label values peducgrp   educ_lbl

label define inc_lbl    1 "Less than $15,000" 2 "$15,000-$30,000" 3 "$30,000-$50,000" ///
                        4 "$50,000-$75,000" 5 "$75,000 or more" 99 "Unknown"
label values hhincome2  inc_lbl

label define hhtype_lbl 1 "Married couple" 2 "Single-mother household" ///
    3 "Other female-HH family" 4 "Single-father household" ///
    5 "Other male-HH family"   6 "Female-HH nonfamily" ///
    7 "Male-HH nonfamily"      8 "Other"
label values hhtypev2   hhtype_lbl

label define metro_lbl  1 "Metropolitan" 2 "Nonmetropolitan" 3 "Not identified"
label values gtmetsta   metro_lbl

label define disab_lbl  1 "Disabled (age 25-64)" 2 "Not disabled (age 25-64)"
label values pdisabl_age25to64 disab_lbl

/*----------------------------------------------------------------------------
  SECTION 3 — HELPER PROGRAM: fdic_demo_table
  Writes a standard demographic breakdown table for one binary (0/1) outcome.

  Syntax:
    fdic_demo_table outcome,
        xlfile(path)          full path to output .xlsx file
        sheet(name)           worksheet name
        title(text)           table title (row 1)
        [subpop(condition)]   if-condition to restrict universe
                              e.g. subpop("banked==1")
                              Uses svy subpop for correct variance estimation.
        [replace]             pass on FIRST call to each .xlsx file to create
                              it fresh; subsequent calls use modify (default)

  The program writes:
    Col A = characteristic label
    Col B = weighted % (current year)
    Col C = standard error
    Col D = unweighted N (for this subgroup × demographic cell)
    Col E = placeholder for prior-year % (fill via fdic_add_diff.do)
    Col F = placeholder for significance stars
----------------------------------------------------------------------------*/
capture program drop fdic_demo_table
program define fdic_demo_table
    syntax varname, xlfile(string) sheet(string) title(string) ///
           [subpop(string) REPLACE]

    local outcome `varlist'
    local fmode   modify
    if "`replace'" != "" local fmode replace

    * Helper: run one svy mean call and store pct/se in locals p and s
    * group_cond = extra condition (combined with subpop via &)
    * Returns: r_pct r_se r_n in locals (sets to "." if subpop empty)
    tempname B V

    * Open / create the Excel file
    putexcel set "`xlfile'", sheet("`sheet'") `fmode'

    * ── Header rows ──────────────────────────────────────────────────────────
    putexcel A1 = "`title'", bold
    putexcel A2 = "Source: ${YEAR} FDIC National Survey of Unbanked and Underbanked Households. Weight: hhsupwgt."
    putexcel A4 = "Characteristic",           bold
    putexcel B4 = "${YEAR} %",                bold
    putexcel C4 = "SE",                        bold
    putexcel D4 = "Unweighted N",              bold
    putexcel E4 = "Prior year % (fill in)",    bold
    putexcel F4 = "Diff (cur - prior)",        bold
    putexcel G4 = "Sig",                       bold

    local row = 5

    * ── Internal helper macro: write one estimate row ─────────────────────────
    * Uses e(b), e(V), e(N_sub) from most-recent svy call.
    * Writes label to col A, pct/SE/N to B/C/D.

    * Build base subpop condition
    local base_sp "`subpop'"   // may be ""

    * ── Utility: one row given extra_cond (combined with base_sp via &) ───────
    * We define this as an inner loop to avoid program-within-program issues.

    * ─────────────────────────────────────────────────────────────────────────
    * NATIONAL (All households / all in subpop)
    * ─────────────────────────────────────────────────────────────────────────
    putexcel A`row' = "All households", bold

    if "`base_sp'" == "" {
        capture quietly svy: mean `outcome'
    }
    else {
        capture quietly svy, subpop(if `base_sp'): mean `outcome'
    }
    if _rc == 0 {
        local pct = e(b)[1,1] * 100
        local se  = sqrt(e(V)[1,1]) * 100
        local n   = e(N_sub)
        if missing(`pct') | e(N_sub) < 5 {
            putexcel B`row' = "—"
        }
        else {
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
    }
    else {
        putexcel B`row' = "—"
    }
    local ++row

    * ─────────────────────────────────────────────────────────────────────────
    * RACE / ETHNICITY
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Race/Ethnicity", bold
    local ++row

    local race_lbls `""Black" "Hispanic" "Asian" "AIAN" "NHOPI" "White" "Other/Multiracial""'
    forvalues r = 1/7 {
        local lbl : word `r' of `race_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "praceeth == `r'"
        }
        else {
            local cond "(`base_sp') & praceeth == `r'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    * ─────────────────────────────────────────────────────────────────────────
    * AGE GROUP
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Age", bold
    local ++row

    local age_lbls `""15-24" "25-34" "35-44" "45-54" "55-64" "65+""'
    forvalues a = 1/6 {
        local lbl : word `a' of `age_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "pagegrp == `a'"
        }
        else {
            local cond "(`base_sp') & pagegrp == `a'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    * ─────────────────────────────────────────────────────────────────────────
    * EDUCATION
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Education", bold
    local ++row

    local educ_lbls `""Less than high school" "High school diploma" "Some college" "College degree""'
    forvalues e = 1/4 {
        local lbl : word `e' of `educ_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "peducgrp == `e'"
        }
        else {
            local cond "(`base_sp') & peducgrp == `e'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    * ─────────────────────────────────────────────────────────────────────────
    * HOUSEHOLD INCOME
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Annual Household Income", bold
    local ++row

    local inc_vals  1 2 3 4 5 99
    local inc_lbls  `""Less than $15,000" "$15,000-$30,000" "$30,000-$50,000" "$50,000-$75,000" "$75,000 or more" "Unknown""'
    local ic = 0
    foreach iv of local inc_vals {
        local ++ic
        local lbl : word `ic' of `inc_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "hhincome2 == `iv'"
        }
        else {
            local cond "(`base_sp') & hhincome2 == `iv'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    * ─────────────────────────────────────────────────────────────────────────
    * DISABILITY STATUS (restricted to age 25-64)
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Disability Status (Age 25-64 only)", bold
    local ++row

    local disab_lbls `""Disabled" "Not disabled""'
    forvalues d = 1/2 {
        local lbl : word `d' of `disab_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "pdisabl_age25to64 == `d'"
        }
        else {
            local cond "(`base_sp') & pdisabl_age25to64 == `d'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    * ─────────────────────────────────────────────────────────────────────────
    * HOUSEHOLD TYPE
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Household Type", bold
    local ++row

    local hhtype_lbls `""Married couple" "Single-mother household" "Other female-HH family" "Single-father household" "Other male-HH family" "Female nonfamily HH" "Male nonfamily HH" "Other""'
    forvalues h = 1/8 {
        local lbl : word `h' of `hhtype_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "hhtypev2 == `h'"
        }
        else {
            local cond "(`base_sp') & hhtypev2 == `h'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    * ─────────────────────────────────────────────────────────────────────────
    * METROPOLITAN STATUS
    * ─────────────────────────────────────────────────────────────────────────
    local ++row
    putexcel A`row' = "Metropolitan Status", bold
    local ++row

    local metro_lbls `""Metropolitan" "Nonmetropolitan" "Not identified""'
    forvalues m = 1/3 {
        local lbl : word `m' of `metro_lbls'
        putexcel A`row' = "  `lbl'"
        if "`base_sp'" == "" {
            local cond "gtmetsta == `m'"
        }
        else {
            local cond "(`base_sp') & gtmetsta == `m'"
        }
        capture quietly svy, subpop(if `cond'): mean `outcome'
        if _rc == 0 & e(N_sub) >= 5 & !missing(e(b)[1,1]) {
            local pct = e(b)[1,1] * 100
            local se  = sqrt(e(V)[1,1]) * 100
            local n   = e(N_sub)
            putexcel B`row' = (`pct'), nformat("0.0")
            putexcel C`row' = (`se'),  nformat("0.0")
            putexcel D`row' = (`n')
        }
        else {
            putexcel B`row' = "—"
        }
        local ++row
    }

    di "  --> `title'"
    di "      written to: `xlfile'  [sheet: `sheet']"
end


/*============================================================================
  TABLE 1.1 — UNBANKED RATES BY HOUSEHOLD CHARACTERISTICS
  Report p.23: % unbanked 2019 / 2021 / 2023 + diff
  Universe: All households
============================================================================*/
di _n "===== TABLE 1.1: Unbanked Rates ====="
fdic_demo_table unbanked, replace ///
    xlfile("${data}/table_1_1_unbanked_rates.xlsx") ///
    sheet("Table 1.1") ///
    title("Table 1.1. Unbanked Rates by Household Characteristics, ${YEAR}")


/*============================================================================
  TABLE 2.1 — PRIMARY METHOD OF BANK ACCOUNT ACCESS (DISTRIBUTION)
  Report p.31: share using each method as PRIMARY, among banked HHs that accessed
  Variable: hbnkaccm (1=Teller 2=ATM 3=Phone 4=Online 5=Mobile 6=Other)
  Universe: Banked HHs that accessed account (hbnkaccm not -1 or missing)
============================================================================*/
di _n "===== TABLE 2.1: Primary Access Method (distribution) ====="

putexcel set "${data}/table_2_bank_access.xlsx", sheet("Table 2.1") replace
putexcel A1 = "Table 2.1. Most Common Way to Access Bank Account, ${YEAR}", bold
putexcel A2 = "Universe: Banked households that accessed account in past 12 months"
putexcel A4 = "Method",        bold
putexcel B4 = "${YEAR} %",    bold
putexcel C4 = "SE",            bold
putexcel D4 = "Unweighted N", bold

local method_lbls `""Bank teller" "ATM or bank kiosk" "Telephone banking" "Online banking" "Mobile banking" "Other""'
local row = 5

forvalues m = 1/6 {
    local lbl : word `m' of `method_lbls'
    putexcel A`row' = "`lbl'"
    * Indicator for each method as the primary method
    capture quietly svy, subpop(if !inlist(hbnkaccm, -1, .)): mean byte(hbnkaccm == `m')
    if _rc == 0 {
        local pct = e(b)[1,1] * 100
        local se  = sqrt(e(V)[1,1]) * 100
        putexcel B`row' = (`pct'), nformat("0.0")
        putexcel C`row' = (`se'),  nformat("0.0")
    }
    else {
        putexcel B`row' = "—"
    }
    local ++row
}
* Unweighted N for the universe
quietly count if !inlist(hbnkaccm, -1, .) & !missing(hbnkaccm)
putexcel D5 = (r(N))
di "  --> Table 2.1 written"


/*============================================================================
  TABLE 2.2 — BANK TELLER AND MOBILE BANKING BY CHARACTERISTICS
  Report p.33: % used teller / % used mobile, 2019/2021/2023 + diff
  Universe: Banked HHs that accessed account (hbnkaccm1v2 / hbnkaccm5v2 not -1)
============================================================================*/
di _n "===== TABLE 2.2: Teller & Mobile Banking by Characteristics ====="

fdic_demo_table hbnkaccm1v2_b, ///
    xlfile("${data}/table_2_bank_access.xlsx") ///
    sheet("Table 2.2a Teller") ///
    title("Table 2.2a. Used Bank Teller, by Household Characteristics, ${YEAR}") ///
    subpop("!inlist(hbnkaccm1v2, -1, .)")

fdic_demo_table hbnkaccm5v2_b, ///
    xlfile("${data}/table_2_bank_access.xlsx") ///
    sheet("Table 2.2b Mobile") ///
    title("Table 2.2b. Used Mobile Banking, by Household Characteristics, ${YEAR}") ///
    subpop("!inlist(hbnkaccm5v2, -1, .)")


/*============================================================================
  TABLE 2.3 — ALL METHODS OF ACCOUNT ACCESS (MULTIPLE RESPONSE)
  Report p.35: % using each method (sums > 100% because multiple methods allowed)
  Universe: Banked HHs that accessed account
  Variables: hbnkaccm1v2_b through hbnkaccm6v2_b
============================================================================*/
di _n "===== TABLE 2.3: All Methods of Access ====="

putexcel set "${data}/table_2_bank_access.xlsx", sheet("Table 2.3") modify
putexcel A1 = "Table 2.3. All Methods Used to Access Bank Account, ${YEAR}", bold
putexcel A2 = "Universe: Banked households that accessed account in past 12 months"
putexcel A3 = "Note: Rows sum to more than 100% because multiple methods allowed"
putexcel A5 = "Method",        bold
putexcel B5 = "${YEAR} %",    bold
putexcel C5 = "SE",            bold

local method_lbls `""Bank teller" "ATM or bank kiosk" "Telephone banking" "Online banking" "Mobile banking" "Other""'
local acc_vars    hbnkaccm1v2_b hbnkaccm2v2_b hbnkaccm3v2_b hbnkaccm4v2_b hbnkaccm5v2_b hbnkaccm6v2_b
local row = 6
local ac = 0

foreach v of local acc_vars {
    local ++ac
    local lbl : word `ac' of `method_lbls'
    putexcel A`row' = "`lbl'"
    capture quietly svy, subpop(if !inlist(hbnkaccm1v2, -1, .)): mean `v'
    if _rc == 0 {
        local pct = e(b)[1,1] * 100
        local se  = sqrt(e(V)[1,1]) * 100
        putexcel B`row' = (`pct'), nformat("0.0")
        putexcel C`row' = (`se'),  nformat("0.0")
    }
    else {
        putexcel B`row' = "—"
    }
    local ++row
}
di "  --> Table 2.3 written"


/*============================================================================
  TABLE 3.1 — ONLINE PAYMENT SERVICES AND PREPAID CARDS BY CHARACTERISTICS
  Report p.40: % using OPS / % using prepaid, 2021/2023 + diff
  Universe: All households
============================================================================*/
di _n "===== TABLE 3.1: OPS & Prepaid Cards by Characteristics ====="

fdic_demo_table husenowops_b, replace ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.1a OPS") ///
    title("Table 3.1a. Used Nonbank Online Payment Service, by Household Characteristics, ${YEAR}")

fdic_demo_table husenowpp_b, ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.1b Prepaid") ///
    title("Table 3.1b. Used Prepaid Card, by Household Characteristics, ${YEAR}")


/*============================================================================
  TABLE 3.2 — OPS AND PREPAID CARD COMBINATIONS AMONG UNBANKED
  Report p.43: % in each category (PP only / OPS only / Both / Neither)
  Universe: Unbanked households
  Variable: husenowppops (1=PP only  2=OPS only  3=Both  4=Neither)
============================================================================*/
di _n "===== TABLE 3.2: OPS/Prepaid combinations among unbanked ====="

putexcel set "${data}/table_3_ops_prepaid.xlsx", sheet("Table 3.2") modify
putexcel A1 = "Table 3.2. Use of Prepaid Cards and Online Payment Services Among Unbanked Households, ${YEAR}", bold
putexcel A2 = "Universe: Unbanked households"
putexcel A4 = "Category",      bold
putexcel B4 = "${YEAR} %",    bold
putexcel C4 = "SE",            bold
putexcel D4 = "Unweighted N", bold

local cat_lbls `""Prepaid card only" "Online payment service only" "Both prepaid and OPS" "Neither""'
local row = 5
forvalues c = 1/4 {
    local lbl : word `c' of `cat_lbls'
    putexcel A`row' = "`lbl'"
    capture quietly svy, subpop(if unbanked == 1): mean byte(husenowppops == `c')
    if _rc == 0 {
        local pct = e(b)[1,1] * 100
        local se  = sqrt(e(V)[1,1]) * 100
        putexcel B`row' = (`pct'), nformat("0.0")
        putexcel C`row' = (`se'),  nformat("0.0")
    }
    else {
        putexcel B`row' = "—"
    }
    local ++row
}
quietly count if unbanked == 1 & !missing(husenowppops)
putexcel D5 = (r(N))
di "  --> Table 3.2 written"


/*============================================================================
  TABLE 3.3 — UNBANKED AND CASH-ONLY UNBANKED RATES BY CHARACTERISTICS
  Report p.44: % unbanked / % cash-only unbanked, 2023
  Note: Unbanked column = Table 1.1. This produces the cash-only column.
  Merge both in Excel for the full table.
============================================================================*/
di _n "===== TABLE 3.3: Unbanked and Cash-Only Unbanked ====="

fdic_demo_table cashonly_unbanked, ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.3 Cash-only") ///
    title("Table 3.3. Cash-Only Unbanked Rate by Household Characteristics, ${YEAR}")


/*============================================================================
  TABLE 4.1 — MONEY ORDERS, CHECK CASHING, MONEY TRANSFER BY CHARACTERISTICS
  Report p.47: % using each, 2021/2023 + diff
  Universe: All households
============================================================================*/
di _n "===== TABLE 4.1: Money Orders / Check Cashing / Money Transfer ====="

fdic_demo_table huse12mo_b, replace ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1a Money Order") ///
    title("Table 4.1a. Used Nonbank Money Order, by Household Characteristics, ${YEAR}")

fdic_demo_table huse12cc_b, ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1b Check Cash") ///
    title("Table 4.1b. Used Nonbank Check Cashing, by Household Characteristics, ${YEAR}")

fdic_demo_table huse12mt_b, ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1c Money Transfer") ///
    title("Table 4.1c. Used Nonbank Money Transfer, by Household Characteristics, ${YEAR}")


/*============================================================================
  TABLE 4.2 — INTERNATIONAL REMITTANCES BY BANK OWNERSHIP AND CHARACTERISTICS
  Report p.50: % sending remittance (any method / nonbank only), 2023
============================================================================*/
di _n "===== TABLE 4.2: International Remittances ====="

* Full demographic breakdown: any remittance
fdic_demo_table huse12rmany_b, ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.2a Any remit") ///
    title("Table 4.2a. Sent International Remittance (Any Method), by Household Characteristics, ${YEAR}")

* Crosstab: by banking status × remittance type
putexcel set "${data}/table_4_transaction_afs.xlsx", sheet("Table 4.2b Bank status") modify
putexcel A1 = "Table 4.2b. International Remittance Rate by Banking Status, ${YEAR}", bold
putexcel A3 = "Banking Status",           bold
putexcel B3 = "Any Remittance %",         bold
putexcel C3 = "SE (any)",                  bold
putexcel D3 = "Nonbank Remittance %",     bold
putexcel E3 = "SE (nonbank)",              bold
putexcel F3 = "Unweighted N",             bold

local bstat_lbls `""Unbanked" "Underbanked" "Fully banked""'
local row = 4
forvalues s = 1/3 {
    local lbl : word `s' of `bstat_lbls'
    putexcel A`row' = "`lbl'"
    capture quietly svy, subpop(if hbankstatv6 == `s'): mean huse12rmany_b
    local pct_any = e(b)[1,1] * 100
    local se_any  = sqrt(e(V)[1,1]) * 100
    local n_s     = e(N_sub)
    capture quietly svy, subpop(if hbankstatv6 == `s'): mean huse12rm_b
    local pct_nb = e(b)[1,1] * 100
    local se_nb  = sqrt(e(V)[1,1]) * 100
    putexcel B`row' = (`pct_any'), nformat("0.0")
    putexcel C`row' = (`se_any'),  nformat("0.0")
    putexcel D`row' = (`pct_nb'),  nformat("0.0")
    putexcel E`row' = (`se_nb'),   nformat("0.0")
    putexcel F`row' = (`n_s')
    local ++row
}
di "  --> Table 4.2 written"


/*============================================================================
  TABLE 5.1 — MAINSTREAM CREDIT PRODUCTS BY CHARACTERISTICS
  Report p.54: % holding each product × demographic, 2023
  Universe: All households
============================================================================*/
di _n "===== TABLE 5.1: Mainstream Credit Products ====="

putexcel set "${data}/table_5_credit.xlsx", sheet("Table 5.1") replace
putexcel A1 = "Table 5.1. Mainstream Credit Products by Household Characteristics, ${YEAR}", bold
putexcel A2 = "Universe: All households"
putexcel A4 = "Characteristic",    bold
putexcel B4 = "Any Mainstream %",  bold
putexcel C4 = "Credit Card %",     bold
putexcel D4 = "Store Card %",      bold
putexcel E4 = "Auto Loan %",       bold
putexcel F4 = "Mortgage/HE %",     bold
putexcel G4 = "Student Loan %",    bold

local cred_vars hcred12any_b hcred12cc_b hcred12sc_b hcred12car_b hcred12hmln_b hcred12sl_b
local col_letters B C D E F G

* Macro for writing all credit columns for a given subpop condition
* subpop condition passed as `1' arg
local row = 5

* National
putexcel A`row' = "All households", bold
local ci = 0
foreach v of local cred_vars {
    local ++ci
    local col : word `ci' of `col_letters'
    capture quietly svy: mean `v'
    if _rc == 0 {
        putexcel `col'`row' = (`=e(b)[1,1]*100'), nformat("0.0")
    }
}
local ++row

* Race/ethnicity
local ++row
putexcel A`row' = "Race/Ethnicity", bold
local ++row
local race_lbls `""Black" "Hispanic" "Asian" "AIAN" "NHOPI" "White" "Other/Multiracial""'
forvalues r = 1/7 {
    local lbl : word `r' of `race_lbls'
    putexcel A`row' = "  `lbl'"
    local ci = 0
    foreach v of local cred_vars {
        local ++ci
        local col : word `ci' of `col_letters'
        capture quietly svy, subpop(if praceeth == `r'): mean `v'
        if _rc == 0 & e(N_sub) >= 5 {
            putexcel `col'`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        }
    }
    local ++row
}

* Banking status breakdown
local ++row
putexcel A`row' = "Banking Status", bold
local ++row
local bstat_lbls `""Unbanked" "Underbanked" "Fully banked""'
forvalues s = 1/3 {
    local lbl : word `s' of `bstat_lbls'
    putexcel A`row' = "  `lbl'"
    local ci = 0
    foreach v of local cred_vars {
        local ++ci
        local col : word `ci' of `col_letters'
        capture quietly svy, subpop(if hbankstatv6 == `s'): mean `v'
        if _rc == 0 & e(N_sub) >= 5 {
            putexcel `col'`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        }
    }
    local ++row
}

* Income
local ++row
putexcel A`row' = "Annual Household Income", bold
local ++row
local inc_vals  1 2 3 4 5 99
local inc_lbls  `""Less than $15,000" "$15,000-$30,000" "$30,000-$50,000" "$50,000-$75,000" "$75,000 or more" "Unknown""'
local ic = 0
foreach iv of local inc_vals {
    local ++ic
    local lbl : word `ic' of `inc_lbls'
    putexcel A`row' = "  `lbl'"
    local ci = 0
    foreach v of local cred_vars {
        local ++ci
        local col : word `ci' of `col_letters'
        capture quietly svy, subpop(if hhincome2 == `iv'): mean `v'
        if _rc == 0 & e(N_sub) >= 5 {
            putexcel `col'`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        }
    }
    local ++row
}
di "  --> Table 5.1 written"


/*============================================================================
  TABLE 5.2 — AFS CREDIT PRODUCTS BY BANK OWNERSHIP (TREND TABLE)
  Report p.56: % using each AFS credit product by banking status, 2019/2021/2023
  Universe: All households
============================================================================*/
di _n "===== TABLE 5.2: AFS Credit by Bank Ownership ====="

putexcel set "${data}/table_5_credit.xlsx", sheet("Table 5.2") modify
putexcel A1 = "Table 5.2. Nonbank Credit Products by Banking Status, ${YEAR}", bold
putexcel A3 = "Banking Status",     bold
putexcel B3 = "Payday Loan %",      bold
putexcel C3 = "Pawn Shop %",        bold
putexcel D3 = "Rent-to-Own %",      bold
putexcel E3 = "Auto Title Loan %",  bold
putexcel F3 = "Tax Refund Loan %",  bold
putexcel G3 = "Any AFS Credit %",   bold

local afs_vars   huse12pdl_b huse12pwn_b huse12rto_b huse12atl_b huse12ral_b anyafs_credit
local afs_cols   B C D E F G
local bstat_lbls `""All households" "Unbanked" "Underbanked" "Fully banked""'

local row = 4
forvalues s = 0/3 {
    local lbl : word `=`s'+1' of `bstat_lbls'
    putexcel A`row' = "`lbl'"
    local ci = 0
    foreach v of local afs_vars {
        local ++ci
        local col : word `ci' of `afs_cols'
        if `s' == 0 {
            capture quietly svy: mean `v'
        }
        else {
            capture quietly svy, subpop(if hbankstatv6 == `s'): mean `v'
        }
        if _rc == 0 & e(N_sub) >= 5 {
            putexcel `col'`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        }
    }
    local ++row
}
di "  --> Table 5.2 written"


/*============================================================================
  TABLE 5.3 — ANY AFS CREDIT BY HOUSEHOLD CHARACTERISTICS
  Report p.57: % using any nonbank credit, 2019/2021/2023 + diff
  Universe: All households
============================================================================*/
di _n "===== TABLE 5.3: Any AFS Credit by Characteristics ====="
fdic_demo_table anyafs_credit, ///
    xlfile("${data}/table_5_credit.xlsx") ///
    sheet("Table 5.3 Any AFS credit") ///
    title("Table 5.3. Used Any Nonbank Credit Product, by Household Characteristics, ${YEAR}")


/*============================================================================
  TABLE 5.4 — BUY NOW PAY LATER BY BANK OWNERSHIP AND CHARACTERISTICS
  Report p.59: % using BNPL, 2023 (new in 2023)
  Universe: All households
============================================================================*/
di _n "===== TABLE 5.4: Buy Now Pay Later ====="

fdic_demo_table hcred12bnpl_b, ///
    xlfile("${data}/table_5_credit.xlsx") ///
    sheet("Table 5.4 BNPL") ///
    title("Table 5.4. Used Buy Now, Pay Later, by Household Characteristics, ${YEAR}")

* Supplement: BNPL by banking status
putexcel set "${data}/table_5_credit.xlsx", sheet("Table 5.4 by bankstatus") modify
putexcel A1 = "Table 5.4 (supplement). BNPL Use by Banking Status, ${YEAR}", bold
putexcel A3 = "Banking Status", bold
putexcel B3 = "BNPL %",        bold
putexcel C3 = "SE",            bold
putexcel D3 = "Unweighted N",  bold
local bstat_lbls `""Unbanked" "Underbanked" "Fully banked""'
local row = 4
forvalues s = 1/3 {
    local lbl : word `s' of `bstat_lbls'
    putexcel A`row' = "`lbl'"
    capture quietly svy, subpop(if hbankstatv6 == `s'): mean hcred12bnpl_b
    if _rc == 0 {
        putexcel B`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        putexcel C`row' = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")
        putexcel D`row' = (`=e(N_sub)')
    }
    local ++row
}
di "  --> Table 5.4 written"


/*============================================================================
  TABLE 5.5 — MISSED OR LATE BNPL PAYMENTS BY CHARACTERISTICS
  Report p.59: % missing/late BNPL payment, among BNPL users, 2023
  Universe: Households that used BNPL
============================================================================*/
di _n "===== TABLE 5.5: Missed/Late BNPL Payments ====="
fdic_demo_table hcred12bnpldq_b, ///
    xlfile("${data}/table_5_credit.xlsx") ///
    sheet("Table 5.5 BNPL late") ///
    title("Table 5.5. Missed or Late BNPL Payment, by Household Characteristics, ${YEAR}") ///
    subpop("hcred12bnpl == 1")


/*============================================================================
  TABLE 6.1 — CRYPTOCURRENCY BY BANK OWNERSHIP AND CHARACTERISTICS
  Report p.62: % owning/using crypto, 2023 (new in 2023)
  Universe: All households
============================================================================*/
di _n "===== TABLE 6.1: Cryptocurrency ====="

fdic_demo_table huse12cryp_b, replace ///
    xlfile("${data}/table_6_crypto.xlsx") ///
    sheet("Table 6.1") ///
    title("Table 6.1. Owned or Used Cryptocurrency, by Household Characteristics, ${YEAR}")

* Supplement: crypto by banking status
putexcel set "${data}/table_6_crypto.xlsx", sheet("Table 6.1 by bankstatus") modify
putexcel A1 = "Table 6.1 (supplement). Crypto by Banking Status, ${YEAR}", bold
putexcel A3 = "Banking Status", bold
putexcel B3 = "Crypto %",       bold
putexcel C3 = "SE",             bold
putexcel D3 = "Unweighted N",   bold
local bstat_lbls `""Unbanked" "Underbanked" "Fully banked""'
local row = 4
forvalues s = 1/3 {
    local lbl : word `s' of `bstat_lbls'
    putexcel A`row' = "`lbl'"
    capture quietly svy, subpop(if hbankstatv6 == `s'): mean huse12cryp_b
    if _rc == 0 {
        putexcel B`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        putexcel C`row' = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")
        putexcel D`row' = (`=e(N_sub)')
    }
    local ++row
}
di "  --> Table 6.1 written"


/*============================================================================
  TABLE 7.1 — UNDERBANKED RATES BY HOUSEHOLD CHARACTERISTICS
  Report (Section 7): same format as Table 1.1 but for underbanked
  Universe: All households (underbanked=0 for unbanked HHs in hbankstatv6)
============================================================================*/
di _n "===== TABLE 7.1: Underbanked Rates ====="
fdic_demo_table underbanked, replace ///
    xlfile("${data}/table_7_underbanked_rates.xlsx") ///
    sheet("Table 7.1") ///
    title("Table 7.1. Underbanked Rates by Household Characteristics, ${YEAR}")


/*============================================================================
  SECTION 20 — FIGURE-READY DATA (LONG FORMAT)
  One Excel workbook, one sheet per figure concept.
  Each row: year | group_label | group_value | pct | se
  Append rows from prior years by copying this workbook and pasting below
  the previous year's rows.
============================================================================*/
di _n "===== FIGURE DATA ====="

putexcel set "${data}/figures_data.xlsx", sheet("Fig1 National rates") replace
putexcel A1 = "Year"   , bold
putexcel B1 = "Series" , bold
putexcel C1 = "Pct"    , bold
putexcel D1 = "SE"     , bold

* National unbanked rate
quietly svy: mean unbanked
putexcel A2 = (${YEAR})
putexcel B2 = "Unbanked"
putexcel C2 = (`=e(b)[1,1]*100'), nformat("0.0")
putexcel D2 = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")

* National underbanked rate
quietly svy: mean underbanked
putexcel A3 = (${YEAR})
putexcel B3 = "Underbanked"
putexcel C3 = (`=e(b)[1,1]*100'), nformat("0.0")
putexcel D3 = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")

* Note: to build a multi-year trend, append rows from prior years manually
* or re-run on each year's data and stack the resulting sheets

* ── Fig 2: Unbanked by race ───────────────────────────────────────────────────
putexcel set "${data}/figures_data.xlsx", sheet("Fig2 Unbanked by race") modify
putexcel A1 = "Year" , bold
putexcel B1 = "Race" , bold
putexcel C1 = "Pct"  , bold
putexcel D1 = "SE"   , bold

local race_lbls `""Black" "Hispanic" "Asian" "AIAN" "NHOPI" "White" "Other/Multiracial""'
local row = 2
forvalues r = 1/7 {
    local lbl : word `r' of `race_lbls'
    capture quietly svy, subpop(if praceeth == `r'): mean unbanked
    if _rc == 0 & e(N_sub) >= 5 {
        putexcel A`row' = (${YEAR})
        putexcel B`row' = "`lbl'"
        putexcel C`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        putexcel D`row' = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")
        local ++row
    }
}

* ── Fig 3: Primary access method distribution ────────────────────────────────
putexcel set "${data}/figures_data.xlsx", sheet("Fig3 Primary access method") modify
putexcel A1 = "Year"   , bold
putexcel B1 = "Method" , bold
putexcel C1 = "Pct"    , bold
putexcel D1 = "SE"     , bold

local method_lbls `""Bank teller" "ATM/kiosk" "Telephone" "Online banking" "Mobile banking" "Other""'
local row = 2
forvalues m = 1/6 {
    local lbl : word `m' of `method_lbls'
    capture quietly svy, subpop(if !inlist(hbnkaccm, -1, .)): mean byte(hbnkaccm == `m')
    if _rc == 0 {
        putexcel A`row' = (${YEAR})
        putexcel B`row' = "`lbl'"
        putexcel C`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        putexcel D`row' = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")
        local ++row
    }
}

* ── Fig 4: Mobile banking by age group ───────────────────────────────────────
putexcel set "${data}/figures_data.xlsx", sheet("Fig4 Mobile by age") modify
putexcel A1 = "Year"      , bold
putexcel B1 = "Age_group" , bold
putexcel C1 = "Mobile_pct", bold
putexcel D1 = "SE"        , bold

local age_lbls `""15-24" "25-34" "35-44" "45-54" "55-64" "65+""'
local row = 2
forvalues a = 1/6 {
    local lbl : word `a' of `age_lbls'
    capture quietly svy, subpop(if !inlist(hbnkaccm5v2, -1, .) & pagegrp == `a'): mean hbnkaccm5v2_b
    if _rc == 0 & e(N_sub) >= 5 {
        putexcel A`row' = (${YEAR})
        putexcel B`row' = "`lbl'"
        putexcel C`row' = (`=e(b)[1,1]*100'), nformat("0.0")
        putexcel D`row' = (`=sqrt(e(V)[1,1])*100'), nformat("0.0")
        local ++row
    }
}

* ── Fig 5: OPS and prepaid by income ─────────────────────────────────────────
putexcel set "${data}/figures_data.xlsx", sheet("Fig5 OPS Prepaid by income") modify
putexcel A1 = "Year"    , bold
putexcel B1 = "Income"  , bold
putexcel C1 = "OPS_pct" , bold
putexcel D1 = "SE_OPS"  , bold
putexcel E1 = "PP_pct"  , bold
putexcel F1 = "SE_PP"   , bold

local inc_vals 1 2 3 4 5
local inc_lbls `""Less than $15K" "$15-30K" "$30-50K" "$50-75K" "$75K+""'
local row = 2
local ic = 0
foreach iv of local inc_vals {
    local ++ic
    local lbl : word `ic' of `inc_lbls'
    capture quietly svy, subpop(if hhincome2 == `iv'): mean husenowops_b
    if _rc == 0 {
        local pct_ops = e(b)[1,1] * 100
        local se_ops  = sqrt(e(V)[1,1]) * 100
    }
    else {
        local pct_ops = .
        local se_ops  = .
    }
    capture quietly svy, subpop(if hhincome2 == `iv'): mean husenowpp_b
    if _rc == 0 {
        local pct_pp  = e(b)[1,1] * 100
        local se_pp   = sqrt(e(V)[1,1]) * 100
    }
    else {
        local pct_pp = .
        local se_pp  = .
    }
    putexcel A`row' = (${YEAR})
    putexcel B`row' = "`lbl'"
    putexcel C`row' = (`pct_ops'), nformat("0.0")
    putexcel D`row' = (`se_ops'),  nformat("0.0")
    putexcel E`row' = (`pct_pp'),  nformat("0.0")
    putexcel F`row' = (`se_pp'),   nformat("0.0")
    local ++row
}

di "  --> Figure data written"


/*============================================================================
  SECTION 21 — EXPORT ESTIMATES FOR YEAR-OVER-YEAR DIFFERENCES
  Saves a flat .dta file with point estimates for all major outcomes × groups.
  In a future run (e.g., 2025 data), load this file and compute differences
  with significance tests using the standard z-test:
    z = (pct_new - pct_old) / sqrt(se_new^2 + se_old^2)
  Flag as * if |z| > 1.96, ** if |z| > 2.576.
  The estimates file has one row per (outcome, group_var, group_val).
============================================================================*/
di _n "===== Saving estimates for future year comparisons ====="

local all_outcomes ///
    unbanked underbanked cashonly_unbanked ///
    huse12mo_b huse12cc_b huse12mt_b huse12rmany_b huse12rm_b ///
    husenowops_b husenowpp_b ///
    huse12pdl_b huse12pwn_b huse12ral_b huse12atl_b huse12rto_b anyafs_credit ///
    hcred12bnpl_b hcred12bnpldq_b huse12cryp_b ///
    hcred12cc_b hcred12sc_b hcred12car_b hcred12hmln_b hcred12sl_b hcred12any_b ///
    hbnkaccm1v2_b hbnkaccm5v2_b

* Group vars and their value ranges
local demo_gvars   praceeth  pagegrp  peducgrp  hhincome2  hhtypev2  gtmetsta  hbankstatv6
local demo_maxvals 7         6        4         99         8         3         3

postfile pf_est                   ///
    str40  outcome                ///
    str40  group_var              ///
    double group_val              ///
    double mean_pct               ///
    double se_pct                 ///
    double n_sub                  ///
    using "${data}/estimates_${YEAR}.dta", replace

foreach o of local all_outcomes {
    * National
    capture quietly svy: mean `o'
    if _rc == 0 {
        post pf_est ("`o'") ("national") (0) (`=e(b)[1,1]*100') (`=sqrt(e(V)[1,1])*100') (`=e(N)')
    }

    * Demographic groups
    local gi = 0
    foreach gv of local demo_gvars {
        local ++gi
        local maxv : word `gi' of `demo_maxvals'
        * For hhincome2: values 1-5 + 99
        if "`gv'" == "hhincome2" {
            local vals "1 2 3 4 5 99"
        }
        else {
            local vals ""
            forvalues v = 1/`maxv' {
                local vals `vals' `v'
            }
        }
        foreach val of local vals {
            capture quietly svy, subpop(if `gv' == `val'): mean `o'
            if _rc == 0 & e(N_sub) >= 5 {
                post pf_est ("`o'") ("`gv'") (`val') ///
                    (`=e(b)[1,1]*100') (`=sqrt(e(V)[1,1])*100') (`=e(N_sub)')
            }
        }
    }
}
postclose pf_est

di "  Estimates saved to: ${data}/estimates_${YEAR}.dta"
di "  Load this file in a future run to populate the Diff and Sig columns."


di _n(2) "============================================="
di        " All tables written to: ${data}/"
di        "============================================="
di        " Files produced:"
di        "   table_1_1_unbanked_rates.xlsx    (Table 1.1)"
di        "   table_2_bank_access.xlsx         (Tables 2.1, 2.2, 2.3)"
di        "   table_3_ops_prepaid.xlsx         (Tables 3.1, 3.2, 3.3)"
di        "   table_4_transaction_afs.xlsx     (Tables 4.1, 4.2)"
di        "   table_5_credit.xlsx              (Tables 5.1-5.5)"
di        "   table_6_crypto.xlsx              (Table 6.1)"
di        "   table_7_underbanked_rates.xlsx   (Table 7.1)"
di        "   figures_data.xlsx               (Figure-ready data)"
di        "   estimates_${YEAR}.dta            (For future year diffs)"
di        "============================================="
