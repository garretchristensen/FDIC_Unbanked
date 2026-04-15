/*============================================================================
  FDIC Unbanked Survey — Add Year-over-Year Differences to Excel Tables

  Run AFTER fdic_unbanked_tables.do for the NEW year's data.
  Requires:
    - estimates_PREV_YEAR.dta  from running fdic_unbanked_tables.do on prior data
    - estimates_CUR_YEAR.dta   from the current run
    - The .xlsx table files produced by the current run

  What this script does:
    1. Merges current and prior year estimates by (outcome, group_var, group_val)
    2. Computes difference and z-test for significance
    3. Opens each Excel file and fills the "Diff" (col F) and "Sig" (col G) columns
       using the row labels in col A to match rows

  Significance convention (two-tailed):
    |z| > 1.96  → *    (p < .05)
    |z| > 2.576 → **   (p < .01)

  NOTE: Because the Excel files use merged / unlabeled rows in column A,
  this script works by reading row labels from Excel and matching them to
  the estimates data. The matching is done via the label-to-(group_var, group_val)
  mapping defined below. Adjust if labels change across years.
============================================================================*/

clear all
set more off

global CUR_YEAR  2023    // year whose tables you want to update
global PREV_YEAR 2021    // year to compare against

global data "../data"

/*----------------------------------------------------------------------------
  STEP 1 — Merge estimates
----------------------------------------------------------------------------*/
use "${data}/estimates_${CUR_YEAR}.dta",  clear
rename mean_pct mean_pct_cur
rename se_pct   se_pct_cur
rename n_sub    n_sub_cur

merge 1:1 outcome group_var group_val ///
    using "${data}/estimates_${PREV_YEAR}.dta", ///
    keepusing(mean_pct se_pct) nogen

rename mean_pct mean_pct_prev
rename se_pct   se_pct_prev

/*----------------------------------------------------------------------------
  STEP 2 — Compute differences and significance
----------------------------------------------------------------------------*/
gen double diff = mean_pct_cur - mean_pct_prev

* Two-sample z-test (independent samples from independent surveys)
gen double se_diff = sqrt(se_pct_cur^2 + se_pct_prev^2)
gen double z_stat  = diff / se_diff

gen str3 sig = ""
replace sig = "**" if abs(z_stat) > 2.576 & !missing(z_stat)
replace sig = "*"  if abs(z_stat) > 1.96  & abs(z_stat) <= 2.576 & !missing(z_stat)

save "${data}/estimates_diff_${PREV_YEAR}_to_${CUR_YEAR}.dta", replace

/*----------------------------------------------------------------------------
  STEP 3 — Utility program to write diff + sig to one Excel table

  The program reads the table's column A (row labels) and looks up each
  label in a supplied Stata dataset (already loaded/merged).

  For each table we define a lookup: row_label → (group_var, group_val)
  Then we find the matching diff and sig in the merged estimates data.

  putexcel uses col E for diff and col F for sig (matching fdic_unbanked_tables.do
  column layout: A=label B=cur% C=SE D=N E=prior% F=diff G=sig).

  This approach is intentionally simple: we just write values to known row
  positions based on the fixed structure written by fdic_demo_table.
  The row offset structure is:
    Row 5:  All households     (national, group_var=national, group_val=0)
    Row 7:  Race/Ethnicity header  (skip)
    Row 8:  Black              (praceeth=1)
    Row 9:  Hispanic           (praceeth=2)
    ...
    Row 14: Other/Multiracial  (praceeth=7)
    Row 16: Age header (skip)
    Row 17: 15-24              (pagegrp=1)
    ...
    Row 22: 65+                (pagegrp=6)
    Row 24: Education header (skip)
    Row 25: Less than HS       (peducgrp=1)
    ...
    Row 28: College degree     (peducgrp=4)
    Row 30: Income header (skip)
    Row 31: <$15K              (hhincome2=1)
    ...
    Row 36: Unknown            (hhincome2=99)
    Row 38: Disability header (skip)
    Row 39: Disabled           (pdisabl_age25to64=1)
    Row 40: Not disabled       (pdisabl_age25to64=2)
    Row 42: HH type header (skip)
    Row 43: Married couple     (hhtypev2=1)
    ...
    Row 50: Other              (hhtypev2=8)
    Row 52: Metro header (skip)
    Row 53: Metropolitan       (gtmetsta=1)
    Row 54: Nonmetropolitan    (gtmetsta=2)
    Row 55: Not identified     (gtmetsta=3)

  These row numbers correspond to what fdic_demo_table writes starting
  at data row 5 (rows 1-4 are title/headers).
----------------------------------------------------------------------------*/

* Row-to-(group_var, group_val) mapping for fdic_demo_table layout
local row_gvar  national  national  praceeth  praceeth  praceeth  praceeth  praceeth  praceeth  praceeth  pagegrp   pagegrp   pagegrp   pagegrp   pagegrp   pagegrp   peducgrp  peducgrp  peducgrp  peducgrp  hhincome2 hhincome2 hhincome2 hhincome2 hhincome2 hhincome2 pdisabl_age25to64 pdisabl_age25to64 hhtypev2 hhtypev2 hhtypev2 hhtypev2 hhtypev2 hhtypev2 hhtypev2 hhtypev2 gtmetsta  gtmetsta  gtmetsta
local row_gval  0         0         1         2         3         4         5         6         7         1         2         3         4         5         6         1         2         3         4         1         2         3         4         5         99        1                 2                 1        2        3        4        5        6        7        8        1         2         3
local data_rows 5         5         8         9         10        11        12        13        14        17        18        19        20        21        22        25        26        27        28        31        32        33        34        35        36        39                40                43       44       45       46       47       48       49       50       53        54        55

capture program drop write_diffs_to_excel
program define write_diffs_to_excel
    /*
       write_diffs_to_excel using diff_dataset, outcome(varname) xlfile(path) sheet(name)
    */
    syntax using/, outcome(string) xlfile(string) sheet(string)

    use "`using'", clear

    * Build lookup of diff and sig by (group_var, group_val)
    keep if outcome == "`outcome'"
    keep group_var group_val diff sig

    * (outcome-specific estimates loaded; now iterate over rows)
    local row_gvars `c(row_gvar)'     // inherited from caller's locals
    local row_gvals `c(row_gval)'
    local data_rowlist `c(data_rows)'

    * Because we can't pass long local lists through -using-, write a small
    * tempfile with the row-number lookup
    tempfile lookup
    quietly {
        clear
        input str40 gv double gval double xlrow
        "national"           0   5
        "praceeth"           1   8
        "praceeth"           2   9
        "praceeth"           3  10
        "praceeth"           4  11
        "praceeth"           5  12
        "praceeth"           6  13
        "praceeth"           7  14
        "pagegrp"            1  17
        "pagegrp"            2  18
        "pagegrp"            3  19
        "pagegrp"            4  20
        "pagegrp"            5  21
        "pagegrp"            6  22
        "peducgrp"           1  25
        "peducgrp"           2  26
        "peducgrp"           3  27
        "peducgrp"           4  28
        "hhincome2"          1  31
        "hhincome2"          2  32
        "hhincome2"          3  33
        "hhincome2"          4  34
        "hhincome2"          5  35
        "hhincome2"         99  36
        "pdisabl_age25to64"  1  39
        "pdisabl_age25to64"  2  40
        "hhtypev2"           1  43
        "hhtypev2"           2  44
        "hhtypev2"           3  45
        "hhtypev2"           4  46
        "hhtypev2"           5  47
        "hhtypev2"           6  48
        "hhtypev2"           7  49
        "hhtypev2"           8  50
        "gtmetsta"           1  53
        "gtmetsta"           2  54
        "gtmetsta"           3  55
        end
        rename gv group_var
        rename gval group_val
        save `lookup'
    }

    use "`using'", clear
    keep if outcome == "`outcome'"
    keep group_var group_val diff sig
    merge 1:1 group_var group_val using `lookup', keep(match) nogen

    putexcel set "`xlfile'", sheet("`sheet'") modify

    * Write prior year % (col E), diff (col F), sig (col G) for each row
    forvalues i = 1/`=_N' {
        local xl_row = xlrow[`i']
        local prev_pct  = .   // not stored in this file; fill by hand or add to estimates dta
        local d  = diff[`i']
        local s  = sig[`i']

        if !missing(`d') {
            putexcel F`xl_row' = (`d'), nformat("0.0")
        }
        if "`s'" != "" & "`s'" != "." {
            putexcel G`xl_row' = "`s'"
        }
    }
    di "  --> Diffs written to `xlfile' [sheet: `sheet']"
end

/*----------------------------------------------------------------------------
  STEP 4 — Apply diffs to each table
  The differences dataset is already in memory after Step 2.
  We save it to a tempfile and pass it to write_diffs_to_excel.
----------------------------------------------------------------------------*/
use "${data}/estimates_diff_${PREV_YEAR}_to_${CUR_YEAR}.dta", clear
tempfile alldiffs
save `alldiffs'

* Table 1.1
write_diffs_to_excel using `alldiffs', ///
    outcome("unbanked") ///
    xlfile("${data}/table_1_1_unbanked_rates.xlsx") ///
    sheet("Table 1.1")

* Table 2.2 — teller and mobile
write_diffs_to_excel using `alldiffs', ///
    outcome("hbnkaccm1v2_b") ///
    xlfile("${data}/table_2_bank_access.xlsx") ///
    sheet("Table 2.2a Teller")

write_diffs_to_excel using `alldiffs', ///
    outcome("hbnkaccm5v2_b") ///
    xlfile("${data}/table_2_bank_access.xlsx") ///
    sheet("Table 2.2b Mobile")

* Table 3.1 — OPS and prepaid
write_diffs_to_excel using `alldiffs', ///
    outcome("husenowops_b") ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.1a OPS")

write_diffs_to_excel using `alldiffs', ///
    outcome("husenowpp_b") ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.1b Prepaid")

* Table 3.3 — cash-only unbanked
write_diffs_to_excel using `alldiffs', ///
    outcome("cashonly_unbanked") ///
    xlfile("${data}/table_3_ops_prepaid.xlsx") ///
    sheet("Table 3.3 Cash-only")

* Table 4.1 — transaction AFS
write_diffs_to_excel using `alldiffs', ///
    outcome("huse12mo_b") ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1a Money Order")

write_diffs_to_excel using `alldiffs', ///
    outcome("huse12cc_b") ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1b Check Cash")

write_diffs_to_excel using `alldiffs', ///
    outcome("huse12mt_b") ///
    xlfile("${data}/table_4_transaction_afs.xlsx") ///
    sheet("Table 4.1c Money Transfer")

* Table 5.3 — any AFS credit
write_diffs_to_excel using `alldiffs', ///
    outcome("anyafs_credit") ///
    xlfile("${data}/table_5_credit.xlsx") ///
    sheet("Table 5.3 Any AFS credit")

* Table 7.1 — underbanked
write_diffs_to_excel using `alldiffs', ///
    outcome("underbanked") ///
    xlfile("${data}/table_7_underbanked_rates.xlsx") ///
    sheet("Table 7.1")

di _n "All diffs written."
di "Review Diff (col F) and Sig (col G) columns in each Excel file."
di "Stars: * = p<.05, ** = p<.01  (two-tailed z-test, independent surveys)"
