set type double

* --- Load main multiyear analysis file ---
* Stata treats "." in the CSV as missing automatically for numeric columns.
cd \\wsl.localhost\Ubuntu-24.04\home\gsc-ubuntu\FDIC_Unbanked\
import delimited using "./data/hhmultiyear/hh_multiyear_analys.csv", clear

* --- Keep only supplement respondents ---
keep if hsupresp == 1

* --- Declare survey design ---
svyset [pw=hhsupwgt]

save "data/hhmultiyear_analys.dta", replace
