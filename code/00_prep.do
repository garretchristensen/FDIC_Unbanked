set type double

* --- Load main multiyear analysis file ---
*change this line to your directory
cd \\wsl.localhost\Ubuntu-24.04\home\gsc-ubuntu\FDIC_Unbanked\
import delimited using "./data/hhmultiyear/hh_multiyear_analys.csv", clear

* --- Keep only supplement respondents ---
keep if hsupresp == 1

* --- Declare survey design ---
svyset [pw=hhsupwgt]

compress
save "data/hhmultiyear_analys.dta", replace
