#=============================================================================
#  01_fdic_tables_r.R
#
#  R translation of 01_fdic_tables_putexcel.do
#
#  PURPOSE: Reproduce the standard demographic tables from the 2023 FDIC
#           National Survey of Unbanked and Underbanked Households.
#           Each table has the same layout:
#
#             Characteristic | 2019 | 2021 | 2023 | Difference (2023-2021)
#
#           Point estimates use the household supplement weight (hhsupwgt).
#           Standard errors use 160 Census Bureau BRR replicate weights
#           with Fay factor = 0.5, matching the report's variance methodology.
#
#  INPUTS:
#    data/hhmultiyear_analys.dta    — Stata dataset, all survey years 2009-2023
#    data/hhmultiyear/hhrep19_23.csv — 160 BRR replicate weights for 2019-2023
#
#  OUTPUTS:
#    output/table_1_1_unbanked_rates.xlsx
#    output/table_7_1_underbanked_rates.xlsx
#    output/table_2_2_bank_access.xlsx
#    output/table_3_1_ops_prepaid.xlsx
#    output/table_3_3_cashonly.xlsx
#    output/table_4_1_transaction_afs.xlsx
#    output/table_5_3_afs_credit.xlsx
#
#  HOW TO RUN:
#    Set your working directory to the repo root, then:
#      source("code/01_fdic_tables_r.R")
#
#  ADAPTING FOR 2025 DATA:
#    1. Update Y1, Y2, Y3 in Section 0 below.
#    2. Update the replicate weight file path in Section 1.
#    3. Check for any new or renamed variables in Section 2.
#    4. Run the script — one function call per table in Section 5.
#=============================================================================

# ── SECTION 0: SETUP & GLOBALS ─────────────────────────────────────────────

# These three packages are the core dependencies:
#   haven    — reads Stata .dta files (like Stata's "use" command)
#   survey   — handles survey-weighted estimation (like Stata's "svy:")
#   openxlsx — writes Excel files (like Stata's "putexcel")
library(haven)      # read Stata .dta files
library(survey)     # survey-weighted estimation (BRR, Fay, etc.)
library(openxlsx)   # write .xlsx Excel workbooks

# ── Year globals ──
# These define which three survey years appear as columns in every table.
# Change these when you have new data (e.g., Y1=2021, Y2=2023, Y3=2025).
Y1 <- 2019   # first year column
Y2 <- 2021   # reference year for Difference column
Y3 <- 2023   # most recent year

# Minimum unweighted subpopulation N to publish an estimate.
# If the subgroup has fewer than MIN_N observations, we write "NA" instead
# of a number, because the estimate would be unreliable.
MIN_N <- 30

# File paths (relative to repo root)
DATA_DIR   <- "data"
OUTPUT_DIR <- "output"


# ── SECTION 1: LOAD DATA AND DECLARE SURVEY DESIGN ─────────────────────────

cat("Loading main dataset...\n")

# read_dta() from the haven package reads Stata .dta files into R.
# This is like "use data/hhmultiyear_analys.dta, clear" in Stata.
df <- read_dta(file.path(DATA_DIR, "hhmultiyear_analys.dta"))

# Keep only the three analysis years (like "keep if inlist(hryear4, ...)")
df <- df[df$hryear4 %in% c(Y1, Y2, Y3), ]
cat(sprintf("  Main data: %d rows after keeping years %d, %d, %d\n",
            nrow(df), Y1, Y2, Y3))

# Load BRR replicate weights from CSV.
# These are 160 Census Bureau replicate weights used to compute standard
# errors via the Balanced Repeated Replication (BRR) method.
cat("Loading replicate weights...\n")
repwgts <- read.csv(file.path(DATA_DIR, "hhmultiyear", "hhrep19_23.csv"))

# The CSV has an "X" column (row index from R export) — drop it.
repwgts$X <- NULL

# Also drop "occurnum" — not needed for the merge, and could cause issues.
repwgts$occurnum <- NULL

# Keep only the three analysis years.
repwgts <- repwgts[repwgts$hryear4 %in% c(Y1, Y2, Y3), ]

# Merge replicate weights onto the main dataset.
# In Stata: merge 1:1 hryear4 qstnum using `repwgts', keep(master match) nogen
# In R: merge() is the equivalent of Stata's "merge" command.
n_before <- nrow(df)
df <- merge(df, repwgts, by = c("hryear4", "qstnum"), all.x = TRUE)
n_after <- nrow(df)

# Merge diagnostic: verify no rows were lost or duplicated.
cat(sprintf("  Merge: %d rows before, %d rows after (should be equal)\n",
            n_before, n_after))
stopifnot(n_before == n_after)  # halt if merge changed row count


# ── SECTION 2: RECODE OUTCOME VARIABLES ────────────────────────────────────

# In the raw data, most outcome variables are coded 1 = Yes, 2 = No.
# For survey-weighted means, we need 0/1 binary variables where 1 = Yes.
# The weighted mean of a 0/1 variable gives a proportion (= percent / 100).
#
# Any value that is not 1 or 2 (e.g., -1 = "Not in Universe") gets set to NA,
# which excludes it from the denominator of the weighted mean.
# This is equivalent to Stata's: gen byte var_b = (var == 1) if inlist(var, 1, 2)

# Helper function: recode a 1=Yes/2=No variable to 1/0, with NA for anything else.
recode_binary <- function(x) {
  # Start with all NA
  out <- rep(NA_real_, length(x))
  # Set to 1 where original == 1 (Yes), 0 where original == 2 (No)
  out[x == 1] <- 1
  out[x == 2] <- 0
  return(out)
}

# ── Banking status ──
df$unbanked          <- recode_binary(df$hunbnk)
df$underbanked       <- ifelse(df$hbankstatv6 %in% c(1, 2, 3),
                               as.numeric(df$hbankstatv6 == 2), NA)
df$cashonly_unbanked  <- recode_binary(df$hunbnkcashonly)

# ── Transaction AFS products ──
# These are "used in past 12 months" indicators (1=Yes, 2=No).
for (v in c("huse12mo", "huse12cc", "huse12mt", "huse12rm", "huse12rmany")) {
  if (v %in% names(df)) {
    df[[paste0(v, "_b")]] <- recode_binary(df[[v]])
  }
}

# ── Online payment services and prepaid cards (2021+ only) ──
# In 2019 these variables do not exist or are all -1 (NIU).
# The function will produce all NAs for 2019, which is correct — the table
# will show "NA" for the 2019 column.
for (v in c("husenowops", "husenowpp")) {
  if (v %in% names(df)) {
    df[[paste0(v, "_b")]] <- recode_binary(df[[v]])
  }
}

# ── AFS credit products ──
for (v in c("huse12pdl", "huse12pwn", "huse12ral", "huse12atl", "huse12rto")) {
  if (v %in% names(df)) {
    df[[paste0(v, "_b")]] <- recode_binary(df[[v]])
  }
}

# Any AFS credit: variable name changed in 2023 multiyear file.
# Try huse12afscv3 first (2023 name), fall back to huse12afsc.
if ("huse12afscv3" %in% names(df)) {
  df$anyafs_credit <- recode_binary(df$huse12afscv3)
} else if ("huse12afsc" %in% names(df)) {
  df$anyafs_credit <- recode_binary(df$huse12afsc)
}

# ── New 2023 products ──
for (v in c("hcred12bnpl", "huse12cryp")) {
  if (v %in% names(df)) {
    df[[paste0(v, "_b")]] <- recode_binary(df[[v]])
  }
}

# ── Bank account access methods ──
for (v in c("hbnkaccm1v2", "hbnkaccm5v2")) {
  if (v %in% names(df)) {
    df[[paste0(v, "_b")]] <- recode_binary(df[[v]])
  }
}

cat("  Outcome variables recoded.\n")


# ── Declare the survey design ──
# This tells the survey package how to compute standard errors.
# It is analogous to Stata's:
#   svyset [pw=hhsupwgt], vce(brr) brrweight(repwgt1-repwgt160) fay(0.5) mse
#
# Parameters:
#   weights    = probability weight (hhsupwgt), used for point estimates
#   repweights = pattern matching repwgt1 through repwgt160 (BRR replicates)
#   type       = "Fay" (a variant of BRR that shrinks replicate weights
#                toward the full-sample weight; reduces variance of the
#                variance estimator)
#   rho        = 0.5 (the Fay factor — Census Bureau standard for CPS)
#   mse        = TRUE (compute variance as mean squared error, not centered)
#   combined.weights = TRUE (replicate weights are full probability weights
#                in their own right; each one can be used as a standalone
#                weight to re-estimate the statistic. This matches Stata's
#                default BRR behavior where replicate weights are used
#                directly to compute replicate estimates.)

cat("Declaring survey design (BRR with Fay factor = 0.5)...\n")
des <- svrepdesign(
  data             = df,
  weights          = ~hhsupwgt,
  repweights       = "repwgt[0-9]+",
  type             = "Fay",
  rho              = 0.5,
  mse              = TRUE,
  combined.weights = TRUE
)
cat("  Survey design declared.\n")


# ── SECTION 3: ESTIMATION HELPER FUNCTION ──────────────────────────────────

# est_cell() estimates the weighted proportion for a single cell:
#   one outcome variable, one year, one demographic subgroup.
#
# Returns a named list with:
#   est = point estimate as a percentage (e.g., 5.4 means 5.4%)
#   se  = standard error as a percentage
#   n   = unweighted subpopulation count
#
# If the subgroup has fewer than MIN_N observations or estimation fails,
# returns list(est = NA, se = NA, n = 0).
#
# Arguments:
#   des       = the full survey design object (created above)
#   outcome   = name of the 0/1 outcome variable (as a string)
#   year      = survey year (e.g., 2023)
#   group_var = name of the demographic variable (as a string),
#               or "national" for the overall estimate
#   group_val = value of group_var to select (ignored if group_var == "national")

est_cell <- function(des, outcome, year, group_var = "national", group_val = 0) {

  # Build the subpopulation indicator: a 0/1 variable that is 1 for
  # observations in the target year AND demographic group.
  # In Stata terms: svy, subpop(if hryear4 == year & group_var == group_val): mean outcome
  if (group_var == "national") {
    subpop_indicator <- as.numeric(des$variables$hryear4 == year)
  } else {
    subpop_indicator <- as.numeric(
      des$variables$hryear4 == year &
      des$variables[[group_var]] == group_val
    )
  }

  # Add the subpopulation indicator to the design object's data.
  des$variables$.subpop <- subpop_indicator

  # Count how many observations are in this subpopulation AND have
  # a non-missing value of the outcome variable.
  n_sub <- sum(subpop_indicator == 1 & !is.na(des$variables[[outcome]]),
               na.rm = TRUE)

  # If too few observations, return NA.
  if (n_sub < MIN_N) {
    return(list(est = NA_real_, se = NA_real_, n = n_sub))
  }

  # Build the formula: ~outcome_variable
  fmla <- as.formula(paste0("~", outcome))

  # Run the survey-weighted mean within the subpopulation.
  # svymean() with na.rm = TRUE excludes missing values from the outcome.
  # The subset argument restricts to our subpopulation.
  result <- tryCatch({
    svymean(fmla, design = subset(des, .subpop == 1), na.rm = TRUE)
  }, error = function(e) {
    return(NULL)
  })

  if (is.null(result)) {
    return(list(est = NA_real_, se = NA_real_, n = n_sub))
  }

  # svymean returns a proportion (0-1). Multiply by 100 for percentage.
  est_pct <- as.numeric(coef(result)) * 100
  se_pct  <- as.numeric(SE(result)) * 100

  return(list(est = est_pct, se = se_pct, n = n_sub))
}


# ── SECTION 4: STANDARD TABLE WRITER ──────────────────────────────────────

# write_std_table() produces one complete table matching the standard layout
# of Table 1.1 (and all other standard demographic tables in the report).
#
# It loops over all demographic rows and all three years, calls est_cell()
# for each cell, computes the Difference column with significance star,
# and writes everything to an Excel worksheet.
#
# Arguments:
#   outcome_var = name of the 0/1 outcome variable (string)
#   wb          = an openxlsx workbook object (created with createWorkbook())
#   sheet_name  = name for the Excel worksheet
#   title       = table title text for row 1

write_std_table <- function(outcome_var, wb, sheet_name, title) {

  cat(sprintf("  Writing %s...\n", sheet_name))

  # Create the worksheet
  addWorksheet(wb, sheetName = sheet_name)

  # ── Define the row structure ──
  # Each row is defined by:
  #   label    = text for column A
  #   gv       = demographic variable name ("national" for all, "header" for section title)
  #   gval     = value to filter on (ignored for "national" and "header")
  #   skip_y2  = TRUE if this variable was not collected in Y2 (2021)
  #
  # This matches the exact row order in the Stata script:
  #   Row 5: All
  #   Rows 6-11: Family Income (hhincome2 = 1..5)
  #   Rows 12-16: Education (peducgrp = 1..4)
  #   Rows 17-23: Age Group (pagegrp = 1..6)
  #   Rows 24-31: Race/Ethnicity (praceeth = 1..7)
  #   Rows 32-34: Disability Status (pdisabl_age25to64 = 1, 2)
  #   Rows 35-38: Monthly Income Volatility (hincvol = 1..3, skip Y2)

  rows <- list(
    # Row 5: All households
    list(label = "All",                                        gv = "national",          gval = 0, skip_y2 = FALSE, bold = TRUE),

    # Row 6: Family Income (section header)
    list(label = "Family Income",                              gv = "header",            gval = 0, skip_y2 = FALSE, bold = TRUE),
    # Rows 7-11: income subgroups
    list(label = "  Less Than $15,000",                        gv = "hhincome2",         gval = 1, skip_y2 = FALSE, bold = FALSE),
    list(label = "  $15,000 to $30,000",                       gv = "hhincome2",         gval = 2, skip_y2 = FALSE, bold = FALSE),
    list(label = "  $30,000 to $50,000",                       gv = "hhincome2",         gval = 3, skip_y2 = FALSE, bold = FALSE),
    list(label = "  $50,000 to $75,000",                       gv = "hhincome2",         gval = 4, skip_y2 = FALSE, bold = FALSE),
    list(label = "  At Least $75,000",                         gv = "hhincome2",         gval = 5, skip_y2 = FALSE, bold = FALSE),

    # Row 12: Education (section header)
    list(label = "Education",                                  gv = "header",            gval = 0, skip_y2 = FALSE, bold = TRUE),
    # Rows 13-16
    list(label = "  No High School Diploma",                   gv = "peducgrp",          gval = 1, skip_y2 = FALSE, bold = FALSE),
    list(label = "  High School Diploma",                      gv = "peducgrp",          gval = 2, skip_y2 = FALSE, bold = FALSE),
    list(label = "  Some College",                             gv = "peducgrp",          gval = 3, skip_y2 = FALSE, bold = FALSE),
    list(label = "  College Degree",                           gv = "peducgrp",          gval = 4, skip_y2 = FALSE, bold = FALSE),

    # Row 17: Age Group (section header)
    list(label = "Age Group",                                  gv = "header",            gval = 0, skip_y2 = FALSE, bold = TRUE),
    # Rows 18-23
    list(label = "  15 to 24 Years",                           gv = "pagegrp",           gval = 1, skip_y2 = FALSE, bold = FALSE),
    list(label = "  25 to 34 Years",                           gv = "pagegrp",           gval = 2, skip_y2 = FALSE, bold = FALSE),
    list(label = "  35 to 44 Years",                           gv = "pagegrp",           gval = 3, skip_y2 = FALSE, bold = FALSE),
    list(label = "  45 to 54 Years",                           gv = "pagegrp",           gval = 4, skip_y2 = FALSE, bold = FALSE),
    list(label = "  55 to 64 Years",                           gv = "pagegrp",           gval = 5, skip_y2 = FALSE, bold = FALSE),
    list(label = "  65 Years or More",                         gv = "pagegrp",           gval = 6, skip_y2 = FALSE, bold = FALSE),

    # Row 24: Race/Ethnicity (section header)
    list(label = "Race/Ethnicity",                             gv = "header",            gval = 0, skip_y2 = FALSE, bold = TRUE),
    # Rows 25-31
    list(label = "  Black",                                    gv = "praceeth",          gval = 1, skip_y2 = FALSE, bold = FALSE),
    list(label = "  Hispanic",                                 gv = "praceeth",          gval = 2, skip_y2 = FALSE, bold = FALSE),
    list(label = "  Asian",                                    gv = "praceeth",          gval = 3, skip_y2 = FALSE, bold = FALSE),
    list(label = "  American Indian or Alaska Native",         gv = "praceeth",          gval = 4, skip_y2 = FALSE, bold = FALSE),
    list(label = "  Native Hawaiian or Other Pacific Islander",gv = "praceeth",          gval = 5, skip_y2 = FALSE, bold = FALSE),
    list(label = "  White",                                    gv = "praceeth",          gval = 6, skip_y2 = FALSE, bold = FALSE),
    list(label = "  Two or More Races",                        gv = "praceeth",          gval = 7, skip_y2 = FALSE, bold = FALSE),

    # Row 32: Disability Status (section header)
    list(label = "Disability Status",                          gv = "header",            gval = 0, skip_y2 = FALSE, bold = TRUE),
    # Rows 33-34
    list(label = "  Disabled, Aged 25 to 64",                  gv = "pdisabl_age25to64", gval = 1, skip_y2 = FALSE, bold = FALSE),
    list(label = "  Not Disabled, Aged 25 to 64",              gv = "pdisabl_age25to64", gval = 2, skip_y2 = FALSE, bold = FALSE),

    # Row 35: Monthly Income Volatility (section header)
    list(label = "Monthly Income Volatility",                  gv = "header",            gval = 0, skip_y2 = FALSE, bold = TRUE),
    # Rows 36-38: skip_y2 = TRUE because hincvol was not collected in 2021
    list(label = "  Income Was About the Same Each Month",     gv = "hincvol",           gval = 1, skip_y2 = TRUE, bold = FALSE),
    list(label = "  Income Varied Somewhat From Month to Month", gv = "hincvol",         gval = 2, skip_y2 = TRUE, bold = FALSE),
    list(label = "  Income Varied a Lot From Month to Month",  gv = "hincvol",           gval = 3, skip_y2 = TRUE, bold = FALSE)
  )

  # ── Write headers ──
  # Row 1: table title (bold)
  bold_style <- createStyle(textDecoration = "bold")
  # Number format style: forces one decimal place display in Excel
  num_style <- createStyle(numFmt = "0.0")
  writeData(wb, sheet_name, x = title, startCol = 1, startRow = 1)
  addStyle(wb, sheet_name, style = bold_style, rows = 1, cols = 1)

  # Row 2: subtitle
  writeData(wb, sheet_name, x = "All Households, Row Percent", startCol = 1, startRow = 2)

  # Row 4: column headers
  col_headers <- c("Characteristic", as.character(Y1), as.character(Y2),
                    as.character(Y3), paste0("Difference (", Y3, "\u2013", Y2, ")"))
  for (j in seq_along(col_headers)) {
    writeData(wb, sheet_name, x = col_headers[j], startCol = j, startRow = 4)
    addStyle(wb, sheet_name, style = bold_style, rows = 4, cols = j)
  }

  # ── Estimate and write each row ──
  # Excel rows start at 5 (row 1 = title, row 2 = subtitle, row 3 = blank,
  # row 4 = headers, row 5 = first data row).
  excel_row <- 5

  for (r in seq_along(rows)) {
    row_def <- rows[[r]]

    # Write the row label in column A
    writeData(wb, sheet_name, x = row_def$label, startCol = 1, startRow = excel_row)
    if (row_def$bold) {
      addStyle(wb, sheet_name, style = bold_style, rows = excel_row, cols = 1)
    }

    # If this is a section header, skip to next row (no numbers to estimate)
    if (row_def$gv == "header") {
      excel_row <- excel_row + 1
      next
    }

    # Estimate for each year: Y1, Y2, Y3
    years <- c(Y1, Y2, Y3)
    est_list <- vector("list", 3)   # will hold list(est, se, n) for each year

    for (yidx in 1:3) {
      # If skip_y2 is TRUE and this is the Y2 column, leave blank
      if (yidx == 2 && row_def$skip_y2) {
        est_list[[yidx]] <- list(est = NA_real_, se = NA_real_, n = 0)
        # Don't write anything — leave cell blank (intentionally not collected)
      } else {
        est_list[[yidx]] <- est_cell(des, outcome_var, years[yidx],
                                      row_def$gv, row_def$gval)
        # Write the point estimate in columns B, C, D (= cols 2, 3, 4)
        if (!is.na(est_list[[yidx]]$est)) {
          writeData(wb, sheet_name,
                    x = round(est_list[[yidx]]$est, 1),
                    startCol = yidx + 1, startRow = excel_row)
        } else {
          # Write "NA" for cells with too few observations
          # But only if skip_y2 is FALSE (we don't write "NA" for intentionally blank cells)
          if (!(yidx == 2 && row_def$skip_y2)) {
            writeData(wb, sheet_name, x = "NA",
                      startCol = yidx + 1, startRow = excel_row)
          }
        }
      }
    }

    # ── Difference column (col E = col 5): Y3 - Y2 ──
    if (row_def$skip_y2) {
      # Not applicable — leave blank (question not asked in Y2)
    } else if (!is.na(est_list[[3]]$est) && !is.na(est_list[[2]]$est)) {
      diff_val <- est_list[[3]]$est - est_list[[2]]$est
      se_diff  <- sqrt(est_list[[3]]$se^2 + est_list[[2]]$se^2)

      # Format: one decimal place, with * if |z| > 1.645 (p < .10 two-tailed)
      diff_str <- sprintf("%.1f", diff_val)
      if (se_diff > 0 && abs(diff_val) / se_diff > 1.645) {
        diff_str <- paste0(diff_str, "*")
      }
      writeData(wb, sheet_name, x = diff_str, startCol = 5, startRow = excel_row)
    } else {
      # At least one year unavailable — cannot compute difference
      writeData(wb, sheet_name, x = "NA", startCol = 5, startRow = excel_row)
    }

    excel_row <- excel_row + 1
  }

  # ── Apply number format to all numeric data cells ──
  # This ensures Excel displays exactly one decimal place (e.g., "16.2" not
  # "16.1974..."), matching the Stata putexcel nformat("0.0") behavior.
  addStyle(wb, sheet_name, style = num_style,
           rows = 5:(excel_row - 1), cols = 2:4,
           gridExpand = TRUE, stack = TRUE)

  # ── Row 39: Note ──
  note_text <- paste0(
    "Note: Monthly income volatility not available for ", Y2,
    ". * = statistically significant at the 10 percent level.",
    " NA = sample too small to produce a reliable estimate.",
    " Standard errors from 160 Census Bureau BRR replicate weights (Fay factor = 0.5)."
  )
  writeData(wb, sheet_name, x = note_text, startCol = 1, startRow = excel_row)

  cat(sprintf("    --> %s done.\n", sheet_name))
}


# ── SECTION 5: WRITE ALL TABLES ───────────────────────────────────────────

cat("\nWriting tables...\n")

# Each Excel file gets its own workbook object, created with createWorkbook().
# Multiple sheets can be added to the same workbook before saving.
# This is like Stata's putexcel set "file.xlsx", sheet("...") replace/modify.

# ── Table 1.1: Unbanked rates ──
wb1 <- createWorkbook()
write_std_table("unbanked", wb1, "Table 1.1",
  paste0("TABLE 1.1 Unbanked Rates by Selected Household Characteristics, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb1, file.path(OUTPUT_DIR, "table_1_1_unbanked_rates.xlsx"),
             overwrite = TRUE)

# ── Table 7.1: Underbanked rates ──
wb7 <- createWorkbook()
write_std_table("underbanked", wb7, "Table 7.1",
  paste0("TABLE 7.1 Underbanked Rates by Selected Household Characteristics, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb7, file.path(OUTPUT_DIR, "table_7_1_underbanked_rates.xlsx"),
             overwrite = TRUE)

# ── Table 2.2: Bank account access methods (two sheets in one workbook) ──
wb2 <- createWorkbook()
write_std_table("hbnkaccm1v2_b", wb2, "Table 2.2a Teller",
  paste0("TABLE 2.2a Bank Teller as Primary Method of Bank Account Access, ",
         Y1, "\u2013", Y3))
write_std_table("hbnkaccm5v2_b", wb2, "Table 2.2b Mobile",
  paste0("TABLE 2.2b Mobile Banking as Primary Method of Bank Account Access, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb2, file.path(OUTPUT_DIR, "table_2_2_bank_access.xlsx"),
             overwrite = TRUE)

# ── Table 3.1: Nonbank online payment services and prepaid cards ──
# Note: husenowops and husenowpp first collected in 2021; Y1 column shows NA
wb3 <- createWorkbook()
write_std_table("husenowops_b", wb3, "Table 3.1a OPS",
  paste0("TABLE 3.1a Use of Nonbank Online Payment Services, ",
         Y1, "\u2013", Y3))
write_std_table("husenowpp_b", wb3, "Table 3.1b Prepaid",
  paste0("TABLE 3.1b Use of Prepaid Cards, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb3, file.path(OUTPUT_DIR, "table_3_1_ops_prepaid.xlsx"),
             overwrite = TRUE)

# ── Table 3.3: Cash-only unbanked ──
wb33 <- createWorkbook()
write_std_table("cashonly_unbanked", wb33, "Table 3.3",
  paste0("TABLE 3.3 Cash-Only Unbanked Rates by Selected Household Characteristics, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb33, file.path(OUTPUT_DIR, "table_3_3_cashonly.xlsx"),
             overwrite = TRUE)

# ── Table 4.1: Nonbank transaction services ──
wb4 <- createWorkbook()
write_std_table("huse12mo_b", wb4, "Table 4.1a Money Order",
  paste0("TABLE 4.1a Use of Nonbank Money Orders, ",
         Y1, "\u2013", Y3))
write_std_table("huse12cc_b", wb4, "Table 4.1b Check Cash",
  paste0("TABLE 4.1b Use of Nonbank Check Cashing, ",
         Y1, "\u2013", Y3))
write_std_table("huse12mt_b", wb4, "Table 4.1c Money Transfer",
  paste0("TABLE 4.1c Use of Nonbank Money Transfer Services, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb4, file.path(OUTPUT_DIR, "table_4_1_transaction_afs.xlsx"),
             overwrite = TRUE)

# ── Table 5.3: Any AFS credit product ──
wb5 <- createWorkbook()
write_std_table("anyafs_credit", wb5, "Table 5.3 Any AFS credit",
  paste0("TABLE 5.3 Use of Any Alternative Financial Credit Products, ",
         Y1, "\u2013", Y3))
saveWorkbook(wb5, file.path(OUTPUT_DIR, "table_5_3_afs_credit.xlsx"),
             overwrite = TRUE)

cat("\nDone. All tables written to ", OUTPUT_DIR, "/\n")
cat(sprintf("Layout: Characteristic | %d | %d | %d | Difference (%d\u2013%d)\n",
            Y1, Y2, Y3, Y3, Y2))
cat(sprintf("  * p<.10 (two-tailed)   NA = N < %d   %d blank for Monthly Income Volatility\n",
            MIN_N, Y2))
