#=============================================================================
#  03a_fdic_templates_r.R
#
#  R translation of 03a_fdic_templates_putexcel.do
#
#  PURPOSE: Four self-contained, beginner-friendly table templates that
#           demonstrate every common table layout in the FDIC report.
#           Copy one template, change the CONFIGURE block, and you have
#           a new table.
#
#  TABLE FORMAT TYPES
#  ------------------
#  A  Binary outcome x years as columns x demographics as rows
#       Report examples: Table 1.1, 2.2, 3.1, 4.1
#
#  B  Multiple binary outcomes as columns x demographics as rows, one year
#       Report examples: Table 3.3, 5.1
#
#  C  Categorical distribution: rows = categories, one column of percentages
#       Report examples: Table 1.2, 3.2
#
#  D  Rows = survey years, columns = ordered response categories
#       Report examples: Table 1.4, 1.3, 2.1, 2.3
#
#  HOW THESE TEMPLATES WORK
#  -------------------------
#  Each template follows the same three-step pattern:
#
#    Step 1. CONFIGURE: Set the outcome(s), year(s), universe, and output
#            file name. These are the only lines you normally need to change.
#
#    Step 2. ESTIMATE: A loop runs svymean() for each demographic subgroup
#            and stores the result (as a percentage) in a data frame.
#
#    Step 3. EXPORT: One call writes the data frame to Excel using openxlsx.
#
#  PREREQUISITES
#  -------------
#    * R 4.0 or later
#    * Packages: haven, survey, openxlsx
#    * Run the SETUP section once per R session before any template
#    * Output goes to output/ as .xlsx files
#
#  INPUTS:
#    data/hhmultiyear_analys.dta
#    data/hhmultiyear/hhrep19_23.csv
#
#  OUTPUTS:
#    output/table_A_example.xlsx
#    output/table_B_example.xlsx
#    output/table_C_example.xlsx
#    output/table_D_example.xlsx
#=============================================================================


#=============================================================================
# SETUP — Run this block once per R session
#=============================================================================

library(haven)      # reads Stata .dta files
library(survey)     # survey-weighted estimation
library(openxlsx)   # writes Excel workbooks

DATA_DIR   <- "data"
OUTPUT_DIR <- "output"

# ── Load data ───────────────────────────────────────────────────────────────
# We use a single year for most templates. Change YEAR as needed.
YEAR <- 2023

cat("Loading data...\n")
df_all <- read_dta(file.path(DATA_DIR, "hhmultiyear_analys.dta"))

# ── Load BRR replicate weights ──────────────────────────────────────────────
# These 160 replicate weights are needed for proper standard errors.
# For the templates, we merge them in so the survey design is fully specified,
# even though we do not always report SEs in the output.
repwgts <- read.csv(file.path(DATA_DIR, "hhmultiyear", "hhrep19_23.csv"))
repwgts$X <- NULL          # drop row-index column from R's write.csv
repwgts$occurnum <- NULL   # not needed for merge

# Merge replicate weights onto the main dataset
df_all <- merge(df_all, repwgts, by = c("hryear4", "qstnum"), all.x = TRUE)

# ── Recode 1=Yes / 2=No variables to 1/0 ────────────────────────────────────
# In the raw data, binary variables are coded 1=Yes, 2=No, -1=NIU.
# For svymean(), we need 0/1 where 1=Yes, and NA for everything else.
# The weighted mean of a 0/1 variable gives a proportion.
#
# This helper function does the recoding:
recode_binary <- function(x) {
  out <- rep(NA_real_, length(x))
  out[x == 1] <- 1
  out[x == 2] <- 0
  return(out)
}

# Recode all commonly used binary variables.
# (In Stata: foreach v of varlist ... { gen byte `v'_b = (`v' == 1) if inlist(`v', 1, 2) })
binary_vars <- c(
  "huse12mo", "huse12cc", "huse12mt", "huse12rm", "huse12rmany",
  "husenowops", "husenowpp",
  "huse12pdl", "huse12pwn", "huse12ral", "huse12atl", "huse12rto",
  "hcred12cc", "hcred12sc", "hcred12car", "hcred12hmln", "hcred12sl", "hcred12any",
  "hcred12bnpl", "hcred12bnpldq", "huse12cryp",
  "hbnkaccm1v2", "hbnkaccm2v2", "hbnkaccm3v2",
  "hbnkaccm4v2", "hbnkaccm5v2", "hbnkaccm6v2"
)
# Only recode variables that actually exist in the dataset
for (v in binary_vars) {
  if (v %in% names(df_all)) {
    df_all[[paste0(v, "_b")]] <- recode_binary(df_all[[v]])
  }
}

# Banking status convenience variables (0/1):
df_all$unbanked          <- recode_binary(df_all$hunbnk)
df_all$underbanked       <- ifelse(df_all$hbankstatv6 %in% 1:3,
                                   as.numeric(df_all$hbankstatv6 == 2), NA)
df_all$cashonly_unbanked <- recode_binary(df_all$hunbnkcashonly)

# Any AFS credit (variable name changed across years)
if ("huse12afscv3" %in% names(df_all)) {
  df_all$anyafs_credit <- recode_binary(df_all$huse12afscv3)
} else if ("huse12afsc" %in% names(df_all)) {
  df_all$anyafs_credit <- recode_binary(df_all$huse12afsc)
}

cat("Setup complete.\n\n")


#=============================================================================
# TEMPLATE A — Binary outcome x years as columns x demographics as rows
#
# Report examples: Table 1.1 (unbanked), 2.2 (bank access), 3.1 (OPS/prepaid),
#                  4.1 (transaction AFS)
#
# Output layout:
#   Group                | 2019  | 2021  | 2023  | Diff (2023-2021)
#   All households       |  5.4  |  4.5  |  4.2  |  -0.2
#   Race/Ethnicity       |       |       |       |
#     Black              | 13.8  | 11.3  | 10.6  |  -0.8
#     Hispanic           | ...
#
# For a table with TWO outcomes side by side (e.g., Table 2.2 with teller
# and mobile), run this template twice with different sheet names in the
# same workbook.
#=============================================================================

# ---- CONFIGURE (edit these lines) ------------------------------------------
outcome   <- "unbanked"            # 0/1 outcome variable name
years     <- c(2019, 2021, 2023)   # years shown as columns
diff_from <- 2021                  # "prior" year for the Diff column
diff_to   <- 2023                  # "current" year for the Diff column
subpop_condition <- NULL            # NULL = all HH; or a quoted expression
                                    #   like quote(hunbnk == 2) for banked HH
xlfile    <- file.path(OUTPUT_DIR, "table_A_example.xlsx")
sheet     <- "Table 1.1"
# ---- END CONFIGURE ----------------------------------------------------------

# ── Define demographic rows ──
# Each row is defined by a label (text for column A), a variable name
# (group_var), and a value (group_val). "all" means the national total,
# "header" means a section title with no numbers.
#
# This is equivalent to the parallel-locals block in the Stata version.

demo_rows <- data.frame(
  label     = character(),
  group_var = character(),
  group_val = numeric(),
  stringsAsFactors = FALSE
)

# Helper: add a row to the demo_rows data frame.
add_row <- function(label, gv, gval) {
  data.frame(label = label, group_var = gv, group_val = gval,
             stringsAsFactors = FALSE)
}

demo_rows <- rbind(
  add_row("All households",                    "all",              0),
  add_row("Race/Ethnicity",                    "header",           0),
  add_row("  Black",                           "praceeth",         1),
  add_row("  Hispanic",                        "praceeth",         2),
  add_row("  Asian",                           "praceeth",         3),
  add_row("  AIAN",                            "praceeth",         4),
  add_row("  NHOPI",                           "praceeth",         5),
  add_row("  White",                           "praceeth",         6),
  add_row("  Other/Multiracial",               "praceeth",         7),
  add_row("Age",                               "header",           0),
  add_row("  15-24",                           "pagegrp",          1),
  add_row("  25-34",                           "pagegrp",          2),
  add_row("  35-44",                           "pagegrp",          3),
  add_row("  45-54",                           "pagegrp",          4),
  add_row("  55-64",                           "pagegrp",          5),
  add_row("  65+",                             "pagegrp",          6),
  add_row("Education",                         "header",           0),
  add_row("  Less than HS",                    "peducgrp",         1),
  add_row("  HS diploma/GED",                  "peducgrp",         2),
  add_row("  Some college",                    "peducgrp",         3),
  add_row("  College degree",                  "peducgrp",         4),
  add_row("Annual Household Income",           "header",           0),
  add_row("  Less than $15,000",               "hhincome2",        1),
  add_row("  $15,000-$30,000",                 "hhincome2",        2),
  add_row("  $30,000-$50,000",                 "hhincome2",        3),
  add_row("  $50,000-$75,000",                 "hhincome2",        4),
  add_row("  $75,000 or more",                 "hhincome2",        5),
  add_row("  Unknown",                         "hhincome2",       99),
  add_row("Disability Status (ages 25-64)",    "header",           0),
  add_row("  Disabled",                        "pdisabl_age25to64",1),
  add_row("  Not disabled",                    "pdisabl_age25to64",2),
  add_row("Household Type",                    "header",           0),
  add_row("  Married couple",                  "hhtypev2",         1),
  add_row("  Single mother",                   "hhtypev2",         2),
  add_row("  Other female-HH",                 "hhtypev2",         3),
  add_row("  Single father",                   "hhtypev2",         4),
  add_row("  Other male-HH",                   "hhtypev2",         5),
  add_row("  Female nonfamily",                "hhtypev2",         6),
  add_row("  Male nonfamily",                  "hhtypev2",         7),
  add_row("  Other",                           "hhtypev2",         8),
  add_row("Metro Status",                      "header",           0),
  add_row("  Metropolitan",                    "gtmetsta",         1),
  add_row("  Nonmetropolitan",                 "gtmetsta",         2),
  add_row("  Not identified",                  "gtmetsta",         3)
)

nrows <- nrow(demo_rows)
nyears <- length(years)
ncols <- nyears + 1   # one column per year + one Diff column

# Find which column indices correspond to diff_from and diff_to
col_from <- which(years == diff_from)
col_to   <- which(years == diff_to)
stopifnot(length(col_from) == 1, length(col_to) == 1)

# ── Initialize results matrix ──
# Rows = demographic groups; Cols = one per year + one for diff.
# NA means either a header row or too few obs to estimate.
M <- matrix(NA_real_, nrow = nrows, ncol = ncols)

# ── Estimation loop ──
# For each demographic row and each year:
#   - Subset the data to that year and demographic group
#   - Compute the survey-weighted mean (= proportion for a 0/1 variable)
#   - Store estimate * 100 (percentage) in the matrix
#
# We use a simple weighted mean here (hhsupwgt only, no BRR) because
# Template A is designed to be beginner-friendly. For BRR standard errors
# (needed for significance testing), see 01_fdic_tables_r.R.

cat("Template A: estimating...\n")

for (i in 1:nrows) {
  gv   <- demo_rows$group_var[i]
  gval <- demo_rows$group_val[i]

  # Skip header rows (section titles with no numbers)
  if (gv == "header") next

  for (j in seq_along(years)) {
    yr <- years[j]

    # Build the subset condition
    if (gv == "all") {
      idx <- which(df_all$hryear4 == yr & !is.na(df_all[[outcome]]))
    } else {
      idx <- which(df_all$hryear4 == yr &
                   df_all[[gv]] == gval &
                   !is.na(df_all[[outcome]]))
    }

    # Skip if too few observations (N < 5, a loose threshold for templates)
    if (length(idx) < 5) next

    # Compute survey-weighted mean using hhsupwgt.
    # This is equivalent to: svy, subpop(...): mean outcome
    # weighted.mean(x, w) = sum(x*w) / sum(w)
    est <- weighted.mean(df_all[[outcome]][idx],
                         w = df_all$hhsupwgt[idx],
                         na.rm = TRUE)
    M[i, j] <- est * 100   # convert proportion to percentage
  }

  # Difference column (last column)
  if (!is.na(M[i, col_to]) && !is.na(M[i, col_from])) {
    M[i, ncols] <- M[i, col_to] - M[i, col_from]
  }
}

# ── Export to Excel ──
wb <- createWorkbook()
addWorksheet(wb, sheetName = sheet)

# Column headers in row 1 (B1, C1, ...)
writeData(wb, sheet, x = "Group", startCol = 1, startRow = 1)
for (j in seq_along(years)) {
  writeData(wb, sheet, x = as.character(years[j]),
            startCol = j + 1, startRow = 1)
}
writeData(wb, sheet,
          x = paste0("Diff (", diff_to, "\u2013", diff_from, ")"),
          startCol = nyears + 2, startRow = 1)

# Row labels in column A (rows 2 onward) and numeric values
for (i in 1:nrows) {
  writeData(wb, sheet, x = demo_rows$label[i],
            startCol = 1, startRow = i + 1)
  for (j in 1:ncols) {
    if (!is.na(M[i, j])) {
      writeData(wb, sheet, x = round(M[i, j], 1),
                startCol = j + 1, startRow = i + 1)
    }
  }
}

saveWorkbook(wb, xlfile, overwrite = TRUE)
cat(sprintf("Template A complete. Output: %s [sheet: %s]\n\n", xlfile, sheet))

# TO PRODUCE TABLE 2.2 (two outcomes: teller and mobile app):
# Run Template A twice, changing outcome and sheet each time:
#
#   First run:   outcome <- "hbnkaccm1v2_b"   (teller)
#                sheet   <- "Teller"
#                (also adjust subpop_condition for banked HH)
#
#   Second run:  outcome <- "hbnkaccm5v2_b"   (mobile app)
#                sheet   <- "Mobile"
#                (add sheet to existing workbook instead of creating new one)


#=============================================================================
# TEMPLATE B — Multiple binary outcomes as columns x demographics as rows
#
# Report examples: Table 3.3 (cash-only unbanked, OPS, prepaid),
#                  Table 5.1 (mainstream credit products)
#
# Output layout:
#   Group             | Money order | Check cash | Money transfer | ...
#   All households    |    14.2     |    3.1     |     7.0        | ...
#   Race/Ethnicity    |             |            |                |
#     Black           |    25.4     |    ...
#
# NOTE: This is the same as Template A but with outcomes-as-columns instead
# of years-as-columns. Change the CONFIGURE block to use it for a different
# table.
#=============================================================================

# ---- CONFIGURE (edit these lines) ------------------------------------------
outcomes   <- c("huse12mo_b", "huse12cc_b", "huse12mt_b")
out_labels <- c("Money order", "Check cashing", "Money transfer")
year_B     <- 2023           # single year
subpop_B   <- NULL           # NULL = all HH
xlfile_B   <- file.path(OUTPUT_DIR, "table_B_example.xlsx")
sheet_B    <- "Table B"
# ---- END CONFIGURE ----------------------------------------------------------

nout   <- length(outcomes)

# Reuse the same demo_rows from Template A (same demographic breakdown).
# If you want different rows, edit the demo_rows data frame above.

# ── Initialize results matrix ──
# Rows = demographics, Cols = one per outcome (no diff column in Type B)
MB <- matrix(NA_real_, nrow = nrows, ncol = nout)

# ── Estimation loop ──
cat("Template B: estimating...\n")

for (i in 1:nrows) {
  gv   <- demo_rows$group_var[i]
  gval <- demo_rows$group_val[i]

  if (gv == "header") next

  for (j in seq_along(outcomes)) {
    out_var <- outcomes[j]

    if (gv == "all") {
      idx <- which(df_all$hryear4 == year_B & !is.na(df_all[[out_var]]))
    } else {
      idx <- which(df_all$hryear4 == year_B &
                   df_all[[gv]] == gval &
                   !is.na(df_all[[out_var]]))
    }

    if (length(idx) < 5) next

    est <- weighted.mean(df_all[[out_var]][idx],
                         w = df_all$hhsupwgt[idx],
                         na.rm = TRUE)
    MB[i, j] <- est * 100
  }
}

# ── Export to Excel ──
wb_B <- createWorkbook()
addWorksheet(wb_B, sheetName = sheet_B)

# Column headers
writeData(wb_B, sheet_B, x = "Group", startCol = 1, startRow = 1)
for (j in seq_along(out_labels)) {
  writeData(wb_B, sheet_B, x = out_labels[j],
            startCol = j + 1, startRow = 1)
}

# Row labels and values
for (i in 1:nrows) {
  writeData(wb_B, sheet_B, x = demo_rows$label[i],
            startCol = 1, startRow = i + 1)
  for (j in 1:nout) {
    if (!is.na(MB[i, j])) {
      writeData(wb_B, sheet_B, x = round(MB[i, j], 1),
                startCol = j + 1, startRow = i + 1)
    }
  }
}

saveWorkbook(wb_B, xlfile_B, overwrite = TRUE)
cat(sprintf("Template B complete. Output: %s [sheet: %s]\n\n", xlfile_B, sheet_B))


#=============================================================================
# TEMPLATE C — Categorical distribution: rows = categories, one column
#
# Report examples:
#   Table 1.2  Account-ownership transition status (4 categories, all HH)
#   Table 3.2  OPS/prepaid combination status (4 categories, unbanked HH)
#
# Output layout:
#   Category                     | Percent
#   Longer-term unbanked         |   3.7
#   Recently unbanked            |   0.5
#   Recently banked              |   5.9
#   Longer-term banked           |  89.9
#
# How it works: for a categorical variable with K levels, we compute the
# weighted proportion in each level. Percentages sum to 100.
#=============================================================================

# ---- CONFIGURE (edit these lines) ------------------------------------------
catvar     <- "hbankstatv6"          # categorical variable (integer-coded)
cat_vals   <- c(1, 2, 3)            # category values (in order)
cat_labels <- c("Unbanked", "Underbanked", "Fully banked")
year_C     <- 2023
subpop_C   <- NULL                   # NULL = all HH; or quote(hunbnk == 1)
xlfile_C   <- file.path(OUTPUT_DIR, "table_C_example.xlsx")
sheet_C    <- "Table C"
# ---- END CONFIGURE ----------------------------------------------------------

cat("Template C: estimating...\n")

# Subset to the target year
df_C <- df_all[df_all$hryear4 == year_C, ]

# Apply universe restriction if specified
if (!is.null(subpop_C)) {
  df_C <- df_C[eval(subpop_C, envir = df_C), ]
}

# Keep only rows where the categorical variable is one of the target values.
# This excludes -1 (NIU), 99 (unknown), etc.
df_C <- df_C[df_C[[catvar]] %in% cat_vals, ]

# Compute weighted proportions for each category.
# For each category value k, the proportion is:
#   sum(hhsupwgt where catvar == k) / sum(hhsupwgt for all valid rows)
total_wgt <- sum(df_C$hhsupwgt, na.rm = TRUE)

MC <- numeric(length(cat_vals))
for (k in seq_along(cat_vals)) {
  wgt_k <- sum(df_C$hhsupwgt[df_C[[catvar]] == cat_vals[k]], na.rm = TRUE)
  MC[k] <- (wgt_k / total_wgt) * 100
}

# ── Export to Excel ──
wb_C <- createWorkbook()
addWorksheet(wb_C, sheetName = sheet_C)

writeData(wb_C, sheet_C, x = "Category", startCol = 1, startRow = 1)
writeData(wb_C, sheet_C, x = as.character(year_C), startCol = 2, startRow = 1)

for (k in seq_along(cat_labels)) {
  writeData(wb_C, sheet_C, x = cat_labels[k],
            startCol = 1, startRow = k + 1)
  writeData(wb_C, sheet_C, x = round(MC[k], 1),
            startCol = 2, startRow = k + 1)
}

saveWorkbook(wb_C, xlfile_C, overwrite = TRUE)
cat(sprintf("Template C complete. Output: %s [sheet: %s]\n\n", xlfile_C, sheet_C))

# CAUTION: The proportions are computed in the order of cat_vals.
# Make sure cat_vals and cat_labels are in the same order.
# You can verify by running:
#   table(df_C[[catvar]])
# and checking that the values match your cat_vals list.


#=============================================================================
# TEMPLATE D — Rows = survey years, columns = ordered response categories
#
# Report examples:
#   Table 1.4  Interest in having bank account (4 levels, unbanked HH)
#   Table 1.3  Unbanked prior banking status (2 categories, unbanked HH)
#   Table 2.1  Primary account-access method (6 methods, banked HH)
#
# Output layout:
#   Year | Very interested | Somewhat | Not very | Not at all
#   2019 |      XX.X       |  XX.X    |  XX.X    |   XX.X
#   2021 |      XX.X       |  ...
#   2023 |      ...
#
# How it works: for each year, compute weighted proportions across
# response categories. Each row sums to 100%.
#=============================================================================

# ---- CONFIGURE (edit these lines) ------------------------------------------
# NOTE: "hbankint" is not in the multiyear analysis file. This is a
# demonstration — replace with a variable that exists in your data.
# For example, try catvar_D <- "hbankstatv6" with cat_vals_D <- c(1,2,3)
# and subpop_D <- NULL to see the banking status distribution over time.
catvar_D    <- "hbankint"              # categorical response variable
cat_vals_D  <- c(1, 2, 3, 4)          # category values in column order
cat_labs_D  <- c("Very interested", "Somewhat interested",
                 "Not very interested", "Not at all interested")
years_D     <- c(2019, 2021, 2023)    # years to show as rows
subpop_D    <- quote(hunbnk == 1)     # universe: unbanked HH (NULL = all)
xlfile_D    <- file.path(OUTPUT_DIR, "table_D_example.xlsx")
sheet_D     <- "Table 1.4"
# ---- END CONFIGURE ----------------------------------------------------------

ncat_D   <- length(cat_vals_D)
nyears_D <- length(years_D)

cat("Template D: estimating...\n")

# Initialize results matrix (rows = years, cols = categories)
MD <- matrix(NA_real_, nrow = nyears_D, ncol = ncat_D)

for (row_idx in seq_along(years_D)) {
  yr <- years_D[row_idx]

  # Subset to this year
  df_D <- df_all[df_all$hryear4 == yr, ]

  # Apply universe restriction if specified
  if (!is.null(subpop_D)) {
    df_D <- df_D[eval(subpop_D, envir = df_D), ]
  }

  # Keep only valid category values
  df_D <- df_D[df_D[[catvar_D]] %in% cat_vals_D, ]

  if (nrow(df_D) == 0) next

  total_wgt <- sum(df_D$hhsupwgt, na.rm = TRUE)

  for (k in seq_along(cat_vals_D)) {
    wgt_k <- sum(df_D$hhsupwgt[df_D[[catvar_D]] == cat_vals_D[k]], na.rm = TRUE)
    MD[row_idx, k] <- (wgt_k / total_wgt) * 100
  }
}

# ── Export to Excel ──
wb_D <- createWorkbook()
addWorksheet(wb_D, sheetName = sheet_D)

# Column headers
writeData(wb_D, sheet_D, x = "Year", startCol = 1, startRow = 1)
for (k in seq_along(cat_labs_D)) {
  writeData(wb_D, sheet_D, x = cat_labs_D[k],
            startCol = k + 1, startRow = 1)
}

# Row labels (years) and numeric values
for (row_idx in seq_along(years_D)) {
  writeData(wb_D, sheet_D, x = as.character(years_D[row_idx]),
            startCol = 1, startRow = row_idx + 1)
  for (k in 1:ncat_D) {
    if (!is.na(MD[row_idx, k])) {
      writeData(wb_D, sheet_D, x = round(MD[row_idx, k], 1),
                startCol = k + 1, startRow = row_idx + 1)
    }
  }
}

saveWorkbook(wb_D, xlfile_D, overwrite = TRUE)
cat(sprintf("Template D complete. Output: %s [sheet: %s]\n\n", xlfile_D, sheet_D))

# SAME CAUTION AS TEMPLATE C: make sure cat_vals_D and cat_labs_D are in
# the same order. Verify by checking:
#   table(df_all[[catvar_D]][df_all$hryear4 == 2023])
#
# TABLE 2.3 NOTE: Bank-access methods are "check all that apply" (multiple
# response). Use Template B instead with one year, restricted to banked HH,
# with each access method as a separate outcome column. Row percentages
# will sum to more than 100%.

cat("All templates complete.\n")
