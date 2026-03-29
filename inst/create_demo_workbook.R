# =============================================================================
# Create a demo Excel workbook with complex multi-sheet formulas
# Run this once to generate inst/extdata/sales_report_demo.xlsx
# =============================================================================

library(openxlsx2)

wb <- wb_workbook()

# ============================================================
# Sheet 1: "Products" -- master product list with pricing
# ============================================================
wb$add_worksheet("Products")

# Headers (row 1)
wb$add_data(x = data.frame(
  A = c("Product_ID", "P001", "P002", "P003", "P004", "P005",
        "P006", "P007", "P008", "P009", "P010"),
  B = c("Category", "Electronics", "Electronics", "Furniture", "Furniture", "Clothing",
        "Clothing", "Electronics", "Furniture", "Clothing", "Electronics"),
  C = c("Product_Name", "Laptop", "Phone", "Desk", "Chair", "Jacket",
        "Shirt", "Tablet", "Bookshelf", "Shoes", "Monitor"),
  D = c("Unit_Price", "1200", "800", "350", "275", "150",
        "45", "500", "180", "95", "450"),
  E = c("Cost", "850", "520", "200", "160", "80",
        "22", "310", "100", "50", "280"),
  F = c("Margin_Pct", "", "", "", "", "",
        "", "", "", "", ""),
  G = c("Status", "Active", "Active", "Active", "Discontinued", "Active",
        "Active", "Active", "Active", "Active", "Active")
), sheet = "Products", col_names = FALSE)

# Margin formulas in column F (rows 2-11): =(D-E)/D
for (r in 2:11) {
  wb$add_formula(sheet = "Products", x = sprintf("(D%d-E%d)/D%d", r, r, r),
                  dims = sprintf("F%d", r))
}

# Summary rows
wb$add_data(x = data.frame(A = c("", "TOTALS", "AVG_PRICE", "MAX_PRICE", "MIN_PRICE",
                                   "ACTIVE_COUNT")),
            sheet = "Products", dims = "A12", col_names = FALSE)

# Total cost, total price
wb$add_formula(sheet = "Products", x = "SUM(D2:D11)", dims = "D13")
wb$add_formula(sheet = "Products", x = "SUM(E2:E11)", dims = "E13")
wb$add_formula(sheet = "Products", x = "AVERAGE(F2:F11)", dims = "F13")

# AVG, MAX, MIN
wb$add_formula(sheet = "Products", x = "AVERAGE(D2:D11)", dims = "D14")
wb$add_formula(sheet = "Products", x = "MAX(D2:D11)", dims = "D15")
wb$add_formula(sheet = "Products", x = "MIN(D2:D11)", dims = "D16")

# COUNTIF for active products
wb$add_formula(sheet = "Products", x = 'COUNTIF(G2:G11,"Active")', dims = "D17")


# ============================================================
# Sheet 2: "Q1 Sales" -- quarterly sales data
# ============================================================
wb$add_worksheet("Q1 Sales")

# Headers
wb$add_data(x = data.frame(
  A = c("Date", "2024-01-05", "2024-01-12", "2024-01-20", "2024-02-03",
        "2024-02-14", "2024-02-28", "2024-03-05", "2024-03-15", "2024-03-22",
        "2024-01-08", "2024-01-25", "2024-02-10", "2024-02-20", "2024-03-01",
        "2024-03-10", "2024-03-18", "2024-03-28", "2024-01-15", "2024-02-05"),
  B = c("Product_ID", "P001", "P002", "P003", "P001", "P005",
        "P006", "P001", "P007", "P002", "P004",
        "P008", "P009", "P010", "P003", "P006",
        "P001", "P002", "P005", "P007"),
  C = c("Region", "North", "South", "North", "East", "West",
        "North", "South", "East", "North", "West",
        "North", "South", "East", "West", "North",
        "East", "South", "West", "North"),
  D = c("Qty", "3", "5", "2", "1", "10",
        "20", "4", "3", "6", "1",
        "2", "8", "2", "3", "15",
        "2", "4", "7", "1"),
  E = c("Unit_Price", "1200", "800", "350", "1200", "150",
        "45", "1200", "500", "800", "275",
        "180", "95", "450", "350", "45",
        "1200", "800", "150", "500"),
  F = c("Revenue", "", "", "", "", "",
        "", "", "", "", "",
        "", "", "", "", "",
        "", "", "", ""),
  G = c("Discount_Pct", "0", "0.05", "0", "0.1", "0",
        "0", "0.05", "0.1", "0", "0.15",
        "0", "0", "0.05", "0", "0",
        "0.1", "0.05", "0", "0.1"),
  H = c("Net_Revenue", "", "", "", "", "",
        "", "", "", "", "",
        "", "", "", "", "",
        "", "", "", "")
), sheet = "Q1 Sales", col_names = FALSE)

# Revenue = Qty * Unit_Price
for (r in 2:20) {
  wb$add_formula(sheet = "Q1 Sales", x = sprintf("D%d*E%d", r, r),
                  dims = sprintf("F%d", r))
}

# Net_Revenue = Revenue * (1 - Discount_Pct)
for (r in 2:20) {
  wb$add_formula(sheet = "Q1 Sales", x = sprintf("F%d*(1-G%d)", r, r),
                  dims = sprintf("H%d", r))
}

# Summary rows
wb$add_data(x = data.frame(A = c("", "TOTAL_QTY", "TOTAL_REVENUE",
                                   "TOTAL_NET_REV", "AVG_DISCOUNT",
                                   "NORTH_REVENUE", "SOUTH_REVENUE",
                                   "EAST_REVENUE", "WEST_REVENUE")),
            sheet = "Q1 Sales", dims = "A21", col_names = FALSE)

wb$add_formula(sheet = "Q1 Sales", x = "SUM(D2:D20)", dims = "D22")
wb$add_formula(sheet = "Q1 Sales", x = "SUM(F2:F20)", dims = "F23")
wb$add_formula(sheet = "Q1 Sales", x = "SUM(H2:H20)", dims = "H24")
wb$add_formula(sheet = "Q1 Sales", x = "AVERAGE(G2:G20)", dims = "G25")

# SUMIFS -- revenue by region
wb$add_formula(sheet = "Q1 Sales", x = 'SUMIFS(F2:F20,C2:C20,"North")', dims = "F26")
wb$add_formula(sheet = "Q1 Sales", x = 'SUMIFS(F2:F20,C2:C20,"South")', dims = "F27")
wb$add_formula(sheet = "Q1 Sales", x = 'SUMIFS(F2:F20,C2:C20,"East")', dims = "F28")
wb$add_formula(sheet = "Q1 Sales", x = 'SUMIFS(F2:F20,C2:C20,"West")', dims = "F29")


# ============================================================
# Sheet 3: "Q2 Sales" -- same structure, different data
# ============================================================
wb$add_worksheet("Q2 Sales")

wb$add_data(x = data.frame(
  A = c("Date", "2024-04-02", "2024-04-15", "2024-04-22", "2024-05-01",
        "2024-05-12", "2024-05-25", "2024-06-03", "2024-06-14", "2024-06-20",
        "2024-04-08", "2024-04-20", "2024-05-05", "2024-05-18", "2024-06-01",
        "2024-06-10"),
  B = c("Product_ID", "P001", "P003", "P005", "P002", "P007",
        "P001", "P006", "P010", "P002", "P008",
        "P009", "P001", "P003", "P005", "P004"),
  C = c("Region", "South", "North", "East", "West", "North",
        "East", "South", "North", "East", "West",
        "North", "South", "East", "West", "North"),
  D = c("Qty", "2", "4", "12", "3", "2",
        "5", "25", "3", "7", "1",
        "6", "3", "5", "9", "2"),
  E = c("Unit_Price", "1200", "350", "150", "800", "500",
        "1200", "45", "450", "800", "180",
        "95", "1200", "350", "150", "275"),
  F = c("Revenue", "", "", "", "", "",
        "", "", "", "", "",
        "", "", "", "", ""),
  G = c("Discount_Pct", "0.05", "0", "0", "0.1", "0",
        "0.05", "0", "0.1", "0", "0",
        "0", "0.1", "0", "0.05", "0.15"),
  H = c("Net_Revenue", "", "", "", "", "",
        "", "", "", "", "",
        "", "", "", "", "")
), sheet = "Q2 Sales", col_names = FALSE)

for (r in 2:16) {
  wb$add_formula(sheet = "Q2 Sales", x = sprintf("D%d*E%d", r, r),
                  dims = sprintf("F%d", r))
  wb$add_formula(sheet = "Q2 Sales", x = sprintf("F%d*(1-G%d)", r, r),
                  dims = sprintf("H%d", r))
}

wb$add_data(x = data.frame(A = c("", "TOTAL_QTY", "TOTAL_REVENUE",
                                   "TOTAL_NET_REV", "AVG_DISCOUNT",
                                   "NORTH_REVENUE", "SOUTH_REVENUE",
                                   "EAST_REVENUE", "WEST_REVENUE")),
            sheet = "Q2 Sales", dims = "A17", col_names = FALSE)

wb$add_formula(sheet = "Q2 Sales", x = "SUM(D2:D16)", dims = "D18")
wb$add_formula(sheet = "Q2 Sales", x = "SUM(F2:F16)", dims = "F19")
wb$add_formula(sheet = "Q2 Sales", x = "SUM(H2:H16)", dims = "H20")
wb$add_formula(sheet = "Q2 Sales", x = "AVERAGE(G2:G16)", dims = "G21")

wb$add_formula(sheet = "Q2 Sales", x = 'SUMIFS(F2:F16,C2:C16,"North")', dims = "F22")
wb$add_formula(sheet = "Q2 Sales", x = 'SUMIFS(F2:F16,C2:C16,"South")', dims = "F23")
wb$add_formula(sheet = "Q2 Sales", x = 'SUMIFS(F2:F16,C2:C16,"East")', dims = "F24")
wb$add_formula(sheet = "Q2 Sales", x = 'SUMIFS(F2:F16,C2:C16,"West")', dims = "F25")


# ============================================================
# Sheet 4: "Annual Summary" -- cross-sheet references & complex formulas
# ============================================================
wb$add_worksheet("Annual Summary")

wb$add_data(x = data.frame(
  A = c("Metric", "Total Units Sold", "Total Revenue", "Total Net Revenue",
        "Average Discount", "Revenue Growth Q1->Q2",
        "", "Region", "North", "South", "East", "West", "TOTAL",
        "", "Category Analysis", "Avg Product Margin",
        "Active Products", "Revenue per Active Product"),
  B = c("Q1", "", "", "", "", "",
        "", "Q1 Revenue", "", "", "", "", "",
        "", "", "", "", ""),
  C = c("Q2", "", "", "", "", "",
        "", "Q2 Revenue", "", "", "", "", "",
        "", "", "", "", ""),
  D = c("Total / Change", "", "", "", "", "",
        "", "Total", "", "", "", "", "",
        "", "", "", "", "")
), sheet = "Annual Summary", col_names = FALSE)

# Q1 references
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!D22", dims = "B2")
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!F23", dims = "B3")
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!H24", dims = "B4")
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!G25", dims = "B5")

# Q2 references
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!D18", dims = "C2")
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!F19", dims = "C3")
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!H20", dims = "C4")
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!G21", dims = "C5")

# Totals
wb$add_formula(sheet = "Annual Summary", x = "B2+C2", dims = "D2")
wb$add_formula(sheet = "Annual Summary", x = "B3+C3", dims = "D3")
wb$add_formula(sheet = "Annual Summary", x = "B4+C4", dims = "D4")
wb$add_formula(sheet = "Annual Summary", x = "AVERAGE(B5,C5)", dims = "D5")

# Revenue growth: (Q2-Q1)/Q1 -- uses IFERROR for safety
wb$add_formula(sheet = "Annual Summary", x = "IFERROR((C3-B3)/B3,0)", dims = "D6")

# Region breakdown -- cross-sheet refs
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!F26", dims = "B9")
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!F27", dims = "B10")
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!F28", dims = "B11")
wb$add_formula(sheet = "Annual Summary", x = "'Q1 Sales'!F29", dims = "B12")
wb$add_formula(sheet = "Annual Summary", x = "SUM(B9:B12)", dims = "B13")

wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!F22", dims = "C9")
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!F23", dims = "C10")
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!F24", dims = "C11")
wb$add_formula(sheet = "Annual Summary", x = "'Q2 Sales'!F25", dims = "C12")
wb$add_formula(sheet = "Annual Summary", x = "SUM(C9:C12)", dims = "C13")

# Total per region
wb$add_formula(sheet = "Annual Summary", x = "B9+C9", dims = "D9")
wb$add_formula(sheet = "Annual Summary", x = "B10+C10", dims = "D10")
wb$add_formula(sheet = "Annual Summary", x = "B11+C11", dims = "D11")
wb$add_formula(sheet = "Annual Summary", x = "B12+C12", dims = "D12")
wb$add_formula(sheet = "Annual Summary", x = "SUM(D9:D12)", dims = "D13")

# Category analysis -- cross-sheet to Products
wb$add_formula(sheet = "Annual Summary", x = "'Products'!F13", dims = "B16")
wb$add_formula(sheet = "Annual Summary", x = "'Products'!D17", dims = "B17")

# Revenue per active product: IF active > 0
wb$add_formula(sheet = "Annual Summary", x = "IF(B17>0,D3/B17,0)", dims = "B18")


# ============================================================
# Sheet 5: "Pivot Analysis" -- more SUMIFS, COUNTIFS, nested IF
# ============================================================
wb$add_worksheet("Pivot Analysis")

wb$add_data(x = data.frame(
  A = c("Category", "Electronics", "Furniture", "Clothing", "", "Grand Total",
        "", "Discount Impact", "With Discount", "Without Discount",
        "", "Top Region by Q1 Revenue"),
  B = c("Product Count", "", "", "", "", "",
        "", "", "", "",
        "", ""),
  C = c("Total Q1 Revenue", "", "", "", "", "",
        "", "Q1 Revenue", "", "",
        "", ""),
  D = c("Total Q2 Revenue", "", "", "", "", "",
        "", "Q2 Revenue", "", "",
        "", ""),
  E = c("Combined Revenue", "", "", "", "", "",
        "", "Combined", "", "",
        "", "")
), sheet = "Pivot Analysis", col_names = FALSE)

# COUNTIFS for product count per category
wb$add_formula(sheet = "Pivot Analysis",
                x = 'COUNTIF(Products!B2:B11,"Electronics")', dims = "B2")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'COUNTIF(Products!B2:B11,"Furniture")', dims = "B3")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'COUNTIF(Products!B2:B11,"Clothing")', dims = "B4")
wb$add_formula(sheet = "Pivot Analysis",
                x = "SUM(B2:B4)", dims = "B6")

# SUMIFS for Q1 revenue by category (using product ID prefix matching is complex,
# so we use a simpler approach: sum revenue where product matches)
# For demo, we'll use cross-sheet SUMIFS
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P001")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P002")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P007")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P010")',
                dims = "C2")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P003")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P004")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P008")',
                dims = "C3")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P005")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P006")+SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!B2:B20,"P009")',
                dims = "C4")
wb$add_formula(sheet = "Pivot Analysis", x = "SUM(C2:C4)", dims = "C6")

# Q2 same pattern
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P001")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P002")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P007")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P010")',
                dims = "D2")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P003")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P004")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P008")',
                dims = "D3")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P005")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P006")+SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!B2:B16,"P009")',
                dims = "D4")
wb$add_formula(sheet = "Pivot Analysis", x = "SUM(D2:D4)", dims = "D6")

# Combined
for (r in c(2, 3, 4, 6)) {
  wb$add_formula(sheet = "Pivot Analysis", x = sprintf("C%d+D%d", r, r),
                  dims = sprintf("E%d", r))
}

# Discount impact: SUMIFS with condition on discount > 0
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!G2:G20,">0")',
                dims = "C9")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q1 Sales\'!F2:F20,\'Q1 Sales\'!G2:G20,"0")',
                dims = "C10")

wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!G2:G16,">0")',
                dims = "D9")
wb$add_formula(sheet = "Pivot Analysis",
                x = 'SUMIFS(\'Q2 Sales\'!F2:F16,\'Q2 Sales\'!G2:G16,"0")',
                dims = "D10")

wb$add_formula(sheet = "Pivot Analysis", x = "C9+D9", dims = "E9")
wb$add_formula(sheet = "Pivot Analysis", x = "C10+D10", dims = "E10")

# Nested IF: determine top region
wb$add_formula(sheet = "Pivot Analysis",
                x = 'IF(\'Annual Summary\'!D9>=\'Annual Summary\'!D10,IF(\'Annual Summary\'!D9>=\'Annual Summary\'!D11,IF(\'Annual Summary\'!D9>=\'Annual Summary\'!D12,"North","West"),"East"),IF(\'Annual Summary\'!D10>=\'Annual Summary\'!D11,IF(\'Annual Summary\'!D10>=\'Annual Summary\'!D12,"South","West"),"East"))',
                dims = "B12")


# ============================================================
# Save the workbook
# ============================================================
out_file <- "inst/extdata/sales_report_demo.xlsx"
wb_save(wb, out_file, overwrite = TRUE)
cat("Demo workbook saved to:", out_file, "\n")
cat("Sheets:", paste(wb$get_sheet_names(), collapse = ", "), "\n")
