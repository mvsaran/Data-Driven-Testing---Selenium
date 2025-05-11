# 🧪 Data-Driven Testing Framework using Selenium & Apache POI

This repository showcases a robust **data-driven testing framework** built using Java, Selenium WebDriver, and Apache POI to automate and validate a Fixed Deposit (FD) calculator application.

## 🚀 Features

- ✅ Read input test data from Excel (`.xlsx`) file
- ✅ Write results back into Excel
- ✅ Write to a specific cell in Excel
- ✅ Dynamically write data into specific cells
- ✅ Excel utility class (`ExcelUtils.java`) to centralize all reusable Excel methods
- ✅ End-to-end automation script (`FDCalculator.java`) that:
  - Reads input values (principal, interest, tenure, etc.)
  - Enters them into a browser app using Selenium
  - Captures the output and compares it with expected values
  - Writes test status (Pass/Fail/Error) back to the Excel sheet

## 📁 Tech Stack

- **Java**
- **Selenium WebDriver**
- **Apache POI** (for Excel interaction)
- **TestNG/JUnit** *(Optional for integration)*
- **Excel (.xlsx) as Data Source**
- **Maven** for dependency management

## 🧰 Utilities

- `ExcelUtils.java`: Contains reusable methods like `getCellData`, `setCellData`, `getRowCount`, and cell color formatting.
- `FDCalculator.java`: The main script performing web automation and Excel-driven testing.
