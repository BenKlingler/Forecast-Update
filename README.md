# Forecast-Update
## Overview

This suite of VBA macros is designed to manage and forecast financial data within an Excel workbook, particularly across "UPCOMING PROJECTS" and "Revenue Summary" sheets. The macros streamline the process of updating financial forecasts, distributing project hours, and maintaining the accuracy of the revenue summary.

## Macro Descriptions

### Macro: `RunAllMacros`

**Purpose:** Serves as an orchestrator to run a sequence of macros in a specific order to update the financial forecast and revenue summary.

**Executed Macros:**
- `ExtendForecastValues`
- `UpdateTotal2025`
- `UpdateRemainingContractAmount`

### Macro: `ExtendForecastValues`

**Functionality:**
- Identifies the last date column with content in the "Revenue Summary" sheet.
- Processes each row to extend the forecast values by copying the previous month's value into the new month if the project's end date extends into the new month.

### Macro: `UpdateTotal2025`

**Functionality:**
- Finds the "Total 2025" column in the "Revenue Summary" sheet.
- Sums up values across each row for the year 2025 and populates the "Total 2025" column with the calculated sum.

### Macro: `UpdateRemainingContractAmount`

**Functionality:**
- Updates the "Remaining Contract Amount" by subtracting the sum of all monthly values from the "CONTRACT $" amount.
- Adjusts future forecast values in case of a negative remaining contract amount, ensuring no over-extension of the budget.

### Helper Functions

- `FindNextVisibleRow`: Finds the next visible row within a specified column that is not hidden and contains data, used across various macros to locate the relevant rows for processing.
