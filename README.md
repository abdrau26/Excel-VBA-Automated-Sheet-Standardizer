
A VBA automation tool designed to insert standardized column headers and apply uniform formatting across multiple worksheets within a Microsoft Excel workbook.

---
## Project Overview
When processing raw data exports across multiple sheets, manual header creation and styling are repetitive and prone to inconsistency. This project automates data preparation by combining macro-recorded formatting procedures with a programmatic VBA loop. 

The script dynamically iterates through evewry sheet in the target workbook (`CleaningUpData.xlsm`), applies standardized header titles and visual styling, and skips template or reference sheets such as `RECORDING`.

---

## Key Features

* **Automated Row Insertion**: Dynamically creates a dedicated header row across targeted sheets.
* **Uniform Styling Protocol**: Enforces standard typography, bold weight, background fill, and border styling across header rows.
* **Selective Processing**: Includes logic to bypass specific reference worksheets.
* **Extensible Macro Architecture**: Modular design separating the execution loop from specific layout routines for easy maintenance.

---

## Technical Stack

| Component       | Technology                          |
| :-------------- | :---------------------------------- |
| **Platform**    | Microsoft Excel                     |
| **Language**    | Visual Basic for Applications (VBA) |
| **File Format** | Macro-Enabled Workbook (`.xlsm`)    |

---

## Implementation Details

The module consists of three core components:

1. **`AddHeader` (Subroutine)**: Inserts a new row at position 1 and populates standard column headers.
2. **`FormatHeaders` (Subroutine)**: Applies typography, font coloring, background fill, and thick outer borders to the header range.
3. **`CleanUpData` (Main Driver Routine)**: Iterates across the `Worksheets` collection, evaluating worksheet names before invoking the header insertion and formatting macros.

### Implementation Code


```vba
Public Sub CleanUpData()

    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
        If Worksheets(i).Name <> "RECORDING" Then
            Worksheets(i).Select
            AddHeader
            FormatHeaders
            
        End If
        i = i + 1
    Loop


End Sub
```

## Prerequisites
- Microsoft 365 Apps
- Macros enabled in Excel Security Settings
## Execution Steps
1. Open `CleaningUpData.xlsm`
2. Press `ALT + F11` to launch the Visual Basic for Applications editor.
3. Verify that `AddHeader`, `FormatHeaders`, and `CleanUpData` are present inside the standard module.
4. Press `ALT + F8`, select `CleanUpData`, and click **Run**
## Author
**Developer**: Abdul Rauff