---
name: recon-compliance
description: Reviews reconciliation workbooks in the `Recon_Files` folder, extracts sign-off information from the **SignOff** worksheet, validates compliance rules, and generates a consolidated Excel compliance report in the `reports` folder.
---

# Recon Compliance Skill

## Purpose
Validate reconciliation sign-off compliance.

## Input
- Folder: `Recon_Files`
- Worksheet: **SignOff**

## Fields Extracted

| Label | Cell |
|---|---|
| Performed_By | D30 |
| Performed_Date | D31 |
| Reviewed_By | G30 |
| Reviewed_Date | G31 |

## Compliance Rules

- **compliance1**: `Performed_By` must be different from `Reviewed_By`.
- **compliance2**: `Reviewed_Date` must be greater than or equal to `Performed_Date`.

## Workflow

1. Scan all reconciliation workbooks in `Recon_Files`.
2. Open the **SignOff** worksheet.
3. Read cells D30, D31, G30 and G31.
4. Append the values to a pandas DataFrame.
5. Calculate `compliance1` and `compliance2`.
6. Export the DataFrame to `reports/Recon_Compliance_Report.xlsx`.
7. Highlight all FALSE values in the compliance columns.
8. If the worksheet is missing or cannot be processed, record the failure in a **Remarks** column.
