---
name: recon-ai-summary
description: Reviews every reconciliation workbook in the `Recon_Files` folder, analyzes the reconciliation table from the SignOff worksheet, generates an AI summary highlighting key observations, risks, and recommendations, and writes the summary into cell E10.
---

# Recon AI Summary Skill

## Purpose
Generate an executive reconciliation summary for every workbook.

## Input
- Folder: `Recon_Files`
- Worksheet: **SignOff**
- Source table: `C19:G27`
- Header row: `C19:G19`

## AI Requirements
- Maximum 5 lines.
- Summarize reconciliation status.
- Highlight risks and unusual items.
- Provide actionable recommendations.
- Use professional finance/audit language.

## Workflow
1. Scan every workbook in `Recon_Files`.
2. Open the **SignOff** worksheet.
3. Read table `C19:G27`.
4. Convert the table to a structured dataset.
5. Send the dataset to the AI model.
6. Write the generated summary into cell `E10`.
7. Save the workbook.
