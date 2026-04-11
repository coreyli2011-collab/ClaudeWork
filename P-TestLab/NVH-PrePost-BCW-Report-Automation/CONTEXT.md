## Project Name
NVH Pre/Post BCW Report Automation

## Goal
Automate the repetitive parts of building NVH Pre/Post BCW durability reports in PowerPoint.
Reports contain TestLab ActivePictures and follow a fixed 6-slide structure (Drive/Coast x 3 torque loads).

## Status
In progress — Step 2 (generate reports) and Step 7 (update results) implemented and tested.

## What Is Automated
- Step 2: Generate one PPTX per sample with all test information filled in automatically
- Step 7: Update existing reports with actual dB values and Pass/Fail result

## What Remains Manual
- Step 1: User prepares template PPTX (replace actual dB values with XXdB placeholder)
- Steps 3-6: Replace ActivePictures, format them, review plots, fill in dB values
- Step 7 visual: Pass/Fail visual indicator in the table — user manually verifies

## How To Use — Step by Step

### STEP 1 — Prepare the template (manual, once per campaign)
1. Open `working/report_template.pptx` in PowerPoint
2. Replace every actual dB result value (e.g. -23dB, -30dB) with `XXdB`
   - Important: type exactly `XXdB` with no spaces
3. Save the file — keep the same name and location

### STEP 2 — Generate reports for all samples (automated)
1. Open `working/NVH_Report_Input.xlsx`
2. Fill in the **Campaign_Info** sheet with the programme and test details
3. Fill in the **Samples** sheet — one row per sample (Step 2 columns only)
4. Save and close the Excel file
5. Double-click `working/Step2_Generate_Reports.bat`
6. Reports appear in the `reports/` folder — one PPTX per sample

### STEPS 3–6 — Your engineering work (manual)
- Open each report from the `reports/` folder
- Replace the ActivePictures with the correct TestLab plots
- Format the ActivePictures (use copy/paste format from your reference slide)
- Review each plot and fill in the actual dB values directly in PowerPoint
- Add comments as needed

### STEP 7 — Update Pass/Fail results (automated)
1. Open `working/NVH_Report_Input.xlsx`
2. In the **Samples** sheet, fill in the dB values in the **Step 7 columns**
   - Enter as a number only — e.g. `-23` not `-23dB`
   - Negative values = amplitude below target = likely PASS
   - Positive values = amplitude above target = check carefully
3. Save and close the Excel file
4. Double-click `working/Step7_Update_Results.bat`
5. Reports are updated automatically — verify and adjust manually if needed

### PASS / FAIL RULES
| Result | Condition |
|---|---|
| PASS | dB value ≤ +3 (all negative values qualify) |
| CONDITIONAL FAIL | dB value between +3 and +6 |
| ABSOLUTE FAIL | dB value > +6 |

## Key Files (working/)
- report_template.pptx       — User's prepared template (replace dB values with XXdB)
- NVH_Report_Input.xlsx      — Excel input: campaign info + one row per sample
- generate_reports.py        — Step 2 script
- update_results.py          — Step 7 script
- create_input_template.py   — Creates the Excel template (run once)
- Step2_Generate_Reports.bat — Double-click to run Step 2
- Step7_Update_Results.bat   — Double-click to run Step 7

## Slide Structure (all 6 slides per report)
- Slide 1: Drive +450Nm
- Slide 2: Drive +75Nm
- Slide 3: Drive -75Nm
- Slide 4: Coast +450Nm
- Slide 5: Coast +75Nm
- Slide 6: Coast -75Nm

## Pass/Fail Rules
- PASS             : dB value <= +3  (all negative values qualify)
- CONDITIONAL FAIL : dB value between +3 and +6
- ABSOLUTE FAIL    : dB value > +6

## Notes
- Text replacement uses python-pptx, preserving run-level formatting
- Text in grouped shapes handled recursively
- Report filename: NVH_Report_{PartNumber}_SN{SerialNumber}.pptx
- ActivePicture copy-format API not available — investigated TLB and web sources
- Future: Investigate Windows UI Automation for ActivePicture format copying

## Next Steps
- Test with real campaign data
- Confirm Pass/Fail visual indicator behaviour (need a FAIL example to compare)
- Investigate Windows UI Automation for Step 5 formatting automation
