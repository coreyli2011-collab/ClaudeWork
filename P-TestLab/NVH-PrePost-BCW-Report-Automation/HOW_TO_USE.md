# NVH Pre/Post BCW Report Generator
## How To Use

---

## What This Tool Does

Every time you have a new test campaign, this tool:
- Automatically generates one PowerPoint report per sample with all test information filled in
- After your engineering review, automatically fills in the dB values and Pass/Fail results

You only need to do the engineering work — the repetitive copy/paste and data entry is handled for you.

---

## Files You Will Use

| File | What It Is |
|---|---|
| `working/report_template.pptx` | Your blank template — prepared once per campaign |
| `working/NVH_Report_Input.xlsx` | Your input form — fill in test and sample data here |
| `working/Step2_Generate_Reports.bat` | Double-click to generate all reports |
| `working/Step7_Update_Results.bat` | Double-click to update reports with results |
| `reports/` | Folder where your completed reports appear |

---

## Step-by-Step Instructions

---

### STEP 1 — Prepare Your Template
*Do this once at the start of each campaign.*

1. Open **`working/report_template.pptx`** in PowerPoint
2. Find every dB result value on every slide (e.g. `-23dB`, `-30dB`, `-16dB`)
3. Replace each one by typing exactly **`XXdB`** — no spaces, capital XX
4. Save the file — keep the same filename and location
5. Close PowerPoint

> **Tip:** The template already contains a copy of your previous report as a starting point.

---

### STEP 2 — Fill In Your Test Information
*Do this before generating reports.*

1. Open **`working/NVH_Report_Input.xlsx`**
2. Go to the **Campaign_Info** sheet — fill in the programme and test details
   - Programme name, test order numbers, test dyno, dates, part ratio, etc.
3. Go to the **Samples** sheet — fill in one row per sample *(blue columns only)*
   - Part number, serial number, sample number, published date
4. **Save and close** the Excel file
5. Double-click **`Step2_Generate_Reports.bat`**
6. A window will appear showing progress — wait for it to finish
7. Your reports appear in the **`reports/`** folder — one per sample

---

### STEPS 3 TO 6 — Your Engineering Work
*This part remains manual — it requires your expertise.*

Open each report from the `reports/` folder and:

- Replace the ActivePictures with the correct TestLab plots
- Format the ActivePictures to match your standard layout
- Review each plot carefully
- Fill in the actual dB values directly in the PowerPoint slides
- Add any comments

---

### STEP 7 — Fill In Results and Update Pass/Fail
*Do this after completing your engineering review.*

1. Open **`working/NVH_Report_Input.xlsx`**
2. Go to the **Samples** sheet — fill in the dB values in the **green columns** (Step 7)
   - Enter numbers only — e.g. type `-23` not `-23dB`
   - Negative numbers = amplitude below target = likely PASS
   - Positive numbers = amplitude above target = review carefully
3. **Save and close** the Excel file
4. Double-click **`Step7_Update_Results.bat`**
5. A window will appear showing progress — wait for it to finish
6. Open each report and **verify** the Pass/Fail results
7. Adjust manually if needed

---

## Pass/Fail Reference

| Result | When It Applies |
|---|---|
| **PASS** | dB value is +3 or below (all negative values are PASS) |
| **CONDITIONAL FAIL** | dB value is between +3 and +6 |
| **ABSOLUTE FAIL** | dB value is above +6 |

---

## Something Not Working?

Contact the person who set this up, or open a new conversation with Claude Code and describe the problem. The technical details are in `CONTEXT.md` in this same folder.
