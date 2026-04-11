## Goal
Update LMS_Run_Selection_2206_-_Rev1.xlsm to work with 
Simcenter Testlab 2406 (currently works on 2206 only)

## Key Finding
Likely fix is updating the TLB reference path in VBA from:
- 2206 TLB path: C:\Users\lic\OneDrive - American Axle & Manufacturing, Inc\ClaudeWork\TestLab\LMSTestLabAutomation_2206.tlb
- 2406 TLB path: C:\Users\lic\OneDrive - American Axle & Manufacturing, Inc\ClaudeWork\TestLab\LMSTestLabAutomation_2406.tlb

## InputBasket API
Confirmed unchanged between 2206 and 2406 versions.
RInputBasket, Clear, InputBasketObject, Count, Item all identical.

## Files Available
- LMS_Run_Selection_2206_-_Rev1.xlsm (working 2206 version but with some bugs)
- LMSTestLabAutomation_2206.tlb (2206 type library)
- LMSTestLabAutomation_2406.tlb (2406 type library)
- Simcenter_Testlab_Automation_Excel_Demo_Script.txt
- 2025 August - Simcenter Testlab Automation for Beginners.pdf (reference)
- Automation Manual.pdf (reference)
- Automation_Overview.pdf (reference)