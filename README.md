# Intelligent Excel Report Generator (VBA)

## Overview

This VBA (Visual Basic for Applications) macro automates the final review process for Excel-based Garments Quality Control (QC) inspection reports. It eliminates the manual effort required to check multiple conditions, determine the overall PASS/FAIL status, and write a consistent summary comment.

## Features

* **One-Click Operation:** Runs directly within the Excel template via a button click or shortcut.
* **Multi-Sheet Analysis:** Analyzes data across multiple "Page" worksheets containing detailed inspection results (e.g., defect points, shading values).
* **PASS/FAIL Logic:** Automatically determines the final report status (stored in cell `B49` of the "Summary" sheet) based on complex business rules:
    * Checks if the Average Defect Point (`B41`) exceeds the Standard Point (`B43`).
    * Checks if the percentage of rolls with Critical Shading (rating <= 4) meets or exceeds a defined threshold (e.g., 20%).
* **Intelligent Comment Generation:** Auto-generates a detailed and structured summary comment (in cell `F47` of the "Summary" sheet) explaining the result:
    * Clearly states the primary reason(s) for failure (e.g., "DUE TO HIGH AVG POINT & CRITICAL SHADING.").
    * Lists the top N (e.g., 3) defects contributing the most points.
    * Includes details about shading issues found (e.g., "CSV, SSV SHADE RANGE-3 TO 3/4").
    * Appends other relevant findings like length/width shortages, bowing/skewing, or GSM/Moisture variations in a numbered list.
* **Individual Roll Check:** Calculates and reports the number of rolls and total yards failing the individual standard point (updates cells `B45`, `B47`).
* **User Input:** Prompts the user to specify fabric type (Solid/Stripe) for correct bowing/skewing analysis and the individual standard point.

## Technologies Used

* **Microsoft Excel**
* **VBA (Visual Basic for Applications)**

## How to Use (Conceptual)

1.  Ensure the VBA code is added as a module within your Excel inspection report template (`.xlsm` file).
2.  Assign the `GenerateResultAndComment` macro to a button on the "Summary" sheet or run it directly (Alt+F8).
3.  Input the Fabric Type and Individual Standard Point when prompted.
4.  The macro runs, updates the PASS/FAIL status (`B49`), the detailed comment (`F47`), and other relevant cells (`B45`, `B47`, `E12`).

*(This project demonstrates proficiency in VBA programming for automating complex decision-making and reporting tasks directly within Microsoft Excel.)*
