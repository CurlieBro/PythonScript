Excel File Comparison & Update Tool
===================================

Created By: Nizam Rahim
Version: 1.0.0
Date: 01 Oct 2025

Overview
--------
This script compares two Excel files — **HC Report.xlsx** and **Laptop Hostname.xlsx** —
and updates the **User Name** column in *Laptop Hostname.xlsx* using the matching
**Login ID** values from *HC Report.xlsx*. It is designed to be resilient to column
header variations and stray whitespace, and it provides a concise update summary in
the console.

What the Script Does
--------------------
1) **Loads dependencies safely** and exits with a helpful message if required
   packages are missing (pandas, openpyxl).
2) **Reads source files** from `C:\Temp`:
     - `HC Report.xlsx` (source of truth for Login ID → User Name mapping)
     - `Laptop Hostname.xlsx` (target to be updated)
3) **Normalizes headers** so that column names can be matched case-insensitively and
   without trailing spaces. Typical headers it looks for:
     - Login ID: "Login ID", "LoginID", "UserID", "User ID"
     - User Name: "User Name", "Username", "User", "Display Name", "Full Name"
4) **Validates required columns** and aborts with a clear error if any are missing.
5) **Cleans key columns** (strips spaces, lowercases Login IDs) and removes invalid rows.
6) **Left-joins** Hostname data with HC data on Login ID to pull the corresponding
   User Name from HC where available.
7) **Overwrites/fills** the **User Name** column in *Laptop Hostname.xlsx* with values
   from *HC Report.xlsx* when a match is found. If **User Name** does not exist in the
   Hostname file, it creates the column.
8) **Prints an update summary**:
     - Records with Login ID
     - Records updated from HC
     - Login IDs with no match
   A small **sample of updated rows** is shown when updates occur.
9) **Creates a one-click backup** of the original *Laptop Hostname.xlsx* as
   `Laptop Hostname_backup.xlsx` in `C:\Temp` (skips creation if it already exists).
10) **Saves the updated file** back to `C:\Temp\Laptop Hostname.xlsx` (helper join columns
    are dropped to keep the file clean).

Files & Paths
-------------
• Folder: `C:\Temp`
• Source: `C:\Temp\HC Report.xlsx`
• Target: `C:\Temp\Laptop Hostname.xlsx`
• Backup: `C:\Temp\Laptop Hostname_backup.xlsx`

Requirements
------------
• Python 3.8+
• Packages: pandas, openpyxl

Install the required packages (Windows PowerShell/CMD):
    py -m pip install --user pandas openpyxl

How to Run
----------
• Double-click (if associated) or run from terminal:
    py your_script_name.py

• The console will display progress logs and a final success/fail status.
  If dependencies are missing, a clear installation hint is printed and the script exits.

Console Messages & Errors
-------------------------
• Missing dependencies → install hint printed; script exits.
• Missing files → explicit path shown for the file not found.
• Missing required columns → lists which columns were not detected and tips to fix headers.
• Update summary → shows counts of matched/updated records and no-match cases.

Customization Tips
------------------
• **Change folder location**: Update the `temp_dir` Path in the code.
• **Adjust column name variants**: Add your own header aliases in the `normalize_columns`
  lookups for Login ID and User Name.
• **Change backup behavior**: Modify the backup file name or always overwrite by removing
  the "exists" check before backup creation.

Safety Notes
------------
• The script **never modifies the HC Report.xlsx** file.
• A backup of `Laptop Hostname.xlsx` is created before any write (if it does not already exist).
• Only the **User Name** column in the Hostname file is updated; other data remains intact.

Changelog
---------
v1.0.0 (01 Oct 2025) — Initial release:
  • Adds robust header normalization and validation
  • Adds backup creation before save
  • Adds concise update summary and sample preview
  • Cleans helper columns before writing
