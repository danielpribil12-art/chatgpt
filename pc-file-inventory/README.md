# Windows File Inventory

A PowerShell utility with an Arch-inspired pastel neon Windows desktop interface for scanning files by type and size, exporting CSV reports, and reviewing old temporary files. The GUI is named Archiva and is available from the PC File Inventory desktop shortcut.

## Open the desktop app

Run FileInventory-GUI.ps1 in Windows PowerShell, or open the PC File Inventory icon on the desktop. It lets you choose drives or add folders, scan in the background, review totals by file type, and view the 500 largest files. Reports are saved in Documents\FileInventoryReports by default. The full inventory streams to disk as it scans to keep memory use bounded on large drives.

The desktop app can also find old .tmp and .temp files in the current user's temp folders. Select the items you want and confirm before deletion. It does not clean Downloads, the Recycle Bin, browser caches, Windows system folders, or files outside those temp folders.

## Run the command-line scanner

    powershell -ExecutionPolicy Bypass -File .\FileInventory.ps1 -OutputDirectory "$env:USERPROFILE\Documents\FileInventoryReports"

Pass one or more folders or drives explicitly with -Root:

    .\FileInventory.ps1 -Root 'C:\Users\Dan\Documents','D:\' -OutputDirectory '.\reports'

The scan skips folders it cannot read and does not follow directory junctions or symbolic links, preventing loops and duplicate results.

## Find cleanup candidates from the command line

This only creates a CSV. It does not delete anything. Candidates are limited to .tmp and .temp files older than seven days in the current user's temporary folders:

    .\FileInventory.ps1 -FindJunk -OutputDirectory '.\reports'

Change the age threshold with -OlderThanDays 14. Review the generated cleanup-candidates-*.csv before continuing.

## Delete reviewed candidates from the command line

Deletion requires both -DeleteCandidates and -ApplyCleanup, plus typing DELETE at the prompt. The script rechecks every path and only allows old .tmp or .temp files inside the user's temp folders. -WhatIf previews PowerShell's delete operations:

    .\FileInventory.ps1 -DeleteCandidates '.\reports\cleanup-candidates-20261113-120000.csv' -ApplyCleanup

    .\FileInventory.ps1 -DeleteCandidates '.\reports\cleanup-candidates-20261113-120000.csv' -ApplyCleanup -WhatIf

The command-line cleanup permanently deletes only confirmed files.

