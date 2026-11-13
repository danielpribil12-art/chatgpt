# Windows File Inventory

A PowerShell utility that inventories files by extension/type and size, exports a full CSV plus a size summary, and can prepare a conservative list of old temporary files for review.

## Scan fixed drives

    powershell -ExecutionPolicy Bypass -File .\FileInventory.ps1 -OutputDirectory "$env:USERPROFILE\Desktop\Inventory"

Pass one or more folders or drives explicitly with -Root:

    .\FileInventory.ps1 -Root 'C:\Users\Dan\Documents','D:\' -OutputDirectory '.\reports'

The scan skips folders it cannot read and does not follow directory junctions or symbolic links, preventing loops and duplicate results.

## Find cleanup candidates

This only creates a CSV. It does not delete anything. Candidates are limited to .tmp and .temp files older than seven days in the current user's temporary folders:

    .\FileInventory.ps1 -FindJunk -OutputDirectory '.\reports'

Change the age threshold with -OlderThanDays 14. Review the generated cleanup-candidates-*.csv before continuing.

## Delete reviewed candidates

Deletion requires both -DeleteCandidates and -ApplyCleanup, plus typing DELETE at the prompt. The script rechecks every path and only allows old .tmp or .temp files inside the user's temp folders. -WhatIf previews PowerShell's delete operations:

    .\FileInventory.ps1 -DeleteCandidates '.\reports\cleanup-candidates-20261113-120000.csv' -ApplyCleanup

    .\FileInventory.ps1 -DeleteCandidates '.\reports\cleanup-candidates-20261113-120000.csv' -ApplyCleanup -WhatIf

The tool permanently deletes only confirmed files. It does not clean Downloads, Recycle Bin, browser caches, Windows system folders, or arbitrary files named junk.

