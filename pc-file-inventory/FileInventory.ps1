[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]] $Root,
    [string] $OutputDirectory = (Get-Location).Path,
    [switch] $FindJunk,
    [string] $DeleteCandidates,
    [switch] $ApplyCleanup,
    [ValidateRange(1, 3650)]
    [int] $OlderThanDays = 7
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-DefaultRoots {
    try {
        $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' -ErrorAction Stop
        return @($drives | ForEach-Object { $_.DeviceID + '\' })
    } catch {
        Write-Warning ('Could not list fixed drives. Scanning the current user profile instead. ' + $_.Exception.Message)
        return @($env:USERPROFILE)
    }
}

function Get-FileInventory {
    param([string[]] $ScanRoots)
    foreach ($scanRoot in $ScanRoots) {
        if (-not (Test-Path -LiteralPath $scanRoot -PathType Container)) {
            Write-Warning ('Skipping missing or inaccessible root: ' + $scanRoot)
            continue
        }
        Write-Host ('Scanning ' + $scanRoot + ' ...')
        $stack = [System.Collections.Generic.Stack[string]]::new()
        $stack.Push((Get-Item -LiteralPath $scanRoot -Force).FullName)
        while ($stack.Count -gt 0) {
            $current = $stack.Pop()
            try { $children = Get-ChildItem -LiteralPath $current -Force -ErrorAction Stop }
            catch { Write-Verbose ('Skipped inaccessible folder: ' + $current); continue }
            foreach ($item in $children) {
                try {
                    if ($item.PSIsContainer) {
                        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) { $stack.Push($item.FullName) }
                        continue
                    }
                    $extension = $item.Extension.ToLowerInvariant()
                    if ([string]::IsNullOrWhiteSpace($extension)) { $extension = '[no extension]' }
                    [pscustomobject]@{
                        Path = $item.FullName
                        Name = $item.Name
                        Extension = $extension
                        SizeBytes = [long]$item.Length
                        SizeMiB = [math]::Round($item.Length / 1MB, 3)
                        LastWriteTimeUtc = $item.LastWriteTimeUtc.ToString('o')
                        Type = if ($extension -eq '[no extension]') { 'No extension' } else { $extension.TrimStart('.').ToUpperInvariant() }
                    }
                } catch { Write-Verbose ('Skipped file that could not be read: ' + $item.FullName) }
            }
        }
    }
}

function Get-AllowedTempRoots {
    @($env:TEMP, (Join-Path $env:LOCALAPPDATA 'Temp')) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Container) } |
        ForEach-Object { [IO.Path]::GetFullPath($_).TrimEnd('\') } |
        Select-Object -Unique
}

function Test-InAllowedTempRoot {
    param([string] $Path, [string[]] $AllowedRoots)
    $fullPath = [IO.Path]::GetFullPath($Path)
    foreach ($allowedRoot in $AllowedRoots) {
        if ($fullPath.StartsWith($allowedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Export-JunkCandidates {
    param([string] $Destination, [int] $AgeDays)
    $allowedRoots = @(Get-AllowedTempRoots)
    if ($allowedRoots.Count -eq 0) { throw 'No user temporary folders were found.' }
    $cutoff = (Get-Date).AddDays(-$AgeDays)
    $results = foreach ($rootPath in $allowedRoots) {
        Get-ChildItem -LiteralPath $rootPath -File -Force -Recurse -ErrorAction SilentlyContinue |
            Where-Object {
                $_.LastWriteTime -lt $cutoff -and $_.Extension -match '^\.(tmp|temp)$' -and
                (($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0)
            } |
            ForEach-Object {
                [pscustomobject]@{
                    Path = $_.FullName
                    SizeBytes = [long]$_.Length
                    LastWriteTimeUtc = $_.LastWriteTimeUtc.ToString('o')
                    Reason = 'Old .' + $_.Extension.TrimStart('.') + ' file inside user temp folder'
                }
            }
    }
    $results | Export-Csv -LiteralPath $Destination -NoTypeInformation -Encoding UTF8
    Write-Host ('Wrote ' + @($results).Count + ' cleanup candidates to ' + $Destination)
    Write-Host ('Candidates are limited to .tmp/.temp files older than ' + $AgeDays + ' days in your user temp folders.')
}

function Remove-ReviewedCandidates {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string] $CandidateFile)
    if (-not (Test-Path -LiteralPath $CandidateFile -PathType Leaf)) { throw ('Candidate file not found: ' + $CandidateFile) }
    $allowedRoots = @(Get-AllowedTempRoots)
    if ($allowedRoots.Count -eq 0) { throw 'No user temporary folders were found.' }
    $rows = @(Import-Csv -LiteralPath $CandidateFile)
    if ($rows.Count -eq 0) { Write-Host 'No cleanup candidates to process.'; return }

    $validPaths = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $rows) {
        $candidatePath = [string]$row.Path
        if ([string]::IsNullOrWhiteSpace($candidatePath) -or -not (Test-InAllowedTempRoot -Path $candidatePath -AllowedRoots $allowedRoots)) {
            Write-Warning ('Rejected path outside user temp folders: ' + $candidatePath)
            continue
        }
        try {
            $file = Get-Item -LiteralPath $candidatePath -Force -ErrorAction Stop
            if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
                $file.Extension -notmatch '^\.(tmp|temp)$' -or $file.LastWriteTime -ge (Get-Date).AddDays(-$OlderThanDays)) {
                Write-Warning ('Rejected file that no longer matches cleanup rules: ' + $candidatePath)
                continue
            }
            $validPaths.Add($file.FullName)
        } catch { Write-Verbose ('Skipped unavailable candidate: ' + $candidatePath) }
    }
    if ($validPaths.Count -eq 0) { Write-Host 'No valid candidates remain.'; return }

    Write-Host ('Files eligible for deletion (' + $validPaths.Count + '):')
    $validPaths | ForEach-Object { Write-Host ('  ' + $_) }
    $answer = Read-Host 'Permanently delete these files? Type DELETE to continue'
    if ($answer -cne 'DELETE') { Write-Host 'Cleanup cancelled. No files were deleted.'; return }
    foreach ($path in $validPaths) {
        if ($PSCmdlet.ShouldProcess($path, 'Permanently delete reviewed temporary file')) {
            try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; Write-Host ('Deleted: ' + $path) }
            catch { Write-Warning ('Could not delete ' + $path + ' : ' + $_.Exception.Message) }
        }
    }
}

if ($DeleteCandidates) {
    if (-not $ApplyCleanup) { throw 'Deletion is disabled by default. Add -ApplyCleanup to review a candidate CSV and continue.' }
    Remove-ReviewedCandidates -CandidateFile $DeleteCandidates
    return
}
if ($ApplyCleanup) { throw 'Use -DeleteCandidates <CSV> together with -ApplyCleanup to delete reviewed items.' }
if (-not $Root -or $Root.Count -eq 0) { $Root = @(Get-DefaultRoots) }
if (-not $Root -or $Root.Count -eq 0) { throw 'No scan roots were supplied or discovered.' }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$inventoryPath = Join-Path $OutputDirectory ('file-inventory-' + $stamp + '.csv')
$summaryPath = Join-Path $OutputDirectory ('file-summary-' + $stamp + '.csv')
$inventory = @(Get-FileInventory -ScanRoots $Root)
$inventory | Export-Csv -LiteralPath $inventoryPath -NoTypeInformation -Encoding UTF8
$summary = $inventory | Group-Object Type | ForEach-Object {
    $totalBytes = ($_.Group | Measure-Object -Property SizeBytes -Sum).Sum
    [pscustomobject]@{ Type = $_.Name; FileCount = $_.Count; TotalBytes = [long]$totalBytes; TotalMiB = [math]::Round($totalBytes / 1MB, 3) }
} | Sort-Object TotalBytes -Descending
$summary | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8
Write-Host ('Inventory: ' + $inventoryPath)
Write-Host ('Summary:   ' + $summaryPath)
Write-Host ('Files found: ' + $inventory.Count)
Write-Host 'Largest types:'
$summary | Select-Object -First 15 | Format-Table -AutoSize
if ($FindJunk) {
    $candidatePath = Join-Path $OutputDirectory ('cleanup-candidates-' + $stamp + '.csv')
    Export-JunkCandidates -Destination $candidatePath -AgeDays $OlderThanDays
}

