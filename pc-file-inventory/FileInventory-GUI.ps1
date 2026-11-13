[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:reportDir = Join-Path $env:USERPROFILE 'Documents\FileInventoryReports'
New-Item -ItemType Directory -Force -Path $script:reportDir | Out-Null
$script:scanProcess = $null
$script:scanStarted = $null
$script:deleteCandidates = @()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PC File Inventory'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1120, 760)
$form.MinimumSize = New-Object System.Drawing.Size(850, 600)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 252)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'PC File Inventory'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 19)
$title.Location = New-Object System.Drawing.Point(20, 15)
$title.AutoSize = $true
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'Scan drives, review file sizes, and safely check temporary files.'
$subtitle.Location = New-Object System.Drawing.Point(23, 54)
$subtitle.AutoSize = $true
$subtitle.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($subtitle)

$driveLabel = New-Object System.Windows.Forms.Label
$driveLabel.Text = 'Drives to scan'
$driveLabel.Location = New-Object System.Drawing.Point(22, 95)
$driveLabel.AutoSize = $true
$form.Controls.Add($driveLabel)

$driveList = New-Object System.Windows.Forms.CheckedListBox
$driveList.Location = New-Object System.Drawing.Point(22, 119)
$driveList.Size = New-Object System.Drawing.Size(245, 95)
$driveList.CheckOnClick = $true
$driveList.BorderStyle = 'FixedSingle'
$form.Controls.Add($driveList)

try {
    Get-CimInstance Win32_LogicalDisk -Filter 'DriveType = 3' | ForEach-Object {
        $index = $driveList.Items.Add(('{0} ({1})' -f $_.DeviceID, $_.VolumeName))
        $driveList.SetItemChecked($index, $true)
    }
} catch {
    $index = $driveList.Items.Add(($env:SystemDrive + '\'))
    $driveList.SetItemChecked($index, $true)
}

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = 'Add folder'
$folderButton.Location = New-Object System.Drawing.Point(22, 221)
$folderButton.Size = New-Object System.Drawing.Size(105, 32)
$form.Controls.Add($folderButton)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = 'Reports folder'
$outputLabel.Location = New-Object System.Drawing.Point(290, 95)
$outputLabel.AutoSize = $true
$form.Controls.Add($outputLabel)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Text = $script:reportDir
$outputBox.Location = New-Object System.Drawing.Point(290, 119)
$outputBox.Size = New-Object System.Drawing.Size(555, 27)
$form.Controls.Add($outputBox)

$outputBrowse = New-Object System.Windows.Forms.Button
$outputBrowse.Text = 'Browse...'
$outputBrowse.Location = New-Object System.Drawing.Point(855, 117)
$outputBrowse.Size = New-Object System.Drawing.Size(92, 30)
$form.Controls.Add($outputBrowse)

$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Text = 'Scan selected locations'
$scanButton.Location = New-Object System.Drawing.Point(290, 164)
$scanButton.Size = New-Object System.Drawing.Size(200, 40)
$scanButton.BackColor = [System.Drawing.Color]::FromArgb(95, 58, 200)
$scanButton.ForeColor = [System.Drawing.Color]::White
$scanButton.FlatStyle = 'Flat'
$form.Controls.Add($scanButton)

$cleanupButton = New-Object System.Windows.Forms.Button
$cleanupButton.Text = 'Review temporary files'
$cleanupButton.Location = New-Object System.Drawing.Point(504, 164)
$cleanupButton.Size = New-Object System.Drawing.Size(195, 40)
$form.Controls.Add($cleanupButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = 'Delete selected'
$deleteButton.Location = New-Object System.Drawing.Point(713, 164)
$deleteButton.Size = New-Object System.Drawing.Size(145, 40)
$deleteButton.Enabled = $false
$form.Controls.Add($deleteButton)

$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = 'Open reports'
$openButton.Location = New-Object System.Drawing.Point(872, 164)
$openButton.Size = New-Object System.Drawing.Size(105, 40)
$form.Controls.Add($openButton)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Ready'
$status.Location = New-Object System.Drawing.Point(290, 218)
$status.Size = New-Object System.Drawing.Size(760, 25)
$status.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 80)
$form.Controls.Add($status)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(20, 270)
$tabs.Size = New-Object System.Drawing.Size(1060, 435)
$form.Controls.Add($tabs)

$typeTab = New-Object System.Windows.Forms.TabPage
$typeTab.Text = 'Files by type'
$tabs.TabPages.Add($typeTab) | Out-Null
$fileTab = New-Object System.Windows.Forms.TabPage
$fileTab.Text = 'Largest files'
$tabs.TabPages.Add($fileTab) | Out-Null
$tempTab = New-Object System.Windows.Forms.TabPage
$tempTab.Text = 'Cleanup review'
$tabs.TabPages.Add($tempTab) | Out-Null

function New-Grid {
    param([System.Windows.Forms.Control] $Parent)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.RowHeadersVisible = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $Parent.Controls.Add($grid)
    return $grid
}
$typeGrid = New-Grid $typeTab
$fileGrid = New-Grid $fileTab
$tempGrid = New-Grid $tempTab
$tempGrid.ReadOnly = $false
$tempGrid.Columns.Clear()
$tempGrid.AutoGenerateColumns = $false
$selectColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$selectColumn.Name = 'Select'
$selectColumn.HeaderText = 'Select'
$selectColumn.FillWeight = 25
$tempGrid.Columns.Add($selectColumn) | Out-Null
foreach ($column in @(
    @{ Name = 'Path'; Header = 'Temporary file'; Weight = 200 },
    @{ Name = 'SizeBytes'; Header = 'Size (bytes)'; Weight = 45 },
    @{ Name = 'LastWriteTime'; Header = 'Last changed'; Weight = 70 },
    @{ Name = 'Reason'; Header = 'Why listed'; Weight = 80 }
)) {
    $gridColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $gridColumn.Name = $column.Name
    $gridColumn.HeaderText = $column.Header
    $gridColumn.FillWeight = $column.Weight
    $gridColumn.ReadOnly = $true
    $tempGrid.Columns.Add($gridColumn) | Out-Null
}

$script:timer = New-Object System.Windows.Forms.Timer
$script:timer.Interval = 800
$script:timer.Add_Tick({
    if ($null -ne $script:scanProcess -and $script:scanProcess.HasExited) {
        $script:timer.Stop()
        $script:scanProcess = $null
        $summaryFiles = @(Get-ChildItem -LiteralPath $script:reportDir -Filter 'file-summary-*.csv' -File | Where-Object { $_.LastWriteTime -ge $script:scanStarted } | Sort-Object LastWriteTime -Descending)
        $inventoryFiles = @(Get-ChildItem -LiteralPath $script:reportDir -Filter 'file-inventory-*.csv' -File | Where-Object { $_.LastWriteTime -ge $script:scanStarted } | Sort-Object LastWriteTime -Descending)
        if ($summaryFiles.Count -gt 0) {
            $typeGrid.DataSource = @(Import-Csv -LiteralPath $summaryFiles[0].FullName)
            if ($inventoryFiles.Count -gt 0) {
                $largest = @(Import-Csv -LiteralPath $inventoryFiles[0].FullName | Sort-Object { [long]$_.SizeBytes } -Descending | Select-Object -First 500)
                $fileGrid.DataSource = $largest
            }
            $status.Text = 'Scan complete. Full CSV reports are saved in ' + $script:reportDir
        } else {
            $status.Text = 'Scan ended without a report. Check the selected locations and folder permissions.'
        }
        $scanButton.Enabled = $true
        $cleanupButton.Enabled = $true
    }
})

$folderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Choose an additional folder to scan'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $index = $driveList.Items.Add($dialog.SelectedPath)
        $driveList.SetItemChecked($index, $true)
    }
})
$outputBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Choose where scan reports should be saved'
    $dialog.SelectedPath = $outputBox.Text
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $outputBox.Text = $dialog.SelectedPath }
})
$openButton.Add_Click({
    if (Test-Path -LiteralPath $outputBox.Text) { Start-Process explorer.exe -ArgumentList ('"' + $outputBox.Text + '"') }
})
$scanButton.Add_Click({
    if ($null -ne $script:scanProcess) { return }
    $roots = @()
    foreach ($selected in $driveList.CheckedItems) {
        $value = [string]$selected
        if ($value -match '^([A-Za-z]:) \(') { $roots += $Matches[1] + '\' } else { $roots += $value }
    }
    if ($roots.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select a drive or add a folder first.', 'No locations selected', 'OK', 'Information') | Out-Null
        return
    }
    $reportPath = [IO.Path]::GetFullPath($outputBox.Text)
    New-Item -ItemType Directory -Force -Path $reportPath | Out-Null
    $script:reportDir = $reportPath
    $rootValues = ($roots | ForEach-Object { "'" + $_.Replace("'", "''") + "'" }) -join ','
    $scriptPath = (Join-Path $PSScriptRoot 'FileInventory.ps1').Replace("'", "''")
    $outValue = $reportPath.Replace("'", "''")
    $runner = Join-Path $env:TEMP ('file-inventory-scan-' + [guid]::NewGuid().ToString('N') + '.ps1')
    $runnerContent = "& '" + $scriptPath + "' -Root @(" + $rootValues + ") -OutputDirectory '" + $outValue + "'"
    Set-Content -LiteralPath $runner -Value $runnerContent -Encoding UTF8
    $script:scanStarted = Get-Date
    $psExe = Join-Path $PSHOME 'powershell.exe'
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $runner + '"'
    $script:scanProcess = Start-Process -FilePath $psExe -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $scanButton.Enabled = $false
    $cleanupButton.Enabled = $false
    $status.Text = 'Scanning in the background. Large drives can take several minutes.'
    $script:timer.Start()
})
$cleanupButton.Add_Click({
    $status.Text = 'Looking for temporary files older than seven days...'
    $tempGrid.Rows.Clear()
    $script:deleteCandidates = @()
    $deleteButton.Enabled = $false
    $cutoff = (Get-Date).AddDays(-7)
    $tempRoots = @($env:TEMP, (Join-Path $env:LOCALAPPDATA 'Temp')) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Container) } |
        ForEach-Object { [IO.Path]::GetFullPath($_).TrimEnd('\') } |
        Select-Object -Unique
    foreach ($tempRoot in $tempRoots) {
        $stack = [System.Collections.Generic.Stack[string]]::new()
        $stack.Push($tempRoot)
        while ($stack.Count -gt 0) {
            $folder = $stack.Pop()
            try { $children = Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop } catch { continue }
            foreach ($item in $children) {
                if ($item.PSIsContainer) {
                    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) { $stack.Push($item.FullName) }
                    continue
                }
                if ($item.Extension -match '^\.(tmp|temp)$' -and $item.LastWriteTime -lt $cutoff -and
                    ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) {
                    $script:deleteCandidates += $item.FullName
                    $idx = $tempGrid.Rows.Add($false, $item.FullName, [long]$item.Length, $item.LastWriteTime.ToString('g'), 'Old temporary file')
                }
            }
        }
    }
    $deleteButton.Enabled = ($tempGrid.Rows.Count -gt 0)
    $status.Text = ('Found ' + $tempGrid.Rows.Count + ' old temporary files. Select only the items you want to remove.')
    $tabs.SelectedTab = $tempTab
})
$deleteButton.Add_Click({
    $tempGrid.EndEdit()
    $chosen = @()
    foreach ($row in $tempGrid.Rows) { if ($row.Cells['Select'].Value -eq $true) { $chosen += [string]$row.Cells['Path'].Value } }
    if ($chosen.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select the files you want to delete first.', 'Nothing selected', 'OK', 'Information') | Out-Null
        return
    }
    $total = 0L
    foreach ($path in $chosen) {
        $row = $tempGrid.Rows | Where-Object { $_.Cells['Path'].Value -eq $path } | Select-Object -First 1
        if ($null -ne $row) { $total += [long]$row.Cells['SizeBytes'].Value }
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ('Permanently delete ' + $chosen.Count + ' selected temporary files (' + [math]::Round($total / 1MB, 2) + ' MiB)?'),
        'Confirm deletion', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $allowed = @($env:TEMP, (Join-Path $env:LOCALAPPDATA 'Temp')) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Container) } |
        ForEach-Object { [IO.Path]::GetFullPath($_).TrimEnd('\') } |
        Select-Object -Unique
    $deleted = 0
    foreach ($path in $chosen) {
        $full = [IO.Path]::GetFullPath($path)
        $safe = $false
        foreach ($root in $allowed) {
            if ($full.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { $safe = $true; break }
        }
        if (-not $safe) { continue }
        try {
            $file = Get-Item -LiteralPath $full -Force -ErrorAction Stop
            if (-not $file.PSIsContainer -and ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -and
                $file.Extension -match '^\.(tmp|temp)$' -and $file.LastWriteTime -lt (Get-Date).AddDays(-7)) {
                Remove-Item -LiteralPath $full -Force -ErrorAction Stop
                $deleted++
            }
        } catch { }
    }
    $cleanupButton.PerformClick()
    $status.Text = ('Deleted ' + $deleted + ' selected files. Remaining candidates are still listed for review.')
})

[void]$form.ShowDialog()

