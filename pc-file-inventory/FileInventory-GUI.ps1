[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:palette = @{
    Background = [System.Drawing.ColorTranslator]::FromHtml('#11111B')
    Panel = [System.Drawing.ColorTranslator]::FromHtml('#181825')
    Surface = [System.Drawing.ColorTranslator]::FromHtml('#1E1E2E')
    Border = [System.Drawing.ColorTranslator]::FromHtml('#313244')
    Text = [System.Drawing.ColorTranslator]::FromHtml('#CDD6F4')
    Muted = [System.Drawing.ColorTranslator]::FromHtml('#A6ADC8')
    Blue = [System.Drawing.ColorTranslator]::FromHtml('#89B4FA')
    Cyan = [System.Drawing.ColorTranslator]::FromHtml('#89DCEB')
    Mauve = [System.Drawing.ColorTranslator]::FromHtml('#CBA6F7')
    Pink = [System.Drawing.ColorTranslator]::FromHtml('#F5C2E7')
    Green = [System.Drawing.ColorTranslator]::FromHtml('#A6E3A1')
    Red = [System.Drawing.ColorTranslator]::FromHtml('#F38BA8')
}

$script:reportDir = Join-Path $env:USERPROFILE 'Documents\FileInventoryReports'
New-Item -ItemType Directory -Force -Path $script:reportDir | Out-Null
$script:scanProcess = $null
$script:scanStarted = $null
$script:runnerPath = $null
$script:operationType = $null
$script:reviewSourcePath = $null
$script:reviewResultPath = $null
$script:reviewCountPath = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = 'archiva.exe // drive scout'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1120, 760)
$form.MinimumSize = New-Object System.Drawing.Size(850, 600)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = $script:palette.Background
$form.ForeColor = $script:palette.Text
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'archiva.exe'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 22)
$title.ForeColor = $script:palette.Pink
$title.Location = New-Object System.Drawing.Point(82, 14)
$title.AutoSize = $true
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = '[ dan@winbox ~ ]  arch energy, Windows disks  ·  nya.sys ♡'
$subtitle.Location = New-Object System.Drawing.Point(84, 58)
$subtitle.AutoSize = $true
$subtitle.ForeColor = $script:palette.Cyan
$form.Controls.Add($subtitle)

$brandBadge = New-Object System.Windows.Forms.Label
$brandBadge.Text = 'A'
$brandBadge.TextAlign = 'MiddleCenter'
$brandBadge.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 25)
$brandBadge.ForeColor = $script:palette.Background
$brandBadge.BackColor = $script:palette.Mauve
$brandBadge.Location = New-Object System.Drawing.Point(22, 20)
$brandBadge.Size = New-Object System.Drawing.Size(48, 48)
$form.Controls.Add($brandBadge)

$driveLabel = New-Object System.Windows.Forms.Label
$driveLabel.Text = 'TARGETS // MOUNT POINTS'
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
$scanButton.Text = '>_ scan targets'
$scanButton.Location = New-Object System.Drawing.Point(290, 164)
$scanButton.Size = New-Object System.Drawing.Size(200, 40)
$scanButton.BackColor = $script:palette.Mauve
$scanButton.ForeColor = $script:palette.Background
$scanButton.FlatStyle = 'Flat'
$form.Controls.Add($scanButton)

$cleanupButton = New-Object System.Windows.Forms.Button
$cleanupButton.Text = '~ review temp files'
$cleanupButton.Location = New-Object System.Drawing.Point(504, 164)
$cleanupButton.Size = New-Object System.Drawing.Size(195, 40)
$form.Controls.Add($cleanupButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = 'rm -i selected'
$deleteButton.Location = New-Object System.Drawing.Point(713, 164)
$deleteButton.Size = New-Object System.Drawing.Size(145, 40)
$deleteButton.Enabled = $false
$form.Controls.Add($deleteButton)

$openButton = New-Object System.Windows.Forms.Button
$reviewButton = New-Object System.Windows.Forms.Button
$reviewButton.Text = 'review saved CSV'
$reviewButton.Location = New-Object System.Drawing.Point(872, 164)
$reviewButton.Size = New-Object System.Drawing.Size(108, 40)
$form.Controls.Add($reviewButton)

$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = 'open /reports'
$openButton.Location = New-Object System.Drawing.Point(990, 164)
$openButton.Size = New-Object System.Drawing.Size(100, 40)
$form.Controls.Add($openButton)

$status = New-Object System.Windows.Forms.Label
$status.Text = '● ready · no sudo required'
$status.Location = New-Object System.Drawing.Point(290, 218)
$status.Size = New-Object System.Drawing.Size(760, 25)
$status.ForeColor = $script:palette.Green
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

$searchPanel = New-Object System.Windows.Forms.Panel
$searchPanel.Location = New-Object System.Drawing.Point(7, 7)
$searchPanel.Size = New-Object System.Drawing.Size(1035, 36)
$searchPanel.BackColor = $script:palette.Panel
$fileTab.Controls.Add($searchPanel)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = 'Search full CSV'
$searchLabel.Location = New-Object System.Drawing.Point(5, 9)
$searchLabel.AutoSize = $true
$searchLabel.ForeColor = $script:palette.Cyan
$searchPanel.Controls.Add($searchLabel)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(112, 5)
$searchBox.Size = New-Object System.Drawing.Size(660, 26)
$searchBox.BackColor = $script:palette.Surface
$searchBox.ForeColor = $script:palette.Text
$searchBox.BorderStyle = 'FixedSingle'
$searchPanel.Controls.Add($searchBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = 'find matches'
$searchButton.Location = New-Object System.Drawing.Point(782, 4)
$searchButton.Size = New-Object System.Drawing.Size(130, 28)
$searchButton.FlatStyle = 'Flat'
$searchButton.BackColor = $script:palette.Blue
$searchButton.ForeColor = $script:palette.Background
$searchPanel.Controls.Add($searchButton)

$revealButton = New-Object System.Windows.Forms.Button
$revealButton.Text = 'show selected'
$revealButton.Location = New-Object System.Drawing.Point(922, 4)
$revealButton.Size = New-Object System.Drawing.Size(105, 28)
$revealButton.FlatStyle = 'Flat'
$revealButton.BackColor = $script:palette.Surface
$revealButton.ForeColor = $script:palette.Cyan
$searchPanel.Controls.Add($revealButton)

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

function Get-LargestInventoryFiles {
    param([string] $InventoryPath, [int] $Limit = 500)
    $largest = [System.Collections.Generic.SortedDictionary[string, object]]::new([System.StringComparer]::Ordinal)
    $sequence = 0L
    Import-Csv -LiteralPath $InventoryPath | ForEach-Object {
        $key = '{0:D20}-{1:D20}' -f [long]$_.SizeBytes, $sequence
        $largest[$key] = $_
        $sequence++
        if ($largest.Count -gt $Limit) {
            $iterator = $largest.GetEnumerator()
            [void]$iterator.MoveNext()
            [void]$largest.Remove($iterator.Current.Key)
        }
    }
    @($largest.Values | Sort-Object { [long]$_.SizeBytes } -Descending)
}

function Test-SafeTempFilePath {
    param([string] $Path, [string[]] $AllowedRoots)
    try {
        $fullPath = [IO.Path]::GetFullPath($Path)
        foreach ($allowedRoot in $AllowedRoots) {
            $fullRoot = [IO.Path]::GetFullPath($allowedRoot).TrimEnd('\')
            if (-not $fullPath.StartsWith($fullRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { continue }
            $currentPath = $fullPath
            while ($currentPath) {
                $current = Get-Item -LiteralPath $currentPath -Force -ErrorAction Stop
                if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $false }
                if ([string]::Equals([IO.Path]::GetFullPath($current.FullName).TrimEnd('\'), $fullRoot, [StringComparison]::OrdinalIgnoreCase)) { return $true }
                $parent = [IO.Directory]::GetParent($currentPath)
                if ($null -eq $parent -or $parent.FullName -eq $currentPath) { break }
                $currentPath = $parent.FullName
            }
        }
    } catch { return $false }
    return $false
}

function ConvertTo-PowerShellLiteral {
    param([string] $Value)
    "'" + $Value.Replace("'", "''") + "'"
}

function Start-ReportReviewWorker {
    param([ValidateSet('Largest', 'Search')][string] $Mode, [string] $SourcePath, [string] $Query = '')
    if ($null -ne $script:scanProcess) { return }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        $status.ForeColor = $script:palette.Red
        $status.Text = 'The inventory CSV could not be found.'
        return
    }

    $token = [guid]::NewGuid().ToString('N')
    $script:runnerPath = Join-Path $env:TEMP ('archiva-review-' + $token + '.ps1')
    $script:reviewResultPath = Join-Path $env:TEMP ('archiva-results-' + $token + '.csv')
    $script:reviewCountPath = Join-Path $env:TEMP ('archiva-count-' + $token + '.txt')
    $runnerTemplate = @'
$source = __SOURCE__
$output = __OUTPUT__
$countFile = __COUNTFILE__
$mode = __MODE__
$query = __QUERY__
$rows = [System.Collections.Generic.List[object]]::new()
$count = 0L
if ($mode -eq 'Largest') {
    $largest = [System.Collections.Generic.SortedDictionary[string, object]]::new([System.StringComparer]::Ordinal)
    $sequence = 0L
    Import-Csv -LiteralPath $source | ForEach-Object {
        $key = '{0:D20}-{1:D20}' -f [long]$_.SizeBytes, $sequence
        $largest[$key] = $_
        $sequence++
        if ($largest.Count -gt 500) {
            $iterator = $largest.GetEnumerator()
            [void]$iterator.MoveNext()
            [void]$largest.Remove($iterator.Current.Key)
        }
    }
    foreach ($item in ($largest.Values | Sort-Object { [long]$_.SizeBytes } -Descending)) { $rows.Add($item) }
    $count = $rows.Count
} else {
    Import-Csv -LiteralPath $source | ForEach-Object {
        if ($_.Path.IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $_.Name.IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $_.Extension.IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
            $_.Type.IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $count++
            if ($rows.Count -lt 500) { $rows.Add($_) }
        }
    }
}
if ($rows.Count -gt 0) {
    $rows | Export-Csv -LiteralPath $output -NoTypeInformation -Encoding UTF8
} else {
    '"Path","Name","Extension","SizeBytes","SizeMiB","LastWriteTimeUtc","Type"' | Set-Content -LiteralPath $output -Encoding UTF8
}
Set-Content -LiteralPath $countFile -Value $count -Encoding ASCII
'@
    $runnerContent = $runnerTemplate
    $runnerContent = $runnerContent.Replace('__SOURCE__', (ConvertTo-PowerShellLiteral $SourcePath))
    $runnerContent = $runnerContent.Replace('__OUTPUT__', (ConvertTo-PowerShellLiteral $script:reviewResultPath))
    $runnerContent = $runnerContent.Replace('__COUNTFILE__', (ConvertTo-PowerShellLiteral $script:reviewCountPath))
    $runnerContent = $runnerContent.Replace('__MODE__', (ConvertTo-PowerShellLiteral $Mode))
    $runnerContent = $runnerContent.Replace('__QUERY__', (ConvertTo-PowerShellLiteral $Query))
    Set-Content -LiteralPath $script:runnerPath -Value $runnerContent -Encoding UTF8
    $psExe = Join-Path $PSHOME 'powershell.exe'
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $script:runnerPath + '"'
    try {
        $script:operationType = 'Review' + $Mode
        $script:scanProcess = Start-Process -FilePath $psExe -ArgumentList $arguments -WindowStyle Hidden -PassThru
        $scanButton.Enabled = $false
        $cleanupButton.Enabled = $false
        $reviewButton.Enabled = $false
        $searchButton.Enabled = $false
        $status.ForeColor = $script:palette.Cyan
        $status.Text = if ($Mode -eq 'Largest') { 'Reading the report in the background and finding the largest files...' } else { 'Searching the full report in the background...' }
        $script:timer.Start()
    } catch {
        Remove-Item -LiteralPath $script:runnerPath -Force -ErrorAction SilentlyContinue
        $script:runnerPath = $null
        $script:operationType = $null
        $scanButton.Enabled = $true
        $cleanupButton.Enabled = $true
        $reviewButton.Enabled = $true
        $searchButton.Enabled = $true
        $status.ForeColor = $script:palette.Red
        $status.Text = 'Could not start report review: ' + $_.Exception.Message
    }
}

$typeGrid = New-Grid $typeTab
$fileGrid = New-Grid $fileTab
$tempGrid = New-Grid $tempTab
$fileGrid.Dock = 'None'
$fileGrid.Location = New-Object System.Drawing.Point(7, 49)
$fileGrid.Size = New-Object System.Drawing.Size(1035, 340)
$fileGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
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

foreach ($page in @($typeTab, $fileTab, $tempTab)) {
    $page.BackColor = $script:palette.Panel
    $page.ForeColor = $script:palette.Text
}
foreach ($grid in @($typeGrid, $fileGrid, $tempGrid)) {
    $grid.BackgroundColor = $script:palette.Panel
    $grid.GridColor = $script:palette.Border
    $grid.BorderStyle = 'None'
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:palette.Surface
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:palette.Cyan
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:palette.Surface
    $grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $script:palette.Cyan
    $grid.DefaultCellStyle.BackColor = $script:palette.Panel
    $grid.DefaultCellStyle.ForeColor = $script:palette.Text
    $grid.DefaultCellStyle.SelectionBackColor = $script:palette.Mauve
    $grid.DefaultCellStyle.SelectionForeColor = $script:palette.Background
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:palette.Surface
    $grid.RowTemplate.Height = 28
}
foreach ($page in @($typeTab, $fileTab, $tempTab)) {
    $page.UseVisualStyleBackColor = $false
}
foreach ($control in $form.Controls) {
    if ($control -is [System.Windows.Forms.Label]) {
        if ($control -eq $brandBadge) { continue }
        if ($control -notin @($title, $subtitle, $status)) { $control.ForeColor = $script:palette.Muted }
        $control.BackColor = $script:palette.Background
    } elseif ($control -is [System.Windows.Forms.TextBox] -or $control -is [System.Windows.Forms.CheckedListBox]) {
        $control.BackColor = $script:palette.Surface
        $control.ForeColor = $script:palette.Text
        $control.BorderStyle = 'FixedSingle'
    } elseif ($control -is [System.Windows.Forms.Button]) {
        $control.FlatStyle = 'Flat'
        $control.FlatAppearance.BorderColor = $script:palette.Border
        $control.BackColor = $script:palette.Surface
        $control.ForeColor = $script:palette.Cyan
        $control.Cursor = [System.Windows.Forms.Cursors]::Hand
    }
}
$scanButton.BackColor = $script:palette.Mauve
$scanButton.ForeColor = $script:palette.Background
$scanButton.FlatAppearance.BorderSize = 0
$cleanupButton.BackColor = $script:palette.Blue
$cleanupButton.ForeColor = $script:palette.Background
$deleteButton.BackColor = $script:palette.Red
$deleteButton.ForeColor = $script:palette.Background
$reviewButton.BackColor = $script:palette.Pink
$reviewButton.ForeColor = $script:palette.Background
$reviewButton.FlatAppearance.BorderSize = 0
$openButton.BackColor = $script:palette.Surface
$tabs.BackColor = $script:palette.Panel
$tabs.ForeColor = $script:palette.Text
$tabs.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabs.Add_DrawItem({
    param($sender, $eventArgs)
    $page = $sender.TabPages[$eventArgs.Index]
    $selected = ($sender.SelectedIndex -eq $eventArgs.Index)
    $fill = if ($selected) { $script:palette.Mauve } else { $script:palette.Surface }
    $ink = if ($selected) { $script:palette.Background } else { $script:palette.Muted }
    $fillBrush = [System.Drawing.SolidBrush]::new($fill)
    $textBrush = [System.Drawing.SolidBrush]::new($ink)
    $eventArgs.Graphics.FillRectangle($fillBrush, $eventArgs.Bounds)
    $flags = [System.Drawing.StringFormat]::GenericDefault
    $flags.Alignment = [System.Drawing.StringAlignment]::Center
    $flags.LineAlignment = [System.Drawing.StringAlignment]::Center
    $bounds = $eventArgs.Bounds
    $layout = [System.Drawing.RectangleF]::new(
        [single]$bounds.X, [single]$bounds.Y, [single]$bounds.Width, [single]$bounds.Height
    )
    $eventArgs.Graphics.DrawString(
        [string]$page.Text,
        [System.Drawing.Font]$sender.Font,
        [System.Drawing.Brush]$textBrush,
        [System.Drawing.RectangleF]$layout,
        [System.Drawing.StringFormat]$flags
    )
    $fillBrush.Dispose()
    $textBrush.Dispose()
})

$script:timer = New-Object System.Windows.Forms.Timer
$script:timer.Interval = 800
$script:timer.Add_Tick({
    if ($null -ne $script:scanProcess -and $script:scanProcess.HasExited) {
        $script:timer.Stop()
        $scanExitCode = $script:scanProcess.ExitCode
        $completedOperation = $script:operationType
        $script:scanProcess = $null
        if ($script:runnerPath -and (Test-Path -LiteralPath $script:runnerPath)) {
            Remove-Item -LiteralPath $script:runnerPath -Force -ErrorAction SilentlyContinue
            $script:runnerPath = $null
        }
        if ($scanExitCode -ne 0) {
            $status.Text = if ($completedOperation -eq 'Scan') { 'Scan stopped with an error. Check the selected locations and folder permissions.' } else { 'Report review stopped with an error. Check that the CSV can be read.' }
            $status.ForeColor = $script:palette.Red
            $scanButton.Enabled = $true
            $cleanupButton.Enabled = $true
            $reviewButton.Enabled = $true
            $searchButton.Enabled = $true
            $script:operationType = $null
            return
        }
        if ($completedOperation -eq 'Scan') {
            try {
                $summaryFiles = @(Get-ChildItem -LiteralPath $script:reportDir -Filter 'file-summary-*.csv' -File | Where-Object { $_.LastWriteTime -ge $script:scanStarted } | Sort-Object LastWriteTime -Descending)
                $inventoryFiles = @(Get-ChildItem -LiteralPath $script:reportDir -Filter 'file-inventory-*.csv' -File | Where-Object { $_.LastWriteTime -ge $script:scanStarted } | Sort-Object LastWriteTime -Descending)
                if ($summaryFiles.Count -gt 0 -and $inventoryFiles.Count -gt 0) {
                    $typeGrid.DataSource = @(Import-Csv -LiteralPath $summaryFiles[0].FullName)
                    $script:reviewSourcePath = $inventoryFiles[0].FullName
                    $script:operationType = $null
                    Start-ReportReviewWorker -Mode Largest -SourcePath $script:reviewSourcePath
                    return
                }
                $status.Text = 'Scan ended without complete reports. Check the selected locations and folder permissions.'
                $status.ForeColor = $script:palette.Red
            } catch {
                $status.Text = 'Scan report could not be loaded: ' + $_.Exception.Message
                $status.ForeColor = $script:palette.Red
            }
        } elseif ($completedOperation -like 'Review*') {
            try {
                $rows = @(Import-Csv -LiteralPath $script:reviewResultPath)
                $count = 0
                if (Test-Path -LiteralPath $script:reviewCountPath) { $count = [long](Get-Content -LiteralPath $script:reviewCountPath -Raw) }
                $fileGrid.DataSource = $rows
                $tabs.SelectedTab = $fileTab
                $status.ForeColor = $script:palette.Green
                if ($completedOperation -eq 'ReviewSearch') {
                    $status.Text = ('Found ' + $count + ' matches; showing up to 500. Use Show selected to open its folder.')
                } else {
                    $status.Text = ('Loaded the 500 largest files from ' + $script:reviewSourcePath)
                }
            } catch {
                $status.Text = 'Could not load review results: ' + $_.Exception.Message
                $status.ForeColor = $script:palette.Red
            } finally {
                foreach ($path in @($script:reviewResultPath, $script:reviewCountPath)) {
                    if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
                }
                $script:reviewResultPath = $null
                $script:reviewCountPath = $null
            }
        }
        $scanButton.Enabled = $true
        $cleanupButton.Enabled = $true
        $reviewButton.Enabled = $true
        $searchButton.Enabled = $true
        $script:operationType = $null
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
$reviewButton.Add_Click({
    if ($null -ne $script:scanProcess) { return }
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Choose an Archiva CSV report'
    $dialog.InitialDirectory = if (Test-Path -LiteralPath $outputBox.Text) { $outputBox.Text } else { $script:reportDir }
    $dialog.Filter = 'Archiva reports|file-inventory-*.csv;file-summary-*.csv|All CSV files|*.csv'
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $directory = Split-Path -Parent $dialog.FileName
    $name = [IO.Path]::GetFileName($dialog.FileName)
    if ($name -match '^file-inventory-(.+)\.csv$') {
        $stamp = $Matches[1]
        $inventoryPath = $dialog.FileName
        $summaryPath = Join-Path $directory ('file-summary-' + $stamp + '.csv')
    } elseif ($name -match '^file-summary-(.+)\.csv$') {
        $stamp = $Matches[1]
        $summaryPath = $dialog.FileName
        $inventoryPath = Join-Path $directory ('file-inventory-' + $stamp + '.csv')
    } else {
        [System.Windows.Forms.MessageBox]::Show('Choose a file named file-summary-*.csv or file-inventory-*.csv.', 'Not an Archiva report', 'OK', 'Information') | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show('The matching summary CSV was not found beside this report.', 'Incomplete report set', 'OK', 'Warning') | Out-Null
        return
    }
    $script:reportDir = $directory
    $outputBox.Text = $directory
    $script:reviewSourcePath = if (Test-Path -LiteralPath $inventoryPath -PathType Leaf) { $inventoryPath } else { $null }
    $typeGrid.DataSource = @(Import-Csv -LiteralPath $summaryPath)
    $tabs.SelectedTab = $typeTab
    if ($script:reviewSourcePath) {
        Start-ReportReviewWorker -Mode Largest -SourcePath $script:reviewSourcePath
    } else {
        $status.ForeColor = $script:palette.Red
        $status.Text = 'Summary loaded, but its full inventory CSV is missing. Search and file actions need both files.'
    }
})
$searchButton.Add_Click({
    if ($null -ne $script:scanProcess) { return }
    if (-not $script:reviewSourcePath) {
        [System.Windows.Forms.MessageBox]::Show('Load a saved scan first using review saved CSV.', 'No report loaded', 'OK', 'Information') | Out-Null
        return
    }
    $query = $searchBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($query)) {
        [System.Windows.Forms.MessageBox]::Show('Enter part of a path, file name, extension, or type.', 'Search query needed', 'OK', 'Information') | Out-Null
        return
    }
    Start-ReportReviewWorker -Mode Search -SourcePath $script:reviewSourcePath -Query $query
})
$typeGrid.Add_DoubleClick({
    if ($null -eq $typeGrid.CurrentRow) { return }
    $type = [string]$typeGrid.CurrentRow.Cells['Type'].Value
    if ([string]::IsNullOrWhiteSpace($type)) { return }
    $searchBox.Text = if ($type -eq 'No extension') { $type } else { '.' + $type.ToLowerInvariant() }
    $searchButton.PerformClick()
})
$searchBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $searchButton.PerformClick()
        $_.SuppressKeyPress = $true
    }
})
$revealButton.Add_Click({
    if ($null -eq $fileGrid.CurrentRow) {
        [System.Windows.Forms.MessageBox]::Show('Select a file row first.', 'No file selected', 'OK', 'Information') | Out-Null
        return
    }
    $path = [string]$fileGrid.CurrentRow.Cells['Path'].Value
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show('That file is no longer available at this path.', 'File not found', 'OK', 'Information') | Out-Null
        return
    }
    Start-Process -FilePath explorer.exe -ArgumentList ('/select,"' + $path + '"')
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
    try {
        if ([string]::IsNullOrWhiteSpace($outputBox.Text)) { throw 'Choose a reports folder first.' }
        $reportPath = [IO.Path]::GetFullPath($outputBox.Text)
        New-Item -ItemType Directory -Force -Path $reportPath | Out-Null
    } catch {
        $status.ForeColor = $script:palette.Red
        $status.Text = 'Reports folder is not available: ' + $_.Exception.Message
        return
    }
    $script:reportDir = $reportPath
    $rootValues = ($roots | ForEach-Object { "'" + $_.Replace("'", "''") + "'" }) -join ','
    $scriptPath = (Join-Path $PSScriptRoot 'FileInventory.ps1').Replace("'", "''")
    $outValue = $reportPath.Replace("'", "''")
    $runner = Join-Path $env:TEMP ('file-inventory-scan-' + [guid]::NewGuid().ToString('N') + '.ps1')
    $script:runnerPath = $runner
    $runnerContent = "& '" + $scriptPath + "' -Root @(" + $rootValues + ") -OutputDirectory '" + $outValue + "'"
    Set-Content -LiteralPath $runner -Value $runnerContent -Encoding UTF8
    $script:scanStarted = Get-Date
    $script:operationType = 'Scan'
    $psExe = Join-Path $PSHOME 'powershell.exe'
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $runner + '"'
    try {
        $script:scanProcess = Start-Process -FilePath $psExe -ArgumentList $arguments -WindowStyle Hidden -PassThru
    } catch {
        Remove-Item -LiteralPath $runner -Force -ErrorAction SilentlyContinue
        $script:runnerPath = $null
        $script:operationType = $null
        $scanButton.Enabled = $true
        $cleanupButton.Enabled = $true
        $status.ForeColor = $script:palette.Red
        $status.Text = 'Could not start the scan: ' + $_.Exception.Message
        return
    }
    $scanButton.Enabled = $false
    $cleanupButton.Enabled = $false
    $status.ForeColor = $script:palette.Cyan
    $status.Text = 'Scanning in the background... archive your patience, cutie.'
$script:timer.Start()
})
$cleanupButton.Add_Click({
    $status.Text = 'Looking for temporary files older than seven days...'
    $tempGrid.Rows.Clear()
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
                    $idx = $tempGrid.Rows.Add($false, $item.FullName, [long]$item.Length, $item.LastWriteTime.ToString('g'), 'Old temporary file')
                }
            }
        }
    }
    $deleteButton.Enabled = ($tempGrid.Rows.Count -gt 0)
    $status.Text = ('☆ found ' + $tempGrid.Rows.Count + ' old temp files. pick your targets, then confirm.')
    $tabs.SelectedTab = $tempTab
})
$deleteButton.Add_Click({
    $tempGrid.EndEdit()
    $chosen = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $tempGrid.Rows) { if ($row.Cells['Select'].Value -eq $true) { $chosen.Add($row) } }
    if ($chosen.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select the files you want to delete first.', 'Nothing selected', 'OK', 'Information') | Out-Null
        return
    }
    $total = 0L
    foreach ($row in $chosen) { $total += [long]$row.Cells['SizeBytes'].Value }
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
    foreach ($row in $chosen) {
        $path = [string]$row.Cells['Path'].Value
        $full = [IO.Path]::GetFullPath($path)
        $safe = $false
        foreach ($root in $allowed) {
            if ($full.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { $safe = $true; break }
        }
        if (-not $safe) { continue }
        try {
            $file = Get-Item -LiteralPath $full -Force -ErrorAction Stop
            if (-not $file.PSIsContainer -and (Test-SafeTempFilePath -Path $file.FullName -AllowedRoots $allowed) -and
                $file.Extension -match '^\.(tmp|temp)$' -and $file.LastWriteTime -lt (Get-Date).AddDays(-7)) {
                Remove-Item -LiteralPath $full -Force -ErrorAction Stop
                $deleted++
            }
        } catch { }
    }
    $cleanupButton.PerformClick()
    $status.Text = ('Deleted ' + $deleted + ' selected files. Remaining candidates are still listed for review.')
})

$form.Add_FormClosing({
    if ($null -ne $script:scanProcess -and -not $script:scanProcess.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            'A scan or report review is still running. Stop it and close Archiva?',
            'Operation in progress', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $_.Cancel = $true
            return
        }
        Stop-Process -Id $script:scanProcess.Id -Force -ErrorAction SilentlyContinue
        $script:scanProcess = $null
        if ($script:runnerPath) { Remove-Item -LiteralPath $script:runnerPath -Force -ErrorAction SilentlyContinue }
        $script:timer.Stop()
    }
})

[void]$form.ShowDialog()
