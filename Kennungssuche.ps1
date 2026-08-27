Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:CsvPath = Join-Path -Path $PSScriptRoot -ChildPath 'kennungen.csv'
$script:Entries = @()

function Import-EntriesCsv {
    $bytes = [System.IO.File]::ReadAllBytes($script:CsvPath)
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        $content = $utf8.GetString($bytes)
    }
    catch [System.Text.DecoderFallbackException] {
        $content = [System.Text.Encoding]::Default.GetString($bytes)
    }

    if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) {
        $content = $content.Substring(1)
    }

    return $content | ConvertFrom-Csv -Delimiter ';'
}

function Load-Entries {
    if (-not (Test-Path -LiteralPath $script:CsvPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Die Datei kennungen.csv wurde nicht gefunden.",
            'Kennungssuche',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }

    try {
        $script:Entries = @(Import-EntriesCsv | Where-Object {
            $_.Kennung -and $_.Vorname -and $_.Nachname
        })
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Die CSV-Datei konnte nicht gelesen werden:`n$($_.Exception.Message)",
            'Kennungssuche',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

function Get-FilteredEntries {
    $term = $searchBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($term)) {
        return $script:Entries
    }

    return @($script:Entries | Where-Object {
        $fullName = "$($_.Vorname) $($_.Nachname)"
        @($_.Kennung, $_.Vorname, $_.Nachname, $fullName) | Where-Object {
            $_.IndexOf($term, [System.StringComparison]::CurrentCultureIgnoreCase) -ge 0
        }
    })
}

function ConvertTo-CheckboxValue($value) {
    if ($value -is [bool]) {
        return $value
    }

    return [string]$value -match '^(1|true|yes|ja)$'
}

function Save-Entries {
    $checkedByKennung = @{}
    foreach ($row in $grid.Rows) {
        $checkedByKennung[$row.Cells['Kennung'].Value] = [pscustomobject]@{
            Handy = ConvertTo-CheckboxValue $row.Cells['Handy'].Value
            Recherche = ConvertTo-CheckboxValue $row.Cells['Recherche'].Value
            'AZR/LMR' = ConvertTo-CheckboxValue $row.Cells['AZR/LMR'].Value
            KBA = ConvertTo-CheckboxValue $row.Cells['KBA'].Value
            SARS = ConvertTo-CheckboxValue $row.Cells['SARS'].Value
            'USB MIK' = ConvertTo-CheckboxValue $row.Cells['USB MIK'].Value
            'USB VS' = ConvertTo-CheckboxValue $row.Cells['USB VS'].Value
        }
    }

    $rows = foreach ($entry in $script:Entries) {
        $checks = $checkedByKennung[$entry.Kennung]
        [pscustomobject]@{
            Kennung = $entry.Kennung
            Vorname = $entry.Vorname
            Nachname = $entry.Nachname
            Handy = if ($null -ne $checks) { $checks.Handy } else { ConvertTo-CheckboxValue $entry.Handy }
            Recherche = if ($null -ne $checks) { $checks.Recherche } else { ConvertTo-CheckboxValue $entry.Recherche }
            SARS = if ($null -ne $checks) { $checks.SARS } else { ConvertTo-CheckboxValue $entry.SARS }
            KBA = if ($null -ne $checks) { $checks.KBA } else { ConvertTo-CheckboxValue $entry.KBA }
            'AZR/LMR' = if ($null -ne $checks) { $checks.'AZR/LMR' } else { ConvertTo-CheckboxValue $entry.'AZR/LMR' }
            'USB MIK' = if ($null -ne $checks) { $checks.'USB MIK' } else { ConvertTo-CheckboxValue $entry.'USB MIK' }
            'USB VS' = if ($null -ne $checks) { $checks.'USB VS' } else { ConvertTo-CheckboxValue $entry.'USB VS' }
        }
    }

    try {
        $csv = $rows | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
        [System.IO.File]::WriteAllLines($script:CsvPath, $csv, (New-Object System.Text.UTF8Encoding($false)))
        $statusLabel.Text = 'Änderungen gespeichert'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Die Änderungen konnten nicht gespeichert werden:`n$($_.Exception.Message)",
            'Kennungssuche',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Update-Grid {
    $grid.Rows.Clear()
    $matches = @(Get-FilteredEntries)

    foreach ($entry in $matches) {
        [void]$grid.Rows.Add(
            $entry.Kennung,
            $entry.Vorname,
            $entry.Nachname,
            (ConvertTo-CheckboxValue $entry.Handy),
            (ConvertTo-CheckboxValue $entry.Recherche),
            (ConvertTo-CheckboxValue $entry.SARS),
            (ConvertTo-CheckboxValue $entry.KBA),
            (ConvertTo-CheckboxValue $entry.'AZR/LMR'),
            (ConvertTo-CheckboxValue $entry.'USB MIK'),
            (ConvertTo-CheckboxValue $entry.'USB VS')
        )
    }

    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $statusLabel.Text = "$($script:Entries.Count) Einträge geladen"
    }
    else {
        $statusLabel.Text = "$($matches.Count) Treffer für '$($searchBox.Text.Trim())'"
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Kennungssuche'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(850, 570)
$form.MinimumSize = New-Object System.Drawing.Size(700, 450)
$form.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 251)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Kennungssuche'
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 22)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 45, 75)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(28, 22)
$form.Controls.Add($titleLabel)

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = 'Nach Vorname, Nachname, vollständigem Namen oder Kennung suchen'
$hintLabel.ForeColor = [System.Drawing.Color]::FromArgb(90, 105, 125)
$hintLabel.AutoSize = $true
$hintLabel.Location = New-Object System.Drawing.Point(31, 62)
$form.Controls.Add($hintLabel)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(30, 96)
$searchBox.Size = New-Object System.Drawing.Size(535, 32)
$searchBox.Font = New-Object System.Drawing.Font('Segoe UI', 12)
$searchBox.Anchor = 'Top,Left,Right'
$form.Controls.Add($searchBox)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = 'Leeren'
$clearButton.Location = New-Object System.Drawing.Point(575, 95)
$clearButton.Size = New-Object System.Drawing.Size(105, 34)
$clearButton.Anchor = 'Top,Right'
$clearButton.FlatStyle = 'Flat'
$clearButton.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($clearButton)

$reloadButton = New-Object System.Windows.Forms.Button
$reloadButton.Text = 'Daten neu laden'
$reloadButton.Location = New-Object System.Drawing.Point(690, 95)
$reloadButton.Size = New-Object System.Drawing.Size(125, 34)
$reloadButton.Anchor = 'Top,Right'
$reloadButton.FlatStyle = 'Flat'
$reloadButton.BackColor = [System.Drawing.Color]::FromArgb(25, 107, 146)
$reloadButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($reloadButton)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(30, 150)
$grid.Size = New-Object System.Drawing.Size(785, 330)
$grid.Anchor = 'Top,Bottom,Left,Right'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.ReadOnly = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'None'
$grid.GridColor = [System.Drawing.Color]::FromArgb(225, 230, 236)
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(25, 45, 75)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$grid.EnableHeadersVisualStyles = $false
[void]$grid.Columns.Add('Kennung', 'Kennung')
[void]$grid.Columns.Add('Vorname', 'Vorname')
[void]$grid.Columns.Add('Nachname', 'Nachname')
$handyColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$handyColumn.Name = 'Handy'
$handyColumn.HeaderText = 'Handy'
$handyColumn.TrueValue = $true
$handyColumn.FalseValue = $false
$handyColumn.Width = 105
[void]$grid.Columns.Add($handyColumn)
$rechercheColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$rechercheColumn.Name = 'Recherche'
$rechercheColumn.HeaderText = 'Recherche'
$rechercheColumn.TrueValue = $true
$rechercheColumn.FalseValue = $false
$rechercheColumn.Width = 125
[void]$grid.Columns.Add($rechercheColumn)
$categoryColumns = @(
    @{ Name = 'SARS'; Width = 75 },
    @{ Name = 'KBA'; Width = 75 },
    @{ Name = 'AZR/LMR'; Width = 75 },
    @{ Name = 'USB MIK'; Width = 75 },
    @{ Name = 'USB VS'; Width = 75 }
)
foreach ($category in $categoryColumns) {
    if ($category.Name -in @('USB MIK', 'USB VS')) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.ReadOnly = $true
    }
    else {
        $column = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    }
    $column.Name = $category.Name
    $column.HeaderText = $category.Name
    if ($column -is [System.Windows.Forms.DataGridViewCheckBoxColumn]) {
        $column.TrueValue = $true
        $column.FalseValue = $false
    }
    $column.Width = $category.Width
    [void]$grid.Columns.Add($column)
}
foreach ($column in $grid.Columns) {
    $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
}
$form.Controls.Add($grid)

$grid.Add_CellFormatting({
    param($sender, $eventArgs)

    if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -lt 0) {
        return
    }

    $columnName = $sender.Columns[$eventArgs.ColumnIndex].Name
    if ($columnName -notin @('USB MIK', 'USB VS')) {
        return
    }

    $isAvailable = ConvertTo-CheckboxValue $eventArgs.Value
    $eventArgs.CellStyle.BackColor = if ($isAvailable) {
        [System.Drawing.Color]::FromArgb(76, 175, 80)
    }
    else {
        [System.Drawing.Color]::FromArgb(220, 70, 70)
    }
    $eventArgs.CellStyle.ForeColor = [System.Drawing.Color]::White
    $eventArgs.CellStyle.SelectionBackColor = $eventArgs.CellStyle.BackColor
    $eventArgs.CellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $eventArgs.Value = ''
    $eventArgs.FormattingApplied = $true
})

$grid.Add_CellClick({
    param($sender, $eventArgs)

    if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -lt 0) {
        return
    }

    $columnName = $sender.Columns[$eventArgs.ColumnIndex].Name
    if ($columnName -in @('USB MIK', 'USB VS')) {
        $cell = $sender.Rows[$eventArgs.RowIndex].Cells[$eventArgs.ColumnIndex]
        $cell.Value = -not (ConvertTo-CheckboxValue $cell.Value)
        $sender.InvalidateCell($cell)
    }
})

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = 'Änderungen speichern'
$saveButton.Location = New-Object System.Drawing.Point(650, 492)
$saveButton.Size = New-Object System.Drawing.Size(165, 34)
$saveButton.Anchor = 'Bottom,Right'
$saveButton.FlatStyle = 'Flat'
$saveButton.BackColor = [System.Drawing.Color]::FromArgb(25, 107, 146)
$saveButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($saveButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(90, 105, 125)
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(31, 498)
$statusLabel.Anchor = 'Bottom,Left'
$form.Controls.Add($statusLabel)

$searchBox.Add_TextChanged({ Update-Grid })
$clearButton.Add_Click({ $searchBox.Clear(); $searchBox.Focus() })
$saveButton.Add_Click({ Save-Entries })
$reloadButton.Add_Click({
    if (Load-Entries) {
        Update-Grid
    }
})
$form.Add_Shown({
    $searchBox.Focus()
    if (Load-Entries) {
        Update-Grid
    }
})

[void]$form.ShowDialog()
