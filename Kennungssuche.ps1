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

function Update-Grid {
    $grid.Rows.Clear()
    $matches = @(Get-FilteredEntries)

    foreach ($entry in $matches) {
        [void]$grid.Rows.Add($entry.Kennung, $entry.Vorname, $entry.Nachname, "$($entry.Vorname) $($entry.Nachname)")
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
$grid.ReadOnly = $true
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
$grid.EnableHeadersVisualStyles = $false
[void]$grid.Columns.Add('Kennung', 'Kennung')
[void]$grid.Columns.Add('Vorname', 'Vorname')
[void]$grid.Columns.Add('Nachname', 'Nachname')
[void]$grid.Columns.Add('Vollname', 'Vollständiger Name')
$form.Controls.Add($grid)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(90, 105, 125)
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(31, 498)
$statusLabel.Anchor = 'Bottom,Left'
$form.Controls.Add($statusLabel)

$searchBox.Add_TextChanged({ Update-Grid })
$clearButton.Add_Click({ $searchBox.Clear(); $searchBox.Focus() })
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
