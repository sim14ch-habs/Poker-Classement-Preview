param(
    [string]$WorkbookPath = ".\\Poker Stanley Hiver 2026.xlsx",
    [string]$ImagePath = "",
    [string]$OutputDir = "",
    [string]$DataDir = "",
    [string]$Model = "gpt-5-mini",
    [string]$EmailTo = "sim_621@hotmail.com",
    [string]$PublicUrl = "",
    [string]$PublishRepoDir = "",
    [switch]$NoOcr,
    [switch]$SkipExport
)
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PSCommandPath) -or $MyInvocation.MyCommand.CommandType -ne 'ExternalScript') {
    throw "Ce script doit etre execute depuis un fichier .ps1 (PowerShell -File), pas en commande inline."
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ValidationMutex = New-Object System.Threading.Mutex($false, 'Local\PokerValidationPhotoUi')
$script:ValidationMutexAcquired = $false
try {
    $script:ValidationMutexAcquired = $script:ValidationMutex.WaitOne(0, $false)
}
catch {
    # If mutex state cannot be read, continue rather than blocking the workflow.
    $script:ValidationMutexAcquired = $true
}

if (-not $script:ValidationMutexAcquired) {
    [System.Windows.Forms.MessageBox]::Show('Une autre fenetre de validation est deja ouverte.', 'Validation Poker', 'OK', 'Information') | Out-Null
    Write-Output 'Une autre instance de validation est deja en cours.'
    exit 2
}

function Release-ComObjectSafe {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) } catch {}
    }
}

function Normalize-Plain {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $norm = $Text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().ToLowerInvariant().Trim()
}

function Get-DateFromPhotoName {
    param([string]$FileName)
    $base = [IO.Path]::GetFileNameWithoutExtension($FileName)
    if ($base -notmatch '^\s*(\d{1,2})\s+([\p{L}\-]+)\s+(\d{4})\s*$') { return $null }

    $day = [int]$matches[1]
    $monthRaw = Normalize-Plain $matches[2]
    $year = [int]$matches[3]

    $monthMap = @{
        "janvier" = 1; "fevrier" = 2; "mars" = 3; "avril" = 4; "mai" = 5; "juin" = 6
        "juillet" = 7; "aout" = 8; "septembre" = 9; "octobre" = 10; "novembre" = 11; "decembre" = 12
    }

    if (-not $monthMap.ContainsKey($monthRaw)) { return $null }
    try { return [datetime]::new($year, [int]$monthMap[$monthRaw], $day) } catch { return $null }
}

function Get-LatestDatedPhoto {
    param([string]$Folder)

    $files = Get-ChildItem -LiteralPath $Folder -File |
        Where-Object {
            $_.Extension.ToLowerInvariant() -in @('.jpg', '.jpeg', '.png') -and
            $_.Name -notlike 'Classement semaine *.png' -and
            $_.Name -notlike '*_table.png' -and
            $_.Name -notlike '*_upscaled.png'
        }

    $dated = foreach ($f in $files) {
        $d = Get-DateFromPhotoName -FileName $f.Name
        if ($null -ne $d) { [pscustomobject]@{ File = $f; Date = $d } }
    }

    if (-not $dated) { return $null }
    return $dated | Sort-Object Date, @{ Expression = { $_.File.LastWriteTime } } | Select-Object -Last 1
}

function Get-MimeType {
    param([string]$Path)
    switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.jpg' { return 'image/jpeg' }
        '.jpeg' { return 'image/jpeg' }
        '.png' { return 'image/png' }
        default { throw "Unsupported image extension: $Path" }
    }
}

function To-DataUrl {
    param([string]$Path)
    $mime = Get-MimeType -Path $Path
    $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Path))
    return "data:$mime;base64,$b64"
}

function Invoke-PhotoOcr {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$ImagePath,
        [datetime]$ExpectedDate,
        [string[]]$Players,
        [int]$MaxPositions
    )

    $dataUrl = To-DataUrl -Path $ImagePath

    $systemPrompt = @(
        'You read handwritten weekly poker ranking sheets.',
        'Return ONLY a valid JSON object with this shape:',
        '{ "rows": [ { "position": integer, "registered_player": string|null, "is_replacement": boolean, "replacement_name": string|null, "kills": integer, "confidence": number } ], "notes": [string] }',
        'Rules:',
        '- Joueurs column is official player.',
        '- Remplacants column means replacement for official player.',
        '- If replacement_name exists then is_replacement=true.',
        '- If is_replacement=true then kills must be 0 because replacement kills do not count for the registered player.',
        '- kills is last column value (blank -> 0).',
        '- Do not invent names outside provided list.'
    ) -join "`n"

    $userPrompt = @(
        "Expected photo date: $($ExpectedDate.ToString('yyyy-MM-dd'))",
        "Expected max positions: $MaxPositions",
        '',
        'Official players:',
        ($Players -join "`n")
    ) -join "`n"

    $payload = @{
        model = $Model
        response_format = @{ type = "json_object" }
        messages = @(
            @{ role = "system"; content = $systemPrompt },
            @{ role = "user"; content = @(
                @{ type = "text"; text = $userPrompt },
                @{ type = "image_url"; image_url = @{ url = $dataUrl; detail = "high" } }
            ) }
        )
    }

    $headers = @{ Authorization = "Bearer $ApiKey"; "Content-Type" = "application/json" }
    $body = $payload | ConvertTo-Json -Depth 20 -Compress
    $resp = Invoke-RestMethod -Method POST -Uri "https://api.openai.com/v1/chat/completions" -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($body))

    $content = [string]$resp.choices[0].message.content
    if ([string]::IsNullOrWhiteSpace($content)) { throw "Empty OCR response" }

    $clean = $content.Trim()
    $fence = [string]([char]96) + [string]([char]96) + [string]([char]96)
    if ($clean.StartsWith($fence)) {
        $lines = $clean -split "`r?`n"
        if ($lines.Count -ge 2) { $clean = ($lines[1..($lines.Count - 1)] -join "`n") }
        if ($clean.EndsWith($fence)) { $clean = $clean.Substring(0, $clean.Length - $fence.Length) }
        $clean = $clean.Trim()
    }

    try { $obj = $clean | ConvertFrom-Json } catch { throw "Invalid OCR JSON: $clean" }
    if (-not ($obj.PSObject.Properties.Name -contains 'rows')) { $obj | Add-Member -NotePropertyName rows -NotePropertyValue @() -Force }
    if (-not ($obj.PSObject.Properties.Name -contains 'notes')) { $obj | Add-Member -NotePropertyName notes -NotePropertyValue @() -Force }
    return $obj
}
function Get-WorkbookContext {
    param(
        [string]$WorkbookPath,
        [datetime]$PhotoDate
    )

    $xlUp = -4162
    $excel = $null
    $wb = $null
    $wsHebdo = $null
    $wsKills = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($WorkbookPath)
        $wsHebdo = $wb.Worksheets.Item("Classement Hebdo")
        $wsKills = $wb.Worksheets.Item("Kills")

        $lastRowHebdo = $wsHebdo.Cells($wsHebdo.Rows.Count, 1).End($xlUp).Row
        $players = New-Object System.Collections.Generic.List[string]
        $hebdoRows = @{}

        for ($r = 3; $r -le $lastRowHebdo; $r++) {
            $name = [string]$wsHebdo.Cells($r, 1).Text
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $name = $name.Trim()
            [void]$players.Add($name)
            $hebdoRows[$name] = $r
        }

        $lastRowKills = $wsKills.Cells($wsKills.Rows.Count, 1).End($xlUp).Row
        $killsRows = @{}
        for ($r = 3; $r -le $lastRowKills; $r++) {
            $name = [string]$wsKills.Cells($r, 1).Text
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $killsRows[$name.Trim()] = $r
        }

        $weekCol = $null
        for ($c = 2; $c -le 15; $c++) {
            $v = $wsHebdo.Cells(1, $c).Value2
            if ($null -eq $v -or "$v" -eq "") { continue }
            $d = [datetime]::FromOADate([double]$v).Date
            if ($d -eq $PhotoDate.Date) { $weekCol = $c; break }
        }
        if ($null -eq $weekCol) {
            throw "No week column found in Classement Hebdo for date $($PhotoDate.ToString('yyyy-MM-dd'))"
        }

        return [pscustomobject]@{
            Players = @($players)
            PlayerRowsHebdo = $hebdoRows
            PlayerRowsKills = $killsRows
            WeekCol = [int]$weekCol
            WeekNum = [int]$weekCol - 1
        }
    }
    finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObjectSafe $wsKills
        Release-ComObjectSafe $wsHebdo
        Release-ComObjectSafe $wb
        Release-ComObjectSafe $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Build-InitialRows {
    param(
        [string[]]$Players,
        [object]$OcrResult
    )

    $rows = New-Object System.Collections.Generic.List[object]
    for ($pos = 1; $pos -le $Players.Count; $pos++) {
        [void]$rows.Add([pscustomobject]@{
            Position = $pos
            Player = ""
            IsAbsent = $false
            IsReplacement = $false
            ReplacementName = ""
            Kills = 0
            Confidence = ""
        })
    }

    $candidates = New-Object System.Collections.Generic.List[object]

    if ($null -ne $OcrResult) {
        foreach ($r in @($OcrResult.rows)) {
            $pos = 0
            try { $pos = [int]$r.position } catch { continue }
            if ($pos -lt 1 -or $pos -gt $Players.Count) { continue }

            $player = [string]$r.registered_player
            if ([string]::IsNullOrWhiteSpace($player)) { continue }
            if ($Players -notcontains $player) { continue }

            $isRep = $false
            try { $isRep = [bool]$r.is_replacement } catch {}

            $repName = [string]$r.replacement_name
            if (-not [string]::IsNullOrWhiteSpace($repName)) { $isRep = $true }

            $kills = 0
            try { $kills = [int]$r.kills } catch { $kills = 0 }
            if ($kills -lt 0) { $kills = 0 }
            if ($isRep) { $kills = 0 }

            $confVal = -1.0
            try { $confVal = [double]$r.confidence } catch { $confVal = -1.0 }
            $confTxt = ""
            if ($confVal -ge 0) { $confTxt = $confVal.ToString('0.00') }

            [void]$candidates.Add([pscustomobject]@{
                Position = $pos
                Player = $player
                IsReplacement = $isRep
                ReplacementName = $repName
                Kills = $kills
                Confidence = $confTxt
                ConfidenceValue = $confVal
            })
        }
    }

    $assignedPlayers = @{}
    $assignedPositions = @{}

    foreach ($cand in ($candidates | Sort-Object @{ Expression = { $_.ConfidenceValue }; Descending = $true }, @{ Expression = { $_.Position }; Ascending = $true })) {
        $pos = [int]$cand.Position
        $player = [string]$cand.Player

        if ($assignedPositions.ContainsKey($pos)) { continue }
        if ($assignedPlayers.ContainsKey($player)) { continue }

        $idx = $pos - 1
        $rows[$idx].Player = $player
        $rows[$idx].IsAbsent = $false
        $rows[$idx].IsReplacement = [bool]$cand.IsReplacement
        $rows[$idx].ReplacementName = [string]$cand.ReplacementName
        $rows[$idx].Kills = [int]$cand.Kills
        $rows[$idx].Confidence = [string]$cand.Confidence

        $assignedPositions[$pos] = $true
        $assignedPlayers[$player] = $true
    }

    $remainingPlayers = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Players) {
        if (-not $assignedPlayers.ContainsKey($p)) {
            [void]$remainingPlayers.Add($p)
        }
    }

    $nextRemaining = 0
    for ($idx = 0; $idx -lt $rows.Count; $idx++) {
        if (-not [string]::IsNullOrWhiteSpace([string]$rows[$idx].Player)) { continue }
        if ($nextRemaining -ge $remainingPlayers.Count) { break }

        $rows[$idx].Player = $remainingPlayers[$nextRemaining]
        $rows[$idx].IsAbsent = $false
        $rows[$idx].IsReplacement = $false
        $rows[$idx].ReplacementName = ""
        $rows[$idx].Kills = 0
        $rows[$idx].Confidence = "deduction"
        $nextRemaining++
    }

    return $rows.ToArray()
}

function Is-LowConfidence {
    param([string]$ConfidenceText)

    if ([string]::IsNullOrWhiteSpace($ConfidenceText)) { return $true }
    if ($ConfidenceText.Trim().ToLowerInvariant() -eq 'deduction') { return $true }

    $v = 0.0
    if ([double]::TryParse($ConfidenceText, [ref]$v)) {
        return ($v -lt 0.75)
    }

    return $true
}

function Swap-RowData {
    param(
        [System.Data.DataTable]$Table,
        [int]$A,
        [int]$B
    )
    $cols = @('Player', 'IsAbsent', 'IsReplacement', 'ReplacementName', 'Kills', 'Confidence')
    foreach ($c in $cols) {
        $tmp = $Table.Rows[$A][$c]
        $Table.Rows[$A][$c] = $Table.Rows[$B][$c]
        $Table.Rows[$B][$c] = $tmp
    }
}
function Show-ValidationForm {
    param(
        [string]$ImagePath,
        [string[]]$Players,
        [object[]]$Rows,
        [int]$WeekNum,
        [string]$DateKey,
        [string]$OcrStatus,
        [string]$Notes
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Validation classement semaine $WeekNum ($DateKey)"
    $form.StartPosition = 'CenterScreen'
    $form.Width = 1600
    $form.Height = 980
    $form.MinimumSize = New-Object System.Drawing.Size(1200, 760)
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    $split.Orientation = 'Vertical'
    $split.SplitterDistance = 620

    $imgPanel = New-Object System.Windows.Forms.Panel
    $imgPanel.Dock = 'Fill'

    $imgLabel = New-Object System.Windows.Forms.Label
    $imgLabel.Text = 'Photo'
    $imgLabel.Dock = 'Top'
    $imgLabel.Height = 28
    $imgLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

    $pb = New-Object System.Windows.Forms.PictureBox
    $pb.Dock = 'Fill'
    $pb.SizeMode = 'Zoom'
    $pb.Image = [System.Drawing.Image]::FromFile($ImagePath)

    $imgPanel.Controls.Add($pb)
    $imgPanel.Controls.Add($imgLabel)
    $split.Panel1.Controls.Add($imgPanel)

    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = 'Fill'

    $topInfo = New-Object System.Windows.Forms.Label
    $topInfo.Dock = 'Top'
    $topInfo.Height = 64
    $topInfo.Text = "OCR: $OcrStatus`r`nNotes: $Notes"
    $topInfo.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $instruction = New-Object System.Windows.Forms.Label
    $instruction.Dock = 'Top'
    $instruction.Height = 24
    $instruction.Text = 'Swap: click one row header then another row header. Auto-swap keeps unique names. Use Absent for ABS players. Replacement rows must keep Kills at 0. Yellow rows = low confidence, verify first.'

    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Position', [int])
    [void]$table.Columns.Add('Player', [string])
    [void]$table.Columns.Add('IsAbsent', [bool])
    [void]$table.Columns.Add('IsReplacement', [bool])
    [void]$table.Columns.Add('ReplacementName', [string])
    [void]$table.Columns.Add('Kills', [int])
    [void]$table.Columns.Add('Confidence', [string])

    foreach ($r in $Rows) {
        $dr = $table.NewRow()
        $dr['Position'] = [int]$r.Position
        $dr['Player'] = [string]$r.Player
        $dr['IsAbsent'] = [bool]$r.IsAbsent
        $dr['IsReplacement'] = [bool]$r.IsReplacement
        $dr['ReplacementName'] = [string]$r.ReplacementName
        $dr['Kills'] = [int]$r.Kills
        $dr['Confidence'] = [string]$r.Confidence
        [void]$table.Rows.Add($dr)
    }

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AutoGenerateColumns = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'

    $cPos = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cPos.DataPropertyName = 'Position'
    $cPos.HeaderText = 'Pos'
    $cPos.Width = 55
    $cPos.ReadOnly = $true

    $cPlayer = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $cPlayer.DataPropertyName = 'Player'
    $cPlayer.HeaderText = 'Joueur'
    $cPlayer.Width = 240
    $cPlayer.FlatStyle = 'Flat'
    [void]$cPlayer.Items.Add('')
    foreach ($p in $Players) { [void]$cPlayer.Items.Add($p) }

    $cAbs = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $cAbs.DataPropertyName = 'IsAbsent'
    $cAbs.HeaderText = 'Absent'
    $cAbs.Width = 70

    $cRep = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $cRep.DataPropertyName = 'IsReplacement'
    $cRep.HeaderText = 'Remplacant'
    $cRep.Width = 90

    $cRepName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cRepName.DataPropertyName = 'ReplacementName'
    $cRepName.HeaderText = 'Nom remplacant'
    $cRepName.Width = 170

    $cKills = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cKills.DataPropertyName = 'Kills'
    $cKills.HeaderText = 'Kills'
    $cKills.Width = 65

    $cConf = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cConf.DataPropertyName = 'Confidence'
    $cConf.HeaderText = 'Conf'
    $cConf.Width = 60
    $cConf.ReadOnly = $true

    [void]$grid.Columns.Add($cPos)
    [void]$grid.Columns.Add($cPlayer)
    [void]$grid.Columns.Add($cAbs)
    [void]$grid.Columns.Add($cRep)
    [void]$grid.Columns.Add($cRepName)
    [void]$grid.Columns.Add($cKills)
    [void]$grid.Columns.Add($cConf)

    $grid.DataSource = $table

    $lowConfidenceColor = [System.Drawing.Color]::FromArgb(255, 249, 196)
    $normalColor = [System.Drawing.Color]::White

    $applyConfidenceColors = {
        foreach ($gridRow in $grid.Rows) {
            if ($gridRow.IsNewRow) { continue }

            $gridRow.DefaultCellStyle.BackColor = $normalColor
            $gridRow.Cells[1].Style.BackColor = $normalColor
            $gridRow.Cells[1].Style.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
            $gridRow.Cells[1].Style.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText

            $confText = [string]$gridRow.Cells[6].Value
            if (Is-LowConfidence -ConfidenceText $confText) {
                $gridRow.Cells[1].Style.BackColor = $lowConfidenceColor
                $gridRow.Cells[1].Style.SelectionBackColor = $lowConfidenceColor
                $gridRow.Cells[1].Style.SelectionForeColor = [System.Drawing.Color]::Black
            }
        }
    }

    $grid.add_DataBindingComplete({
        param($sender, $e)
        & $applyConfidenceColors
    })

    $grid.add_CurrentCellDirtyStateChanged({
        param($sender, $e)
        if ($sender.IsCurrentCellDirty) {
            [void]$sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    $playerColIndex = 1
    $absentColIndex = 2
    $replacementColIndex = 3
    $replacementNameColIndex = 4
    $killsColIndex = 5
    $editedPlayerRow = -1
    $editedPlayerOldValue = ""
    $suppressAutoPlayerSwap = $false

    $grid.add_CellBeginEdit({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }

        if ($e.ColumnIndex -eq $killsColIndex) {
            $isAbsent = [bool]$sender.Rows[$e.RowIndex].Cells[$absentColIndex].Value
            $isReplacement = [bool]$sender.Rows[$e.RowIndex].Cells[$replacementColIndex].Value
            if ($isAbsent -or $isReplacement) {
                $sender.Rows[$e.RowIndex].Cells[$killsColIndex].Value = 0
                if ($null -ne $status) {
                    $status.Text = "Kills forces a 0 pour ABS/remplacement (ligne $($e.RowIndex + 1))"
                }
                $e.Cancel = $true
                return
            }
        }

        if ($e.ColumnIndex -ne $playerColIndex) { return }

        $editedPlayerRow = $e.RowIndex
        $editedPlayerOldValue = [string]$sender.Rows[$e.RowIndex].Cells[$playerColIndex].Value
    })

    $grid.add_CellValueChanged({
        param($sender, $e)

        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq $playerColIndex -and -not $suppressAutoPlayerSwap) {
            $newPlayer = [string]$sender.Rows[$e.RowIndex].Cells[$playerColIndex].Value
            $oldPlayer = ""
            if ($editedPlayerRow -eq $e.RowIndex) { $oldPlayer = $editedPlayerOldValue }

            if (-not [string]::IsNullOrWhiteSpace($newPlayer) -and $newPlayer -ne $oldPlayer) {
                $otherRow = -1
                foreach ($row in $sender.Rows) {
                    if ($row.IsNewRow -or $row.Index -eq $e.RowIndex) { continue }
                    $candidate = [string]$row.Cells[$playerColIndex].Value
                    if ($candidate -eq $newPlayer) {
                        $otherRow = $row.Index
                        break
                    }
                }

                if ($otherRow -ge 0) {
                    $suppressAutoPlayerSwap = $true
                    try {
                        $sender.Rows[$otherRow].Cells[$playerColIndex].Value = $oldPlayer
                        if ($null -ne $status) {
                            $status.Text = "Echange auto des noms: ligne $($e.RowIndex + 1) <-> ligne $($otherRow + 1)"
                        }
                    }
                    finally {
                        $suppressAutoPlayerSwap = $false
                    }
                }
            }

            $editedPlayerRow = -1
            $editedPlayerOldValue = ""
        }


        if ($e.RowIndex -ge 0 -and -not $suppressAutoPlayerSwap) {
            if ($e.ColumnIndex -eq $absentColIndex) {
                $isAbsent = [bool]$sender.Rows[$e.RowIndex].Cells[$absentColIndex].Value
                if ($isAbsent) {
                    $suppressAutoPlayerSwap = $true
                    try {
                        $sender.Rows[$e.RowIndex].Cells[$replacementColIndex].Value = $false
                        $sender.Rows[$e.RowIndex].Cells[$replacementNameColIndex].Value = ''
                        $sender.Rows[$e.RowIndex].Cells[$killsColIndex].Value = 0
                        if ($null -ne $status) {
                            $status.Text = "Ligne $($e.RowIndex + 1) marquee ABS"
                        }
                    }
                    finally {
                        $suppressAutoPlayerSwap = $false
                    }
                }
            }
            elseif ($e.ColumnIndex -eq $replacementColIndex) {
                $isReplacement = [bool]$sender.Rows[$e.RowIndex].Cells[$replacementColIndex].Value
                if ($isReplacement) {
                    $suppressAutoPlayerSwap = $true
                    try {
                        $sender.Rows[$e.RowIndex].Cells[$absentColIndex].Value = $false
                        $sender.Rows[$e.RowIndex].Cells[$killsColIndex].Value = 0
                        if ($null -ne $status) {
                            $status.Text = "Kills retires pour le remplacement (ligne $($e.RowIndex + 1))"
                        }
                    }
                    finally {
                        $suppressAutoPlayerSwap = $false
                    }
                }
            }
        }

        & $applyConfidenceColors
    })

    & $applyConfidenceColors

    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Dock = 'Bottom'
    $bottom.Height = 54
    $bottom.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 8)

    $status = New-Object System.Windows.Forms.Label
    $status.Dock = 'Fill'
    $status.Text = 'Pret'

    $btnApprove = New-Object System.Windows.Forms.Button
    $btnApprove.Text = 'Appliquer (Approve and apply)'
    $btnApprove.Width = 220
    $btnApprove.Dock = 'Right'

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Annuler'
    $btnCancel.Width = 100
    $btnCancel.Dock = 'Right'

    $bottom.Controls.Add($status)
    $bottom.Controls.Add($btnCancel)
    $bottom.Controls.Add($btnApprove)

    $rightPanel.Controls.Add($grid)
    $rightPanel.Controls.Add($instruction)
    $rightPanel.Controls.Add($topInfo)
    $rightPanel.Controls.Add($bottom)
    $bottom.BringToFront()
    $split.Panel2.Controls.Add($rightPanel)

    $form.Controls.Add($split)
    $form.AcceptButton = $btnApprove
    $form.CancelButton = $btnCancel

    $swapSource = -1
    $grid.add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }

        if ($swapSource -lt 0) {
            $swapSource = $e.RowIndex
            $status.Text = "Swap source selected: row $($swapSource + 1). Click target row header."
            return
        }

        if ($swapSource -eq $e.RowIndex) {
            $swapSource = -1
            $status.Text = 'Swap canceled'
            return
        }

        Swap-RowData -Table $table -A $swapSource -B $e.RowIndex
        $status.Text = "Swapped row $($swapSource + 1) with row $($e.RowIndex + 1)"
        & $applyConfidenceColors
        $swapSource = -1
    })
    $grid.add_CellValidating({
        param($sender, $e)
        if ($e.ColumnIndex -eq $killsColIndex) {
            $txt = [string]$e.FormattedValue
            if ([string]::IsNullOrWhiteSpace($txt)) { return }
            $v = 0
            if (-not [int]::TryParse($txt, [ref]$v) -or $v -lt 0) {
                [System.Windows.Forms.MessageBox]::Show('Kills must be a non-negative integer.', 'Validation', 'OK', 'Warning') | Out-Null
                $e.Cancel = $true
            }
        }
    })

    $script:formResult = $null

    $btnApprove.Add_Click({
        $grid.EndEdit()

        $seen = @{}
        foreach ($dr in $table.Rows) {
            $player = [string]$dr['Player']
            if ([string]::IsNullOrWhiteSpace($player)) { continue }
            if ($seen.ContainsKey($player)) {
                [System.Windows.Forms.MessageBox]::Show("Duplicate player selected: $player", 'Validation', 'OK', 'Error') | Out-Null
                return
            }
            $seen[$player] = $true
        }

        $resultRows = New-Object System.Collections.Generic.List[object]
        foreach ($dr in $table.Rows) {
            $isAbsent = [bool]$dr['IsAbsent']
            $isReplacement = [bool]$dr['IsReplacement']
            $replacementName = [string]$dr['ReplacementName']
            $kills = [int]$dr['Kills']

            if ($isAbsent) {
                $isReplacement = $false
                $replacementName = ''
                $kills = 0
            }
            elseif ($isReplacement) {
                $kills = 0
            }

            [void]$resultRows.Add([pscustomobject]@{
                Position = [int]$dr['Position']
                Player = [string]$dr['Player']
                IsAbsent = $isAbsent
                IsReplacement = $isReplacement
                ReplacementName = $replacementName
                Kills = $kills
                Confidence = [string]$dr['Confidence']
            })
        }

        $script:formResult = $resultRows.ToArray()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $dialog = $form.ShowDialog()
    try { $pb.Image.Dispose() } catch {}

    if ($dialog -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $script:formResult
}

function Apply-ValidatedData {
    param(
        [string]$WorkbookPath,
        [object]$Context,
        [object[]]$Rows
    )

    $excel = $null
    $wb = $null
    $wsHebdo = $null
    $wsKills = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($WorkbookPath)
        $wsHebdo = $wb.Worksheets.Item('Classement Hebdo')
        $wsKills = $wb.Worksheets.Item('Kills')

        $assigned = @{}
        foreach ($row in $Rows) {
            $player = [string]$row.Player
            if ([string]::IsNullOrWhiteSpace($player)) { continue }
            $assigned[$player] = $row
        }

        foreach ($player in $Context.Players) {
            $rh = [int]$Context.PlayerRowsHebdo[$player]
            $rk = if ($Context.PlayerRowsKills.ContainsKey($player)) { [int]$Context.PlayerRowsKills[$player] } else { $null }

            if ($assigned.ContainsKey($player)) {
                $r = $assigned[$player]
                $isAbsent = $false
                $isReplacement = $false
                try { $isAbsent = [bool]$r.IsAbsent } catch {}
                try { $isReplacement = [bool]$r.IsReplacement } catch {}

                if ($isAbsent) {
                    $wsHebdo.Cells($rh, $Context.WeekCol).Value2 = 'ABS'
                    if ($null -ne $rk) { $wsKills.Cells($rk, $Context.WeekCol).Value2 = '' }
                }
                else {
                    $pos = [int]$r.Position
                    $code = if ($isReplacement) { 100 + $pos } else { $pos }
                    $wsHebdo.Cells($rh, $Context.WeekCol).Value2 = $code

                    if ($null -ne $rk) {
                        $kills = 0
                        if (-not $isReplacement) {
                            try { $kills = [int]$r.Kills } catch { $kills = 0 }
                        }
                        if ($kills -gt 0) { $wsKills.Cells($rk, $Context.WeekCol).Value2 = $kills }
                        else { $wsKills.Cells($rk, $Context.WeekCol).Value2 = '' }
                    }
                }
            }
            else {
                $wsHebdo.Cells($rh, $Context.WeekCol).Value2 = 'ABS'
                if ($null -ne $rk) { $wsKills.Cells($rk, $Context.WeekCol).Value2 = '' }
            }
        }

        $wb.Save()
    }
    finally {
        if ($null -ne $wb) { try { $wb.Close($true) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObjectSafe $wsKills
        Release-ComObjectSafe $wsHebdo
        Release-ComObjectSafe $wb
        Release-ComObjectSafe $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$workspaceDir = Split-Path -Path $resolvedWorkbookPath -Parent

if ([string]::IsNullOrWhiteSpace($DataDir)) { $DataDir = Join-Path $workspaceDir 'data' }
if (-not (Test-Path -LiteralPath $DataDir)) { New-Item -ItemType Directory -Path $DataDir | Out-Null }

if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = $DataDir }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

$photoFile = $null
$photoDate = $null
if ([string]::IsNullOrWhiteSpace($ImagePath)) {
    $latest = Get-LatestDatedPhoto -Folder $DataDir
    if ($null -eq $latest) { throw "No dated photo found. Example: 2 mars 2026.jpeg" }
    $photoFile = $latest.File.FullName
    $photoDate = $latest.Date
}
else {
    $photoFile = (Resolve-Path -LiteralPath $ImagePath).Path
    $photoDate = Get-DateFromPhotoName -FileName ([IO.Path]::GetFileName($photoFile))
    if ($null -eq $photoDate) { throw "Cannot parse date from image file name: $photoFile" }
}

$context = Get-WorkbookContext -WorkbookPath $resolvedWorkbookPath -PhotoDate $photoDate
Write-Output ("Preparation terminee: photo {0}, semaine {1}." -f [IO.Path]::GetFileName($photoFile), $context.WeekNum)

$ocrObj = $null
$ocrStatus = 'manual'
$ocrNotes = 'none'

if (-not $NoOcr) {
    Write-Output ("OCR en cours via {0}... (20-90 sec)" -f $Model)
    $apiKey = $env:OPENAI_API_KEY
    if ([string]::IsNullOrWhiteSpace($apiKey)) { $apiKey = [Environment]::GetEnvironmentVariable('OPENAI_API_KEY', 'User') }
    if ([string]::IsNullOrWhiteSpace($apiKey)) { throw 'OPENAI_API_KEY is not set. Run setup first.' }

    try {
        $ocrObj = Invoke-PhotoOcr -ApiKey $apiKey -Model $Model -ImagePath $photoFile -ExpectedDate $photoDate -Players $context.Players -MaxPositions $context.Players.Count
        $ocrStatus = 'ok'
        $ocrNotes = ((@($ocrObj.notes)) -join '; ')
        if ([string]::IsNullOrWhiteSpace($ocrNotes)) { $ocrNotes = 'none' }
        Write-Output 'OCR termine.'
    }
    catch {
        $ocrStatus = 'failed'
        $ocrNotes = $_.Exception.Message
        Write-Output ("OCR echec: {0}" -f $ocrNotes)
        [System.Windows.Forms.MessageBox]::Show("OCR failed: $ocrNotes`r`nContinue with manual mode.", 'OCR', 'OK', 'Warning') | Out-Null
        $ocrObj = $null
    }
}

Write-Output "Ouverture de la fenetre de validation..."
$rows = Build-InitialRows -Players $context.Players -OcrResult $ocrObj
$validatedRows = Show-ValidationForm -ImagePath $photoFile -Players $context.Players -Rows $rows -WeekNum $context.WeekNum -DateKey $photoDate.ToString('yyyy-MM-dd') -OcrStatus $ocrStatus -Notes $ocrNotes
if ($null -eq $validatedRows) { Write-Output 'Canceled by user'; exit 1 }

Apply-ValidatedData -WorkbookPath $resolvedWorkbookPath -Context $context -Rows $validatedRows

$report = [pscustomobject]@{
    created_at = (Get-Date).ToString('s')
    workbook = $resolvedWorkbookPath
    image = $photoFile
    week = $context.WeekNum
    date = $photoDate.ToString('yyyy-MM-dd')
    ocr_status = $ocrStatus
    rows = $validatedRows
}
$reportPath = Join-Path $OutputDir ("validation_ui_semaine_{0}_{1}.json" -f $context.WeekNum, (Get-Date -Format 'yyyyMMdd_HHmmss'))
($report | ConvertTo-Json -Depth 20) | Set-Content -LiteralPath $reportPath -Encoding UTF8
Write-Output "Validation report: $reportPath"

if (-not $SkipExport) {
    $scriptDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { Split-Path -Path $MyInvocation.MyCommand.Path -Parent } else { $PSScriptRoot }
    $exportScript = Join-Path $scriptDir 'export_classement_hebdo.ps1'
    if (-not (Test-Path -LiteralPath $exportScript)) { throw "Missing export script: $exportScript" }
    & $exportScript -WorkbookPath $resolvedWorkbookPath -OutputDir $DataDir -EmailTo $EmailTo -PublicUrl $PublicUrl -PublishRepoDir $PublishRepoDir
}

Write-Output ("Done. Week {0} updated with validated UI." -f $context.WeekNum)






