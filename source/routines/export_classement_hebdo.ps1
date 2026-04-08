param(
    [string]$WorkbookPath = ".\Poker Stanley Hiver 2026.xlsx",
    [string]$OutputDir = "",
    [switch]$SaveWorkbook,
    [string]$EmailTo = "",
    [string]$PublicUrl = "",
    [string]$PublishRepoDir = ""
)

$ErrorActionPreference = "Stop"

function Release-ComObjectSafe {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) } catch {}
    }
}

function Get-PointsMapFromSheet {
    param([object]$WsCalc)

    $map = @{}
    $lastRow = $WsCalc.Cells($WsCalc.Rows.Count, 1).End($xlUp).Row
    for ($r = 1; $r -le $lastRow; $r++) {
        $keyTxt = [string]$WsCalc.Cells($r, 1).Text
        if ([string]::IsNullOrWhiteSpace($keyTxt)) { continue }

        $key = $keyTxt.Trim().ToUpperInvariant()
        $points = 0
        $ok = $false

        $v = $WsCalc.Cells($r, 2).Value2
        if ($null -ne $v) {
            try {
                $points = [int][Math]::Round([double]$v)
                $ok = $true
            }
            catch {}
        }

        if (-not $ok) {
            $txt = ([string]$WsCalc.Cells($r, 2).Text).Trim()
            $tmp = 0
            if ([int]::TryParse($txt, [ref]$tmp)) {
                $points = $tmp
                $ok = $true
            }
        }

        if ($ok) { $map[$key] = $points }
    }

    if (-not $map.ContainsKey('ABS')) { $map['ABS'] = 0 }
    return $map
}

function Normalize-IneligibleWeeklyKills {
    param(
        [object]$WsHebdo,
        [object]$WsKills,
        [int]$LastPlayerRow,
        [int]$LastWeekCol
    )

    $killsRows = @{}
    $lastKillsRow = $WsKills.Cells($WsKills.Rows.Count, 1).End($xlUp).Row
    for ($r = 3; $r -le $lastKillsRow; $r++) {
        $name = ([string]$WsKills.Cells($r, 1).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $killsRows[$name] = $r
        }
    }

    $changes = New-Object System.Collections.Generic.List[object]

    for ($r = 3; $r -le $LastPlayerRow; $r++) {
        $player = ([string]$WsHebdo.Cells($r, 1).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($player) -or -not $killsRows.ContainsKey($player)) { continue }

        $killsRow = [int]$killsRows[$player]
        for ($weekCol = 2; $weekCol -le $LastWeekCol; $weekCol++) {
            $rawCode = ([string]$WsHebdo.Cells($r, $weekCol).Text).Trim()
            $upperCode = $rawCode.ToUpperInvariant()
            $isIneligible = $false
            $reason = ''

            if ([string]::IsNullOrWhiteSpace($rawCode) -or $upperCode -eq 'ABS') {
                $isIneligible = $true
                $reason = 'absent'
            }
            else {
                $codeNum = 0
                if ([int]::TryParse($rawCode, [ref]$codeNum) -and $codeNum -ge 100) {
                    $isIneligible = $true
                    $reason = 'replacement'
                }
            }

            if (-not $isIneligible) { continue }

            $killsText = ([string]$WsKills.Cells($killsRow, $weekCol).Text).Trim()
            if ([string]::IsNullOrWhiteSpace($killsText)) { continue }

            $WsKills.Cells($killsRow, $weekCol).Value2 = ''
            [void]$changes.Add([pscustomobject]@{
                Week = $weekCol - 1
                Player = $player
                PreviousKills = $killsText
                Reason = $reason
            })
        }
    }

    return [pscustomobject]@{
        ChangeCount = $changes.Count
        Changes = $changes.ToArray()
    }
}

function Test-WeeklyConsistency {
    param(
        [object]$WsHebdo,
        [int]$LastPlayerRow,
        [int]$LastWeekCol
    )

    $fatalIssues = New-Object System.Collections.Generic.List[object]

    for ($weekCol = 2; $weekCol -le $LastWeekCol; $weekCol++) {
        $weekNum = $weekCol - 1
        $weekLabel = ([string]$WsHebdo.Cells(2, $weekCol).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($weekLabel)) { $weekLabel = "Semaine $weekNum" }

        $seenPositions = @{}
        $activePositions = New-Object System.Collections.Generic.List[int]

        for ($r = 3; $r -le $LastPlayerRow; $r++) {
            $player = ([string]$WsHebdo.Cells($r, 1).Text).Trim()
            if ([string]::IsNullOrWhiteSpace($player)) { continue }

            $rawCode = ([string]$WsHebdo.Cells($r, $weekCol).Text).Trim()
            $upperCode = $rawCode.ToUpperInvariant()
            if ([string]::IsNullOrWhiteSpace($rawCode) -or $upperCode -eq 'ABS') { continue }

            $codeNum = 0
            if (-not [int]::TryParse($rawCode, [ref]$codeNum)) {
                [void]$fatalIssues.Add([pscustomobject]@{
                    Week = $weekNum
                    Label = $weekLabel
                    Player = $player
                    Issue = "Code invalide '$rawCode'"
                })
                continue
            }

            $position = if ($codeNum -ge 100) { $codeNum - 100 } else { $codeNum }
            if ($position -lt 1 -or $position -gt ($LastPlayerRow - 2)) {
                [void]$fatalIssues.Add([pscustomobject]@{
                    Week = $weekNum
                    Label = $weekLabel
                    Player = $player
                    Issue = "Position invalide $position"
                })
                continue
            }

            if ($seenPositions.ContainsKey($position)) {
                [void]$fatalIssues.Add([pscustomobject]@{
                    Week = $weekNum
                    Label = $weekLabel
                    Player = $player
                    Issue = "Position $position dupliquee avec $($seenPositions[$position])"
                })
                continue
            }

            $seenPositions[$position] = $player
            [void]$activePositions.Add($position)
        }

        $sortedPositions = @($activePositions | Sort-Object)
        for ($idx = 0; $idx -lt $sortedPositions.Count; $idx++) {
            $expected = $idx + 1
            if ([int]$sortedPositions[$idx] -ne $expected) {
                [void]$fatalIssues.Add([pscustomobject]@{
                    Week = $weekNum
                    Label = $weekLabel
                    Player = ''
                    Issue = "Sequence incoherente: position $expected manquante"
                })
                break
            }
        }
    }

    return [pscustomobject]@{
        FatalCount = $fatalIssues.Count
        FatalIssues = $fatalIssues.ToArray()
    }
}

function Build-WeeklyTabs {
    param(
        [object]$WsHebdo,
        [object]$WsKills,
        [hashtable]$PointsMap,
        [int]$LastPlayerRow,
        [int]$LastWeekCol
    )

    $killsRows = @{}
    $lastKillsRow = $WsKills.Cells($WsKills.Rows.Count, 1).End($xlUp).Row
    for ($r = 3; $r -le $lastKillsRow; $r++) {
        $name = ([string]$WsKills.Cells($r, 1).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $killsRows[$name] = $r
        }
    }

    $buttons = New-Object System.Text.StringBuilder
    $panels = New-Object System.Text.StringBuilder
    $hasAnyWeek = $false

    for ($weekCol = 2; $weekCol -le $LastWeekCol; $weekCol++) {
        $weekNum = $weekCol - 1
        $weekLabel = ([string]$WsHebdo.Cells(2, $weekCol).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($weekLabel)) { $weekLabel = "Semaine $weekNum" }
        $weekDate = ([string]$WsHebdo.Cells(1, $weekCol).Text).Trim()

        $tabId = "week-$weekNum"
        $isActive = ($weekCol -eq $LastWeekCol)
        $btnClass = if ($isActive) { 'tab-btn active' } else { 'tab-btn' }
        $panelClass = if ($isActive) { 'tab-panel active' } else { 'tab-panel' }

        [void]$buttons.AppendLine("<button class=""$btnClass"" data-tab=""$tabId"">$([System.Net.WebUtility]::HtmlEncode($weekLabel))</button>")

        $rows = New-Object System.Collections.Generic.List[object]
        for ($r = 3; $r -le $LastPlayerRow; $r++) {
            $player = ([string]$WsHebdo.Cells($r, 1).Text).Trim()
            if ([string]::IsNullOrWhiteSpace($player)) { continue }

            $rawCode = ([string]$WsHebdo.Cells($r, $weekCol).Text).Trim()
            $upperCode = $rawCode.ToUpperInvariant()
            $isAbsent = $false
            $isReplacement = $false
            $pointsKey = ''
            $positionDisplay = ''
            $positionSort = 999

            if ([string]::IsNullOrWhiteSpace($rawCode) -or $upperCode -eq 'ABS') {
                $isAbsent = $true
                $pointsKey = 'ABS'
                $positionDisplay = 'ABS'
            }
            else {
                $codeNum = 0
                if ([int]::TryParse($rawCode, [ref]$codeNum)) {
                    if ($codeNum -ge 100) { $isReplacement = $true }
                    $pointsKey = $codeNum.ToString()
                    $posNum = if ($codeNum -ge 100) { $codeNum - 100 } else { $codeNum }
                    if ($posNum -lt 1) { $posNum = $codeNum }
                    $positionDisplay = $posNum.ToString()
                    $positionSort = $posNum
                }
                else {
                    $pointsKey = $upperCode
                    $positionDisplay = $rawCode
                    $positionSort = 998
                }
            }

            $points = 0
            if ($PointsMap.ContainsKey($pointsKey)) {
                $points = [int]$PointsMap[$pointsKey]
            }

            $killsDisplay = '0'
            if (-not $isAbsent -and -not $isReplacement -and $killsRows.ContainsKey($player)) {
                $killsTxt = ([string]$WsKills.Cells([int]$killsRows[$player], $weekCol).Text).Trim()
                if (-not [string]::IsNullOrWhiteSpace($killsTxt)) {
                    $killsDisplay = $killsTxt
                }
            }

            $nameClass = ''
            if ($isAbsent) { $nameClass = 'absent' }
            elseif ($isReplacement) { $nameClass = 'replacement' }

            [void]$rows.Add([pscustomobject]@{
                Player = $player
                NameClass = $nameClass
                PositionDisplay = $positionDisplay
                PositionSort = $positionSort
                Points = $points
                Kills = $killsDisplay
                IsAbsent = if ($isAbsent) { 1 } else { 0 }
            })
        }

        $sortedRows = $rows | Sort-Object IsAbsent, @{ Expression = { $_.PositionSort }; Ascending = $true }, @{ Expression = { $_.Player }; Ascending = $true }

        [void]$panels.AppendLine("<section id=""$tabId"" class=""$panelClass"">")
        [void]$panels.AppendLine("<h2>$([System.Net.WebUtility]::HtmlEncode($weekLabel)) <span class=""week-date"">$([System.Net.WebUtility]::HtmlEncode($weekDate))</span></h2>")
        [void]$panels.AppendLine('<table class="week-table">')
        [void]$panels.AppendLine('<thead><tr><th>Position</th><th>Points</th><th>Joueurs</th><th>Kills</th></tr></thead><tbody>')

        foreach ($row in $sortedRows) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$row.Player)
            $kills = [System.Net.WebUtility]::HtmlEncode([string]$row.Kills)
            $positionText = [System.Net.WebUtility]::HtmlEncode([string]$row.PositionDisplay)
            $nameCellClass = if ([string]::IsNullOrWhiteSpace([string]$row.NameClass)) { 'player' } else { "player $([string]$row.NameClass)" }
            [void]$panels.AppendLine("<tr><td class=""num"">$positionText</td><td class=""num"">$([int]$row.Points)</td><td class=""$nameCellClass"">$name</td><td class=""num"">$kills</td></tr>")
        }

        [void]$panels.AppendLine('</tbody></table>')
        [void]$panels.AppendLine('</section>')
        $hasAnyWeek = $true
    }

    return [pscustomobject]@{
        ButtonsHtml = $buttons.ToString()
        PanelsHtml = $panels.ToString()
        HasAnyWeek = $hasAnyWeek
    }
}

function Convert-ExcelColorToCss {
    param([object]$Cell)

    if ($null -eq $Cell) { return '' }

    $sources = New-Object System.Collections.Generic.List[object]
    try {
        $display = $Cell.DisplayFormat
        if ($null -ne $display) { [void]$sources.Add($display) }
    }
    catch {}
    [void]$sources.Add($Cell)

    foreach ($src in $sources) {
        try {
            $interior = $src.Interior
            if ($null -eq $interior) { continue }

            $pattern = 0
            $colorIndex = -4142
            $color = 0

            try { $pattern = [int]$interior.Pattern } catch {}
            try { $colorIndex = [int]$interior.ColorIndex } catch {}
            try { $color = [int]$interior.Color } catch {}

            if ($pattern -eq 0 -or $colorIndex -eq -4142 -or $color -le 0) { continue }

            $r = $color -band 255
            $g = ($color -shr 8) -band 255
            $b = ($color -shr 16) -band 255

            if ($r -ge 250 -and $g -ge 250 -and $b -ge 250) { continue }
            return ('#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
        }
        catch {}
    }

    return ''
}

function Build-TotalTabHtml {
    param(
        [object]$WsTotal,
        [object]$WsHebdo
    )

    $headerRow = 2
    $candidateLastRow = [int]$WsTotal.Cells($WsTotal.Rows.Count, 2).End($xlUp).Row
    if ($candidateLastRow -lt $headerRow) { $candidateLastRow = $headerRow }

    $lastRow = $headerRow
    for ($r = $candidateLastRow; $r -ge 3; $r--) {
        $name = ([string]$WsTotal.Cells($r, 2).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $lastRow = $r
            break
        }
    }

    $cols = New-Object System.Collections.Generic.List[int]
    [void]$cols.Add(1)
    [void]$cols.Add(2)

    for ($weekCol = 2; $weekCol -le 15; $weekCol++) {
        $totalCol = $weekCol + 1
        if ($totalCol -le 30) {
            [void]$cols.Add($totalCol)
        }
    }

    for ($c = 17; $c -le 30; $c++) {
        $head = ([string]$WsTotal.Cells($headerRow, $c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($head) -and -not ($cols -contains $c)) {
            [void]$cols.Add($c)
        }
    }

    $totalCol = 17
    $killsCol = 18
    foreach ($c in $cols) {
        $headUpper = ([string]$WsTotal.Cells($headerRow, $c).Text).Trim().ToUpperInvariant()
        if ($headUpper -eq 'TOTAL') { $totalCol = $c }
        if (-not [string]::IsNullOrWhiteSpace($headUpper) -and $headUpper -like '*KILL*') { $killsCol = $c }
    }

    $maxKills = $null
    for ($r = 3; $r -le $lastRow; $r++) {
        $name = ([string]$WsTotal.Cells($r, 2).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $killsTxt = ([string]$WsTotal.Cells($r, $killsCol).Text).Trim()
        $killsVal = 0
        if ([int]::TryParse($killsTxt, [ref]$killsVal)) {
            if ($null -eq $maxKills -or $killsVal -gt $maxKills) {
                $maxKills = $killsVal
            }
        }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<section>')
    [void]$sb.AppendLine('<h2>Classement Total</h2>')
    [void]$sb.AppendLine('<table class="total-table">')
    [void]$sb.AppendLine('<thead><tr>')

    foreach ($c in $cols) {
        if ($c -ge 3 -and $c -le 16) {
            $weekTxt = ([string]$WsTotal.Cells($headerRow, $c).Text).Trim()
            if ([string]::IsNullOrWhiteSpace($weekTxt)) { $weekTxt = "Semaine $($c - 2)" }
            $dateTxt = ([string]$WsHebdo.Cells(1, $c - 1).Text).Trim()
            $weekHtml = [System.Net.WebUtility]::HtmlEncode($weekTxt)
            $dateHtml = [System.Net.WebUtility]::HtmlEncode($dateTxt)
            if ([string]::IsNullOrWhiteSpace($dateTxt)) {
                [void]$sb.AppendLine("<th>$weekHtml</th>")
            }
            else {
                [void]$sb.AppendLine("<th>$weekHtml<br><span class=""subdate"">$dateHtml</span></th>")
            }
        }
        else {
            $th = [System.Net.WebUtility]::HtmlEncode([string]$WsTotal.Cells($headerRow, $c).Text)
            [void]$sb.AppendLine("<th>$th</th>")
        }
    }

    [void]$sb.AppendLine('</tr></thead>')
    [void]$sb.AppendLine('<tbody>')

    for ($r = 3; $r -le $lastRow; $r++) {
        $name = ([string]$WsTotal.Cells($r, 2).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $rank = 0
        $isTop10 = [int]::TryParse((([string]$WsTotal.Cells($r, 1).Text).Trim()), [ref]$rank) -and $rank -ge 1 -and $rank -le 10

        $killsVal = 0
        $hasKillsValue = [int]::TryParse((([string]$WsTotal.Cells($r, $killsCol).Text).Trim()), [ref]$killsVal)
        $isBestKills = ($null -ne $maxKills) -and $hasKillsValue -and ($killsVal -eq $maxKills)

        [void]$sb.AppendLine('<tr>')
        foreach ($c in $cols) {
            $txt = [System.Net.WebUtility]::HtmlEncode([string]$WsTotal.Cells($r, $c).Text)

            $classes = New-Object System.Collections.Generic.List[string]
            if ($c -eq 2) { [void]$classes.Add('total-player') } else { [void]$classes.Add('num') }
            if ($isTop10 -and ($c -eq 1 -or $c -eq 2)) { [void]$classes.Add('qualif') }
            if ($isBestKills -and $c -eq $killsCol) { [void]$classes.Add('best-kills') }
            if ($c -eq $totalCol) { [void]$classes.Add('total-score') }

            $fillColor = ''
            if ($c -ge 3 -and $c -le 16) {
                $fillColor = Convert-ExcelColorToCss -Cell $WsTotal.Cells($r, $c)
            }
            $styleAttr = ''
            if (-not [string]::IsNullOrWhiteSpace($fillColor)) {
                $styleAttr = " style=""background-color:$fillColor;"""
            }

            $classAttr = [string]::Join(' ', $classes.ToArray())
            [void]$sb.AppendLine("<td class=""$classAttr""$styleAttr>$txt</td>")
        }
        [void]$sb.AppendLine('</tr>')
    }

    [void]$sb.AppendLine('</tbody></table>')
    [void]$sb.AppendLine('</section>')
    return $sb.ToString()
}

function Build-ProgressionData {
    param(
        [object]$WsHebdo,
        [object]$WsTotal,
        [int]$LastWeekCol
    )

    $weeks = New-Object System.Collections.Generic.List[object]
    for ($c = 2; $c -le $LastWeekCol; $c++) {
        $n = $c - 1
        $label = ([string]$WsHebdo.Cells(2, $c).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($label)) { $label = "Semaine $n" }
        $dateTxt = ([string]$WsHebdo.Cells(1, $c).Text).Trim()
        [void]$weeks.Add([pscustomobject]@{
            label = $label
            shortLabel = "S$n"
            date = $dateTxt
        })
    }

    $headerRow = 2
    $candidateLastRow = [int]$WsTotal.Cells($WsTotal.Rows.Count, 2).End($xlUp).Row
    if ($candidateLastRow -lt $headerRow) { $candidateLastRow = $headerRow }

    $playerRows = New-Object System.Collections.Generic.List[object]
    for ($r = 3; $r -le $candidateLastRow; $r++) {
        $name = ([string]$WsTotal.Cells($r, 2).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        [void]$playerRows.Add([pscustomobject]@{
            Name = $name
            Row = $r
        })
    }

    if ($playerRows.Count -eq 0) {
        return [pscustomobject]@{
            weeks = $weeks.ToArray()
            players = @()
            maxPosition = 1
        }
    }

    $weekCount = [Math]::Max(0, $LastWeekCol - 1)
    $cumulative = @{}
    $positionsByName = @{}

    foreach ($p in $playerRows) {
        $cumulative[$p.Name] = 0.0
        $positionsByName[$p.Name] = New-Object System.Collections.Generic.List[object]
    }

    for ($week = 1; $week -le $weekCount; $week++) {
        $totalCol = $week + 2

        foreach ($p in $playerRows) {
            $points = 0.0
            $v = $WsTotal.Cells($p.Row, $totalCol).Value2
            if ($null -ne $v) {
                try {
                    $points = [double]$v
                }
                catch {
                    $txt = ([string]$WsTotal.Cells($p.Row, $totalCol).Text).Trim()
                    $tmp = 0.0
                    if ([double]::TryParse($txt, [ref]$tmp)) {
                        $points = $tmp
                    }
                }
            }
            $cumulative[$p.Name] = [double]$cumulative[$p.Name] + $points
        }

        $ranked = $playerRows | Sort-Object @{ Expression = { [double]$cumulative[$_.Name] }; Descending = $true }, @{ Expression = { [string]$_.Name }; Ascending = $true }

        $pos = 1
        foreach ($rp in $ranked) {
            [void]$positionsByName[$rp.Name].Add($pos)
            $pos++
        }
    }

    $players = New-Object System.Collections.Generic.List[object]
    foreach ($p in $playerRows) {
        [void]$players.Add([pscustomobject]@{
            name = [string]$p.Name
            positions = $positionsByName[$p.Name].ToArray()
        })
    }

    return [pscustomobject]@{
        weeks = $weeks.ToArray()
        players = $players.ToArray()
        maxPosition = [Math]::Max(1, $playerRows.Count)
    }
}
function Export-PublicHtml {
    param(
        [object]$WsHebdo,
        [object]$WsKills,
        [object]$WsTotal,
        [object]$WsCalc,
        [string]$OutputPath,
        [int]$WeekNumber,
        [int]$LastPlayerRow,
        [int]$LastWeekColHebdo,
        [object]$SiteConfig = $null
    )

    $pointsMap = Get-PointsMapFromSheet -WsCalc $WsCalc
    $weeklyTabs = Build-WeeklyTabs -WsHebdo $WsHebdo -WsKills $WsKills -PointsMap $pointsMap -LastPlayerRow $LastPlayerRow -LastWeekCol $LastWeekColHebdo
    $totalHtml = Build-TotalTabHtml -WsTotal $WsTotal -WsHebdo $WsHebdo
    $progressionData = Build-ProgressionData -WsHebdo $WsHebdo -WsTotal $WsTotal -LastWeekCol $LastWeekColHebdo
    $progressionJson = $progressionData | ConvertTo-Json -Depth 12 -Compress
    if ($null -eq $SiteConfig) {
        $SiteConfig = [ordered]@{
            enableFinalStackCalculator = $false
            relaxDesktopHorizontalScroll = $false
        }
    }
    $siteConfigJson = $SiteConfig | ConvertTo-Json -Depth 8 -Compress
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    $scriptDir = Split-Path -Parent $PSCommandPath
    $templatePath = Join-Path $scriptDir 'site_template.html'
    if (-not (Test-Path -LiteralPath $templatePath)) {
        throw "Template introuvable: $templatePath"
    }

    $html = Get-Content -Raw -LiteralPath $templatePath
    $tokens = @('{{STAMP}}', '{{WEEK_PANELS}}', '{{TOTAL_HTML}}', '{{PROGRESSION_JSON}}', '{{SITE_CONFIG_JSON}}')
    foreach ($token in $tokens) {
        if (-not $html.Contains($token)) {
            throw "Template invalide: placeholder manquant $token"
        }
    }

    $html = $html.Replace('{{STAMP}}', $stamp)
    $html = $html.Replace('{{WEEK_PANELS}}', $weeklyTabs.PanelsHtml)
    $html = $html.Replace('{{TOTAL_HTML}}', $totalHtml)
    $html = $html.Replace('{{PROGRESSION_JSON}}', $progressionJson)
    $html = $html.Replace('{{SITE_CONFIG_JSON}}', $siteConfigJson)

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
}

function Send-OutlookMailWithAttachment {
    param(
        [string]$To,
        [string]$Subject,
        [string]$AttachmentPath,
        [string]$Body = ""
    )

    if (-not (Test-Path -LiteralPath $AttachmentPath)) {
        throw "Attachment not found: $AttachmentPath"
    }

    $outlook = $null
    $mail = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = $To
        $mail.Subject = $Subject
        if ([string]::IsNullOrWhiteSpace($Body)) {
    $mail.Body = "Classement en pièce jointe."
        }
        else {
            $mail.Body = $Body
        }
        [void]$mail.Attachments.Add($AttachmentPath)
        $mail.Send()
    }
    catch {
        throw "Envoi email impossible via Outlook: $($_.Exception.Message)"
    }
    finally {
        Release-ComObjectSafe $mail
        Release-ComObjectSafe $outlook
    }
}

function Resolve-GitExecutable {
    $cmd = Get-Command git -ErrorAction SilentlyContinue
    if ($cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) {
        return [string]$cmd.Source
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    [void]$candidates.Add((Join-Path $env:ProgramFiles 'Git\cmd\git.exe'))
    [void]$candidates.Add((Join-Path $env:ProgramFiles 'Git\bin\git.exe'))
    [void]$candidates.Add((Join-Path $env:LocalAppData 'Programs\Git\cmd\git.exe'))

    $ghDesktopRoot = Join-Path $env:LocalAppData 'GitHubDesktop'
    if (Test-Path -LiteralPath $ghDesktopRoot) {
        $apps = Get-ChildItem -LiteralPath $ghDesktopRoot -Directory -Filter 'app-*' -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        foreach ($app in $apps) {
            [void]$candidates.Add((Join-Path $app.FullName 'resources\app\git\cmd\git.exe'))
        }
    }

    foreach ($path in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    return $null
}

function Invoke-GitChecked {
    param(
        [string]$GitExe,
        [string]$RepoDir,
        [string[]]$GitArgs
    )

    $previousPreference = $ErrorActionPreference
    try {
        # git writes progress to stderr even on success; do not treat that as a terminating error.
        $ErrorActionPreference = 'Continue'
        $out = & $GitExe -C $RepoDir @GitArgs 2>&1
        $exitCode = $LASTEXITCODE
    }
    catch {
        throw "git $($GitArgs -join ' ') invocation failed: $($_.Exception.Message)"
    }
    finally {
        $ErrorActionPreference = $previousPreference
    }

    $lines = @($out | ForEach-Object { $_.ToString() })
    if ($exitCode -ne 0) {
        $msg = ($lines -join ' ')
        throw "git $($GitArgs -join ' ') failed: $msg"
    }
    return ,$lines
}

function Auto-CommitPushRepo {
    param(
        [string]$RepoDir,
        [int]$WeekNumber,
        [string[]]$FilesToAdd
    )

    try {
        if (-not (Test-Path -LiteralPath (Join-Path $RepoDir '.git'))) {
        return "Git auto-push ignoré: '$RepoDir' n'est pas un repo git."
        }

        $gitExe = Resolve-GitExecutable
        if ([string]::IsNullOrWhiteSpace($gitExe)) {
        return 'Git auto-push ignoré: git.exe introuvable.'
        }

        $existing = New-Object System.Collections.Generic.List[string]
        foreach ($f in $FilesToAdd) {
            if (Test-Path -LiteralPath (Join-Path $RepoDir $f)) {
                [void]$existing.Add($f)
            }
        }
        if ($existing.Count -eq 0) {
        return 'Git auto-push ignoré: aucun fichier à ajouter.'
        }

        # Keep auto-push resilient on Windows when line-ending warnings are enabled locally.
        Invoke-GitChecked -GitExe $gitExe -RepoDir $RepoDir -GitArgs (@('-c', 'core.safecrlf=false', 'add', '--') + $existing.ToArray()) | Out-Null
        $status = Invoke-GitChecked -GitExe $gitExe -RepoDir $RepoDir -GitArgs (@('status', '--porcelain', '--') + $existing.ToArray())
        if (-not $status -or (($status -join '').Trim().Length -eq 0)) {
        return 'Git auto-push: aucun changement à publier.'
        }

        $commitMsg = "Mise à jour classement semaine $WeekNumber"
        Invoke-GitChecked -GitExe $gitExe -RepoDir $RepoDir -GitArgs @('commit', '-m', $commitMsg) | Out-Null
        Invoke-GitChecked -GitExe $gitExe -RepoDir $RepoDir -GitArgs @('push', 'origin', 'main') | Out-Null
        return "Git auto-push OK: $commitMsg"
    }
    catch {
        return "Git auto-push échec: $($_.Exception.Message)"
    }
}

function Sync-SourceFilesToPublishRepo {
    param(
        [string]$WorkspaceDir,
        [string]$ScriptDir,
        [string]$PublishRepoDir
    )

    if ([string]::IsNullOrWhiteSpace($PublishRepoDir) -or -not (Test-Path -LiteralPath $PublishRepoDir)) {
        return [pscustomobject]@{
            Summary = ''
            RelativeFiles = @()
        }
    }

    $sourceRoot = Join-Path $PublishRepoDir 'source'
    $sourceRoutines = Join-Path $sourceRoot 'routines'
    if (-not (Test-Path -LiteralPath $sourceRoutines)) {
        New-Item -ItemType Directory -Path $sourceRoutines -Force | Out-Null
    }

    $filesToCopy = @(
        @{ Source = (Join-Path $ScriptDir 'site_template.html'); Destination = (Join-Path $sourceRoutines 'site_template.html'); Relative = 'source/routines/site_template.html' },
        @{ Source = (Join-Path $ScriptDir 'export_classement_hebdo.ps1'); Destination = (Join-Path $sourceRoutines 'export_classement_hebdo.ps1'); Relative = 'source/routines/export_classement_hebdo.ps1' },
        @{ Source = (Join-Path $ScriptDir 'validation_photo_classement_ui.ps1'); Destination = (Join-Path $sourceRoutines 'validation_photo_classement_ui.ps1'); Relative = 'source/routines/validation_photo_classement_ui.ps1' },
        @{ Source = (Join-Path $WorkspaceDir 'lancer_validation_photo_ui.bat'); Destination = (Join-Path $sourceRoot 'lancer_validation_photo_ui.bat'); Relative = 'source/lancer_validation_photo_ui.bat' },
        @{ Source = (Join-Path $WorkspaceDir 'lancer_validation_photo_ui_preview.bat'); Destination = (Join-Path $sourceRoot 'lancer_validation_photo_ui_preview.bat'); Relative = 'source/lancer_validation_photo_ui_preview.bat' },
        @{ Source = (Join-Path $WorkspaceDir 'lancer_export_preview.bat'); Destination = (Join-Path $sourceRoot 'lancer_export_preview.bat'); Relative = 'source/lancer_export_preview.bat' },
        @{ Source = (Join-Path $WorkspaceDir 'PREVIEW_WORKFLOW.md'); Destination = (Join-Path $sourceRoot 'PREVIEW_WORKFLOW.md'); Relative = 'source/PREVIEW_WORKFLOW.md' }
    )

    $synced = New-Object System.Collections.Generic.List[string]
    foreach ($item in $filesToCopy) {
        if (-not (Test-Path -LiteralPath $item.Source)) { continue }
        Copy-Item -LiteralPath $item.Source -Destination $item.Destination -Force
        [void]$synced.Add([string]$item.Relative)
    }

    $readmePath = Join-Path $sourceRoot 'README.md'
    $readme = @'
# Source du site Poker Stanley

Ce dossier contient une copie des fichiers source utilisés pour générer et publier le site.

- `routines/site_template.html` : interface web source
- `routines/export_classement_hebdo.ps1` : génération HTML/PNG + publication
- `routines/validation_photo_classement_ui.ps1` : validation photo/OCR
- `lancer_validation_photo_ui*.bat` : lanceurs de validation
- `lancer_export_preview.bat` : export preview

Le site public continue d'être servi par `index.html` à la racine du repo.
'@
    Set-Content -LiteralPath $readmePath -Value $readme -Encoding UTF8
    [void]$synced.Add('source/README.md')

    return [pscustomobject]@{
        Summary = "Sources synchronisées dans: $sourceRoot"
        RelativeFiles = $synced.ToArray()
    }
}

$xlUp = -4162
$xlToLeft = -4159
$xlNormal = -4143
$xlYes = 1
$xlSortOnValues = 0
$xlAscending = 1
$xlDescending = 2
$xlTopToBottom = 1
$xlPinYin = 1
$xlScreen = 1
$xlPrinter = 2
$xlBitmap = 2
$xlPicture = -4147

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$workspaceDir = Split-Path -Path $resolvedWorkbookPath -Parent
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $workspaceDir 'data'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
if ([string]::IsNullOrWhiteSpace($PublicUrl)) {
    $PublicUrl = $env:POKER_PUBLIC_URL
}
if ([string]::IsNullOrWhiteSpace($PublishRepoDir)) {
    $defaultRepoDir = Join-Path $workspaceDir 'Poker-Classement'
    if (Test-Path -LiteralPath $defaultRepoDir) {
        $PublishRepoDir = $defaultRepoDir
    }
}

$excel = $null
$workbook = $null
$wsHebdo = $null
$wsKills = $null
$wsTotal = $null
$wsCalc = $null
$tableTotal = $null
$chartObj = $null
$killsNormalizationSummary = ""
$killsNormalizationChangeCount = 0
$consistencySummary = ""
$sourceSyncSummary = ""
$sourceSyncFiles = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.WindowState = $xlNormal
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $true

    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath)
    $wsHebdo = $workbook.Worksheets.Item("Classement Hebdo")
    $wsKills = $workbook.Worksheets.Item("Kills")
    $wsTotal = $workbook.Worksheets.Item("Classement Total")
    $wsCalc = $workbook.Worksheets.Item("Calcul de points")

    $lastPlayerRow = $wsHebdo.Cells($wsHebdo.Rows.Count, 1).End($xlUp).Row
    if ($lastPlayerRow -lt 3) {
        throw "Aucun joueur détecté dans 'Classement Hebdo'."
    }

    $lastWeekColHebdo = 0
    for ($col = 15; $col -ge 2; $col--) {
        $rng = $wsHebdo.Range($wsHebdo.Cells(3, $col), $wsHebdo.Cells($lastPlayerRow, $col))
        if ([int]$excel.WorksheetFunction.CountA($rng) -gt 0) {
            $lastWeekColHebdo = $col
            break
        }
    }
    if ($lastWeekColHebdo -eq 0) {
        throw "Aucune semaine renseignée dans 'Classement Hebdo'."
    }

    $killsNormalization = Normalize-IneligibleWeeklyKills -WsHebdo $wsHebdo -WsKills $wsKills -LastPlayerRow $lastPlayerRow -LastWeekCol $lastWeekColHebdo
    $killsNormalizationChangeCount = [int]$killsNormalization.ChangeCount
    if ($killsNormalizationChangeCount -gt 0) {
        $examples = @(
            $killsNormalization.Changes |
                Select-Object -First 5 |
                ForEach-Object { "S$($_.Week) $($_.Player)=$($_.PreviousKills)" }
        )
        $killsNormalizationSummary = "Kills corrigés pour joueurs absents/remplacés: $killsNormalizationChangeCount"
        if ($examples.Count -gt 0) {
            $killsNormalizationSummary += " (" + ($examples -join ', ') + ")"
        }
    }

    $consistencyReport = Test-WeeklyConsistency -WsHebdo $wsHebdo -LastPlayerRow $lastPlayerRow -LastWeekCol $lastWeekColHebdo
    if ([int]$consistencyReport.FatalCount -gt 0) {
        $examples = @(
            $consistencyReport.FatalIssues |
                Select-Object -First 5 |
                ForEach-Object {
                    $prefix = if ([string]::IsNullOrWhiteSpace($_.Player)) { $_.Label } else { "$($_.Label) / $($_.Player)" }
                    "${prefix}: $($_.Issue)"
                }
        )
        throw ("Contrôle de cohérence échoué: " + ($examples -join ' | '))
    }
    $consistencySummary = "Contrôle de cohérence OK: aucune anomalie bloquante."

    $weekNumber = $lastWeekColHebdo - 1
    $lastWeekColTotal = $lastWeekColHebdo + 1

    $wsTotal.Range($wsTotal.Columns(3), $wsTotal.Columns(16)).EntireColumn.Hidden = $false
    if ($lastWeekColTotal -lt 16) {
        $wsTotal.Range($wsTotal.Columns($lastWeekColTotal + 1), $wsTotal.Columns(16)).EntireColumn.Hidden = $true
    }

    $tableTotal = $wsTotal.ListObjects.Item("Table1345")
    $tableTotal.Sort.SortFields.Clear()
    [void]$tableTotal.Sort.SortFields.Add($tableTotal.ListColumns("TOTAL").Range, $xlSortOnValues, $xlDescending)
    [void]$tableTotal.Sort.SortFields.Add($tableTotal.ListColumns("Joueurs").Range, $xlSortOnValues, $xlAscending)
    $tableTotal.Sort.Header = $xlYes
    $tableTotal.Sort.Orientation = $xlTopToBottom
    $tableTotal.Sort.SortMethod = $xlPinYin
    $tableTotal.Sort.Apply()

    $excel.CalculateFull()

    $lastExportRow = 24
    for ($r = 31; $r -ge 3; $r--) {
        $txt = [string]$wsTotal.Cells($r, 2).Text
        if (-not [string]::IsNullOrWhiteSpace($txt)) {
            $lastExportRow = $r
            break
        }
    }

    $exportRange = $wsTotal.Range("A1:S$lastExportRow")
    $outputFile = Join-Path $OutputDir ("Classement semaine {0}.png" -f $weekNumber)

    $workbook.Activate() | Out-Null
    $wsTotal.Activate() | Out-Null
    $exportRange.Select() | Out-Null
    Start-Sleep -Milliseconds 400

    if (Test-Path -LiteralPath $outputFile) {
        Remove-Item -LiteralPath $outputFile -Force
    }

    $ok = $false
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $copied = $false
        $copyErrors = @()

        foreach ($mode in @(
            @{ Appearance = $xlScreen;  Format = $xlBitmap },
            @{ Appearance = $xlScreen;  Format = $xlPicture },
            @{ Appearance = $xlPrinter; Format = $xlBitmap }
        )) {
            try {
                $exportRange.CopyPicture($mode.Appearance, $mode.Format)
                $copied = $true
                break
            }
            catch {
                $copyErrors += $_.Exception.Message
            }
        }

        if (-not $copied) {
            throw ("CopyPicture impossible: " + ($copyErrors -join " | "))
        }

        Start-Sleep -Milliseconds 350

        $chartObj = $wsTotal.ChartObjects().Add($exportRange.Left, $exportRange.Top, $exportRange.Width, $exportRange.Height)
        try {
            $chartObj.Chart.Paste() | Out-Null
        }
        catch {
            $chartObj.Delete()
            $chartObj = $null
            throw "Paste dans le chart impossible."
        }

        Start-Sleep -Milliseconds 350
        [void]$chartObj.Chart.Export($outputFile, "PNG")
        $chartObj.Delete()
        $chartObj = $null

        if ((Test-Path -LiteralPath $outputFile) -and ((Get-Item -LiteralPath $outputFile).Length -gt 5000)) {
            $ok = $true
            break
        }
    }

    if (-not $ok) {
        throw "Export PNG invalide (fichier vide/blanc)."
    }

    $htmlOutputFile = Join-Path $OutputDir "classement_public.html"
    $siteConfig = [ordered]@{
        enableFinalStackCalculator = $false
        relaxDesktopHorizontalScroll = $false
    }
    $publishFingerprint = ((@([string]$PublishRepoDir, [string]$PublicUrl) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ')
    if ($publishFingerprint -match '(?i)preview') {
        $siteConfig.enableFinalStackCalculator = $true
        $siteConfig.relaxDesktopHorizontalScroll = $true
    }

    Export-PublicHtml -WsHebdo $wsHebdo -WsKills $wsKills -WsTotal $wsTotal -WsCalc $wsCalc -OutputPath $htmlOutputFile -WeekNumber $weekNumber -LastPlayerRow $lastPlayerRow -LastWeekColHebdo $lastWeekColHebdo -SiteConfig $siteConfig

    $repoSyncSummary = ""
    $gitPushSummary = ""
    if (-not [string]::IsNullOrWhiteSpace($PublishRepoDir)) {
        if (-not (Test-Path -LiteralPath $PublishRepoDir)) {
            New-Item -ItemType Directory -Path $PublishRepoDir | Out-Null
        }

        $repoIndexFile = Join-Path $PublishRepoDir 'index.html'
        $repoLatestPng = Join-Path $PublishRepoDir 'latest.png'
        $repoWeekPng = Join-Path $PublishRepoDir ("Classement semaine {0}.png" -f $weekNumber)

        Copy-Item -LiteralPath $htmlOutputFile -Destination $repoIndexFile -Force
        Copy-Item -LiteralPath $outputFile -Destination $repoLatestPng -Force
        Copy-Item -LiteralPath $outputFile -Destination $repoWeekPng -Force

        $sourceSync = Sync-SourceFilesToPublishRepo -WorkspaceDir $workspaceDir -ScriptDir $PSScriptRoot -PublishRepoDir $PublishRepoDir
        $sourceSyncSummary = [string]$sourceSync.Summary
        $sourceSyncFiles = @($sourceSync.RelativeFiles)

        $repoSyncSummary = "Fichiers web synchronisés dans: $PublishRepoDir"
        $filesToCommit = New-Object System.Collections.Generic.List[string]
        foreach ($path in @(
            'index.html',
            'latest.png',
            ("Classement semaine {0}.png" -f $weekNumber)
        )) {
            [void]$filesToCommit.Add($path)
        }
        foreach ($path in $sourceSyncFiles) {
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                [void]$filesToCommit.Add([string]$path)
            }
        }

        $gitPushSummary = Auto-CommitPushRepo -RepoDir $PublishRepoDir -WeekNumber $weekNumber -FilesToAdd $filesToCommit.ToArray()
    }

    if ($SaveWorkbook -or $killsNormalizationChangeCount -gt 0) {
        $workbook.Save()
    }
    else {
        $workbook.Saved = $true
    }

    Write-Output ("PNG généré: {0}" -f $outputFile)
    Write-Output ("HTML public généré: {0}" -f $htmlOutputFile)
    Write-Output ("Semaine détectée: {0}" -f $weekNumber)
    if (-not [string]::IsNullOrWhiteSpace($consistencySummary)) {
        Write-Output $consistencySummary
    }
    if (-not [string]::IsNullOrWhiteSpace($killsNormalizationSummary)) {
        Write-Output $killsNormalizationSummary
    }
    if (-not [string]::IsNullOrWhiteSpace($repoSyncSummary)) {
        Write-Output $repoSyncSummary
    }
    if (-not [string]::IsNullOrWhiteSpace($sourceSyncSummary)) {
        Write-Output $sourceSyncSummary
    }
    if (-not [string]::IsNullOrWhiteSpace($gitPushSummary)) {
        Write-Output $gitPushSummary
    }

    if (-not [string]::IsNullOrWhiteSpace($EmailTo)) {
        $subject = [IO.Path]::GetFileName($outputFile)

        $emailBodyLines = New-Object System.Collections.Generic.List[string]
        [void]$emailBodyLines.Add("Classement en pièce jointe.")
        [void]$emailBodyLines.Add("")
        if (-not [string]::IsNullOrWhiteSpace($PublicUrl)) {
            [void]$emailBodyLines.Add("URL publique: $PublicUrl")
        }
        else {
            [void]$emailBodyLines.Add("URL publique: (non configurée)")
            [void]$emailBodyLines.Add("Page locale à publier: $htmlOutputFile")
        }
        if (-not [string]::IsNullOrWhiteSpace($repoSyncSummary)) {
            [void]$emailBodyLines.Add("")
            [void]$emailBodyLines.Add($repoSyncSummary)
        }
        if (-not [string]::IsNullOrWhiteSpace($sourceSyncSummary)) {
            [void]$emailBodyLines.Add($sourceSyncSummary)
        }
        if (-not [string]::IsNullOrWhiteSpace($consistencySummary)) {
            [void]$emailBodyLines.Add($consistencySummary)
        }
        if (-not [string]::IsNullOrWhiteSpace($killsNormalizationSummary)) {
            [void]$emailBodyLines.Add($killsNormalizationSummary)
        }
        if (-not [string]::IsNullOrWhiteSpace($gitPushSummary)) {
            [void]$emailBodyLines.Add($gitPushSummary)
        }
        $emailBody = ($emailBodyLines.ToArray() -join "`r`n")

        Send-OutlookMailWithAttachment -To $EmailTo -Subject $subject -AttachmentPath $outputFile -Body $emailBody
        Write-Output ("Email envoyé: {0} (objet: {1})" -f $EmailTo, $subject)
        if ([string]::IsNullOrWhiteSpace($PublicUrl)) {
            Write-Output "Info: définis -PublicUrl (ou variable d'environnement POKER_PUBLIC_URL) pour l'ajouter au courriel."
        }
    }
}
finally {
    if ($null -ne $chartObj) {
        try { $chartObj.Delete() } catch {}
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }

    Release-ComObjectSafe $chartObj
    Release-ComObjectSafe $tableTotal
    Release-ComObjectSafe $wsCalc
    Release-ComObjectSafe $wsTotal
    Release-ComObjectSafe $wsKills
    Release-ComObjectSafe $wsHebdo
    Release-ComObjectSafe $workbook
    Release-ComObjectSafe $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}












