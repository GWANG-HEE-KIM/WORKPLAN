#requires -Version 5.1
<#
  월 주조 생산계획 수립 프로그램
  - 수주(월별 출하계획) + 금형/가공 재고 + 소재 매핑표 + 전월 계획 시트(제품별 일일 생산수량)를 읽어
    품번 매칭 → 소재 단위 필요량 산출 → 일별 주조 계획을 만들고 Excel/CSV로 저장한다.
  - 설정은 같은 폴더의 설정.json 에서 바꾼다.
#>
param(
  [string]$ConfigPath = (Join-Path $PSScriptRoot '설정.json'),
  [int]$Year = 0,
  [int]$Month = 0
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

# ============================================================ 공통 함수
function Num($v) { $d = 0.0; if ([double]::TryParse("$v", [ref]$d)) { $d } else { 0 } }
function Base($pn) { ("$pn" -replace '\s*\([^)]*\)', '').Trim() }
function Norm($pn) { ((Base $pn).ToUpper() -replace '[\s\-_./]', '') }
# 끝의 개정 표기를 뗀 키: 1401-850066-C03 → 1401850066, 18486-002 → 18486
function TrimKey($pn) { $b = Base ("$pn" -replace '\s+(LH|RH)\b', ''); if ($b -match '^(.+)-[^-]+$') { Norm $Matches[1] } else { Norm $b } }
function Squash($s) { "$s" -replace '\s', '' }
function ColName($n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char][int](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
function ColIndex($s) { $n = 0; foreach ($ch in $s.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; $n }
function Csv($v) { $s = "$v"; if ($s -match '[",\r\n]' -or $s -ne $s.Trim()) { '"' + $s.Replace('"', '""') + '"' } else { $s } }
function Write-Csv($path, $cols, $rows) {
  $sb = New-Object Text.StringBuilder
  [void]$sb.AppendLine((($cols | ForEach-Object { Csv $_ }) -join ','))
  foreach ($r in $rows) { [void]$sb.AppendLine((($cols | ForEach-Object { Csv $r.$_ }) -join ',')) }
  [IO.File]::WriteAllText($path, $sb.ToString(), (New-Object Text.UTF8Encoding $true))
}
# Excel에서 다시 저장된 CSV(탭 구분, CP949)도 읽는다
function Read-AnyCsv($path) {
  $bytes = [IO.File]::ReadAllBytes($path)
  $enc = if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB) { [Text.Encoding]::UTF8 } else { [Text.Encoding]::GetEncoding(949) }
  $txt = [IO.File]::ReadAllText($path, $enc)
  $txt | ConvertFrom-Csv -Delimiter $(if ($txt.Split("`n")[0] -match "`t") { "`t" } else { ',' })
}
function Step($msg) { Write-Host ''; Write-Host "▶ $msg" -ForegroundColor Cyan }

# ============================================================ xlsx 읽기 (Excel 없이 XML 직접)
$script:TempRoot = Join-Path $env:TEMP ("prodplan_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $script:TempRoot | Out-Null

function Open-Xlsx($path) {
  if (-not (Test-Path -LiteralPath $path)) { throw "파일이 없습니다: $path" }
  $dir = Join-Path $script:TempRoot ([guid]::NewGuid().ToString('N'))
  $zip = "$dir.zip"
  Copy-Item -LiteralPath $path $zip                          # 열려 있는 파일도 복사본으로 읽는다
  Expand-Archive $zip -DestinationPath $dir -Force
  $strings = New-Object System.Collections.Generic.List[string]
  if (Test-Path "$dir\xl\sharedStrings.xml") {
    $r = [Xml.XmlReader]::Create("$dir\xl\sharedStrings.xml")
    while ($r.Read()) {
      if ($r.NodeType -eq 'Element' -and $r.LocalName -eq 'si') {
        $sub = $r.ReadSubtree(); $sb = New-Object Text.StringBuilder
        while ($sub.Read()) { if ($sub.NodeType -eq 'Element' -and $sub.LocalName -eq 't') { [void]$sb.Append($sub.ReadElementContentAsString()) } }
        $strings.Add(($sb.ToString() -replace "[`r`n`t]", ' '))
      }
    }
    $r.Close()
  }
  [xml]$wb = Get-Content "$dir\xl\workbook.xml" -Raw -Encoding UTF8
  [xml]$rels = Get-Content "$dir\xl\_rels\workbook.xml.rels" -Raw -Encoding UTF8
  $target = @{}; foreach ($x in $rels.Relationships.Relationship) { $target[$x.Id] = $x.Target }
  $sheets = [ordered]@{}
  foreach ($s in $wb.workbook.sheets.sheet) {
    $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $t = $target[$rid] -replace '^/xl/', ''
    $sheets[$s.name] = Join-Path "$dir\xl" ($t -replace '/', '\')
  }
  [pscustomobject]@{ Path = $path; Strings = $strings; Sheets = $sheets }
}

# 시트 → 행 목록. 각 행은 @{ _r = 행번호; A = 값; B = 값 ... } (수식은 저장된 계산값)
function Read-Sheet($book, [string]$name) {
  $file = $null
  if ($name) {
    foreach ($k in $book.Sheets.Keys) { if ($k.Trim() -eq $name.Trim()) { $file = $book.Sheets[$k] } }
    if (-not $file) { throw "시트를 찾을 수 없습니다: '$name' (파일: $($book.Path))`n  있는 시트: $(@($book.Sheets.Keys) -join ', ')" }
  } else { $file = @($book.Sheets.Values)[0] }
  $rows = New-Object System.Collections.Generic.List[object]
  $r = [Xml.XmlReader]::Create($file); $cur = $null
  while ($r.Read()) {
    if ($r.NodeType -ne 'Element') { continue }
    if ($r.LocalName -eq 'row') { $cur = @{ _r = [int]$r.GetAttribute('r') }; $rows.Add($cur) }
    elseif ($r.LocalName -eq 'c') {
      $ref = $r.GetAttribute('r'); $t = $r.GetAttribute('t'); $v = $null
      if (-not $r.IsEmptyElement) {
        $sub = $r.ReadSubtree()
        while ($sub.Read()) { if ($sub.NodeType -eq 'Element' -and ($sub.LocalName -eq 'v' -or $sub.LocalName -eq 't')) { $v = $sub.ReadElementContentAsString() } }
      }
      if ($null -eq $v -or $v -eq '') { continue }
      if ($t -eq 's') { $v = $book.Strings[[int]$v] }
      $v = ($v -replace "[`r`n`t]", ' ').Trim()
      if ($v -ne '' -and $v -ne '-') { $cur[($ref -replace '\d+$', '')] = $v }
    }
  }
  $r.Close()
  , $rows
}

# 머리글 행 찾기: 지정한 이름이 모두 있는 첫 행 → @{ Row; Col = @{이름 = 열} }
function Find-Header($rows, [string[]]$names, [int]$maxRow = 15) {
  foreach ($row in $rows) {
    if ($row._r -gt $maxRow) { break }
    $map = @{}
    foreach ($k in $row.Keys) { if ($k -ne '_r') { $n = Squash $row[$k]; foreach ($want in $names) { if ($n -eq (Squash $want) -and -not $map[$want]) { $map[$want] = $k } } } }
    if (@($names | Where-Object { -not $map[$_] }).Count -eq 0) { return @{ Row = $row._r; Col = $map; Cells = $row } }
  }
  throw "머리글을 찾을 수 없습니다: $($names -join ', ')"
}

# ============================================================ 설정
if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "설정 파일이 없습니다: $ConfigPath" }
$cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$progDir = Split-Path -Parent (Resolve-Path -LiteralPath $ConfigPath)
$dataDir = if ([IO.Path]::IsPathRooted($cfg.'자료폴더')) { $cfg.'자료폴더' } else { [IO.Path]::GetFullPath((Join-Path $progDir $cfg.'자료폴더')) }
if (-not $Year) { $Year = [int]$cfg.'계획연도' }
if (-not $Month) { $Month = [int]$cfg.'계획월' }
function DataFile($name) { if ([IO.Path]::IsPathRooted($name)) { $name } else { Join-Path $dataDir $name } }

$planStart = [datetime]::new($Year, $Month, 1)
$prevStart = $planStart.AddMonths(-1); $nextStart = $planStart.AddMonths(1)
$daysInMonth = [datetime]::DaysInMonth($Year, $Month)
$holidays = @($cfg.'휴무일' | ForEach-Object { [int]$_ })
$workDays = @(1..$daysInMonth | Where-Object { $dw = [datetime]::new($Year, $Month, $_).DayOfWeek; $dw -ne 'Saturday' -and $dw -ne 'Sunday' -and $_ -notin $holidays })
$offDays = @(1..$daysInMonth | Where-Object { $_ -notin $workDays })
$safety = [double]$cfg.'안전율'; $nextShare = [double]$cfg.'익월선행비율'
$machOkTypes = @($cfg.'가공재고인정_품목구분'); $machOkSuffix = @($cfg.'가공재고인정_품번표시')
$mLabel = "${Month}월"; $nLabel = "$($nextStart.Month)월"
$outDir = Join-Path $progDir (Join-Path $cfg.'출력폴더' ('{0:yyyy-MM}' -f $planStart))
New-Item -ItemType Directory -Force $outDir | Out-Null

Write-Host "==============================================="
Write-Host " ${Year}년 ${Month}월 주조 생산계획 수립"
Write-Host "  근무일 $($workDays.Count)일 / 안전율 $safety / $nLabel 선행 $([int]($nextShare*100))%"
Write-Host "==============================================="

try {
# ============================================================ 1. 입력 파일 읽기
Step "1/5 입력 파일 읽기"
# --- 수주: 업체명/품명/품번/단중 + 당월·익월 출하계획 (머리글의 날짜 열로 찾음)
$ordBook = Open-Xlsx (DataFile $cfg.'수주파일')
$ordRows = Read-Sheet $ordBook $cfg.'수주시트'
$oh = Find-Header $ordRows @('업체명', '품명', '품번', '단중')
$serial = { param($d) [int]($d - [datetime]'1899-12-30').TotalDays }
$curCol = $null; $nextCol = $null
foreach ($k in $oh.Cells.Keys) { if ($k -eq '_r') { continue }; $v = Num $oh.Cells[$k]; if ($v -eq (& $serial $planStart)) { $curCol = $k }; if ($v -eq (& $serial $nextStart)) { $nextCol = $k } }
if (-not $curCol) { throw "수주 파일 머리글에서 $Year-$Month 열을 찾을 수 없습니다." }
# '구분' 열이 둘(국내/해외, 출하계획/납품실적)이라 '출하계획' 값이 있는 열을 고른다
$kindCol = $null
foreach ($k in $oh.Cells.Keys) { if ($k -ne '_r' -and (Squash $oh.Cells[$k]) -eq '구분') { if ($ordRows | Where-Object { $_[$k] -eq '출하계획' } | Select-Object -First 1) { $kindCol = $k } } }
$orders = @(foreach ($r in $ordRows) {
  if ($r._r -le $oh.Row -or -not $r[$oh.Col['품번']]) { continue }
  if ($kindCol -and $r[$kindCol] -ne '출하계획') { continue }
  [pscustomobject]@{
    Cust = $r[$oh.Col['업체명']]; Name = $r[$oh.Col['품명']]; PN = $r[$oh.Col['품번']]; Key = (Norm $r[$oh.Col['품번']])
    W = (Num $r[$oh.Col['단중']]); Cur = (Num $r[$curCol]); Next = $(if ($nextCol) { Num $r[$nextCol] } else { 0 })
  }
})
Write-Host "  수주: $($orders.Count)건 ($mLabel 출하 합계 $('{0:N0}' -f ($orders | Measure-Object Cur -Sum).Sum), $nLabel $('{0:N0}' -f ($orders | Measure-Object Next -Sum).Sum))"
if (-not $nextCol) { Write-Host "  ※ 수주 파일에 $nLabel 열이 없어 선행분은 0으로 계산합니다." -ForegroundColor Yellow }

# --- 재고: 고객사/품번/품명/품목구분/SUM
function Load-Stock($file, $sheet) {
  $book = Open-Xlsx (DataFile $file); $rows = Read-Sheet $book $sheet
  $h = Find-Header $rows @('고객사', '품번', '품명', 'SUM')
  $typeCol = $null; foreach ($k in $h.Cells.Keys) { if ($k -ne '_r' -and (Squash $h.Cells[$k]) -eq '품목구분') { $typeCol = $k } }
  $by = @{}; $seq = 0
  foreach ($r in $rows) {
    if ($r._r -le $h.Row -or -not $r[$h.Col['품번']]) { continue }
    $pn = $r[$h.Col['품번']]
    $o = [pscustomobject]@{ Cust = $r[$h.Col['고객사']]; Name = $r[$h.Col['품명']]; PN = $pn; Key = (Norm $pn)
      Sfx = $(if ($pn -match '\(([^)]*)\)') { $Matches[1] } else { '' }); Type = $(if ($typeCol) { $r[$typeCol] } else { '' }); Qty = (Num $r[$h.Col['SUM']]); Seq = $seq++ }
    if (-not $by[$o.Key]) { $by[$o.Key] = New-Object System.Collections.ArrayList }
    [void]$by[$o.Key].Add($o)
  }
  $by
}
$moldBy = Load-Stock $cfg.'금형재고파일' $cfg.'금형재고시트'
$machBy = Load-Stock $cfg.'가공재고파일' $cfg.'가공재고시트'
Write-Host "  금형 재고: 기본품번 $($moldBy.Count)개 / 가공 재고: 기본품번 $($machBy.Count)개"
function MachOk($row) { ($row.Type -in $machOkTypes) -or ($machOkSuffix | Where-Object { $row.Sfx -match [regex]::Escape($_) }) }

# --- 전월 계획 시트: 품명·단중·단가·일별 생산수량 → 제품별 일일 생산수량, 요일별 수량
$prevBook = Open-Xlsx (DataFile $cfg.'전월계획파일')
$prevRows = Read-Sheet $prevBook $cfg.'전월계획시트'
$ph = Find-Header $prevRows @('품명', '품번', '단중')
$priceCol = $null; foreach ($k in $ph.Cells.Keys) { if ($k -ne '_r' -and $ph.Cells[$k] -match '단가') { $priceCol = $k } }
$dayCol1 = $null
foreach ($r in $prevRows) {
  if ($r._r -lt $ph.Row -or $r._r -gt $ph.Row + 4) { continue }
  # 1,2,3 이 연속된 열 중 가장 왼쪽 (오른쪽에 중량용 일별 열이 또 있을 수 있음)
  $hits = @(foreach ($k in $r.Keys) { if ($k -ne '_r' -and $r[$k] -eq '1') { $i = ColIndex $k; if ($r[(ColName ($i + 1))] -eq '2' -and $r[(ColName ($i + 2))] -eq '3') { $i } } })
  if ($hits) { $dayCol1 = ($hits | Measure-Object -Minimum).Minimum; break }
}
if (-not $dayCol1) { throw "전월 계획 시트에서 일별(1,2,3…) 머리글을 찾을 수 없습니다." }
$prevDays = [datetime]::DaysInMonth($prevStart.Year, $prevStart.Month)
$master = @{}; $master2 = @{}
foreach ($r in $prevRows) {
  if ($r._r -le $ph.Row + 1 -or -not $r[$ph.Col['품번']] -or "$($r['A'])" -notmatch '^\d+$') { continue }
  $qty = @(for ($d = 1; $d -le $prevDays; $d++) { Num $r[(ColName ($dayCol1 + $d - 1))] })
  $nz = @($qty | Where-Object { $_ -gt 0 })
  $mode = if ($nz) { [double](($nz | Group-Object | Sort-Object Count, { [double]$_.Name } -Descending | Select-Object -First 1).Name) } else { $null }
  $week = foreach ($wd in 1..5) {   # 월~금 요일별 최빈값
    $v = @(for ($d = 1; $d -le $prevDays; $d++) { if ([int][datetime]::new($prevStart.Year, $prevStart.Month, $d).DayOfWeek -eq $wd -and $qty[$d - 1] -gt 0) { $qty[$d - 1] } })
    if ($v) { [double](($v | Group-Object | Sort-Object Count, { [double]$_.Name } -Descending | Select-Object -First 1).Name) } else { 0 }
  }
  $rec = [pscustomobject]@{ Name = $r[$ph.Col['품명']]; PN = $r[$ph.Col['품번']]; W = (Num $r[$ph.Col['단중']]); Price = $(if ($priceCol) { Num $r[$priceCol] } else { 0 }); Daily = $mode; Week = @($week) }
  $k = Norm ($rec.PN -replace '\s+(LH|RH)\b', '')
  if (-not $master[$k] -or ($mode -and -not $master[$k].Daily)) { $master[$k] = $rec }
  $k2 = TrimKey $rec.PN
  if (-not $master2[$k2] -or ($mode -and -not $master2[$k2].Daily)) { $master2[$k2] = $rec }
}
Write-Host "  전월 계획: 품목 $($master.Count)개 (일일 생산수량 있는 품목 $(@($master.Values | Where-Object Daily).Count)개)"
function Lookup($pn) {
  $k = Norm $pn; $t = TrimKey $pn
  $c = @($master[$k], $master[$t], $master2[$t] | Where-Object { $_ })
  $wd = @($c | Where-Object Daily); if ($wd) { return $wd[0] } elseif ($c) { return $c[0] }
  if ($k.Length -ge 6) {
    $hits = @($master.Keys | Where-Object { $_.Length -ge 6 -and ($_.StartsWith($k) -or $k.StartsWith($_)) -and $master[$_].Daily })
    if ($hits.Count -eq 1) { return $master[$hits[0]] }
  }
}

# --- 소재 매핑표 (완제품품번 → 확정 소재품번)
$mapPath = Join-Path $progDir $cfg.'소재매핑표'
$mapping = @{}
if (Test-Path -LiteralPath $mapPath) {
  foreach ($m in (Read-AnyCsv $mapPath)) { $f = "$($m.'완제품품번(수주)')".Trim(); $c = "$($m.'확정_소재품번')".Trim(); if ($f -and $c) { $mapping[(Norm $f)] = $c } }
}
Write-Host "  소재 매핑표: 확정 $($mapping.Count)건"

# ============================================================ 2. 품번 매칭
Step "2/5 수주 ↔ 금형·가공 재고 품번 매칭"
function Tokens($s) {
  $syn = @{ '타이페이' = 'TAIPEI'; '싱가폴' = 'SINGAPORE'; 'ASSEMBLY' = 'ASSY'; 'STERING' = 'STEERING' }
  @("$s".ToUpper() -split '[^0-9A-Z가-힣]+' | Where-Object { $_ } | ForEach-Object { if ($syn[$_]) { $syn[$_] } else { $_ } }) | Sort-Object -Unique
}
function Side($tok) { ($tok | Where-Object { $_ -in 'LH', 'RH', 'PORT', 'STBD' }) -join ',' }
function Lcs($a, $b) {
  $best = 0
  for ($i = 0; $i -lt $a.Length; $i++) { for ($j = 0; $j -lt $b.Length; $j++) {
    $k = 0; while ($i + $k -lt $a.Length -and $j + $k -lt $b.Length -and $a[$i + $k] -eq $b[$j + $k]) { $k++ }
    if ($k -gt $best) { $best = $k } } }
  $best
}
function Fuzzy($key, $table) {
  if ($key.Length -lt 5) { return $null }
  $hits = @($table.Keys | Where-Object { $_.Length -ge 5 -and ($_.StartsWith($key) -or $key.StartsWith($_)) })
  if ($hits.Count -eq 1) { $hits[0] } else { $null }
}
$usedMold = @{}
foreach ($o in $orders) {
  $o | Add-Member -NotePropertyMembers @{ MoldKeys = @(); Method = @(); Cand = ''; Memo = '' }
  $mk = if ($moldBy.ContainsKey($o.Key)) { $o.Key } else { Fuzzy $o.Key $moldBy }
  if ($mk) { $o.MoldKeys = @($mk); $usedMold[$mk] = 1; $o.Method += $(if ($mk -eq $o.Key) { '금형:정확' } else { '금형:유사' }) }
}
# 소재 품번 ≠ 완제품 품번: 같은 고객사 안에서 품번 공통부분 + 품명 유사도로 연결 (점수 높은 쌍부터)
$pairs = @()
foreach ($k in @($moldBy.Keys)) {
  if ($usedMold[$k]) { continue }
  $m = $moldBy[$k][0]; $mt = Tokens $m.Name; $ms = Side $mt
  foreach ($o in $orders) {
    if ($o.MoldKeys -or $o.Cust -ne $m.Cust) { continue }
    $ot = Tokens $o.Name; $os = Side $ot
    if ($ms -and $os -and $ms -ne $os) { continue }
    $inter = @($mt | Where-Object { $_ -in $ot }).Count; $union = @(@($mt) + @($ot) | Sort-Object -Unique).Count
    $j = $(if ($union) { $inter / $union } else { 0 }); $l = Lcs $k $o.Key
    $tier = $(if ($l -ge 5 -and $j -ge 0.5) { 1 } elseif ($l -ge 7 -and $j -gt 0) { 2 } elseif ($j -eq 1) { 3 } else { 0 })
    if ($tier) { $pairs += [pscustomobject]@{ Key = $k; O = $o; Tier = $tier; L = $l; J = $j } }
  }
}
$linkSig = @{}
foreach ($p in ($pairs | Where-Object Tier -lt 3 | Sort-Object Tier, @{ e = 'L'; d = $true }, @{ e = 'J'; d = $true })) {
  if ($p.O.MoldKeys) { continue }
  $sig = "$($p.Tier)/$($p.L)/$($p.J)"
  if ($linkSig[$p.Key] -and $linkSig[$p.Key] -ne $sig) { continue }   # 한 소재 → 여러 완제품은 동점일 때만
  $p.O.MoldKeys = @($p.Key); $p.O.Method += $(if ($p.Tier -eq 1) { '소재연결:품번+품명유사' } else { '소재연결:품번유사' }); $linkSig[$p.Key] = $sig; $usedMold[$p.Key] = 1
}
# 확정 매핑표가 최우선
foreach ($o in $orders) {
  $m = $mapping[$o.Key]; if (-not $m) { continue }
  if ($m -eq '없음') { $o.MoldKeys = @(); $o.Method = @($o.Method | Where-Object { $_ -notmatch '^(금형|소재연결)' }) + '매핑표:소재없음'; continue }
  $keys = @($m -split '/' | ForEach-Object { Norm $_ } | Where-Object { $_ })
  $miss = @($keys | Where-Object { -not $moldBy.ContainsKey($_) })
  $o.MoldKeys = @($keys | Where-Object { $moldBy.ContainsKey($_) })
  $o.Method = @($o.Method | Where-Object { $_ -notmatch '^(금형|소재연결)' }) + '금형:매핑표확정'
  if ($miss) { $o.Memo = "매핑표 소재품번이 금형 재고에 없음: $($miss -join ',')" }
  $os = Side (Tokens $o.Name)
  foreach ($k in $o.MoldKeys) { $usedMold[$k] = 1; $ms = Side (Tokens $moldBy[$k][0].Name); if ($os -and $ms -and $os -ne $ms) { $o.Memo = (@($o.Memo, "좌우 불일치: 수주 $os / 소재 $($moldBy[$k][0].PN) $ms") | Where-Object { $_ }) -join ' ; ' } }
}
# 품명만 같은 소재는 후보로만 표시
foreach ($p in ($pairs | Where-Object Tier -eq 3)) {
  if ($p.O.MoldKeys) { continue }
  $g = $moldBy[$p.Key][0]
  $p.O.Cand = (@($p.O.Cand, "$(Base $g.PN)(재고$($g.Qty))") | Where-Object { $_ }) -join ' | '
}
# 가공 재고: 완제품 품번 + 연결된 소재 품번의 가공동 재고
foreach ($o in $orders) {
  $rows = @(); if ($machBy[$o.Key]) { $rows += $machBy[$o.Key] }
  foreach ($k in $o.MoldKeys) { if ($k -ne $o.Key -and $machBy[$k]) { $rows += $machBy[$k] } }
  if (-not $rows -and -not $o.MoldKeys) { $fk = Fuzzy $o.Key $machBy; if ($fk) { $rows += $machBy[$fk] } }
  $o | Add-Member -NotePropertyMembers @{ MachRows = $rows; Mold = (($o.MoldKeys | ForEach-Object { $moldBy[$_] } | Measure-Object Qty -Sum).Sum) }
  if ($rows) { $o.Method += '가공:있음' }
}
$noMold = @($orders | Where-Object { -not $_.MoldKeys })
Write-Host "  수주 $($orders.Count)건 중 소재(금형) 연결 $(@($orders | Where-Object { $_.MoldKeys }).Count)건, 미연결 $($noMold.Count)건 (후보 있음 $(@($noMold | Where-Object Cand).Count)건)"

$matchOut = foreach ($o in $orders) {
  [pscustomobject]@{
    '고객사' = $o.Cust; '품명' = $o.Name; '출하품번' = $o.PN; "${mLabel}출하" = $o.Cur; "${nLabel}출하" = $o.Next
    '소재품번' = (($o.MoldKeys | ForEach-Object { $moldBy[$_][0].PN }) -join ' / '); '금형재고' = $o.Mold
    '가공품번' = (($o.MachRows | ForEach-Object { $_.PN }) -join ' | ')
    '가공재고_전체' = (($o.MachRows | Measure-Object Qty -Sum).Sum); '가공재고_인정' = (($o.MachRows | Where-Object { MachOk $_ } | Measure-Object Qty -Sum).Sum)
    '매칭방법' = ($o.Method -join ' / '); '소재후보_품명일치' = $o.Cand; '검토메모' = $o.Memo
  }
}
$matchCsv = Join-Path $outDir '품번매칭_결과.csv'
Write-Csv $matchCsv @('고객사', '품명', '출하품번', "${mLabel}출하", "${nLabel}출하", '소재품번', '금형재고', '가공품번', '가공재고_전체', '가공재고_인정', '매칭방법', '소재후보_품명일치', '검토메모') $matchOut

# 매핑표에 새 미연결 품번 추가 (기존 행은 그대로 둔다)
$newMap = @($noMold | Where-Object { -not $mapping.ContainsKey($_.Key) -and ($_.Cur -gt 0 -or $_.Next -gt 0) })
if ($newMap) {
  $existing = if (Test-Path -LiteralPath $mapPath) { @(Read-AnyCsv $mapPath) } else { @() }
  $already = @{}; foreach ($e in $existing) { $already[(Norm $e.'완제품품번(수주)')] = 1 }
  $add = @($newMap | Where-Object { -not $already[$_.Key] } | ForEach-Object {
    [pscustomobject]@{ '고객사' = $_.Cust; '완제품품번(수주)' = $_.PN; '품명' = $_.Name; '후보' = $_.Cand; '확정_소재품번' = ''; '비고' = "$('{0:yyyy-MM}' -f $planStart) 추가 ($mLabel 출하 $($_.Cur))" } })
  if ($add) {
    $cols = @('고객사', '완제품품번(수주)', '품명', '후보', '확정_소재품번', '비고')
    $all = @($existing | ForEach-Object { $e = $_; $o = [ordered]@{}; foreach ($c in $cols) { $o[$c] = $e.$c }; [pscustomobject]$o }) + $add
    try { Write-Csv $mapPath $cols $all; Write-Host "  소재 매핑표에 확정이 필요한 품번 $($add.Count)건을 추가했습니다 → $($cfg.'소재매핑표')" -ForegroundColor Yellow }
    catch { Write-Host "  ※ 소재 매핑표가 열려 있어 추가하지 못했습니다. 파일을 닫고 다시 실행하세요." -ForegroundColor Yellow }
  }
}

# ============================================================ 3. 소재 단위 필요량
Step "3/5 소재 단위 필요량 계산"
$parent = @{}
function Find($x) { while ($parent[$x] -ne $x) { $x = $parent[$x] }; $x }
for ($i = 0; $i -lt $orders.Count; $i++) { $parent["i$i"] = "i$i" }
$owner = @{}
for ($i = 0; $i -lt $orders.Count; $i++) {
  foreach ($k in $orders[$i].MoldKeys) { if ($owner[$k]) { $a = Find "i$i"; $b = Find $owner[$k]; if ($a -ne $b) { $parent[$a] = $b } } else { $owner[$k] = "i$i" } }
}
$groups = @{}
for ($i = 0; $i -lt $orders.Count; $i++) { $g = Find "i$i"; if (-not $groups[$g]) { $groups[$g] = @() }; $groups[$g] += $orders[$i] }

$plans = @(foreach ($g in $groups.Values) {
  $keys = @($g | ForEach-Object { $_.MoldKeys } | Select-Object -Unique)
  $machU = @{}; foreach ($o in $g) { foreach ($m in $o.MachRows) { if (MachOk $m) { $machU[$m.PN] = $m.Qty } } }
  $mold = (@($keys | ForEach-Object { $moldBy[$_] }) | Measure-Object Qty -Sum).Sum
  $mach = ($machU.Values | Measure-Object -Sum).Sum
  $cur = ($g | Measure-Object Cur -Sum).Sum; $next = ($g | Measure-Object Next -Sum).Sum
  $pre = [math]::Ceiling($next * $nextShare)
  if ($cur -le 0 -and $pre -le 0) { continue }
  $need = [math]::Max(0, [math]::Ceiling(($cur + $pre) * $safety - $mold - $mach))
  $castPNs = @($keys | ForEach-Object { Base $moldBy[$_][0].PN })
  $ms = @(@($keys | ForEach-Object { $moldBy[$_][0].PN }) + @($g | ForEach-Object { $_.PN }) | ForEach-Object { Lookup $_ } | Where-Object { $_ })
  $ms = @(@($ms | Where-Object Daily) + @($ms | Where-Object { -not $_.Daily }))
  $top = $g | Sort-Object Cur -Descending | Select-Object -First 1
  [pscustomobject]@{
    Cust = $top.Cust; Orders = $g; Keys = $keys; NoCast = (-not $keys)
    CastPN = $(if ($keys) { $castPNs -join ' / ' } else { "(소재 미확인) $($top.PN)" })
    Name = $(if ($ms) { $ms[0].Name } elseif ($keys) { $moldBy[$keys[0]][0].Name } else { $top.Name })
    W = $(if ($ms -and $ms[0].W) { $ms[0].W } else { $top.W }); Price = $(if ($ms) { $ms[0].Price } else { 0 })
    Daily = $(if ($ms) { $ms[0].Daily } else { $null }); Week = $(if ($ms) { $ms[0].Week } else { $null })
    Stock = $mold + $mach; Cur = $cur; Next = $next; Pre = $pre; Need = $need; Plan = @{}; Short = 0; Rule = ''; Alt = $false; No = 0
  }
})
Write-Host "  소재 품목 $($plans.Count)개 / 생산 필요 $(@($plans | Where-Object Need -gt 0).Count)개 / 필요량 합계 $('{0:N0}' -f ($plans | Measure-Object Need -Sum).Sum)"

# ============================================================ 4. 일별 배치
Step "4/5 일별 주조 계획 배치"
$load = @{}; foreach ($d in $workDays) { $load[$d] = 0 }
$weekOf = @{}; $wi = -1; $lastMonday = $null
foreach ($d in $workDays) { $dt = [datetime]::new($Year, $Month, $d); $mon = $dt.AddDays( - (([int]$dt.DayOfWeek + 6) % 7)); if ($mon -ne $lastMonday) { $wi++; $lastMonday = $mon }; $weekOf[$d] = $wi }

# 4a. 격주 생산 라인 (설정.json '격주생산')
foreach ($line in @($cfg.'격주생산')) {
  if (-not $line) { continue }
  $setA = @($line.A | ForEach-Object { Norm $_ }); $setB = @($line.B | ForEach-Object { Norm $_ })
  $firstIsA = ("$($line.'첫주')" -eq 'A')
  foreach ($p in ($plans | Where-Object { $_.Need -gt 0 })) {
    $inA = @($p.Keys | Where-Object { $_ -in $setA }).Count -gt 0; $inB = @($p.Keys | Where-Object { $_ -in $setB }).Count -gt 0
    if (-not ($inA -or $inB)) { continue }
    $myParity = if ($inA -eq $firstIsA) { 0 } else { 1 }        # 0 = 첫째 주부터
    $days = @($workDays | Where-Object { $weekOf[$_] % 2 -eq $myParity })
    $sum = 0
    foreach ($d in $days) {
      if ($sum -ge $p.Need) { break }
      $wd = [int][datetime]::new($Year, $Month, $d).DayOfWeek
      $q = if ($p.Week -and $wd -ge 1 -and $wd -le 5) { $p.Week[$wd - 1] } else { 0 }
      if (-not $q) { $q = $p.Daily }; if (-not $q) { continue }
      $p.Plan[$d] = $q; $load[$d] += $q; $sum += $q
    }
    $p.Alt = $true; $p.Short = [math]::Max(0, $p.Need - $sum)
    $wk = (@($days | ForEach-Object { $weekOf[$_] } | Select-Object -Unique) | ForEach-Object { $w = $_; "$(@($days | Where-Object { $weekOf[$_] -eq $w })[0])일주" }) -join '·'
    $p.Rule = "격주($($line.'이름')): $wk, 전월 요일별 $(@($p.Week) -join '·')" + $(if ($p.Short) { " / 미달 $($p.Short)" } else { '' })
  }
}
# 4b. 나머지: 전월 일일 생산수량 고정, 연속 근무일을 부하가 가장 적은 구간에 배치
foreach ($p in ($plans | Where-Object { $_.Need -gt 0 -and -not $_.NoCast -and -not $_.Alt } | Sort-Object Need -Descending)) {
  if ($p.Daily) {
    $daily = $p.Daily; $p.Short = [math]::Max(0, $p.Need - $daily * $workDays.Count)
    $p.Rule = "전월 일생산 $daily" + $(if ($p.Short) { " / $($workDays.Count)일 생산해도 미달 $($p.Short)" } else { '' })
  } else {
    $daily = [math]::Max(10, [math]::Ceiling($p.Need / $workDays.Count / 10) * 10); $p.Rule = "전월 생산이력 없음 → 월 균등배분 일 $daily"
  }
  $n = [math]::Min($workDays.Count, [math]::Ceiling($p.Need / $daily))
  $best = 0; $bestLoad = [double]::MaxValue
  for ($s = 0; $s -le $workDays.Count - $n; $s++) { $l = 0; for ($j = $s; $j -lt $s + $n; $j++) { $l += $load[$workDays[$j]] }; if ($l -lt $bestLoad) { $bestLoad = $l; $best = $s } }
  for ($j = $best; $j -lt $best + $n; $j++) { $p.Plan[$workDays[$j]] = $daily; $load[$workDays[$j]] += $daily }
}
foreach ($p in $plans) { if ($p.NoCast -and $p.Need -gt 0) { $p.Rule = '소재 품번 확인 후 계획 (일별 미배정)' } elseif ($p.Need -eq 0) { $p.Rule = '재고로 충족' } }
$prodSum = ($plans | ForEach-Object { ($_.Plan.Values | Measure-Object -Sum).Sum } | Measure-Object -Sum).Sum
Write-Host "  주조 계획 합계 $('{0:N0}' -f $prodSum)개 / 하루 부하 $(($load.Values | Measure-Object -Minimum).Minimum)~$(($load.Values | Measure-Object -Maximum).Maximum)개"

# ============================================================ 5. Excel 출력
Step "5/5 Excel 작성"
$custOrder = @($plans | Group-Object Cust | Sort-Object { ($_.Group | Measure-Object Cur -Sum).Sum } -Descending | ForEach-Object Name)
$xlsxPath = Join-Path $outDir ("주조생산계획_{0:yyyy-MM}_{1:MMdd-HHmm}.xlsx" -f $planStart, (Get-Date))
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  # 이 PC의 COM 바인더는 Value2에 문자열만 받으므로 숫자·수식은 Formula 문자열로 넣는다
  function Put($sh, $r, $c, $v) {
    if ($null -eq $v -or "$v" -eq '') { return }
    $cell = $sh.Cells.Item($r, $c)
    if ($v -is [string] -and $v -notmatch '^=') { $cell.NumberFormat = '@'; $cell.Value2 = $v } else { $cell.Formula = [string]$v }
  }
  $ws = $wb.Worksheets.Item(1); $ws.Name = "$mLabel 주조 생산계획"; $ws.Cells.Font.Name = '맑은 고딕'; $ws.Cells.Font.Size = 10
  $dayCol = 15; $sumCol = $dayCol + $daysInMonth; $carryCol = $sumCol + 1; $noteCol = $carryCol + 1
  $dF = ColName $dayCol; $dL = ColName ($dayCol + $daysInMonth - 1); $cS = ColName $sumCol; $cC = ColName $carryCol; $cN = ColName $noteCol
  Put $ws 1 1 ("{0}. {1}. 1.~" -f $Year, $Month); $ws.Range('A1').Font.Size = 12
  Put $ws 2 1 '금형공장'; Put $ws 2 3 "근무일수 : $($workDays.Count) 일"; $ws.Range('A2:C2').Font.Size = 14; $ws.Range('A2:C2').Font.Bold = $true
  Put $ws 2 $dayCol ("※ 재고=금형 전부+가공({0}), 생산=({1} 출하+{2} 출하×{3}%)×{4}−재고, 일생산=전월 계획 제품별 수량 (일별 칸 수정 시 합계 자동 계산)" -f (@($machOkTypes) + @($machOkSuffix) -join '·'), $mLabel, $nLabel, [int]($nextShare * 100), $safety)
  $ws.Cells.Item(2, $dayCol).Font.Size = 9; $ws.Cells.Item(2, $dayCol).Font.Color = 0x595959
  foreach ($x in @(@(1, 'NO'), @(2, '품 명'), @(3, '품 번 (소재)'), @(4, '단중'), @(5, '소재 단가'), @(14, '구분'))) { Put $ws 3 $x[0] $x[1]; $ws.Range("$(ColName $x[0])3:$(ColName $x[0])4").Merge() }
  $grp = @(@(6, '재고 (금형+가공 인정분)'), @(8, "$mLabel 출하계획 (영업요청)"), @(10, "$nLabel 선행분 ($nLabel 출하×$([int]($nextShare*100))%)"), @(12, "$mLabel 주조 생산계획 총 생산"))
  foreach ($x in $grp) { Put $ws 3 $x[0] $x[1]; $ws.Range("$(ColName $x[0])3:$(ColName ($x[0] + 1))3").Merge(); Put $ws 4 $x[0] '수량(EA)'; Put $ws 4 ($x[0] + 1) '중량(kg)' }
  Put $ws 3 $dayCol '일별 생산 계획(주조)'; $ws.Range("${dF}3:${dL}3").Merge()
  for ($d = 1; $d -le $daysInMonth; $d++) { Put $ws 4 ($dayCol + $d - 1) ([string]$d); if ($d -in $offDays) { $ws.Cells.Item(4, $dayCol + $d - 1).Font.Color = 0x0000FF } }
  Put $ws 3 $sumCol '주조합계'; $ws.Range("${cS}3:${cS}4").Merge()
  Put $ws 3 $carryCol "$nLabel 이월재고 예상"; $ws.Range("${cC}3:${cC}4").Merge()
  Put $ws 3 $noteCol "출하품번($mLabel/$nLabel) / 산출 근거"; $ws.Range("${cN}3:${cN}4").Merge()
  $hdr = $ws.Range("A3:${cN}4"); $hdr.Interior.Color = 0xFFFFCC; $hdr.Font.Bold = $true; $hdr.HorizontalAlignment = -4108; $hdr.VerticalAlignment = -4108; $hdr.WrapText = $true
  $ws.Range('H3:K4').Interior.Color = 0xD6E4FC

  $r = 5; $no = 0; $subRows = @(); $numCols = @(6..13) + @($dayCol..$carryCol)
  foreach ($cu in $custOrder) {
    $first = $r
    foreach ($p in ($plans | Where-Object Cust -eq $cu | Sort-Object CastPN)) {
      $no++; $p.No = $no
      Put $ws $r 1 ([double]$no); Put $ws $r 2 $p.Name; Put $ws $r 3 $p.CastPN; Put $ws $r 4 $p.W; if ($p.Price) { Put $ws $r 5 ([math]::Round($p.Price)) }
      Put $ws $r 6 $p.Stock; Put $ws $r 7 "=F$r*D$r"; Put $ws $r 8 $p.Cur; Put $ws $r 9 "=H$r*D$r"; Put $ws $r 10 $p.Pre; Put $ws $r 11 "=J$r*D$r"
      Put $ws $r 12 "=$cS$r"; Put $ws $r 13 "=L$r*D$r"; Put $ws $r 14 '생산계획'
      foreach ($d in $p.Plan.Keys) { Put $ws $r ($dayCol + $d - 1) $p.Plan[$d] }
      Put $ws $r $sumCol "=SUM($dF${r}:$dL$r)"; Put $ws $r $carryCol "=F$r+L$r-H$r"
      Put $ws $r $noteCol ((($p.Orders | Sort-Object PN | ForEach-Object { "$($_.PN)($($_.Cur)/$($_.Next))" }) -join ', ') + ' / ' + $p.Rule)
      if ($p.NoCast) { $ws.Range("A${r}:${cN}$r").Font.Color = 0x0000C0 } elseif ($p.Short) { $ws.Range("${cN}$r").Font.Color = 0x0000C0 }
      $r++
    }
    Put $ws $r 1 "$cu 전체"; $ws.Range("A${r}:E$r").Merge()
    foreach ($c in $numCols) { $cl = ColName $c; Put $ws $r $c "=SUM(${cl}${first}:$cl$($r - 1))" }
    $subRows += $r; $r++
  }
  Put $ws $r 1 '전체 합계'; $ws.Range("A${r}:E$r").Merge()
  foreach ($c in $numCols) { $cl = ColName $c; Put $ws $r $c ("=" + (($subRows | ForEach-Object { "$cl$_" }) -join '+')) }
  $end = $r
  $all = $ws.Range("A3:${cN}$end"); $all.Borders.LineStyle = 1; $all.Borders.Color = 0x808080
  $ws.Range("A5:A$end").HorizontalAlignment = -4108; $ws.Range("D5:D$end").HorizontalAlignment = -4108; $ws.Range("N5:N$end").HorizontalAlignment = -4108
  $ws.Range("E5:M$end").NumberFormat = '#,##0;-#,##0;"-"'; $ws.Range("D5:D$end").NumberFormat = '0.00'
  $ws.Range("${dF}5:${cS}$end").NumberFormat = '#,##0;-#,##0;""'; $ws.Range("${dF}5:${cS}$end").HorizontalAlignment = -4108
  $ws.Range("${cC}5:${cC}$end").NumberFormat = '#,##0;-#,##0;"-"'
  foreach ($d in $offDays) { $ws.Range("$(ColName ($dayCol + $d - 1))5:$(ColName ($dayCol + $d - 1))$end").Interior.Color = 0xF2F2F2 }
  foreach ($sr in $subRows + $end) { $rg = $ws.Range("A${sr}:${cN}$sr"); $rg.Interior.Color = $(if ($sr -eq $end) { 0x00C0FF } else { 0x00FFFF }); $rg.Font.Bold = $true }
  $ws.Range("${cN}5:${cN}$end").Font.Size = 9
  $widths = @{ 1 = 5; 2 = 22; 3 = 18; 4 = 6; 5 = 10; 14 = 9 }; foreach ($c in 6..13) { $widths[$c] = 9 }
  foreach ($c in $widths.Keys) { $ws.Columns.Item($c).ColumnWidth = $widths[$c] }
  foreach ($c in $dayCol..($dayCol + $daysInMonth - 1)) { $ws.Columns.Item($c).ColumnWidth = 5.4 }
  $ws.Columns.Item($sumCol).ColumnWidth = 9; $ws.Columns.Item($carryCol).ColumnWidth = 10; $ws.Columns.Item($noteCol).ColumnWidth = 60
  $ws.Rows.Item(3).RowHeight = 42; $ws.Rows.Item(4).RowHeight = 24
  $ws.Activate(); $ws.Range('D5').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true
  $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = $false; $ws.PageSetup.PrintTitleRows = '$3:$4'

  # --- 시트 2: 출하품번별 산출 근거
  $w2 = $wb.Worksheets.Item(2); $w2.Name = '산출근거(출하품번별)'; $w2.Cells.Font.Name = '맑은 고딕'; $w2.Cells.Font.Size = 10
  $h2 = '계획 NO', '고객사', '출하품번', '품명', "$mLabel 출하", "$nLabel 출하", "$nLabel 선행분", '소재품번', '금형 재고', '가공 재고(전체)', '가공 재고(인정)', '그룹 출하+선행', '그룹 가용재고', '주조 필요량', '산출 근거', '매칭 방법', '검토 메모'
  for ($c = 0; $c -lt $h2.Count; $c++) { Put $w2 1 ($c + 1) $h2[$c] }
  $last2 = ColName $h2.Count
  $w2.Range("A1:${last2}1").Interior.Color = 0xFFFFCC; $w2.Range("A1:${last2}1").Font.Bold = $true; $w2.Range("A1:${last2}1").WrapText = $true; $w2.Range("A1:${last2}1").HorizontalAlignment = -4108; $w2.Rows.Item(1).RowHeight = 32
  $r = 2
  foreach ($p in ($plans | Sort-Object No)) {
    foreach ($o in ($p.Orders | Sort-Object PN)) {
      $vals = @([double]$p.No, $o.Cust, $o.PN, $o.Name, $o.Cur, $o.Next, [math]::Ceiling($o.Next * $nextShare), (($o.MoldKeys | ForEach-Object { $moldBy[$_][0].PN }) -join ' / '), $o.Mold,
        (($o.MachRows | Measure-Object Qty -Sum).Sum), (($o.MachRows | Where-Object { MachOk $_ } | Measure-Object Qty -Sum).Sum), ($p.Cur + $p.Pre), $p.Stock, $p.Need, $p.Rule, ($o.Method -join ' / '), $o.Memo)
      for ($c = 0; $c -lt $vals.Count; $c++) { Put $w2 $r ($c + 1) $vals[$c] }
      $r++
    }
  }
  $w2.Range("A1:${last2}$($r - 1)").Borders.LineStyle = 1; $w2.Range("A1:${last2}$($r - 1)").Borders.Color = 0x808080
  $w2.Range("E2:G$($r - 1)").NumberFormat = '#,##0'; $w2.Range("I2:N$($r - 1)").NumberFormat = '#,##0'
  $wd2 = 8, 18, 15, 28, 10, 10, 10, 20, 10, 10, 10, 11, 11, 10, 40, 30, 30; for ($c = 0; $c -lt $wd2.Count; $c++) { $w2.Columns.Item($c + 1).ColumnWidth = $wd2[$c] }
  $w2.Activate(); $w2.Range('A2').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true

  # --- 시트 3: 확인 필요
  $w3 = $wb.Worksheets.Item(3); $w3.Name = '확인 필요'; $w3.Cells.Font.Name = '맑은 고딕'; $w3.Cells.Font.Size = 10
  $issues = @()
  $issues += $plans | Where-Object Short -gt 0 | Sort-Object Short -Descending | ForEach-Object { [pscustomobject]@{ K = '생산 미달'; C = $_.Cust; P = $_.CastPN; Q = $_.Short; T = $_.Rule } }
  $issues += $plans | Where-Object { $_.NoCast -and $_.Need -gt 0 } | ForEach-Object { [pscustomobject]@{ K = '소재 미확인'; C = $_.Cust; P = $_.CastPN; Q = $_.Need; T = '소재 매핑표에 확정 소재품번을 입력하세요' } }
  $issues += $plans | Where-Object { $_.Rule -match '이력 없음' } | ForEach-Object { [pscustomobject]@{ K = '전월 일생산 없음'; C = $_.Cust; P = $_.CastPN; Q = $_.Need; T = $_.Rule } }
  $issues += $orders | Where-Object { $_.Memo } | ForEach-Object { [pscustomobject]@{ K = '매핑 검토'; C = $_.Cust; P = $_.PN; Q = $null; T = $_.Memo } }
  $issues += $orders | Where-Object { -not $_.MoldKeys -and $_.Cand -and ($_.Cur -gt 0 -or $_.Next -gt 0) } | ForEach-Object { [pscustomobject]@{ K = '소재 후보 확정 필요'; C = $_.Cust; P = $_.PN; Q = $_.Cur; T = "후보: $($_.Cand)" } }
  $issues = @($issues | Where-Object { $_ })
  $h3 = '구분', '고객사', '품번', '수량', '내용'
  for ($c = 0; $c -lt $h3.Count; $c++) { Put $w3 1 ($c + 1) $h3[$c] }
  $w3.Range('A1:E1').Interior.Color = 0xFFFFCC; $w3.Range('A1:E1').Font.Bold = $true
  $r = 2; foreach ($x in $issues) { Put $w3 $r 1 $x.K; Put $w3 $r 2 $x.C; Put $w3 $r 3 $x.P; if ($null -ne $x.Q) { Put $w3 $r 4 ([double]$x.Q) }; Put $w3 $r 5 $x.T; $r++ }
  if ($r -gt 2) { $w3.Range("A1:E$($r - 1)").Borders.LineStyle = 1; $w3.Range("A1:E$($r - 1)").Borders.Color = 0x808080; $w3.Range("D2:D$($r - 1)").NumberFormat = '#,##0' } else { Put $w3 2 1 '확인할 항목이 없습니다.' }
  $wd3 = 18, 20, 24, 10, 80; for ($c = 0; $c -lt $wd3.Count; $c++) { $w3.Columns.Item($c + 1).ColumnWidth = $wd3[$c] }

  $ws.Activate(); $xl.ScreenUpdating = $true
  $wb.SaveAs($xlsxPath, 51); $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }

# ============================================================ 요약
Write-Host ''
Write-Host "=============== 결과 ===============" -ForegroundColor Green
$plans | Group-Object Cust | Sort-Object { ($_.Group | Measure-Object Cur -Sum).Sum } -Descending | ForEach-Object {
  "  {0,-24} {1} 출하 {2,8:N0}  선행 {3,7:N0}  재고 {4,8:N0}  주조 {5,8:N0}" -f $_.Name, $mLabel, ($_.Group | Measure-Object Cur -Sum).Sum, ($_.Group | Measure-Object Pre -Sum).Sum, ($_.Group | Measure-Object Stock -Sum).Sum, (($_.Group | ForEach-Object { ($_.Plan.Values | Measure-Object -Sum).Sum }) | Measure-Object -Sum).Sum
} | Write-Host
Write-Host ("  주조 계획 합계 {0:N0}개" -f $prodSum)
$short = @($plans | Where-Object Short -gt 0)
if ($short) { Write-Host "  생산 미달 $($short.Count)건: $(($short | Sort-Object Short -Descending | ForEach-Object { "$($_.CastPN)($($_.Short))" }) -join ', ')" -ForegroundColor Yellow }
Write-Host ''
Write-Host "  계획 파일 : $xlsxPath"
Write-Host "  매칭 결과 : $matchCsv"
if ($cfg.'완료후열기') { Invoke-Item -LiteralPath $xlsxPath }
}
finally {
  Remove-Item -Recurse -Force $script:TempRoot -ErrorAction SilentlyContinue
}
