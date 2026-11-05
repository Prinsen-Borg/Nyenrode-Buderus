# Builds v2-business-case-model.xlsx with live formulas, so the group can change an
# assumption on the Inputs sheet and see benefits, costs and payback update.
param([string]$Out = "$PSScriptRoot\v2-business-case-model.xlsx")
$ErrorActionPreference = 'Stop'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(2).Delete() }
$in = $wb.Worksheets.Item(1); $in.Name = 'Inputs'
$md = $wb.Worksheets.Add([Type]::Missing, $in); $md.Name = 'Model'

function Put($ws, $addr, $val) { $ws.Range($addr).Formula = [string]$val }
function F($ws, $addr, $f) { $ws.Range($addr).Formula = $f }
function A([int]$row, [int]$col) { return ([string][char](64+$col)) + $row }

# ---- Inputs: label | low | base | high | label (case/source/assumption) ----
Put $in 'A1' 'Buderus – business case OEE & predictive maintenance (v2)'
Put $in 'A2' 'Change the yellow cells. Labels: [case] case text · [source] public benchmark · [group] group concept · [assumption] to validate.'
$hdr = @('Input','Low','Base','High','Unit','Label')
for ($c=0; $c -lt 6; $c++) { $in.Range((A 4 ($c+1))).Formula = $hdr[$c] }
$rows = @(
  @('Revenue',                         2000,  2000,  2000, '€ m / year',   '[group]'),
  @('Production lines',                 120,   120,   120, 'lines',        '[group]'),
  @('Legacy lines',                      18,    18,    18, 'lines',        '[group]'),
  @('Plants',                            10,    10,    10, 'plants',       '[group]'),
  @('Planned hours per line',          4000,  4000,  4000, 'h / year',     '[assumption] two shifts'),
  @('Contribution margin',              0.30,  0.30,  0.30, '%',           '[assumption]'),
  @('Idle crew cost per downtime hour',  500,   500,   500, '€ / h',       '[assumption]'),
  @('Sellable share of extra output',   1.00,  1.00,  1.00, '%',           '[assumption] main sensitivity'),
  @('OEE today',                        0.70,  0.70,  0.70, '%',           '[assumption]'),
  @('Unplanned downtime',               0.06,  0.06,  0.06, '% of planned','[assumption]'),
  @('Share caused by equipment failure',0.50,  0.50,  0.50, '%',           '[assumption]'),
  @('Reduction of failure downtime',    0.20,  0.30,  0.50, '%',           '[source] McKinsey 30–50% (mature)'),
  @('Maintenance spend',                  30,    30,    30, '€ m / year',  '[assumption] 1.5% of revenue'),
  @('Maintenance cost reduction',       0.05,  0.08,  0.10, '%',           '[source] Deloitte 5–10%'),
  @('Scrap and rework cost',              30,    30,    30, '€ m / year',  '[assumption] 1.5% of revenue'),
  @('Quality loss reduction',           0.05,  0.07,  0.10, '%',           '[assumption]'),
  @('Central platform (one-off)',          5,     5,     5, '€ m',         '[assumption]'),
  @('OT-DMZ and edge per plant',         0.5,   0.5,   0.5, '€ m / plant', '[assumption]'),
  @('Retrofit per standard line',       0.05,  0.05,  0.05, '€ m / line',  '[assumption]'),
  @('Retrofit per legacy line',          0.2,   0.2,   0.2, '€ m / line',  '[assumption]'),
  @('Programme and integration',           3,     3,     3, '€ m',         '[assumption]'),
  @('Training and change',                 2,     2,     2, '€ m',         '[assumption]'),
  @('Recurring platform cost',             2,     2,     2, '€ m / year',  '[assumption]'),
  @('Recurring support and team',        1.5,   1.5,   1.5, '€ m / year',  '[assumption]'),
  @('Discount rate',                    0.08,  0.08,  0.08, '%',           '[assumption] Finance to confirm'))
# Write the inputs as one 2D block; per-cell COM writes of numbers fail in Windows PowerShell 5.1
$arr = New-Object 'object[,]' $rows.Count, 6
for ($i=0; $i -lt $rows.Count; $i++) { for ($c=0; $c -lt 6; $c++) { $arr[$i,$c] = $rows[$i][$c] } }
# Value2 setters are bound as string here, so numbers go in through .Formula (invariant culture)
for ($i=0; $i -lt $rows.Count; $i++) { for ($c=0; $c -lt 6; $c++) { $v = $arr[$i,$c]; $s = if ($v -is [string]) { $v } else { '=' + ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }; $in.Range((A (5+$i) ($c+1))).Formula = $s } }
$last = 4 + $rows.Count
$in.Range("B5:D$last").Interior.Color = 13434879   # light yellow = editable
foreach ($r in @(10,12,13,14,15,16,18,20,29)) { $in.Range("B${r}:D${r}").NumberFormat = '0%' }

# Named cells per scenario column: R(row) returns e.g. Inputs!C12 for the chosen column
# ---- Model: one column per scenario (B low, C base, D high) ----
Put $md 'A1' 'Model (formulas – do not type over)'
$mh = @('Line','Low','Base','High','Note'); for ($c=0; $c -lt 5; $c++) { $md.Range((A 3 ($c+1))).Formula = $mh[$c] }
$lines = @(
  @('Revenue per line (€ m)',                 '=Inputs!{c}5/Inputs!{c}6'),
  @('Value per downtime hour (€)',            '={c}4*1000000/Inputs!{c}9*Inputs!{c}10*Inputs!{c}12+Inputs!{c}11'),
  @('Failure downtime per line (h)',          '=Inputs!{c}9*Inputs!{c}14*Inputs!{c}15'),
  @('Hours regained per line (h)',            '={c}6*Inputs!{c}16'),
  @('Benefit: downtime (€ m / year)',         '={c}7*{c}5*Inputs!{c}6/1000000'),
  @('Benefit: maintenance (€ m / year)',      '=Inputs!{c}17*Inputs!{c}18'),
  @('Benefit: quality (€ m / year)',          '=Inputs!{c}19*Inputs!{c}20'),
  @('Gross benefits (€ m / year)',            '=SUM({c}8:{c}10)'),
  @('OEE effect (points)',                    '={c}7/Inputs!{c}9*Inputs!{c}13*100'),
  @('One-off investment (€ m)',               '=Inputs!{c}21+Inputs!{c}22*Inputs!{c}8+Inputs!{c}23*(Inputs!{c}6-Inputs!{c}7)+Inputs!{c}24*Inputs!{c}7+Inputs!{c}25+Inputs!{c}26'),
  @('Recurring cost (€ m / year)',            '=Inputs!{c}27+Inputs!{c}28'),
  @('Net benefits at full roll-out (€ m / year)', '={c}11-{c}14'))
for ($i=0; $i -lt $lines.Count; $i++) {
  $r = 4+$i; $md.Range("A$r").Formula = $lines[$i][0]
  foreach ($c in @('B','C','D')) { F $md "$c$r" ($lines[$i][1].Replace('{c}', $c)) } }

# Cash flow over 5 years: investment phased 1/3, 5/12, 1/4; benefits ramp 5%, 35%, 75%, 100%, 100%
Put $md 'A18' 'Cash flow (€ m)'; Put $md 'B18' 'Year 1'; Put $md 'C18' 'Year 2'; Put $md 'D18' 'Year 3'; Put $md 'E18' 'Year 4'; Put $md 'F18' 'Year 5'
Put $md 'A19' 'Investment phasing'; $ph=@(0.333,0.417,0.25,0,0); for($y=0;$y -lt 5;$y++){ $md.Range((A 19 (2+$y))).Formula="=$($ph[$y])" }
Put $md 'A20' 'Benefit ramp-up';    $rp=@(0.05,0.35,0.75,1,1);    for($y=0;$y -lt 5;$y++){ $md.Range((A 20 (2+$y))).Formula="=$($rp[$y])" }
$md.Range('B19:F20').Interior.Color = 13434879; $md.Range('B19:F20').NumberFormat = '0%'
$scen = @(@('Low','B',21),@('Base','C',24),@('High','D',27))
foreach ($s in $scen) {
  $r=$s[2]; $col=$s[1]
  Put $md "A$r" "$($s[0]): net cash flow"; Put $md "A$($r+1)" "$($s[0]): cumulative"
  for ($y=0; $y -lt 5; $y++) {
    $cl=[char](66+$y)
    F $md "$cl$r" "=-`$${col}`$13*${cl}`$19+`$${col}`$15*${cl}`$20"
    if ($y -eq 0) { F $md "$cl$($r+1)" "=$cl$r" } else { $pv=[char](65+$y); F $md "$cl$($r+1)" "=$pv$($r+1)+$cl$r" } }
  Put $md "G$r" 'NPV'; F $md "H$r" "=NPV(Inputs!${col}29,B${r}:F${r})" }

foreach ($ws in @($in,$md)) {
  $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 14
  $ws.Cells.Font.Name = 'Calibri'; $ws.Columns.Item(1).ColumnWidth = 44
  for ($c=2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 12 } }
$in.Columns.Item(6).ColumnWidth = 40
$in.Range('A4:F4').Font.Bold = $true; $md.Range('A3:E3').Font.Bold = $true; $md.Range('A18:F18').Font.Bold = $true
$md.Range('B4:D16').NumberFormatLocal = '#.##0,0'; $md.Range('B5:D5').NumberFormatLocal = '#.##0'
$md.Range('B21:F28').NumberFormatLocal = '#.##0,0'; foreach ($h in @('H21','H24','H27')) { $md.Range($h).NumberFormatLocal = '#.##0,0' }
$md.Range('A11:D11').Font.Bold = $true; $md.Range('A15:D15').Font.Bold = $true

# NumberFormatLocal: this machine runs Excel with Dutch separators; Excel stores the format locale-neutrally
$wb.SaveAs($Out, 51)
$xl.Calculate()
$res = foreach ($r in 4..16) { '{0,-46} {1,8:N1} {2,8:N1} {3,8:N1}' -f $md.Cells.Item($r,1).Text, $md.Cells.Item($r,2).Value2, $md.Cells.Item($r,3).Value2, $md.Cells.Item($r,4).Value2 }
$res += foreach ($r in 21..28) { if ($md.Cells.Item($r,1).Text) { '{0,-24} {1} | NPV {2:N1} shown as {3}' -f $md.Cells.Item($r,1).Text, (($md.Range("B${r}:F${r}").Value2 | ForEach-Object { '{0:N1}' -f $_ }) -join ' '), $md.Cells.Item($r,8).Value2, $md.Cells.Item($r,8).Text } }
$wb.Close($false); $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
$res
