# Builds v2-board-presentation.pptx (9 slides, 16:9) in the DESIGN.md palette.
# Runs PowerPoint through COM, so it needs Windows with Office installed.
param([string]$Out = "$PSScriptRoot\v2-board-presentation.pptx", [string]$PngDir = '')
$ErrorActionPreference = 'Stop'

# DESIGN.md tokens; COM wants BGR integers
function C($hex) { [Convert]::ToInt32($hex.Substring(0,2),16) + 256*[Convert]::ToInt32($hex.Substring(2,2),16) + 65536*[Convert]::ToInt32($hex.Substring(4,2),16) }
$INK=C '18242F'; $BODY=C '3D4B57'; $MUTED=C '6B7A86'; $LINE=C 'DDE3E8'; $CANVAS=C 'F4F6F8'; $WHITE=C 'FFFFFF'
$NAVY=C '1F3A5F'; $NAVY2=C '2D5487'; $TINT=C 'E9EFF6'; $FLAG=C 'B8741A'; $FLAGBG=C 'FBF1E3'; $OK=C '2F7A4F'; $OKBG=C 'E6F3EB'
$FONT='Segoe UI'; $FONTB='Segoe UI Semibold'

function TB($s, $l, $t, $w, $h, $text, $size, $color, $font=$FONT, $align=1) {
  $b = $s.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
  $tf = $b.TextFrame; $tf.WordWrap = -1; $tf.AutoSize = 0
  $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
  $r = $tf.TextRange; $r.Text = $text; $r.Font.Name = $font; $r.Font.Size = [single]$size; $r.Font.Color.RGB = $color
  $r.ParagraphFormat.Alignment = [int]$align; return $b }

function BOX($s, $l, $t, $w, $h, $fill, $line=$null, $lw=1, $round=$true) {
  $type = if ($round) { 5 } else { 1 }
  $b = $s.Shapes.AddShape($type, [single]$l, [single]$t, [single]$w, [single]$h)
  if ($round) { $b.Adjustments.Item(1) = [single]0.08 }
  if ($null -eq $fill) { $b.Fill.Visible = 0 } else { $b.Fill.Visible = -1; $b.Fill.Solid(); $b.Fill.ForeColor.RGB = $fill }
  if ($null -eq $line) { $b.Line.Visible = 0 } else { $b.Line.Visible = -1; $b.Line.ForeColor.RGB = $line; $b.Line.Weight = [single]$lw }
  $b.Shadow.Visible = 0; return $b }

# Text inside a shape: first paragraph bold, rest body
function FILL($b, $text, $size, $color, $boldFirst=$true, $align=1, $anchor=1, $margin=10) {
  $tf = $b.TextFrame; $tf.WordWrap = -1
  $tf.MarginLeft = [single]$margin; $tf.MarginRight = [single]$margin; $tf.MarginTop = [single]8; $tf.MarginBottom = [single]8; $tf.VerticalAnchor = [int]$anchor
  $r = $tf.TextRange; $r.Text = $text; $r.Font.Name = $FONT; $r.Font.Size = [single]$size; $r.Font.Color.RGB = $color
  $r.ParagraphFormat.Alignment = [int]$align; $r.ParagraphFormat.SpaceAfter = [single]3
  if ($boldFirst) { $r.Paragraphs(1).Font.Name = $FONTB } }

function Title($s, $text, $n) {
  TB $s 48 36 864 60 $text 24 $INK $FONTB | Out-Null
  TB $s 48 512 600 14 'Buderus · Point of View for the Board · Group 3' 9 $MUTED | Out-Null
  TB $s 852 512 60 14 "$n" 9 $MUTED $FONT 3 | Out-Null }

function Section($s, $text) { TB $s 48 18 600 14 $text.ToUpper() 10 $NAVY2 $FONTB | Out-Null }

function TABLE($s, $l, $t, $w, $rows, $widths, $size=11) {
  $nr = $rows.Count; $nc = $rows[0].Count
  $sh = $s.Shapes.AddTable($nr, $nc, [single]$l, [single]$t, [single]$w, [single](22*$nr))
  $tbl = $sh.Table; $tbl.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}', $false)
  for ($c=1; $c -le $nc; $c++) { $tbl.Columns.Item($c).Width = [single]$widths[$c-1] }
  for ($r=1; $r -le $nr; $r++) { for ($c=1; $c -le $nc; $c++) {
    $cell = $tbl.Cell($r,$c); $cs = $cell.Shape; $tf = $cs.TextFrame
    $tf.MarginLeft = [single]6; $tf.MarginRight = [single]6; $tf.MarginTop = [single]4; $tf.MarginBottom = [single]4
    $tr = $tf.TextRange; $tr.Text = [string]$rows[$r-1][$c-1]; $tr.Font.Name = $FONT; $tr.Font.Size = [single]$size; $tr.Font.Color.RGB = $BODY
    if ($r -eq 1) { $cs.Fill.Visible = -1; $cs.Fill.Solid(); $cs.Fill.ForeColor.RGB = $NAVY; $tr.Font.Color.RGB = $WHITE; $tr.Font.Name = $FONTB }
    else { $cs.Fill.Visible = 0; if ($c -eq 1) { $tr.Font.Name = $FONTB; $tr.Font.Color.RGB = $INK } }
    $bd = $cell.Borders.Item(3); $bd.Visible = -1; $bd.ForeColor.RGB = $LINE; $bd.Weight = [single]0.75 } }
  for ($r=1; $r -le $nr; $r++) { $tbl.Rows.Item($r).Height = [single]10 }
  return $tbl }

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add(0)
$pres.PageSetup.SlideWidth = [single]960; $pres.PageSetup.SlideHeight = [single]540
$blank = $pres.SlideMaster.CustomLayouts.Item(7)
function NewSlide($i) { $s = $pres.Slides.AddSlide($i, $blank); $s.FollowMasterBackground = 0; $s.Background.Fill.Solid(); $s.Background.Fill.ForeColor.RGB = $WHITE; return $s }

# ---------- 1 Title ----------
$s = NewSlide 1; $s.Background.Fill.ForeColor.RGB = $NAVY
TB $s 64 150 832 20 'GROUP 3 · CASE BUDERUS · ADVANCED IT PROGRAM, NYENRODE' 12 (C 'BCD0E8') $FONTB | Out-Null
TB $s 64 182 832 130 'Scaling AI and IoT in our plants: group standards, local control' 40 $WHITE $FONTB | Out-Null
TB $s 64 330 760 60 'Point of View for the Board. Decisions on architecture, governance, investment, risk and people.' 18 (C 'DBE6F2') | Out-Null
TB $s 64 440 600 20 'Louw Fokker, CIO  ·  Board meeting 16 December 2026' 13 (C 'C9D7E7') | Out-Null
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Open with the question of the CEO and the core tension: how do we standardise digitally at group level without ignoring plant-specific OT constraints? The answer comes first, on the next slide.'

# ---------- 2 Decisions requested ----------
$s = NewSlide 2; Section $s '01 · Decisions requested'; Title $s 'We ask the board for five decisions today' 2
$dec = @(
  @('Direction','Adopt a hybrid, federated OT/IoT/data/AI architecture with group-wide minimum standards for segmentation, edge, data platform and security.'),
  @('Ownership','Install a Digital Operations Steering Group with the mandates and escalation lines proposed, including a local go/no-go for production impact.'),
  @('Investment','Release the approved budget in phases through investment gates with financial, technical and safety criteria. No automatic group-wide roll-out.'),
  @('Risk limits','Confirm that safety and continuity lead: AI advises, people decide. Writing back to machine control needs separate approval.'),
  @('People','Fund OT/IT integration roles, a digital champion per plant and a learning path for operators, maintenance and engineering.'))
for ($i=0; $i -lt 5; $i++) {
  $y = 116 + $i*76
  BOX $s 48 $y 864 66 $CANVAS | Out-Null
  $o = $s.Shapes.AddShape(9, [single]64, [single]($y+17), [single]32, [single]32); $o.Fill.ForeColor.RGB = $NAVY; $o.Line.Visible = 0; $o.Shadow.Visible = 0
  FILL $o "$($i+1)" 14 $WHITE $true 2 3 0
  TB $s 112 ($y+12) 150 20 $dec[$i][0] 15 $INK $FONTB | Out-Null
  TB $s 262 ($y+12) 560 46 $dec[$i][1] 13 $BODY | Out-Null
  $tag = BOX $s 832 ($y+20) 66 24 $TINT; FILL $tag 'DECIDE' 10 $NAVY $true 2 3 0 }
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Answer first. Each decision maps to a later section: 1 architecture, 2 governance, 3 business case and roadmap, 4 risk, 5 organisation and talent. The supervisory board already approved the budget; the question is how and under which conditions we spend it.'

# ---------- 3 Management summary ----------
$s = NewSlide 3; Section $s '02 · Management summary'; Title $s 'The pilot proves value; scaling needs a shared foundation and mandates' 3
$scr = @(
  @('Situation','The energy transition drives demand and the pressure on efficiency, quality and cost. One plant ran a successful pilot: sensor data enables predictive maintenance, higher availability and better quality.'),
  @('Complication','OT is business-critical and plant-specific: 120 lines, two PLC families, 18 legacy lines on Windows XP and Profibus. No standard OT-DMZ, OEM remote tunnels, an isolated cloud proof of concept, and a skills gap in data, AI and OT/IT.'),
  @('Resolution','Standardise the digital foundation at group level. Plants apply it within local OT, safety and continuity limits. Scale per plant and per use case, only when value, technical fit and continuity are proven.'))
for ($i=0; $i -lt 3; $i++) {
  $x = 48 + $i*296
  $b = BOX $s $x 116 272 288 $(if ($i -eq 2) { $TINT } else { $CANVAS })
  FILL $b ($scr[$i][0] + "`r" + $scr[$i][1]) 13 $BODY $true 1 1 16
  $b.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]16; $b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $INK }
$k = BOX $s 48 420 864 70 $NAVY
FILL $k "Key message`rScale on proven value and repeatable standards, with local control where production requires it." 15 $WHITE $true 1 3 16
$k.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]11; $k.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = (C 'BCD0E8')
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Situation, complication, resolution. The real bottleneck is not the AI model but the OT/IT boundary: segmentation, integration and ownership.'

# ---------- 4 Business case ----------
$s = NewSlide 4; Section $s '03 · Business case & opportunities'; Title $s 'A modest but credible case: €12 m a year for €24 m, payback in year 5' 4
$kpi = @(@('€12 m','gross benefits per year (base)'),@('€24 m','one-off investment'),@('≈ 4.5 yrs','payback (base)'),@('+0.6 pt','OEE improvement'))
for ($i=0; $i -lt 4; $i++) { $x = 48 + $i*216; $b = BOX $s $x 112 200 76 $CANVAS
  FILL $b ($kpi[$i][0] + "`r" + $kpi[$i][1]) 11 $MUTED $true 1 3 14
  $p1 = $b.TextFrame.TextRange.Paragraphs(1); $p1.Font.Size = [single]26; $p1.Font.Color.RGB = $NAVY }
$rows = @(
  @('€ m per year','Low','Base','High'),
  @('Less unplanned downtime','5.0','7.6','12.6'),
  @('Lower maintenance cost','1.5','2.4','3.0'),
  @('Less quality loss','1.5','2.1','3.0'),
  @('Gross benefits','8.0','12.1','18.6'),
  @('Recurring cost','−3.5','−3.5','−3.5'),
  @('Net benefits at full roll-out','4.5','8.6','15.1'))
$tbl = TABLE $s 48 206 520 $rows @(268,84,84,84) 12
foreach ($r in @(5,7)) { for ($c=1; $c -le 4; $c++) { $tbl.Cell($r,$c).Shape.TextFrame.TextRange.Font.Name = $FONTB } }
for ($r=1; $r -le 7; $r++) { for ($c=2; $c -le 4; $c++) { $tbl.Cell($r,$c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 3 } }
$b = BOX $s 592 206 320 230 $FLAGBG (C 'ECD2AD') 0.75
FILL $b "Why this differs from our first version`rv1 counted revenue: 15–25% extra output, €300–500 m. v2 counts value: avoided downtime × margin, plus maintenance and quality savings, minus run costs.`rDemand is the main sensitivity. If extra output cannot be sold, the downtime benefit drops to about €2–3 m.`rThe platform is a fixed cost: the case improves with every line connected, which is why it is a group decision." 12 $BODY $true 1 1 14
$b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = (C '7A4A0C')
TB $s 48 452 864 40 'All figures are assumptions to validate with pilot data. Benchmarks: McKinsey (30–50% less downtime in mature programmes), Deloitte (5–10% lower maintenance cost). Model: deliverables/business-case/v2-business-case-model.xlsx' 9.5 $MUTED | Out-Null
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'For the CFO: positive but modest is credible. Scale model from our v1 (10 plants, 120 lines, €2 bn revenue). Value per downtime hour €1,750 (margin plus idle crew). Payback about 4.5 years base, 3.5 high, not within 5 years in the low case. NPV at 8% is about zero after 5 years and +€10 m after 7. The case gives no pilot figures: we show which data we need (OEE baseline, downtime, MTTR, false positives, costs).'

# ---------- 5 Governance ----------
$s = NewSlide 5; Section $s '04 · Governance'; Title $s 'A Steering Group decides; plants keep a veto on safety' 5
$rows = @(
  @('Role','Mandate'),
  @('Board','Sets goals, risk appetite and investment frame; decides on group-wide scaling'),
  @('Digital Operations Steering Group (chair: COO)','Prioritises use cases, tracks value, risk and capacity; prepares gate decisions'),
  @('Operations / business owner','Owns the use case, the process change and the outcome'),
  @('CFO / Finance','Validates business case, benefit measurement and budget deviations'),
  @('CIO / Group IT','Owns architecture, platform, integration standards and IT roadmap'),
  @('OT / Engineering','Machine integration, OT safety, test and change procedures'),
  @('Plant manager','Local go/no-go for production impact, maintenance windows and acceptance'),
  @('CISO','Sets security and privacy requirements; can block a roll-out'))
TABLE $s 48 110 864 $rows @(300,564) 12 | Out-Null
$esc = @(@('Plant risk','Plant manager pauses'),@('Exception to a standard','Architecture / Security Board'),@('Budget deviation','Steering Group, then Board'),@('Multi-plant incident','CIO, COO and CISO directly'))
for ($i=0; $i -lt 4; $i++) { $x = 48 + $i*216; $b = BOX $s $x 418 204 70 $TINT
  FILL $b ($esc[$i][0] + "`r→ " + $esc[$i][1]) 12 $NAVY $true 1 3 12 }
TB $s 48 400 300 14 'Escalation' 10 $NAVY2 $FONTB | Out-Null
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Governance between Operations, Business, Finance, IT and OT. The COO chairs because the value is created in the plants; the CIO is vice-chair. COBIT 2019 is the accountability frame behind these roles.'

# ---------- 6 Future state architecture ----------
$s = NewSlide 6; Section $s '05 · Future state architecture'; Title $s 'Hybrid: control stays local, intelligence central, a hard boundary between' 6
$layers = @(
  @('L4 · Group platform (cloud)','IoT, data and AI platform: storage, data quality, model training, model registry, MLOps. ERP, PLM and data warehouse.', $TINT, 'Group IT'),
  @('L3.5 · OT-DMZ','Brokers and gateways; only necessary data flows; managed remote access instead of OEM tunnels.', $FLAGBG, 'Group standard'),
  @('L3 · Plant: MES and edge','Local MES instance from a central template. Edge gateways collect and buffer data and run validated models.', $CANVAS, 'Plant'),
  @('L0–L2 · Machine and control','PLC, SCADA and safety functions stay local and autonomous. AI never takes over safety-critical control.', $CANVAS, 'Plant / OT'))
for ($i=0; $i -lt 4; $i++) { $y = 112 + $i*84
  $b = BOX $s 48 $y 560 74 $layers[$i][2]
  FILL $b ($layers[$i][0] + "`r" + $layers[$i][1]) 12 $BODY $true 1 3 14
  $b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $INK; $b.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]13
  TB $s 616 ($y+28) 90 18 $layers[$i][3] 10 $NAVY2 $FONTB | Out-Null }
$pr = BOX $s 712 112 200 326 $NAVY
FILL $pr "Design principles`r1  Plants keep running when the link to the platform fails.`r2  Reusable group standards, minimal coupling to control networks.`r3  Traceable data and models, and an explicit exit from supplier lock-in.`r4  Advice flows back through Maximo and MES; write-back to control only after separate approval." 12 $WHITE $true 1 1 14
$pr.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]14
TB $s 48 452 864 30 'Platform products are chosen after a requirements and supplier review. Today: no standard OT-DMZ, OEM TeamViewer/VPN tunnels, manual CSV transfers and an isolated cloud proof of concept.' 10 $MUTED | Out-Null
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Purdue / ISA-95 layering. The group concept placed a central MES in Amsterdam; we recommend a central MES template with local instances, because MES runs production execution and must keep working when the network fails.'

# ---------- 7 IT organisation design ----------
$s = NewSlide 7; Section $s '05 · IT organisation design'; Title $s 'Group standards where work is common, local control where it is specific' 7
TB $s 48 108 420 16 'Business Activity Model: where each activity belongs' 11 $NAVY2 $FONTB | Out-Null
$q = @(
  @(48,128,'Specialized · local','Safe control of production lines · OT security per plant', $CANVAS),
  @(260,128,'Distinct · product team','Predicting failures with AI (pilot) · new AI use cases', $TINT),
  @(48,236,'Common · group standard','Finance and ERP processes · maintenance planning (one group process)', $CANVAS),
  @(260,236,'Adaptive · group platform','Collecting and exposing machine data · OEE reporting', $CANVAS))
foreach ($c in $q) { $b = BOX $s $c[0] $c[1] 204 100 $c[4]; FILL $b ($c[2] + "`r" + $c[3]) 12 $BODY $true 1 1 12; $b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $INK }
TB $s 48 344 416 30 'Specific ↑  /  Generic ↓    ·    Stable ←  /  Responsive →' 9.5 $MUTED | Out-Null
$b = BOX $s 488 108 424 132 $CANVAS
FILL $b "SAFe: one Agile Release Train 'Digital Operations'`rProduct teams per use case (predictive maintenance, quality, energy) with operations, OT, data and IT in the team.`rPlatform teams for edge and connectivity, data and AI platform, OT security and integration.`rQuarterly PI planning with the plants at the table." 12 $BODY $true 1 1 14
$b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $INK
$b = BOX $s 488 252 424 150 $CANVAS
FILL $b "Talent: measure first, then build`rSkills scan per role and plant, then learning paths for operators and maintenance, OT/IT integration, data and analytics, AI and model management, and leaders.`rA digital champion per plant; rotation between plants and Group IT.`rSuccess = teams can assess, run and safely change use cases themselves." 12 $BODY $true 1 1 14
$b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $INK
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Business Activity Model (multimodality): Common and Adaptive go to Group IT as standards; Specialized stays local under plant management and OT; Distinct becomes a product team with a business case until the use case is proven and moves to Adaptive. Roles: Group IT, plant management, engineering/OT, agile teams and security.'

# ---------- 8 Risk and control ----------
$s = NewSlide 8; Section $s '05 · Risk and control'; Title $s 'The biggest risks sit on the OT/IT boundary, each tied to a COBIT objective' 8
$rows = @(
  @('Risk','Measure','COBIT 2019'),
  @('Network segmentation','Zones and conduits (IEC 62443); only necessary data flows; monitoring','APO13, DSS05, BAI10'),
  @('Remote access by suppliers','One managed access point; personal, time-bound accounts; logging','APO10, APO13, DSS05'),
  @('Patching and change (incl. Windows XP lines)','OT change procedure, test and fallback plan; isolate unpatchable systems','BAI06, BAI10, DSS01'),
  @('Supplier dependency','Assess support, data export, lifecycle and exit before selection','APO10, APO12, APO05'),
  @('Data and model integrity','Provenance, time stamps, quality checks, drift monitoring','APO14, BAI10, MEA02'),
  @('Production continuity','Local autonomy, buffering, continuity tests, stop criteria','EDM03, DSS01, DSS04'),
  @('Wrong AI advice','Human review, clear limits, manual fallback','APO12, BAI03, DSS06'))
TABLE $s 48 108 864 $rows @(250,454,160) 11.5 | Out-Null
$b = BOX $s 48 418 864 72 $TINT
FILL $b "Gartner Top Strategic Technology Trends 2026: our top-3`rPhysical AI (condition monitoring now, autonomous action only after safety validation)  ·  AI Security Platforms (part of the security architecture, not a replacement for OT segmentation)  ·  Domain-Specific Language Models (maintenance knowledge for technicians, never for control)" 11.5 $NAVY $true 1 3 14
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'COBIT 2019 is the governance and accountability frame; IEC 62443 and ISO 27001 cover technical OT and information security; ISO/IEC 42001 covers AI management. Privacy: data minimisation and retention for personal data in telemetry and logs (APO14, APO13, MEA03).'

# ---------- 9 Program structure and roadmap ----------
$s = NewSlide 9; Section $s '06 · Program structure & roadmap'; Title $s 'Four phases, each behind a gate: no automatic group-wide roll-out' 9
$ph = @(
  @('Phase 1 · Q1 2027','Validate and lay the foundation','Make pilot figures verifiable. OT-DMZ and remote-access standard. Platform choice.','Gate: business case validated by Finance'),
  @('Phase 2 · Q2–Q3 2027','Second plant','One standard plant with S7 and ControlLogix. Test the edge blueprint; predictive maintenance on 2–3 lines.','Gate: value proven outside the pilot site'),
  @('Phase 3 · Q4 2027–2028','Roll-out standard lines','Plant by plant, by profile. Second use case. Isolate legacy lines first.','Gate: go/no-go per plant plus security review'),
  @('Phase 4 · from 2028','Legacy and depth','Retrofit or replace legacy lines on their own business case. New AI use cases.','Gate: investment decision per line group'))
for ($i=0; $i -lt 4; $i++) { $x = 48 + $i*218
  $b = BOX $s $x 112 204 150 $CANVAS
  FILL $b ($ph[$i][0] + "`r" + $ph[$i][1] + "`r" + $ph[$i][2]) 12 $BODY $false 1 1 14
  $tr = $b.TextFrame.TextRange; $tr.Paragraphs(1).Font.Size = [single]10; $tr.Paragraphs(1).Font.Color.RGB = $MUTED
  $tr.Paragraphs(2).Font.Name = $FONTB; $tr.Paragraphs(2).Font.Size = [single]15; $tr.Paragraphs(2).Font.Color.RGB = $INK
  $g = BOX $s $x 272 204 50 $TINT; FILL $g $ph[$i][3] 11 $NAVY $false 1 3 12 }
$b = BOX $s 48 352 864 88 $NAVY
FILL $b "Program structure`rDigital Operations Steering Group (chair COO)  →  one SAFe Agile Release Train 'Digital Operations'  →  product teams per use case + platform teams. Quarterly PI planning with the plants. Each gate is a board-level checkpoint on value, technical fit and safety." 12.5 $WHITE $true 1 3 16
$b.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]11; $b.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = (C 'BCD0E8')
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Timings are the group''s proposal, not from the case. Close by repeating the five decisions.'

$pres.SaveAs($Out, 24)
if ($PngDir) { $n = 0; foreach ($sl in $pres.Slides) { $n++; $sl.Export("$PngDir\v2-slide$n.png", 'PNG', 1600, 900) } }
$pres.Close()
"saved $Out"
