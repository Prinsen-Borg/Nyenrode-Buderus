# Bouwt v5-vergelijking-met-concept.pdf: één pagina, voor de groep. Run: python3 build-vergelijking-v5.py
import os, subprocess
HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "v5-vergelijking-met-concept.html")
PDF = os.path.join(HERE, "v5-vergelijking-met-concept.pdf")
CSS = """
@page { size: A4 landscape; margin: 14mm 16mm; }
:root { --ink:#161616; --sec:#525252; --help:#6f6f6f; --line:#e0e0e0; --layer:#f4f4f4; --blue:#0f62fe; --blue-bg:#edf5ff;
        --green:#24a148; --green-bg:#defbe6; --red:#da1e28; --amber-bg:#fcf4d6; }
* { box-sizing:border-box; }
html,body { margin:0; background:#fff; color:var(--ink); font-family:"IBM Plex Sans","Helvetica Neue",Arial,sans-serif; font-size:9.8pt; line-height:1.4; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
.page { page-break-after:always; height:181mm; position:relative; }
.page:last-child { page-break-after:auto; }
.eyebrow { font-size:8.5pt; letter-spacing:.06em; text-transform:uppercase; color:var(--sec); }
h1 { font-weight:300; font-size:21pt; margin:1.5mm 0 1mm; line-height:1.15; }
.lead { color:var(--sec); font-size:10pt; margin:0 0 4mm; max-width:230mm; }
h2 { font-size:11pt; font-weight:600; margin:0 0 1.5mm; padding-bottom:1.5mm; border-bottom:2px solid var(--blue); }
.cols { display:grid; grid-template-columns:1fr 1fr; gap:9mm; }
.cols3 { display:grid; grid-template-columns:1fr 1fr 1fr; gap:7mm; }
table { width:100%; border-collapse:collapse; font-size:9pt; }
th { background:var(--line); text-align:left; font-weight:600; padding:1.2mm 2mm; }
td { padding:1mm 2mm; border-bottom:1px solid var(--line); vertical-align:top; }
td.n, th.n { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
tr.tot td { font-weight:600; border-top:1px solid #8d8d8d; }
tr.key td { background:var(--blue-bg); font-weight:600; }
.tag { font-size:8pt; color:var(--sec); white-space:nowrap; }
.tag.b { color:#0043ce; } .tag.a { color:#8a3800; }
.kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:4mm; margin-bottom:4mm; }
.kpi { background:var(--layer); padding:2.5mm 4mm; border-top:3px solid var(--blue); }
.kpi .v { font-size:17pt; font-weight:300; line-height:1.1; }
.kpi .l { font-size:8.5pt; color:var(--sec); }
.box { background:var(--blue-bg); padding:3.5mm 4.5mm; border-left:3px solid var(--blue); margin-bottom:5mm; }
.flag { background:var(--amber-bg); padding:3mm 4mm; border-left:3px solid #b28600; font-size:9.5pt; }
.lever { background:var(--layer); padding:3.5mm 4mm; }
.lever b { display:block; margin-bottom:1mm; }
.lever .v { color:var(--blue); font-weight:600; }
ol.steps { margin:0; padding-left:5mm; } ol.steps li { margin-bottom:1.5mm; }
.foot { position:absolute; bottom:0; left:0; right:0; font-size:8pt; color:var(--help); border-top:1px solid var(--line); padding-top:1.5mm; display:flex; justify-content:space-between; }
.title .band { position:absolute; top:0; left:0; width:60mm; height:3px; background:var(--blue); }
.small { font-size:8pt; color:var(--sec); }
"""
SAME = [
    ("Schaal", "€ 2 mld, 10 fabrieken, 120 lijnen, 18 legacy"),
    ("Architectuur", "Centrale diensten; SCADA/PLC blijven lokaal; edge-gateway per fabriek"),
    ("Rol van AI", "AI adviseert, de medewerker beslist"),
    ("Governance", "Steering Group; plantmanager heeft lokaal go/no-go"),
    ("Fasering", "Valideren → beperkte uitrol → opschalen op criteria (v5: drie golven met poorten)"),
    ("Vijf besluiten", "Architectuur, steering group, gefaseerde financiering, risicogrenzen, capaciteit en opleiding"),
]
DIFF = [
    ("Productiewinst", "15 / 20 / 25%", "5% (OEE 65% → 68%)", "k"),
    ("Wat telt als waarde", "Extra omzet € 300–500 mln", "Marge op verkoopbaar deel, ± € 24 mln per jaar", "k"),
    ("Verkoopbaar deel", "100% (75% als voorbeeld)", "50%", ""),
    ("Reikwijdte", "OEE + voorspellend onderhoud", "+ AI-kwaliteitscontrole en AI-energiesturing", "k"),
    ("Kosten", "Posten genoemd, geen bedragen", "€ 37 mln investering, € 7 mln run per jaar", ""),
    ("Terugverdientijd / NCW", "Nog niet te berekenen", "2,5 jaar; € 9,4 mln na 3 jaar", ""),
    ("Legacy-lijnen", "Zelfde winst als andere lijnen", "Laatste golf, hogere kosten per lijn", ""),
    ("Locatie", "Amsterdam als hoofdlocatie, tien hoofdsteden", "Alleen \"centraal\" en \"tien fabrieken\"", "t"),
    ("Productiemix", "72 accu- en 48 assemblagelijnen", "Niet genoemd", ""),
    ("Benodigde pilotdata", "Aparte lijst (§12)", "Alleen \"valideren in golf 1\"", "t"),
]
ACTIONS = [
    ("Bewust accepteren", "5% in plaats van 15–25%: het concept vraagt een OEE van bijna world class en 4 à 6× de groei van peers. Leg dit in de presentatie kort uit."),
    ("Bewust accepteren", "Twee extra toepassingen leveren samen ± € 10 mln per jaar en maken de 2,5 jaar haalbaar. Met alleen OEE en voorspellend onderhoud ligt de terugverdientijd hoger."),
    ("Terugzetten (v6)", "Amsterdam als centrale locatie en een korte lijst \"gegevens die we in golf 1 meten\". Beide zijn onbedoeld weggevallen."),
]
html = f"""<!doctype html><html lang="nl"><head><meta charset="utf-8"><title>Vergelijking business case met concept</title><style>{CSS}
body {{ font-size:9pt; }} table {{ font-size:8.4pt; }} td {{ padding:.8mm 1.6mm; }} th {{ padding:1mm 1.6mm; }}
h1 {{ font-size:18pt; }} .lead {{ margin-bottom:3mm; }}
.pill {{ font-size:7.5pt; font-weight:600; padding:.3mm 1.6mm; white-space:nowrap; }}
.pill.k {{ background:var(--amber-bg); color:#8a3800; }} .pill.ok {{ background:var(--green-bg); color:#0e6027; }}
</style></head><body>
<section class="page">
  <div class="eyebrow">Intern · Groep 3 · Toets van de business case (v5) aan het concept-docx</div>
  <h1>Zelfde opzet en besturing, andere cijfers en een bredere reikwijdte</h1>
  <p class="lead">De business case voor de Raad van Bestuur bouwt voort op het groepsconcept. De structuur is overgenomen; de cijfers zijn
  getoetst aan peers en sectorbenchmarks en daardoor fors bijgesteld.</p>
  <div class="cols" style="grid-template-columns:1fr 2fr;gap:7mm">
    <div>
      <h2>Gelijk gebleven</h2>
      <table>{''.join(f'<tr><td style="width:24mm"><b>{a}</b></td><td>{b}</td></tr>' for a,b in SAME)}</table>
    </div>
    <div>
      <h2>Afwijkingen</h2>
      <table>
        <tr><th style="width:34mm">Onderwerp</th><th>Concept</th><th>Business case v5</th><th></th></tr>
        {''.join(f'<tr><td><b>{a}</b></td><td>{b}</td><td>{c}</td><td>{"<span class=\'pill k\'>groepskeuze</span>" if k=="k" else ("<span class=\'pill ok\'>terugzetten</span>" if k=="t" else "")}</td></tr>' for a,b,c,k in DIFF)}
      </table>
    </div>
  </div>
  <h2 style="margin-top:3mm">Wat de groep moet doen</h2>
  <table>{''.join(f'<tr><td style="width:34mm"><span class="pill {"ok" if a.startswith("Terug") else "k"}">{a}</span></td><td>{b}</td></tr>' for a,b in ACTIONS)}</table>
  <div class="foot"><span>Buderus · Groep 3 · intern werkdocument · fictieve casus</span><span>Concept: materials/group-drafts · Business case: deliverables/business-case/v5</span></div>
</section></body></html>"""
open(HTML, "w", encoding="utf-8").write(html)
JS = ("const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();const p=await b.newPage();"
      "await p.goto('file://'+process.argv[1]);await p.pdf({path:process.argv[2],preferCSSPageSize:true,printBackground:true});await b.close();})();")
subprocess.run(["node", "-e", JS, HTML, PDF], check=True)
print("saved", PDF)
