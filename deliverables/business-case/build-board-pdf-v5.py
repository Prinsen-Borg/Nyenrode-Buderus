# Bouwt v5-business-case-board.pdf (4 pagina's, board-niveau) uit één set aannames.
# Run: python3 build-board-pdf-v5.py   (vereist Node met Playwright + Chromium)
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "v5-business-case-board.html")
PDF = os.path.join(HERE, "v5-business-case-board.pdf")

# ---------------- Aannames (alle bedragen in € miljoen) ----------------
A = dict(omzet=2000, fabrieken=10, lijnen=120, legacy=18, oee0=0.65, oee1=0.68, capaciteit=0.05, verkoopbaar=0.50,
         marge=0.47, onderhoud=40, onderhoud_red=0.10, afkeur=40, afkeur_red=0.20, energie=30, energie_red=0.08,
         ramp=[0.20, 0.60, 1.00], invest=[20, 12, 5], run=[4, 6, 7], afschrijving=5, wacc=0.08)
BATEN = [
    ("Meer output: minder stilstand, hogere OEE", "IoT-sensoren + AI voorspellen storingen; live lijndata",
     A["omzet"] * A["capaciteit"] * A["verkoopbaar"] * A["marge"]),
    ("Lagere onderhoudskosten", "Onderhoud op conditie i.p.v. kalender", A["onderhoud"] * A["onderhoud_red"]),
    ("Minder afkeur en herstelwerk", "AI-camera's op kwaliteitskritische stations", A["afkeur"] * A["afkeur_red"]),
    ("Lagere energiekosten", "Submetering + AI-sturing van verbruik", A["energie"] * A["energie_red"]),
]
INVEST_SPLIT = [("Centraal data- en AI-platform, integratie ERP/MES", 8), ("OT-segmentatie, beveiliging en edge per fabriek", 7),
                ("Sensoren en lijnkoppeling, 120 lijnen (incl. 18 legacy)", 12), ("AI-camera's kwaliteit (≈ 60 stations)", 5),
                ("Energiemetering, 10 fabrieken", 2), ("Verandering, training en programma", 3)]
RUN_SPLIT = [("Platform, cloud en licenties", 3), ("Data- en AI-team, OT-security monitoring", 2.5), ("Beheer hardware en modellen", 1.5)]
assert sum(v for _, v in INVEST_SPLIT) == sum(A["invest"])
assert abs(sum(v for _, v in RUN_SPLIT) - A["run"][2]) < 1e-9

RR = sum(b[2] for b in BATEN)


def reken(rr, invest, run):
    rows, cum, pb = [], 0.0, None
    for i in range(3):
        baten = rr * A["ramp"][i]
        netto = baten - run[i] - invest[i]
        prev, cum = cum, cum + netto
        if pb is None and cum >= 0:
            pb = i + (-prev) / netto
        rows.append(dict(baten=baten, run=run[i], invest=invest[i], netto=netto, cum=cum))
    if pb is None:  # jaar 4 op vol tempo
        n4 = rr - run[2]
        pb = 3 + (-cum) / n4
    return rows, pb


ROWS, PB = reken(RR, A["invest"], A["run"])
NPV3 = sum(r["netto"] / (1 + A["wacc"]) ** (i + 1) for i, r in enumerate(ROWS))
INV_TOT = sum(A["invest"])
# eenvoudige P&L
afschr = []
for i in range(3):
    afschr.append(sum(A["invest"][:i + 1]) / A["afschrijving"])
PL = [dict(baten=r["baten"], run=-r["run"], ebitda=r["baten"] - r["run"], afschr=-afschr[i], ebit=r["baten"] - r["run"] - afschr[i])
      for i, r in enumerate(ROWS)]
# gevoeligheid
SENS = [("Basis", PB), ("Vraag lager: 25% i.p.v. 50% verkoopbaar", reken(RR - BATEN[0][2] / 2, A["invest"], A["run"])[1]),
        ("Kosten 25% hoger", reken(RR, [x * 1.25 for x in A["invest"]], [x * 1.25 for x in A["run"]])[1]),
        ("Uitrol 1 jaar later (baten 0 / 20 / 60%)", None)]
# vertraging: baten schuiven een jaar
cum, pbd = 0.0, None
ramp_d = [0, 0.2, 0.6, 1.0, 1.0]; inv_d = A["invest"] + [0, 0]; run_d = A["run"] + [7, 7]
for i in range(5):
    n = RR * ramp_d[i] - run_d[i] - inv_d[i]; prev, cum = cum, cum + n
    if pbd is None and cum >= 0: pbd = i + (-prev) / n
SENS[3] = (SENS[3][0], pbd)


def m(x, d=1):
    s = f"{abs(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if x < -1e-9 else "") + s


def pct(x):
    return f"{x*100:.0f}%".replace(".", ",")


def jr(x):
    return f"{x:.1f}".replace(".", ",") + " jaar"


YEARS = ["2027", "2028", "2029"]

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


def foot(n):
    return (f'<div class="foot"><span>Buderus · Business case Digitale Fabriek 2027–2029 · ter besluitvorming Raad van Bestuur · bedragen in € miljoen</span>'
            f'<span>{n}</span></div>')


# ---------------- Pagina 1: uitgangspunten ----------------
p1_rows = [
    ("Omzet alle fabrieken", "€ 2,0 mld", "Buderus", "Past tussen peers: Ariston € 2,7 mld, BDR Thermea € 1,9 mld"),
    ("Fabrieken / productielijnen", "10 / 120 (18 legacy)", "Buderus", "Uitrol in drie golven, legacy als laatste"),
    ("OEE vandaag → doel", f"{pct(A['oee0'])} → {pct(A['oee1'])}", "aanname", "Typische fabriek 60–67%, world class 85%"),
    ("Extra capaciteit", pct(A["capaciteit"]), "aanname", "Gevolg van +3 OEE-punten; ruim onder best in class"),
    ("Verkoopbaar deel extra capaciteit", pct(A["verkoopbaar"]), "aanname", "Markt groeit ≈ 3% per jaar; Sales valideert"),
    ("Brutomarge op extra omzet", pct(A["marge"]), "bron", "Materiaal ≈ 45% van omzet (Ariston)"),
    ("Onderhoudskosten per jaar", f"€ {A['onderhoud']} mln, −{pct(A['onderhoud_red'])}", "bron", "Deloitte: −5–10%; wij nemen de bovenkant"),
    ("Afkeur en herstelwerk per jaar", f"€ {A['afkeur']} mln, −{pct(A['afkeur_red'])}", "bron", "P&G met AI-camera's: −10–20%"),
    ("Energiekosten per jaar", f"€ {A['energie']} mln, −{pct(A['energie_red'])}", "bron", "IEA: AI bespaart 3–10% energie"),
    ("Baten bereikt per jaar", " / ".join(pct(x) for x in A["ramp"]), "aanname", "2027 / 2028 / 2029, volgt de uitrol"),
    ("Investering / jaarlijkse runkosten", f"€ {INV_TOT} mln / € {m(A['run'][2],0)} mln", "aanname", "Opbouw op pagina 4"),
]
p1 = f"""
<section class="page">
  <div class="eyebrow">2 · Uitgangspunten en aannames</div>
  <h1>Voorzichtig gerekend en getoetst aan de sector</h1>
  <p class="lead">De pilot laat het potentieel zien, maar is nog geen basis voor groepsbrede cijfers. Daarom rekenen we met
  interne kerngegevens, openbare cijfers van vergelijkbare bedrijven en sectorbenchmarks. Alle aannames valideren we in de eerste uitrolgolf.</p>
  <div class="cols" style="grid-template-columns:2.2fr 1fr">
    <div>
      <h2>Aannames</h2>
      <table><tr><th>Aanname</th><th class="n">Waarde</th><th>Label</th><th>Onderbouwing</th></tr>
      {''.join(f'<tr><td>{a}</td><td class="n">{b}</td><td class="tag {"b" if c=="bron" else ("a" if c=="aanname" else "")}">{c}</td><td class="small">{d}</td></tr>' for a,b,c,d in p1_rows)}
      </table>
    </div>
    <div>
      <h2>Waarom dit realistisch is</h2>
      <table>
        <tr><th></th><th class="n">Buderus</th><th class="n">Sector</th></tr>
        <tr><td>OEE</td><td class="n">65% → 68%</td><td class="n">60–67%; top 85%</td></tr>
        <tr><td>Afkeur</td><td class="n">−20%</td><td class="n">−10–20%; top −41%</td></tr>
        <tr><td>Energie</td><td class="n">−8%</td><td class="n">−3–10%; top −28%</td></tr>
        <tr><td>Extra groei</td><td class="n">+2,5%</td><td class="n">markt ≈ +3%/jaar</td></tr>
      </table>
      <p class="small" style="margin-top:3mm">We rekenen alleen met de marge op verkoopbare extra productie, niet met omzet.
      Onze doelen liggen binnen de normale sectorbandbreedte en ruim onder wat de beste fabrieken ter wereld halen.</p>
      <div class="flag" style="margin-top:4mm"><b>Labels.</b> <i>Buderus</i> = interne kerngegevens · <i>bron</i> = openbare benchmark (pagina 5) ·
      <i>aanname</i> = onze inschatting, te valideren met pilotdata, Finance, Operations en Sales.</div>
    </div>
  </div>
  {foot(2)}
</section>"""

# ---------------- Pagina 2: business case uitgelegd ----------------
levers = [
    ("Voorspellend onderhoud", "IoT-sensoren op machines; AI voorspelt storingen dagen vooruit. Reparatie in gepland onderhoud.", BATEN[0][2] * 0.55 + BATEN[1][2]),
    ("Live OEE per lijn", "Lijndata via edge naar MES; één OEE-definitie voor alle fabrieken. AI wijst de grootste verliezen aan.", BATEN[0][2] * 0.45),
    ("AI-kwaliteitscontrole", "Camera's met AI vangen fouten bij het station, niet pas aan het eind van de lijn of bij de klant.", BATEN[2][2]),
    ("AI-energiesturing", "Meten per lijn; AI schakelt stilstaand verbruik en pieken af. Lagere kosten en minder CO₂.", BATEN[3][2]),
]
p2 = f"""
<section class="page">
  <div class="eyebrow">3 · Het innovatieprogramma</div>
  <h1>Eén digitaal fundament, vier toepassingen</h1>
  <p class="lead">Het programma koppelt de fabrieken aan één gezamenlijk data- en AI-platform. Daarop draaien vier toepassingen
  die stilstand, afkeur en energieverbruik verlagen. De besturing van de machines blijft lokaal: AI adviseert, mensen beslissen.</p>
  <div class="box"><b>Bestuurskeuze.</b> Standaardiseer het digitale fundament op groepsniveau; fabrieken passen het toe binnen
  hun lokale OT-, veiligheids- en continuïteitsgrenzen; schaal per fabriek en per toepassing op zodra waarde, technische geschiktheid
  en continuïteit zijn aangetoond.</div>
  <h2>Vier toepassingen en hun jaarlijkse waarde bij volledige uitrol</h2>
  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5mm;margin-bottom:6mm">
    {''.join(f'<div class="lever"><b>{t}</b><span class="small">{d}</span><div class="v" style="margin-top:2mm">€ {m(v)} mln per jaar</div></div>' for t,d,v in levers)}
  </div>
  <div class="cols">
    <div>
      <h2>Waarom dit werkt</h2>
      <ul style="margin:0;padding-left:5mm">
        <li><b>Marge, geen omzet.</b> Alleen verkoopbare extra output telt, tegen de brutomarge.</li>
        <li><b>Schaal betaalt het platform.</b> Het fundament is een vaste kost; elke extra lijn en toepassing voegt waarde toe tegen lage meerkosten.</li>
        <li><b>Veilig en continu.</b> OT-segmentatie per fabriek; bij uitval van de verbinding draait de fabriek gewoon door.</li>
      </ul>
    </div>
    <div>
      <h2>Uitrol in drie golven, met poorten</h2>
      <ol class="steps">
        <li><b>2027:</b> pilotfabriek + 2 fabrieken. Meet OEE, juistheid van AI-meldingen en adoptie.</li>
        <li><b>2028:</b> standaardlijnen in alle fabrieken, als golf 1 de criteria haalt.</li>
        <li><b>2029:</b> legacy-lijnen en de laatste toepassingen. De plantmanager houdt go/no-go bij veiligheidsrisico.</li>
      </ol>
    </div>
  </div>
  {foot(3)}
</section>"""

# ---------------- Pagina 3: kosten en baten 3 jaar ----------------
def row(label, vals, cls="", tot=None):
    t = sum(vals) if tot is None else tot
    return f'<tr class="{cls}"><td>{label}</td>' + "".join(f'<td class="n">{m(v)}</td>' for v in vals) + f'<td class="n">{m(t)}</td></tr>'


p3 = f"""
<section class="page">
  <div class="eyebrow">4 · Kosten en baten 2027–2029</div>
  <h1>€ {INV_TOT} mln investering, terugverdiend na ≈ {jr(PB)}</h1>
  <div class="kpis">
    <div class="kpi"><div class="v">€ {INV_TOT} mln</div><div class="l">Investering, gespreid over 3 jaar</div></div>
    <div class="kpi"><div class="v">€ {m(RR)} mln</div><div class="l">Baten per jaar bij volledige uitrol (2029)</div></div>
    <div class="kpi"><div class="v">{jr(PB)}</div><div class="l">Terugverdientijd (kasstroom, vóór belasting)</div></div>
    <div class="kpi"><div class="v">€ {m(NPV3)} mln</div><div class="l">Netto contante waarde na 3 jaar (8%)</div></div>
  </div>
  <div class="cols" style="grid-template-columns:1.25fr 1fr">
    <div>
      <h2>Kasstroom</h2>
      <table>
        <tr><th>€ mln</th>{''.join(f'<th class="n">{y}</th>' for y in YEARS)}<th class="n">Totaal</th></tr>
        {''.join(row(b[0], [b[2]*A['ramp'][i] for i in range(3)]) for b in BATEN)}
        {row('Totaal baten', [r['baten'] for r in ROWS], 'tot')}
        {row('Runkosten (platform, licenties, team)', [-r['run'] for r in ROWS])}
        {row('Investering', [-r['invest'] for r in ROWS])}
        {row('Netto kasstroom', [r['netto'] for r in ROWS], 'tot')}
        {row('Cumulatief', [r['cum'] for r in ROWS], 'key', tot=ROWS[-1]['cum'])}
      </table>
      <h2 style="margin-top:3.5mm">Effect op de winst- en verliesrekening</h2>
      <table>
        <tr><th>€ mln</th>{''.join(f'<th class="n">{y}</th>' for y in YEARS)}<th class="n">Totaal</th></tr>
        {row('Baten', [p['baten'] for p in PL])}
        {row('Runkosten', [p['run'] for p in PL])}
        {row('EBITDA-effect', [p['ebitda'] for p in PL], 'tot')}
        {row(f'Afschrijving ({A["afschrijving"]} jaar)', [p['afschr'] for p in PL])}
        {row('EBIT-effect', [p['ebit'] for p in PL], 'tot')}
      </table>
    </div>
    <div>
      <h2>Waar de investering in gaat</h2>
      <table>
        {''.join(f'<tr><td>{a}</td><td class="n">{m(v,0)}</td></tr>' for a,v in INVEST_SPLIT)}
        <tr class="tot"><td>Totaal investering</td><td class="n">{INV_TOT}</td></tr>
        {''.join(f'<tr><td class="small">Run: {a}</td><td class="n small">{m(v)}</td></tr>' for a,v in RUN_SPLIT)}
      </table>
      <h2 style="margin-top:3.5mm">Gevoeligheid terugverdientijd</h2>
      <table>
        {''.join(f'<tr class="{"key" if i==0 else ""}"><td>{a}</td><td class="n">{jr(v)}</td></tr>' for i,(a,v) in enumerate(SENS))}
      </table>
      <p class="small" style="margin-top:2mm">Vraag is de grootste gevoeligheid: kan de extra output niet verkocht worden, dan schuift de terugverdientijd naar ruim 3 jaar.</p>
    </div>
  </div>
  {foot(4)}
</section>"""

# ---------------- Pagina 4: referenties ----------------
REFS = [
    ("Omzet en kostenstructuur", "Ariston: € 2,7 mld omzet, 32 fabrieken, materiaal 45%, EBITDA 11,5%, groei +3%", "Omzetschaal, marge 47%", "Ariston Group (2026), FY2025 results"),
    ("Peers", "BDR Thermea: € 1,9 mld, 14 fabrieken, EBITDA ≈ 8%. Bosch Home Comfort: ≈ € 4,4 mld", "€ 2 mld is plausibel", "BDR Thermea (2026); Bosch (2026)"),
    ("OEE", "Typische fabriek 60–67%; world class 85%", "Van 65% naar 68%", "WorkTrek (2025); TeepTrak (2026); Nakajima (1988)"),
    ("Kosten van stilstand", "Ongeplande stilstand kost de 500 grootste bedrijven ≈ 11% van de omzet", "Stilstand is de grootste hefboom", "Siemens (2024), True cost of downtime"),
    ("Voorspellend onderhoud", "Uptime +10–20%; onderhoudskosten −5–10%", "Onderhoud −10%", "Deloitte (2024), Predictive maintenance"),
    ("AI-kwaliteitscontrole", "Afkeur −10–20% met AI-camera's (P&G met Siemens)", "Afkeur −20%", "Manufacturing Digital (2024)"),
    ("AI-energiesturing", "AI-procesoptimalisatie bespaart 3–10% energie", "Energie −8%", "IEA (2025), Energy and AI"),
    ("Best in class (plafond)", "WEF Lighthouse-fabrieken: defecten −41%, energie −28%", "Wij blijven ruim onder dit plafond", "World Economic Forum (2025)"),
    ("Kosten sensoren", "Draadloos meetpunt ≈ $ 200–500 geïnstalleerd; abonnement $ 200–800 per machine per jaar", "≈ € 100k per lijn incl. engineering", "Tractian (2025)"),
    ("Kapitaalkosten", "Gemiddelde WACC 7,9%", "Disconto 8%", "KPMG (2026), Cost of capital study"),
]
p4 = f"""
<section class="page">
  <div class="eyebrow">5 · Referenties en benchmarks</div>
  <h1>Onze aannames liggen binnen de sectorbandbreedte</h1>
  <p class="lead">Openbare cijfers dienen als toets op plausibiliteit, niet als pilotresultaat. Leveranciers- en secundaire bronnen
  gebruiken we als bandbreedte.</p>
  <table>
    <tr><th style="width:17%">Onderwerp</th><th style="width:38%">Benchmark</th><th style="width:20%">Onze aanname</th><th>Bron</th></tr>
    {''.join(f'<tr><td><b>{a}</b></td><td>{b}</td><td>{c}</td><td class="small">{d}</td></tr>' for a,b,c,d in REFS)}
  </table>
  <p class="small" style="margin-top:4mm">Volledige bronvermelding met links is op verzoek beschikbaar. Buderus-kerngegevens zijn afkomstig uit de
  eigen administratie; aannames worden per uitrolgolf vervangen door gemeten waarden.</p>
  {foot(5)}
</section>"""

TITLE = f"""
<section class="page title" style="padding-top:12mm">
  <div class="eyebrow">Buderus · Point of View CIO · ter besluitvorming</div>
  <h1 style="font-size:34pt;margin-top:4mm;padding-top:4mm;border-top:3px solid var(--blue);display:inline-block">Digitale Fabriek 2027–2029</h1>
  <p class="lead" style="font-size:15pt;max-width:200mm">Business case voor data, AI en IoT in productie en onderhoud</p>
  <div class="kpis" style="margin-top:8mm;max-width:230mm">
    <div class="kpi"><div class="v">€ {INV_TOT} mln</div><div class="l">Investering over 3 jaar</div></div>
    <div class="kpi"><div class="v">€ {m(RR)} mln</div><div class="l">Baten per jaar vanaf 2029</div></div>
    <div class="kpi"><div class="v">{jr(PB)}</div><div class="l">Terugverdientijd</div></div>
    <div class="kpi"><div class="v">5</div><div class="l">Besluiten gevraagd</div></div>
  </div>
  <table style="width:140mm;margin-top:8mm;font-size:10pt">
    <tr><td style="color:var(--sec);width:30mm">Aan</td><td>Raad van Bestuur: Bert Smit (CEO), Hubert Zwiers (CFO)</td></tr>
    <tr><td style="color:var(--sec)">Ter info</td><td>Raad van Commissarissen</td></tr>
    <tr><td style="color:var(--sec)">Van</td><td>Louw Fokker, CIO</td></tr>
    <tr><td style="color:var(--sec)">Datum</td><td>16 december 2026</td></tr>
  </table>
  <p class="small" style="margin-top:30mm;border-top:1px solid var(--line);padding-top:1.5mm">Fictieve casus · Advanced IT Program no. 15, Nyenrode Business University · Groep 3</p>
</section>"""

DECISIONS = [
    ("Architectuur", "Eén centraal data- en AI-platform voor de groep; besturing van machines (PLC/SCADA) blijft lokaal in elke fabriek."),
    ("Governance", "Een Digital Operations Steering Group (Operations, Finance, IT, OT, Security, fabrieken) met heldere mandaten; de plantmanager houdt go/no-go bij risico voor veiligheid of continuïteit."),
    ("Financiering", f"€ {INV_TOT} mln in drie golven (€ {A['invest'][0]} / {A['invest'][1]} / {A['invest'][2]} mln). Elke golf wordt pas vrijgegeven als de vorige de afgesproken waarde-, veiligheids- en continuïteitscriteria haalt."),
    ("Risicogrenzen", "Vaste spelregels voor OT-koppelingen, remote access, wijzigingen, data-integriteit en AI-advies, ingericht volgens COBIT 2019."),
    ("Mensen", "Capaciteit en opleiding voor IT/OT-, data- en AI-vaardigheden: een centraal team plus een digital champion per fabriek."),
]
SUMMARY = f"""
<section class="page">
  <div class="eyebrow">1 · Management summary</div>
  <h1 style="margin-bottom:5mm">Vraag aan de Raad van Bestuur: investeer € {INV_TOT} mln, terugverdiend in {jr(PB)}</h1>
  <div class="cols" style="grid-template-columns:1.15fr 1fr;gap:8mm">
    <div>
      <h2>Aanleiding</h2>
      <p style="margin:0 0 3mm">De energietransitie en onze groei vergroten de vraag, en daarmee de druk op efficiëntie, kwaliteit en kosten
      in onze tien fabrieken. De pilot met sensordata in één fabriek laat zien dat voorspellend onderhoud, hogere beschikbaarheid en betere
      kwaliteit binnen bereik liggen. De Raad van Commissarissen heeft budget vrijgemaakt voor IoT, AI, dataplatformen en training.
      De vraag is nu hoe we dit groepsbreed opschalen zonder de productiecontinuïteit in gevaar te brengen.</p>
      <h2>Ons advies</h2>
      <div class="box" style="margin-bottom:3mm">Standaardiseer het digitale fundament op groepsniveau; fabrieken passen het toe binnen hun lokale OT-,
      veiligheids- en continuïteitsgrenzen; schaal per fabriek en per toepassing op zodra waarde, technische geschiktheid en continuïteit zijn aangetoond.</div>
      <h2>Wat het oplevert</h2>
      <p style="margin:0">Vier toepassingen op één platform: voorspellend onderhoud, live OEE, AI-kwaliteitscontrole en AI-energiesturing.
      De OEE stijgt van {pct(A['oee0'])} naar {pct(A['oee1'])}, goed voor {pct(A['capaciteit'])} extra capaciteit. Vanaf 2029 levert dit
      € {m(RR)} mln per jaar op, waarvan € {m(PL[2]['ebitda'])} mln EBITDA na runkosten. Netto contante waarde na 3 jaar: € {m(NPV3)} mln.</p>
    </div>
    <div>
      <h2>Vijf besluiten die wij u vragen</h2>
      <table>
        {''.join(f'<tr><td style="width:7mm;font-weight:600;color:var(--blue)">{i+1}</td><td><b>{a}.</b> {b}</td></tr>' for i,(a,b) in enumerate(DECISIONS))}
      </table>
      <div class="flag" style="margin-top:4mm"><b>Risico.</b> De grootste gevoeligheid is de vraag. Als maar een kwart van de extra productie
      verkoopbaar is, wordt de terugverdientijd {jr(SENS[1][1])}. Bij 25% hogere kosten is dat {jr(SENS[2][1])}. De poorten tussen de golven beperken dit risico.</div>
    </div>
  </div>
  {foot(1)}
</section>"""

html = f"<!doctype html><html lang='nl'><head><meta charset='utf-8'><title>Business case Digitale Fabriek 2027–2029</title><style>{CSS}</style></head><body>{TITLE}{SUMMARY}{p1}{p2}{p3}{p4}</body></html>"
open(HTML, "w", encoding="utf-8").write(html)
JS = ("const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();const p=await b.newPage();"
      "await p.goto('file://'+process.argv[1]);await p.pdf({path:process.argv[2],preferCSSPageSize:true,printBackground:true});await b.close();})();")
subprocess.run(["node", "-e", JS, HTML, PDF], check=True)
print(f"payback {PB:.2f} jr, NPV3 {NPV3:.1f}, baten {RR:.1f}", [round(s[1], 2) for s in SENS])
print("saved", PDF)
