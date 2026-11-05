# Business case v2: OEE and predictive maintenance

**Latest version, 5 November 2026. This is the version we work in together.**

| Version | File | What it is |
|---|---|---|
| v1 | [`v1-concept-groep.docx`](v1-concept-groep.docx) | The group's first concept: scale model (10 plants, 120 lines, €2 bn) and revenue scenarios of 15–25% extra output |
| **v2** | this file + [`v2-business-case-model.xlsx`](v2-business-case-model.xlsx) | Sharpened version: keeps the v1 scale model, but calculates value from avoided downtime, maintenance and quality with public benchmarks, adds costs, payback and NPV |

Change assumptions in the yellow cells of the Excel model (sheet *Inputs*); the sheet *Model* recalculates. If you change the structure of the model, update [`build-model.ps1`](build-model.ps1) too, so the file can be rebuilt.

What changed from v1 to v2:

- v1 counted **revenue**; v2 counts **contribution margin and costs**.
- v1 assumed 15–25% extra output; v2 uses benchmark ranges for predictive maintenance (+0.4 to +1.1 OEE points).
- v2 adds the investment (≈ €24 m), recurring costs (≈ €3.5 m/year), payback and NPV.
- v2 keeps v1's scale model, plant list and legacy split.

The case gives no pilot figures. Every number below is labelled:

- **[case]** from the case description
- **[source]** public benchmark, linked
- **[group]** from the group's concept business case
- **[assumption]** our estimate, to be validated with pilot data, Finance, Operations and Sales

## 1. Scale model

| Item | Value | Label |
|---|---|---|
| Revenue | €2.0 bn | [group]; plausible next to Bosch Home Comfort's ≈ €4.4 bn ([source](https://www.bosch-presse.de/pressportal/de/en/business-year-2025-of-the-bosch-home-comfort-group-in-figures-280903.html)) |
| Plants / lines | 10 / 120, of which 18 legacy | [group] |
| Revenue per line | €16.7 m per year ≈ €4,200 per production hour | derived; 4,000 planned hours per line per year [assumption] |
| Contribution margin | 30% | [assumption] |
| OEE today | 70% | [assumption] |
| Unplanned downtime | 6% of planned time = 240 h per line per year | [assumption] |
| Share caused by equipment failure | 50% = 120 h per line | [assumption]; the rest is material, changeover, staffing |
| Value of one hour of downtime | €1,750 (€1,250 margin + €500 idle crew) | [assumption]; only if the extra output can be sold |
| Maintenance spend | €30 m per year (1.5% of revenue) | [assumption] |
| Internal scrap and rework | €30 m per year (1.5% of revenue) | [assumption] |

## 2. Annual benefits at full roll-out

| Benefit | Low | Base | High | Driver |
|---|---|---|---|---|
| Less unplanned downtime | €5.0 m | €7.6 m | €12.6 m | 20 / 30 / 50% less failure-driven downtime; McKinsey cites 30–50% for mature programmes ([source](https://reliamag.com/guides/predictive-maintenance-roi-benchmarks-what-the-studies-show/)) |
| Lower maintenance cost | €1.5 m | €2.4 m | €3.0 m | 5 / 8 / 10%; Deloitte cites 5–10% ([source](https://reliamag.com/guides/predictive-maintenance-roi-benchmarks-what-the-studies-show/)) |
| Less quality loss | €1.5 m | €2.1 m | €3.0 m | 5 / 7 / 10% [assumption] |
| **Gross benefits** | **€8 m** | **€12 m** | **€19 m** | |
| OEE effect | +0.4 pt | +0.6 pt | +1.0 pt | 70% → 70.4–71.0% |

The group concept assumed 15–25% extra output (€300–500 m revenue). That would require OEE to rise from 70% to 80–88%, which maintenance alone cannot deliver. The realistic benefit is a factor 25–40 lower.

## 3. Costs

| Item | One-off | Per year |
|---|---|---|
| Central data and AI platform, MLOps, standards | €5.0 m | €2.0 m |
| OT-DMZ, segmentation and edge per plant (10 × €0.5 m) | €5.0 m | |
| Sensors and retrofit, 102 standard lines × €50k | €5.1 m | |
| Retrofit, 18 legacy lines × €200k | €3.6 m | |
| Programme and integration | €3.0 m | |
| Training and change | €2.0 m | |
| Support, connectivity, licences per line; data and model team | | €1.5 m |
| **Total** | **≈ €24 m** | **≈ €3.5 m** |

All cost items are [assumption].

## 4. Return, base scenario, phased roll-out over three years

| | Year 1 | Year 2 | Year 3 | Year 4 | Year 5 |
|---|---|---|---|---|---|
| Investment | −€8.0 m | −€10.0 m | −€6.0 m | | |
| Net benefits | +€0.5 m | +€3.0 m | +€6.5 m | +€8.5 m | +€8.5 m |
| Cumulative | −€7.5 m | −€14.5 m | −€14.0 m | −€5.5 m | **+€3.0 m** |

- Payback ≈ 4.5 years (base), ≈ 3.5 years (high), not within 5 years (low).
- NPV at 8%: ≈ €0 after 5 years, ≈ +€10 m after 7 years.
- Per standard line: ≈ €70k net per year for ≈ €50k investment. The platform is the large fixed cost, so the case improves with every line connected. This is the financial argument for a group standard.

## 5. Sensitivities and messages

1. **Demand is the biggest sensitivity.** If extra output cannot be sold, an hour of downtime is worth only the idle crew (≈ €500) and the downtime benefit drops from €7.6 m to ≈ €2–3 m.
2. **False positives erode value.** Reliamag reports that a ~10% false-positive rate can wipe out the savings of a programme.
3. **Legacy lines decide separately.** ≈ 3 years payback per line but higher risk: isolate first, retrofit in phase 4.
4. **Positive but modest** is credible for a critical CFO and supports investment gates.

## 6. Data we still need from the pilot

| Data | Owner |
|---|---|
| Baseline and OEE per line (availability, performance, quality) | Operations, pilot site |
| Unplanned downtime, MTTR, maintenance hours before and after | Maintenance |
| Useful, false-positive and missed alerts | Maintenance + data team |
| Pilot costs versus recurring costs | Finance |
| Sellable share of extra output and contribution margin | Sales + Finance |
| Integration cost per plant type (standard, legacy) | OT/Engineering + Group IT |
