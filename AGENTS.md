# AGENTS.md: Instructions for AI agents working on the Buderus case

This file is the **single source of truth** for every AI assistant that works in this repository (Claude, GitHub Copilot, Cursor, Codex and others). Tool-specific files such as `CLAUDE.md` only point here.

---

## 1. What this repository is

- The shared workspace of **Group 3** in the **Advanced IT Program at Nyenrode Business University** (cohort no. 15, 2026).
- The group assignment is a **Point of View (PoV)** for the fictional company **Buderus**: how to optimise IT/OT governance so that data analytics, AI and IoT become a structural part of production and maintenance, with grip on safety, continuity and scalability.
- The PoV is presented to the board (and a professional jury) on **16 December 2026**.

| Is | Is not |
|---|---|
| Case material, group drafts and analysis for the Buderus PoV | A place for anyone's employer data (Moba or otherwise) |
| Context files that let an AI assistant help consistently | A record of real Buderus or Bosch internal information |
| Small tools such as the PoV workpage | The final deliverable store (that is the group workspace on Howspace) |

## 2. Repository map

```
AGENTS.md                        ← you are here: rules for every agent
CLAUDE.md                        ← Claude-specific additions; imports this file
.github/copilot-instructions.md  ← Copilot-specific additions; points to this file
CASE.md                          ← the assignment, requirements and deliverables
DESIGN.md                        ← visual rules for pages and slides in this repo
VERSION_CONTROL.md               ← branching and commit conventions
README.md                        ← human introduction
context/                         ← programme, frameworks, company reference, open points
analysis/                        ← worked analysis (business case, ...)
materials/                       ← source files: case, group drafts, writing guides
tools/pov-workpage/              ← HTML workpage that walks through the whole PoV
.claude/skills/                  ← procedures an agent follows (e.g. pov-check)
```

## 3. Working rules for agents

1. **Start from `CASE.md`.** Every output must trace back to a requirement in the case. Use the `pov-check` skill to verify coverage.
2. **Keep the central message.** The PoV argues one board choice: *standardise the digital foundation at group level; plants apply it within local OT, safety and continuity limits; scale per plant and per use case only when value, technical fit and continuity are proven.* Do not drift from it without the group deciding so.
3. **Separate facts from assumptions.** The case gives **no pilot figures**. Never present an invented number as a pilot result. Label every figure as *case*, *public source* (with link) or *assumption*. See `analysis/business-case.md`.
4. **Fiction stays fiction.** Buderus resembles the real Buderus brand of the Bosch Home Comfort Group. Use public figures only as a plausibility check and cite them. Do not imply the case describes the real company's internal situation, and do not name real employees.
5. **Respect the deliverable limits.** Paper 6–8 pages (position-paper structure), presentation max 10 slides / 15 minutes, max 5 board decisions.
6. **Use the course frameworks** where they fit: Business Activity Model (multimodality), Business Model & AI Canvas, COBIT 2019, Gartner Top Strategic Technology Trends 2026 (top-3 only, motivated), SAFe. See `context/frameworks.md`.
7. **Write plainly.** Short sentences, concrete numbers, no hype words. Dutch for content the group presents unless the group chooses English; English for repository files.
8. **Make small, reviewable changes.** Follow `VERSION_CONTROL.md`.
9. **Leave no secrets and no personal data.** No passwords, tokens, personal contact details or employer-confidential information.

## 4. Data and publication

This repository is **public on GitHub** unless the owner has made it private. Before adding a file, check:

| Content | Allowed? |
|---|---|
| Our own analysis, notes, pages and summaries | Yes |
| Group drafts from classmates | Only with the group's consent |
| Nyenrode course material (case text, lecturer slides, readings) | Only if the repository is private; otherwise link or summarise |
| Third-party literature (McKinsey, Gartner, HBR, ...) | No: reference it in `context/literature.md` |
| Employer data (Moba or other) | **Never** |

## 5. When in doubt

Ask the user one precise question instead of guessing on something that is hard to undo. For anything easy to redo, make a reasonable assumption, state it and proceed.
