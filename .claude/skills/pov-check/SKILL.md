---
name: pov-check
description: Use before sharing any draft of the Buderus Point of View (paper, slides or workpage). Checks the draft against the case requirements, deliverable limits and the facts-versus-assumptions rule, and returns a short gap list.
metadata: owner Group 3 · version 0.1
---

## Goal
A short, honest list of what the draft still misses, so the group can fix it before the jury sees it.

## Steps
1. **Read the rules.** Read `CASE.md` (requirements R1–R10 and deliverables) and `AGENTS.md` §4.
2. **Read the draft** the user points to (a file in this repo, the Google Slides deck, or pasted text).
3. **Check coverage.** For each requirement R1–R10, mark: covered, partly, missing. Quote the section or slide where it is covered.
4. **Check limits.**
   - Paper: 6–8 pages, position-paper structure (introduction with thesis, counter-arguments, three arguments, conclusion, APA references).
   - Slides: max. 10, the five decisions early, titles that state a conclusion.
   - Board decisions: max. five, each a decision rather than a status update.
5. **Check figures.** Every number must be labelled case, public source (with link) or assumption. Flag any number presented as a pilot result: the case has none.
6. **Check consistency** with `deliverables/business-case/v2-business-case.md`, the latest presentation in `deliverables/presentation/`, `context/open-points.md` and the central message in `AGENTS.md` §4.
7. **Check hygiene.** No template leftovers, no AI leftovers ("If you want, I can…"), consistent company name (Buderus), one language.

## Output
- A table: requirement · status · where · what is missing.
- Max. five concrete fixes, most important first.
- No rewrite of the draft unless the user asks for it.

## Never
- Never invent pilot figures to fill a gap.
- Never add employer data or real people.
