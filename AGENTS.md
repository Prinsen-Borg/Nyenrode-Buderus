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
.github/copilot-instructions.md  ← pointer only (Copilot reads AGENTS.md itself)
CASE.md                          ← the assignment, requirements and deliverables
DESIGN.md                        ← visual rules; values live in design/tokens.css
VERSION_CONTROL.md               ← branching and commit conventions
README.md                        ← human introduction, including what is where
deliverables/                    ← what we hand in, with earlier and latest versions
  presentation/                  ← v1 = working deck export, v2 = latest board deck
  business-case/                 ← v1 = group concept, v2 = latest (md + Excel model)
  paper/                         ← the PoV paper (6–8 pages)
context/                         ← programme, frameworks, company reference, open points
literature/                      ← module 1 readings
materials/                       ← source files: case description, group drafts, writing guides
design/tokens.css                ← colours and fonts, used by the workpage
tools/pov-workpage/              ← HTML workpage that walks through the whole PoV
.claude/skills/                  ← procedures an agent follows (e.g. pov-check)
```

### Versions

- Files in `deliverables/` start with a version prefix. **The highest version is the one we work in.** Older versions stay so people can recognise their earlier work; do not edit them.
- To start a new version, copy the latest file to the next number, note what changed at the top or in `deliverables/README.md`, and update the README table.

## 3. First conversation: "how do I work with this repo?"

When someone asks how to work with this repository (or opens with a general question such as "what is this?" or "where do I start?"), answer **in their language** and follow this shape. Keep it short; read `deliverables/README.md` first so the versions you name are the current ones.

1. **One line on what this is**: the shared workspace of Group 3 for the Buderus Point of View, due 16 December 2026.
2. **The latest versions**, with paths: the presentation, the business case (markdown + Excel model), the workpage, and the paper status.
3. **What you can do here**, as a short numbered list:
   1. *Look*: I summarise the latest deck, business case or workpage for you.
   2. *Give feedback*: tell me what should be better and I make a new version (next version number) next to the old one, visible to everyone.
   3. *Change the design*: colours and fonts are in `design/tokens.css`, the rules in `DESIGN.md`; I update both.
   4. *Check*: I test a piece against the ten requirements in `CASE.md` (`pov-check`).
   5. *Write*: I help with the paper; the outline is in `deliverables/paper/README.md`.
4. **The reassurance message**, in these words for Dutch speakers (translate for others):
   > Probeer gerust dingen uit, je kunt niks stukmaken. Elke versie blijft bewaard en alles is terug te draaien. Dit is een oefening, dus juist goed om te experimenteren.
5. **End with one question**: "Wat wil je doen?" (or the equivalent in their language).

Be honest about what your tool can do. If you cannot commit and push (for example ChatGPT or claude.ai in a browser), say so: you can read the repository and produce a new version as a file to download, and the person (or a teammate with access) adds it to the repository.

### What makes "you cannot break it" true

These rules are not optional; the reassurance above depends on them.

- **Never overwrite or delete an older version** in `deliverables/`. Always create the next version number.
- **Never force-push, rebase pushed commits or rewrite history.** Every state must stay recoverable.
- **Never delete files from `materials/` or `literature/`.**
- If a change goes wrong, fix it with a new commit (for example `git revert`), not by rewriting history.

## 4. Working rules for agents

1. **Start from `CASE.md`.** Every output must trace back to a requirement in the case. Use the `pov-check` skill to verify coverage.
2. **Keep the central message.** The PoV argues one board choice: *standardise the digital foundation at group level; plants apply it within local OT, safety and continuity limits; scale per plant and per use case only when value, technical fit and continuity are proven.* Do not drift from it without the group deciding so.
3. **Separate facts from assumptions.** The case gives **no pilot figures**. Never present an invented number as a pilot result. Label every figure as *case*, *public source* (with link) or *assumption*. See `deliverables/business-case/v2-business-case.md`.
4. **Fiction stays fiction.** Buderus resembles the real Buderus brand of the Bosch Home Comfort Group. Use public figures only as a plausibility check and cite them. Do not imply the case describes the real company's internal situation, and do not name real employees.
5. **Respect the deliverable limits.** Paper 6–8 pages (position-paper structure), presentation max 10 slides / 15 minutes, max 5 board decisions.
6. **Use the course frameworks** where they fit: Business Activity Model (multimodality), Business Model & AI Canvas, COBIT 2019, Gartner Top Strategic Technology Trends 2026 (top-3 only, motivated), SAFe. See `context/frameworks.md`.
7. **Write plainly.** Short sentences, concrete numbers, no hype words. Dutch for content the group presents unless the group chooses English; English for repository files.
8. **Make small, reviewable changes.** Follow `VERSION_CONTROL.md`.
9. **Leave no secrets and no personal data.** No passwords, tokens, personal contact details or employer-confidential information.

## 5. Data and publication

This repository is **public on GitHub**. On 5 November 2026 the owner decided to publish the case description, group drafts, writing guides and module 1 readings, because the case is fictional and the repository is meant for practice. For anything new, check:

| Content | Allowed? |
|---|---|
| Our own analysis, notes, pages and summaries | Yes |
| Group drafts from classmates | Yes, as agreed for this group |
| Case description and module 1 readings | Yes, already included |
| Lecturers' presentations | Not included; link to Howspace |
| Employer data (Moba or other) | **Never** |

## 6. When in doubt

Ask the user one precise question instead of guessing on something that is hard to undo. For anything easy to redo, make a reasonable assumption, state it and proceed.
