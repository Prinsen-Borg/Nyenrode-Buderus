# Nyenrode Advanced IT: Buderus case (Group 3)

Shared workspace for the group assignment in the Advanced IT Program at Nyenrode Business University: a **Point of View** for the fictional company Buderus on integrating data analytics, AI and IoT into production and maintenance, presented to the board on **16 December 2026**.

When you open this repository in an AI assistant, it knows the case, the requirements and how our output should look.

## Where to find what

| I am looking for… | Go to |
|---|---|
| **The latest presentation** | [`deliverables/presentation/v2-board-presentation.pptx`](deliverables/presentation/v2-board-presentation.pptx) |
| **The latest business case** | [`deliverables/business-case/v2-business-case.md`](deliverables/business-case/v2-business-case.md) and the Excel model [`v2-business-case-model.xlsx`](deliverables/business-case/v2-business-case-model.xlsx) |
| The paper | [`deliverables/paper/`](deliverables/paper/README.md) (outline; not started) |
| Our earlier versions | [`deliverables/README.md`](deliverables/README.md): every version with date and what changed |
| The assignment and its ten requirements | [`CASE.md`](CASE.md) |
| A walk through the whole PoV (Dutch) | [`tools/pov-workpage/index.html`](tools/pov-workpage/index.html) |
| What the group still has to decide | [`context/open-points.md`](context/open-points.md) |
| Frameworks (BAM, COBIT, SAFe, Gartner, ...) | [`context/frameworks.md`](context/frameworks.md) |
| Real-company figures used as a plausibility check | [`context/company-reference.md`](context/company-reference.md) |
| Programme, deadlines and paper structure | [`context/programme.md`](context/programme.md) |
| The live working deck | [Google Slides](https://docs.google.com/presentation/d/1kC8puREHKKjRJ-3MsSN5h03dvNwbiGbqCRT9cOtdZKU/edit) |

## Earlier versions and the latest version

We keep every version so everyone recognises their earlier work, and we work together in the latest one.

| Deliverable | Earlier version (unchanged) | Latest version (work here) |
|---|---|---|
| Presentation | v1: export of the Google Slides working deck (5 Nov) | **v2**: 9-slide board deck with speaker notes |
| Business case | v1: the group's concept (Word, revenue scenarios) | **v2**: markdown + Excel model with value, costs, payback and NPV |
| Paper | – | not started |

Rule: the highest version number is leading. To make a new version, copy the latest file to the next number and add a line to [`deliverables/README.md`](deliverables/README.md).

## What is in this repository, and what is not

| Included | Where |
|---|---|
| Case description (group 3) | [`materials/case/`](materials/case) |
| Group drafts: approach document, architecture and OEE dashboard | [`materials/group-drafts/`](materials/group-drafts) |
| Writing guides: Harvard paper structure, position paper | [`materials/writing-guides/`](materials/writing-guides) |
| Module 1 readings (Business Activity Model, McKinsey, Anderson MacGyver, ...) | [`literature/`](literature/README.md) |

| Not included | Where to find it |
|---|---|
| Lecturers' presentations (Wijers, De Wit, Schrijvers, Rietdijk) | Howspace |
| Other groups' cases | Howspace |
| Anything from our employers | Never in this repository |

## Markdown files for AI assistants

| File | Purpose | Who reads it |
|---|---|---|
| [`AGENTS.md`](AGENTS.md) | **All rules**: central message, facts versus assumptions, limits, versions, publication | Copilot, Codex, Cursor and most coding agents read it automatically |
| [`CLAUDE.md`](CLAUDE.md) | Imports `AGENTS.md` and `CASE.md`; adds only Claude specifics (skills) | Claude Code and Claude apps (Claude does not read `AGENTS.md` by itself) |
| [`.github/copilot-instructions.md`](.github/copilot-instructions.md) | Pointer to `AGENTS.md`, no rules of its own | Older Copilot setups |
| [`CASE.md`](CASE.md) | The assignment, requirements R1–R10, deliverables | Everyone |
| [`DESIGN.md`](DESIGN.md) | Visual rules; the values are in [`design/tokens.css`](design/tokens.css) | Anyone making pages or slides |
| [`VERSION_CONTROL.md`](VERSION_CONTROL.md) | Branches, commits, pull requests | Everyone |

Rule of thumb: content goes in `AGENTS.md`; tool files only hold what is truly tool-specific. There is no separate file for ChatGPT, because ChatGPT does not read repository files.

## Using it with your assistant

- **Claude Code or the Claude app**: open the folder; `CLAUDE.md` loads the rules. Try *"Run pov-check on the latest presentation."*
- **GitHub Copilot (VS Code)**: open the folder; Copilot reads `AGENTS.md`.
- **ChatGPT**: create a ChatGPT project and paste `AGENTS.md` and `CASE.md` into the project instructions, or upload them as files.
- **Codex or Cursor**: open the folder; they read `AGENTS.md`.

## Tweaking the design

The look is neutral on purpose. Change colours and fonts in [`design/tokens.css`](design/tokens.css); the workpage follows at once. The presentation uses the same values at the top of [`build-v2.ps1`](deliverables/presentation/build-v2.ps1). See [`DESIGN.md`](DESIGN.md).

## Getting started

```bash
git clone https://github.com/Prinsen-Borg/Nyenrode-Buderus.git
```

Then open the folder in your assistant of choice. See [`VERSION_CONTROL.md`](VERSION_CONTROL.md) for how we commit and review.
