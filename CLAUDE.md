# CLAUDE.md: Claude-specific instructions

The shared rules for every AI tool live in **AGENTS.md**. Claude loads them here:

@AGENTS.md
@CASE.md

This file only adds what is specific to Claude (the Claude apps, Claude Code and Cowork).

## Skills

- Skills live in `.claude/skills/<name>/SKILL.md`. Claude discovers them automatically.
- `pov-check`: run it before sharing any draft of the paper, the slides or the workpage.

## Visual output

- Read `DESIGN.md` first. The workpage in `tools/pov-workpage/index.html` is the reference implementation.
- Slides for the jury follow the same palette; the group may switch to a Nyenrode template if one is provided.

## Working documents outside this repo

- The group's working deck is a Google Slides file (link in `README.md`). Only slides 1–12 are group content; the rest is a layout template.
- Claude may read it, but changes to it go through the user.
