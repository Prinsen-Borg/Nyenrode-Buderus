# VERSION_CONTROL.md: Git strategy & collaboration

## 1. Remote

- **Remote:** `https://github.com/Prinsen-Borg/Nyenrode-Buderus`
- **Default branch:** `main`

```bash
git clone https://github.com/Prinsen-Borg/Nyenrode-Buderus.git
cd Nyenrode-Buderus
```

## 2. Branches

| Prefix | Use | Example |
|---|---|---|
| `feature/` | New section, analysis or tool | `feature/business-case-scenarios` |
| `fix/` | Correction | `fix/slide4-line-count` |
| `docs/` | Text-only changes | `docs/case-requirements` |

Small text fixes may go straight to `main`; anything that changes the argument goes through a pull request so the group can review it.

## 3. Commit messages (Conventional Commits)

`type(scope): short description in the imperative`

- `feat(analysis): add business case with public benchmarks`
- `docs(case): summarise requirements`
- `fix(workpage): correct legacy line count`
- `chore(repo): initial structure`

AI-assisted commits keep the co-author line.

## 4. Pull requests

1. Create a branch, make the change, commit and push.
2. Open a pull request with what changed and why.
3. At least one other group member reviews before merging. **AI may draft, but the group decides.**
