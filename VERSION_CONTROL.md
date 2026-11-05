# VERSION_CONTROL.md: Git strategy & collaboration

## 1. Remote

- **Remote:** `https://github.com/Prinsen-Borg/Nyenrode-Buderus`
- **Branch:** `main`, and only `main`

```bash
git clone https://github.com/Prinsen-Borg/Nyenrode-Buderus.git
cd Nyenrode-Buderus
```

## 2. Push straight to `main`

We work directly on `main` and push right away, so everyone sees the newest version at once. No feature branches, no pull requests waiting for someone to approve them in GitHub.

This is safe because nothing gets lost:

- Every deliverable keeps its older versions (`v1`, `v2`, ...); a change becomes a **new** version file, never an edit of an old one.
- Git keeps the full history; every earlier state can be brought back.
- Nobody force-pushes or rewrites history (see `AGENTS.md` §3).

Before you push:

1. `git pull` first, so you build on the latest work.
2. Commit with a clear message (below).
3. `git push`. If the push is rejected because someone pushed in the meantime, `git pull` again and push once more. Never use `--force`.

## 3. Commit messages (Conventional Commits)

`type(scope): short description in the imperative`

- `feat(presentation): add v3 with shorter governance slide`
- `docs(case): summarise requirements`
- `fix(workpage): correct legacy line count`
- `chore(repo): initial structure`

AI-assisted commits keep the co-author line.

## 4. Undoing something

Made a mistake? Fix it with a new commit, for example `git revert <commit>`, and push. Older versions in `deliverables/` and the Git history stay intact.
