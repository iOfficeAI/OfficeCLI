# Branch Strategy

## Current State
- **main**: mirrors upstream + compliance docs. Pushed to origin.
- **agent**: working branch for cli-jaw modifications. Created and pushed to origin.
- Both branches currently at commit `9d60efc`.

## Branches
- **main**: mirrors upstream iOfficeAI/OfficeCLI main. Pull-only after initial setup.
- **agent**: our primary working branch. All cli-jaw modifications go here.
- **feature/***: optional short-lived branches for isolated work.

## Sync Procedure
```bash
git fetch upstream
git checkout main
git merge --ff-only upstream/main
git checkout agent
git rebase main
```

## Rules
- Never push to upstream
- All CJK/cli-jaw changes on agent branch
- Tag releases as `cjk-v{version}` on agent branch
