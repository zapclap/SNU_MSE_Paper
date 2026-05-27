# GitHub Sync Workflow

Use this reference when the user asks to keep the SNU thesis-production skill, review workflow, or project artifacts synchronized with GitHub.

## Purpose

The sync workflow commits reusable skill changes and project-local artifacts without mixing manuscript-specific files into the skill root. It is generic and can be reused for any later materials-engineering review or thesis project.

## Required Checks

Before the first push, confirm or obtain:

- GitHub repository URL, such as `git@github.com:USER/REPO.git` or `https://github.com/USER/REPO.git`
- Git commit identity: user name and email
- Whether generated project outputs such as DOCX, PDF, and source figures should be tracked

If any of these are missing, prepare the local repository and ask the user for the missing values instead of guessing.

## Standard Command

From the skill root:

```bash
python scripts/github_sync.py \
  --set-remote git@github.com:USER/REPO.git \
  --user-name "USER NAME" \
  --user-email "USER_EMAIL@example.com" \
  --message "Update thesis production workflow"
```

After the remote and identity are configured once, future syncs can use:

```bash
python scripts/github_sync.py --message "Describe the change"
```

Use `--dry-run` to inspect the commands without changing git state. Use `--no-push` when the user wants a local commit only.

## Behavior

`scripts/github_sync.py`:

1. Initializes a git repository on branch `main` if needed.
2. Verifies git identity before committing.
3. Adds or updates the `origin` remote when `--set-remote` is provided.
4. Stages all non-ignored files by default.
5. Creates a commit only when staged changes exist.
6. Pushes to GitHub if a remote is configured and `--no-push` is not set.

## What To Track

Track reusable workflow files:

- `SKILL.md`
- `references/`
- `scripts/`
- `agents/`
- `assets/`

Track manuscript projects inside `projects/<project-name>/` when the user wants reproducible outputs. Keep all topic-specific configs, notes, ledgers, scripts, DOCX/PDF files, figures, and QA files inside that project folder.

Do not add topic-specific scripts or manuscript outputs directly to the skill root.

## After Future Workflow Edits

When the user changes a reusable workflow rule:

1. Update the skill/reference/script first.
2. Run relevant validation, such as `python -m py_compile scripts/*.py`.
3. Run `python scripts/github_sync.py --message "<concise change summary>"`.
4. If sync cannot push because the remote or credentials are missing, tell the user exactly which value or login step is required.

