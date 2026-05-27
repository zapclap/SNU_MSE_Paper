# Codex Entry Instructions

This repository is a Codex workflow package for SNU/materials-engineering thesis and review-paper production.

When a user opens this downloaded folder as a Codex project and asks to create, revise, or automate a thesis/review paper:

1. Read `SKILL.md`.
2. If the user is starting from scratch or seems unsure what to provide, read `references/new-user-onboarding.md`.
3. Run `python3 scripts/first_run_questions.py` to get the beginner-friendly starting checklist.
4. Ask only the needed questions in plain user language. Do not say "preflight" to the user unless explaining the internal workflow.
5. Create or select a project folder under `projects/` before putting any manuscript-specific files anywhere.
6. Keep reusable workflow changes in the skill root and project-specific artifacts inside the selected project folder.
7. If GitHub sync is configured, run `python3 scripts/github_sync.py --message "<change summary>"` after validation.

