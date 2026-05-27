#!/usr/bin/env python3
"""Initialize, commit, and optionally push this skill workspace to GitHub."""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path


def run(
    args: list[str],
    cwd: Path,
    *,
    dry_run: bool = False,
    check: bool = True,
    capture: bool = False,
) -> subprocess.CompletedProcess[str]:
    if dry_run:
        print("$ " + " ".join(args))
        return subprocess.CompletedProcess(args, 0, "", "")
    proc = subprocess.run(
        args,
        cwd=cwd,
        text=True,
        capture_output=capture,
        check=False,
    )
    if check and proc.returncode != 0:
        if capture:
            if proc.stdout:
                print(proc.stdout, end="")
            if proc.stderr:
                print(proc.stderr, end="", file=sys.stderr)
        raise subprocess.CalledProcessError(proc.returncode, args)
    return proc


def git_text(args: list[str], cwd: Path) -> str:
    proc = run(["git", *args], cwd, capture=True, check=False)
    if proc.returncode != 0:
        return ""
    return proc.stdout.strip()


def in_git_repo(cwd: Path) -> bool:
    return git_text(["rev-parse", "--is-inside-work-tree"], cwd) == "true"


def ensure_repo(cwd: Path, branch: str, dry_run: bool) -> None:
    if not in_git_repo(cwd):
        run(["git", "init", "-b", branch], cwd, dry_run=dry_run)
        return

    current = git_text(["branch", "--show-current"], cwd)
    if current and current != branch:
        run(["git", "checkout", "-B", branch], cwd, dry_run=dry_run)


def ensure_identity(
    cwd: Path,
    *,
    user_name: str | None,
    user_email: str | None,
    dry_run: bool,
) -> None:
    if user_name:
        run(["git", "config", "user.name", user_name], cwd, dry_run=dry_run)
    if user_email:
        run(["git", "config", "user.email", user_email], cwd, dry_run=dry_run)

    name = user_name or git_text(["config", "user.name"], cwd)
    email = user_email or git_text(["config", "user.email"], cwd)
    if dry_run:
        return
    if not name or not email:
        raise SystemExit(
            "Missing git identity. Run again with "
            "--user-name 'Your Name' --user-email 'you@example.com', "
            "or configure git user.name and user.email."
        )


def set_remote(cwd: Path, remote: str, url: str, dry_run: bool) -> None:
    existing = git_text(["remote", "get-url", remote], cwd)
    if existing:
        if existing != url:
            run(["git", "remote", "set-url", remote, url], cwd, dry_run=dry_run)
    else:
        run(["git", "remote", "add", remote, url], cwd, dry_run=dry_run)


def has_remote(cwd: Path, remote: str) -> bool:
    return bool(git_text(["remote", "get-url", remote], cwd))


def has_staged_changes(cwd: Path) -> bool:
    proc = run(["git", "diff", "--cached", "--quiet"], cwd, check=False)
    return proc.returncode == 1


def has_any_changes(cwd: Path) -> bool:
    return bool(git_text(["status", "--porcelain"], cwd))


def default_message() -> str:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"Update thesis production workflow ({stamp})"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Commit the current skill workspace and push to a configured "
            "GitHub remote when available."
        )
    )
    parser.add_argument(
        "--root",
        default=".",
        help="Workspace root that should be treated as the git repository.",
    )
    parser.add_argument("--branch", default="main", help="Branch to use.")
    parser.add_argument("--remote", default="origin", help="Git remote name.")
    parser.add_argument(
        "--set-remote",
        metavar="URL",
        help="Add or update the GitHub remote before syncing.",
    )
    parser.add_argument(
        "--message",
        default=default_message(),
        help="Commit message for this sync batch.",
    )
    parser.add_argument(
        "--include",
        action="append",
        default=[],
        help="Path to stage. May be repeated. Defaults to the whole repo.",
    )
    parser.add_argument(
        "--no-push",
        action="store_true",
        help="Commit locally but skip git push.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print git commands without changing files.",
    )
    parser.add_argument(
        "--user-name",
        help="Set repository-local git user.name before committing.",
    )
    parser.add_argument(
        "--user-email",
        help="Set repository-local git user.email before committing.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = Path(args.root).expanduser().resolve()
    if not root.exists():
        raise SystemExit(f"Root does not exist: {root}")

    ensure_repo(root, args.branch, args.dry_run)
    ensure_identity(
        root,
        user_name=args.user_name,
        user_email=args.user_email,
        dry_run=args.dry_run,
    )

    if args.set_remote:
        set_remote(root, args.remote, args.set_remote, args.dry_run)

    stage_paths = args.include or ["."]
    run(["git", "add", "--", *stage_paths], root, dry_run=args.dry_run)

    if args.dry_run:
        print("# Dry run only; no commit or push was made.")
        return 0

    if not has_staged_changes(root):
        if not has_any_changes(root):
            print("No changes to sync.")
        else:
            print("No staged changes to commit. Check .gitignore or include paths.")
        return 0

    run(["git", "commit", "-m", args.message], root)
    print(f"Committed: {args.message}")

    if args.no_push:
        print("Skipped push because --no-push was set.")
        return 0

    if not has_remote(root, args.remote):
        print(
            f"Committed locally. Add a GitHub remote with "
            f"--set-remote <url> to push to {args.remote}."
        )
        return 0

    run(["git", "push", "-u", args.remote, args.branch], root)
    print(f"Pushed {args.branch} to {args.remote}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

