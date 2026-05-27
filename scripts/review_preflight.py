#!/usr/bin/env python3
"""Check review-thesis metadata before manuscript generation.

The script intentionally asks questions instead of letting a draft silently fill
front matter with vague placeholders. It accepts YAML via PyYAML when available
and falls back to a small text scan for common missing markers.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


MISSING_MARKERS = {"", "TODO", "TBD", "N/A", "null", "None", "[확인 필요]", "확인 필요"}

REQUIRED_FIELDS = [
    ("metadata.author", "저자 이름을 알려주세요."),
    ("metadata.korean_title", "국문 제목을 확정해 주세요."),
    ("metadata.english_title", "영문 제목을 확정해 주세요."),
    ("metadata.degree_name", "학위논문 종류를 알려주세요. 예: 석사학위논문 / 박사학위논문"),
    ("metadata.department", "소속 학과를 알려주세요."),
    ("metadata.major", "전공명을 알려주세요."),
    ("metadata.advisor", "지도교수 성함을 알려주세요."),
    ("metadata.submission_month", "논문 제출월을 알려주세요. 예: 2026년 8월"),
    ("metadata.approval_month", "논문 인준월을 알려주세요. 예: 2026년 8월"),
    ("metadata.committee.chair", "위원장 성함을 알려주세요."),
    ("metadata.committee.vice_chair", "부위원장 성함을 알려주세요."),
    ("metadata.committee.members", "심사위원 명단을 알려주세요. 외부심사위원이면 역할도 함께 적어주세요."),
    (
        "review_project.page_count_policy.allow_over_target",
        "쪽수는 목표 페이지에 엄격히 맞출까요, 아니면 표/그림 전체 페이지 규칙 때문에 목표보다 많아져도 괜찮을까요?",
    ),
]


def load_yaml(path: Path) -> dict[str, Any]:
    try:
        import yaml  # type: ignore
    except Exception:
        return fallback_parse(path)
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise SystemExit(f"Config root must be a mapping: {path}")
    return data


def parse_scalar(value: str) -> Any:
    cleaned = value.strip()
    if cleaned in {"", "null", "None"}:
        return None
    if cleaned in {"true", "True"}:
        return True
    if cleaned in {"false", "False"}:
        return False
    if cleaned == "[]":
        return []
    if cleaned.startswith("[") and cleaned.endswith("]"):
        inner = cleaned[1:-1].strip()
        if not inner:
            return []
        return [part.strip().strip('"').strip("'") for part in inner.split(",")]
    if (cleaned.startswith('"') and cleaned.endswith('"')) or (cleaned.startswith("'") and cleaned.endswith("'")):
        return cleaned[1:-1]
    if re.fullmatch(r"-?\d+", cleaned):
        return int(cleaned)
    return cleaned


def strip_comment(line: str) -> str:
    in_single = False
    in_double = False
    for idx, ch in enumerate(line):
        if ch == "'" and not in_double:
            in_single = not in_single
        elif ch == '"' and not in_single:
            in_double = not in_double
        elif ch == "#" and not in_single and not in_double:
            return line[:idx]
    return line


def fallback_parse(path: Path) -> dict[str, Any]:
    """Small indentation-based YAML fallback for simple project config files."""
    root: dict[str, Any] = {}
    stack: list[tuple[int, Any, Any, str | None]] = [(-1, root, None, None)]

    for raw in path.read_text(encoding="utf-8").splitlines():
        line = strip_comment(raw).rstrip()
        if not line.strip():
            continue
        indent = len(line) - len(line.lstrip(" "))
        content = line.strip()

        while stack and indent <= stack[-1][0]:
            stack.pop()
        parent = stack[-1][1]

        if content.startswith("- "):
            item_text = content[2:].strip()
            if not isinstance(parent, list):
                grand_parent = stack[-1][2]
                key = stack[-1][3]
                new_list: list[Any] = []
                if isinstance(grand_parent, dict) and key is not None:
                    grand_parent[key] = new_list
                    stack[-1] = (stack[-1][0], new_list, grand_parent, key)
                    parent = new_list
                else:
                    continue
            if ":" in item_text:
                key, value = item_text.split(":", 1)
                item: dict[str, Any] = {key.strip(): parse_scalar(value)}
                parent.append(item)
                stack.append((indent, item, parent, None))
            else:
                parent.append(parse_scalar(item_text))
            continue

        if ":" not in content:
            continue
        key, value = content.split(":", 1)
        key = key.strip()
        value = value.strip()
        parsed = parse_scalar(value)
        if value == "":
            parsed = {}
            stack.append((indent, parsed, parent, key))
        if isinstance(parent, dict):
            parent[key] = parsed

    return root


def get_path(data: dict[str, Any], dotted: str) -> Any:
    cur: Any = data
    for part in dotted.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return None
    return cur


def is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() in MISSING_MARKERS or "확인 필요" in value
    if isinstance(value, list):
        return len(value) == 0 or all(is_missing(v) for v in value)
    return False


def placeholder_approved(data: dict[str, Any]) -> bool:
    value = get_path(data, "review_project.placeholder_draft_approved")
    if isinstance(value, str):
        return value.strip().lower() in {"true", "yes", "1"}
    return bool(value)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("config", type=Path)
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON.")
    args = parser.parse_args()

    data = load_yaml(args.config)
    missing = [
        {"field": field, "question": question}
        for field, question in REQUIRED_FIELDS
        if is_missing(get_path(data, field))
    ]

    result = {
        "config": str(args.config.resolve()),
        "ok_to_generate": not missing or placeholder_approved(data),
        "placeholder_draft_approved": placeholder_approved(data),
        "missing": missing,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif missing:
        print("원고 생성 전에 아래 정보를 사용자에게 먼저 확인하세요:\n")
        for idx, item in enumerate(missing, start=1):
            print(f"{idx}. {item['question']} ({item['field']})")
        if placeholder_approved(data):
            print("\nplaceholder_draft_approved=true 이므로 초안 생성은 가능하지만, 최종본 전에는 반드시 반영하세요.")
    else:
        print("Preflight OK: required review-thesis metadata is complete.")

    return 0 if result["ok_to_generate"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
