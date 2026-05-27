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


def fallback_parse(path: Path) -> dict[str, Any]:
    """Tiny fallback for simple metadata blocks when PyYAML is unavailable."""
    text = path.read_text(encoding="utf-8")
    data: dict[str, Any] = {"metadata": {"committee": {}}}
    current = None
    in_committee = False
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        if not raw.startswith(" ") and raw.endswith(":"):
            current = raw[:-1]
            data.setdefault(current, {})
            in_committee = False
            continue
        if current == "metadata" and re.match(r"^  committee:\s*$", raw):
            data["metadata"]["committee"] = {}
            in_committee = True
            continue
        m = re.match(r"^  ([A-Za-z0-9_]+):\s*(.*)$", raw)
        if current and m:
            key, val = m.groups()
            cleaned = val.strip().strip('"')
            if cleaned == "[]":
                parsed: Any = []
            elif cleaned == "" and key == "committee":
                parsed = {}
            else:
                parsed = cleaned
            data.setdefault(current, {})[key] = parsed
            in_committee = key == "committee"
        m = re.match(r"^    ([A-Za-z0-9_]+):\s*(.*)$", raw)
        if current == "metadata" and in_committee and m:
            key, val = m.groups()
            cleaned = val.strip().strip('"')
            if key == "members" and cleaned in {"", "[]"}:
                data["metadata"].setdefault("committee", {})[key] = []
            else:
                data["metadata"].setdefault("committee", {})[key] = cleaned
        if current == "metadata" and in_committee and re.match(r"^      -\s+", raw):
            data["metadata"].setdefault("committee", {}).setdefault("members", []).append(raw.strip())
    return data


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
