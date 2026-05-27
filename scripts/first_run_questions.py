#!/usr/bin/env python3
"""Print beginner-friendly starting questions for a thesis/review project."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


QUESTION_GROUPS = [
    {
        "id": "project_intent",
        "title": "첫 시작 확인",
        "intro": "시작 전에 필요한 정보만 먼저 확인합니다. 모르는 항목은 '나중에'라고 답해도 됩니다.",
        "questions": [
            {
                "id": "document_type",
                "question": "만들 문서 종류가 무엇인가요? 예: 리뷰논문, 학위논문, 기존 DOCX 수정, 일반 논문 초안.",
            },
            {
                "id": "topic_mode",
                "question": "주제는 정해져 있나요, 아니면 Codex가 최신/핫한 주제를 찾아야 하나요?",
            },
            {
                "id": "format_scope",
                "question": "사용할 학교/학과 양식이 있나요? 서울대 재료공학 기본값을 쓸까요?",
            },
            {
                "id": "language",
                "question": "작성 언어는 무엇인가요? 예: 한국어, 영어, 한영 혼합.",
            },
            {
                "id": "page_count_policy",
                "question": "목표 쪽수가 있나요? 그 쪽수를 엄격히 맞출까요, 아니면 표/그림 배치 때문에 더 길어져도 괜찮나요?",
            },
            {
                "id": "image_mode",
                "question": "그림은 새로 그릴까요, 논문 원문 그림을 가져와 출처를 남길까요, 둘 다 쓸까요?",
            },
            {
                "id": "deliverables",
                "question": "최종 산출물은 무엇이 필요한가요? 예: DOCX, PDF, 둘 다.",
            },
        ],
    },
    {
        "id": "thesis_metadata",
        "title": "학위논문 정보",
        "intro": "학위논문 또는 SNU 리뷰논문이면 아래 정보가 필요합니다.",
        "questions": [
            {"id": "author", "question": "저자 이름을 알려주세요."},
            {
                "id": "titles",
                "question": "국문 제목과 영문 제목을 알려주세요. 아직 모르면 Codex가 제안해도 되는지 알려주세요.",
            },
            {
                "id": "degree_name",
                "question": "학위논문 종류를 알려주세요. 예: 석사학위논문 / 박사학위논문.",
            },
            {"id": "department_major", "question": "학과와 전공을 알려주세요."},
            {"id": "advisor", "question": "지도교수 성함을 알려주세요."},
            {
                "id": "months",
                "question": "논문 제출월과 인준월을 알려주세요. 예: 2026년 8월.",
            },
            {
                "id": "committee",
                "question": "위원장, 부위원장, 심사위원 이름과 역할을 알려주세요.",
            },
            {
                "id": "acknowledgements",
                "question": "감사의 글을 넣을까요? 기본값은 제외입니다.",
            },
        ],
    },
]


def detect_projects(root: Path) -> list[str]:
    projects_root = root / "projects"
    if not projects_root.exists():
        return []
    return sorted(path.name for path in projects_root.iterdir() if path.is_dir())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=".", help="Downloaded workflow folder.")
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON.")
    args = parser.parse_args()

    root = Path(args.root).expanduser().resolve()
    payload = {
        "root": str(root),
        "existing_projects": detect_projects(root),
        "question_groups": QUESTION_GROUPS,
        "next_steps": [
            "Create/select a project folder under projects/.",
            "Store answers in the project-local config files.",
            "Run scripts/review_preflight.py before final manuscript generation.",
        ],
    }

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    if payload["existing_projects"]:
        print("기존 프로젝트 폴더가 있습니다:")
        for name in payload["existing_projects"]:
            print(f"- {name}")
        print("먼저 기존 프로젝트를 이어서 수정할지, 새 프로젝트를 만들지 물어보세요.")
        print()

    for group in QUESTION_GROUPS:
        print(f"[{group['title']}]")
        print(group["intro"])
        for idx, item in enumerate(group["questions"], start=1):
            print(f"{idx}. {item['question']}")
        print()

    print("답변을 받으면 projects/<project-name>/ 아래에 config, ledgers, output 폴더를 만들고 진행합니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
