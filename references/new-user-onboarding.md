# New User Onboarding

Use this reference when a user downloads this folder, opens it as a Codex project, and wants to start a thesis/review paper without knowing the workflow.

## User-Facing Tone

Do not start by saying "preflight." Say:

> 시작 전에 필요한 정보만 먼저 확인할게요. 답을 모르는 항목은 "나중에"라고 적어도 됩니다. 최종본에는 빈칸을 넣지 않고, 확인된 값만 반영하겠습니다.

Ask in Korean by default when the user writes in Korean. Use plain language and explain that the questions prevent placeholders from entering the manuscript.

## First Questions

If existing folders are present under `projects/`, first ask:

- 기존 프로젝트를 이어서 수정할까요, 아니면 새 프로젝트를 만들까요?

Ask this first batch before creating a manuscript:

1. 만들 문서 종류가 무엇인가요? 예: 리뷰논문, 학위논문, 기존 DOCX 수정, 일반 논문 초안.
2. 주제는 정해져 있나요, 아니면 Codex가 최신/핫한 주제를 찾아야 하나요?
3. 사용할 학교/학과 양식이 있나요? 서울대 재료공학 기본값을 쓸까요?
4. 작성 언어는 무엇인가요? 예: 한국어, 영어, 한영 혼합.
5. 목표 쪽수가 있나요? 그 쪽수를 엄격히 맞출까요, 아니면 표/그림 배치 때문에 더 길어져도 괜찮나요?
6. 그림은 새로 그릴까요, 논문 원문 그림을 가져와 출처를 남길까요, 둘 다 쓸까요?
7. 최종 산출물은 무엇이 필요한가요? 예: DOCX, PDF, 둘 다.

If the answer indicates an SNU thesis/review thesis, ask the second batch:

1. 저자 이름
2. 국문 제목과 영문 제목. 제목을 아직 모르면 주제 확정 후 Codex가 제안해도 되는지 확인
3. 학위논문 종류. 예: 석사학위논문 / 박사학위논문
4. 학과와 전공
5. 지도교수 성함
6. 논문 제출월과 인준월
7. 위원장, 부위원장, 심사위원 이름과 역할
8. 감사의 글 포함 여부. 기본값은 제외

## How To Act On Answers

After the user answers:

1. If the user wants a new project, run `scripts/init_review_project.py` to create a project folder. If they want to continue an existing project, select that project folder.
2. Put all answers into the project-local `config/review-config.yaml` and `config/project-config.yaml`.
3. Run `scripts/review_preflight.py <project-folder>/config/review-config.yaml`.
4. If anything is still missing, ask only those missing questions.
5. Do not generate a final DOCX/PDF until required metadata and page-count strictness are confirmed, unless the user explicitly approves a placeholder draft.

## Unknown Answers

If the user says they do not know a value:

- For topic/title: offer to research and propose options.
- For degree, department, advisor, committee, submission/approval month: keep it unresolved and ask again before final generation.
- For page count: recommend "flexible" when figure/table pages must be isolated and centered.
- For image mode: recommend generated schematics plus cited source figures only when source traceability is important.

## Folder Behavior

New users should not need to understand the internal file layout. Codex should create and manage:

```text
projects/<project-name>/
  config/
  ledgers/
  notes/
  scripts/
  output/docx/
  output/pdf/
  output/figures/
  output/source_figures/
  output/qa/
  output/page_maps/
```

Do not place manuscript-specific drafts, scripts, images, or outputs in the repository root.
