---
name: git-commit-helper
description: Git 변경사항을 분석하여 Conventional Commits 규칙에 맞춘 명확하고 체계적인 커밋 메시지와 PR 템플릿을 생성하는 스킬.
---

# Git 커밋 헬퍼(Git Commit Helper) 스킬

일관되고 의미 있는 버전 관리 이력을 생성하기 위한 가이드입니다.

## 커밋 컨벤션 형식
```text
<type>(<scope>): <subject>

[optional body]

[optional footer(s)]
```

### Type 목록
- `feat`: 새로운 기능 추가
- `fix`: 버그 수정
- `docs`: 문서 수정
- `style`: 코드 포맷팅, 세미콜론 누락 등 코드 변경 없는 스타일 수정
- `refactor`: 코드 리팩토링 (기능 변경 없음)
- `test`: 테스트 코드 추가 및 수정
- `chore`: 빌드 업무, 패키지 매니저 설정 등 기타 작업

## 작성 원칙
1. 제목은 50자 이내로 명확하게 요약
2. 본문에는 변경 이유와 주요 변경 사항을 불릿 포인트로 작성
3. 관련된 이슈 번호가 있을 경우 푸터에 명시 (`Fixes #123`)
