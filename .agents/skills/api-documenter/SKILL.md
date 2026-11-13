---
name: api-documenter
description: REST, GraphQL, gRPC 등 API 인터페이스 사양을 문서화하고 OpenAPI(Swagger) 명세를 자동 생성/정리하는 스킬.
---

# API 문서화(API Documenter) 스킬

일관성 있고 명확한 API 사양서 작성 및 문서 최신화 가이드입니다.

## 문서화 표준 포맷
1. **엔드포인트 기본 정보**:
   - HTTP Method & Path (예: `POST /api/v1/auth/login`)
   - 기능 요약 및 설명
   - 인증/인가 요구사항 (Bearer Token, Role 등)
2. **요청 사양 (Request)**:
   - Path/Query Parameters
   - Request Headers
   - Request Body (JSON Schema, 타입, 필수 여부, 필드 설명)
3. **응답 사양 (Response)**:
   - 성공 응답 (HTTP 200, 201 등) 예시 및 필드 정의
   - 에러 응답 (HTTP 400, 401, 403, 404, 500 등) 코드 및 에러 메시지 규격
4. **OpenAPI / Swagger 연동**:
   - 필요 시 OpenAPI 3.0+ YAML/JSON 스펙 또는 프레임워크별 주석(Swagger JSDoc, FastAPI docs 등) 생성
