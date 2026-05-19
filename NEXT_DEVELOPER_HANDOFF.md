# Next Developer Handoff

작성일: 2026-05-19

## 목적

이 문서는 `check.html -> qa.html -> WBS 후처리`로 이어지는 산출물 정리 파이프라인의 현재 구현 상태를 다음 작업자가 바로 이어받을 수 있도록 정리한다.

현재 저장소:

```text
https://github.com/jea0123/atoz_ai_integrated
```

현재 주 작업 경로:

```text
C:\atoz_ai_integrated
```

## 실행

프로젝트 루트 `.env`에는 Ollama 설정이 필요하다.

```env
OLLAMA_BASE_URL="http://192.168.1.7:11434"
OLLAMA_MODEL="exaone3.5:7.8b"
```

로컬 실행 예시:

```powershell
cd C:\atoz_ai_integrated
.\venv310\Scripts\python.exe web_app.py --host 127.0.0.1 --port 8001
```

접속:

```text
http://127.0.0.1:8001/check.html
http://127.0.0.1:8001/qa.html
```

## 현재 반영된 핵심 변경

### 1. check.html 결과 폴더 유지

기존 `/api/folder-apply` 흐름은 덤프 결과를 ZIP으로 만들고 중간 복사폴더를 정리하는 구조였다.

현재는 QA 후속 작업을 위해 결과 복사폴더를 유지한다.

결과 폴더 기본 위치:

```text
C:\atoz_ai_integrated\web_runtime\results\folder-dumps
```

반복 실행 시 결과 폴더명은 버전 형태로 생성된다.

```text
테스트_v0.1
테스트_v0.2
테스트_v0.3
```

원본 폴더는 직접 수정하지 않고, 복사된 결과 폴더 안에서 문서관리표준 반영 작업이 진행된다.

### 2. check.html에서 qa.html로 연결

`check.html`의 `덤프 후 반영` 성공 결과에는 `dump_root`가 포함된다.

프론트는 이 경로를 `localStorage`에 저장하고, `qa.html?dump_root=...`로 넘길 수 있다.

### 3. qa.html 결과 폴더 QA 생성

`qa.html`에는 `결과폴더 QA 생성` 영역이 추가되었다.

입력:

- `check 결과 폴더`
- `QA 원천 폴더 경로`
- `QA 원천 폴더 선택`
- `화면설계서 폴더 경로`
- `단위시험 폴더 경로`
- `통합시험 폴더 경로`
- `화면설계서 폴더 선택`

`QA 원천 폴더`는 화면설계서, 단위시험케이스, 통합시험시나리오가 함께 있는 폴더를 한 번에 받기 위한 입력이다.

지원 확장자:

```text
화면설계서: hwp, hwpx, pdf
단위시험케이스: hwpx
통합시험시나리오: xlsx
```

폴더 업로드는 `webkitdirectory`를 사용하며, 선택 후 화면에 처리 대상 파일 목록과 파일 수가 표시된다.

### 4. QA 매칭 기준

QA 파일 탐색은 현재 Ollama가 아니라 규칙 기반이다.

매칭 기준:

- 파일명/경로에 포함된 `SFR-...` 요구사항 ID
- 역할 키워드
- 일부 fallback에서는 문서 내부 텍스트의 `SFR-...` ID

역할별 주요 키워드:

```text
화면설계서:
사용자인터페이스설계서, 사용자 인터페이스 설계서, 화면설계서, 화면정의서, UI설계서

단위시험케이스:
단위시험케이스, 단위시험 케이스, 단위테스트, 단위 테스트, unit test case

통합시험시나리오:
통합시험시나리오, 통합시험 시나리오, 통합시험결과서, 통합시험 결과서, 통합테스트, integration test scenario
```

주의:

- `단위시험결과서`는 `단위시험케이스`와 별도 산출물이므로 단위시험케이스 후보에서 제외한다.
- `통합시험결과서`는 현 테스트 폴더에서 통합시험 계열 양식으로 쓰이는 경우가 있어 통합시험 후보에 포함했다.

세 역할이 같은 요구사항 ID로 모두 잡혀야 해당 요구사항 QA 생성이 실행된다.

예:

```text
SFR-ESS-001 화면설계서
SFR-ESS-001 단위시험케이스
SFR-ESS-001 통합시험시나리오
```

### 5. QA 생성 결과 배치

`qa_generation/folder_pipeline.py`가 폴더 단위 QA 흐름을 담당한다.

처리 순서:

```text
1. 결과 폴더와 QA 원천 폴더 후보 스캔
2. 요구사항 ID 기준으로 화면설계서/단위시험케이스/통합시험시나리오 매칭
3. 화면설계서에서 단위시험 케이스 생성
4. 생성된 단위시험 XLSX로 통합시험 시나리오 생성
5. 생성된 HWPX/XLSX를 기존 대상 위치에 배치
6. 기존 대상 파일은 같은 폴더의 bak 하위로 이동
```

`bak` 폴더가 없으면 생성한다. 같은 이름이 이미 있으면 타임스탬프를 붙인다.

### 6. HWP 보안 팝업 완화

한글 COM 자동화에서 파일 접근 허용 팝업이 떠서 자동 처리가 막히는 문제가 있었다.

다음 파일에서 `FilePathCheckerModule` 등록 및 호출을 보강했다.

```text
document_update/hwpx_text.py
document_update/hwp_convert.py
```

현재 사용자 레지스트리에는 다음 DLL 경로를 등록해 두었다.

```text
C:\atoz_ai_integrated\venv310\Lib\site-packages\pyhwpx\FilePathCheckerModule.dll
```

다른 PC에서 실행할 경우 같은 레지스트리 등록이 필요할 수 있다.

## 주요 변경 파일

```text
web_app.py
web/check.html
web/static/check.js
web/static/check.css
web/qa.html
web/static/qa.js
web/static/qa.css
output_file_check/folder_apply_ops.py
qa_generation/folder_pipeline.py
qa_generation/generate_tc.py
qa_generation/generate_ts.py
document_update/hwpx_text.py
document_update/hwp_convert.py
CHECK_QA_PIPELINE_ISSUES.md
PROJECT_REVIEW_AND_WBS_POSTPROCESS.md
```

## 신규 API

### POST /api/run-qa-folder

`check.html` 결과 폴더 기준으로 QA 산출물을 생성하고 기존 파일 위치에 배치한다.

주요 필드:

```text
dump_root
qa_source_root
ui_design_root
tc_source_root
ts_source_root
qa_source_files
ui_design_files
```

실패 시에도 가능한 경우 검출 현황을 내려준다.

```json
{
  "ok": false,
  "role_counts": {
    "ui_design": 9,
    "tc_template": 0,
    "ts_template": 9
  },
  "source_files": [],
  "missing_requirements": []
}
```

## 현재 중요한 판단

### 파일 서칭은 AI가 아니다

현재 파일 역할 판정과 매칭은 규칙 기반이다. Ollama는 파일 탐색이 아니라 테스트케이스 내용 생성 단계에서 사용한다.

향후 정말 AI 검색을 붙이려면 다음 방식이 적절하다.

```text
1. 규칙 기반 후보를 먼저 만든다.
2. 애매한 파일만 파일명/상대경로/일부 추출 텍스트를 Ollama에 전달한다.
3. 모델이 role, requirement_id, confidence를 반환한다.
4. confidence가 낮으면 화면에서 수동 확인하게 한다.
```

처음부터 모든 파일을 Ollama에 넣는 방식은 느리고, HWP/HWPX/PDF 대량 텍스트 추출 비용도 크다.

### 단위시험케이스와 단위시험결과서는 별개

사용자가 확인한 업무 기준으로 `단위시험케이스`와 `단위시험결과서`는 둘 다 존재하는 별도 문서다.

따라서 현재 매칭은 `단위시험결과서`를 단위시험케이스로 취급하지 않는다.

## 검증한 내용

아래 검증을 수행했다.

```text
Python AST syntax check
Node --check web/static/qa.js
QA 원천 폴더 합성 매칭 테스트
단위시험결과서 제외/단위시험케이스 포함 테스트
브라우저에서 QA 원천 폴더 UI 반영 확인
서버 127.0.0.1:8001 재시작
```

합성 매칭 테스트에서는 같은 `SFR-ESS-001`을 가진 화면설계서, 단위시험케이스, 통합시험시나리오가 정상적으로 하나의 work item으로 묶였다.

## 남은 작업

### 1. 실제 샘플 전체 QA 실행 검증

사용자가 누락된 단위시험케이스 파일을 포함한 QA 원천 폴더를 다시 제공하면 실제 생성까지 실행해서 확인해야 한다.

확인 포인트:

- 화면설계서 9개
- 단위시험케이스 9개
- 통합시험 9개
- 요구사항 9개가 모두 매칭되는지
- 생성된 HWPX/XLSX가 기존 대상 위치로 들어가는지
- 기존 파일이 `bak`로 이동하는지

### 2. 단위시험케이스 HWPX 양식 구조 검증

`generate_tc.py`는 HWPX 안에서 아래 텍스트를 찾는다.

```text
단위시험 ID
단위시험 명
사전조건
화면 ID
수행 결과
```

양식 구조가 다르면 생성 실패할 수 있다.

### 3. WBS 후처리 연결

최종 목표는 `check + qa` 완료 폴더에 WBS 기반 작성자, 최초 작성 시작일자, 개정이력 0.1을 붙이는 것이다.

관련 메모:

```text
PROJECT_REVIEW_AND_WBS_POSTPROCESS.md
```

우선은 최종 폴더를 manifest로 만들고 review report/sidecar JSON을 붙이는 방식이 안전하다. 실제 문서 내부 쓰기는 HWPX/XLSX writer가 더 보강된 뒤 적용한다.

### 4. AI 기반 파일 역할 fallback

규칙 기반으로 잡히지 않는 파일이 남을 경우 Ollama를 이용한 파일 역할 분류 fallback을 추가한다.

권장 응답 스키마:

```json
{
  "role": "ui_design|tc_template|ts_template|ignore",
  "requirement_id": "SFR-ESS-001",
  "confidence": 0.0,
  "reason": "..."
}
```

## 참고 로그

서버 로그:

```text
C:\atoz_ai_integrated\web_runtime\web_app.log
```

QA 실패 시 우선 확인할 값:

```text
role_counts
source_files
missing_requirements
requirement_items
```

