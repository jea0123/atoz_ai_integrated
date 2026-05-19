# ATOZ AI Integrated 프로젝트 정리 및 WBS 후처리 검토

작성일: 2026-05-18

## 0. 추가 요구사항 반영 요약

추가 확인 결과, 이 프로젝트의 최종 흐름은 `check.html`과 `qa.html`을 따로 운영하는 구조가 아니라 하나의 산출물 완성 파이프라인으로 묶는 방향이 맞다.

목표 흐름은 다음과 같다.

```text
1. check.html
   문서관리표준 PDF + 기존 산출물 폴더를 기준으로 기존 산출물의 표지, 문서번호, 문서명, 프로젝트명, 파일명을 정규화한다.

2. qa.html
   check.html이 만든 정규화 산출물 폴더를 기반으로 단위시험 케이스와 통합시험 시나리오를 생성한다.
   이때 사용자는 테스트 문서 양식과 내용을 가져올 화면설계서/사용자인터페이스설계서 타입만 제공한다.

3. QA 결과 배치
   qa.html 생성 결과물은 별도 다운로드 파일로 끝나지 않고 check.html 결과 폴더의 해당 산출물 경로에 들어가야 한다.
   기존에 그 위치에 있던 파일은 같은 폴더의 bak 폴더로 이동한다.
   bak 폴더가 없으면 생성한다.

4. 중간 결과 폴더 유지
   check.html의 덤프 후 반영 결과는 ZIP으로 끝내지 않고 복사폴더 상태로 유지한다.
   이 복사폴더가 이후 QA 생성, QA 결과 교체, WBS 후처리의 작업 기준 폴더가 된다.

5. WBS 후처리
   check + qa가 모두 끝난 최종 산출물 폴더를 WBS와 매칭한다.
   산출물별 작성자/담당자와 최초 작성 시작일자를 실제 문서에 반영한다.

6. 개정이력 보강
   가능하면 각 문서의 개정이력에 0.1 버전 최초 작성 이력을 추가한다.

7. 최종 전달
   ZIP이 꼭 필요하다면 check + qa + WBS + 개정이력까지 모두 끝난 최종 폴더를 기준으로 마지막에만 만든다.
   기본 결과는 최종 복사폴더 경로를 유지하는 방식이 더 자연스럽다.
```

따라서 WBS 후처리는 더 이상 `check.html` 단독 결과 뒤에 붙는 작업이 아니라, `check.html + qa.html` 통합 결과 뒤에 붙는 최종 메타데이터 보강 단계로 보는 것이 정확하다.

## 1. 현재 프로젝트 위치와 실행 상태

- 프로젝트 경로: `C:\atoz_ai_integrated`
- 저장소: `https://github.com/jea0123/atoz_ai_integrated`
- 실행 방식: Python 자체 HTTP 서버
- 현재 기동 URL: `http://127.0.0.1:8001`
- 주요 화면:
  - `/check.html`: 산출물 매핑 확인 및 폴더 반영
  - `/qa.html`: 단위시험 케이스, 통합시험 시나리오 생성
- 가상환경: `C:\atoz_ai_integrated\venv310`
- Ollama 설정 파일: `C:\atoz_ai_integrated\.env`

현재 `.env`는 다음 설정을 사용한다.

```env
OLLAMA_BASE_URL="http://192.168.1.7:11434"
OLLAMA_MODEL="exaone3.5:7.8b"
```

서버 런타임 API 기준으로 Ollama 설정은 `AI 우선` 모드로 인식된다.

## 2. 프로젝트 목적

이 프로젝트는 문서관리표준 PDF, 기존 산출물 폴더, UI 설계서 PDF, QA 양식을 입력받아 산출물 정리와 QA 문서 생성을 돕는 Windows 기반 웹 도구다.

크게 두 계열의 기능이 있다.

1. 산출물 매핑 및 반영
   - 문서관리표준 PDF에서 산출물 번호, 산출물명, 예상 폴더 경로, 프로젝트명을 읽는다.
   - 검사 대상 산출물 폴더를 스캔한다.
   - 표준 산출물과 실제 파일을 AI 우선 또는 규칙 기반으로 매칭한다.
   - 원본을 직접 수정하지 않고 덤프 폴더에 복사한 뒤 복사본에 문서번호, 문서명, 프로젝트명, 파일명을 반영한다.
   - 현재 구현은 반영 결과 폴더를 ZIP으로 묶어 다운로드하지만, 통합 목표에서는 이 복사폴더를 유지해야 한다.

2. QA 산출물 생성
   - UI 설계서 PDF와 단위시험 케이스 HWPX 양식을 기반으로 단위시험 케이스 XLSX/HWPX를 생성한다.
   - 기존 통합시험 시나리오 XLSX, 단위시험 케이스 XLSX, UI 설계서 PDF를 기반으로 통합시험 시나리오 XLSX를 생성한다.

## 3. 주요 코드 구조

```text
C:\atoz_ai_integrated
  app_runtime.py
  web_app.py
  web_uploads.py
  folder_apply.py
  requirements.txt
  output_file_check\
  document_update\
  qa_generation\
  templates\
  web\
  web_runtime\
```

주요 역할은 다음과 같다.

- `web_app.py`
  - HTTP 서버 진입점.
  - `/api/check`, `/api/folder-apply`, `/api/generate-tc`, `/api/generate-ts`, `/api/runtime-mode`, `/download/{token}`을 처리한다.
  - 현재 폴더 반영 결과는 ZIP 다운로드 토큰으로 연결한다.
  - 통합 목표에서는 `/api/folder-apply`가 최종 ZIP을 만들기보다 복사폴더 경로와 후속 작업 상태를 유지해야 한다.

- `app_runtime.py`
  - 프로젝트 기준 경로, `.env`, 로그, Ollama URL, 런타임 모드 판단을 담당한다.
  - `OLLAMA_BASE_URL`을 읽어 `/api/generate`, `/api/chat` 엔드포인트를 내부에서 조합한다.

- `output_file_check/`
  - 문서관리표준 PDF 파싱, 검사 폴더 스캔, 폴더 정책, 매칭, 결과 직렬화, 덤프 반영을 담당한다.
  - `folder_workflow.py`가 웹/CLI 공통 흐름을 조립한다.
  - `folder_apply_ops.py`가 원본 폴더 복사, 문서값 반영, 파일명 변경, `apply_items` 생성까지 맡는다.

- `document_update/`
  - HWP/HWPX/XLSX 등 문서 내부의 문서번호, 문서명, 프로젝트명 치환을 담당한다.
  - HWP 처리는 한글 프로그램 COM 자동화와 `pywin32` 환경에 영향을 받는다.

- `qa_generation/`
  - `generate_tc.py`: UI 설계서 PDF에서 화면/처리흐름을 추출하고 Ollama로 단위시험 케이스를 만든다.
  - `generate_ts.py`: 단위시험 케이스 XLSX와 UI 설계서 PDF의 요구사항 ID 매핑을 이용해 통합시험 시나리오를 만든다.

- `web_runtime/`
  - 서버 실행 중 업로드 파일, 임시 결과, 덤프 결과, 다운로드 ZIP, 로그가 생성되는 작업 영역이다.

## 4. 산출물 매핑 및 반영 흐름

산출물 매핑 확인 흐름은 다음과 같다.

```text
브라우저 업로드
-> web_uploads.save_check_uploads
-> output_file_check.folder_workflow.run_web_check
-> output_file_check.folder_mapping.build_folder_mapping
-> output_file_check.folder_serialization.serialize_check_result
-> JSON 응답
```

폴더 반영 흐름은 다음과 같다.

```text
브라우저 업로드
-> run_web_folder_apply
-> copy_folder_to_dump
-> apply_dumped_folder
-> apply_batch_candidate
-> document_update.document_number.write_updated_document
-> apply_items 포함 JSON 응답
-> attach_folder_download
-> ZIP 다운로드
```

위 흐름은 현재 구현 기준이다. 통합 목표 기준으로는 `attach_folder_download`와 `dump_root` 삭제가 중간 단계에서 실행되면 안 된다. `dump_root`는 QA 결과 교체와 WBS 후처리가 끝날 때까지 살아 있어야 한다.

중요한 점은 원본 폴더를 직접 수정하지 않고 `dump_root` 아래 복사본만 수정한다는 것이다. 이 때문에 WBS 후처리를 붙일 때도 원본 훼손 위험이 낮다.

폴더 반영 완료 payload에는 후속 작업에 필요한 핵심 정보가 이미 들어간다.

```text
dump_root
updated_file_count
failed_file_count
apply_target_file_count
apply_items[]
  status
  output_id
  output_name
  old_path
  new_path
  backup_path
  converted_to_hwpx
  error
```

## 5. 산출물 생성 위치

서버 실행 중 결과는 주로 아래 위치에 생긴다.

- 폴더 반영 덤프: `C:\atoz_ai_integrated\web_runtime\results\folder-dumps\{request_id}\...`
- 폴더 반영 ZIP: `C:\atoz_ai_integrated\web_runtime\results\...zip`
- QA 생성 임시 결과: `C:\atoz_ai_integrated\web_runtime\temp\qa-tc-{request_id}\output`
- 서버 로그: `C:\atoz_ai_integrated\web_runtime\web_app.log`

QA 생성 결과는 현재 다운로드 후 정리될 수 있다. 폴더 반영 덤프도 현재는 사용자가 `dump_path`를 지정하지 않으면 ZIP 생성 후 내부 정리 흐름이 들어간다. 통합 목표에서는 이 동작을 바꿔서 `dump_root`를 유지하고, QA/WBS/개정이력까지 모두 끝난 뒤 최종 폴더를 사용자에게 알려주는 방식이 맞다. ZIP은 최종 전달 옵션으로만 두는 것이 좋다.

## 6. NEXT_THREAD_HANDOFF.md 핵심 내용

`C:\testdocWriter2\NEXT_THREAD_HANDOFF.md`의 핵심은 이 프로젝트가 만든 치환 결과물을 WBS 후처리 입력으로 사용하자는 것이다.

후처리 목표는 다음과 같다.

- 현정씨 프로그램의 폴더 반영 결과물을 입력으로 받는다.
- WBS와 산출물을 매칭한다.
- 각 산출물에 담당자와 산출물별 작성시작일을 붙인다.
- 자동 확정이 어려운 항목은 사유 보고서에 남긴다.
- 은애씨 담당 문서와 프로그램 소스 등은 우선 제외한다.

이미 `C:\testdocWriter2`에는 이 목적을 위한 스크립트가 있다.

```text
C:\testdocWriter2\scripts\extract_wbs_contract_from_xlsm.ps1
C:\testdocWriter2\scripts\export_hyunjung_dump_manifest.ps1
C:\testdocWriter2\scripts\build_wbs_artifact_match_plan.ps1
C:\testdocWriter2\scripts\apply_wbs_metadata.ps1
```

## 7. 산출물 뒤에 WBS 후처리를 덧댈 수 있는지

결론부터 말하면 가능하다. 특히 이 프로젝트의 `폴더 반영` 산출물에는 바로 후속 작업을 붙이기 좋다.

가능한 이유는 다음과 같다.

- 폴더 반영 결과가 `dump_root`라는 명확한 결과 폴더로 분리된다.
- 각 반영 파일에 대해 `apply_items`가 남아서 산출물 ID, 산출물명, 원본 경로, 변경 경로, 오류 사유를 추적할 수 있다.
- `C:\testdocWriter2\scripts\export_hyunjung_dump_manifest.ps1`는 덤프 폴더를 스캔해 WBS 후처리용 `generated_artifact_manifest`를 만들 수 있다.
- `build_wbs_artifact_match_plan.ps1`는 manifest와 WBS contract를 비교해 `owner`, `startDate`, `matchedWbsId`, `confidence`, `manualReview`를 산출한다.
- `apply_wbs_metadata.ps1`는 매칭 계획을 바탕으로 미리보기 폴더와 sidecar JSON, 검토 보고서를 만들 수 있다.

다만 “후처리 보고서를 최종 폴더에 덧붙인다”와 “문서 내부의 담당자/작성시작일 필드에 정확히 쓴다”는 난이도가 다르다.

- 후속 작업 메모나 검토 보고서를 최종 복사폴더에 함께 넣는 것은 바로 가능하다.
- WBS manifest, match plan, review report를 최종 복사폴더 옆이나 내부에 생성하는 것도 바로 가능하다.
- 실제 HWPX/XLSX 문서 표지 안의 담당자/작성시작일 위치에 값을 넣는 것은 writer 보강이 필요하다.

## 8. 붙일 수 있는 지점

가장 자연스러운 연결 지점은 `web_app.py`의 `handle_folder_apply_post`다.

현재 흐름은 다음 순서다.

```text
payload = run_web_folder_apply(...)
attach_folder_download(payload)
필요 시 dump_root 정리
self.send_json(payload)
```

현재 목표에 맞게 바꾸려면 `attach_folder_download(payload)`를 중간 단계에서 빼고, `dump_root` 정리도 하지 않아야 한다. 그 다음 같은 `dump_root`에서 아래 단계를 이어간다.

```text
1. payload["dump_root"]에서 generated_artifact_manifest 생성
2. 최신 WBS contract와 manifest로 match plan 생성
3. review.md 생성
4. 자동 적용 가능 항목은 preview 또는 실제 후처리 결과 폴더 생성
5. 최종 복사폴더 경로와 보고서 경로를 응답에 포함
6. 필요할 때만 최종 단계에서 ZIP 생성
```

초기 통합은 Python에서 PowerShell 스크립트를 호출하는 방식이 가장 빠르다. 안정화 후에는 manifest 생성과 WBS 매칭 로직을 Python 모듈로 옮기는 편이 배포와 유지보수에 유리하다.

## 9. 현재 제약과 위험

1. WBS 후보 모호성
   - `데이터전환결과서`처럼 같은 산출물명이 WBS에 여러 번 나오고 파일명에 요구사항 ID가 없으면 자동 적용이 막힌다.
   - `NEXT_THREAD_HANDOFF.md` 기준으로 `3.5.4.1 / 통합시험 I / SFR-IIN-002 / 데이터전환결과서` 행은 WBS 수정 여부 확인이 필요하다.

2. 문서 내부 writer 미완성
   - `apply_wbs_metadata.ps1`는 현재 마커 기반 치환이다.
   - 실제 파일에 `{{담당자}}`, `{{작성시작일}}` 같은 마커가 없으면 `copied_no_marker`가 난다.
   - 실사용을 위해서는 HWPX/XLSX 표지 셀 또는 문단 위치 기반 writer/profile이 필요하다.

3. HWP 처리 이슈
   - 기존 인계 메모 기준으로 `.hwp` 4건이 pywin32와 한글 COM 자동화 문제로 실패했다.
   - 이 프로젝트도 HWP 변환은 Windows 한글 설치와 COM 환경에 의존한다.

4. QA 산출물과 WBS 후처리의 관계
   - 기존 인계 메모에서는 단위시험 케이스, 통합시험 시나리오, 인수인계시험 관련 문서를 은애씨 프로그램 담당 산출물로 보고 WBS 후처리 대상에서 우선 제외했다.
   - 추가 요구사항 기준으로는 이 제외 정책을 그대로 두면 안 된다.
   - `qa.html`이 생성한 테스트 산출물은 `check.html` 결과 폴더의 정식 산출물 경로에 배치되어야 하며, 그 이후 최종 산출물 전체가 WBS 후처리 대상이 된다.
   - 다만 WBS 자동 매칭에서는 테스트 산출물명이 여러 요구사항에 반복될 수 있으므로 `요구사항 ID`, `화면 ID`, 산출물명, 기존 경로를 함께 써야 한다.

5. 결과 폴더 생명주기
   - 현재 웹 요청에서 `dump_path`가 없으면 서버가 덤프 폴더를 정리한다.
   - 통합 목표에서는 이 동작이 맞지 않는다.
   - 확인하기에서 감지한 폴더 구조를 그대로 복사한 결과 폴더를 유지하고, 그 복사폴더를 QA/WBS 후속 작업의 기준으로 써야 한다.
   - ZIP은 최종 폴더가 완성된 뒤 선택적으로 생성해야 한다.

## 10. 권장 통합안

### 1단계: check + qa 통합 파이프라인 추가

새 API를 하나 두는 것이 가장 자연스럽다.

```text
POST /api/integrated-artifact-build
```

이 API는 내부에서 다음 순서로 실행한다.

```text
1. run_web_folder_apply
   기존 check.html의 덤프 후 반영을 실행한다.

2. QA 템플릿 탐색
   dump_root 안에서 단위시험 케이스 HWPX, 통합시험 시나리오 XLSX 같은 테스트 산출물 위치를 찾는다.

3. generate_test_cases
   사용자인터페이스설계서 PDF를 읽어 단위시험 케이스 XLSX/HWPX를 생성한다.

4. replace_artifact_with_backup
   생성된 단위시험 케이스 결과물을 dump_root 안의 기존 단위시험 케이스 위치에 배치한다.
   기존 파일은 같은 폴더의 bak 폴더로 이동한다.

5. generate_test_scenarios
   생성된 단위시험 케이스 XLSX와 기존 통합시험 시나리오 XLSX 양식을 이용해 통합시험 시나리오 XLSX를 만든다.

6. replace_artifact_with_backup
   생성된 통합시험 시나리오 결과물을 dump_root 안의 기존 통합시험 시나리오 위치에 배치한다.
   기존 파일은 같은 폴더의 bak 폴더로 이동한다.

7. WBS 후처리
   최종 dump_root를 manifest로 만들고 WBS match plan을 생성한다.
   자동 적용 가능한 항목은 작성자/최초 작성 시작일자를 넣고, 애매한 항목은 review report에 남긴다.

8. 개정이력 0.1 추가
   작성자/최초 작성 시작일자와 함께 개정이력 표에 0.1 최초 작성 행을 넣는다.

9. 최종 결과 유지
   check + qa + WBS 메타데이터까지 반영된 최종 복사폴더를 유지하고 사용자에게 경로를 제공한다.
   ZIP이 필요하면 이 최종 복사폴더를 기준으로 마지막에만 선택적으로 만든다.
```

기존 `check.html`과 `qa.html`을 그대로 유지하더라도, 운영 화면은 `check.html`에 QA 입력 섹션을 추가하는 방식이 좋다. 이유는 최종 배치 기준 경로가 `check.html`의 `dump_root`이기 때문이다.

### 2단계: QA 결과를 기존 산출물 위치로 교체 배치

새 공통 함수가 필요하다.

```text
replace_artifact_with_backup(target_path, generated_path)
```

정책은 다음과 같다.

- `target_path`가 있으면 `target_path.parent\bak` 폴더로 이동한다.
- `bak` 폴더가 없으면 만든다.
- bak 안에 같은 이름이 있으면 타임스탬프를 붙여 충돌을 피한다.
- `generated_path`는 `target_path` 위치로 이동하거나 복사한다.
- 생성 파일명이 임의 이름이어도 최종 파일명은 기존 target 파일명을 따른다.
- 교체 결과는 report에 `replaced`, `backup_path`, `target_path`로 남긴다.

이 단계가 있어야 `qa.html` 결과물이 별도 다운로드로 흩어지지 않고, 문서관리표준 기준의 산출물 폴더 구조 안에 정확히 들어간다.

### 3단계: 외부 WBS 후처리로 연결

현재 구조를 크게 바꾸지 않고 아래 순서로 운영한다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\testdocWriter2\scripts\export_hyunjung_dump_manifest.ps1" `
  -DumpRoot "{folder_apply_dump_root}" `
  -OutputPath "C:\testdocWriter2\samples\generated_artifact_manifest.from-atoz-ai-integrated.json" `
  -ProjectName "2026년도 수입식품통합정보시스템 고도화" `
  -SystemName "수입식품통합정보시스템" `
  -ExcludeFolderNames "양식,자체테스트(사업관리)" `
  -ExcludeDocumentNames "통합시험결과서,인수인계시험결과서,단위시험케이스,통합시험시나리오,인수인계시험시나리오"
```

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\testdocWriter2\scripts\build_wbs_artifact_match_plan.ps1" `
  -ManifestPath "C:\testdocWriter2\samples\generated_artifact_manifest.from-atoz-ai-integrated.json" `
  -WbsContractPath "C:\testdocWriter2\samples\wbs_contract.from-current-wbs.json" `
  -OutputPath "C:\testdocWriter2\outputs\wbs_artifact_match_plan.from-atoz-ai-integrated.json"
```

초기에는 자동 문서 수정까지 한 번에 욕심내기보다 `.review.md`와 sidecar JSON을 최종 산출물 폴더에 붙이는 방식이 안전하다. 이후 writer가 안정화되면 작성자/최초 작성 시작일자/개정이력을 실제 문서 내부에 적용한다.

### 4단계: 웹 앱 내부 후처리 옵션 추가

`/check.html` 화면의 폴더 반영 버튼 근처에 다음 옵션을 추가한다.

- WBS 후처리 실행
- WBS 파일 경로 또는 업로드
- 은애씨 담당 문서 제외
- 수동 검토 항목 포함 여부
- 후처리 결과 보고서 다운로드

서버는 `/api/folder-apply` 응답에 다음 필드를 추가할 수 있다.

```json
{
  "wbs_postprocess": {
    "manifest_path": "...json",
    "match_plan_path": "...json",
    "review_report_path": "...review.md",
    "preview_dir": "...",
    "auto_apply_count": 0,
    "manual_review_count": 1,
    "download_url": "/download/..."
  }
}
```

### 5단계: 실제 문서 writer 보강

마지막 단계에서 파일 유형별 writer를 만든다.

- XLSX/XLSM: openpyxl로 표지 시트의 라벨 기반 셀 탐색 후 담당자/작성시작일 입력
- HWPX: 압축 XML 내부에서 표지 문단/표 구조 탐색 후 입력
- DOCX: OOXML 문단/표 라벨 기반 입력
- HWP: 가능하면 HWPX 변환 후 처리하거나 한글 COM writer로 분리

이 단계에서 개정이력 writer도 같이 들어가야 한다.

개정이력 0.1 행의 기본 정책은 다음처럼 잡을 수 있다.

```text
버전: 0.1
일자: WBS 최초 작성 시작일자
작성자: WBS 담당자
개정내용: 최초 작성
```

이 단계가 끝나야 “담당자/작성시작일/개정이력이 실제 문서 내부에 들어간 최종 산출물”이라고 볼 수 있다.

## 11. 판단 요약

- `NEXT_THREAD_HANDOFF.md`의 후순위 작업은 이 프로젝트 산출물 뒤에 붙일 수 있다.
- 최종 기준은 `/check.html` 단독 결과가 아니라 `check.html + qa.html`을 모두 돌린 통합 결과 폴더다.
- `qa.html` 결과물은 별도 다운로드로 끝내지 말고 `check.html` 결과 폴더의 동일 산출물 경로에 교체 배치해야 한다.
- 교체 전 기존 파일은 같은 폴더의 `bak` 폴더로 이동해야 하며, `bak`이 없으면 생성한다.
- 그 다음 최종 폴더를 manifest로 바꿔 WBS match plan과 review report를 만들 수 있다.
- 자동 적용까지 가려면 WBS 모호성 해소, 파일 유형별 writer, 작성자/최초 작성 시작일자 위치 탐색, 개정이력 0.1 행 삽입 로직이 필요하다.
- ZIP은 중간 결과로 만들 필요가 없고, 최종 복사폴더가 완성된 뒤 선택적으로만 만들면 된다.
- 실무적으로는 먼저 “복사폴더 유지 + QA 결과 교체 배치 + WBS 검토 보고서/sidecar JSON 포함”을 완성하고, 이후 실제 문서 내부 쓰기와 개정이력 삽입을 단계적으로 완성하는 것이 가장 안전하다.
