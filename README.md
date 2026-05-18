# AI Integrated QA Tool

산출물 매핑 확인 기능과 QA 산출물 생성 기능을 하나의 자체 HTTP 서버로 통합한 웹 도구입니다.

## 주요 기능

- 문서관리표준 PDF와 산출물 폴더를 비교해 파일 매핑 후보를 확인합니다.
- 매핑 결과를 기준으로 폴더 복사본에 문서값을 반영하고 ZIP으로 다운로드합니다.
- 사용자인터페이스설계서 PDF와 기존 단위시험 케이스 HWPX 양식을 기반으로 단위시험 케이스를 생성합니다.
- 기존 통합시험 시나리오 XLSX, 단위시험 케이스 XLSX, 사용자인터페이스설계서 PDF를 기반으로 통합시험 시나리오를 생성합니다.

## 실행 환경

- Windows
- Python 3.12 권장
- 한글 프로그램 설치 필요
- Ollama API 접근 가능 필요

HWPX 생성은 `pyhwpx`와 한글 COM 자동화를 사용하므로 Windows와 한글 설치 환경에서 실행해야 합니다.

## 설치

```powershell
cd C:\dev\workspace\ai_integrated
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

## 환경 변수

프로젝트 루트에 `.env` 파일을 만들고 Ollama 설정을 입력합니다.

```env
OLLAMA_BASE_URL=http://127.0.0.1:11434
OLLAMA_MODEL=exaone3.5:7.8b
```

`OLLAMA_BASE_URL`에는 `/api/chat` 또는 `/api/generate`를 붙이지 않는 것을 권장합니다. 서버가 필요한 엔드포인트를 내부에서 조합합니다.

## 실행

```powershell
.\venv\Scripts\activate
python web_app.py --host 0.0.0.0 --port 8000
```

브라우저에서 아래 주소로 접속합니다.

```text
http://127.0.0.1:8000
```

주요 화면은 다음과 같습니다.

- `/check.html`: 산출물 매핑 확인
- `/qa.html`: QA 생성

## 사용 흐름

### 산출물 매핑

1. 문서관리표준 PDF를 업로드합니다.
2. 검사할 산출물 폴더를 업로드합니다.
3. 매핑 후보를 확인합니다.
4. 필요한 경우 폴더 반영을 실행하고 ZIP 파일을 다운로드합니다.

### 단위시험 케이스 생성

1. 기존 단위시험 케이스 HWPX 파일을 업로드합니다.
2. 사용자인터페이스설계서 PDF를 업로드합니다.
3. 생성 버튼을 누릅니다.
4. 생성된 XLSX/HWPX 파일을 다운로드합니다.

### 통합시험 시나리오 생성

1. 기존 통합시험 시나리오 XLSX 파일을 업로드합니다.
2. 1단계에서 생성한 단위시험 케이스 XLSX 파일을 업로드합니다.
3. 사용자인터페이스설계서 PDF를 업로드합니다.
4. 생성 버튼을 누릅니다.
5. 생성된 XLSX 파일을 다운로드합니다.

## 런타임 파일

서버 실행 중 생성되는 임시 파일과 결과 파일은 `web_runtime` 아래에 만들어집니다.

- `web_runtime/temp`: 업로드 파일과 QA 생성 임시 결과
- `web_runtime/results`: 산출물 반영 ZIP 등 다운로드 결과
- `web_runtime/web_app.log`: 서버 로그

QA 생성 결과는 요청별 임시 폴더에 생성되고 다운로드 후 정리됩니다. 서버 시작 시 중단된 요청의 임시 폴더도 함께 정리됩니다.

## 프로젝트 구조

```text
ai_integrated/
  app_runtime.py              # 런타임 경로, 로그, 환경 설정
  web_app.py                  # 자체 HTTP 서버
  web_uploads.py              # 업로드 처리
  output_file_check/          # 산출물 매핑 확인/반영
  qa_generation/              # TC/TS 생성 로직
  templates/                  # 기본 양식 파일
  web/                        # HTML, CSS, JS
  web_runtime/                # 실행 중 생성되는 파일
```

## 주의 사항

- `.env`, `web_runtime`, `venv`는 Git에 올리지 않는 것을 권장합니다.
- HWPX 생성 중 한글 COM 자동화 오류가 발생하면 실행 중인 한글 프로세스를 종료한 뒤 다시 시도하세요.
- Ollama 서버와 모델이 준비되어 있지 않으면 QA 생성 기능이 실패할 수 있습니다.
