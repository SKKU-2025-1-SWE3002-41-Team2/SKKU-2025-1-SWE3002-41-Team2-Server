# 🧐 Excel Command Platform (Backend)

LLM 기반 자연어 엑셀 명령어 분석 및 처리 서버입니다.\
사용자가 자연어로 입력한 명령을 분석하고, 실제 엑셀 파일에 수식 또는 스타일을 적용해주는 기능을 제공합니다.

> 📌 프론트엔드 저장소는 [여기](https://github.com/SKKU-2025-1-SWE3002-41-Team2/frontend)에서 확인할 수 있습니다.

---

## 프로젝트 개요

이 프로젝트는 자연어를 통해 엑셀 파일을 조작하는 시스템입니다.\
예를 들어, "A1에서 A10까지 1\~10을 넣고 평균을 구해줘" 와 같은 명령어를 입력하면, AI가 이를 해석하여 엑셀 명령어로 변환하고, 실제 시트 데이터를 수정합니다.

---

## 기술 스택

- **Backend Framework**: FastAPI
- **Language**: Python 3.11+
- **AI 모델**: OpenAI GPT (chat API 사용)
- **Excel 조작**: openpyxl
- **DBMS**: MySQL 8 (Docker 사용)
- **ORM**: SQLAlchemy
- **API 문서화**: Swagger UI (`/docs`)
- **테스트**: pytest

---

## 프로젝트 구조

```
SKKU-2025-1-SWE3002-41-Team2-Server/
├── app/
│   ├── api/                 # API 라우터
│   ├── services/            # Excel 조작, LLM 처리
│   ├── schemas/             # Pydantic 스키마
│   ├── models/              # SQLAlchemy 모델
│   ├── utils/               # 유틸리티 함수
│   └── main.py              # FastAPI 엔트리트포인트
├── tests/                   # 유닛 테스트
├── requirements.txt
├── Dockerfile
└── docker-compose.yml
```

---

## 실행 방법

### 1. MySQL (Docker) 실행

```bash
docker compose up -d
```
- 포트: 3307 → 내부 3306
- 계정 정보: `excel` / `1234`
- 데이터베이스: `excel_platform`

### 2. FastAPI 서버 실행

```bash
# 가상환경 설정 및 패키지 설치
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

### 3. .env 파일 설정

.env 파일에는 `DATABASE_URL`과 `OPENAI_API_KEY`가 필요합니다. 예시는 다음과 같습니다:

```
DATABASE_URL=mysql+pymysql://excel:1234@localhost:3307/excel_platform
OPENAI_API_KEY=[gpt_api_key]
```


# 서버 실행
uvicorn app.main:app --reload
```

---

## 예시 명령

### 명령:

> A1부터 A10까지 1\~10 넣고, 평균을 B1에 표시해줘

### 실행되는 내부 명령 목록:

1. `set_value`: A1~~A10에 1~~10 삽입
2. `average`: B1에 `=AVERAGE(A1:A10)` 삽입

---

## 주요 기능

- 자연어 기반 Excel 명령어 처리
- 수식 함수 지원 (SUM, AVERAGE, COUNT, IFS 등)
- 셀 서식 지정 (포트, 테두리, 배경색, 크기 등)
- 대화 세션 관리 (ChatSession + Message)
- 시트 데이터 저장 및 처리 (ChatSheet)

---

## LLM 처리 방식

- 사용자의 자연어 명령 → GPT API로 파시드
- 응답 JSON 내 `commands` 배열 파시드
- 각 명령어를 openpyxl 기반으로 엑셀 파일에 적용

---

## 프론트엔드

UniverJS 기반의 웹 엑셀 인터페이스를 사용하여 엑셀 시트를 렌더링합니다.\
해당 프로젝트는 [frontend](https://github.com/SKKU-2025-1-SWE3002-41-Team2/frontend) 폴더 또는 별도 저장소에서 확인 가능합니다.

---

## 라이선스

본 프로젝트는 MIT License 하에 배포됩니다.

