# 🧐 Excel Command Platform (Backend)
[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.104+-green.svg)](https://fastapi.tiangolo.com/)
[![React](https://img.shields.io/badge/React-18.0+-blue.svg)](https://reactjs.org/)
[![MySQL](https://img.shields.io/badge/MySQL-8.0+-orange.svg)](https://www.mysql.com/)

LLM 기반 자연어 엑셀 명령어 분석 및 처리 서버입니다.\
사용자가 자연어로 입력한 명령을 분석하고, 실제 엑셀 파일에 수식 또는 스타일을 적용해주는 기능을 제공합니다.

> 📌 프론트엔드 저장소는 [여기](https://github.com/SKKU-2025-1-SWE3002-41-Team2/frontend)에서 확인할 수 있습니다.

---
## 목차
- [프로젝트 개요](#프로젝트-개요)
- [주요 기능](#주요-기능)
- [기술 스택](#기술-스택)
- [시스템 아키텍처](#시스템-아키텍처)
- [프로젝트 구조](#프로젝트-구조)
- [설치 및 실행](#설치-및-실행)
- [개발 가이드](#개발-가이드)
- [예시 명령](#예시-명령)
- [LLM 처리 방식](#llm-처리-방식)
- [팀원 소개](#팀원-소개)
- [라이선스](#라이선스)

---
## 프로젝트 개요

이 프로젝트는 자연어를 통해 엑셀 파일을 조작하는 시스템입니다.\
예를 들어, "A1에서 A10까지 1\~10을 넣고 평균을 구해줘" 와 같은 명령어를 입력하면, AI가 이를 해석하여 엑셀 명령어로 변환하고, 실제 시트 데이터를 수정합니다.
### 배경
- 기존 AI 도구들은 엑셀 함수만 알려주고, 사용자가 직접 입력해야 하는 번거로움 존재
- 엑셀 함수에 대한 기초 지식 없이는 활용이 어려움
- 복잡한 함수 조합이나 다중 작업 시 반복적인 설명과 적용 필요

### 목표
-  **자연어 명령**으로 엑셀 작업 수행
-  **AI가 자동으로** 엑셀 함수 생성 및 적용
-  **실시간 시각화**로 결과 즉시 확인
-  **원클릭 다운로드**로 완성된 엑셀 파일 저장

---

##  주요 기능

### 1. 자연어 명령 처리
```
사용자: "A1부터 A10까지 1에서 10까지 숫자를 넣고, B1에 평균을 구해줘"
AI: [작업 수행 후] "A1:A10에 1-10을 입력하고, B1에 평균값 5.5를 계산했습니다."
```

### 2. 지원하는 엑셀 기능
- **기본 함수**: SUM, AVERAGE, COUNT, MAX, MIN
- **논리 함수**: IF, AND, OR, IFS, IFERROR
- **텍스트 함수**: CONCATENATE, LEFT, RIGHT, MID, TRIM
- **검색 함수**: VLOOKUP, HLOOKUP, INDEX, MATCH
- **통계 함수**: MEDIAN, MODE, STDEV, RANK
- **조건부 함수**: COUNTIF, SUMIF, AVERAGEIF

### 3. 고급 기능
-  **파일 업로드/다운로드**: 기존 엑셀 파일 편집 가능
-  **대화 히스토리**: 이전 작업 내역 저장 및 불러오기
-  **반응형 UI**: 다양한 화면 크기 지원

---


## 기술 스택
### Backend
- **Backend Framework**: FastAPI
- **Language**: Python 3.11+
- **AI 모델**: OpenAI GPT (chat API 사용)
- **Excel 조작**: openpyxl
- **DBMS**: MySQL 8 (Docker 사용)
- **ORM**: SQLAlchemy
- **API 문서화**: Swagger UI (`/docs`)
- **테스트**: pytest

### DevOps
- **컨테이너**: Docker & Docker Compose
- **API 문서**: Swagger UI
- **테스트**: pytest (Backend)
---

## 시스템 아키텍처

```mermaid
graph TB
    subgraph "Frontend (React)"
        A[React App]
        B[Univer Editor<br/>- Excel 시트 렌더링<br/>- 실시간 편집]
        C[Chat Interface<br/>- 메시지 입력<br/>- 대화 히스토리]
        
        A --> B
        A --> C
    end
    
    subgraph "Backend (FastAPI)"
        D[FastAPI Server<br/>main.py]
        
        subgraph "Routers"
            R1[auth_router]
            R2[chat_router]
            R3[llm_router]
        end
        
        subgraph "Services"
            S1[auth_service<br/>- 로그인 처리]
            S2[chat_service<br/>- 세션 관리<br/>- 메시지 저장<br/>- 워크플로우 조정]
            S3[llm_service<br/>- GPT 호출<br/>- 명령어 변환]
            S4[excel_service<br/>- 엑셀 조작<br/>- 함수 적용]
        end
        
        D --> R1
        D --> R2
        D --> R3
        
        R1 --> S1
        R2 --> S2
        R3 --> S3
        
        S2 --> S3
        S2 --> S4
    end
    
    subgraph "Database"
        DB[(MySQL<br/>- User<br/>- ChatSession<br/>- Message<br/>- ChatSheet)]
    end
    
    subgraph "External API"
        GPT[OpenAI GPT-4o-mini<br/>자연어 → 엑셀 명령어]
    end
    
    %% Frontend-Backend 통신
    B <-->|"엑셀 데이터<br/>(Base64)"| R2
    C <-->|"채팅 메시지"| R2
    A <-->|"로그인"| R1
    
    %% Service 간 통신
    S3 <-->|"API 호출"| GPT
    S2 <-->|"DB 조작"| DB
    
    %% 데이터 흐름 표시
    style S2 fill:#f9f,stroke:#333,stroke-width:4px
    style S3 fill:#bbf,stroke:#333,stroke-width:2px
    style S4 fill:#bfb,stroke:#333,stroke-width:2px
```

'''
sequenceDiagram
    participant U as 사용자
    participant F as Frontend
    participant R as chat_router
    participant CS as chat_service
    participant LS as llm_service
    participant ES as excel_service
    participant DB as Database
    participant GPT as OpenAI GPT

    U->>F: "A1부터 A10까지 합계 구해줘"
    F->>R: POST /api/chat/sessions/{id}/message
    R->>CS: save_message_and_response()
    
    CS->>DB: 사용자 메시지 저장
    CS->>LS: get_llm_response()
    LS->>GPT: 자연어 분석 요청
    GPT-->>LS: 명령어 시퀀스 반환
    
    Note over LS: 응답 예시:<br/>commands: [<br/>  {type: "set_value", target: "A1:A10", params: [1,2,3...]},<br/>  {type: "sum", target: "A11", params: ["A1:A10"]}<br/>]
    
    LS-->>CS: ResponseResult 반환
    CS->>ES: process_excel_with_commands()
    ES->>ES: 엑셀 파일 수정
    ES-->>CS: 수정된 엑셀 bytes
    
    CS->>DB: AI 응답 저장
    CS->>DB: 엑셀 파일 저장
    CS->>DB: 세션 요약 업데이트
    
    CS-->>R: LLMMessageResponse
    R-->>F: 응답 (메시지 + 엑셀 데이터)
    F-->>U: UI 업데이트
'''

### 데이터 흐름
1. 사용자가 자연어 명령 입력
2. FastAPI 서버가 명령을 GPT에 전달
3. GPT가 엑셀 명령어 시퀀스로 변환
4. Excel Service가 실제 파일 수정
5. 결과를 Frontend에 전송 및 렌더링

---

## 프로젝트 구조

```
backend/
├── app/
│   ├── main.py              # FastAPI 엔트리포인트
│   ├── database.py          # DB 연결 설정
│   ├── models/              # SQLAlchemy 모델
│   │   ├── user.py
│   │   ├── chat_session.py
│   │   ├── message.py
│   │   └── chat_sheet.py
│   ├── schemas/             # Pydantic 스키마
│   │   ├── auth_schema.py
│   │   ├── chat_schema.py
│   │   └── excel_schema.py
│   ├── services/            # 비즈니스 로직
│   │   ├── auth_service.py
│   │   ├── chat_service.py
│   │   ├── llm_service.py
│   │   └── excel_service.py
│   └── routers/             # API 라우터
│       ├── auth_router.py
│       └── chat_router.py
├── tests/                   # 테스트 코드
├── requirements.txt
├── docker-compose.yml
└── README.md
```

---

##  설치 및 실행

### 사전 요구사항
- Python 3.11+
- Node.js 18+
- Docker & Docker Compose
- OpenAI API Key

### 1. 저장소 클론
```bash
git clone https://github.com/SKKU-2025-1-SWE3002-41-Team2/backend.git
cd backend
```

### 2. 환경 변수 설정
`.env` 파일 생성:
```env
DATABASE_URL=mysql+pymysql://excel:1234@localhost:3307/excel_platform
OPENAI_API_KEY=your-openai-api-key-here
```

### 3. Docker로 MySQL 실행
```bash
docker-compose up -d
```

### 4. Backend 실행
```bash
# 가상환경 생성 및 활성화
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 패키지 설치
pip install -r requirements.txt

# 서버 실행
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

---
##  개발 가이드

### 코드 스타일
- **Python**: PEP 8 준수
- **JavaScript**: ESLint + Prettier
- **명명 규칙**:
  - 클래스: PascalCase
  - 함수/변수: camelCase
  - 상수: UPPER_SNAKE_CASE

---
## 예시 명령

### 명령:

> A1부터 A10까지 1\~10 넣고, 평균을 B1에 표시해줘

### 실행되는 내부 명령 목록:

1. `set_value`: A1-A10에 1-10 삽입
2. `average`: B1에 `=AVERAGE(A1:A10)` 삽입

---

## LLM 처리 방식

- 사용자의 자연어 명령 → GPT API로 파시드
- 응답 JSON 내 `commands` 배열 파시드
- 각 명령어를 openpyxl 기반으로 엑셀 파일에 적용


---

## 👥 팀원 소개

| 이름  | 역할             | 담당 업무                              |
|-----|----------------|------------------------------------|
| 방기호 | 팀장, 프론트엔드, 백엔드 | 프로젝트 총괄, UI/UX 디자인, 엑셀 서비스         |
| 이상훈 | 프론트엔드, 백엔드     | React 개발, Univer 통합, 엑셀 서비스        |
| 윤규성 | 백엔드            | LLM 프롬포팅, 엑셀 서비스            |
| 정관용 | 백엔드            | LLM 서비스, LLM 프롬포팅, 엑셀 서비스 ,  문서 관리 |
| 박동민 | 백엔드            | DB 설계, API 설계, LLM 프롬포팅, 문서 관리     |

---

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.