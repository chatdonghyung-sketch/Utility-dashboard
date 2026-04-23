# CHANGELOG

모든 버전별 변경사항을 이 파일에 기록합니다.  
형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.0.0/)를 따르며,  
버전은 [Semantic Versioning](https://semver.org/lang/ko/) 규칙을 따릅니다.

---

## [1.0.0] - 2026-04-14

### 최초 릴리즈

#### 추가
- FastAPI + SQLite 기반 백엔드 서버 (`backend/app.py`, `backend/database.py`)
- 단일 실행 런처 스크립트 (`start.py`) — 패키지 자동 설치 후 uvicorn 기동
- 냉동기·냉각탑 운전일지 API (`/api/chiller-log`)
- 냉동기 가동 분석 API (`/api/chiller-analysis`)
- FFU 운전일지 API (`/api/ffu-log`)
- 보일러 운전일지 API (`/api/boiler-log`)
- 안전 점검일지 API (`/api/safety-log`)
- HMI 트렌드 조회 API (`/api/trend/*`) — 일별·범위 조회 및 FAB/Process 필터
- 작업 지시서(Work Order) 업로드·조회·KPI API (`/api/work-orders/*`)
- 에너지 사용량(LNG·STEAM·NItrogen·ARGON) 엑셀 자동 스캔 API (`/api/energy-files`)
- Vue.js 빌드 결과물 정적 서빙 (`dist/`)
- 에너지 데이터 초기 샘플 (1·2·3공장 × 4개 유틸리티, 2026-04-14 기준)
- HMI 트렌드 샘플 데이터 (`HMI_Trend_data/TREND_20260331.xls`)
- 작업 지시서 샘플 (`Work_Order/작업보고서_20260414.xlsx`)

---

<!-- 새 버전 추가 시 위에 아래 형식으로 추가 -->
<!--
## [X.Y.Z] - YYYY-MM-DD

### 추가
-

### 변경
-

### 수정
-

### 제거
-
-->
