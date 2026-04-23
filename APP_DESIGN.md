# Attainment Report Generator - UI Design Notes

## 1) 목적
이 문서는 현재 Power Apps 코드앱(`attainment-report-powerapp`)의 UI 디자인 기준과 작성 방향을 정리합니다.
목표는 Quota Validator와 유사한 엔터프라이즈 톤을 유지하면서, 보고서 생성/이메일 Draft 작업의 가독성과 작업 속도를 높이는 것입니다.

## 2) 참고 기준
- Quota Validator의 Tailwind 기본 토큰
  - `c:\Codex\PowerApps\Quota Validator\node_modules\tailwindcss\index.css`
  - `c:\Codex\PowerApps\Quota Validator\node_modules\tailwindcss\theme.css`
- 특히 다음 개념을 반영:
  - Neutral + Blue 중심의 안정적인 색상 체계
  - 단순한 card/surface 계층
  - 과도한 장식 대신 선명한 상태 표현(활성/비활성/성공/오류)

## 3) 디자인 시스템 요약
- Typography
  - 기본 폰트: Segoe UI 계열
  - 전역 폰트 크기: `src/index.css`의 `body { font-size }`를 기준으로 전체 스케일 제어
  - 타이틀/섹션/버튼/힌트/메시지가 비례적으로 확장/축소되도록 구성

- Color
  - 배경: 연한 gray (`#edeff3`)
  - 카드: white surface + 얇은 border
  - Primary 액션: blue (`#2563eb` 계열)
  - 상태 메시지:
    - success: green tint
    - error: red tint

- Radius/Border/Spacing
  - radius: 8~12px 중심
  - border: gray-200/300 톤
  - panel 내부 여백과 컴포넌트 간 간격은 10~16px 단위

## 4) 주요 UI 구성
- 상단 헤더
  - 앱 이름 + 보조 설명
- Step 1: 파일 업로드
  - 버튼 업로드 + Drag & Drop 동시 지원
  - 드래그 중 하이라이트(`.drag-active`) 표시
- Step 2: 리포트 생성
  - region 선택 pill + 진행률 표시 + 통계 카드
- Step 3: Outlook Draft
  - region filter / manager list / template editor 3열 구성

## 5) 인터랙션 가이드
- 업로드
  - `.xlsx` 외 확장자는 즉시 에러 노출
  - 처리 중(`isGenerating`, `isDrafting`)에는 업로드 비활성화
- 버튼 상태
  - 가능 조건이 충족될 때만 활성화
  - disabled 시 명확한 투명도 처리
- 진행 상태
  - progress text + progress bar 동시 제공

## 5-1) 리포트 생성 로직
- Attainment row의 `LI_EMP_ID`를 Sales Compensation Report의 `Employee ID`와 매칭합니다.
- SCR에 `Supervisory Manager`가 있으면 리포트 생성 시 `Level_1_Manager` 대신 SCR 기준 매니저를 사용합니다.
- 생성될 리포트 row가 모두 `Employee Status = Terminated`인 매니저는 리포트/Draft 대상에서 제외합니다.
- 각 매니저 리포트의 row는 `LI_EMP_ID` 오름차순, `Quota Start Date` 오름차순, `Measure Weight` 내림차순으로 정렬합니다.

## 6) 접근성/운영 포인트
- 텍스트 대비를 유지(특히 힌트/라벨 영역)
- 모바일에서는 2열/3열 레이아웃을 1열로 전환
- 폰트 변경 요청은 우선 `src/index.css`의 body font-size 조정 후 필요 시 `src/App.css` 개별 요소 보정

## 7) 파일 위치
- 전역 스타일: `attainment-report-powerapp/src/index.css`
- 화면 스타일: `attainment-report-powerapp/src/App.css`
- 화면 구조: `attainment-report-powerapp/src/App.tsx`

## 8) 변경 이력
- 2026-04-23
  - Sales Compensation Report의 `Supervisory Manager`를 기준으로 현재 매니저 조직도를 보정
  - 퇴사자만 포함된 매니저 리포트 생성 제외
  - 매니저별 리포트 row 정렬 기준 추가: `LI_EMP_ID`, `Quota Start Date`, `Measure Weight`
- 2026-02-27
  - 루트 폴더에서 레거시 Python/Streamlit 기반 파일 및 미사용 산출물 정리
  - React Power App(`attainment-report-powerapp`) 중심 구조로 재정렬
  - 저장소 반영 대상: `https://github.com/LI-JBLEE/PA-Attainment-Report-Automator`
