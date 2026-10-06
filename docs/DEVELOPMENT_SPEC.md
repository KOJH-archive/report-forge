# Report Forge 개발명세서

## 1. 문서 목적

Report Forge는 원고, 데이터, 출처와 선택한 디자인 템플릿을 바탕으로 편집 가능한 PowerPoint, Word, Excel 결과물을 생성하는 로컬 우선 제작기다. 이 문서는 현재 구현 범위와 추가 개발 방향, 사용자가 조정할 수 있는 설정을 정의한다.

## 2. 제품 목표

- 회사 PPTX의 디자인 언어를 보존하면서 새 내용과 데이터를 교체한다.
- 회사 서식이 없을 때 미리 정의한 보고서 패턴에서 적절한 구성을 선택한다.
- 같은 원고와 데이터로 PPT, Word, Excel의 근거를 일관되게 생성한다.
- 데이터·출처·계산식을 결과물과 함께 추적할 수 있게 한다.
- 외부 API 없이 로컬에서 생성하며 Codex, Antigravity, Cursor, VS Code와 일반 터미널에서 실행한다.

## 3. 현재 아키텍처

```text
report.json + CSV/XLSX + optional PPTX
              ↓
schema validation / safe input checks
              ↓
data metrics + narrative interpolation
              ↓
slide planning / template slot mapping
              ↓
PPTX + DOCX + XLSX + HTML preview + manifest
```

주요 모듈:

- `src/core.mjs`: 스키마 검증, 입력 파일 보호, 데이터 로딩, 지표 계산, 원고 해석, 슬라이드 계획
- `src/template.mjs`: PPTX OOXML 분석, 역할 추정, 템플릿 슬라이드 복제, 표·차트·노트 교체
- `src/renderers.mjs`: 기본 PPTX, Word, Excel, HTML 생성
- `src/build.mjs`: 출력 검증, 템플릿 적용, 결과 생성과 manifest 작성
- `src/server.mjs`: 로컬 Studio API와 업로드·미리보기·다운로드
- `src/studio.html`: 원고 편집, 디자인 매핑, 미리보기 UI
- `schema/report.schema.json`: 입력 계약

## 4. 입력 계약

원고는 `schema/report.schema.json`을 따르는 JSON이다.

- `title`, `subtitle`, `author`, `date`, `disclosure`
- `sources`: 출처명, 발행기관, 날짜, URL, 위치
- `datasets`: 프로젝트 폴더 안의 CSV/XLSX 연결
- `metrics`: `sum`, `mean`, `first`, `last`, `count`, `growthPct`
- `sections`: summary, metrics, chart, table, comparison, timeline, text, conclusion
- `template`: 기존 회사 PPTX 경로와 슬라이드·텍스트 역할 매핑

모든 파일 경로는 프로젝트 폴더 안의 상대 경로여야 한다. 출력은 입력 원고나 데이터 파일을 덮어쓰지 않는다.

## 5. 기존 회사 PPTX 처리

### 5.1 처리 절차

1. `.pptx` 압축 패키지와 XML 관계를 검사한다.
2. 슬라이드 크기, 슬라이드 순서, 마스터, 레이아웃, 텍스트 상자, 표, 차트를 추출한다.
3. PowerPoint 자리표시자와 `rf:title`, `rf:body` 같은 이름을 우선 사용한다.
4. 이름이 없으면 위치·크기·반복 문구·글꼴 크기를 이용해 역할을 추정한다.
5. 원본 슬라이드를 복제하고 지정된 텍스트·표·차트·노트를 교체한다.
6. 사용하지 않는 원본 슬라이드와 오래된 차트 데이터·노트를 결과에서 제거한다.

### 5.2 보존 대상

- 슬라이드 크기
- 마스터와 레이아웃
- 배경, 로고, 이미지와 색상
- 도형 좌표와 텍스트 영역
- 기존 글꼴과 표·차트 스타일
- 편집 가능한 차트 구조

### 5.3 현재 제한

SmartArt, OLE, 이미지로 된 차트·표, 병합 셀 표, 원형·복합 차트의 데이터 교체는 지원하지 않는다. 그룹 내부 텍스트와 마스터의 고정 문구는 보존 문구로 취급될 수 있으므로 원본 디자인을 정리하거나 역할을 직접 지정해야 한다.

## 6. 사용자 커스터마이즈 영역

### 6.1 기본 테마

`themes/corporate.json`, `themes/editorial.json`을 복사해 새 테마를 만들 수 있다.

```json
{
  "name": "my-company",
  "font": "Malgun Gothic",
  "background": "FFFFFF",
  "foreground": "1D2C36",
  "accent": "165D8C",
  "secondary": "6B8E23",
  "muted": "647783",
  "rule": "DCE3E7",
  "tableHeader": "165D8C",
  "tableStripe": "F4F6F7"
}
```

### 6.2 회사 템플릿 매핑

```json
{
  "template": {
    "file": "templates/company/source.pptx",
    "layouts": {
      "cover": 1,
      "flow": 2,
      "metrics": 3,
      "chart": 4,
      "table": 5,
      "references": 2
    },
    "roles": {
      "2": {
        "3": "title",
        "4": "lead",
        "5": "body",
        "6": "footer",
        "7": "page"
      }
    }
  }
}
```

역할에는 `title`, `lead`, `body`, `footer`, `page`, `meta`, `disclosure`, `unit`, `note`, `keep`, `clear`, `metric1.label`, `metric1.value`, `metric1.description` 등이 있다.

### 6.3 원고 블록

사용자는 섹션 순서와 유형을 바꿀 수 있다. 차트의 형태, 범주 열, 계열, 단위, 표의 열, 비교 선택지와 일정 단계를 직접 지정한다. `{{metricId}}` 표기로 데이터에서 계산한 값을 본문에 삽입한다.

### 6.4 패턴 라이브러리 확장

향후 `patterns/` 디렉터리에 발표 유형을 추가한다.

- `executive-summary`: 경영진 요약
- `sales-performance`: 매출·성과 분석
- `strategy-proposal`: 선택지·전략 제안
- `project-status`: 일정·진척 보고
- `research-report`: 조사·리서치 보고
- `personal-project`: 개인 프로젝트 설명

패턴은 권장 섹션 순서, 필요한 입력, 사용 가능한 데이터 유형, 기본 테마, 슬라이드 수 제한을 선언한다.

## 7. 추가 개발 항목

### 우선순위 1

- 자연어 입력을 `report.json`으로 변환하는 로컬 LLM 어댑터
- 패턴 라이브러리와 입력 내용 기반 패턴 추천
- Studio에서 패턴 선택·복제·삭제·저장
- 실제 회사 PPTX를 사용한 PowerPoint 검증 자동화

### 우선순위 2

- Word 보고서에도 회사 브랜드 토큰 적용
- 사용자 정의 슬라이드 영역 편집기
- 표·차트 계열명과 스타일 매핑 UI
- 데이터 파일 업로드 및 열 매핑 화면
- CSV/XLSX 변경 감지와 재생성

### 우선순위 3

- SmartArt·병합표·복합 차트 지원
- PDF 및 이미지 내보내기
- 검토 이력과 결과 버전 비교
- 로컬 패턴 저장소와 팀 공유 규칙
- 접근성 검사와 색상 대비 검사

## 8. 로컬 LLM 연동 설계

현재 Report Forge 자체는 LLM을 호출하지 않는다. 향후 LLM을 붙일 때도 생성기와 모델을 분리한다.

```text
자연어 입력
   ↓
LocalModelAdapter
   ↓
validated report.json
   ↓
기존 결정론적 생성 파이프라인
```

지원 후보는 Ollama와 LM Studio다. 모델은 숫자를 직접 계산하지 않고 원고 JSON 초안을 작성한다. 숫자 계산, 출처 연결, 파일 생성은 기존 로컬 코드가 담당한다.

## 9. 검증 기준

- `npm test`: 단위·통합 테스트 19개 이상 통과
- 기본 출력의 PPTX 패키지 구조 검사
- 회사 템플릿의 디자인 자산 해시 보존 검사
- 기존 업무 문구 제거 검사
- 새 차트의 독립적인 내장 워크북 검사
- 표 전체 행 보존 검사
- Studio 업로드·매핑·저장·재열기·다운로드 검사
- 실제 배포 전 PowerPoint와 Word에서 최종 표시 확인

## 10. 실행 명령

```sh
npm ci --ignore-scripts
npm test
npm run demo
npm run studio
npm run demo:company
node src/cli.mjs validate projects/my-report/report.json
node src/cli.mjs build projects/my-report/report.json --out outputs/my-report
```

## 11. 비기능 요구사항

- Node.js 22 이상
- Windows, macOS, Linux의 표준 Node.js 실행 환경
- 외부 API 없이 기본 기능 실행
- 프로젝트 폴더 밖 파일 접근 차단
- 입력 파일 크기와 압축 해제 크기 제한
- 생성 실패 시 기존 결과 보존
- 제품 코드에 특정 IDE의 절대 경로 또는 전용 API를 포함하지 않음
