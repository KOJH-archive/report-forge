# Report Forge

> Local-first, source-aware presentation and report generator that reuses company PPTX designs and creates editable PowerPoint, Word, and Excel outputs from structured content and data.

GitHub: [KOJH-archive/report-forge](https://github.com/KOJH-archive/report-forge)

원고와 데이터, 출처를 한 번 정리하면 같은 내용으로 발표 자료와 보고서를 만드는 로컬 제작기입니다. Codex, Antigravity, VS Code, Cursor 또는 일반 터미널에서 같은 명령을 사용합니다.

주제별로 레이아웃을 다시 만드는 수고를 줄이는 데 집중합니다. AI 도구가 공통 원고를 작성하고, 독립 실행 프로그램이 정해진 서식에 맞춰 자료를 생성합니다. **0.2에서는 기존  PPTX의 디자인을 가져와 내용만 교체할 수 있습니다.** 설치 후 샘플 생성에는 AI 계정, API 키, 인터넷 연결이 필요하지 않습니다.

## 시작하기

Node.js 22 이상을 설치한 뒤 이 폴더에서 실행합니다.

```sh
npm ci --ignore-scripts
npm run demo
npm run studio
```

브라우저에서 `http://127.0.0.1:8765`를 열면 자료 정보, 목차와 원고, 디자인을 수정할 수 있습니다. **원고 저장**은 입력한 원고 파일을 갱신합니다. **자료 내보내기**는 현재 화면의 원고로 자료를 생성하고 다운로드 버튼을 표시합니다. 저장하지 않은 수정도 내보내기에 반영합니다.

샘플은 `outputs/demo/`에 다음 파일을 만듭니다.

- `presentation.pptx`: 기본 16:9 또는  PPTX의 원래 크기를 유지하는 발표 자료, 편집 가능한 표와 Office 차트, 출처와 발표자 노트
- `report.docx`: 상세 설명을 포함한 Word 보고서, 편집 가능한 표와 Office 차트, 참고 자료 목록
- `evidence.xlsx`: 사용한 데이터 스냅샷, 계산 지표, 출처 목록
- `preview.html`: 브라우저에서 발표 자료와 보고서 보기를 전환하는 미리보기
- `resolved-report.json`: 계산을 반영한 원고와 데이터
- `manifest.json`: 입력 및 출력 파일 해시, 생성 정보, 검증 범위

샘플의 매출과 계획은 서식 검증용 **가상 데이터**입니다.

## 새 주제로 제작하기

```sh
node src/cli.mjs init projects/my-report
node src/cli.mjs studio projects/my-report/report.json
```

`report.json`의 원고, 출처, 데이터 연결을 수정합니다. CSV 또는 XLSX 파일은 그 원고 폴더 안에 넣습니다. 샘플 폴더는 새 프로젝트를 시작하기 위한 예제이며 실제 보고 자료로 바꿔서 사용해야 합니다.

IDE의 AI에게는 [공통 작업 지침](prompts/BUILD_REPORT.md)을 읽고 작업하도록 요청하면 됩니다. 지침의 예시 프롬프트에 주제, 발표 대상, 핵심 결론, 원본 파일을 넣습니다. 특정 IDE의 전용 명령이나 계정에 의존하지 않습니다.

화면 없이 실행할 수도 있습니다.

```sh
node src/cli.mjs validate projects/my-report/report.json
node src/cli.mjs build projects/my-report/report.json --out outputs/my-report
```

기존 결과를 갱신하려면 `--force`를 지정합니다. 원고와 데이터 검증이 실패하면 기존 출력 파일을 수정하지 않습니다. 파일 복사 도중 발생하는 디스크 오류에 대해서는 전체 출력의 원자적 교체를 보장하지 않습니다.

## 원고와 디자인

지원하는 블록은 핵심 요약, 주요 지표, 차트와 해석, 수치 표, 선택지 비교, 실행 일정, 상세 설명, 결론과 제안입니다. 표의 행이나 본문이 많으면 발표 자료에서는 이어지는 페이지를 생성합니다. 데이터 행을 임의로 생략하지 않습니다.

`theme`을 `corporate` 또는 `editorial`로 바꾸면 같은 내용에 다른 색상 서식을 적용합니다. 글꼴과 색상은 `themes/`의 일반 JSON 파일에 있습니다. 용 테마를 따로 만들 때는 복사한 테마 파일을 사용합니다.

```sh
node src/cli.mjs build projects/my-report/report.json --out outputs/my-report --theme themes/my-company.json
```

기본 글꼴은 맑은 고딕입니다. macOS 또는 Linux에서는 사용 환경에 설치된 한국어 글꼴 이름을 테마의 `font`에 지정해야 합니다. 첫 버전은 글꼴 파일을 내장하지 않습니다.

##  PPTX 디자인 가져오기

자료 제작 화면의 **자료 정보 →  PPTX 디자인 → 디자인 가져오기**에서 기존 PPTX를 선택합니다. 자동으로 표지·본문·지표·차트·표에 사용할 원본과 텍스트 역할을 찾습니다. 잘못 찾은 위치는 같은 화면의 슬라이드 선택과 역할 설정에서 수정합니다. **원고 저장**을 누르면 선택한  디자인이 다음 실행에도 적용됩니다.

의 PPTX는 원고 폴더의 `templates/`에 복사됩니다. 원본 파일을 변경하거나 외부로 전송하지 않습니다. 그림으로 다시 만드는 방식이 아니라 원본 슬라이드의 내부 구조를 복제하므로 마스터, 레이아웃, 배경, 로고, 이미지, 좌표, 글꼴, 색상과 표·차트의 기존 서식을 유지합니다. 표 데이터, 차트 캐시와 편집용 엑셀 데이터, 발표자 노트는 새 내용으로 교체합니다. 사용하지 않는 원본 슬라이드, 차트 데이터와 기존 노트는 출력 패키지에서 제거합니다.

명령으로 가져올 수도 있습니다.

```sh
node src/cli.mjs import-template company.pptx --out projects/my-report/templates/company
```

이후 해당 원고의 `report.json`에 다음 속성을 추가합니다.

```json
"template": { "file": "templates/company/source.pptx" }
```

`design.json`은 분석 결과이며 직접 편집해도 생성에 반영되지 않습니다. 사용자 수정은 원고의 `template.layouts`와 `template.roles`에 저장합니다. 사용 방법과 지원 범위는 [ 디자인 안내](docs/COMPANY_DESIGN.md)에 있습니다.

검증용  디자인 예제도 제공됩니다. 실제  자료가 아닌 가상의 4:3 서식입니다.

```sh
npm run demo:company
```

원본은 `outputs/company-demo/company-source.pptx`, 생성 결과는 `outputs/company-demo/result/presentation.pptx`에 있습니다. 원본을 화면에서 가져와 기능을 시험할 수 있습니다.

 디자인 복제는 PowerPoint에 적용됩니다. Word는 기본 보고서 서식을 사용합니다. 브라우저 미리보기와 위치 지도는 내용과 교체 위치를 확인하기 위한 화면이며 원본 PPTX의 실제 표시를 재현하지 않습니다.

## 데이터와 숫자

CSV의 첫 행과 XLSX 시트의 첫 행은 고유한 열 이름이어야 합니다. XLSX는 지정한 시트를 읽고, 시트를 지정하지 않으면 첫 번째 시트를 읽습니다. 원고 폴더 밖의 파일이나 폴더 밖을 가리키는 링크는 읽지 않습니다.

지원하는 계산은 `sum`, `mean`, `first`, `last`, `count`, `growthPct`입니다. `count`는 행 수를 셉니다. `growthPct`는 첫 행 대비 마지막 행의 변화율을 계산하며, 원본의 행 순서를 그대로 사용합니다. 첫 값이 0이면 계산을 중단합니다.

원고에 `{{onlineTotal}}` 같은 지표 ID를 넣으면 표시 단위를 포함한 계산값으로 치환합니다. 수동으로 입력한 수치의 의미까지 자동 검증하지는 않습니다.

엑셀 수식은 파일에 저장된 계산 결과를 읽습니다. 수식을 다시 계산하지 않습니다. 계산 결과가 없으면 Excel 또는 LibreOffice에서 계산한 뒤 저장하라는 오류를 표시합니다. 원본 엑셀의 수식과 서식을 그대로 복제하는 기능은 없습니다.

차트는 선 그래프 또는 막대그래프, 최대 24개의 범주와 3개 계열을 지원합니다. 한 차트의 계열은 같은 단위를 사용해야 합니다. 범주가 많으면 먼저 원본 데이터를 집계합니다.

## 출처

자료마다 고유한 `source.id`, 자료명, 작성 기관, 날짜, URL 또는 파일 내 위치를 적습니다. 데이터의 출처는 연결된 표, 차트, 지표에 자동으로 전달합니다. 발표 자료에는 짧은 출처 표시와 참고 자료 페이지를, Word에는 블록별 출처와 참고 자료 목록을 생성합니다.

웹 자료를 자동 수집하거나 출처가 주장을 실제로 뒷받침하는지 판정하는 기능은 첫 버전에 포함하지 않습니다. 주장의 정확성과 출처의 대응 관계는 작성자가 확인해야 합니다.

## 여러 IDE에서 사용하기

프로젝트 전체를 다른 IDE에서 열고 동일한 `npm` 명령을 실행합니다. 원고는 JSON, 데이터는 CSV 또는 XLSX, 디자인은 JSON이므로 특정 앱의 내부 아티팩트 형식에 묶이지 않습니다. 자동 원고 작성은 각 IDE의 AI가 [공통 작업 지침](prompts/BUILD_REPORT.md)을 읽는 방식입니다. 별도 LLM API 호출은 하지 않습니다.

표준 Windows Node.js와 한국어 경로, 프로젝트 외부 작업 디렉터리에서 실행을 검증했습니다. Antigravity의 실제 UI 및 macOS/Linux 실행은 이 환경에서 직접 검증하지 않았습니다.

## 검증과 현재 범위

```sh
npm test
```

자동 테스트는 수치 계산, 데이터 연결, 잘못된 출처, 빈 숫자, 엑셀 수식 결과, 표 페이지 분할, Office 패키지 안의 차트와 표, 반복 생성, 브라우저 내보내기, 일반 실행 환경을 확인합니다.

HTML 미리보기는 내용과 구성의 확인용입니다. PowerPoint와 Word가 파일을 표시한 화면을 재현하는 것은 아닙니다. 현재 환경에 Office와 LibreOffice가 없어 실제 Office의 최종 표시를 검증하지 못했습니다. 배포 전에는 각 프로그램에서 글꼴, 차트, 표의 줄바꿈과 페이지 구성을 확인해야 합니다.

현재 지원하지 않는 기능은 PDF 직접 생성, 실시간 공동 편집, 웹 검색과 출처 자동 검증입니다.  디자인에서는 이미지로 된 표·차트, SmartArt의 내용 교체, 병합 표, 복합·원형 차트의 데이터 교체를 지원하지 않습니다. 지원하지 않는 개체나 부족한 교체 영역은 오류로 알려 줍니다. 연구 참고 자료와 구현 선택은 [조사 기록](docs/RESEARCH.md)에 있습니다.

## 개발명세서와 확장

현재 아키텍처, 입력 계약,  PPTX 처리 방식, 테마·패턴·레이아웃 커스터마이즈 방법, 로컬 LLM 연동 계획과 추가 개발 로드맵은 [개발명세서](docs/DEVELOPMENT_SPEC.md)에 정리되어 있습니다. 기존  PPTX를 가져오는 상세 사용법은 [ 디자인 안내](docs/COMPANY_DESIGN.md)를 참고하세요.

## GitHub Topics

`presentation-generator` `pptx` `report-generator` `document-automation` `local-first` `nodejs` `pptx-template` `data-visualization` `office-automation`
