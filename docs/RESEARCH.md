# 발표 자료 자동화 프로젝트 조사

조사일은 2026년 10월 6일입니다. GitHub의 구현 문서와 라이선스를 확인하고, Reddit과 Hacker News의 실제 문제 제기를 함께 검토했습니다. 프로젝트 소개에서 주장하는 지원 범위는 직접 실행해 검증한 범위와 구분합니다.

## 가장 관련성이 높은 구현 참고 자료

### Presenton

- 저장소: https://github.com/presenton/presenton
- 구체적 아티팩트: https://github.com/presenton/presenton/blob/main/docs/template-v2.md
- 라이선스: https://github.com/presenton/presenton/blob/main/LICENSE, Apache 2.0

프로젝트 소개는 자체 호스팅, AI 모델 선택, 사용자 서식, 편집 가능한 PPTX와 API를 제공합니다. Template V2 문서는 기존 PPTX에서 재사용 레이아웃을 만들고, 레이아웃별 콘텐츠 스키마에 맞는 내용을 생성하는 흐름을 설명합니다. 기존 PPTX를 분석해 서식을 만드는 경로에는 슬라이드 미리보기와 이미지 입력이 가능한 모델이 필요합니다.

사용자의 문제와 가장 가까운 부분은 **내용을 생성하는 단계와 서식을 적용하는 단계의 분리**입니다. Report Forge는 공통 원고, 데이터, 테마를 분리하는 구조를 참고했습니다. 기존 PPTX 자동 분석이나 Presenton의 실행 환경 전체를 첫 버전의 필수 의존성으로 가져오지는 않았습니다. 따라서 첫 버전에서 회사 PPTX를 그대로 복제할 수 있다는 의미는 아닙니다.

### PPT Master

- 저장소: https://github.com/hugohe3/ppt-master
- 소개: https://github.com/hugohe3/ppt-master/blob/main/README.md
- 라이선스: https://github.com/hugohe3/ppt-master/blob/main/LICENSE, MIT

프로젝트 소개는 여러 AI 작업 도구에서 실행하는 워크플로, 재사용 가능한 브랜드와 레이아웃, 데이터 기반의 PowerPoint 차트와 표를 설명합니다. 여러 도구에서 작동한다는 설명은 프로젝트 작성자의 지원 주장입니다. 조사 과정에서 Antigravity의 실제 UI에 설치하거나 동작을 확인하지는 않았습니다.

Report Forge는 **AI 작업 지침을 파일로 전달하고 독립 실행 도구로 결과를 만드는 방식**, **Office 차트와 표를 그대로 편집할 수 있게 유지하는 방향**을 참고했습니다. PPT Master의 코드를 복사하거나 해당 워크플로를 설치하지 않았습니다.

### PptxGenJS

- 저장소: https://github.com/gitbrent/PptxGenJS
- 차트 API: https://gitbrent.github.io/PptxGenJS/docs/api-charts/
- 실제 구현 예제: https://github.com/gitbrent/PptxGenJS/blob/master/demos/modules/demo_chart.mjs
- 라이선스: https://github.com/gitbrent/PptxGenJS/blob/master/LICENSE, MIT

JavaScript로 PowerPoint 파일을 생성하는 라이브러리입니다. 차트 API와 예제의 사용 방식을 확인했습니다. Report Forge는 이 공개 라이브러리를 실제 출력 엔진으로 사용합니다. 의존성 버전은 4.0.1로 고정했습니다.

검증 중 같은 프로세스에서 여러 번 생성하면 차트 파일 번호가 계속 증가하는 것을 확인했습니다. Word 출력은 차트 번호가 항상 1에서 시작한다고 가정하지 않고 실제 슬라이드 관계 파일을 따라 차트를 연결합니다. 반복 내보내기 테스트가 이 경우를 검증합니다. 구조 검사에서는 실제 파일이 없는 추가 슬라이드 마스터의 콘텐츠 유형 선언을 발견했습니다. 실제 개체는 보존하고 대상 파일이 없는 선언만 정리한 뒤 재검증합니다.

### Quarto

- 저장소: https://github.com/quarto-dev/quarto-cli
- PowerPoint 문서: https://quarto.org/docs/presentations/powerpoint.html
- Word 문서: https://quarto.org/docs/output-formats/ms-word.html

텍스트와 분석 코드를 바탕으로 여러 문서 형식을 생성하는 대안입니다. 분석 중심의 정기 보고서라면 채택할 가치가 있습니다. 이 첫 버전에서는 페이지별 블록 구성과 Office 차트를 직접 다루기 위해 독립 생성기를 구현했습니다. Quarto는 현재 제품 의존성에 포함하지 않았습니다.

## 커뮤니티에서 확인한 문제

### Reddit

- 논의: https://www.reddit.com/r/ClaudeAI/comments/1v7xmn8/suggestions_for_presentation_generation/

회사 서식에 맞는 발표 자동화를 논의하며, 승인한 레이아웃을 먼저 정하고 AI가 구조화된 콘텐츠를 작성하게 하는 접근이 제안됩니다. 일부 댓글에는 요약·표·차트용 슬라이드와 페이지별 입력 지침을 준비한 작업 경험도 소개됩니다.

이는 활용 사례와 요구를 파악하는 참고입니다. 댓글의 제작 시간이나 품질 주장을 성능 검증으로 사용하지 않았습니다. Report Forge의 숫자 처리와 출력 기능에 대한 기술 판단은 라이브러리 문서와 직접 실행한 테스트에 근거합니다.

### Hacker News

- 논의: https://news.ycombinator.com/item?id=48518648
- 제목: Show HN: Brightdeck – an OOXML-compatible AI presentation maker

제작자 게시물은 기존 PowerPoint의 가져오기와 편집, 레이아웃 반복, PPTX 호환성을 문제로 제시합니다. 사용자가 기존 회사 자료를 계속 활용하려는 경우 브라우저 화면의 완성도만으로 충분하지 않다는 점을 보여 주는 사례입니다.

Brightdeck은 여기에서 공개 라이브러리로 채택한 아티팩트가 아닙니다. 게시물과 댓글만으로 다른 제품보다 우수하다고 판단하지 않았습니다. 이 논의에서 도출한 개발 판단은 **편집 가능한 Office 개체를 만들고 실제 Office 표시 검증을 별도로 남겨야 한다**는 것입니다.

## 개발에 반영한 선택

1. 원고와 데이터, 출처를 공통 모델로 저장하고 서식을 별도로 적용합니다.
2. AI는 각 IDE에서 원고를 작성합니다. 자료 생성은 일반 Node.js 프로그램이 실행합니다.
3. 숫자는 데이터에서 계산합니다. 동일한 Office 차트 패키지와 내장 워크북을 PPT와 Word에서 사용합니다.
4. 자료가 많으면 표와 본문을 여러 슬라이드로 분할합니다. 근거를 임의로 잘라 내지 않습니다.
5. API 키가 없어도 샘플 생성, 원고 편집, 내보내기가 가능합니다. 입력 자료를 외부 서비스로 전송하지 않습니다.

## 검증 결과와 남은 작업

19개의 자동 테스트로 계산, 출처 연결, CSV/XLSX 입력, 음수 차트, 표 행 보존, 편집 가능한 차트 패키지, 반복 생성, 로컬 브라우저 내보내기, 한국어 경로 및 외부 작업 디렉터리 실행과 기존 회사 PPTX의 디자인 가져오기를 확인했습니다.

표준 Windows Node.js 실행을 확인합니다. Antigravity 실제 UI와 macOS/Linux는 직접 검증하지 않았습니다. Word 렌더러를 실행했지만 이 컴퓨터에 LibreOffice가 없어 실제 Word 페이지 렌더링은 확인하지 못했습니다. 브라우저 미리보기와 Office 파일의 실제 표시를 구분합니다.

0.2에서는 기존 PPTX를 분석해 재사용 레이아웃을 추정하고 원본 OOXML 슬라이드를 복제하는 기능을 구현했습니다. Presenton Template V2의 고정 디자인·가변 콘텐츠 분리와 PPT Master의 기존 디자인 재사용 원칙을 참고했습니다. 원본 코드나 해당 제품의 AI API를 가져오지 않았습니다. 실제 회사 자료가 없어 가상의 4:3 회사 서식으로 디자인 보존과 내용 교체를 검증했습니다.

Microsoft 공식 [슬라이드 구조](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-presentation-slides), [마스터 구조](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-slide-masters), [노트 구조](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides)를 참고해 슬라이드·차트·워크북의 연결을 보존했습니다. 다음 검증 대상은 실제 회사 PPTX와 Office에서의 최종 표시입니다. 데이터 파일 업로드 화면과 지원 개체 확대도 후속 개발 범위입니다.
