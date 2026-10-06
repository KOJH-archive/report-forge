# 참고 프로젝트와 공개 라이브러리

Presenton과 PPT Master는 설계 참고 자료입니다. 해당 저장소의 코드는 이 프로젝트로 복사하지 않았습니다. 구체적인 조사 자료는 RESEARCH.md에 기록했습니다.

프로그램은 다음 공개 패키지에 의존합니다. 버전과 전체 의존성 트리는 package-lock.json에 고정되어 있습니다. 각 패키지의 원본 라이선스는 설치된 패키지와 공식 저장소에서 확인할 수 있습니다.

- pptxgenjs 4.0.1: PowerPoint 생성, https://github.com/gitbrent/PptxGenJS
- docx 9.6.1: Word 문서 생성, https://github.com/dolanmiu/docx
- exceljs 4.4.0: Excel 입력과 데이터 스냅샷, https://github.com/exceljs/exceljs
- ajv 8.17.1: 공통 원고 스키마 검증, https://github.com/ajv-validator/ajv
- csv-parse 5.6.0: CSV 입력, https://github.com/adaltas/node-csv
- jszip 3.10.1: Office 차트 패키지 연결과 테스트, https://github.com/Stuk/jszip
- @xmldom/xmldom 0.9.12: 기존 PPTX의 XML 분석과 서식 보존, https://github.com/xmldom/xmldom (MIT)

현재 프로젝트는 개인 또는 회사의 로컬 개발용이며 공개 배포나 npm 게시를 수행하지 않았습니다. 제품 코드에는 Codex의 관리 런타임 경로나 전용 API를 넣지 않았습니다.
