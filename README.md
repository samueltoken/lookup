# lookup

로그인/결제 없이 바로 쓰는 Windows 문서 뷰어입니다.

## 지원 형식

- `.pdf`
- `.hwp`, `.hwpx`
- `.doc`, `.docx`
- `.xls`, `.xlsx`

비-PDF 문서는 앱 내부에서 PDF로 임시 변환해 표시합니다.  
원본 파일은 수정하지 않습니다.

## 주요 기능

- 파일 열기: `lookup.exe [문서 경로]`
- 연속 스크롤 + 썸네일 패널
- `Ctrl + 마우스 휠` 부드러운 확대/축소
- 검색 결과 목록 + 단어 위치 하이라이트
- 주석(형광펜/펜/텍스트 메모), 페이지 삭제/회전/순서 변경
- 전체화면(`F11`) + 연속/현재페이지 모드 전환
- 인쇄(`Ctrl+P`)
- 설정 모달:
  - 한국어/영어 전환
  - 개발자 문의 이메일 복사
- GitHub Release 기반 자동 업데이트(진행률 표시)

## 개발 실행

```bash
npm install
npm start
```

## 설치 파일 빌드

```bash
npm install
npm run dist
```

생성 파일:

- `release/lookup-Setup-x.y.z.exe`
- `release/latest.yml`
- `release/*.blockmap`

## 자동 업데이트 설정

`update-config.json`에 저장소 정보를 넣으면 GitHub Release로 업데이트를 확인합니다.

```json
{
  "owner": "깃허브아이디",
  "repo": "저장소이름"
}
```

비워 두면 `package.json`의 `repository` URL에서 자동 추론합니다.

## 제거(언인스톨)

- Windows 설정 > 앱 > 설치된 앱 > `lookup` 제거
- 또는 시작 메뉴의 `lookup Uninstall` 실행
