# lookup (윈도우 PDF 뷰어)

로그인/결제 없이 로컬 PDF만 바로 열어보는 데스크톱 앱입니다.

## 핵심 기능

- 파일 열기 (`lookup.exe [PDF경로]` 지원)
- 왼쪽 미리보기 썸네일
- 연속 스크롤 보기 (기본)
- 확대/축소 (`Ctrl + 마우스 휠` 포함)
- 텍스트 검색 + 노란 하이라이트
- 페이지 회전/삭제/순서 변경(드래그)
- 주석(형광펜/펜/텍스트 메모)
- 저장: 다른 이름 저장 / 원본 덮어쓰기
- 전체화면 (`F11`) + 보기 모드 전환
- 인쇄 (`Ctrl+P`)
- 다크모드

## 설치 파일 만들기

```bash
npm install
npm run dist
```

생성 파일:

- `release/lookup-Setup-1.1.0.exe`

설치 중에 `.pdf 파일을 lookup으로 열기` 체크를 선택하면 더블클릭 연결이 설정됩니다.

## 개발 실행

```bash
npm install
npm start
```

## 자동 업데이트 (GitHub Release)

1. `update-config.json`에 저장소 정보 입력
2. 앱 버전(예: `1.1.0`) 올리고 새 설치 파일 빌드
3. GitHub 릴리즈에 아래 파일 업로드
   - `lookup-Setup-x.y.z.exe`
   - `latest.yml`
   - `*.blockmap`

`update-config.json` 예시:

```json
{
  "owner": "깃허브아이디",
  "repo": "저장소이름"
}
```

## 제거

- Windows `설정 > 앱 > 설치된 앱`에서 `lookup` 제거
- 또는 시작 메뉴의 `lookup Uninstall` 실행
