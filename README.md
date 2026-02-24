# PassNote Excel Tool

서버 없이(브라우저에서만) 엑셀을 처리하는 작은 도구입니다.

- A/B 엑셀을 `ISBN`으로 비교해서 **교집합만 추출**
- 단일 엑셀에서
  - **상품명에 `예약판매` 포함 행 제외**
  - **출판연도(2016~2026) 선택 필터**
- 결과를 계속 **누적**해두었다가 마지막에 **엑셀로 저장(XLSX 다운로드)**

## 컬럼명 규칙

기본적으로 아래 컬럼명을 사용합니다.

- `ISBN`
- `상품명`
- `출판날짜`

## 로컬 실행

```bash
npm install
npm run dev
```

## 빌드

```bash
npm run typecheck
npm run build
```

## GitHub Pages 배포

이 레포에는 GitHub Pages 자동 배포 워크플로우가 포함되어 있습니다. (`.github/workflows/pages.yml`)

1) `main` 브랜치에 푸시  
2) GitHub 레포 설정에서 Pages → Source를 **GitHub Actions**로 선택  
3) Actions가 빌드 후 `dist/`를 Pages로 배포합니다.

## 성능 메모

- 행이 2만 줄 이상이어도, 표 미리보기는 **가상 스크롤**이라 브라우저가 버티도록 설계했습니다.
- 다만 엑셀 파싱은 파일 크기에 따라 시간이 걸릴 수 있습니다.
# passnote-excel-tool
