# 조건부확률 이중교차표

GitHub Pages에서 바로 실행할 수 있는 바닐라 JavaScript 기반 교차표 웹앱입니다.

## 파일 구성

- `index.html`: 정적 웹앱 본문입니다.
- `data.xlsx`: 기본으로 자동 로드되는 설문 응답 파일입니다.
- `crosstab_app.py`: 기존 Python 로컬 서버 버전입니다.
- `run_crosstab.bat`: 기존 Windows 실행 파일입니다.

## GitHub Pages 배포

1. GitHub 저장소 루트에 `index.html`과 `data.xlsx`를 올립니다.
2. 저장소에서 `Settings` > `Pages`로 이동합니다.
3. `Build and deployment`의 `Source`를 `Deploy from a branch`로 선택합니다.
4. `Branch`를 `main`, 폴더를 `/root`로 선택한 뒤 저장합니다.
5. 잠시 후 `https://사용자명.github.io/저장소명/` 주소에서 열 수 있습니다.

## 데이터 공개 주의

GitHub Pages로 공유하면 저장소에 올린 `data.xlsx`도 공개됩니다. 응답 원본을 공개하고 싶지 않다면 `data.xlsx`를 올리지 않고, 웹앱 화면의 엑셀 파일 선택 기능으로만 사용하세요.

## 사용한 외부 라이브러리

엑셀 파일 파싱을 위해 브라우저에서 SheetJS CDN을 불러옵니다. 별도 빌드나 설치 과정은 없습니다.
