# ESP Lab 홈페이지

한양대학교 ERICA 임베디드 보안 및 프라이버시 연구실(ESP Lab) 홈페이지의 소스입니다.
공개 주소는 <https://esplab.hanyang.ac.kr> 이고, GitHub Pages가 이 저장소의 `main` 브랜치에서 빌드해 배포합니다.

사이트는 정적 페이지 하나로 되어 있습니다. 메뉴는 주소의 `#` 뒤 이름으로 바뀌고, 한국어와 영어 전환, 다크 모드가 있습니다.

## 어디를 고치나

| 고칠 내용 | 파일 |
| --- | --- |
| 디자인, 문구, 레이아웃 | `index.tpl.html` (위쪽 `<style>`이 CSS, 가운데가 마크업, 맨 아래 `<script>`가 데이터를 화면에 그리는 코드) |
| 소식, 행사 | `data/news.json` |
| 구성원 (박사, 석사, 학부, 졸업생) | `data/members.json` |
| 교수 소개 (경력, 학력, 수상) | `data/professor.json` |
| 초청 강연 | `data/talks.json` |
| 연구과제 | `data/projects.json` |
| 언론 보도 | `data/press.json` |
| 첫 화면 사진, 연구 분야 설명, 대표 논문 | `data/site.json` |
| 논문 제목에 걸리는 링크 (DOI 등) | `data/pub_links.json` |
| 논문별 연구 분야 보정 | `data/area_overrides.json` |
| 논문, 특허, 소프트웨어 목록 | `data/publications.json`. 손으로 고치지 않고 아래 "논문 목록 갱신"대로 Excel에서 만듭니다. |
| 사진 | `images/`. 파일 이름은 16진수 12자리에 `.jpg`를 붙이고, JSON에서 `images/<이름>.jpg`로 가리킵니다. |

한국어와 영어가 같이 있는 항목은 `ko`/`en` 같은 짝으로 들어 있습니다. 기존 항목 하나를 복사해 고치는 것이 가장 안전합니다.

다음 파일은 빌드가 만드는 결과물이라 git에 넣지 않고, 직접 고쳐도 다음 빌드 때 사라집니다: `index.html`, `files.json`, `dist/`.

## 수정 절차

처음 한 번은 저장소를 받고 Python 패키지를 설치합니다. Python 3.11 이상이면 됩니다.

```bash
git clone https://github.com/Hanyang-ESPLab/Hanyang-ESPLab.github.io.git
cd Hanyang-ESPLab.github.io
pip install -r requirements.txt
```

고칠 때마다 다음 순서로 합니다.

1. 위 표의 파일을 고칩니다.
2. 빌드합니다. 결과는 `dist/`에 생깁니다.

   ```bash
   python build.py --no-excel
   ```

3. 로컬에서 확인합니다. 아래를 실행한 뒤 브라우저에서 <http://127.0.0.1:8765> 를 엽니다.

   ```bash
   python3 -m http.server 8765 --bind 127.0.0.1 --directory dist
   ```

4. 고친 소스 파일만 커밋하고 `main`에 push합니다. GitHub Actions가 빌드해서 1~2분 뒤 사이트에 반영됩니다.
   진행 상황과 실패 여부는 저장소의 Actions 탭에서 봅니다. 빌드가 실패하면 이전 사이트가 그대로 남습니다.

소식 한 줄처럼 JSON만 고치는 작은 수정은 GitHub 웹 편집기에서 바로 고치고 commit해도 됩니다. 이때도 Actions가 빌드합니다.

## 논문 목록 갱신

논문, 특허, 소프트웨어 목록은 연구실 Dropbox의 Excel에서 만듭니다.

- `ESP Lab Dropbox/ESPLAB/연구실 실적/논문실적/ESP 논문실적 (최신).xlsx`
- `ESP Lab Dropbox/ESPLAB/연구실 실적/특허_소프트웨어_실적/ESP 특허실적 (최신).xlsx`

Dropbox가 설치된 컴퓨터에서 `--no-excel` 없이 빌드하면 `data/publications.json`이 Excel 내용으로 다시 만들어집니다.
이 파일을 커밋해서 push하면 사이트에 반영됩니다.

```bash
python build.py
```

Dropbox 폴더가 홈 디렉터리의 `ESP Lab Dropbox/ESPLAB/연구실 실적`이 아니면 환경 변수 `ESPLAB_LAB_DIR`에 그 경로를 넣습니다.

새 논문의 링크는 `data/pub_links.json`에 종류별로 `"논문 제목": "URL"` 한 줄을 더합니다. 제목은 사이트에 보이는 것과 같아야 하고, 대소문자와 공백 수는 달라도 됩니다.
빌드할 때 어느 논문과도 맞지 않는 제목이 있으면 경고를 출력합니다.

## 하지 말 것

- `python build.py --sync-gs`는 쓰지 않습니다. 옛 Google Sites 페이지를 다시 받아 와서 `data/members.json`, `news.json`, `professor.json`, `talks.json`의 손편집을 덮어씁니다.
- `data/publications.json`은 손으로 고치지 않습니다. Excel에서 다시 만들 때 사라집니다.
- `index.html`, `dist/`, `files.json`은 고치거나 커밋하지 않습니다.

## 파일 구성

```
index.tpl.html   페이지 템플릿 (CSS, 마크업, 렌더링 스크립트)
data/            사이트에 들어가는 내용 (JSON)
images/          사진
logo.png         로고
build.py         data/*.json + index.tpl.html -> index.html, 그리고 make_dist.py로 dist/
make_dist.py     index.html을 완전한 HTML 문서로 감싸고 사진을 복사해 dist/를 만듦
gs_parse.py      옛 Google Sites 페이지 해석기 (--sync-gs 전용, 평소에는 쓰지 않음)
requirements.txt openpyxl, beautifulsoup4
.github/workflows/pages.yml   push 때 빌드하고 GitHub Pages에 배포
_drafts/         디자인 시안 등 임시 작업 (git에 넣지 않음)
```
