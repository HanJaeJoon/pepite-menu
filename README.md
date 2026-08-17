# PEPITE 메뉴판

메뉴 내용은 `data/*.json`, 레이아웃은 `build.py` 로 관리하고 docx/pdf 는 빌드해서 만든다.
docx 를 직접 손으로 고치면 정렬과 페이지 구성이 다시 깨지므로 하지 않는다.

## 매장

레스토랑 PÉPITE.

- 네이버 지도: https://naver.me/xQiV4L8E
  (단축 URL 만료 대비 원본: `https://map.naver.com/p/entry/place/2083321247`, place ID `2083321247`)
- 주소: (미기재)
- 전화: (미기재)
- 영업시간: (미기재)

매장 정보는 메뉴판 인쇄물에 들어가지 않는 참고 정보다. 나중에 메뉴판 하단에 주소/전화를
넣기로 하면 여기 대신 `data/store.json` 을 만들어 `build.py` 가 읽게 한다.

## 빌드

```bash
pip install -r requirements.txt
python build.py                                                  # dist/*.docx
powershell -ExecutionPolicy Bypass -File tools/export_pdf.ps1     # dist/*.pdf  (Windows + Word)
```

`python build.py --debug` 로 실행하면 보이지 않는 표 테두리가 빨간 실선으로 그려져서
칸 배치를 눈으로 확인할 수 있다.

## 산출물

| 파일 | 페이지 |
|---|---|
| 페피트 요리 메뉴판 | 4 |
| 페피트 런치 메뉴판 | 2 |
| 페피트 런치 코스메뉴 | 1 (A5) |
| 페피트 와인 메뉴판 | 6 |
| 페피트 위스키 메뉴판 | 1 |

- Word 에서 해당 docx 를 열어둔 채로 빌드하면 저장이 막힌다. 닫고 실행할 것.
- `ChunkFive Ex` 폰트가 설치돼 있어야 로고/섹션 헤더/가격이 의도한 모양으로 나온다.
  없으면 Word 가 대체 폰트로 그린다. (`바탕`, `Times New Roman` 은 Windows 기본)
  폰트 파일은 `fonts/Chunkfive Ex.ttf` 로 동봉했다 (Peter Wiegel 제작, SIL OFL 1.1,
  라이선스 전문도 같은 폴더에 있음). 새 머신에서는 이 파일을 우클릭 -> 설치.
- PDF 에 폰트가 제대로 들어갔는지는 임베드 폰트 목록에 `ChunkFive Ex` 가 있는지로
  확인할 수 있다 (없으면 대체 폰트로 렌더링된 것).

## 메뉴를 바꾸려면

1. `data/food.json`, `data/lunch.json`, `data/wine.json`, `data/course.json`,
   `data/whisky.json` 중 해당 파일 수정 (위스키 병 사진은 `assets/whisky/`)
2. `python build.py`
3. PDF 로 내보내 페이지 수/경계 확인
4. 페이지가 넘치면 `build.py` 의 `PAGES`(페이지별 섹션 배치) 또는 `STYLE`/`WINE_STYLE`(글자 크기, 간격) 조정

자세한 배경과 주의사항은 `CLAUDE.md` 참고.
