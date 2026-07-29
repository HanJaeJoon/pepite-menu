# PEPITE 메뉴판

메뉴 내용은 `data/*.json`, 레이아웃은 `build.py` 로 관리하고 docx/pdf 는 빌드해서 만든다.
docx 를 직접 손으로 고치면 정렬과 페이지 구성이 다시 깨지므로 하지 않는다.

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
| 페피트 런치 메뉴판 | 3 |
| 페피트 런치 코스메뉴 | 1 (A5) |
| 페피트 와인 메뉴판 | 7 |

## 메뉴를 바꾸려면

1. `data/food.json`, `data/lunch.json`, `data/wine.json`, `data/course.json` 중 해당 파일 수정
2. `python build.py`
3. PDF 로 내보내 페이지 수/경계 확인
4. 페이지가 넘치면 `build.py` 의 `PAGES`(페이지별 섹션 배치) 또는 `STYLE`/`WINE_STYLE`(글자 크기, 간격) 조정

자세한 배경과 주의사항은 `CLAUDE.md` 참고.
