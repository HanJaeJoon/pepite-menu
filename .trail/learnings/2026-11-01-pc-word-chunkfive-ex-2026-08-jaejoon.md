---
title: "이 PC 빌드 환경 - Word 와 ChunkFive Ex 폰트 모두 2026-08-17 설치 완료. export_pdf.ps1 + pypdf 로 페이지 검증…"
author: Jaejoon
date: 2026-11-01
id: l-40385d29
scope: repo
status: active
touches:
  - tools/export_pdf.ps1
  - build.py
---

## Decided
이 작업 머신(jj272 데스크톱)의 빌드 환경 이력:

- MS Word: 원래 없었으나 2026-08-17 사용자가 설치. `tools/export_pdf.ps1` 정상 동작.
- `ChunkFive Ex` 폰트: 원래 없었으나 2026-08-17 사용자 폰트로 설치
  (`%LOCALAPPDATA%\Microsoft\Windows\Fonts\Chunkfive Ex.ttf` + HKCU 레지스트리).
  원본 TTF 와 OFL 1.1 라이선스는 저장소 `fonts/` 에 동봉했다.
- `바탕`: Windows 기본, 설치됨.

**Why:** 이제 이 PC 에서 빌드 -> PDF 내보내기 -> 페이지 검증까지 전부 가능하다.

## Watch out
레이아웃을 바꾸면 `python build.py` 후 `export_pdf.ps1` 을 돌리고,
pypdf 로 페이지 수(요리 4, 런치 2, 코스 1, 와인 7)와 임베드 폰트 목록에
`ChunkFive Ex` 가 있는지 확인할 것. 페이지 나눔이 명시적이라 넘침은 페이지 수
증가로 드러난다.

_From Claude Code's memory of this repo, last updated Aug 17._
