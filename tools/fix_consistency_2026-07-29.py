# -*- coding: utf-8 -*-
"""메뉴 데이터 정합성 일괄 수정. data/*.json 을 제자리에서 고친다."""
import json, os, re

DATA = r"C:\Users\jj272\OneDrive\문서\카카오톡 받은 파일\페피트 메뉴판_새버전\data"
log = []


def load(n):
    return json.load(open(os.path.join(DATA, n + '.json'), encoding='utf-8'))


def save(n, d):
    json.dump(d, open(os.path.join(DATA, n + '.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)


def items(d):
    for s in d['sections']:
        for it in s['items']:
            yield s['name'], it


def find(d, sec, kor):
    for s, it in items(d):
        if s == sec and it['kor'] == kor:
            return it
    raise KeyError((sec, kor))


def setf(it, field, new, why):
    old = it[field]
    if old == new:
        return
    it[field] = new
    log.append(f'{why}: {field} {old!r} -> {new!r}')


food, lunch, wine = load('food'), load('lunch'), load('wine')

# ---------------------------------------------------------------- 요리 메뉴판
setf(find(food, 'SIDE', '올리브, 하몽'), 'eng', 'Olives, Jamón',
     '요리/SIDE 올리브,하몽 - 영문명이 깔라마리(Chipirones)로 잘못 붙어 있었음')
setf(find(food, 'SIDE', '트러플감자튀김'), 'eng', 'Truffle Fries',
     '요리/SIDE 트러플감자튀김 - 영문명이 고등어초밥(Mackerel Pressed Suship)이었음')
setf(find(food, 'STARTER', '잿방어 세비체'), 'eng', 'Aged Amberjack Ceviche',
     '요리/STARTER 잿방어 - Gray Mullet(숭어)은 오역, 설명문 Amberjack 기준')
setf(find(food, 'PASTA', '트러플'), 'eng', 'Truffle',
     '요리/PASTA 트러플 - 같은 섹션 다른 항목은 Pasta 접미사 없음')
setf(find(food, 'STARTER', '우니 타르틀렛(2pieces)'), 'kor', '우니 타르틀렛 (2pieces)',
     '요리/STARTER 우니 타르틀렛 - 괄호 앞 띄어쓰기 통일')
setf(find(food, 'DESSERT', '자몽 타르틀렛(2pieces)'), 'kor', '자몽 타르틀렛 (2pieces)',
     '요리/DESSERT 자몽 타르틀렛 - 괄호 앞 띄어쓰기 통일')
setf(find(food, 'PASTA', '화이트 라구'), 'desc_kor', '화이트라구, 화이트포트와인, 포르치니버섯',
     '요리/PASTA 화이트 라구 - 쉼표 뒤 공백')
it = find(food, 'MAIN', '해산물 플레터')
setf(it, 'desc_kor', it['desc_kor'].replace('PEPITE', 'PÉPITE'),
     '요리/MAIN 해산물 플레터 - 브랜드 표기를 PÉPITE 로 통일')
setf(it, 'desc_eng', it['desc_eng'].replace('PEPITE', 'PÉPITE'),
     '요리/MAIN 해산물 플레터 - 브랜드 표기를 PÉPITE 로 통일')
it = find(food, 'STARTER', '무화과 카프레제')
setf(it, 'desc_kor', it['desc_kor'].replace('꼼테치즈', '꽁떼치즈').replace('토사즈크림', '토사즈 크림'),
     '요리/STARTER 무화과 카프레제 - 꼼테/꽁떼, 토사즈크림 표기 통일')

# ---------------------------------------------------------------- 런치 메뉴판
it = find(lunch, 'STARTER', '잿방어세비체')
setf(it, 'kor', '잿방어 세비체', '런치/STARTER 잿방어 - 띄어쓰기')
setf(it, 'eng', 'Aged Amberjack Ceviche', '런치/STARTER 잿방어 - Yellowtail(방어) -> Amberjack(잿방어)')

it = find(lunch, 'STARTER', '한우 비프타르타르')
setf(it, 'kor', '한우 비프 타르타르', '런치/STARTER 한우 타르타르 - 요리 메뉴판과 띄어쓰기 통일')
setf(it, 'eng', 'Hanwoo Beef Tartare',
     '런치/STARTER 한우 타르타르 - Hanwoo 자체가 한우이므로 Korean 중복 제거')

setf(find(lunch, 'STARTER', '무화과 카프레제'), 'eng', 'Fig Caprese',
     '런치/STARTER 무화과 카프레제 - Figs -> Fig')
setf(find(lunch, 'STARTER', '전복 아에모노'), 'eng', 'Abalone Aemono',
     '런치/STARTER 전복 아에모노 - 대소문자')
setf(find(lunch, 'DESSERT', '마스카포네크림, 바닐라아이스크림'), 'eng',
     'Mascarpone Cream & Vanilla Ice Cream',
     '런치/DESSERT 마스카포네크림 - 요리 메뉴판과 영문명 통일')
setf(find(lunch, 'STARTER', '시저 샐러드'), 'desc_kor',
     '로메인, 곡물후레이크, 베이컨크럼블, 레몬오일',
     '런치/STARTER 시저 샐러드 - 쉼표 뒤 공백')

it = find(lunch, 'PASTA', '홍새우 파스타')
setf(it, 'desc_kor', '홍새우비스크, 홍새우, 스트라치아텔라',
     '런치/PASTA 홍새우 - 쉼표 뒤 공백 + 요리 메뉴판과 동일하게 피스타치오 삭제')
setf(it, 'desc_eng', 'Red Shrimp Bisque, Red Shrimp, Stracciatella',
     '런치/PASTA 홍새우 - Pistachio 삭제')

it = find(lunch, 'STARTER', '무화과 카프레제')
setf(it, 'desc_kor', it['desc_kor'].replace('꼼테치즈', '꽁떼치즈').replace('토사즈 크림', '토사즈 크림'),
     '런치/STARTER 무화과 카프레제 - 치즈 표기 통일')

# ---------------------------------------------------------------- 와인 메뉴판
WINE_TERMS = [
    ('샤도네이', '샤르도네', '품종 표기 통일 (샤도네이 -> 샤르도네)'),
    ('쇼비뇽', '소비뇽', '품종 표기 통일 (쇼비뇽 -> 소비뇽)'),
    ('페트릭 자비에', '패트릭 자비에', '생산자 Patrick Javillier 표기 통일'),
    ('피노뫼니에', '피노 뫼니에', '품종 표기 통일'),
    ('뮈니에', '뫼니에', '품종 표기 통일 (Meunier -> 뫼니에)'),
    ('샹볼 뮈니지', '샹볼 뮈지니', 'Chambolle-Musigny 오타'),
]
WINE_ENG = [
    ('Schönheits ', 'Schoenheitz ', '생산자 Schoenheitz 영문 표기 통일'),
    ('Schueller-Heitz ', 'Schoenheitz ', '생산자 Schoenheitz 영문 표기 통일'),
    ('Champagne  ', 'Champagne ', '겹공백 정리'),
]
seen = set()
for sec, it in items(wine):
    for f in ('kor', 'eng', 'region', 'desc_kor', 'desc_eng'):
        v = it.get(f)
        if not v:
            continue
        new = v
        for a, b, why in WINE_TERMS:
            if a in new:
                new = new.replace(a, b)
                if why not in seen:
                    seen.add(why)
                    log.append(f'와인 - {why}')
        for a, b, why in WINE_ENG:
            if a in new:
                new = new.replace(a, b)
                if why not in seen:
                    seen.add(why)
                    log.append(f'와인 - {why}')
        new = new.replace('“', '"').replace('”', '"').replace('’', "'")
        new = re.sub(r'\s{2,}', ' ', new).strip()
        it[f] = new
log.append('와인 - 곡선 따옴표(" ")를 곧은 따옴표로 통일, 겹공백 제거')

save('food', food); save('lunch', lunch); save('wine', wine)
print('총 %d건 수정\n' % len(log))
for l in log:
    print(' -', l)
