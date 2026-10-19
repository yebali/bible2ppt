"""성경 책 목록, 본문 입력("왕상19:1-5") 해석, 번역본 데이터에서 절 찾기."""
import json
import re
import unicodedata
from functools import lru_cache
from pathlib import Path

# (사이트 책 코드, 이름, 약어). 순서가 성경 순서.
BOOKS = [
    ("gen", "창세기", "창"), ("exo", "출애굽기", "출"), ("lev", "레위기", "레"), ("num", "민수기", "민"),
    ("deu", "신명기", "신"), ("jos", "여호수아", "수"), ("jdg", "사사기", "삿"), ("rut", "룻기", "룻"),
    ("1sa", "사무엘상", "삼상"), ("2sa", "사무엘하", "삼하"), ("1ki", "열왕기상", "왕상"), ("2ki", "열왕기하", "왕하"),
    ("1ch", "역대상", "대상"), ("2ch", "역대하", "대하"), ("ezr", "에스라", "스"), ("neh", "느헤미야", "느"),
    ("est", "에스더", "에"), ("job", "욥기", "욥"), ("psa", "시편", "시"), ("pro", "잠언", "잠"),
    ("ecc", "전도서", "전"), ("sng", "아가", "아"), ("isa", "이사야", "사"), ("jer", "예레미야", "렘"),
    ("lam", "예레미야애가", "애"), ("ezk", "에스겔", "겔"), ("dan", "다니엘", "단"), ("hos", "호세아", "호"),
    ("jol", "요엘", "욜"), ("amo", "아모스", "암"), ("oba", "오바댜", "옵"), ("jnh", "요나", "욘"),
    ("mic", "미가", "미"), ("nam", "나훔", "나"), ("hab", "하박국", "합"), ("zep", "스바냐", "습"),
    ("hag", "학개", "학"), ("zec", "스가랴", "슥"), ("mal", "말라기", "말"),
    ("mat", "마태복음", "마"), ("mrk", "마가복음", "막"), ("luk", "누가복음", "눅"), ("jhn", "요한복음", "요"),
    ("act", "사도행전", "행"), ("rom", "로마서", "롬"), ("1co", "고린도전서", "고전"), ("2co", "고린도후서", "고후"),
    ("gal", "갈라디아서", "갈"), ("eph", "에베소서", "엡"), ("php", "빌립보서", "빌"), ("col", "골로새서", "골"),
    ("1th", "데살로니가전서", "살전"), ("2th", "데살로니가후서", "살후"), ("1ti", "디모데전서", "딤전"),
    ("2ti", "디모데후서", "딤후"), ("tit", "디도서", "딛"), ("phm", "빌레몬서", "몬"), ("heb", "히브리서", "히"),
    ("jas", "야고보서", "약"), ("1pe", "베드로전서", "벧전"), ("2pe", "베드로후서", "벧후"), ("1jn", "요한일서", "요일"),
    ("2jn", "요한이서", "요이"), ("3jn", "요한삼서", "요삼"), ("jud", "유다서", "유"), ("rev", "요한계시록", "계"),
]
_BY_NAME = {n: i for i, (_, name, abbr) in enumerate(BOOKS) for n in (name, abbr)}
# 정리된 입력: 책 시작장[:시작절 | C] [-(끝장:끝절 | 끝장C | 숫자)]. C는 '장/편'으로 쓴 장 번호 표시
_REF = re.compile(r"(\D+?)(\d+)(?:C|:(\d+))?(?:-(\d+)(?:(C)|:(\d+))?)?")
_SEP = re.compile(r"[,;]")  # 여러 본문 구분 (쉼표나 세미콜론)
# 절 키: "3", "18-19", 반절 "2상"/"2ㄱ"/"6b-11". 에스더·다니엘 추가본의 "1_2" 같은 키는 장·절로 못 찾으니 건너뜀
_KEY = re.compile(r"(\d+)\D?(?:-(\d+)\D?)?")


def _normalize(part):
    """공백을 없애고 한글 표기를 기호로: '1장 1절'→'1:1', '23편'·'3장'→'23C'·'3C'(장 번호), '~'→'-'.
    숫자 바로 뒤의 장·편·절만 바꿈 (책 이름 '시편'은 그대로)."""
    s = re.sub(r"\s", "", part).replace("~", "-")
    s = re.sub(r"(\d+)[장편](\d+)", r"\1:\2", s)
    s = re.sub(r"(\d+)절", r"\1", s)
    s = re.sub(r"(\d+)[장편]", r"\1C", s)
    return s.rstrip(":")


def parse(text):
    """'왕상19:1-5, 시23' -> [(책번호, 시작장, 시작절|None, 끝장, 끝절|None)]. 절이 None이면 장 처음/끝."""
    passages = []
    for part in _SEP.split(text):
        s = _normalize(part)
        if not s:
            continue
        m = _REF.fullmatch(s)
        if not m:
            raise ValueError(f'"{part.strip()}" 형식을 알 수 없어요. 예) 창1:1-5')
        name, c1, v1, x, x_is_chapter, y = m.groups()
        if name not in _BY_NAME:
            raise ValueError(f'"{name}" 책을 찾을 수 없어요')
        if any(n and int(n) == 0 for n in (c1, v1, x, y)):
            raise ValueError(f'"{part.strip()}" 장·절 번호는 1부터예요')
        c1, v1 = int(c1), v1 and int(v1)
        if y:
            c2, v2 = int(x), int(y)
        elif x_is_chapter:  # '창1:1-3장': 3장 끝까지
            c2, v2 = int(x), None
        elif x and v1:
            c2, v2 = c1, int(x)
        elif x:
            c2, v2 = int(x), None
        else:
            c2, v2 = c1, v1
        if (c2, v2 or 999) < (c1, v1 or 0):
            raise ValueError(f'"{part.strip()}" 범위의 앞뒤가 바뀌었어요')
        passages.append((_BY_NAME[name], c1, v1, c2, v2))
    if not passages:
        raise ValueError("본문을 입력하세요. 예) 창1:1-5")
    return passages


def last_part_start(text):
    """입력 중인 마지막 본문이 시작하는 위치 (마지막 쉼표·세미콜론 다음)."""
    return max((m.end() for m in _SEP.finditer(text)), default=0)


def suggest_books(text):
    """입력 중인 마지막 본문의 책 이름 자동완성 후보 (책번호 목록). 숫자를 치기 시작하면 []."""
    s = re.sub(r"\s", "", text[last_part_start(text):])
    if not s or re.search(r"\d", s):
        return []
    typed = unicodedata.normalize("NFD", s)  # 자모로 풀어서 비교: "여"도 "열왕기상"에 걸림
    hits = [i for i, (_, name, abbr) in enumerate(BOOKS)
            if any(unicodedata.normalize("NFD", n).startswith(typed) for n in (name, abbr))]
    return sorted(hits, key=lambda i: s not in BOOKS[i][1:])  # 이름·약어가 정확히 같은 책을 맨 앞에


def typing_book_chapter(text):
    """입력 중인 마지막 본문이 '책' 또는 '책+장'까지면 (책번호, 장|None), 아니면 None. 안내 문구용."""
    m = re.fullmatch(r"(\D+?)(\d+)?C?", _normalize(text[last_part_start(text):]))
    if m and m.group(1) in _BY_NAME:
        return _BY_NAME[m.group(1)], m.group(2) and int(m.group(2))
    return None


# 성경 파일 이름(영문 코드, 대한성서공회 사이트의 번역본 코드) -> 화면에 보이는 이름.
# 파일 이름을 영문으로 두는 건 맥에서 만든 한글 파일 이름이 윈도우에서 깨질 수 있어서
VERSION_NAMES = {
    "GAE": "개역개정", "HAN": "개역한글", "SAE": "표준새번역", "SAENEW": "새번역",
    "COG": "공동번역", "COGNEW": "공동번역개정판", "CEV": "CEV",
}


class Bible:
    """bibles/<코드>.json 번역본."""

    def __init__(self, path, name):
        self.name = name
        self.data = json.loads(Path(path).read_text(encoding="utf-8"))

    def chapters(self, book):
        return len(self.data.get(BOOKS[book][0], {}))

    def chapter(self, book, c):
        """[(시작절, 끝절, 절키, 본문)] 페이지 순서 그대로. 없는 장이면 None."""
        ch = self.data.get(BOOKS[book][0], {}).get(str(c))
        if ch is None:
            return None
        out = []
        for key, text in ch.items():
            if m := _KEY.fullmatch(key):
                out.append((int(m[1]), int(m[2] or m[1]), key, text))
        return out

    def last_verse(self, book, c):
        ch = self.chapter(book, c)
        return ch and max(e for _, e, _, _ in ch)

    def verses(self, passage):
        """passage -> [(책번호, 장, 절키, 본문)]"""
        book, c1, v1, c2, v2 = passage
        name = BOOKS[book][1]
        out = []
        for c in range(c1, c2 + 1):
            ch = self.chapter(book, c)
            if ch is None:
                raise ValueError(f"{name} 책은 {self.chapters(book)}장까지 있어요 ({self.name})")
            last = max(e for _, e, _, _ in ch)
            lo = v1 if c == c1 and v1 else 1
            hi = v2 if c == c2 and v2 else last
            if max(lo, hi) > last:
                raise ValueError(f"{name} {c}장은 {last}절까지 있어요 ({self.name})")
            out += [(book, c, key, text) for s, e, key, text in ch if e >= lo and s <= hi]
        if not out:  # 번역본에서 빠진 절 (새번역 마17:21 등)
            ref = f"{c1}:{v1}" + (f"-{c2}:{v2}" if c2 != c1 else f"-{v2}" if v2 != v1 else "")
            raise ValueError(f"{name} {ref}은(는) {self.name}에 없는 절이에요. 다른 번역본을 골라 보세요")
        return out


def _files(folder):
    """{화면 이름: 파일}. 아는 코드(GAE 등)는 한글 이름으로, 그 밖에 직접 넣은 파일은 파일 이름 그대로
    (한글 이름이 자모로 분리(NFD)돼 있어도 NFC로 맞춤)."""
    return {VERSION_NAMES.get(p.stem.upper(), unicodedata.normalize("NFC", p.stem)): p
            for p in Path(folder).glob("*.json")}


def versions(folder):
    """번역본 이름 목록 (개역개정이 있으면 맨 앞)."""
    return sorted(sorted(_files(folder)), key=lambda n: n != "개역개정")


@lru_cache(maxsize=None)
def load(folder, name):
    path = _files(folder).get(name)
    if path is None:
        raise FileNotFoundError(f"{folder} 폴더에 {name} 성경 파일이 없어요")
    return Bible(path, name)


def slide_items(bible, text):
    """입력 -> [(책 약어 "왕상", 장:절 "19:2", 본문)]"""
    return [(BOOKS[book][2], f"{c}:{key}", body)
            for passage in parse(text) for book, c, key, body in bible.verses(passage)]
