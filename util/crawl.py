"""대한성서공회 성경읽기 페이지에서 번역본별 본문을 받아 bibles/<코드>.json 으로 저장 (예: bibles/GAE.json).

사용법: 프로젝트 폴더에서 python -m util.crawl [버전코드 ...]   (예: python -m util.crawl GAE HAN, 생략 시 전체)
저장 형식: {"gen": {"1": {"1": "태초에 ...", ...}, ...}, ...}
"""
import html
import json
import re
import sys
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import bible

URL = "https://www.bskorea.or.kr/bible/korbibReadpage.php?version={v}&book={b}&chap={c}"
BOOKS = [code for code, _, _ in bible.BOOKS]  # 번역본마다 장 수가 달라서 빈 장이 나올 때까지 받는다

NUMBER = re.compile(r'<span class="number">([^&<]+)(?:&nbsp;)*</span>')
FOOTNOTE_MARK = re.compile(r"<a class=comment.*?</a>", re.S)
FOOTNOTE_BODY = re.compile(r"<div[^>]*class=D2.*?</div>", re.S)
TITLE = re.compile(r'<font class="smallTitle">.*?</font>(?=\s*<br)', re.S)  # 소제목 (안에 이름 태그가 섞여 있어도 끝까지)
HEADING = re.compile(r"<H VAL=[^>]*>.*?</font>", re.S | re.I)  # 공동번역·CEV의 다른 소제목 표시
TAG = re.compile(r"<[^>]+>")


def parse(page):
    """장 페이지 HTML -> {절번호: 본문}. 각주 표시/내용, 소제목은 버린다.
    절 본문 = 절 번호부터 다음 절 번호 직전까지 (사이트가 본문 일부를 </span> 뒤에 두는 경우가 있어서)."""
    start = page.find('id="tdBible1"')
    if start < 0:  # 점검·차단 페이지 등: '없는 장'으로 착각하지 않게 오류로 (fetch가 다시 시도)
        raise ValueError("성경 본문 페이지가 아니에요")
    end = page.find("preCheck", start)
    cut = page.rfind("<div", start, end)  # 아래 이전/다음 버튼 영역은 뺌
    block = page[start:cut if cut > start else end]
    block = HEADING.sub("", TITLE.sub("", FOOTNOTE_BODY.sub("", FOOTNOTE_MARK.sub("", block))))
    parts = NUMBER.split(block)  # [앞부분, 번호1, 본문1, 번호2, 본문2, ...]
    clean = lambda s: " ".join(html.unescape(TAG.sub("", s)).split())
    lead = parts[0].split('class="chapNum"', 1)[-1].split("</font>", 1)[-1]  # 장 제목('제 3 장') 뒤 첫 절 번호 앞
    verses = {}
    for num, body in zip(parts[1::2], parts[2::2]):
        num, text = num.strip(), clean(body)
        if not text and not verses:  # 1절 본문을 번호 앞(장 제목 바로 뒤)에 두는 페이지 (CEV 사27:1, 단11:1)
            text = clean(lead)
        if not text and verses:  # 번호만 있고 본문이 없으면 앞 절에 합쳐 번역된 절 (CEV 출21:10-11 등): '10-11'로
            prev = next(reversed(verses))
            verses[f"{prev.split('-')[0]}-{num}"] = verses.pop(prev)
        else:
            verses[num] = text
    return verses


def fetch(version, book, chap):
    """없는 장이면 {} (빈 페이지가 두 번 연속일 때), 네트워크 오류가 계속되면 예외."""
    empty, err = 0, "빈 페이지"
    for attempt in range(5):
        if attempt:
            time.sleep(2 ** (attempt - 1))
        try:
            with urllib.request.urlopen(URL.format(v=version, b=book, c=chap), timeout=20) as r:
                verses = parse(r.read().decode("utf-8"))
            if verses:
                return verses
            empty += 1
            if empty == 2:
                return {}
        except Exception as e:  # 네트워크 오류는 재시도
            err = e
    raise RuntimeError(f"{version} {book} {chap}: 본문을 받지 못함 ({err})")


def crawl_book(version, book):
    chapters = {}
    while verses := fetch(version, book, len(chapters) + 1):
        chapters[str(len(chapters) + 1)] = verses
        odd = [k for k in verses if not k.isdigit()]
        if odd:
            print(f"  참고: {version} {book} {len(chapters)}장 절 번호 {odd}")
    return chapters


def crawl(version):
    with ThreadPoolExecutor(4) as pool:  # 사이트 부담 줄이려고 동시 4개로 제한
        data = dict(zip(BOOKS, pool.map(lambda b: crawl_book(version, b), BOOKS)))
    out = Path(__file__).parent.parent / "bibles" / f"{version}.json"  # 프로젝트의 bibles/, 영문 파일 이름
    if out.exists():  # 사이트 오류로 책이 덜 받아졌으면 좋은 기존 파일을 덮어쓰지 않음
        old = json.loads(out.read_text(encoding="utf-8"))
        short = [f"{b} {len(data.get(b, {}))}/{len(old[b])}장" for b in old if len(data.get(b, {})) < len(old[b])]
        if short:
            raise SystemExit(f"{bible.VERSION_NAMES[version]}: 예전보다 장 수가 적은 책이 있어 저장하지 않았어요 ({', '.join(short)}). "
                             "사이트 상태를 확인하고 다시 실행하세요")
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    chapters = sum(len(ch) for ch in data.values())
    print(f"{bible.VERSION_NAMES[version]}: {chapters}장 {sum(len(v) for ch in data.values() for v in ch.values())}절 -> {out}")


if __name__ == "__main__":
    assert parse('<x id="tdBible1"><font class="smallTitle"><font class="name">다윗</font>의 시</font><br />'
                 '<span><span class="number">8&nbsp;</span><a class=comment>1)</a>구원은 <font size=\'1\'>내리</font>'
                 '소서 <div id=a class=D2 >승리는</div></font></span>이어진 문장<br /><font class="smallTitle">소제목</font>'
                 '<br /><span><span class="number">9&nbsp;</span></span><H VAL=1.1>유다</font><br /><span>'
                 '<span class="number">10&nbsp;</span>끝</span></div><div><a onclick="preCheck()">') \
        == {"8-9": "구원은 내리소서 이어진 문장", "10": "끝"}
    for v in sys.argv[1:] or bible.VERSION_NAMES:
        crawl(v)
