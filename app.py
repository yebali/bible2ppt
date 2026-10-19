"""성경 자막 PPT 만들기 - 화면 (방송 모니터형)."""
import json
import os
import re
import subprocess
import sys
import tkinter as tk
from functools import lru_cache
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from PIL import ImageTk

import bible
import ppt

def desktop():
    """바탕화면 폴더. 윈도우는 OneDrive로 옮겨진 바탕화면도 실제 위치로 찾음."""
    if sys.platform == "win32":
        import ctypes
        buf = ctypes.create_unicode_buffer(260)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) == 0:  # CSIDL_DESKTOPDIRECTORY
            return Path(buf.value)
    return Path.home() / "Desktop"


BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
BIBLES = BASE / "bibles"
SETTINGS = BASE / "settings.json"
DEFAULT_TEMPLATE = "template.pptx"
DEFAULTS = {
    "version": "개역개정", "max_lines": ppt.DEFAULT_LINES,
    "template": DEFAULT_TEMPLATE, "out_dir": str(desktop() / "성경자막"), "recent": [],
}
C = {"console": "#2b2f36", "panel": "#353a42", "bezel": "#1c1f24", "line": "#4a505a", "text": "#e9ecef",
     "muted": "#9aa3ad", "tally": "#f4b740", "red": "#e5484d", "sel": "#4a4029"}
EXAMPLE = "예) 창1:1-5    창세기 1:1-5    시23    창1:30-2:3    왕상19:1-5, 약5:17"
GAP = 6  # 필름 띠 그림 사이 간격


def open_file(path):
    if sys.platform == "win32":
        os.startfile(path)
    else:
        subprocess.run(["open", str(path)])


# '이름'이거나 '이름 2'처럼 공백 뒤가 숫자인 문서만 닫음 ('이름 묵상노트' 같은 다른 문서는 그대로)
KEYNOTE_CLOSE = '''on run argv
    set stem to item 1 of argv
    if application "Keynote" is not running then return
    tell application "Keynote"
        repeat with d in (get every document)
            set n to name of d
            if n is stem then
                close d saving no
            else if n starts with (stem & " ") then
                try
                    (text ((length of stem) + 2) thru -1 of n) as integer
                    close d saving no
                end try
            end if
        end repeat
    end tell
end run'''


def close_open_copy(path):
    """같은 파일이 PowerPoint(윈도우)·Keynote(맥)에 이미 열려 있으면 닫음. 새로 만든 파일이 한 번만 열리게.
    열린 내용은 저장하지 않고 닫음 (어차피 새로 만든 파일로 바뀜)."""
    p = str(Path(path).resolve())
    try:
        if sys.platform == "win32":
            # ponytail: OneDrive 바탕화면처럼 경로가 URL로 잡히면 못 찾음 -> 그때는 '이름 (2).pptx'로 저장됨
            ps = ("try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { exit }; "
                  f"foreach ($x in @($a.Presentations)) {{ if ($x.FullName -eq '{p.replace(chr(39), chr(39) * 2)}') {{ $x.Close() }} }}")
            subprocess.run(["powershell", "-NoProfile", "-Command", ps], timeout=20,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        elif sys.platform == "darwin":  # Keynote는 pptx를 열 때마다 '이름', '이름 2'... 새 문서로 가져옴
            subprocess.run(["osascript", "-e", KEYNOTE_CLOSE, Path(path).stem], timeout=20, capture_output=True)
    except (OSError, subprocess.SubprocessError):
        pass  # 못 닫아도 저장은 계속 (윈도우에서 잠겨 있으면 새 이름으로 저장)


def file_name(text):
    return re.sub(r'[\\/*?"<>|\s]', "", text.replace(":", "_").replace(",", "+"))[:80] + ".pptx"


def resolve(path):
    path = Path(path)
    return path if path.is_absolute() else BASE / path


def template_path(p):
    """템플릿 경로 (상대 경로는 프로그램 폴더 기준). 기본 템플릿이 없으면 새로 만듦."""
    path = resolve(p or DEFAULT_TEMPLATE)
    if not path.exists() and path == BASE / DEFAULT_TEMPLATE:
        ppt.make_default_template(path)
    return path


@lru_cache(maxsize=1)
def load_template(path, mtime):
    """수정 시각(mtime)이 바뀌면 다시 읽음 (PowerPoint에서 템플릿을 고친 경우)."""
    try:
        return ppt.Template(path)
    except ValueError:
        raise
    except Exception as e:  # 손상된 파일 등
        raise ValueError(f"템플릿을 읽지 못했어요 ({path.name}): {e}") from e


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = self.load_settings()
        self.versions = bible.versions(BIBLES)
        if not self.versions:
            messagebox.showerror("성경 데이터 없음", f"{BIBLES} 폴더에 성경 파일(.json)이 없어요.")
            root.destroy()
            return
        if self.cfg["version"] not in self.versions:
            self.cfg["version"] = self.versions[0]
        self.slides, self.thumbs, self.current, self.tpl, self.note = [], [], 0, None, ""
        self._after, self._suppress, self._save_warned = None, False, False
        self.popup_items, self.popup_mode = [], ""
        self.k = max(1.0, root.winfo_fpixels("1i") / 96)  # 윈도우 화면 배율(125%, 150%)에 맞춤
        self.thumb_h, self.thumb_w, self._resize_job = 0, 0, None  # 썸네일 크기는 필름 띠 높이에 맞춰 정해짐
        fam = next((f for f in ("맑은 고딕", "Malgun Gothic", "Apple SD Gothic Neo") if f in tkfont.families()),
                   "TkDefaultFont")
        self.ui, self.ui_b, self.big = (fam, 10), (fam, 10, "bold"), (fam, 17, "bold")
        self.style()
        self.build()
        self.on_type()

    # ───────── 화면 구성 ─────────

    def style(self):
        st = ttk.Style(self.root)
        st.theme_use("clam")
        st.configure(".", background=C["panel"], foreground=C["text"], font=self.ui, bordercolor=C["line"],
                     lightcolor=C["panel"], darkcolor=C["panel"], troughcolor=C["console"],
                     fieldbackground=C["bezel"], insertcolor=C["text"], arrowcolor=C["text"])
        st.map("TCombobox", fieldbackground=[("readonly", C["bezel"])], foreground=[("readonly", C["text"])],
               selectbackground=[("readonly", C["bezel"])], selectforeground=[("readonly", C["text"])])
        for k, v in (("background", C["bezel"]), ("foreground", C["text"]),
                     ("selectBackground", C["sel"]), ("font", self.ui)):
            self.root.option_add(f"*TCombobox*Listbox.{k}", v)
        st.configure("TButton", padding=(12, 6))
        st.map("TButton", background=[("active", "#454b55")])
        st.configure("Tally.TButton", background=C["tally"], foreground="#231a05", font=self.ui_b,
                     padding=(18, 8), borderwidth=0)
        st.map("Tally.TButton", background=[("active", "#ffc75a")])
        st.configure("Strip.Horizontal.TScrollbar", background=C["line"], troughcolor=C["console"],
                     bordercolor=C["console"], lightcolor=C["line"], darkcolor=C["line"],
                     arrowcolor=C["muted"], gripcount=0, arrowsize=12)
        st.map("Strip.Horizontal.TScrollbar", background=[("active", C["muted"])])

    def build(self):
        r, k = self.root, self.k
        r.title("성경 자막 PPT 만들기")
        r.configure(bg=C["console"])
        r.geometry(f"{round(1100 * k)}x{round(720 * k)}")
        r.minsize(round(760 * k), round(520 * k))

        top = tk.Frame(r, bg=C["panel"], padx=12, pady=10)
        top.pack(fill="x")
        self.version = ttk.Combobox(top, values=self.versions, state="readonly", width=13)
        self.version.set(self.cfg["version"])
        self.version.bind("<<ComboboxSelected>>", lambda e: self.option_changed())
        self.version.pack(side="left", ipady=3, padx=(0, 12))
        ttk.Button(top, text="PPT 만들기", style="Tally.TButton", command=self.make).pack(side="right")
        self.query = tk.StringVar()
        self.entry = tk.Entry(top, textvariable=self.query, font=self.big, bg=C["bezel"], fg=C["text"],
                              insertbackground=C["text"], relief="flat", highlightthickness=1,
                              highlightbackground=C["line"], highlightcolor=C["tally"])
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 12), ipady=5)

        self.msg = tk.Label(r, bg=C["console"], font=self.ui, anchor="w", justify="left", padx=14, pady=5)
        self.msg.pack(fill="x")
        self.msg.bind("<Configure>", lambda e: self.msg.config(wraplength=max(100, e.width - 28)))  # 좁으면 줄바꿈

        # 창 크기에 맞춰 늘어나고, 경계선(사시)을 끌어서 미리보기·목록·필름 띠 크기를 조절
        pane = dict(bg=C["panel"], bd=0, sashwidth=6, opaqueresize=True)
        self.panes = tk.PanedWindow(r, orient="vertical", **pane)
        self.panes.pack(fill="both", expand=True, padx=12)
        upper = tk.PanedWindow(self.panes, orient="horizontal", **pane)
        mon = tk.Frame(upper, bg=C["bezel"], padx=10, pady=8)
        cap = tk.Frame(mon, bg=C["bezel"])
        cap.pack(fill="x", pady=(0, 6))
        tk.Label(cap, text="미리보기", bg=C["bezel"], fg=C["muted"], font=self.ui).pack(side="left")
        self.pos = tk.Label(cap, bg=C["bezel"], fg=C["tally"], font=self.ui_b)
        self.pos.pack(side="right")
        self.monitor = tk.Canvas(mon, bg=C["bezel"], highlightthickness=0, width=1, height=1)
        self.monitor.pack(fill="both", expand=True)
        self.list = tk.Listbox(upper, bg=C["bezel"], fg=C["text"], selectbackground=C["sel"],
                               selectforeground=C["text"], activestyle="none", relief="flat", highlightthickness=0,
                               font=self.ui, width=1, exportselection=False)
        self.list.bind("<<ListboxSelect>>", lambda e: self.list.curselection() and self.select(self.list.curselection()[0]))
        upper.add(mon, minsize=round(320 * k), stretch="always")
        upper.add(self.list, minsize=round(150 * k), width=round(320 * k), stretch="never")

        stripf = tk.Frame(self.panes, bg=C["console"])
        holder = tk.Frame(stripf, bg=C["console"], height=14)  # 스크롤바가 생겼다 없어져도 화면이 흔들리지 않게 자리 고정
        holder.pack(side="bottom", fill="x")
        holder.pack_propagate(False)
        self.bar = ttk.Scrollbar(holder, orient="horizontal", style="Strip.Horizontal.TScrollbar")
        self.strip = tk.Canvas(stripf, bg=C["console"], highlightthickness=0, height=1)
        self.strip.pack(fill="both", expand=True, pady=(4, 0))
        self.bar.configure(command=self.strip.xview)
        self.strip.configure(xscrollcommand=self.bar.set)
        self.strip.bind("<Button-1>", self.strip_click)
        self.strip.bind("<MouseWheel>", lambda e: self.strip.xview_scroll(-1 if e.delta > 0 else 1, "units"))
        self.panes.add(upper, minsize=round(200 * k), stretch="always")
        self.panes.add(stripf, minsize=round(60 * k), height=round(104 * k), stretch="never")
        for w in (self.monitor, self.strip):
            w.bind("<Configure>", lambda e: self.schedule_relayout())

        foot = tk.Frame(r, bg=C["console"], padx=12, pady=8)
        foot.pack(fill="x")
        ttk.Button(foot, text="설정", command=self.open_settings).pack(side="right")
        self.split_lbl = tk.Label(foot, bg=C["console"], fg=C["red"], font=self.ui)
        self.split_lbl.pack(side="right", padx=12)
        self.save_lbl = tk.Label(foot, bg=C["console"], fg=C["muted"], font=self.ui, anchor="w")
        self.save_lbl.pack(side="left")

        # 자동완성 목록: 입력칸 바로 아래에 겹쳐 띄움
        self.popup = tk.Listbox(r, bg=C["bezel"], fg=C["text"], selectbackground=C["sel"],
                                selectforeground=C["text"], activestyle="none", relief="flat",
                                highlightthickness=1, highlightbackground=C["tally"], font=(self.ui[0], 11),
                                exportselection=False)
        self.popup.bind("<ButtonPress-1>", self.click_popup)

        self.query.trace_add("write", lambda *a: self.on_type())
        e = self.entry
        e.bind("<Down>", lambda ev: self.move_popup(1))
        e.bind("<Up>", lambda ev: self.move_popup(-1))
        e.bind("<Return>", self.on_enter)
        e.bind("<Tab>", lambda ev: self.accept() if self.popup_items else None)
        e.bind("<Escape>", lambda ev: self.hide_popup())
        e.bind("<FocusOut>", lambda ev: r.after(150, self.hide_popup))
        e.bind("<FocusIn>", lambda ev: self.show_recent() if not self.query.get().strip() else None)
        r.bind("<Prior>", lambda ev: self.select(self.current - 1))  # PageUp
        r.bind("<Next>", lambda ev: self.select(self.current + 1))   # PageDown
        r.protocol("WM_DELETE_WINDOW", self.quit)
        e.focus_set()

    # ───────── 입력과 자동완성 ─────────

    def bible(self):
        return bible.load(BIBLES, self.version.get())

    def set_msg(self, text, color=None):
        self.msg.config(text=text, fg=color or C["muted"])

    def on_type(self):
        text = self.query.get()
        if self._suppress:
            self._suppress = False
        else:
            books = bible.suggest_books(text)[:8]
            items = [(f"{bible.BOOKS[i][1]}   {bible.BOOKS[i][2]}", i) for i in books]
            self.show_popup(items, "book") if items else self.hide_popup()
        self.set_msg(*self.hint(text))
        if self._after:
            self.root.after_cancel(self._after)
        self._after = self.root.after(400, lambda: self.preview(False))  # 치는 동안 잠깐 멈추면 미리보기

    def hint(self, text):
        """입력 중인 책/장의 범위 안내. 예) '열왕기상 19장: 1–21절'"""
        bc = bible.typing_book_chapter(text)
        if not bc:
            return (EXAMPLE,) if not text.strip() or bible.suggest_books(text) else ("",)
        b, (book, c) = self.bible(), bc
        name = bible.BOOKS[book][1]
        if c is None:
            return (f"{name}: 1–{b.chapters(book)}장",)
        last = b.last_verse(book, c)
        if last:
            return (f"{name} {c}장: 1–{last}절",)
        return (f"{name} 책은 {b.chapters(book)}장까지 있어요", C["red"])

    def show_popup(self, items, mode):
        self.popup_items, self.popup_mode = items, mode
        self.popup.delete(0, "end")
        for label, _ in items:
            self.popup.insert("end", "  " + label)
        self.popup.config(height=min(8, len(items)))
        self.popup.selection_set(0)
        self.popup.place(in_=self.entry, x=0, rely=1.0, y=2, relwidth=1.0)
        self.popup.lift()

    def hide_popup(self):
        self.popup_items = []
        self.popup.place_forget()

    def show_recent(self):
        if self.cfg["recent"]:
            self.show_popup([(t, t) for t in self.cfg["recent"]], "recent")

    def move_popup(self, d):
        if not self.popup_items:
            if d > 0 and not self.query.get().strip():
                self.show_recent()
            return "break"
        cur = self.popup.curselection()
        i = max(0, min(len(self.popup_items) - 1, (cur[0] if cur else -1) + d))
        self.popup.selection_clear(0, "end")
        self.popup.selection_set(i)
        self.popup.see(i)
        return "break"

    def click_popup(self, e):
        """누르는 순간 고름. 'break'로 목록이 포커스를 가져가지 않게 해서 이어 치는 글자가 입력칸으로 감."""
        self.popup.selection_clear(0, "end")
        self.popup.selection_set(self.popup.nearest(e.y))
        self.accept()
        return "break"

    def accept(self):
        if not self.popup_items:
            return None
        cur = self.popup.curselection()
        value, mode = self.popup_items[cur[0] if cur else 0][1], self.popup_mode
        self.hide_popup()
        self._suppress = True
        if mode == "book":  # 쉼표·세미콜론 뒤 마지막 본문의 책 이름만 약어로 바꿈
            text = self.query.get()
            head = text[:bible.last_part_start(text)]
            self.query.set(f"{head} {bible.BOOKS[value][2]}" if head else bible.BOOKS[value][2])
        else:
            self.query.set(value)
            self.preview(True)
        self.entry.focus_set()
        self.entry.icursor("end")
        return "break"

    def on_enter(self, _):
        if self.popup_items:  # 자동완성 목록이 떠 있으면 고르기만
            return self.accept()
        self.make()
        return "break"

    def option_changed(self):
        self.save_settings()
        if self.query.get().strip():
            self.preview(True)

    # ───────── 미리보기 ─────────

    def template(self):
        path = template_path(self.cfg["template"])
        if not path.exists():
            raise ValueError(f"템플릿 파일이 없어요: {path}")
        return load_template(path, path.stat().st_mtime)

    def preview(self, show_errors):
        if self._after:  # 기다리던 자동 미리보기는 취소
            self.root.after_cancel(self._after)
        self._after = None
        text = self.query.get()
        if not text.strip():
            self.clear_preview()
            return None
        try:
            tpl = self.template()
            items = bible.slide_items(self.bible(), text)
            slides = ppt.layout(items, tpl, self.cfg["max_lines"])
        except (ValueError, OSError) as e:
            self.clear_preview(str(e))  # 이전 본문이 그대로 남아 맞는 미리보기처럼 보이지 않게
            if show_errors:
                self.set_msg(str(e), C["red"])
            return None
        self.show(slides, tpl)
        return slides

    def clear_preview(self, note=""):
        """미리보기를 비우고, 이유가 있으면 미리보기 칸 가운데에 보여 줌."""
        self.slides, self.note = [], note
        self.list.delete(0, "end")
        for lbl in (self.pos, self.split_lbl, self.save_lbl):
            lbl.config(text="")
        self.draw_thumbs()
        self.draw_monitor()

    def show(self, slides, tpl):
        self.slides, self.tpl = slides, tpl
        self.list.delete(0, "end")
        for s in slides:
            self.list.insert("end", f"  {s.verse}{s.piece}   {s.text}")
        self.draw_thumbs()
        split = list(dict.fromkeys(s.verse for s in slides if s.piece))
        more = " …" if len(split) > 4 else ""
        self.split_lbl.config(text=f"{len(split)}절이 나뉘었어요 ({', '.join(split[:4])}{more})" if split else "")
        self.save_lbl.config(text=f"저장: {resolve(self.cfg['out_dir']) / file_name(self.query.get())}")
        if self.cfg["max_lines"] > tpl.fit_lines:
            self.set_msg(f"템플릿 띠에는 한글 {tpl.fit_lines}줄까지 들어가서 {tpl.fit_lines}줄씩 나눴어요. "
                         f"더 넣으려면 템플릿에서 띠를 키우세요.", C["tally"])
        self.select(0)

    def select(self, i):
        if not self.slides:
            return
        self.current = max(0, min(i, len(self.slides) - 1))
        self.draw_monitor()
        self.pos.config(text=f"{self.current + 1}번째 장 / {len(self.slides)}장")
        self.list.selection_clear(0, "end")
        self.list.selection_set(self.current)
        self.list.see(self.current)
        self.highlight()

    # ───────── 창 크기에 맞춰 다시 그리기 ─────────

    def schedule_relayout(self):
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(120, self.relayout)  # 끄는 동안 멈추면 한 번만

    def relayout(self):
        self._resize_job = None
        if self.strip_thumb_h() != self.thumb_h:
            self.draw_thumbs()  # 필름 띠 높이가 바뀌면 썸네일 크기도 바뀜
        else:
            self.update_bar()  # 폭만 바뀌면 보이는 썸네일 개수만 달라짐
        self.draw_monitor()

    def aspect(self):
        return self.tpl.slide_w / self.tpl.slide_h if self.tpl else 16 / 9

    def draw_monitor(self):
        """미리보기를 칸 크기에 맞춰 (비율 유지) 가운데에 그림."""
        self.monitor.delete("all")
        cw, ch = self.monitor.winfo_width(), self.monitor.winfo_height()
        if not self.slides:
            if self.note:
                self.monitor.create_text(cw // 2, ch // 2, text=self.note, fill=C["muted"], font=self.ui,
                                         width=max(100, cw - 40), justify="center")
            return
        w = max(16, min(cw, int(ch * self.aspect())))
        self.mon_img = ImageTk.PhotoImage(ppt.render(self.slides[self.current], self.tpl, w))
        self.monitor.create_image(cw // 2, ch // 2, image=self.mon_img)

    def strip_thumb_h(self):
        """필름 띠 높이에 맞는 썸네일 높이."""
        return max(24, self.strip.winfo_height() - 8)

    @property
    def step(self):
        """필름 띠에서 썸네일 하나가 차지하는 폭 (그림 + 간격)."""
        return self.thumb_w + GAP

    def draw_thumbs(self):
        self.thumb_h = self.strip_thumb_h()
        self.thumb_w = round(self.thumb_h * self.aspect())
        self.strip.delete("all")
        self.thumbs = []
        badge = max(7, self.thumb_h // 9)
        for n, s in enumerate(self.slides):
            x = n * self.step
            self.thumbs.append(ImageTk.PhotoImage(ppt.render(s, self.tpl, self.thumb_w)))
            self.strip.create_image(x, 4, image=self.thumbs[-1], anchor="nw")
            self.strip.create_rectangle(x + 3, 7, x + 6 + badge * 3, 9 + badge * 2, fill="#000000", outline="",
                                        stipple="gray50")
            self.strip.create_text(x + 5 + badge * 1.5, 8 + badge, text=str(n + 1), fill="white",
                                   font=(self.ui[0], badge, "bold"))
        self.strip.configure(scrollregion=(0, 0, len(self.slides) * self.step, self.thumb_h + 8))
        self.update_bar()
        self.highlight()

    def highlight(self):
        """필름 띠에서 지금 장 테두리, 안 보이면 그 자리로 스크롤."""
        self.strip.delete("hl")
        if not self.slides:
            return
        x = self.current * self.step
        self.strip.create_rectangle(x - 2, 2, x + self.thumb_w + 1, self.thumb_h + 5,
                                    outline=C["tally"], width=2, tags="hl")
        total = len(self.slides) * self.step
        left, right = self.strip.xview()
        if not left <= x / total <= (x + self.thumb_w) / total <= right:
            self.strip.xview_moveto(max(0.0, x / total - (right - left) / 2))

    def update_bar(self):
        """썸네일이 화면 폭을 넘을 때만 가로 스크롤바를 보여 줌."""
        if len(self.slides) * self.step > self.strip.winfo_width():
            self.bar.pack(fill="x")
        else:
            self.bar.pack_forget()

    def strip_click(self, e):
        i = int(self.strip.canvasx(e.x) // self.step)
        if i < len(self.slides):
            self.select(i)

    # ───────── 만들기 ─────────

    def make(self):
        self.hide_popup()
        slides = self.preview(True)
        if not slides:
            return
        out = resolve(self.cfg["out_dir"])  # 상대 경로는 프로그램 폴더 기준 (실행 위치와 무관하게)
        path = out / file_name(self.query.get())
        try:
            out.mkdir(parents=True, exist_ok=True)
            saved = self.save_new(slides, path)
        except OSError as e:
            messagebox.showerror("저장하지 못했어요", str(e))
            return
        q = self.query.get().strip()
        self.cfg["recent"] = [q] + [t for t in self.cfg["recent"] if t != q][:19]
        self.save_settings()
        note = "" if saved == path else " (같은 이름 파일이 PowerPoint에서 열려 있어서 새 이름으로 저장)"
        self.set_msg(f"만들었어요: {saved.name} ({len(slides)}장){note}", C["tally"])
        open_file(saved)

    def save_new(self, slides, path):
        """같은 이름이 있으면 묻지 않고 덮어씀. PowerPoint에서 열려 있어 못 쓰면 '이름 (2).pptx'처럼 새 이름으로."""
        if path.exists():
            close_open_copy(path)
        prs = ppt.build(slides, self.tpl)  # 한 번만 만들고, 저장만 다른 이름으로 다시 시도
        for n in range(1, 100):
            p = path if n == 1 else path.with_name(f"{path.stem} ({n}){path.suffix}")
            try:
                prs.save(p)
                return p
            except PermissionError as e:
                if not p.exists():  # 없는 파일도 못 만들면 열린 파일 탓이 아니라 폴더에 쓸 수 없는 것
                    raise PermissionError(f"{path.parent} 폴더에 저장할 수 없어요 (쓰기 권한 없음). "
                                          f"설정에서 다른 저장 폴더를 고르세요.") from e
        raise PermissionError(f"{path.name} 이름의 파일들이 모두 열려 있어서 저장하지 못했어요.")

    # ───────── 설정 ─────────

    def load_settings(self):
        cfg = dict(DEFAULTS, recent=[])
        try:
            saved = json.loads(SETTINGS.read_text(encoding="utf-8"))
            # 형식이 맞는 값만 받음 (직접 고치다 잘못 넣은 값은 기본값으로)
            cfg.update({k: v for k, v in saved.items() if k in DEFAULTS and type(v) is type(DEFAULTS[k])})
        except (OSError, ValueError, AttributeError):
            pass  # 처음 실행이거나 파일이 깨졌으면 기본값
        return cfg

    def save_settings(self):
        self.cfg.update(version=self.version.get())
        tmp = SETTINGS.with_suffix(".tmp")
        try:
            tmp.write_text(json.dumps(self.cfg, ensure_ascii=False, indent=1), encoding="utf-8")
            os.replace(tmp, SETTINGS)  # 쓰는 도중 꺼져도 기존 설정 파일이 깨지지 않게
        except OSError as e:
            if not self._save_warned:  # 한 번만 알림. 프로그램은 계속 동작
                self._save_warned = True
                messagebox.showwarning("설정을 저장하지 못했어요",
                                       f"{SETTINGS.parent} 폴더에 쓸 수 없어서, 바꾼 설정이 다음 실행 때 유지되지 않아요.\n"
                                       f"프로그램 폴더를 바탕화면처럼 쓸 수 있는 곳으로 옮겨 주세요.\n\n({e})")

    def open_template(self, p):
        path = template_path(p)
        if path.exists():
            open_file(path)
        else:
            messagebox.showerror("템플릿 없음", f"파일이 없어요: {path}")

    def open_settings(self):
        w = tk.Toplevel(self.root, bg=C["panel"], padx=18, pady=16)
        w.title("설정")
        w.transient(self.root)
        w.resizable(False, False)
        lines = tk.StringVar(value=str(self.cfg["max_lines"]))
        tpl = tk.StringVar(value=self.cfg["template"])
        out = tk.StringVar(value=self.cfg["out_dir"])

        def row(n, label, widget, *buttons):
            tk.Label(w, text=label, bg=C["panel"], fg=C["text"], font=self.ui).grid(row=n, column=0, sticky="w", pady=6, padx=(0, 14))
            widget.grid(row=n, column=1, sticky="we")
            for k, b in enumerate(buttons):
                b.grid(row=n, column=2 + k, padx=(6, 0))

        def pick_template():
            if p := filedialog.askopenfilename(parent=w, filetypes=[("PowerPoint", "*.pptx")]):
                tpl.set(p)

        def pick_folder():
            if p := filedialog.askdirectory(parent=w):
                out.set(p)

        row(0, "한 장 최대 줄 수", ttk.Spinbox(w, from_=1, to=10, textvariable=lines, width=5))
        row(1, "템플릿 파일", ttk.Entry(w, textvariable=tpl, width=42), ttk.Button(w, text="찾기", command=pick_template),
            ttk.Button(w, text="PowerPoint로 열기", command=lambda: self.open_template(tpl.get().strip())))
        row(2, "저장 폴더", ttk.Entry(w, textvariable=out, width=42), ttk.Button(w, text="변경", command=pick_folder))
        tk.Label(w, text="배경색, 띠 색과 투명도, 글꼴, 글자 크기는 템플릿 파일을 PowerPoint에서 직접 고쳐요.\n"
                         "{본문} 글자는 지우지 마세요.", bg=C["panel"], fg=C["muted"], font=self.ui,
                 justify="left").grid(row=3, column=0, columnspan=4, sticky="w", pady=(10, 14))

        def save():
            try:
                n = max(1, min(10, int(lines.get())))
            except ValueError:
                n = DEFAULTS["max_lines"]
            self.cfg.update(max_lines=n, template=tpl.get().strip() or DEFAULT_TEMPLATE,
                            out_dir=out.get().strip() or DEFAULTS["out_dir"])
            self.save_settings()
            w.destroy()
            if self.query.get().strip():
                self.preview(True)

        btns = tk.Frame(w, bg=C["panel"])
        btns.grid(row=5, column=0, columnspan=4, sticky="e")
        ttk.Button(btns, text="취소", command=w.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="저장", style="Tally.TButton", command=save).pack(side="right")
        w.grab_set()

    def quit(self):
        self.save_settings()
        for job in (self._after, self._resize_job):  # 예약된 작업 정리 (닫은 뒤 오류 방지)
            if job:
                self.root.after_cancel(job)
        self.root.destroy()


if __name__ == "__main__":
    if sys.platform == "win32":
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 흐릿하게 확대되지 않도록
        except (AttributeError, OSError):
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()
