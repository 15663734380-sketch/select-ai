#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
划词问答 —— 阅读文档时选中文字，小窗口即时给出 AI 解答。

工作方式：
  1. 全局鼠标监听：检测「拖选」或「双击选词」结束后松开左键的瞬间；
  2. 向当前前台应用发送 Ctrl+C，把选中的文字取到剪贴板（读完立刻还原原剪贴板）；
  3. 把选中文字发给 OpenAI 兼容接口（DeepSeek / 通义 / 智谱 / 月之暗面等），流式显示回答。

文档上下文：
  点「加载文档」选择 PDF / Word / 文本后，文档只在加载时解析一次并存在内存中；
  之后每次划词，会按选中文字自动检索文档中最相关的若干片段（默认最多约 3000 字）
  随问题一起发送，回答会优先依据文档内容。文档被外部修改保存后会自动重新解析。

依赖：pynput、pyperclip、requests、pypdf（PDF）、python-docx（Word）
      Pillow 为可选依赖，用于在顶栏显示图标里的角色形象，缺少时自动省略
运行：pythonw select_ai.py
"""

import ctypes
import json
import queue
import re
import sys
import threading
import time
import traceback
import tkinter as tk
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, ttk

import pyperclip
import requests
from pynput import keyboard, mouse

if getattr(sys, "frozen", False):
    # 打包成 exe 后：配置、历史、图标都放在 exe 同目录
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
LOG_PATH = APP_DIR / "error.log"
HISTORY_PATH = APP_DIR / "history.json"
HISTORY_MAX = 100    # 最多保留的历史问答条数
FOOTER_HIDE_H = 540  # 窗口高度低于此值（逻辑像素）时折叠底部操作栏
FOOTER_SHOW_H = 565  # 窗口高度回升到此值以上时恢复底部操作栏


def resource_path(name: str) -> Path:
    """图标等只读资源：优先程序目录，打包后回退到内部资源目录。"""
    p = APP_DIR / name
    if p.exists():
        return p
    base = getattr(sys, "_MEIPASS", None)
    if base:
        q = Path(base) / name
        if q.exists():
            return q
    return p

# 浅蓝配色（护眼）：取自应用图标的天蓝发梢与白色荷叶边
# 层次：淡蓝窗口底 → 更浅的近白圆角卡片（文字框）→ 深海军蓝文字
BG = "#dfeefb"           # 窗口底色：淡蓝
CARD = "#fbfdff"        # 文字框 / 卡片：比底色更浅，接近纯白
CARD_HOVER = "#eef6fe"
HEADER_L = "#c9e3fb"     # 顶栏渐变起点（浅天蓝）
HEADER_R = "#dfeefb"     # 顶栏渐变终点（与窗口底色衔接）
HEADER_LINE = "#b9d8f5"  # 顶栏下沿的分隔线
BORDER = "#b7d7f5"       # 淡蓝边框
FG = "#1f3a63"
TEXT_SOFT = "#33517c"
SUB = "#5b7ba6"
ACCENT = "#4189dd"
ACCENT_HOVER = "#3277c8"
OK = "#2ea77c"
WARN = "#c97a12"
SEL = "#cfe4fb"
GOLD = "#a8731f"
SKY = "#3a86d6"
FONT = "Microsoft YaHei UI"
MONO = "Consolas"

DOC_CHUNK = 700      # 文档分块大小（字符）
DOC_TOP_K = 4        # 每次最多引用的片段数
DOC_BUDGET = 50000    # 每次随问题附带的文档片段总字数上限

HEADING_RE = re.compile(
    r"^(#{1,6}\s|第[一二三四五六七八九十百零两\d]+[章节篇部]|chapter\s+\d)",
    re.IGNORECASE,
)

DEFAULT_CONFIG = {
    "base_url": "https://api.deepseek.com/v1",
    "model": "deepseek-flash",
    "api_key": "",
    "system_prompt": (
        "你是精炼的学习助手。用户会发来阅读时选中的一段文字，请立刻用中文解释："
        "它在讲什么、关键概念是什么、为什么重要；若是代码或公式，说明作用与原理。"
        "默认控制在 150 字以内，必要时用短分点，不要复述原文。"
    ),
    "max_chars": 2000,
    "doc_context_chars": DOC_BUDGET,
    "last_doc_path": "",
    "topmost": True,
    "window_pos": None,
}


def log(msg: str) -> None:
    try:
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    try:
        if CONFIG_PATH.exists():
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                cfg.update(data)
    except Exception:
        log("读取配置失败\n" + traceback.format_exc())
    # 旧模型名已在 DeepSeek 官方停用，自动迁移到当前模型
    if "deepseek.com" in str(cfg.get("base_url") or ""):
        if cfg.get("model") in ("deepseek-chat", "deepseek-reasoner", "deepseek-v4-flash"):
            cfg["model"] = "deepseek-flash"
    return cfg


def save_config(cfg: dict) -> None:
    try:
        CONFIG_PATH.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        log("保存配置失败\n" + traceback.format_exc())


def enable_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def clip_get() -> str:
    try:
        return pyperclip.paste() or ""
    except Exception:
        return ""


def clip_set(text: str) -> None:
    try:
        pyperclip.copy(text)
    except Exception:
        pass


def clipboard_seq() -> int:
    """剪贴板修改序号：每次内容被写入都会 +1，用于判断期间是否被别处改动。"""
    try:
        return int(ctypes.windll.user32.GetClipboardSequenceNumber())
    except Exception:
        return 0


def ctrl_or_alt_down() -> bool:
    """Ctrl / Alt 按下时不触发，避免误伤 Ctrl+拖拽等其它操作。"""
    try:
        u = ctypes.windll.user32
        return bool(u.GetAsyncKeyState(0x11) & 0x8000) or bool(
            u.GetAsyncKeyState(0x12) & 0x8000
        )
    except Exception:
        return False


# ---------------------------------------------------------------- 文档处理
def read_document(path: Path) -> str:
    """把 PDF / DOCX / 文本文件读成纯文本（在后台线程调用）。"""
    suffix = path.suffix.lower()
    if suffix in (".doc", ".ppt", ".xls"):
        raise ValueError("不支持旧版 Office 格式，请另存为 .docx / .pdf / .txt")
    if suffix == ".pdf":
        try:
            from pypdf import PdfReader
        except ImportError as e:
            raise RuntimeError("缺少 pypdf，请运行 pip install pypdf") from e
        reader = PdfReader(str(path))
        parts = []
        for page in reader.pages:
            try:
                parts.append(page.extract_text() or "")
            except Exception:
                parts.append("")
        return "\n\n".join(parts)
    if suffix == ".docx":
        try:
            import docx
        except ImportError as e:
            raise RuntimeError("缺少 python-docx，请运行 pip install python-docx") from e
        doc = docx.Document(str(path))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                line = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                if line:
                    parts.append(line)
        return "\n".join(parts)

    raw = path.read_bytes()
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        try:
            return raw.decode("utf-16")
        except UnicodeDecodeError:
            pass
    raise ValueError("无法识别文件编码，请转成 UTF-8 文本")


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def split_chunks(text: str, size: int = DOC_CHUNK) -> list:
    """按段落优先切块，单块不超过 size 字符。"""
    units = []
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        if len(para) <= size:
            units.append(para)
            continue
        for sent in re.split(r"(?<=[。！？；!?;])|\n", para):
            sent = sent.strip()
            while len(sent) > size:
                units.append(sent[:size])
                sent = sent[size:]
            if sent:
                units.append(sent)
    chunks, buf = [], ""
    for u in units:
        is_heading = len(u) < 80 and bool(HEADING_RE.match(u))
        if buf and (
            len(buf) + len(u) + 1 > size
            or (is_heading and len(buf) >= size // 2)
        ):
            chunks.append(buf)
            buf = u
        else:
            buf = (buf + "\n" + u).strip() if buf else u
    if buf:
        chunks.append(buf)
    return chunks


def _bigrams(s: str) -> set:
    """字符二元组，用于中英文混合的轻量相关性打分。"""
    s = re.sub(r"[\s\W_]+", "", s.lower())
    if len(s) < 2:
        return {s} if s else set()
    return {s[i : i + 2] for i in range(len(s) - 1)}


def retrieve_chunks(query: str, chunks, grams, top_k: int = DOC_TOP_K, budget: int = DOC_BUDGET) -> list:
    """挑出与选中文字最相关的片段，按原文顺序返回。"""
    if not chunks or budget <= 0:
        return []
    picked, total = [], 0
    q = _bigrams(query)
    if len(q) >= 4 and grams and len(grams) == len(chunks):
        scored = []
        for i, g in enumerate(grams):
            hit = len(q & g)
            if hit:
                scored.append((hit / len(q), i))
        scored.sort(key=lambda x: (-x[0], x[1]))
        # 只保留与最佳片段相关性相当的块，避免把弱相关的无关章节带进上下文
        cutoff = max(0.08, scored[0][0] * 0.3) if scored else 1.0
        for score, i in scored:
            if score < cutoff or len(picked) >= top_k:
                break
            c = chunks[i]
            if picked and total + len(c) > budget:
                continue
            picked.append(i)
            total += len(c)
        # 附带第一块，帮助模型把握文档主题（预算允许时）
        if picked and 0 not in picked and len(chunks) > 1 and total + len(chunks[0]) <= budget:
            picked.append(0)
    if not picked:  # 没有命中或选中太短：退回开头几块
        for i, c in enumerate(chunks):
            if len(picked) >= top_k:
                break
            if picked and total + len(c) > budget:
                continue
            picked.append(i)
            total += len(c)
    picked.sort()
    return [chunks[i] for i in picked]


class App:
    def __init__(self) -> None:
        self.cfg = load_config()
        self.listening = False
        self.gen = 0
        self.answer_text = ""
        self.last_text = ""
        self.last_text_at = 0.0
        self.last_click = (0.0, 0.0, 0.0)
        self.press = None
        self.press_inside = False
        self.win_rect = (0, 0, 0, 0)
        self.msg_q: "queue.Queue[tuple]" = queue.Queue()
        self._moved_by_self = False
        self.footer_hidden = False

        # 文档上下文（加载后常驻内存，提问时检索相关片段随问题发送）
        self.doc_lock = threading.Lock()
        self.doc_path = None
        self.doc_text = ""
        self.doc_chunks = ()
        self.doc_grams = ()
        self.doc_mtime = 0.0

        # 历史问答（每次成功解答后保存到 history.json）
        self.history = self._load_history()
        self.history_win = None
        self.current_question = ""

        self.root = tk.Tk()
        self.root.withdraw()  # 先隐藏，等界面与尺寸就绪后一次性显示
        self.root.title("划词问答")
        ico = resource_path("app.ico")
        if ico.exists():
            try:
                self.root.iconbitmap(str(ico))
            except Exception:
                pass
        self.root.configure(bg=BG)

        # 界面缩放系数（96 DPI = 1.0）：像素尺寸随屏幕缩放
        self.ui = 1.0
        try:
            self.ui = max(1.0, float(self.root.tk.call("tk", "scaling")) / (96 / 72))
        except Exception:
            pass
        self.head_h = int(56 * self.ui)
        self.brand_w = int(142 * self.ui)
        # 顶栏里的角色形象（取自应用图标；缺少 Pillow 时自动省略）
        self.mascot = None
        try:
            from PIL import Image, ImageTk

            src = Image.open(resource_path("app_mascot.png")).convert("RGBA")
            side = int(40 * self.ui)
            self.mascot = ImageTk.PhotoImage(src.resize((side, side), Image.LANCZOS))
        except Exception:
            self.mascot = None

        self.root.minsize(380, 400)
        self.root.geometry(self._initial_geometry())
        self.root.attributes("-topmost", bool(self.cfg.get("topmost", True)))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Unmap>", self._on_minimize)
        self.root.after(600, self._titlebar)

        self._build_ui()
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        need_w = max(430, min(self.root.winfo_reqwidth(), 560))
        need_h = min(self.root.winfo_reqheight(), sh - 100)
        pos = self.cfg.get("window_pos")
        if isinstance(pos, list) and len(pos) == 2:
            x, y = int(pos[0]), int(pos[1])
        else:
            x, y = sw - need_w - 40, 140
        x = max(20, min(x, sw - need_w - 20))
        y = max(20, min(y, sh - need_h - 60))
        self.root.geometry(f"{need_w}x{need_h}+{x}+{y}")
        self.root.update()  # 先把界面绘制完成，避免首帧空白小窗
        self.root.deiconify()
        self.root.lift()

        # 全局鼠标 & 热键监听
        self.mouse_listener = mouse.Listener(on_click=self._on_click)
        self.mouse_listener.daemon = True
        self.mouse_listener.start()

        self.hotkeys = None
        try:
            self.hotkeys = keyboard.GlobalHotKeys(
                {"<ctrl>+<alt>+q": lambda: self.msg_q.put(("toggle", 0, None))}
            )
            self.hotkeys.daemon = True
            self.hotkeys.start()
        except Exception:
            self.hotkeys = None
            log("注册全局热键失败\n" + traceback.format_exc())

        self._update_win_rect()
        self.root.after(40, self._pump)
        self.root.after(300, self._autoload_doc)

    # ---------------------------------------------------------------- UI
    def _initial_geometry(self) -> str:
        pos = self.cfg.get("window_pos")
        if isinstance(pos, list) and len(pos) == 2:
            return f"430x500+{int(pos[0])}+{int(pos[1])}"
        sw = self.root.winfo_screenwidth()
        return f"430x500+{max(0, sw - 470)}+140"

    def _hwnd(self, win=None) -> int:
        try:
            return int(ctypes.windll.user32.GetAncestor(int((win or self.root).winfo_id()), 2))
        except Exception:
            return 0

    def _fit_on_screen(self) -> None:
        """按窗口的真实像素尺寸把窗口收进屏幕内，避免高 DPI 下溢出。"""
        try:
            self.root.update_idletasks()
            hwnd = self._hwnd()
            rect = wintypes.RECT()
            if not hwnd or not ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return
            sw = ctypes.windll.user32.GetSystemMetrics(0)
            sh = ctypes.windll.user32.GetSystemMetrics(1)
            w, h = rect.right - rect.left, rect.bottom - rect.top
            x = max(20, min(rect.left, sw - w - 20))
            y = max(20, min(rect.top, sh - h - 50))
            if (x, y) != (rect.left, rect.top):
                self.root.geometry(f"+{x}+{y}")
        except Exception:
            log("窗口位置校正失败\n" + traceback.format_exc())

    def _fit_window_on_screen(self, win) -> None:
        """把子窗口按真实像素收进屏幕，避免高 DPI 或居于屏幕右侧时溢出。"""
        try:
            win.update_idletasks()
            hwnd = int(ctypes.windll.user32.GetAncestor(int(win.winfo_id()), 2))
            rect = wintypes.RECT()
            if not hwnd or not ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return
            sw = ctypes.windll.user32.GetSystemMetrics(0)
            sh = ctypes.windll.user32.GetSystemMetrics(1)
            w, h = rect.right - rect.left, rect.bottom - rect.top
            x = max(10, min(rect.left, sw - w - 10))
            y = max(10, min(rect.top, sh - h - 40))
            if (x, y) != (rect.left, rect.top):
                ctypes.windll.user32.SetWindowPos(hwnd, 0, x, y, 0, 0, 0x0001 | 0x0004)
        except Exception:
            log("窗口位置校正失败\n" + traceback.format_exc())

    def _titlebar(self, win=None) -> None:
        try:
            hwnd = self._hwnd(win)
            if not hwnd:
                return
            value = ctypes.c_int(0)  # 0 = 浅色标题栏
            for attr in (20, 19):
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)
                ) == 0:
                    break
        except Exception:
            pass

    @staticmethod
    def _gradient(cv, w: int, h: int, c0: str, c1: str) -> None:
        """在画布上从左到右画一条横向渐变。"""
        a = tuple(int(c0[i : i + 2], 16) for i in (1, 3, 5))
        b = tuple(int(c1[i : i + 2], 16) for i in (1, 3, 5))
        for x in range(w):
            t = x / max(1, w - 1)
            color = "#%02x%02x%02x" % tuple(
                int(a[i] + (b[i] - a[i]) * t) for i in range(3)
            )
            cv.create_line(x, 0, x, h, fill=color)

    def _mkbtn(self, parent, text, cmd, primary=False, font_size=9):
        """主按钮：蓝底白字；次级按钮：白底深蓝字 + 淡蓝边。"""
        if primary:
            bg, abg, edge, fg = ACCENT, ACCENT_HOVER, ACCENT, "#ffffff"
        else:
            bg, abg, edge, fg = CARD, CARD_HOVER, BORDER, TEXT_SOFT
        return tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg,
            fg=fg,
            disabledforeground=SUB,
            activebackground=abg,
            activeforeground=fg,
            bd=0,
            relief="flat",
            font=(FONT, font_size, "bold" if primary else "normal"),
            padx=12 if primary else 10,
            pady=4,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=edge,
            highlightcolor=SKY,
        )

    def _section_label(self, parent, text: str) -> None:
        """带竖条的小节标题。"""
        bar = tk.Frame(parent, bg=SKY, width=3, height=12)
        bar.pack(side="left", padx=(2, 7), pady=1)
        tk.Label(parent, text=text, fg=SUB, bg=BG, font=(FONT, 9)).pack(side="left")

    def _card(self, parent, bg, pad=4, radius=10, **pack_kw):
        """圆角卡片：在画布上画圆角白底，里面放一个 Frame，文字框放在 Frame 里。

        返回 (内部 Frame, 设置边框颜色的函数)。bg 需与父容器底色一致，画布本身才看不出边界。
        """
        cv = tk.Canvas(parent, bg=bg, bd=0, highlightthickness=0)
        cv.pack(**pack_kw)
        inner = tk.Frame(cv, bg=CARD)
        win = cv.create_window(0, 0, window=inner, anchor="nw")
        st = {"outline": BORDER, "req": None}

        def draw(_event=None):
            w, h = cv.winfo_width(), cv.winfo_height()
            if w < 4 or h < 4:
                return
            r = max(1, min(radius, w // 2, h // 2))
            x2, y2 = w - 1, h - 1
            cv.delete("card")
            o = st["outline"]
            boxes = ((0, 0, 2 * r, 2 * r), (x2 - 2 * r, 0, x2, 2 * r),
                     (x2 - 2 * r, y2 - 2 * r, x2, y2), (0, y2 - 2 * r, 2 * r, y2))
            starts = (90, 0, 270, 180)
            for box, start in zip(boxes, starts):
                cv.create_arc(*box, start=start, extent=90, style="pieslice",
                              fill=CARD, outline="", tags="card")
                cv.create_arc(*box, start=start, extent=90, style="arc",
                              outline=o, tags="card")
            cv.create_rectangle(r, 0, x2 - r, y2, fill=CARD, outline="", tags="card")
            cv.create_rectangle(0, r, x2, y2 - r, fill=CARD, outline="", tags="card")
            cv.create_line(r, 0, x2 - r, 0, fill=o, tags="card")
            cv.create_line(r, y2, x2 - r, y2, fill=o, tags="card")
            cv.create_line(0, r, 0, y2 - r, fill=o, tags="card")
            cv.create_line(x2, r, x2, y2 - r, fill=o, tags="card")
            cv.coords(win, pad, pad)
            cv.itemconfigure(win, width=w - 2 * pad, height=h - 2 * pad)

        def fit(_event=None):
            # 画布不会自动跟随内部控件，这里按内部需求尺寸同步画布的请求尺寸
            size = (inner.winfo_reqwidth() + 2 * pad, inner.winfo_reqheight() + 2 * pad)
            if size != st["req"]:
                st["req"] = size
                cv.configure(width=size[0], height=size[1])

        cv.bind("<Configure>", draw)
        inner.bind("<Configure>", fit)
        fit()
        return inner, lambda color: (st.__setitem__("outline", color), draw())

    def _build_ui(self) -> None:
        # 顶栏：整条是浅天蓝渐变，左侧带角色形象和标题，右侧是操作按钮
        header = tk.Frame(self.root, bg=HEADER_R, height=self.head_h)
        header.pack(fill="x")
        header.pack_propagate(False)

        brand = tk.Canvas(
            header, width=self.brand_w, height=self.head_h,
            bg=HEADER_L, bd=0, highlightthickness=0,
        )
        brand.pack(side="left")
        self._gradient(brand, self.brand_w, self.head_h, HEADER_L, HEADER_R)
        tx = int(10 * self.ui)
        if self.mascot is not None:
            brand.create_image(
                tx + self.mascot.width() // 2, self.head_h // 2,
                image=self.mascot, anchor="center",
            )
            tx += self.mascot.width() + int(8 * self.ui)
        brand.create_text(
            tx, self.head_h // 2, text="划词问答", fill=FG,
            font=(FONT, 11, "bold"), anchor="w",
        )

        acts = tk.Frame(header, bg=HEADER_R)
        acts.pack(side="right", fill="x", expand=True)
        pady = max(8, (self.head_h - 32) // 2)
        self.btn_set = self._mkbtn(acts, "设置", self.open_settings)
        self.btn_set.pack(side="right", padx=(6, 12), pady=pady)

        self.btn_listen = self._mkbtn(
            acts, "开始检测", self.toggle_listen, primary=True
        )
        self.btn_listen.pack(side="right", padx=2, pady=pady)

        self.var_top = tk.BooleanVar(value=bool(self.cfg.get("topmost", True)))
        chk = tk.Checkbutton(
            acts,
            text="置顶",
            variable=self.var_top,
            command=self._on_topmost,
            bg=HEADER_R,
            fg=TEXT_SOFT,
            selectcolor="white",
            activebackground=HEADER_R,
            activeforeground=FG,
            font=(FONT, 9),
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        chk.pack(side="right", padx=(0, 8), pady=pady)
        tk.Frame(self.root, bg=HEADER_LINE, height=1).pack(fill="x")

        # 文档栏：和正文同底色，不再单独占一条色带
        docrow = tk.Frame(self.root, bg=BG)
        docrow.pack(fill="x", padx=14, pady=(8, 0))
        self.btn_doc_clear = self._mkbtn(docrow, "移除", self.remove_doc, font_size=8)
        self.btn_doc = self._mkbtn(docrow, "加载文档", self.choose_doc, font_size=8)
        self.btn_doc.pack(side="right")
        self.lbl_doc = tk.Label(
            docrow,
            text="未加载文档 · 回答基于通用知识",
            fg=SUB,
            bg=BG,
            font=(FONT, 9),
            anchor="w",
        )
        self.lbl_doc.pack(side="left", fill="x", expand=True, padx=(2, 8))

        body = tk.Frame(self.root, bg=BG)
        self.body = body
        body.pack(fill="both", expand=True, padx=12, pady=(4, 0))

        # ---- 提问内容：白色圆角卡片，聚焦时边框变亮蓝
        ask_head = tk.Frame(body, bg=BG)
        ask_head.pack(fill="x")
        self._section_label(ask_head, "提问内容")
        tk.Label(
            ask_head, text="划词或输入 · Enter 发送", fg=SUB, bg=BG, font=(FONT, 8)
        ).pack(side="right", padx=4)
        sel_inner, sel_outline = self._card(body, BG, fill="x", pady=(5, 12))
        self.txt_sel = tk.Text(
            sel_inner,
            height=4,
            width=1,
            bg=CARD,
            fg=TEXT_SOFT,
            insertbackground=FG,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=(FONT, 9),
            wrap="word",
            padx=8,
            pady=6,
            spacing1=1,
            spacing3=1,
            selectbackground=SEL,
            selectforeground=FG,
            inactiveselectbackground=SEL,
        )
        self.txt_sel.pack(fill="x", padx=6, pady=3)
        self.txt_sel.bind("<Return>", self._ask_from_input)
        self.txt_sel.bind("<FocusIn>", lambda e: sel_outline(ACCENT))
        self.txt_sel.bind("<FocusOut>", lambda e: sel_outline(BORDER))

        # ---- AI 解答：白色圆角卡片，滚动条收在卡片内
        top_row = tk.Frame(body, bg=BG)
        top_row.pack(fill="x")
        self._section_label(top_row, "AI 解答")
        btn_hist = self._mkbtn(top_row, "历史", self.open_history, font_size=8)
        btn_hist.pack(side="right", padx=(0, 0))
        self.lbl_meta = tk.Label(
            top_row, text="", fg=SUB, bg=BG, font=(FONT, 8)
        )
        self.lbl_meta.pack(side="right", padx=(4, 8))
        ans_inner, _ = self._card(body, BG, fill="both", expand=True, pady=(5, 0))
        self.txt_ans = tk.Text(
            ans_inner,
            height=10,
            width=1,
            bg=CARD,
            fg=FG,
            insertbackground=FG,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=(FONT, 10),
            wrap="word",
            padx=10,
            pady=8,
            spacing1=2,
            spacing3=3,
            selectbackground=SEL,
            selectforeground=FG,
            inactiveselectbackground=SEL,
        )
        sb = ttk.Scrollbar(
            ans_inner, orient="vertical", command=self.txt_ans.yview, style="Ocean.Vertical.TScrollbar"
        )
        self.txt_ans.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(2, 5), pady=6)
        self.txt_ans.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=4)
        self._setup_tags()

        footer = tk.Frame(self.root, bg=BG)
        self.footer = footer
        # 先于正文参与布局并贴底；窗口高度不足时由 _on_root_configure 优先折叠隐藏
        footer.pack(side="bottom", fill="x", padx=14, pady=(6, 10), before=body)

        btn_clear = self._mkbtn(footer, "清空", self.clear_answer, font_size=8)
        btn_clear.pack(side="right")
        btn_copy = self._mkbtn(footer, "复制", self.copy_answer, font_size=8)
        btn_copy.pack(side="right", padx=(0, 6))

        # 状态灯：监听中亮绿色，暂停时为灰蓝
        self.dot = tk.Label(footer, text="●", fg=SUB, bg=BG, font=(FONT, 8))
        self.dot.pack(side="left", padx=(2, 6))
        self.status = tk.Label(
            footer,
            text="已暂停 · 点「开始检测」生效",
            fg=SUB,
            bg=BG,
            font=(FONT, 8),
            anchor="w",
        )
        self.status.pack(side="left", fill="x", expand=True)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Ocean.Vertical.TScrollbar",
            background="#c9e0f7",
            troughcolor=CARD,
            bordercolor=CARD,
            arrowcolor="#6b93c4",
            darkcolor="#c9e0f7",
            lightcolor="#c9e0f7",
            arrowsize=12,
        )
        style.map("Ocean.Vertical.TScrollbar", background=[("active", "#b5d5f3")])
        self.root.bind("<Configure>", self._on_root_configure)

    def _on_root_configure(self, event) -> None:
        """窗口变矮时先折叠底部操作栏，优先保证 AI 解答区可读。"""
        if event.widget is not self.root:
            return
        h = event.height
        try:
            if not self.footer_hidden and h <= FOOTER_HIDE_H:
                self.footer_hidden = True
                self.footer.pack_forget()
            elif self.footer_hidden and h >= FOOTER_SHOW_H:
                self.footer_hidden = False
                self.footer.pack(
                    side="bottom", fill="x", padx=14, pady=(6, 10), before=self.body
                )
        except Exception:
            log("底部操作栏折叠失败\n" + traceback.format_exc())

    def _setup_tags(self) -> None:
        self.txt_ans.tag_configure("h", font=(FONT, 11, "bold"), foreground=GOLD, spacing1=6, spacing3=4)
        self.txt_ans.tag_configure("b", font=(FONT, 10, "bold"), foreground=FG)
        self.txt_ans.tag_configure("code", font=(MONO, 9), background="#eaf3fd", foreground=SKY)
        self.txt_ans.tag_configure("quote", foreground=SUB, lmargin1=10, lmargin2=10)
        self.txt_ans.tag_configure("li", lmargin1=14, lmargin2=24)
        self.txt_ans.configure(state="disabled")
        # 答案区划词可继续追问（提问内容框保持可编辑，不绑定）
        self.txt_ans.bind("<ButtonRelease-1>", self._on_inner_select)

    # ------------------------------------------------------------ 文档上下文
    def choose_doc(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择要作为参考的文档",
            filetypes=[
                ("支持的文档", "*.pdf *.docx *.txt *.md *.markdown"),
                ("PDF", "*.pdf"),
                ("Word", "*.docx"),
                (
                    "文本 / 代码",
                    "*.txt *.md *.markdown *.csv *.json *.py *.js *.ts *.java "
                    "*.c *.cpp *.h *.go *.rs *.sql *.html *.xml *.yaml *.yml",
                ),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.load_document(Path(path))

    def _autoload_doc(self) -> None:
        raw = (self.cfg.get("last_doc_path") or "").strip()
        if not raw:
            return
        p = Path(raw)
        if p.is_file():
            self.load_document(p, quiet=True)
        else:
            self.cfg["last_doc_path"] = ""
            save_config(self.cfg)

    def load_document(self, path: Path, quiet: bool = False) -> None:
        name = path.name
        self.lbl_doc.config(text=f"正在解析 {name[:14]}…", fg=SUB)
        if not quiet:
            self.status.config(text=f"正在解析文档：{name}")
        self.btn_doc.config(state="disabled")

        def work():
            try:
                text = normalize_text(read_document(path))
            except Exception as e:
                self.msg_q.put(("doc_error", 0, f"{name} 解析失败：{e}"))
                return
            if not text.strip():
                self.msg_q.put(
                    ("doc_error", 0, f"{name} 里没有提取到文字（扫描版 PDF 需先 OCR）")
                )
                return
            try:
                chunks = tuple(split_chunks(text))
                grams = tuple(_bigrams(c) for c in chunks)
                mtime = path.stat().st_mtime
            except Exception as e:
                self.msg_q.put(("doc_error", 0, f"{name} 解析失败：{e}"))
                return
            self.msg_q.put(
                (
                    "doc",
                    0,
                    {
                        "path": path,
                        "text": text,
                        "chunks": chunks,
                        "grams": grams,
                        "mtime": mtime,
                        "quiet": quiet,
                    },
                )
            )

        threading.Thread(target=work, daemon=True).start()

    def _on_doc_loaded(self, info: dict) -> None:
        with self.doc_lock:
            self.doc_path = info["path"]
            self.doc_text = info["text"]
            self.doc_chunks = info["chunks"]
            self.doc_grams = info["grams"]
            self.doc_mtime = info["mtime"]
        self.cfg["last_doc_path"] = str(info["path"])
        save_config(self.cfg)
        self.btn_doc.config(state="normal", text="换文档")
        if not self.btn_doc_clear.winfo_manager():
            self.btn_doc_clear.pack(side="right", padx=(0, 6))
        self._update_doc_label()
        if not info.get("quiet"):
            self.status.config(
                text=f"已加载 {info['path'].name[:16]} · 回答优先依据文档"
            )

    def _on_doc_error(self, msg: str) -> None:
        self.btn_doc.config(state="normal")
        if self.doc_path:
            self.btn_doc.config(text="换文档")
            self._update_doc_label()
        else:
            self.btn_doc.config(text="加载文档")
            self.lbl_doc.config(text="文档加载失败", fg=WARN)
        self.status.config(text=msg[:120])
        log(msg)

    def remove_doc(self) -> None:
        with self.doc_lock:
            self.doc_path = None
            self.doc_text = ""
            self.doc_chunks = ()
            self.doc_grams = ()
            self.doc_mtime = 0.0
        self.cfg["last_doc_path"] = ""
        save_config(self.cfg)
        self.btn_doc.config(text="加载文档")
        self.btn_doc_clear.pack_forget()
        self._update_doc_label()
        self.status.config(text="已移除文档，回答将基于通用知识")

    def _update_doc_label(self) -> None:
        with self.doc_lock:
            path = self.doc_path
            total = len(self.doc_text)
            count = len(self.doc_chunks)
        if not path:
            self.lbl_doc.config(text="未加载文档 · 回答基于通用知识", fg=SUB)
            return
        name = path.name
        if len(name) > 16:
            name = name[:15] + "…"
        self.lbl_doc.config(text=f"{name} · {total:,} 字 · {count} 段", fg=TEXT_SOFT)

    def _refresh_doc_if_needed(self) -> None:
        """文档被外部修改保存后，提问时自动重新解析（在后台线程调用）。"""
        with self.doc_lock:
            path = self.doc_path
            old_mtime = self.doc_mtime
        if not path:
            return
        try:
            mtime = path.stat().st_mtime
        except Exception:
            return
        if abs(mtime - old_mtime) < 1e-6:
            return
        try:
            text = normalize_text(read_document(path))
            if not text.strip():
                return
            chunks = tuple(split_chunks(text))
            grams = tuple(_bigrams(c) for c in chunks)
            with self.doc_lock:
                self.doc_text = text
                self.doc_chunks = chunks
                self.doc_grams = grams
                self.doc_mtime = mtime
            log(f"文档已变化，自动重新解析：{path.name}")
        except Exception:
            log("文档自动重新解析失败\n" + traceback.format_exc())

    def _pick_doc_chunks(self, query: str) -> list:
        try:
            budget = int(self.cfg.get("doc_context_chars", DOC_BUDGET))
        except Exception:
            budget = DOC_BUDGET
        if budget <= 0:
            return []
        # 引用片段数随预算放大，避免调大字数后仍只取 4 段
        top_k = max(DOC_TOP_K, budget // DOC_CHUNK + 2)
        with self.doc_lock:
            chunks = self.doc_chunks
            grams = self.doc_grams
        return retrieve_chunks(query, chunks, grams, top_k, budget)

    # ------------------------------------------------------------ 监听状态
    def toggle_listen(self) -> None:
        self.listening = not self.listening
        if self.listening:
            self.dot.config(fg=OK)
            self.btn_listen.config(text="暂停", bg=CARD, activebackground=CARD_HOVER,
                                  fg=TEXT_SOFT, activeforeground=TEXT_SOFT,
                                  highlightbackground=BORDER)
            self.status.config(text="监听中 · Ctrl+Alt+Q 暂停")
            self.root.attributes("-topmost", self.var_top.get())
        else:
            self.dot.config(fg=SUB)
            self.btn_listen.config(text="开始检测", bg=ACCENT, activebackground=ACCENT_HOVER,
                                   fg="#ffffff", activeforeground="#ffffff",
                                   highlightbackground=ACCENT)
            self.status.config(text="已暂停 · 点「开始检测」生效")

    def _on_minimize(self, event=None) -> None:
        """窗口最小化时自动暂停检测，避免在后台无意识地消耗 token。"""
        if event is not None and event.widget is not self.root:
            return
        try:
            if self.root.state() == "iconic" and self.listening:
                self.toggle_listen()
                self.status.config(text="窗口已最小化 · 检测已暂停")
        except Exception:
            log("最小化暂停失败\n" + traceback.format_exc())

    def _on_topmost(self) -> None:
        v = bool(self.var_top.get())
        self.root.attributes("-topmost", v)
        self.cfg["topmost"] = v
        save_config(self.cfg)

    # ------------------------------------------------------------ 选词检测
    def _point_inside(self, x, y) -> bool:
        l, t, r, b = self.win_rect
        if r - l <= 0 or b - t <= 0:
            return False
        return l - 2 <= x <= r + 2 and t - 2 <= y <= b + 2

    def _update_win_rect(self) -> None:
        try:
            hwnd = self._hwnd()
            rect = wintypes.RECT()
            if hwnd and ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                self.win_rect = (rect.left, rect.top, rect.right, rect.bottom)
        except Exception:
            pass
        try:
            self.root.after(400, self._update_win_rect)
        except Exception:
            pass

    def _on_click(self, x, y, button, pressed) -> None:
        try:
            if button != mouse.Button.left:
                return
            if pressed:
                self.press = (x, y, time.time())
                self.press_inside = self._point_inside(x, y)
                return
            if self.press is None:
                return
            px, py, _ = self.press
            self.press = None
            if self.press_inside or not self.listening:
                return
            now = time.time()
            moved = ((x - px) ** 2 + (y - py) ** 2) ** 0.5
            lx, ly, lt = self.last_click
            is_double = (now - lt) < 0.45 and ((x - lx) ** 2 + (y - ly) ** 2) ** 0.5 < 8
            self.last_click = (x, y, now)
            if ctrl_or_alt_down():
                return
            if not (moved >= 5 or is_double):
                return
            threading.Thread(target=self._capture, daemon=True).start()
        except Exception:
            log("鼠标回调异常\n" + traceback.format_exc())

    def _foreground_is_self(self) -> bool:
        try:
            fg = ctypes.windll.user32.GetForegroundWindow()
            if not fg:
                return False
            own = self._hwnd()
            return bool(own) and ctypes.windll.user32.GetAncestor(fg, 2) == own
        except Exception:
            return False

    def _ask_from_input(self, event=None):
        """在提问内容框里输入问题后按 Enter 直接发送。"""
        text = self.txt_sel.get("1.0", "end-1c").strip()
        if not text:
            self.status.config(text="请先划词，或在「提问内容」里输入问题")
            return "break"
        if text == self.last_text and time.time() - self.last_text_at < 3:
            return "break"
        self.last_text = text
        self.last_text_at = time.time()
        self.gen += 1
        gen = self.gen
        self.msg_q.put(("selection", gen, text))
        threading.Thread(target=self._ask, args=(text, gen), daemon=True).start()
        return "break"

    def _on_inner_select(self, event) -> None:
        """在答案区划词时，直接对选中的文字继续提问。"""
        if not self.listening:
            return
        try:
            text = event.widget.get("sel.first", "sel.last")
        except tk.TclError:
            return  # 只是单击或拖动滚动条，没有文字选区
        text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text).strip()
        if not text:
            return
        if text == self.last_text and time.time() - self.last_text_at < 3:
            return
        self.last_text = text
        self.last_text_at = time.time()
        self.gen += 1
        gen = self.gen
        self.msg_q.put(("selection", gen, text))
        threading.Thread(target=self._ask, args=(text, gen), daemon=True).start()

    def _capture(self) -> None:
        try:
            time.sleep(0.1)  # 等选区稳定
            if self._foreground_is_self():
                return
            if ctrl_or_alt_down():
                return  # 用户正按着 Ctrl/Alt（多半要复制），不抢占剪贴板
            old = clip_get()
            seq0 = clipboard_seq()
            kc = keyboard.Controller()
            kc.press(keyboard.Key.ctrl)
            try:
                kc.press("c")
                kc.release("c")
            finally:
                kc.release(keyboard.Key.ctrl)
            time.sleep(0.16)
            text = clip_get()
            seq1 = clipboard_seq()
            if text == old or (seq0 and seq1 == seq0):
                return  # Ctrl+C 没生效：不是文字选区（比如拖动了滚动条）
            # 抓取期间若用户又复制了新内容（序号变化）或正按着 Ctrl，就保留用户剪贴板，
            # 避免把用户刚复制的文字覆盖回旧内容。
            if not (seq1 and clipboard_seq() != seq1) and not ctrl_or_alt_down():
                clip_set(old)  # 还原用户原剪贴板
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            text = re.sub(r"[ \t]+", " ", text)
            text = re.sub(r"\n{3,}", "\n\n", text).strip()
            if not text:
                return
            if text == self.last_text and time.time() - self.last_text_at < 3:
                return
            self.last_text = text
            self.last_text_at = time.time()
            self.gen += 1
            gen = self.gen
            self.msg_q.put(("selection", gen, text))
            self._ask(text, gen)
        except Exception:
            log("抓取选区异常\n" + traceback.format_exc())

    # ------------------------------------------------------------ AI 请求
    def _ask(self, text: str, gen: int) -> None:
        cfg = self.cfg
        key = (cfg.get("api_key") or "").strip()
        if not key:
            self.msg_q.put(
                (
                    "error",
                    gen,
                    "还没有配置 API Key。\n点右上角「设置」，填入 DeepSeek / 通义 / 智谱等任意 OpenAI 兼容接口的 Key，保存后重新选词即可。",
                )
            )
            return

        base = (cfg.get("base_url") or "").strip().rstrip("/")
        url = base + "/chat/completions"
        limit = int(cfg.get("max_chars") or 2000)
        cut = text if len(text) <= limit else text[:limit] + "…"

        # 检索文档中最相关的片段，让回答优先依据文档内容
        self._refresh_doc_if_needed()
        doc_chunks = self._pick_doc_chunks(text)
        user_content = "选中文本：\n" + cut + "\n\n"
        if doc_chunks:
            doc_name = self.doc_path.name if self.doc_path else "参考文档"
            user_content += (
                f"参考文档《{doc_name}》中的相关片段（按原文顺序）：\n<<<\n"
                + "\n\n……\n\n".join(doc_chunks)
                + "\n>>>\n\n"
                "请优先依据上面的文档片段解释选中文本；如果文档片段没有覆盖相关内容，"
                "可以用通用知识补充，但要说明“文档中未涉及”。\n\n"
            )
        user_content += "请解释这段内容。"
        payload = {
            "model": cfg.get("model") or "deepseek-flash",
            "messages": [
                {"role": "system", "content": cfg.get("system_prompt") or DEFAULT_CONFIG["system_prompt"]},
                {"role": "user", "content": user_content},
            ],
            "stream": True,
            "temperature": 0.3,
        }
        # DeepSeek 新模型默认开启思考模式；划词解释求快，这里显式关闭
        if "api.deepseek.com" in base:
            payload["thinking"] = {"type": "disabled"}
        if doc_chunks:
            self.msg_q.put(
                ("status", gen, f"识别 {len(text)} 字 · 引用文档 {len(doc_chunks)} 段")
            )
        headers = {
            "Authorization": "Bearer " + key,
            "Content-Type": "application/json",
            "Accept": "text/event-stream",
        }
        started = time.time()
        buf = []
        try:
            with requests.post(
                url, headers=headers, json=payload, stream=True, timeout=(10, 120)
            ) as r:
                if r.status_code != 200:
                    detail = r.text[:300].replace("\n", " ")
                    self.msg_q.put(
                        ("error", gen, f"请求失败 HTTP {r.status_code}\n{detail}")
                    )
                    return
                r.encoding = "utf-8"
                for raw in r.iter_lines(decode_unicode=True):
                    if self.gen != gen:
                        return
                    if not raw or not raw.startswith("data:"):
                        continue
                    data = raw[5:].strip()
                    if data == "[DONE]":
                        break
                    try:
                        obj = json.loads(data)
                        delta = obj["choices"][0].get("delta") or {}
                        piece = delta.get("content") or ""
                    except Exception:
                        continue
                    if piece:
                        buf.append(piece)
                        self.msg_q.put(("delta", gen, piece))
            if self.gen != gen:
                return
            cost = time.time() - started
            self.msg_q.put(("done", gen, "".join(buf)))
            self.msg_q.put(("status", gen, f"完成 · {cost:.1f}s"))
        except requests.exceptions.Timeout:
            self.msg_q.put(("error", gen, "请求超时，请检查网络或接口地址。"))
        except Exception as e:
            self.msg_q.put(("error", gen, f"网络错误：{e}"))

    # ------------------------------------------------------------ UI 更新
    def _pump(self) -> None:
        try:
            while True:
                kind, gen, data = self.msg_q.get_nowait()
                if kind == "doc":
                    self._on_doc_loaded(data)
                elif kind == "doc_error":
                    self._on_doc_error(data)
                elif kind == "toggle":
                    self.toggle_listen()
                elif kind == "selection" and gen == self.gen:
                    self._show_selection(data)
                elif kind == "delta" and gen == self.gen:
                    self._append_answer(data)
                elif kind == "done" and gen == self.gen:
                    self._render_answer(data)
                elif kind == "error" and gen == self.gen:
                    self._show_error(data)
                elif kind == "status" and gen == self.gen:
                    self.lbl_meta.config(text=data)
        except queue.Empty:
            pass
        except Exception:
            log("界面刷新异常\n" + traceback.format_exc())
        try:
            self.root.after(40, self._pump)
        except Exception:
            pass

    def _show_selection(self, text: str) -> None:
        self.current_question = text
        self.txt_sel.delete("1.0", "end")
        self.txt_sel.insert("1.0", text[:800] + ("…" if len(text) > 800 else ""))
        self.txt_sel.see("1.0")
        self.answer_text = ""
        self.txt_ans.configure(state="normal")
        self.txt_ans.delete("1.0", "end")
        self.txt_ans.configure(state="disabled")
        self.lbl_meta.config(text=f"识别 {len(text)} 字 · 生成中…")
        self.status.config(text="已识别选中文字，正在请求解答…")

    def _append_answer(self, piece: str) -> None:
        self.answer_text += piece
        self.txt_ans.configure(state="normal")
        self.txt_ans.insert("end", piece)
        self.txt_ans.configure(state="disabled")
        self.txt_ans.see("end")

    def _render_answer(self, full: str) -> None:
        if not full.strip():
            return
        self.answer_text = full
        self.txt_ans.configure(state="normal")
        self.txt_ans.delete("1.0", "end")
        for line in full.split("\n"):
            self._insert_md_line(line)
        self.txt_ans.configure(state="disabled")
        self.txt_ans.see("end")
        self._add_history(self.current_question, full)
        self.status.config(text="解答完成 · 可复制或继续选词")

    # ------------------------------------------------------------ 历史问答
    def _load_history(self) -> list:
        try:
            if HISTORY_PATH.exists():
                data = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    items = [x for x in data if isinstance(x, dict) and x.get("q")]
                    return items[-HISTORY_MAX:]
        except Exception:
            log("读取历史记录失败\n" + traceback.format_exc())
        return []

    def _save_history(self) -> None:
        try:
            HISTORY_PATH.write_text(
                json.dumps(self.history[-HISTORY_MAX:], ensure_ascii=False, indent=1),
                encoding="utf-8",
            )
        except Exception:
            log("保存历史记录失败\n" + traceback.format_exc())

    def _add_history(self, question: str, answer: str) -> None:
        q = (question or "").strip()
        a = (answer or "").strip()
        if not q or not a:
            return
        self.history.append(
            {
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "q": q,
                "a": a,
                "doc": self.doc_path.name if self.doc_path else "",
            }
        )
        self.history = self.history[-HISTORY_MAX:]
        self._save_history()

    def open_history(self) -> None:
        """历史问答窗口：查看历史 Q&A，并可导出为 Markdown。"""
        if self.history_win is not None and self.history_win.winfo_exists():
            self.history_win.lift()
            self.history_win.focus_force()
            return

        win = tk.Toplevel(self.root)
        win.withdraw()  # 先隐藏，等内容与尺寸就绪后再显示，避免开场闪黑框
        self.history_win = win
        win.title("历史问答")
        win.configure(bg=BG)
        win.transient(self.root)
        try:
            win.iconbitmap(str(resource_path("app.ico")))
        except Exception:
            pass
        win.minsize(480, 380)

        def on_close() -> None:
            self.history_win = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)

        # 顶栏：与主窗口、设置窗口同款浅天蓝 + 金色竖条
        head = tk.Frame(win, bg=HEADER_L, height=44)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Frame(head, bg=GOLD, width=3, height=14).pack(side="left", padx=(16, 8))
        tk.Label(
            head, text="历史问答", fg=FG, bg=HEADER_L, font=(FONT, 11, "bold")
        ).pack(side="left")
        self._mkbtn(head, "关闭", on_close, font_size=9).pack(
            side="right", padx=(6, 14), pady=8
        )
        btn_export = self._mkbtn(head, "导出", None, primary=True, font_size=9)
        btn_export.pack(side="right", pady=8)
        tk.Frame(win, bg=HEADER_LINE, height=1).pack(fill="x")
        win.after(150, lambda: self._titlebar(win))

        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=12, pady=(10, 0))
        inner, _ = self._card(body, BG, fill="both", expand=True)
        txt = tk.Text(
            inner,
            width=1,
            height=13,
            bg=CARD,
            fg=FG,
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=(FONT, 9),
            wrap="word",
            padx=10,
            pady=8,
            spacing1=2,
            spacing3=3,
            selectbackground=SEL,
            selectforeground=FG,
            inactiveselectbackground=SEL,
        )
        sb = ttk.Scrollbar(
            inner,
            orient="vertical",
            command=txt.yview,
            style="Ocean.Vertical.TScrollbar",
        )
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(2, 5), pady=6)
        txt.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=4)
        txt.tag_configure("t", foreground=SUB, font=(FONT, 8))
        txt.tag_configure("q", foreground=SKY, font=(FONT, 10, "bold"), spacing1=8)
        txt.tag_configure("a", foreground=FG, font=(FONT, 9))
        txt.tag_configure("sep", foreground=BORDER)

        if not self.history:
            txt.insert(
                "end",
                "还没有历史记录。\n\n划词或输入问题并得到解答后，记录会自动保存在这里。",
                "t",
            )
        else:
            for item in reversed(self.history):
                t = str(item.get("time", ""))
                q = str(item.get("q", "")).strip()
                a = str(item.get("a", "")).strip()
                doc = str(item.get("doc", "") or "")
                txt.insert("end", t + (f" · 参考：{doc}" if doc else "") + "\n", "t")
                if len(q) > 240:
                    q = q[:240] + "…"
                if len(a) > 1200:
                    a = a[:1200] + "…"
                txt.insert("end", "问题：" + q + "\n", "q")
                txt.insert("end", "回答：\n" + a + "\n", "a")
                txt.insert("end", "—" * 40 + "\n\n", "sep")
        txt.configure(state="disabled")
        txt.see("1.0")

        foot = tk.Frame(win, bg=BG)
        foot.pack(fill="x", padx=16, pady=(6, 12))
        tip = tk.Label(
            foot,
            text=f"共 {len(self.history)} 条记录 · 导出为 Markdown",
            fg=SUB,
            bg=BG,
            font=(FONT, 8),
            anchor="w",
        )
        tip.pack(side="left", fill="x", expand=True)
        btn_export.configure(command=lambda: self.export_history(win, tip))

        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        w = min(620, sw - 80)
        h = min(max(win.winfo_reqheight(), 500), sh - 100)
        x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - h) // 2
        x = max(20, min(x, sw - w - 20))
        y = max(20, min(y, sh - h - 60))
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.update()  # 先把内容绘制完成再显示，避免首帧深色底一闪
        win.deiconify()
        win.lift()

    def export_history(self, parent, tip_label) -> None:
        if not self.history:
            tip_label.config(text="还没有历史记录，无法导出")
            return
        default = f"划词问答历史_{datetime.now():%Y%m%d_%H%M}.md"
        path = filedialog.asksaveasfilename(
            parent=parent,
            title="导出历史记录",
            defaultextension=".md",
            initialfile=default,
            filetypes=[
                ("Markdown 文档", "*.md"),
                ("Word 文档", "*.docx"),
                ("文本文件", "*.txt"),
            ],
        )
        if not path:
            return
        try:
            out = Path(path)
            if out.suffix.lower() == ".docx":
                self._export_docx(out)
            else:
                out.write_text(self._history_markdown(), encoding="utf-8")
            tip_label.config(text=f"已导出：{out.name}")
            self.status.config(text="历史记录已导出")
        except Exception as exc:
            tip_label.config(text=f"导出失败：{exc}")
            log("导出历史记录失败\n" + traceback.format_exc())

    def _export_docx(self, path: Path) -> None:
        """把历史问答导出为 Word 文档。"""
        from docx import Document
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor

        doc = Document()
        normal = doc.styles["Normal"]
        normal.font.name = "Microsoft YaHei"
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        normal.font.size = Pt(10.5)

        def cn(run) -> None:
            run.font.name = "Microsoft YaHei"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

        def gray(p, size=9):
            for run in p.runs:
                cn(run)
                run.font.size = Pt(size)
                run.font.color.rgb = RGBColor(0x5B, 0x7B, 0xA6)

        p = doc.add_paragraph()
        r = p.add_run("划词问答历史记录")
        cn(r)
        r.bold = True
        r.font.size = Pt(18)
        r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x63)

        p = doc.add_paragraph()
        p.add_run(f"导出时间：{datetime.now():%Y-%m-%d %H:%M:%S}")
        gray(p)

        for i, item in enumerate(reversed(self.history), 1):
            p = doc.add_paragraph()
            r = p.add_run(f"{i}. {item.get('time', '')}")
            cn(r)
            r.bold = True
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x63)

            doc_name = str(item.get("doc", "") or "")
            if doc_name:
                p = doc.add_paragraph()
                p.add_run(f"参考文档：{doc_name}")
                gray(p)

            p = doc.add_paragraph()
            r = p.add_run("问题：")
            cn(r)
            r.bold = True
            r2 = p.add_run(str(item.get("q", "")).strip())
            cn(r2)

            p = doc.add_paragraph()
            r = p.add_run("回答：")
            cn(r)
            r.bold = True
            for line in str(item.get("a", "")).split("\n"):
                if line.strip().startswith("```"):
                    continue
                para = doc.add_paragraph(line.rstrip())
                for run in para.runs:
                    cn(run)

        doc.save(str(path))

    def _history_markdown(self) -> str:
        lines = [
            "# 划词问答历史记录",
            "",
            f"导出时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
            "",
        ]
        for i, item in enumerate(reversed(self.history), 1):
            lines.append(f"## {i}. {item.get('time', '')}")
            lines.append("")
            doc = str(item.get("doc", "") or "")
            if doc:
                lines.append(f"参考文档：{doc}")
                lines.append("")
            lines.append("**问题：**")
            lines.append("")
            lines.append(str(item.get("q", "")).strip())
            lines.append("")
            lines.append("**回答：**")
            lines.append("")
            lines.append(str(item.get("a", "")).strip())
            lines.append("")
            lines.append("---")
            lines.append("")
        return "\n".join(lines)

    def _insert_md_line(self, line: str) -> None:
        s = line.rstrip()
        if s.strip().startswith("```"):
            return
        m = re.match(r"^#{1,4}\s+(.*)$", s)
        if m:
            self._insert_inline(m.group(1), ("h",))
            self.txt_ans.insert("end", "\n")
            return
        if s.startswith(("- ", "* ")):
            self._insert_inline("• " + s[2:], ("li",))
            self.txt_ans.insert("end", "\n")
            return
        if s.startswith(">"):
            self._insert_inline(s.lstrip("> ").strip(), ("quote",))
            self.txt_ans.insert("end", "\n")
            return
        self._insert_inline(s, ())
        self.txt_ans.insert("end", "\n")

    def _insert_inline(self, s: str, base: tuple) -> None:
        pattern = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")
        pos = 0
        for m in pattern.finditer(s):
            if m.start() > pos:
                self.txt_ans.insert("end", s[pos : m.start()], base)
            seg = m.group(0)
            if seg.startswith("**"):
                self.txt_ans.insert("end", seg[2:-2], base + ("b",))
            else:
                self.txt_ans.insert("end", seg[1:-1], base + ("code",))
            pos = m.end()
        if pos < len(s):
            self.txt_ans.insert("end", s[pos:], base)

    def _show_error(self, msg: str) -> None:
        self.answer_text = msg
        self.txt_ans.configure(state="normal")
        self.txt_ans.delete("1.0", "end")
        self.txt_ans.insert("1.0", msg)
        self.txt_ans.configure(state="disabled")
        self.lbl_meta.config(text="出错")
        self.status.config(text="请求失败，详见窗口内提示")

    def copy_answer(self) -> None:
        if self.answer_text.strip():
            clip_set(self.answer_text)
            self.status.config(text="已复制回答到剪贴板")

    def clear_answer(self) -> None:
        self.answer_text = ""
        self.txt_ans.configure(state="normal")
        self.txt_ans.delete("1.0", "end")
        self.txt_ans.configure(state="disabled")
        self.lbl_meta.config(text="")
        self.status.config(text="已清空")

    # ------------------------------------------------------------ 设置
    def open_settings(self) -> None:
        win = tk.Toplevel(self.root)
        win.withdraw()  # 先隐藏，等内容与尺寸就绪后再显示，避免开场闪黑框
        win.title("设置")
        win.configure(bg=BG)
        win.transient(self.root)
        win.resizable(False, False)
        try:
            win.iconbitmap(str(resource_path("app.ico")))
        except Exception:
            pass

        # 设置窗口顶栏：与主窗口同款浅天蓝 + 金色竖条
        shead = tk.Frame(win, bg=HEADER_L, height=44)
        shead.pack(fill="x")
        shead.pack_propagate(False)
        sbar = tk.Frame(shead, bg=GOLD, width=3, height=14)
        sbar.pack(side="left", padx=(16, 8))
        tk.Label(shead, text="设置", fg=FG, bg=HEADER_L, font=(FONT, 11, "bold")).pack(side="left")
        tk.Frame(win, bg=HEADER_LINE, height=1).pack(fill="x")
        win.after(150, lambda: self._titlebar(win))

        def field(label, value, show=None):
            tk.Label(win, text=label, fg=SUB, bg=BG, font=(FONT, 9)).pack(
                anchor="w", padx=16, pady=(10, 2)
            )
            e = tk.Entry(
                win,
                bg=CARD,
                fg=FG,
                insertbackground=FG,
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground=BORDER,
                highlightcolor=ACCENT,
                font=(FONT, 9),
                show=show,
            )
            e.pack(fill="x", padx=16, ipady=5)
            e.insert(0, value or "")
            return e

        e_url = field("接口地址（OpenAI 兼容，通常以 /v1 结尾）", self.cfg.get("base_url", ""))
        e_model = field("模型名", self.cfg.get("model", ""))
        e_key = field("API Key", self.cfg.get("api_key", ""), show="•")

        tk.Label(win, text="提示词（决定回答风格）", fg=SUB, bg=BG, font=(FONT, 9)).pack(
            anchor="w", padx=16, pady=(10, 2)
        )
        t_prompt = tk.Text(
            win,
            height=5,
            bg=CARD,
            fg=FG,
            insertbackground=FG,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            font=(FONT, 9),
            wrap="word",
            padx=8,
            pady=6,
        )
        t_prompt.pack(fill="x", padx=16)
        t_prompt.insert("1.0", self.cfg.get("system_prompt", ""))

        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", padx=16, pady=(10, 0))
        tk.Label(row, text="单次最多发送字数", fg=SUB, bg=BG, font=(FONT, 9)).pack(side="left")
        e_max = tk.Entry(
            row,
            width=8,
            bg=CARD,
            fg=FG,
            insertbackground=FG,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            font=(FONT, 9),
        )
        e_max.pack(side="right", ipady=4)
        e_max.insert(0, str(self.cfg.get("max_chars", 2000)))

        row2 = tk.Frame(win, bg=BG)
        row2.pack(fill="x", padx=16, pady=(8, 0))
        tk.Label(
            row2,
            text="每次附带的文档参考字数（0 = 不引用文档）",
            fg=SUB,
            bg=BG,
            font=(FONT, 9),
        ).pack(side="left")
        e_doc = tk.Entry(
            row2,
            width=8,
            bg=CARD,
            fg=FG,
            insertbackground=FG,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            font=(FONT, 9),
        )
        e_doc.pack(side="right", ipady=4)
        e_doc.insert(0, str(self.cfg.get("doc_context_chars", DOC_BUDGET)))

        tip = (
            "常用地址：\n"
            "DeepSeek  https://api.deepseek.com/v1  模型 deepseek-flash（V4.1）或 deepseek-v4-pro\n"
            "通义千问  https://dashscope.aliyuncs.com/compatible-mode/v1  模型 qwen-plus\n"
            "智谱      https://open.bigmodel.cn/api/paas/v4  模型 glm-4-flash"
        )
        tk.Label(
            win, text=tip, fg=SUB, bg=BG, font=(FONT, 8), justify="left", wraplength=440
        ).pack(anchor="w", padx=16, pady=(10, 0))

        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=16, pady=14)
        self._mkbtn(btns, "保存", lambda: on_save(), primary=True).pack(side="right")
        self._mkbtn(btns, "取消", win.destroy).pack(side="right", padx=8)

        # 按内容计算高度，确保「保存/取消」始终可见（高 DPI 下尤其重要）
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        h = min(win.winfo_reqheight(), sh - 80)
        px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        x = max(20, min(px + (pw - 480) // 2, sw - 480 - 20))
        y = max(20, min(py + (ph - h) // 2, sh - h - 60))
        win.geometry(f"480x{h}+{x}+{y}")
        win.update()  # 先把内容绘制完成再显示，避免首帧深色底一闪
        win.deiconify()
        win.lift()
        win.grab_set()

        def on_save():
            self.cfg["base_url"] = e_url.get().strip()
            self.cfg["model"] = e_model.get().strip()
            self.cfg["api_key"] = e_key.get().strip()
            self.cfg["system_prompt"] = t_prompt.get("1.0", "end").strip()
            try:
                self.cfg["max_chars"] = max(100, int(e_max.get().strip()))
            except Exception:
                self.cfg["max_chars"] = 2000
            try:
                self.cfg["doc_context_chars"] = max(0, int(e_doc.get().strip()))
            except Exception:
                self.cfg["doc_context_chars"] = DOC_BUDGET
            save_config(self.cfg)
            win.destroy()
            self.status.config(text="设置已保存 · 选词试试")

    # ------------------------------------------------------------ 退出
    def on_close(self) -> None:
        try:
            self.cfg["window_pos"] = [self.root.winfo_x(), self.root.winfo_y()]
            save_config(self.cfg)
        except Exception:
            pass
        try:
            self.mouse_listener.stop()
        except Exception:
            pass
        try:
            if self.hotkeys:
                self.hotkeys.stop()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    enable_dpi_awareness()
    # 独立的应用标识：让任务栏 / Alt+Tab 显示 app.ico，而不是 pythonw 自带的图标
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("SelectAI.HuaCiWenDa")
    except Exception:
        pass
    try:
        App().run()
    except Exception:
        log("启动失败\n" + traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
