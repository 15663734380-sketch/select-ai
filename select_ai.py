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
import random
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

try:  # 桌宠动画引擎（只负责画面与动画，不含交互）
    from pet_anim import PetAnimator, layout as pet_layout
except Exception:  # 引擎缺失时桌宠不显示，其余功能不受影响
    PetAnimator = None
    pet_layout = None

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
    "web_search": False,
    "last_doc_path": "",
    "topmost": True,
    "window_pos": None,
    "pet_enabled": False,
    "pet_pos": None,
    "pet_size": 128,
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
        self.cfg["topmost"] = True  # 每次启动都默认置顶
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
        self.pet = None
        self.pet_photo = None
        self.pet_drag = None
        self.pet_size = 128
        self.pet_exited = False
        self.balance_loading = False
        self.balance_bubble = None
        self._bubble_after = None
        self._bubble_size = (0, 0)
        self._fade_start = None
        self._pet_menu_obj = None
        self.pet_anim = None
        self._egg_after = None
        self.ask_grip = None

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
        self._apply_icon(self.root)
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
        self.root.bind("<Map>", self._on_restore)
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
        self.std_h = 0
        self.root.deiconify()
        self.root.lift()
        self.root.update_idletasks()
        self.std_h = self.root.winfo_height()  # 主窗口标准高度（初版布局）

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

    def _apply_icon(self, win) -> None:
        """按 DPI 精确加载图标尺寸，避免任务栏把大图缩放后发虚。"""
        ico = str(resource_path("app.ico"))
        try:
            win.iconbitmap(ico)
        except Exception:
            pass
        try:
            win.update_idletasks()
            hwnd = self._hwnd(win)
            if not hwnd:
                return
            u = ctypes.windll.user32
            u.GetDpiForWindow.argtypes = [ctypes.c_void_p]
            u.GetDpiForWindow.restype = ctypes.c_uint
            dpi = int(u.GetDpiForWindow(ctypes.c_void_p(hwnd))) or 96
            u.LoadImageW.argtypes = [
                ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint,
                ctypes.c_int, ctypes.c_int, ctypes.c_uint,
            ]
            u.LoadImageW.restype = ctypes.c_void_p
            u.SendMessageW.argtypes = [
                ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p
            ]
            if not hasattr(self, "_icon_cache"):
                self._icon_cache = {}
            # 小图标供任务栏/标题栏，大图标供 Alt+Tab；按 96 DPI 基准乘缩放
            for slot, cx in ((0, round(24 * dpi / 96)), (1, round(32 * dpi / 96))):
                key = (slot, cx)
                hicon = self._icon_cache.get(key)
                if not hicon:
                    hicon = u.LoadImageW(None, ico, 1, cx, cx, 0x00000010)
                    if hicon:
                        self._icon_cache[key] = hicon
                if hicon:
                    u.SendMessageW(ctypes.c_void_p(hwnd), 0x80, slot, hicon)
        except Exception:
            log("设置窗口图标失败\n" + traceback.format_exc())

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
        ask_wrap = tk.Frame(body, bg=BG)
        ask_wrap.pack(fill="x", pady=(5, 12))
        self.ask_wrap = ask_wrap
        sel_inner, sel_outline = self._card(ask_wrap, BG, fill="x")
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
        ask_sb = ttk.Scrollbar(
            sel_inner,
            orient="vertical",
            command=self.txt_sel.yview,
            style="Ocean.Vertical.TScrollbar",
        )
        self.txt_sel.configure(yscrollcommand=ask_sb.set)
        ask_sb.pack(side="right", fill="y", padx=(2, 5), pady=3)
        self.txt_sel.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=3)
        self.txt_sel.bind("<MouseWheel>", self._ask_wheel)
        # 提问框下边缘可拖动加高（临时查看更多内容，窗口缩放时恢复标准高度）
        self.ask_canvas = sel_inner.master
        self.ask_inner = sel_inner
        for w in (self.ask_canvas, sel_inner, self.txt_sel):
            w.bind("<ButtonPress-1>", self._ask_grip_press)
            w.bind("<Motion>", self._ask_grip_cursor)
            w.bind("<Leave>", self._ask_grip_leave)
        # 拖动过程中鼠标会移到其它部件上，交由窗口级绑定持续处理
        self.root.bind("<B1-Motion>", self._ask_grip_move, add="+")
        self.root.bind("<ButtonRelease-1>", self._ask_grip_release, add="+")
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

        # 记录提问卡片的默认高度：拖动窗口上边缘时由它吸收高度变化
        self.root.update_idletasks()
        self.ask_h = max(60, ask_wrap.winfo_reqheight())
        self.ask_default_h = self.ask_h  # 初版提问框高度：向上拉伸到此为止
        sel_inner.master.pack_configure(fill="both", expand=True)
        ask_wrap.pack_propagate(False)
        ask_wrap.configure(height=self.ask_h)
        self.ask_min_h = max(56, int(self.txt_sel.winfo_reqheight() * 0.6))
        self._last_h = None
        self.root.bind("<Configure>", self._on_root_configure)

    def _on_root_configure(self, event) -> None:
        """按拖动方向分配高度：下边缘改 AI 解答区，上边缘改提问内容区。"""
        if event.widget is not self.root:
            return
        try:
            if self.root.state() == "iconic":
                return
        except Exception:
            pass
        h = event.height
        lh = self._last_h
        self._last_h = h
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

        # 识别拖的是哪条边：上边缘优先伸缩提问框，下边缘只动解答区。
        try:
            if lh is not None and h != lh and self.ask_h > self.ask_default_h:
                if h < lh:
                    # 缩小窗口时先把临时加高的部分逐步收回来，再按标准逻辑继续缩
                    extra = self.ask_h - self.ask_default_h
                    take = min(extra, lh - h)
                    self._set_ask_h(self.ask_h - take, h, allow_expand=True)
                    if take >= lh - h:
                        return
                    lh -= take  # 剩余压缩量按标准逻辑处理
                else:
                    return  # 放大窗口时保留手动加高的提问框
            if lh is None or h == lh or not self.std_h:
                return
            if self._resize_is_top():
                # 窗口在标准高度以上时提问框保持初版大小；
                # 低于标准高度时，只有超出部分才压缩/恢复提问框。
                change = max(0, self.std_h - lh) - max(0, self.std_h - h)
                self._set_ask_h(self.ask_h + change, h)
            elif h < lh:
                self._clamp_ask_h(h)
        except Exception:
            log("提问区缩放失败\n" + traceback.format_exc())

    def _resize_is_top(self) -> bool:
        """根据鼠标是否停在窗口上边缘判断当前是拖上边还是下边。"""
        try:
            r = wintypes.RECT()
            p = wintypes.POINT()
            hwnd = self._hwnd(self.root)
            if hwnd and ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r)):
                ctypes.windll.user32.GetCursorPos(ctypes.byref(p))
                top_d = abs(p.y - r.top)
                bottom_d = abs(p.y - r.bottom)
                if min(top_d, bottom_d) <= 24:
                    return top_d < bottom_d
        except Exception:
            pass
        return False

    def _set_ask_h(self, value: int, total_h: int = 0, allow_expand: bool = False) -> None:
        """设置提问内容框高度（像素）：默认最大回到初版高度，手动拖动可临时加高。"""
        value = int(value)
        if allow_expand and total_h:
            limit = max(self.ask_default_h, int(total_h) - 340)
        else:
            limit = self.ask_default_h
            if total_h and int(total_h) < self.std_h:
                # 窗口比标准尺寸小时，给解答区留出最低可读空间
                limit = min(limit, max(self.ask_min_h, int(total_h) - 320))
        value = max(self.ask_min_h, min(value, limit))
        if value == self.ask_h:
            return
        self.ask_h = value
        try:
            self.ask_wrap.configure(height=value)
        except Exception:
            pass

    def _clamp_ask_h(self, total_h: int) -> None:
        """下边缘缩小时，提问区若过高则让出空间给解答区。"""
        limit = min(self.ask_default_h, max(self.ask_min_h, int(total_h) - 320))
        if self.ask_h > limit:
            self._set_ask_h(limit)

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
        """窗口最小化时自动暂停检测，并显示悬浮的鲸鱼娘桌宠。"""
        if event is not None and event.widget is not self.root:
            return
        try:
            if self.root.state() == "iconic":
                if self.listening:
                    self.toggle_listen()
                    self.status.config(text="窗口已最小化 · 检测已暂停")
                if self.cfg.get("pet_enabled") and not self.pet_exited:
                    self._show_pet()
                self._sync_pet_base()
        except Exception:
            log("最小化处理失败\n" + traceback.format_exc())

    # ------------------------------------------------------------ 桌宠
    PET_MIN = 48
    PET_MAX = 320
    PET_STEP = 16

    def _clamp_pet_pos(self, x: float, y: float, size: int) -> tuple:
        """把桌宠姿态图左上角限制在屏幕内，保证整个窗口（含跳跃留白）不越界。

        返回的是姿态图左上角坐标；窗口原点需再减去 pet_layout 的边距。
        屏幕分辨率变化、或换机后旧坐标超出范围时，这个方法保证桌宠仍可见。
        """
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        try:
            w, h, lm, tm = pet_layout(size)
        except Exception:
            w = h = size
            lm = tm = 0
        x = max(lm, min(int(x), max(lm, sw - w + lm - 8)))
        y = max(tm, min(int(y), max(tm, sh - h + tm - 8)))
        return x, y

    def _pet_loader(self, name: str):
        """桌宠姿态图：pet/<状态>.png；某个姿态图缺失时由引擎回退到顶栏形象图。"""
        from PIL import Image

        if name == "fallback":
            return Image.open(resource_path("app_mascot.png"))
        return Image.open(resource_path(f"pet/{name}.png"))

    def _pet_trigger(self, state: str, hold: float = None) -> None:
        """触发桌宠的一个临时形态（吓一跳 / 捧元宝），到时自动回到基础形态。"""
        if self.pet_anim is not None:
            self.pet_anim.trigger(state, hold)

    def _sync_pet_base(self) -> None:
        """桌宠的基础形态跟随主窗口：最小化时发呆，否则认真观察。"""
        idle = self.root.state() == "iconic"
        if self.pet_anim is not None:
            self.pet_anim.set_base("idle" if idle else "observe")
        self._schedule_egg()

    def _schedule_egg(self) -> None:
        """发呆彩蛋：主窗口最小化期间，随机 2 分钟到 1 小时后冒一句「有认真学习吗」。"""
        self._cancel_egg()
        if self.pet is None or self.pet_anim is None or self.root.state() != "iconic":
            return
        self._egg_after = self.root.after(random.randint(120, 3600) * 1000, self._egg_fire)

    def _cancel_egg(self) -> None:
        if self._egg_after is not None:
            try:
                self.root.after_cancel(self._egg_after)
            except Exception:
                pass
            self._egg_after = None

    def _egg_fire(self) -> None:
        self._egg_after = None
        if self.pet is None or self.root.state() != "iconic":
            return
        self._show_balance_bubble("有认真学习吗", stay_ms=4000)
        self._schedule_egg()

    def _pet_motion(self, event) -> None:
        """鼠标在桌宠身上移动（未按下）：盯着鼠标方向看。"""
        if self.pet_anim is None or self.pet_drag:
            return
        if self.pet_anim.alpha_at(event.x, event.y) < 24:
            self.pet_anim.hover(False)
            return
        self.pet_anim.hover(True, event.x, event.y)

    def _pet_leave(self, event) -> None:
        if self.pet_anim is not None:
            self.pet_anim.hover(False)

    def _show_pet(self) -> None:
        """最小化后显示桌宠：可拖动换位置、滚轮改大小，单击恢复主窗口。"""
        if self.pet is not None:
            try:
                self.pet.attributes("-topmost", False)
                self.pet.deiconify()
                self.pet.lower()
                return
            except Exception:
                self.pet = None
        if PetAnimator is None or pet_layout is None:
            log("桌宠动画引擎不可用，桌宠不显示")
            return
        try:
            self.pet_size = max(
                self.PET_MIN, min(self.PET_MAX, int(self.cfg.get("pet_size") or 128))
            )
        except Exception:
            self.pet_size = 128
        size = self.pet_size
        w, h, lm, tm = pet_layout(size)

        win = tk.Toplevel(self.root)
        win.withdraw()  # 先隐藏，等画面与位置就绪后再显示，避免闪动
        win.overrideredirect(True)
        win.attributes("-topmost", False)  # 桌宠只待在桌面层，不压住视频等窗口
        try:
            self.pet_anim = PetAnimator(
                win,
                self._pet_loader,
                log=lambda: log("桌宠动画帧渲染失败\n" + traceback.format_exc()),
            )
            self.pet_anim.set_size(size)
        except Exception:
            self.pet_anim = None
            win.destroy()
            log("桌宠图片加载失败\n" + traceback.format_exc())
            return

        win.bind("<ButtonPress-1>", self._pet_press)
        win.bind("<B1-Motion>", self._pet_move)
        win.bind("<ButtonRelease-1>", self._pet_release)
        win.bind("<MouseWheel>", self._pet_wheel)
        win.bind("<Button-3>", self._pet_menu)
        win.bind("<Motion>", self._pet_motion)
        win.bind("<Leave>", self._pet_leave)

        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        pos = self.cfg.get("pet_pos")
        if isinstance(pos, list) and len(pos) == 2:
            x, y = self._clamp_pet_pos(pos[0], pos[1], size)
        else:
            x, y = self._clamp_pet_pos(sw - size - 36, sh - size - 90, size)
        # 窗口比姿态图大一圈（上方留给跳跃），所以窗口原点要往上、往左挪
        win.geometry(f"{w}x{h}+{x - lm}+{y - tm}")

        win.update_idletasks()
        self.pet = win
        self._sync_pet_base()
        self.pet_anim.start()  # 第一帧先画好，再显示窗口
        # 解除与主窗口的 owner 关系，避免主窗口最小化时桌宠被系统一并隐藏
        try:
            hwnd = self._hwnd(win)
            if hwnd:
                ctypes.windll.user32.SetWindowLongPtrW(hwnd, -8, 0)
        except Exception:
            pass
        win.deiconify()
        win.lower()  # 沉到普通窗口之下，回到桌面时仍然可见

    def _pet_rect(self) -> tuple:
        """桌宠姿态图所在方块的实际外框（物理像素），避免 Tk 位置滞后。
        窗口比姿态图大一圈，这里换算回姿态图的左上角：拖动、位置记忆、气泡定位都按它算。"""
        size = self.pet_size
        _, _, lm, tm = pet_layout(size)
        try:
            r = wintypes.RECT()
            hwnd = self._hwnd(self.pet)
            if hwnd and ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r)):
                return r.left + lm, r.top + tm, r.left + lm + size, r.top + tm + size
        except Exception:
            pass
        x, y = self.pet.winfo_x(), self.pet.winfo_y()
        return x + lm, y + tm, x + lm + size, y + tm + size

    def _pet_press(self, event) -> None:
        rx, ry, _, _ = self._pet_rect()
        self.pet_drag = {
            "x": event.x_root,
            "y": event.y_root,
            "wx": rx,
            "wy": ry,
            "moved": False,
        }
        if self.pet_anim is not None:
            self.pet_anim.hover(False)
        try:
            event.widget.grab_set()  # 拖动时持续捕获鼠标，快速拖动也不丢事件
        except Exception:
            pass

    def _pet_move(self, event) -> None:
        d = self.pet_drag
        if not d or self.pet is None:
            return
        dx, dy = event.x_root - d["x"], event.y_root - d["y"]
        if abs(dx) + abs(dy) <= 3:
            return
        d["moved"] = True
        x, y = self._clamp_pet_pos(d["wx"] + dx, d["wy"] + dy, self.pet_size)
        try:
            _, _, lm, tm = pet_layout(self.pet_size)
            self.pet.geometry(f"+{x - lm}+{y - tm}")
            self._move_balance_bubble(x, y)
        except Exception:
            pass

    def _pet_release(self, event) -> None:
        d = self.pet_drag
        self.pet_drag = None
        try:
            event.widget.grab_release()
        except Exception:
            pass
        if not d:
            return
        if not d["moved"]:
            self._pet_trigger("startle")
            self._restore_from_pet()
            return
        try:
            rx, ry, _, _ = self._pet_rect()
            self.cfg["pet_pos"] = [rx, ry]
            save_config(self.cfg)
        except Exception:
            pass

    def _pet_wheel(self, event) -> None:
        if self.pet is None or not event.delta:
            return
        step = self.PET_STEP if event.delta > 0 else -self.PET_STEP
        new_size = max(self.PET_MIN, min(self.PET_MAX, self.pet_size + step))
        if new_size != self.pet_size:
            self._resize_pet(new_size)

    def _resize_pet(self, new_size: int) -> None:
        """滚轮缩放桌宠，以中心为锚点并保持在屏幕内。"""
        if self.pet is None or self.pet_anim is None or pet_layout is None:
            return
        try:
            rx, ry, rr, rb = self._pet_rect()
            cx, cy = (rx + rr) // 2, (ry + rb) // 2
            x, y = self._clamp_pet_pos(cx - new_size // 2, cy - new_size // 2, new_size)
            w, h, lm, tm = pet_layout(new_size)
            self.pet_size = new_size
            self.pet_anim.set_size(new_size)
            self.pet.geometry(f"{w}x{h}+{x - lm}+{y - tm}")
            self.pet.update_idletasks()
            self._move_balance_bubble(x, y)
        except Exception:
            log("桌宠缩放失败\n" + traceback.format_exc())
            return
        self.cfg["pet_size"] = new_size
        self.cfg["pet_pos"] = [x, y]
        save_config(self.cfg)

    # ------------------------------------------------------------ 余额
    def _pet_menu(self, event) -> None:
        """右键桌宠弹出的菜单。"""
        if self.pet is None:
            return
        menu = None
        try:
            menu = tk.Menu(
                self.root,
                tearoff=0,
                bg=CARD,
                fg=FG,
                activebackground=SEL,
                activeforeground=FG,
                font=(FONT, 9),
                bd=0,
            )
            menu.add_command(label="显示余额", command=self._start_balance_query)
            menu.add_command(label="退出", command=self._exit_pet)
            self._pet_menu_obj = menu
            menu.tk_popup(event.x_root, event.y_root)
        except Exception:
            log("桌宠菜单显示失败\n" + traceback.format_exc())
        finally:
            if menu is not None:
                try:
                    menu.grab_release()
                except Exception:
                    pass

    def _start_balance_query(self) -> None:
        """查询 DeepSeek 账户余额，在桌宠左上角用气泡显示。"""
        if self.balance_loading:
            return
        self._pet_trigger("money")
        key = (self.cfg.get("api_key") or "").strip()
        if not key:
            self._show_balance_bubble("请先在设置中填写 API Key")
            return
        if "api.deepseek.com" not in str(self.cfg.get("base_url") or ""):
            self._show_balance_bubble("当前接口无法查询余额\n仅支持 DeepSeek 官方接口")
            return
        self.balance_loading = True
        self._show_balance_bubble("正在查询余额…")

        def work() -> None:
            msg = ""
            try:
                resp = requests.get(
                    "https://api.deepseek.com/user/balance",
                    headers={
                        "Accept": "application/json",
                        "Authorization": f"Bearer {key}",
                    },
                    timeout=12,
                )
                if resp.status_code == 401:
                    msg = "查询失败：API Key 无效"
                elif resp.status_code != 200:
                    msg = f"查询失败（HTTP {resp.status_code}）"
                else:
                    data = resp.json()
                    parts = []
                    order = {"CNY": 0, "USD": 1}
                    infos = sorted(
                        data.get("balance_infos") or [],
                        key=lambda it: order.get(
                            str(it.get("currency") or "").upper(), 9
                        ),
                    )
                    for item in infos:
                        cur = str(item.get("currency") or "").upper()
                        sym = {"CNY": "¥", "USD": "$"}.get(cur, f"{cur} ")
                        parts.append(f"{sym}{item.get('total_balance', '0')}")
                    if not parts:
                        msg = "余额信息为空"
                    else:
                        msg = "余额：" + "   ".join(parts)
                        if not data.get("is_available", True):
                            msg += "\n余额不足，API 调用可能失败"
            except requests.Timeout:
                msg = "查询失败：请求超时"
            except Exception as exc:
                msg = f"查询失败：{str(exc)[:80]}"
            self.msg_q.put(("balance", 0, msg))

        threading.Thread(target=work, daemon=True).start()

    def _finish_balance(self, msg: str) -> None:
        self.balance_loading = False
        self._pet_trigger("money")
        if self.pet is None:
            return
        self._show_balance_bubble(msg)

    def _show_balance_bubble(self, text: str, stay_ms: int = 3000) -> None:
        """透明圆角气泡，定位在桌宠左上角。"""
        self._hide_balance_bubble()
        if self.pet is None:
            return
        try:
            import tkinter.font as tkfont

            font = tkfont.Font(family=FONT, size=10)
            lines = text.split("\n")
            tw = max(font.measure(line) for line in lines)
            th = font.metrics("linespace") * len(lines)
            bw, bh = tw + 28, th + 18
            tail = 10
            key = "#ff00fe"
            cw, ch = bw + 2, bh + tail + 2
            win = tk.Toplevel(self.root)
            win.withdraw()
            win.overrideredirect(True)
            win.attributes("-topmost", False)
            win.configure(bg=key)
            try:
                win.attributes("-transparentcolor", key)
            except Exception:
                pass
            cv = tk.Canvas(win, width=cw, height=ch, bg=key, highlightthickness=0, bd=0)
            cv.pack()
            r = 10
            x0, y0, x1, y1 = 1, 1, bw - 1, bh - 1
            pts = [
                x0 + r, y0, x1 - r, y0, x1 - r, y0, x1, y0,
                x1, y0 + r, x1, y1 - r, x1, y1 - r, x1, y1,
                x1 - r, y1, x0 + r, y1, x0 + r, y1, x0, y1,
                x0, y1 - r, x0, y0 + r, x0, y0 + r, x0, y0,
            ]
            cv.create_polygon(
                pts, smooth=True, splinesteps=24, fill=CARD, outline=BORDER, width=1
            )
            tx = x1 - 26
            cv.create_polygon(
                tx, y1 - 1, tx + 14, y1 - 1, x1 - 6, y1 + tail,
                fill=CARD, outline="",
            )
            cv.create_text(
                (x0 + x1) // 2, (y0 + y1) // 2,
                text=text, fill=FG, font=font, justify="left",
            )
            cv.bind("<Button-1>", lambda e: self._hide_balance_bubble())
            px, py, _, _ = self._pet_rect()
            bx = px + 14 - cw
            by = py - ch - 2
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            bx = max(4, min(bx, sw - cw - 4))
            by = max(4, min(by, sh - ch - 4))
            win.geometry(f"{cw}x{ch}+{bx}+{by}")
            win.update_idletasks()
            win.deiconify()
            win.lift()
            try:
                win.attributes("-alpha", 1.0)
            except Exception:
                pass
            self.balance_bubble = win
            self._bubble_size = (cw, ch)
            self._fade_start = None
            self._bubble_after = self.root.after(
                max(200, int(stay_ms) - 1000), self._begin_bubble_fade
            )
        except Exception:
            log("余额气泡显示失败\n" + traceback.format_exc())

    def _begin_bubble_fade(self) -> None:
        self._fade_start = time.monotonic()
        self._fade_balance_bubble()

    def _fade_balance_bubble(self) -> None:
        """按真实时间以约 60fps 平滑淡出，总时长 1 秒。"""
        win = self.balance_bubble
        if win is None:
            return
        start = self._fade_start if self._fade_start is not None else time.monotonic()
        t = min(1.0, max(0.0, (time.monotonic() - start) / 1.0))
        smooth = t * t * (3.0 - 2.0 * t)  # smoothstep：首尾更柔和
        try:
            win.attributes("-alpha", max(0.0, 1.0 - smooth))
        except Exception:
            pass
        if t < 1.0:
            self._bubble_after = self.root.after(16, self._fade_balance_bubble)
        else:
            self._bubble_after = None
            self._hide_balance_bubble()

    def _move_balance_bubble(self, px: int = None, py: int = None) -> None:
        """让气泡跟随桌宠移动，始终保持在桌宠左上角。"""
        win = self.balance_bubble
        cw, ch = self._bubble_size
        if win is None or cw <= 0 or ch <= 0:
            return
        try:
            if px is None or py is None:
                px, py, _, _ = self._pet_rect()
            bx = px + 14 - cw
            by = py - ch - 2
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            bx = max(4, min(bx, sw - cw - 4))
            by = max(4, min(by, sh - ch - 4))
            win.geometry(f"+{bx}+{by}")
        except Exception:
            pass

    def _hide_balance_bubble(self) -> None:
        self._fade_start = None
        if self._bubble_after is not None:
            try:
                self.root.after_cancel(self._bubble_after)
            except Exception:
                pass
            self._bubble_after = None
        if self.balance_bubble is not None:
            try:
                self.balance_bubble.destroy()
            except Exception:
                pass
            self.balance_bubble = None

    def _hide_pet(self) -> None:
        self._cancel_egg()
        if self.pet_anim is not None:
            self.pet_anim.destroy()
            self.pet_anim = None
        self._hide_balance_bubble()
        if self.pet is not None:
            try:
                self.pet.grab_release()
            except Exception:
                pass
            try:
                self.pet.destroy()
            except Exception:
                pass
        self.pet = None
        self.pet_photo = None
        self.pet_drag = None

    def _restore_from_pet(self) -> None:
        """点击桌宠：立刻恢复主窗口，桌宠保持常驻。"""
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            log("恢复主窗口失败\n" + traceback.format_exc())

    def _on_restore(self, event=None) -> None:
        """主窗口从最小化恢复：桌宠重新出现（若已启用）。"""
        if event is not None and event.widget is not self.root:
            return
        if self.root.state() == "iconic":
            return
        self.pet_exited = False
        if self.cfg.get("pet_enabled"):
            self.root.after(200, self._show_pet)
        else:
            self._hide_pet()
        self._sync_pet_base()

    def _exit_pet(self) -> None:
        """右键菜单「退出」：本次会话先隐藏桌宠，主窗口恢复时再出现。"""
        self.pet_exited = True
        self._hide_pet()

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

    def _ask_grip_press(self, event):
        """按住提问框下边缘时，进入临时加高模式。"""
        w = event.widget
        try:
            near_bottom = event.y >= w.winfo_height() - 8
        except Exception:
            near_bottom = False
        if not near_bottom:
            return None
        self.ask_grip = {"y": event.y_root, "h": self.ask_h, "grab": w}
        try:
            w.grab_set()  # 拖出边缘也能持续收到移动/松开事件
        except Exception:
            pass
        return "break"

    def _ask_grip_move(self, event):
        if not self.ask_grip:
            return None
        new_h = self.ask_grip["h"] + (event.y_root - self.ask_grip["y"])
        self._set_ask_h(new_h, self.root.winfo_height(), allow_expand=True)
        return "break"

    def _ask_grip_release(self, event):
        g = self.ask_grip
        if not g:
            return None
        self.ask_grip = None
        try:
            g["grab"].grab_release()
        except Exception:
            pass
        return "break"

    def _ask_grip_cursor(self, event):
        w = event.widget
        try:
            near_bottom = event.y >= w.winfo_height() - 8
        except Exception:
            near_bottom = False
        try:
            if self.ask_grip or near_bottom:
                w.configure(cursor="sb_v_double_arrow")
            else:
                w.configure(cursor="xterm" if isinstance(w, tk.Text) else "")
        except Exception:
            pass

    def _ask_grip_leave(self, event):
        if self.ask_grip:
            return
        try:
            event.widget.configure(cursor="xterm" if isinstance(event.widget, tk.Text) else "")
        except Exception:
            pass

    def _ask_wheel(self, event):
        """提问框内容超出可视区域时，用滚轮上下查看。"""
        try:
            self.txt_sel.yview_scroll(-3 if event.delta > 0 else 3, "units")
        except Exception:
            pass
        return "break"

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
        # 联网搜索：仅 DeepSeek 官方接口支持 deepseek_web_search 服务端工具
        web = bool(cfg.get("web_search")) and "api.deepseek.com" in base

        # 检索文档中最相关的片段，让回答优先依据文档内容
        self._refresh_doc_if_needed()
        doc_chunks = self._pick_doc_chunks(text)
        user_content = ("用户的问题或选中的文字：\n" if web else "选中文本：\n") + cut + "\n\n"
        if doc_chunks:
            doc_name = self.doc_path.name if self.doc_path else "参考文档"
            user_content += (
                f"参考文档《{doc_name}》中的相关片段（按原文顺序）：\n<<<\n"
                + "\n\n……\n\n".join(doc_chunks)
                + "\n>>>\n\n"
                "请优先依据上面的文档片段解释选中文本；如果文档片段没有覆盖相关内容，"
                "可以用通用知识补充，但要说明“文档中未涉及”。\n\n"
            )
        user_content += (
            "请直接回答；若这段文字本身需要解释，就解释它。"
            if web
            else "请解释这段内容。"
        )

        system_prompt = cfg.get("system_prompt") or DEFAULT_CONFIG["system_prompt"]
        if web:
            now = datetime.now()
            week = "星期" + "一二三四五六日"[now.weekday()]
            system_prompt += (
                f"\n\n已开启联网搜索。当前本机时间：{now:%Y-%m-%d %H:%M:%S}（{week}，北京时间）。"
                "涉及时间或实时信息时，以此时间为准，搜索结果的发布时间只用于判断信息新旧；"
                "如果用户是在提问，请直接回答，不要只解释问题本身。"
            )

        messages = [{"role": "system", "content": system_prompt}]
        if web:
            # DeepSeek 用户元信息标记：让模型知道今天的日期与所在地
            messages.append(
                {
                    "role": "latest_reminder",
                    "content": f"{now:%Y-%m-%d},{week},中国,Web,中文",
                }
            )
        messages.append({"role": "user", "content": user_content})

        payload = {
            "model": cfg.get("model") or "deepseek-flash",
            "messages": messages,
            "stream": True,
            "temperature": 0.3,
        }
        # DeepSeek 新模型默认开启思考模式；划词解释求快，这里显式关闭
        if "api.deepseek.com" in base:
            payload["thinking"] = {"type": "disabled"}
        if web:
            payload["tools"] = [{"type": "deepseek_web_search"}]
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
                elif kind == "balance":
                    self._finish_balance(data)
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
        self._apply_icon(win)
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
        self._apply_icon(win)

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

        v_web = tk.BooleanVar(value=bool(self.cfg.get("web_search")))
        tk.Checkbutton(
            win,
            text="联网搜索（实时信息，DeepSeek 官方接口，额外消耗 token）",
            variable=v_web,
            bg=BG,
            fg=SUB,
            activebackground=BG,
            activeforeground=FG,
            selectcolor=CARD,
            font=(FONT, 9),
            bd=0,
            highlightthickness=0,
        ).pack(anchor="w", padx=14, pady=(8, 0))

        v_pet = tk.BooleanVar(value=bool(self.cfg.get("pet_enabled")))
        tk.Checkbutton(
            win,
            text="开启桌宠（桌面悬浮）",
            variable=v_pet,
            bg=BG,
            fg=SUB,
            activebackground=BG,
            activeforeground=FG,
            selectcolor=CARD,
            font=(FONT, 9),
            bd=0,
            highlightthickness=0,
        ).pack(anchor="w", padx=14, pady=(6, 0))

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
            self.cfg["web_search"] = bool(v_web.get())
            self.cfg["pet_enabled"] = bool(v_pet.get())
            self.pet_exited = False
            if self.cfg["pet_enabled"]:
                self.root.after(120, self._show_pet)
            else:
                self._hide_pet()
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
