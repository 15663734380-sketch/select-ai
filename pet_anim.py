# -*- coding: utf-8 -*-
"""鲸鱼娘桌宠动画引擎（只负责"长什么样、怎么动"，不含任何交互逻辑）。

思路：
  1. 每个状态一张透明底姿态图（idle 发呆 / observe 认真观察 / startle 吓一跳 / watch 盯着看 / money 捧金元宝）；
  2. 状态内部用连续网格变形（Image.MESH 亚像素级双三次/双线性采样）让图"活"起来：
     胸腔自然呼吸起伏、脊柱统一节律的滞后摆动、头部绕颈部平滑歪头与 2.5D 视差转向、
     弹跳与落地缓冲挤压拉伸，全程无硬切线，再叠加随形变同步跟随的手绘小特效；
  3. 任意两个状态之间做加权交叉淡入（预乘 alpha 混合，边缘不发虚），中途反向会接着当前进度倒回去，不重头播；
  4. 画面用逐像素透明的分层窗口输出（UpdateLayeredWindow），旋转缩放都不会出现锯齿和杂色边。
"""
import ctypes
import math
import time
from ctypes import wintypes

from PIL import Image, ImageDraw, ImageFilter

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
user32.GetAncestor.restype = wintypes.HWND
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
user32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetDC.restype = wintypes.HDC
user32.GetDC.argtypes = [wintypes.HWND]
user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC
gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteDC.argtypes = [wintypes.HDC]

WS_EX_LAYERED = 0x00080000
GWL_EXSTYLE = -20
ULW_ALPHA = 0x02
TWO_PI = 2.0 * math.pi

CARD = "#fbfdff"
BORDER = "#b7d7f5"
FG = "#1f3a63"
DOT_OFF = "#bfd6ec"
GOLD = "#f3c13a"


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 1)]


class BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_byte), ("BlendFlags", ctypes.c_byte),
        ("SourceConstantAlpha", ctypes.c_byte), ("AlphaFormat", ctypes.c_byte),
    ]


user32.UpdateLayeredWindow.restype = wintypes.BOOL
user32.UpdateLayeredWindow.argtypes = [
    wintypes.HWND, wintypes.HDC, ctypes.POINTER(wintypes.POINT),
    ctypes.POINTER(wintypes.SIZE), wintypes.HDC, ctypes.POINTER(wintypes.POINT),
    wintypes.COLORREF, ctypes.POINTER(BLENDFUNCTION), wintypes.DWORD,
]
gdi32.CreateDIBSection.restype = wintypes.HBITMAP
gdi32.CreateDIBSection.argtypes = [
    wintypes.HDC, ctypes.POINTER(BITMAPINFO), wintypes.UINT,
    ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD,
]


def layout(size: int):
    """窗口画布布局：返回 (窗口宽, 窗口高, 左边距, 上边距)。

    size×size 的姿态图放在 (左边距, 上边距) 处；上方留出跳跃、挤压拉伸和倾斜的空间，
    四周全是透明像素，不占视觉。
    """
    side = int(round(size * 0.08))
    top = int(round(size * 0.52))
    bottom = int(round(size * 0.04))
    return size + 2 * side, size + top + bottom, side, top


def _alpha_resize(im, size):
    """带 alpha 权重的缩放（先预乘再缩），透明边缘不会发黑边。"""
    return im.convert("RGBa").resize(size, Image.LANCZOS).convert("RGBA")


def _bleed(im):
    """把不透明像素的颜色向透明区域"渗透"。之后的旋转/缩放做双线性插值时，
    边缘才不会混进透明像素里的黑色而出现黑边。"""
    alpha = im.getchannel("A")
    out = im
    for radius in (2, 6, 16):
        soft = out.convert("RGBa").filter(ImageFilter.GaussianBlur(radius)).convert("RGBA")
        holes = out.getchannel("A").point(lambda v: 255 if v == 0 else 0)
        rgb = Image.composite(soft.convert("RGB"), out.convert("RGB"), holes)
        out = Image.merge("RGBA", (*rgb.split(), out.getchannel("A")))
    return Image.merge("RGBA", (*out.convert("RGB").split(), alpha))


def _fade(im, k):
    if k >= 0.999:
        return im
    a = im.getchannel("A").point(lambda v: int(v * k))
    return Image.merge("RGBA", (*im.convert("RGB").split(), a))


def _hex(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _star(radius, colour, core=True):
    """四角星光（4 倍超采样后缩小，边缘平滑）。"""
    k = radius * 0.30
    pts = [(0, -radius), (k, -k), (radius, 0), (k, k), (0, radius), (-k, k), (-radius, 0), (-k, -k)]
    side = int(radius * 2.2) * 4
    c = side / 2.0
    big = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    d = ImageDraw.Draw(big)
    d.polygon([(c + x * 4, c + y * 4) for x, y in pts], fill=_hex(colour) + (255,))
    if core:
        rc = radius * 0.22 * 4
        d.ellipse([c - rc, c - rc, c + rc, c + rc], fill=(255, 255, 255, 255))
    return _alpha_resize(big, (side // 4, side // 4))


def _thought_bubble(w, h, active, fade):
    """发呆时的"…"思考小泡：三个圆点依次亮起。"""
    big = Image.new("RGBA", (w * 4, h * 4), (0, 0, 0, 0))
    d = ImageDraw.Draw(big)
    d.rounded_rectangle(
        [2, 2, w * 4 - 3, h * 4 - 3], radius=int(h * 4 * 0.5),
        fill=_hex(CARD) + (255,), outline=_hex(BORDER) + (255,), width=4,
    )
    r = h * 0.19 * 4
    cy = h * 0.5 * 4
    for i in range(3):
        cx = w * (i + 1) / 4.0 * 4
        on = i == active
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=_hex(FG if on else DOT_OFF) + (255,))
    return _fade(_alpha_resize(big, (w, h)), fade)


def _smoothstep(p):
    p = min(1.0, max(0.0, p))
    return p * p * (3.0 - 2.0 * p)


def _smooth_range(v, edge0, edge1):
    """在 [edge0, edge1] 之间从 0 平滑过渡到 1（三次埃尔米特插值，两端导数为 0）。"""
    span = edge1 - edge0
    if span == 0.0:
        return 0.0
    return _smoothstep((v - edge0) / span)


class PetAnimator:
    """把姿态图渲染成分层窗口的连续帧。状态切换调用 set_base / trigger / hover。"""

    STATES = ("idle", "observe", "startle", "watch", "money")
    HOLD = {"startle": 1.25, "money": 4.6}
    FRAME_MS = 20          # 目标 ~50fps，动作与鼠标跟随更丝滑
    MESH_COLS = 12
    MESH_ROWS = 16
    CX = 0.49              # 角色身体/面部横向中心（归一化坐标）
    FEET_Y = 0.947         # 脚底接地锚点（五张姿态图实测均为 0.947）
    NECK_Y = 0.56          # 颈部枢轴高度

    def __init__(self, win, loader, log=None):
        self.win = win
        self.loader = loader      # loader(name) -> PIL.Image RGBA (512x512)
        self.log = log
        self.size = 128
        self.masters = {}
        self.small = {}
        self.fx = {}
        self.base = "idle"
        self.cur = "idle"
        self.prev = None
        self.t_switch = time.monotonic()
        self.dur = 0.0
        self.entered = {"idle": self.t_switch}
        self.overlay = None
        self.overlay_until = 0.0
        self.hovering = False
        self.gaze = (0.0, 0.0)
        self.gaze_vel = (0.0, 0.0)
        self.gaze_target = (0.0, 0.0)
        self.last = None
        self.last_alpha = None
        self._after = None
        self._last_t = self.t_switch
        self._err = 0
        self.hwnd = 0
        self._dib = None
        self._load_masters()

    # ------------------------------------------------------------ 资源
    def _load_masters(self):
        for name in self.STATES:
            try:
                img = self.loader(name)
                img = img.convert("RGBA") if img is not None else None
            except Exception:
                img = None
            if img is None:
                try:
                    img = self.loader("fallback")
                    img = img.convert("RGBA") if img is not None else None
                except Exception:
                    img = None
            if img is None:
                continue
            try:
                if img.size != (512, 512):
                    img = _alpha_resize(img, (512, 512))
                self.masters[name] = _bleed(img)
            except Exception:
                self.masters.pop(name, None)
        if not self.masters:
            raise RuntimeError("no pet sprite available")

    def set_size(self, size):
        """滚轮改变桌宠大小：按新尺寸重做小图（预乘 alpha 并加 2px 透明安全边）和特效缓存。"""
        size = int(size)
        self.size = size
        self.small = {}
        for name, im in self.masters.items():
            rs = im.convert("RGBa").resize((size, size), Image.LANCZOS)
            padded = Image.new("RGBa", (size + 4, size + 4), (0, 0, 0, 0))
            padded.paste(rs, (2, 2))
            self.small[name] = padded
        self.fx = {}
        self._free_dib()

    def layout(self):
        return layout(self.size)

    # ------------------------------------------------------------ 状态
    def set_base(self, state):
        """基础状态：主界面最小化时为 idle，否则为 observe。"""
        if state not in self.STATES or state == self.base:
            return
        self.base = state

    def trigger(self, state, hold=None):
        """触发一个临时状态（吓一跳 / 捧元宝），持续 hold 秒后自动回到基础状态。"""
        if state not in self.STATES:
            return
        self.overlay = state
        self.overlay_until = time.monotonic() + float(hold or self.HOLD.get(state, 1.0))

    def hover(self, active, x=None, y=None):
        """鼠标在桌宠上移动时调用；x, y 是相对窗口的像素坐标，用来算盯看方向。"""
        self.hovering = bool(active)
        if active and x is not None and y is not None:
            w, h, lm, tm = layout(self.size)
            s = self.size
            dx = (x - lm - s / 2.0) / (s * 0.9)
            dy = (y - tm - s * 0.40) / (s * 0.9)
            self.gaze_target = (max(-1.0, min(1.0, dx)), max(-1.0, min(1.0, dy)))
        elif not active:
            self.gaze_target = (0.0, 0.0)

    def alpha_at(self, x, y):
        """窗口上某点是否画着角色（透明像素上不算"在桌宠身上"）。"""
        if self.last_alpha is None:
            return 0
        try:
            if 0 <= x < self.last_alpha.size[0] and 0 <= y < self.last_alpha.size[1]:
                return self.last_alpha.getpixel((int(x), int(y)))
        except Exception:
            pass
        return 0

    def _dur(self, a, b):
        if b == "startle":
            return 0.10
        if a == "startle":
            return 0.40
        if b == "watch":
            return 0.30
        if a == "watch":
            return 0.35
        if b == "money":
            return 0.30
        if a == "money":
            return 0.45
        return 0.70

    def _step(self, now):
        if self.overlay is not None and now >= self.overlay_until:
            self.overlay = None
        desired = "watch" if self.hovering else self.base
        if self.overlay is not None:
            desired = self.overlay
        # 盯看方向用临界阻尼弹簧平滑追踪，起步柔和、停止无过冲
        dt = min(0.05, max(0.0, now - self._last_t))
        if dt > 0.0:
            omega = 12.5
            e = math.exp(-omega * dt)
            gx, gy = self.gaze
            vx, vy = self.gaze_vel
            tx, ty = self.gaze_target
            dx, dy = gx - tx, gy - ty
            nx = tx + (dx + (vx + omega * dx) * dt) * e
            ny = ty + (dy + (vy + omega * dy) * dt) * e
            nvx = (vx - omega * (vx + omega * dx) * dt) * e
            nvy = (vy - omega * (vy + omega * dy) * dt) * e
            self.gaze = (nx, ny)
            self.gaze_vel = (nvx, nvy)
        self._last_t = now
        if desired != self.cur:
            if self.prev == desired:
                # 正在 A→B 的半路上又要回 A：从当前进度接着倒回去
                p = _smoothstep((now - self.t_switch) / self.dur) if self.dur else 1.0
                self.prev, self.cur = self.cur, self.prev
                self.dur = self._dur(self.prev, self.cur)
                self.t_switch = now - (1.0 - p) * self.dur
            else:
                self.prev = self.cur
                self.cur = desired
                self.dur = self._dur(self.prev, self.cur)
                self.t_switch = now
            self.entered[self.cur] = now

    # ------------------------------------------------------------ 连续形变动作
    def _motion_params(self, state, t):
        """计算某状态在 t 秒时的连续全身形变参数（全部以归一化图幅为单位）。

        返回 dict：
          tx, ty         整体平移（以脚底为锚点）
          sx, sy         挤压与拉伸（以脚底为锚点）
          ang            整体倾角（弧度，正为顺时针右倾）
          breath         呼吸节律 [-1, 1]
          breath_chest   胸腔呼吸扩张幅度
          breath_lift    肩颈随呼吸上提幅度
          sway_amp       脊柱左右摆动幅度
          sway_phase     摆动相位
          sway_lag       由脚底到头顶的连续相位滞后（弧度）
          head_tilt      头部绕颈部枢轴的柔和歪头角（弧度）
          head_nod       头部轻微俯仰位移
          gx, gy         面部 2.5D 转向偏移
        """
        p = {
            "tx": 0.0, "ty": 0.0, "sx": 1.0, "sy": 1.0, "ang": 0.0,
            "breath": 0.0, "breath_chest": 0.0, "breath_lift": 0.0,
            "sway_amp": 0.0, "sway_phase": 0.0, "sway_lag": 0.45,
            "head_tilt": 0.0, "head_nod": 0.0, "gx": 0.0, "gy": 0.0,
        }
        if state == "idle":
            bp = TWO_PI * t / 3.8
            breath = 0.85 * math.sin(bp) + 0.15 * math.sin(2.0 * bp - 0.5)
            sp = TWO_PI * t / 5.4
            p["breath"] = breath
            p["breath_chest"] = 0.014
            p["breath_lift"] = 0.008
            p["sx"] = 1.0 - 0.003 * breath
            p["sy"] = 1.0 + 0.006 * breath
            p["ang"] = 0.017 * math.sin(sp - 0.12)
            p["sway_amp"] = 0.013
            p["sway_phase"] = sp
            p["sway_lag"] = 0.52
            p["head_tilt"] = 0.024 * math.sin(sp - 0.42)
            p["head_nod"] = 0.0035 * breath
            return p

        if state == "observe":
            op = TWO_PI * t / 3.6
            lean = 0.5 - 0.5 * math.cos(op)
            scan = math.sin(op * 0.5)
            breath = math.sin(TWO_PI * t / 3.2)
            sc = 1.0 + 0.022 * lean
            p["sx"] = sc * (1.0 - 0.004 * lean)
            p["sy"] = sc * (1.0 + 0.005 * lean)
            p["tx"] = 0.013 * scan
            p["ang"] = 0.019 * scan
            p["breath"] = breath
            p["breath_chest"] = 0.010
            p["breath_lift"] = 0.005
            p["sway_amp"] = 0.011
            p["sway_phase"] = op * 0.5
            p["sway_lag"] = 0.34
            p["head_tilt"] = 0.022 * scan + 0.011 * math.sin(op - 0.35)
            p["head_nod"] = -0.006 * lean
            return p

        if state == "startle":
            def hop(u, span, amp):
                return -amp * math.sin(math.pi * u / span) if 0.0 <= u < span else 0.0

            def squash_at(tc, amp, width):
                d = (t - tc) / width
                return -amp * math.exp(-d * d) if abs(d) < 3.5 else 0.0

            ty = hop(t, 0.40, 0.28) + hop(t - 0.40, 0.24, 0.095) + hop(t - 0.64, 0.18, 0.036)
            # 起跳与腾空拉长，每次落地缓冲压扁
            air_stretch = 0.095 * math.sin(math.pi * min(1.0, t / 0.40)) * math.exp(-1.5 * t) if t < 0.40 else 0.0
            land_sq = (
                squash_at(0.40, 0.11, 0.065)
                + squash_at(0.64, 0.055, 0.055)
                + squash_at(0.82, 0.025, 0.055)
            )
            st = air_stretch + land_sq
            decay = math.exp(-2.6 * t)
            sh_phase = TWO_PI * t / 0.14
            p["ty"] = ty
            p["sy"] = 1.0 + st
            p["sx"] = 1.0 / math.sqrt(max(0.75, p["sy"]))
            p["tx"] = 0.014 * math.sin(sh_phase) * math.exp(-3.4 * t)
            p["ang"] = 0.055 * math.sin(TWO_PI * t / 0.18) * decay
            p["sway_amp"] = 0.018 * decay
            p["sway_phase"] = sh_phase
            p["sway_lag"] = 0.42
            p["head_tilt"] = 0.035 * math.sin(sh_phase - 0.45) * decay
            p["head_nod"] = -0.016 * math.exp(-2.4 * t)
            return p

        if state == "watch":
            nx, ny = self.gaze
            bp = TWO_PI * t / 3.8
            breath = 0.85 * math.sin(bp) + 0.15 * math.sin(2.0 * bp - 0.5)
            sp = TWO_PI * t / 5.4
            p["breath"] = breath
            p["breath_chest"] = 0.012
            p["breath_lift"] = 0.007
            p["sx"] = 1.0 - 0.002 * breath
            p["sy"] = 1.0 + 0.005 * breath
            p["ang"] = 0.036 * nx + 0.008 * math.sin(sp)
            p["sway_amp"] = 0.005
            p["sway_phase"] = sp
            p["sway_lag"] = 0.42
            p["head_tilt"] = 0.040 * nx + 0.010 * math.sin(sp - 0.3)
            p["head_nod"] = 0.016 * ny + 0.003 * breath
            p["gx"] = 0.034 * nx
            p["gy"] = 0.022 * ny
            return p

        if state == "money":
            # 顺滑的正弦幂弹跳：落地速度为 0，无生硬折角，配合 1:2 节拍左右轻晃
            u = (t % 0.76) / 0.76
            bounce = math.sin(math.pi * u) ** 1.35
            swp = TWO_PI * t / 1.52
            stretch = 0.028 * (2.0 * bounce - 1.0)
            p["ty"] = -0.033 * bounce
            p["sy"] = 1.0 + stretch
            p["sx"] = 1.0 - 0.55 * stretch
            p["ang"] = 0.028 * math.sin(swp)
            p["sway_amp"] = 0.010
            p["sway_phase"] = swp
            p["sway_lag"] = 0.35
            p["head_tilt"] = 0.024 * math.sin(swp - 0.25)
            p["breath"] = bounce - 0.5
            p["breath_chest"] = 0.008
            p["breath_lift"] = 0.004
            return p

        breath = math.sin(TWO_PI * t / 3.6)
        p["breath"] = breath
        p["breath_chest"] = 0.010
        p["breath_lift"] = 0.006
        return p

    @classmethod
    def _disp(cls, u, v, m):
        """计算精灵归一化坐标 (u, v) 在形变参数 m 下的前向平滑位移 (dx, dy)。

        全图共享连续光滑的权重场（脊柱悬臂、肩胸呼吸、颈部过渡、面部 2.5D），
        不存在任何硬切线，确保头、长发、道具、手部与裙摆一体自然联动。
        """
        cx = cls.CX
        fy = cls.FEET_Y
        h_up = (0.95 - v) / 0.90
        if h_up < 0.0:
            h_up = 0.0
        elif h_up > 1.0:
            h_up = 1.0

        w_spine = _smooth_range(v, 0.95, 0.22)
        w_head = _smooth_range(v, 0.78, 0.36)
        w_sh = _smooth_range(v, 0.86, 0.44)

        du = u - cx
        dv_c = v - 0.63
        w_chest = math.exp(-(du * du / 0.09 + dv_c * dv_c / 0.04))

        # 1. 局部胸腔呼吸与肩颈上提
        br = m["breath"]
        lx = m["breath_chest"] * br * du * w_chest
        ly = m["breath_chest"] * 0.8 * br * dv_c * w_chest - m["breath_lift"] * br * w_sh

        # 2. 统一节律的脊柱连续滞后摆动（脚底为 0，向上平滑递增并带相位滞后）
        if m["sway_amp"] != 0.0:
            lx += m["sway_amp"] * w_spine * math.sin(m["sway_phase"] - m["sway_lag"] * h_up)

        # 3. 柔和歪头与俯仰（绕颈部平滑旋转，跨越胸肩 0.78→0.36 渐变至 0）
        ht = m["head_tilt"]
        if ht != 0.0:
            c_h = math.cos(ht)
            s_h = math.sin(ht)
            nx = u - cx
            ny = v - cls.NECK_Y
            lx += w_head * (nx * c_h - ny * s_h - nx)
            ly += w_head * (nx * s_h + ny * c_h - ny)
        if m["head_nod"] != 0.0:
            ly += w_head * m["head_nod"]

        # 4. 面部 2.5D 视差转向（watch 盯看鼠标时眼球面部带动头颈平滑转向）
        gx, gy = m["gx"], m["gy"]
        if gx != 0.0 or gy != 0.0:
            dv_f = v - 0.40
            w_face = math.exp(-(du * du / 0.15 + dv_f * dv_f / 0.11))
            w_gaze = 0.55 * w_face + 0.45 * w_head
            lx += gx * w_gaze + 0.35 * gx * w_spine
            ly += gy * w_gaze

        # 5. 全局挤压拉伸、倾角与跳跃位移（绕脚底中心锚点）
        px = cx + (u + lx - cx) * m["sx"]
        py = fy + (v + ly - fy) * m["sy"]
        ang = m["ang"]
        if ang != 0.0:
            ca = math.cos(ang)
            sa = math.sin(ang)
            rx = px - cx
            ry = py - fy
            px = cx + rx * ca - ry * sa
            py = fy + rx * sa + ry * ca
        px += m["tx"]
        py += m["ty"]
        return px - u, py - v

    def _fx_cache(self, name):
        key = (name, self.size)
        hit = self.fx.get(key)
        if hit is not None:
            return hit
        s = self.size
        if name == "idle":  # "…"思考泡：7.2 秒一轮，前 3.6 秒显示，圆点依次亮起
            w, h = max(8, int(s * 0.34)), max(6, int(s * 0.24))
            n = 72
            frames = [_thought_bubble(w, h, int(i * 0.05 / 0.45) % 3,
                                      min(1.0, i * 0.05 / 0.3, (3.6 - i * 0.05) / 0.4)) for i in range(n)]
            data = (7.2, n, frames)
        elif name == "observe":  # 放大镜上的闪光：2.4 秒一轮，亮 0.9 秒
            r = s * 0.085
            frames = [_fade(_star(r * (0.55 + 0.45 * math.sin(math.pi * i / 24.0)), "#fff6c8"),
                            math.sin(math.pi * i / 24.0) ** 0.6) for i in range(24)]
            data = (2.4, 24, frames)
        elif name == "money":  # 元宝周围的金星：1.6 秒一轮，六颗错相闪烁
            r = s * 0.050
            n = 32
            frames = [_fade(_star(r * (0.6 + 0.4 * math.sin(math.pi * i / n)), GOLD),
                            math.sin(math.pi * i / n) ** 0.5) for i in range(n)]
            data = (1.6, n, frames)
        else:
            data = None
        self.fx[key] = data
        return data

    def _compose(self, state, now):
        """用连续网格变形把某个状态在 now 时刻的姿态画进画布（无硬切接缝，亚像素平滑）。"""
        s = self.size
        w, h, lm, tm = layout(s)
        src = self.small.get(state)
        if src is None:
            return Image.new("RGBA", (w, h), (0, 0, 0, 0))

        t = now - self.entered.get(state, now)
        m = self._motion_params(state, t)

        nx, ny = self.MESH_COLS, self.MESH_ROWS
        inv_s = 1.0 / s
        pts = []
        for j in range(ny + 1):
            vy = (j * h / ny - tm) * inv_s
            row = []
            for i in range(nx + 1):
                ux = (i * w / nx - lm) * inv_s
                dx1, dy1 = self._disp(ux, vy, m)
                dx2, dy2 = self._disp(ux - dx1, vy - dy1, m)
                row.append(((ux - dx2) * s + 2.0, (vy - dy2) * s + 2.0))
            pts.append(row)

        mesh = []
        for j in range(ny):
            y0, y1 = int(j * h / ny), int((j + 1) * h / ny)
            r0, r1 = pts[j], pts[j + 1]
            for i in range(nx):
                x0, x1 = int(i * w / nx), int((i + 1) * w / nx)
                u00, v00 = r0[i]
                u01, v01 = r1[i]
                u11, v11 = r1[i + 1]
                u10, v10 = r0[i + 1]
                mesh.append(((x0, y0, x1, y1), (u00, v00, u01, v01, u11, v11, u10, v10)))

        resample = Image.Resampling.BICUBIC if w * h <= 65000 else Image.Resampling.BILINEAR
        canvas = src.transform((w, h), Image.Transform.MESH, mesh, resample=resample).convert("RGBA")
        self._draw_fx(canvas, state, t, lm, tm, m)
        return canvas

    def _draw_fx(self, canvas, state, t, lm, tm, m):
        data = self._fx_cache(state)
        if not data:
            return
        cycle, count, frames = data
        s = self.size
        phase = t % cycle
        if state == "idle":
            if phase < cycle / 2.0:
                im = frames[int(phase / (cycle / 2.0) * count) % count]
                dx, dy = self._disp(0.62, 0.03, m)
                float_y = -0.012 * s * math.sin(math.pi * phase / (cycle / 2.0))
                self._paste(canvas, im, lm + (0.62 + dx) * s, tm + (0.03 + dy) * s + float_y)
        elif state == "observe":
            if phase < 0.9:
                im = frames[int(phase / 0.9 * count) % count]
                dx, dy = self._disp(0.305, 0.365, m)
                self._paste(canvas, im, lm + (0.305 + dx) * s, tm + (0.365 + dy) * s, centre=True)
        elif state == "money":
            idx = int(phase / cycle * count) % count
            spots = ((0.30, 0.66), (0.70, 0.64), (0.22, 0.46), (0.78, 0.44), (0.50, 0.84), (0.64, 0.80))
            for j, (fx, fy) in enumerate(spots):
                im = frames[(idx + j * 5) % count]
                dx, dy = self._disp(fx, fy, m)
                self._paste(canvas, im, lm + (fx + dx) * s, tm + (fy + dy) * s, centre=True)

    @staticmethod
    def _paste(canvas, im, x, y, centre=False):
        if centre:
            x -= im.size[0] / 2.0
            y -= im.size[1] / 2.0
        x, y = int(round(x)), int(round(y))
        if x < 0 or y < 0 or x + im.size[0] > canvas.size[0] or y + im.size[1] > canvas.size[1]:
            return
        canvas.alpha_composite(im, (x, y))

    # ------------------------------------------------------------ 渲染循环
    def _render(self, now):
        b = self._compose(self.cur, now)
        if self.prev is not None:
            p = _smoothstep((now - self.t_switch) / self.dur) if self.dur else 1.0
            a = self._compose(self.prev, now)
            out = Image.blend(a.convert("RGBa"), b.convert("RGBa"), p).convert("RGBA")
            if p >= 1.0:
                self.prev = None
        else:
            out = b
        self.last = out
        self.last_alpha = out.getchannel("A")
        return out

    def _ensure_dib(self, w, h):
        if self._dib is not None and self._dib[0] == (w, h):
            return self._dib
        self._free_dib()
        hdc = user32.GetDC(None)
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = w
        bmi.bmiHeader.biHeight = -h
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bits = ctypes.c_void_p()
        hbmp = gdi32.CreateDIBSection(hdc, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
        user32.ReleaseDC(None, hdc)
        if not hbmp or not bits.value:
            raise OSError("CreateDIBSection failed")
        mdc = gdi32.CreateCompatibleDC(None)
        gdi32.SelectObject(mdc, hbmp)
        self._dib = ((w, h), hbmp, mdc, bits)
        return self._dib

    def _free_dib(self):
        if self._dib is None:
            return
        _, hbmp, mdc, _ = self._dib
        try:
            gdi32.DeleteDC(mdc)
            gdi32.DeleteObject(hbmp)
        except Exception:
            pass
        self._dib = None

    def _push(self, img):
        if not self.hwnd:
            self.hwnd = user32.GetAncestor(int(self.win.winfo_id()), 2)
            if not self.hwnd:
                return
            style = user32.GetWindowLongPtrW(self.hwnd, GWL_EXSTYLE)
            user32.SetWindowLongPtrW(self.hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED)
        rc = wintypes.RECT()
        if not user32.GetClientRect(self.hwnd, ctypes.byref(rc)):
            return
        w, h = rc.right - rc.left, rc.bottom - rc.top
        if w != img.size[0] or h != img.size[1]:
            return  # 窗口正在改尺寸，下一帧再画
        _, _, mdc, bits = self._ensure_dib(w, h)
        pm = img.convert("RGBa")           # 预乘 alpha，分层窗口要求的格式
        r, g, b, a = pm.split()
        data = Image.merge("RGBa", (b, g, r, a)).tobytes()  # BGRA
        ctypes.memmove(bits, data, len(data))
        size = wintypes.SIZE(w, h)
        src = wintypes.POINT(0, 0)
        blend = BLENDFUNCTION(0, 0, 255, 1)  # AC_SRC_OVER, AC_SRC_ALPHA
        user32.UpdateLayeredWindow(self.hwnd, None, None, ctypes.byref(size), mdc,
                                   ctypes.byref(src), 0, ctypes.byref(blend), ULW_ALPHA)

    def start(self):
        if self._after is None:
            self._last_t = time.monotonic()
            self._tick()

    def stop(self):
        if self._after is not None:
            try:
                self.win.after_cancel(self._after)
            except Exception:
                pass
            self._after = None
        self._free_dib()

    def destroy(self):
        self.stop()
        self.last = None
        self.last_alpha = None
        self.masters = {}
        self.small = {}
        self.fx = {}

    def _tick(self):
        self._after = None
        started = time.monotonic()
        try:
            self._step(started)
            self._push(self._render(started))
        except Exception:
            self._err += 1
            if self._err <= 2 and self.log:
                try:
                    self.log()
                except Exception:
                    pass
            if self._err > 60:
                return
        elapsed = int((time.monotonic() - started) * 1000)
        self._after = self.win.after(max(1, self.FRAME_MS - elapsed), self._tick)
