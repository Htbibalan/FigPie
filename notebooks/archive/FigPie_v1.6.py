import io
import json
import math
import base64
import webbrowser
import html
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple, Union, Dict

import fitz
from PIL import Image, ImageTk, ImageGrab, ImageDraw, ImageFont, ImageChops

Image.MAX_IMAGE_PIXELS = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser

APP_TITLE = "FigPie 🍰"
PROJECT_VERSION = 1.6
DEFAULT_CANVAS_W = 1600
DEFAULT_CANVAS_H = 1200
DEFAULT_BG = "white"
DEFAULT_OUTER_MARGIN = 20
DEFAULT_GAP_X = 1
DEFAULT_GAP_Y = 1
DEFAULT_LABEL_FONT_SIZE = 28
DEFAULT_TEXT_FONT_SIZE = 22
DEFAULT_STAT_TEXT_SIZE = 28
DEFAULT_STAT_BRACKET_HEIGHT = 28
DEFAULT_STAT_TEXT_GAP = 10
DEFAULT_LINE_SPACING = 1.15
HANDLE_SIZE = 10
MIN_PANEL_SIZE = 40
SAFE_GAP = 8
DEFAULT_DPI = 600
MIN_ZOOM = 0.25
MAX_ZOOM = 4.0
ZOOM_STEP = 1.15
SHIFT_MASK = 0x0001
CTRL_MASK = 0x0004

try:
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except Exception:
    RESAMPLE_LANCZOS = Image.LANCZOS


def parse_dnd_files(data: str) -> List[str]:
    files = []
    current = ""
    brace = False
    for ch in data:
        if ch == "{":
            brace = True
            current = ""
        elif ch == "}":
            brace = False
            if current:
                files.append(current)
                current = ""
        elif ch == " " and not brace:
            if current:
                files.append(current)
                current = ""
        else:
            current += ch
    if current:
        files.append(current)
    return files


def fit_size_keep_aspect(src_w: int, src_h: int, max_w: int, max_h: int) -> Tuple[int, int]:
    if src_w <= 0 or src_h <= 0:
        return max(1, max_w), max(1, max_h)
    scale = min(max_w / src_w, max_h / src_h)
    return max(1, int(round(src_w * scale))), max(1, int(round(src_h * scale)))


def next_panel_label(index: int) -> str:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    result = ""
    i = index
    while True:
        result = alphabet[i % 26] + result
        i = i // 26 - 1
        if i < 0:
            break
    return result


def list_font_families() -> List[str]:
    root = tk._default_root
    if root is None:
        return ["Arial", "Helvetica", "Times New Roman", "DejaVu Sans"]
    try:
        families = sorted(set(root.tk.call("font", "families")))
    except Exception:
        families = []
    preferred = ["Arial", "Helvetica", "Times New Roman", "Calibri", "DejaVu Sans", "Courier New", "Cascadia Code", "Segoe UI"]
    ordered = [f for f in preferred if f in families] + [f for f in families if f not in preferred]
    return ordered or preferred


def get_font_path_candidates(font_family: str, bold: bool = False, italic: bool = False) -> List[str]:
    mapping = {
        ("Arial", False, False): ["arial.ttf", "Arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"],
        ("Arial", True, False): ["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"],
        ("Arial", False, True): ["ariali.ttf", "Arial Italic.ttf", "LiberationSans-Italic.ttf", "DejaVuSans-Oblique.ttf"],
        ("Arial", True, True): ["arialbi.ttf", "Arial Bold Italic.ttf", "LiberationSans-BoldItalic.ttf", "DejaVuSans-BoldOblique.ttf"],
        ("Helvetica", False, False): ["Helvetica.ttf", "Arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"],
        ("Helvetica", True, False): ["Helvetica-Bold.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"],
        ("Helvetica", False, True): ["Helvetica-Italic.ttf", "Arial Italic.ttf", "LiberationSans-Italic.ttf", "DejaVuSans-Oblique.ttf"],
        ("Helvetica", True, True): ["Helvetica-BoldItalic.ttf", "Arial Bold Italic.ttf", "LiberationSans-BoldItalic.ttf", "DejaVuSans-BoldOblique.ttf"],
        ("Times New Roman", False, False): ["times.ttf", "Times New Roman.ttf", "LiberationSerif-Regular.ttf", "DejaVuSerif.ttf"],
        ("Times New Roman", True, False): ["timesbd.ttf", "Times New Roman Bold.ttf", "LiberationSerif-Bold.ttf", "DejaVuSerif-Bold.ttf"],
        ("Times New Roman", False, True): ["timesi.ttf", "Times New Roman Italic.ttf", "LiberationSerif-Italic.ttf", "DejaVuSerif-Italic.ttf"],
        ("Times New Roman", True, True): ["timesbi.ttf", "Times New Roman Bold Italic.ttf", "LiberationSerif-BoldItalic.ttf", "DejaVuSerif-BoldItalic.ttf"],
        ("Calibri", False, False): ["calibri.ttf", "Carlito-Regular.ttf", "Arial.ttf", "DejaVuSans.ttf"],
        ("Calibri", True, False): ["calibrib.ttf", "Carlito-Bold.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"],
        ("Calibri", False, True): ["calibrii.ttf", "Carlito-Italic.ttf", "Arial Italic.ttf", "DejaVuSans-Oblique.ttf"],
        ("Calibri", True, True): ["calibriz.ttf", "Carlito-BoldItalic.ttf", "Arial Bold Italic.ttf", "DejaVuSans-BoldOblique.ttf"],
        ("Courier New", False, False): ["cour.ttf", "Courier New.ttf", "LiberationMono-Regular.ttf", "DejaVuSansMono.ttf"],
        ("Courier New", True, False): ["courbd.ttf", "Courier New Bold.ttf", "LiberationMono-Bold.ttf", "DejaVuSansMono-Bold.ttf"],
        ("Courier New", False, True): ["couri.ttf", "Courier New Italic.ttf", "LiberationMono-Italic.ttf", "DejaVuSansMono-Oblique.ttf"],
        ("Courier New", True, True): ["courbi.ttf", "Courier New Bold Italic.ttf", "LiberationMono-BoldItalic.ttf", "DejaVuSansMono-BoldOblique.ttf"],
        ("DejaVu Sans", False, False): ["DejaVuSans.ttf"],
        ("DejaVu Sans", True, False): ["DejaVuSans-Bold.ttf"],
        ("DejaVu Sans", False, True): ["DejaVuSans-Oblique.ttf"],
        ("DejaVu Sans", True, True): ["DejaVuSans-BoldOblique.ttf"],
        ("Cascadia Code", False, False): ["CascadiaCode.ttf", "CascadiaMono.ttf", "DejaVuSansMono.ttf"],
        ("Cascadia Code", True, False): ["CascadiaCodePL-Bold.ttf", "CascadiaMonoPL-Bold.ttf", "DejaVuSansMono-Bold.ttf"],
        ("Cascadia Code", False, True): ["CascadiaCodeItalic.ttf", "DejaVuSansMono-Oblique.ttf"],
        ("Cascadia Code", True, True): ["CascadiaCodeBoldItalic.ttf", "DejaVuSansMono-BoldOblique.ttf"],
        ("Segoe UI", False, False): ["segoeui.ttf", "Arial.ttf", "DejaVuSans.ttf"],
        ("Segoe UI", True, False): ["segoeuib.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"],
        ("Segoe UI", False, True): ["segoeuii.ttf", "Arial Italic.ttf", "DejaVuSans-Oblique.ttf"],
        ("Segoe UI", True, True): ["segoeuiz.ttf", "Arial Bold Italic.ttf", "DejaVuSans-BoldOblique.ttf"],
    }
    base = [font_family, f"{font_family}.ttf", f"{font_family}.otf"]
    return mapping.get((font_family, bold, italic), []) + base


def get_font(font_family: str, size: int, bold: bool = False, italic: bool = False):
    for cand in get_font_path_candidates(font_family, bold=bold, italic=italic):
        try:
            return ImageFont.truetype(cand, size=size)
        except Exception:
            continue
    preferred = []
    if bold and italic:
        preferred.append("DejaVuSans-BoldOblique.ttf")
    elif bold:
        preferred.append("DejaVuSans-Bold.ttf")
    elif italic:
        preferred.append("DejaVuSans-Oblique.ttf")
    preferred.extend(["DejaVuSans.ttf", "arial.ttf"])
    for cand in preferred:
        try:
            return ImageFont.truetype(cand, size=size)
        except Exception:
            continue
    return ImageFont.load_default()


def clamp(n, lo, hi):
    return max(lo, min(hi, n))


def rects_overlap(a, b, pad=0):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    return not (ax2 + pad <= bx1 or bx2 + pad <= ax1 or ay2 + pad <= by1 or by2 + pad <= ay1)


def image_to_base64_png(img: Image.Image) -> str:
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return base64.b64encode(bio.getvalue()).decode("utf-8")


def base64_png_to_image(text: str) -> Image.Image:
    return Image.open(io.BytesIO(base64.b64decode(text))).convert("RGBA")


def make_plain_runs(text: str, bold: bool = False, italic: bool = False) -> List[Dict]:
    return [{"text": text, "bold": bold, "italic": italic}]


def normalize_runs(runs: List[Dict]) -> List[Dict]:
    if not runs:
        return [{"text": "", "bold": False, "italic": False}]
    out = []
    for run in runs:
        txt = run.get("text", "")
        b = bool(run.get("bold", False))
        i = bool(run.get("italic", False))
        if txt == "":
            continue
        if out and out[-1]["bold"] == b and out[-1]["italic"] == i:
            out[-1]["text"] += txt
        else:
            out.append({"text": txt, "bold": b, "italic": i})
    return out or [{"text": "", "bold": False, "italic": False}]


def runs_to_text(runs: List[Dict]) -> str:
    return "".join(r.get("text", "") for r in runs or [])


def rotate_point(px: float, py: float, cx: float, cy: float, degrees: float) -> Tuple[float, float]:
    if not degrees:
        return px, py
    ang = math.radians(degrees)
    cos_a = math.cos(ang)
    sin_a = math.sin(ang)
    dx = px - cx
    dy = py - cy
    return cx + dx * cos_a - dy * sin_a, cy + dx * sin_a + dy * cos_a


def point_line_distance(px: float, py: float, x1: float, y1: float, x2: float, y2: float) -> float:
    dx = x2 - x1
    dy = y2 - y1
    denom = dx * dx + dy * dy
    if denom == 0:
        return math.hypot(px - x1, py - y1)
    t = max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / denom))
    proj_x = x1 + t * dx
    proj_y = y1 + t * dy
    return math.hypot(px - proj_x, py - proj_y)


def color_to_rgb_tuple(color: str) -> Tuple[int, int, int]:
    if not color:
        return (0, 0, 0)
    c = color.strip()
    names = {
        "black": (0, 0, 0),
        "white": (255, 255, 255),
        "red": (255, 0, 0),
        "green": (0, 128, 0),
        "blue": (0, 0, 255),
        "yellow": (255, 255, 0),
        "orange": (255, 165, 0),
        "purple": (128, 0, 128),
        "gray": (128, 128, 128),
        "grey": (128, 128, 128),
    }
    if c.lower() in names:
        return names[c.lower()]
    if c.startswith("#") and len(c) == 7:
        try:
            return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
        except Exception:
            return (0, 0, 0)
    return (0, 0, 0)


def svg_color_opacity(color: str, alpha: int) -> Tuple[str, float]:
    return color if color else "none", max(0.0, min(1.0, alpha / 255.0))





def trim_bbox_from_image(img: Image.Image, tolerance: int = 12) -> Optional[Tuple[int, int, int, int]]:
    rgba = img.convert("RGBA")
    w, h = rgba.size

    alpha = rgba.getchannel("A")
    alpha_bbox = alpha.getbbox()

    if alpha_bbox:
        ax1, ay1, ax2, ay2 = alpha_bbox
        if ax1 > 0 or ay1 > 0 or ax2 < w or ay2 < h:
            return alpha_bbox

    corners = [
        rgba.getpixel((0, 0)),
        rgba.getpixel((w - 1, 0)),
        rgba.getpixel((0, h - 1)),
        rgba.getpixel((w - 1, h - 1)),
    ]

    def rgb_dist(a, b):
        return sum(abs(a[i] - b[i]) for i in range(3))

    bg = max(set(corners), key=corners.count)

    px = rgba.load()
    min_x, min_y = w, h
    max_x, max_y = -1, -1

    for y in range(h):
        for x in range(w):
            p = px[x, y]
            if p[3] == 0:
                continue
            if rgb_dist(p, bg) > tolerance or abs(p[3] - bg[3]) > tolerance:
                if x < min_x:
                    min_x = x
                if y < min_y:
                    min_y = y
                if x > max_x:
                    max_x = x
                if y > max_y:
                    max_y = y

    if max_x == -1 or max_y == -1:
        return None

    return (min_x, min_y, max_x + 1, max_y + 1)





class SplashScreen:
    def __init__(self, root, duration=6200):
        self.root = root
        self.root.overrideredirect(True)
        self.root.attributes("-alpha", 0.0)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")
        self.root.configure(bg="white")
        self.title_label = tk.Label(self.root, text="FigPie 🍰", font=("Cascadia Code", 42, "bold"), bg="white", fg="black")
        self.title_label.place(relx=0.5, rely=0.43, anchor="center")
        self.subtitle_label = tk.Label(self.root, text="Assemble scientific figures as easy as pie...", font=("Cascadia Code", 20, "italic"), bg="white", fg="#585824")
        self.subtitle_label.place(relx=0.5, rely=0.50, anchor="center")
        self.dev_label = tk.Label(self.root, text="Developed by Hamid Taghipour\nUiT The Arctic University of Norway", font=("Cascadia Code", 10), bg="white", fg="black")
        self.dev_label.place(relx=0.5, rely=0.58, anchor="center")
        self.fade_in_out(duration)

    def fade_in_out(self, duration):
        fade_in_ms = min(750, max(250, duration // 3))
        fade_out_ms = min(750, max(250, duration // 3))
        hold_ms = max(0, duration - fade_in_ms - fade_out_ms)
        self.fade_in(fade_in_ms, lambda: self.root.after(hold_ms, lambda: self.fade_out(fade_out_ms, self.close)))

    def fade_in(self, time_ms, callback):
        alpha = 0.0
        increment = 1 / max((time_ms // 30), 1)
        def fade():
            nonlocal alpha
            if alpha < 1.0:
                alpha = min(1.0, alpha + increment)
                self.root.attributes("-alpha", alpha)
                self.root.after(30, fade)
            else:
                callback()
        fade()

    def fade_out(self, time_ms, callback):
        alpha = 1.0
        decrement = 1 / max((time_ms // 30), 1)
        def fade():
            nonlocal alpha
            if alpha > 0.0:
                alpha = max(0.0, alpha - decrement)
                self.root.attributes("-alpha", alpha)
                self.root.after(30, fade)
            else:
                callback()
        fade()

    def close(self):
        self.root.destroy()


@dataclass
class PanelItem:
    id: int
    kind: str = "panel"
    source_path: Optional[str] = None
    pil_image: Optional[Image.Image] = None
    original_size: Tuple[int, int] = (100, 100)
    x: int = 50
    y: int = 50
    w: int = 200
    h: int = 150
    label: str = ""
    show_label: bool = True
    include_in_labels: bool = True
    label_font_size: int = DEFAULT_LABEL_FONT_SIZE
    label_font_family: str = "Arial"
    label_color: str = "black"
    label_offset_x: int = 10
    label_offset_y: int = 10
    border_width: int = 0
    border_color: str = "black"
    z_index: int = 0
    group_id: Optional[int] = None
    tk_preview: Optional[ImageTk.PhotoImage] = None
    _preview_cache_key: Optional[Tuple] = None
    _preview_cache_photo: Optional[ImageTk.PhotoImage] = None

    def bbox(self):
        return self.x, self.y, self.x + self.w, self.y + self.h

    def contains(self, px: int, py: int) -> bool:
        return self.x <= px <= self.x + self.w and self.y <= py <= self.y + self.h

    def resize_handle_hit(self, px: int, py: int, handle_model: int) -> bool:
        x1, y1, x2, y2 = self.bbox()
        return x2 - handle_model <= px <= x2 + handle_model and y2 - handle_model <= py <= y2 + handle_model


@dataclass
class TextItem:
    id: int
    kind: str = "text"
    text: str = "Text"
    x: int = 60
    y: int = 60
    w: int = 250
    h: int = 80
    font_size: int = DEFAULT_TEXT_FONT_SIZE
    font_family: str = "Arial"
    bold: bool = False
    italic: bool = False
    rich_runs: List[Dict] = field(default_factory=lambda: make_plain_runs("Text"))
    fill: str = "black"
    align: str = "left"
    background: Optional[str] = None
    outline: Optional[str] = None
    rotation: int = 0
    line_spacing: float = DEFAULT_LINE_SPACING
    z_index: int = 0
    group_id: Optional[int] = None
    tk_preview: Optional[ImageTk.PhotoImage] = None
    _preview_cache_key: Optional[Tuple] = None
    _preview_cache_photo: Optional[ImageTk.PhotoImage] = None

    def __post_init__(self):
        if not self.rich_runs:
            self.rich_runs = make_plain_runs(self.text, self.bold, self.italic)
        self.sync_text_from_runs()

    def sync_text_from_runs(self):
        self.rich_runs = normalize_runs(self.rich_runs)
        self.text = runs_to_text(self.rich_runs)

    def bbox(self):
        return self.x, self.y, self.x + self.w, self.y + self.h

    def contains(self, px: int, py: int) -> bool:
        return self.x <= px <= self.x + self.w and self.y <= py <= self.y + self.h

    def resize_handle_hit(self, px: int, py: int, handle_model: int) -> bool:
        x1, y1, x2, y2 = self.bbox()
        return x2 - handle_model <= px <= x2 + handle_model and y2 - handle_model <= py <= y2 + handle_model


@dataclass
class ShapeItem:
    id: int
    kind: str = "shape"
    shape_type: str = "rectangle"
    x: int = 100
    y: int = 100
    w: int = 150
    h: int = 100
    color: str = "black"
    width: int = 2
    fill: Optional[str] = None
    fill_alpha: int = 0
    rotation: float = 0.0
    text: str = ""
    font_size: int = 22
    font_family: str = "Arial"
    z_index: int = 0
    group_id: Optional[int] = None
    compare_style: str = "bracket"
    bracket_height: int = DEFAULT_STAT_BRACKET_HEIGHT
    text_gap: int = DEFAULT_STAT_TEXT_GAP
    auto_center_text: bool = True
    snap_to_tops: bool = True

    def is_stat_compare(self) -> bool:
        return self.shape_type in ("stat_compare", "stat_bracket")

    def stat_top_y(self) -> float:
        p1 = (self.x, self.y)
        p2 = (self.x + self.w, self.y + self.h)
        if self.snap_to_tops:
            return min(p1[1], p2[1]) - max(0, self.bracket_height)
        return p1[1] - max(0, self.bracket_height)

    def stat_text_position(self) -> Tuple[float, float]:
        pts = self.base_points()
        top_y = min(p[1] for p in pts)
        if self.auto_center_text:
            mx = (self.x + self.x + self.w) / 2
        else:
            mx = self.x
        return mx, top_y - max(0, self.text_gap)

    def base_points(self) -> List[Tuple[float, float]]:
        if self.shape_type in ("line", "arrow"):
            return [(self.x, self.y), (self.x + self.w, self.y + self.h)]
        if self.shape_type == "arrow_head":
            return [(self.x + self.w, self.y + self.h / 2), (self.x, self.y), (self.x, self.y + self.h)]
        if self.is_stat_compare():
            x1, y1 = self.x, self.y
            x2, y2 = self.x + self.w, self.y + self.h
            top_y = self.stat_top_y()
            if self.compare_style == "line":
                return [(x1, top_y), (x2, top_y)]
            return [(x1, y1), (x1, top_y), (x2, top_y), (x2, y2)]
        return [(self.x, self.y), (self.x + self.w, self.y), (self.x + self.w, self.y + self.h), (self.x, self.y + self.h)]

    def visual_points(self) -> List[Tuple[float, float]]:
        pts = self.base_points()
        if self.is_stat_compare():
            return pts
        if not self.rotation:
            return pts
        cx = self.x + self.w / 2
        cy = self.y + self.h / 2
        return [rotate_point(px, py, cx, cy, self.rotation) for px, py in pts]

    def bbox(self):
        pts = self.visual_points()
        if not pts:
            return self.x, self.y, self.x + self.w, self.y + self.h
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        pad = max(8, self.width + 8)
        if self.is_stat_compare() and self.text:
            pad = max(pad, self.font_size + self.text_gap + 16)
        return int(math.floor(min(xs) - pad)), int(math.floor(min(ys) - pad)), int(math.ceil(max(xs) + pad)), int(math.ceil(max(ys) + pad))

    def contains(self, px: int, py: int) -> bool:
        tol = max(12, self.width + 10)
        if self.shape_type in ("line", "arrow"):
            p1, p2 = self.visual_points()[:2]
            return point_line_distance(px, py, p1[0], p1[1], p2[0], p2[1]) <= tol
        if self.is_stat_compare():
            pts = self.visual_points()
            for a, b in zip(pts, pts[1:]):
                if point_line_distance(px, py, a[0], a[1], b[0], b[1]) <= tol:
                    return True
            if self.text:
                tx, ty = self.stat_text_position()
                return abs(px - tx) <= max(20, len(self.text) * self.font_size * 0.4) and abs(py - ty) <= self.font_size
            x1, y1, x2, y2 = self.bbox()
            return x1 <= px <= x2 and y1 <= py <= y2
        x1, y1, x2, y2 = self.bbox()
        return x1 <= px <= x2 and y1 <= py <= y2

    def resize_handle_hit(self, px: int, py: int, handle_model: int) -> bool:
        x1, y1, x2, y2 = self.bbox()
        return x2 - handle_model <= px <= x2 + handle_model and y2 - handle_model <= py <= y2 + handle_model


CanvasItem = Union[PanelItem, TextItem, ShapeItem]


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, width=260):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0, width=width)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", self._on_frame_configure)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")
        self.inner.bind("<Enter>", lambda e: self.canvas.focus_set())

    def _on_frame_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)
        self.canvas.coords(self.window_id, 0, 0)

    def scroll_to_top(self):
        self.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.yview_moveto(0)

    def _mouse_inside(self, event):
        if not self.winfo_ismapped():
            return False
        try:
            widget = self.winfo_containing(event.x_root, event.y_root)
        except Exception:
            widget = None
        return bool(widget and str(widget).startswith(str(self)))

    def _can_scroll(self):
        bbox = self.canvas.bbox("all")
        if not bbox:
            return False
        return (bbox[3] - bbox[1]) > self.canvas.winfo_height() + 1

    def _on_mousewheel(self, event):
        if self._mouse_inside(event) and self._can_scroll():
            delta = -1 * int(event.delta / 120) if event.delta else 0
            if delta:
                self.canvas.yview_scroll(delta, "units")

    def _on_mousewheel_linux(self, event):
        if self._mouse_inside(event) and self._can_scroll():
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")


class RichTextEditor(tk.Toplevel):
    def __init__(self, parent, item: TextItem, on_save):
        super().__init__(parent)
        self.title("Rich text editor")
        self.geometry("980x670")
        self.transient(parent)
        self.grab_set()
        self.item = item
        self.on_save = on_save
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="Bold", command=self.make_bold).pack(side="left", padx=4)
        ttk.Button(top, text="Italic", command=self.make_italic).pack(side="left", padx=4)
        ttk.Button(top, text="Bold + Italic", command=self.make_bold_italic).pack(side="left", padx=4)
        ttk.Button(top, text="Normal", command=self.make_normal).pack(side="left", padx=4)
        ttk.Label(top, text="Font").pack(side="left", padx=(18, 4))
        self.font_var = tk.StringVar(value=item.font_family)
        ttk.Combobox(top, values=list_font_families(), textvariable=self.font_var, state="readonly", width=22).pack(side="left", padx=4)
        ttk.Label(top, text="Size").pack(side="left", padx=(10, 4))
        self.size_var = tk.IntVar(value=item.font_size)
        ttk.Spinbox(top, from_=6, to=200, textvariable=self.size_var, width=6).pack(side="left", padx=4)
        ttk.Label(top, text="Align").pack(side="left", padx=(10, 4))
        self.align_var = tk.StringVar(value=item.align)
        ttk.Combobox(top, values=["left", "center", "right", "justify"], textvariable=self.align_var, state="readonly", width=10).pack(side="left", padx=4)
        ttk.Label(top, text="Rotation").pack(side="left", padx=(10, 4))
        self.rotation_var = tk.IntVar(value=item.rotation)
        ttk.Combobox(top, values=[0, 90, 180, 270], textvariable=self.rotation_var, state="readonly", width=6).pack(side="left", padx=4)
        ttk.Label(top, text="Line spacing").pack(side="left", padx=(10, 4))
        self.line_spacing_var = tk.DoubleVar(value=item.line_spacing)
        ttk.Spinbox(top, from_=0.8, to=3.0, increment=0.05, textvariable=self.line_spacing_var, width=6).pack(side="left", padx=4)
        text_wrap = ttk.Frame(self)
        text_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.text = tk.Text(text_wrap, wrap="word", undo=True)
        ybar = ttk.Scrollbar(text_wrap, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=ybar.set)
        self.text.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bottom, text="Save", command=self.save_and_close).pack(side="right", padx=4)
        ttk.Button(bottom, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        self._reconfigure_tags()
        self._load_item_content()
        self.bind("<Control-b>", lambda e: self.make_bold())
        self.bind("<Control-i>", lambda e: self.make_italic())

    def _reconfigure_tags(self):
        family = self.font_var.get()
        size = int(self.size_var.get())
        self.text.tag_configure("normal", font=(family, size, "normal"))
        self.text.tag_configure("bold", font=(family, size, "bold"))
        self.text.tag_configure("italic", font=(family, size, "italic"))
        self.text.tag_configure("bolditalic", font=(family, size, "bold italic"))

    def _load_item_content(self):
        self.text.delete("1.0", "end")
        for run in self.item.rich_runs:
            tag = self._style_to_tag(run.get("bold", False), run.get("italic", False))
            self.text.insert("end", run.get("text", ""), (tag,))
        self.text.focus_set()

    def _style_to_tag(self, bold, italic):
        if bold and italic:
            return "bolditalic"
        if bold:
            return "bold"
        if italic:
            return "italic"
        return "normal"

    def _apply_tag_to_selection(self, tag):
        try:
            start = self.text.index("sel.first")
            end = self.text.index("sel.last")
        except Exception:
            return
        for t in ["normal", "bold", "italic", "bolditalic"]:
            self.text.tag_remove(t, start, end)
        self.text.tag_add(tag, start, end)

    def make_bold(self):
        self._apply_tag_to_selection("bold")

    def make_italic(self):
        self._apply_tag_to_selection("italic")

    def make_bold_italic(self):
        self._apply_tag_to_selection("bolditalic")

    def make_normal(self):
        self._apply_tag_to_selection("normal")

    def extract_runs(self) -> List[Dict]:
        content = self.text.get("1.0", "end-1c")
        if not content:
            return [{"text": "", "bold": False, "italic": False}]
        runs = []
        idx = "1.0"
        while self.text.compare(idx, "<", "end-1c"):
            next_idx = self.text.index(f"{idx}+1c")
            ch = self.text.get(idx, next_idx)
            tags = set(self.text.tag_names(idx))
            bold = "bold" in tags or "bolditalic" in tags
            italic = "italic" in tags or "bolditalic" in tags
            if runs and runs[-1]["bold"] == bold and runs[-1]["italic"] == italic:
                runs[-1]["text"] += ch
            else:
                runs.append({"text": ch, "bold": bold, "italic": italic})
            idx = next_idx
        return normalize_runs(runs)

    def save_and_close(self):
        self._reconfigure_tags()
        payload = {
            "rich_runs": self.extract_runs(),
            "font_family": self.font_var.get(),
            "font_size": int(self.size_var.get()),
            "align": self.align_var.get(),
            "rotation": int(self.rotation_var.get()),
            "line_spacing": float(self.line_spacing_var.get()),
        }
        self.on_save(payload)
        self.destroy()


class FigureBoardApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1850x1100")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.canvas_w = DEFAULT_CANVAS_W
        self.canvas_h = DEFAULT_CANVAS_H
        self.bg_color = DEFAULT_BG
        self.view_zoom = 1.0
        self.items: List[CanvasItem] = []
        self.next_id = 1
        self.next_group_id = 1
        self.selected_ids: List[int] = []
        self.anchor_selected_id: Optional[int] = None
        self.selected_label_panel_id: Optional[int] = None
        self.drag_mode: Optional[str] = None
        self.drag_start = (0, 0)
        self.drag_origin_map: Dict[int, Tuple[int, int, int, int]] = {}
        self.before_drag_state = None
        self.hover_item_id: Optional[int] = None
        self.context_click_item_id: Optional[int] = None
        self.erase_mode = False
        self.erase_start = None
        self.erase_current = None
        self.crop_mode = False
        self.crop_start = None
        self.crop_current = None
        self.crop_target_id = None
        self.select_start: Optional[Tuple[int, int]] = None
        self.select_end: Optional[Tuple[int, int]] = None
        self.rect_add_mode = False
        self.stat_mode: Optional[str] = None
        self.stat_first_point: Optional[Tuple[int, int]] = None
        self.label_align_mode = False
        self.label_anchor_panel_id: Optional[int] = None
        self.label_sequence_mode = False
        self.label_sequence_ids: List[int] = []
        self.undo_stack = []
        self.redo_stack = []
        self.dirty = False
        self.project_path: Optional[str] = None
        self.outer_margin = tk.IntVar(value=DEFAULT_OUTER_MARGIN)
        self.gap_x = tk.IntVar(value=DEFAULT_GAP_X)
        self.gap_y = tk.IntVar(value=DEFAULT_GAP_Y)
        self.default_label_size = tk.IntVar(value=DEFAULT_LABEL_FONT_SIZE)
        self.default_text_size = tk.IntVar(value=DEFAULT_TEXT_FONT_SIZE)
        self.default_font_family = tk.StringVar(value="Arial")
        self.default_label_offset_x = tk.IntVar(value=10)
        self.default_label_offset_y = tk.IntVar(value=10)
        self.auto_trim_on_export = tk.BooleanVar(value=True)
        self.keep_label_order = tk.BooleanVar(value=True)
        self.show_grid = tk.BooleanVar(value=False)
        self.canvas_unit = tk.StringVar(value="px")
        self.export_dpi = tk.IntVar(value=DEFAULT_DPI)
        self.status_var = tk.StringVar(value="Ready")
        self.zoom_var = tk.StringVar(value="100%")
        self.stat_symbol_var = tk.StringVar(value="*")
        self.stat_count_var = tk.IntVar(value=3)
        self.stat_text_size_var = tk.IntVar(value=DEFAULT_STAT_TEXT_SIZE)
        self.stat_bracket_height_var = tk.IntVar(value=DEFAULT_STAT_BRACKET_HEIGHT)
        self.stat_text_gap_var = tk.IntVar(value=DEFAULT_STAT_TEXT_GAP)
        self.stat_color_var = tk.StringVar(value="black")
        self.stat_compare_style_var = tk.StringVar(value="Bracket line")
        self.stat_auto_center_var = tk.BooleanVar(value=True)
        self.stat_snap_tops_var = tk.BooleanVar(value=True)
        self.label_align_axis_var = tk.StringVar(value="Y")
        self.font_families = list_font_families()
        self._build_ui()
        self._bind_events()
        self._setup_traces()
        self.redraw()

    def small_button_row(self, parent, specs, padx=6, pady=3):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", padx=padx, pady=pady)
        for text, command, width in specs:
            ttk.Button(frame, text=text, command=command, width=width).pack(side="left", padx=2, pady=1)
        return frame

    def apply_default_label_font_family(self):
        family = self.default_font_family.get()
        self.save_undo_state()
        for item in self.items:
            if isinstance(item, PanelItem):
                item.label_font_family = family
        self.redraw()
        self.set_status("Default label font applied")

    def sx(self, x: float) -> int:
        return int(round(x * self.view_zoom))

    def sy(self, y: float) -> int:
        return int(round(y * self.view_zoom))

    def mx(self, x: float) -> float:
        return x / self.view_zoom

    def my(self, y: float) -> float:
        return y / self.view_zoom

    def handle_model_size(self) -> int:
        return max(4, int(round(HANDLE_SIZE / self.view_zoom)))

    def model_from_event(self, event) -> Tuple[int, int]:
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        return int(round(self.mx(cx))), int(round(self.my(cy)))

    def scaled_scrollregion(self):
        return (0, 0, self.sx(self.canvas_w), self.sy(self.canvas_h))

    def _build_ui(self):
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)
        left_wrap = ttk.Frame(outer, width=315)
        left_wrap.pack(side="left", fill="y", padx=6, pady=6)
        center = ttk.Frame(outer)
        center.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        right_wrap = ttk.Frame(outer, width=390)
        right_wrap.pack(side="right", fill="y", padx=6, pady=6)
        self.left_scroll = ScrollableFrame(left_wrap, width=315)
        self.left_scroll.pack(fill="both", expand=True)
        left = self.left_scroll.inner
        self.right_scroll = ScrollableFrame(right_wrap, width=390)
        self.right_scroll.pack(fill="both", expand=True)
        right = self.right_scroll.inner

        import_box = tk.LabelFrame(left, text="Import", bg="white", fg="green", bd=1, font=("Cascadia Code", 12, "bold"))
        import_box.pack(fill="x", pady=(0, 8))
        self.small_button_row(import_box, [("Files 📁", self.add_files, 10), ("Clipboard 📋", self.paste_from_clipboard, 13)])

        canvas_box = tk.LabelFrame(left, text="Canvas 🖌️", bg="white", fg="darkorange", bd=1, font=("Cascadia Code", 12, "bold"))
        canvas_box.pack(fill="x", pady=(0, 8))
        row = ttk.Frame(canvas_box)
        row.pack(fill="x", padx=6, pady=4)
        ttk.Label(row, text="Unit").pack(side="left")
        ttk.Combobox(row, values=["px", "in", "cm"], textvariable=self.canvas_unit, width=5, state="readonly").pack(side="left", padx=(4, 8))
        ttk.Label(row, text="W").pack(side="left")
        self.canvas_w_entry = ttk.Entry(row, width=7)
        self.canvas_w_entry.insert(0, str(self.canvas_w))
        self.canvas_w_entry.pack(side="left", padx=(4, 8))
        ttk.Label(row, text="H").pack(side="left")
        self.canvas_h_entry = ttk.Entry(row, width=7)
        self.canvas_h_entry.insert(0, str(self.canvas_h))
        self.canvas_h_entry.pack(side="left", padx=4)
        ttk.Button(canvas_box, text="Apply canvas size", command=self.apply_canvas_size).pack(fill="x", padx=6, pady=4)
        settings_row1 = ttk.Frame(canvas_box)
        settings_row1.pack(fill="x", padx=6, pady=(6, 3))
        ttk.Label(settings_row1, text="Margin").pack(side="left")
        ttk.Spinbox(settings_row1, from_=-500, to=1000, textvariable=self.outer_margin, width=6, command=self.redraw).pack(side="left", padx=(2, 12))
        ttk.Label(settings_row1, text="Gap X").pack(side="left")
        ttk.Spinbox(settings_row1, from_=-100, to=1000, textvariable=self.gap_x, width=6, command=self.on_gap_change).pack(side="left", padx=(2, 0))
        settings_row2 = ttk.Frame(canvas_box)
        settings_row2.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Label(settings_row2, text="Gap Y").pack(side="left")
        ttk.Spinbox(settings_row2, from_=-100, to=1000, textvariable=self.gap_y, width=6, command=self.on_gap_change).pack(side="left", padx=(2, 12))
        ttk.Checkbutton(settings_row2, text="Show grid", variable=self.show_grid, command=self.redraw).pack(side="left")

        arrange_box = tk.LabelFrame(left, text="Arrange", bg="white", fg="blue", bd=1, font=("Cascadia Code", 12, "bold"))
        arrange_box.pack(fill="x", pady=(0, 8))
        grid_row = ttk.Frame(arrange_box)
        grid_row.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Label(grid_row, text="Rows").pack(side="left")
        self.grid_rows_entry = ttk.Entry(grid_row, width=5)
        self.grid_rows_entry.insert(0, "2")
        self.grid_rows_entry.pack(side="left", padx=(4, 10))
        ttk.Label(grid_row, text="Cols").pack(side="left")
        self.grid_cols_entry = ttk.Entry(grid_row, width=5)
        self.grid_cols_entry.insert(0, "2")
        self.grid_cols_entry.pack(side="left", padx=4)
        self.small_button_row(arrange_box, [("Custom grid", self.apply_custom_grid, 13), ("Auto grid", self.auto_grid, 10)])
        self.small_button_row(arrange_box, [("Dist H", self.distribute_h, 8), ("Dist V", self.distribute_v, 8), ("Overlaps", self.resolve_all_overlaps, 9)])
        ttk.Label(arrange_box, text="Align to anchor", anchor="center").pack(fill="x", padx=6, pady=(6, 2))
        align_frame = ttk.Frame(arrange_box)
        align_frame.pack(anchor="center", padx=6, pady=(0, 4))
        for sym, where in [("⬅", "left"), ("⬆", "top"), ("⬇", "bottom"), ("➡", "right"), ("↔", "hcenter"), ("↕", "vcenter")]:
            tk.Button(align_frame, text=sym, width=3, font=("Arial", 15, "bold"), command=lambda w=where: self.align_to_anchor(w)).pack(side="left", padx=2)
        self.small_button_row(arrange_box, [("Same W", self.same_widths, 8), ("Same H", self.same_heights, 8)])

        group_box = tk.LabelFrame(left, text="Selection and groups", bg="white", fg="violet", bd=1, font=("Cascadia Code", 12, "bold"))
        group_box.pack(fill="x", pady=(0, 8))
        self.small_button_row(group_box, [("Select all", self.select_all_items, 10), ("Group", self.group_selected, 8), ("Ungroup", self.ungroup_selected, 9)])
        self.small_button_row(group_box, [("Duplicate", self.duplicate_selected, 10), ("Delete", self.delete_selected, 8)])

       



        edit_box = tk.LabelFrame(left, text="Edit", bg="white", fg="purple", bd=1, font=("Cascadia Code", 12, "bold"))
        edit_box.pack(fill="x", pady=(0, 8))

        self.small_button_row(
            edit_box,
            [("Undo", self.undo, 7), ("Redo", self.redo, 7), ("Colour", self.pick_selected_text_colour, 8), ("Text", self.add_text_box, 7)],
            pady=3
        )

        ttk.Button(edit_box, text="Add caption below panels", command=self.add_caption_below_last_figure).pack(
            fill="x", padx=6, pady=(1, 5)
        )

        self.small_button_row(edit_box, [("Trim image", self.trim_selected_image, 11), ("Crop canvas", self.crop_canvas_to_content, 12)], pady=3)

        mode_row = ttk.Frame(edit_box)
        mode_row.pack(fill="x", padx=6, pady=(3, 4))
        self.crop_btn = tk.Button(mode_row, text="Crop OFF", command=self.toggle_crop_mode)
        self.erase_btn = tk.Button(mode_row, text="Erase OFF", command=self.toggle_erase_mode)
        self.crop_btn.pack(side="left", fill="x", expand=True, padx=2)
        self.erase_btn.pack(side="left", fill="x", expand=True, padx=2)



















        export_box = tk.LabelFrame(left, text="Project and export 🎛️", bg="white", fg="red", bd=1, font=("Cascadia Code", 12, "bold"))
        export_box.pack(fill="x", pady=(0, 8))
        self.small_button_row(export_box, [("New", self.new_canvas, 7), ("Save", self.save_project, 7), ("Open", self.open_project, 7)])
        ttk.Checkbutton(export_box, text="Trim whitespace on export ✂️", variable=self.auto_trim_on_export).pack(anchor="w", padx=6, pady=4)
        dpi_row = ttk.Frame(export_box)
        dpi_row.pack(fill="x", padx=6, pady=4)
        ttk.Label(dpi_row, text="DPI").pack(side="left")
        ttk.Spinbox(dpi_row, from_=72, to=2400, textvariable=self.export_dpi, width=8).pack(side="left", padx=6)
        self.small_button_row(export_box, [("PNG", lambda: self.export_image("PNG"), 7), ("TIFF", lambda: self.export_image("TIFF"), 7), ("PDF", self.export_pdf, 7), ("SVG", self.export_svg, 7)])

        canvas_frame = ttk.Frame(center)
        canvas_frame.pack(fill="both", expand=True)
        self.hbar = ttk.Scrollbar(canvas_frame, orient="horizontal")
        self.vbar = ttk.Scrollbar(canvas_frame, orient="vertical")
        self.canvas = tk.Canvas(canvas_frame, bg=self.bg_color, xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set, highlightthickness=1, highlightbackground="#888")
        self.hbar.config(command=self.canvas.xview)
        self.vbar.config(command=self.canvas.yview)
        self.hbar.pack(side="bottom", fill="x")
        self.vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.config(scrollregion=self.scaled_scrollregion())
        if DND_AVAILABLE:
            try:
                self.canvas.drop_target_register(DND_FILES)
                self.canvas.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

        prop_box = tk.LabelFrame(right, text="Selected item info", bg="white", fg="darkgreen", bd=1, font=("Cascadia Code", 12, "bold"))
        prop_box.pack(fill="x", pady=(0, 8))
        self.sel_type_var = tk.StringVar(value="-")
        self.sel_label_var = tk.StringVar(value="")
        self.sel_text_var = tk.StringVar(value="")
        self.sel_x_var = tk.StringVar(value="0")
        self.sel_y_var = tk.StringVar(value="0")
        self.sel_w_var = tk.StringVar(value="0")
        self.sel_h_var = tk.StringVar(value="0")
        self.sel_font_var = tk.StringVar(value=str(DEFAULT_TEXT_FONT_SIZE))
        self.sel_font_family_var = tk.StringVar(value="Arial")
        self.sel_show_label_var = tk.BooleanVar(value=True)
        self.sel_include_label_var = tk.BooleanVar(value=True)
        self.sel_align_var = tk.StringVar(value="left")
        self.sel_bold_var = tk.BooleanVar(value=False)
        self.sel_italic_var = tk.BooleanVar(value=False)
        self.sel_rotation_var = tk.StringVar(value="0")
        self.sel_line_spacing_var = tk.StringVar(value=str(DEFAULT_LINE_SPACING))
        ttk.Label(prop_box, text="Type").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        ttk.Label(prop_box, textvariable=self.sel_type_var).grid(row=0, column=1, sticky="w", padx=6, pady=3)
        ttk.Label(prop_box, text="Label").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(prop_box, textvariable=self.sel_label_var).grid(row=1, column=1, sticky="ew", padx=6, pady=3)
        ttk.Label(prop_box, text="Text").grid(row=2, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(prop_box, textvariable=self.sel_text_var).grid(row=2, column=1, sticky="ew", padx=6, pady=3)
        geom_row = ttk.Frame(prop_box)
        geom_row.grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=3)
        for label, var in [("X", self.sel_x_var), ("Y", self.sel_y_var), ("W", self.sel_w_var), ("H", self.sel_h_var)]:
            ttk.Label(geom_row, text=label).pack(side="left")
            ttk.Entry(geom_row, textvariable=var, width=6).pack(side="left", padx=(2, 7))
        ttk.Label(prop_box, text="Font/Width").grid(row=4, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(prop_box, textvariable=self.sel_font_var, width=10).grid(row=4, column=1, sticky="w", padx=6, pady=3)
        ttk.Label(prop_box, text="Font family").grid(row=5, column=0, sticky="w", padx=6, pady=3)
        ttk.Combobox(prop_box, values=self.font_families, textvariable=self.sel_font_family_var, state="readonly").grid(row=5, column=1, sticky="ew", padx=6, pady=3)
        ttk.Label(prop_box, text="Align").grid(row=6, column=0, sticky="w", padx=6, pady=3)
        ttk.Combobox(prop_box, values=["left", "center", "right", "justify"], textvariable=self.sel_align_var, state="readonly").grid(row=6, column=1, sticky="ew", padx=6, pady=3)
        ttk.Label(prop_box, text="Rotation°").grid(row=7, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(prop_box, textvariable=self.sel_rotation_var, width=10).grid(row=7, column=1, sticky="w", padx=6, pady=3)
        ttk.Label(prop_box, text="Line spacing").grid(row=8, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(prop_box, textvariable=self.sel_line_spacing_var, width=10).grid(row=8, column=1, sticky="w", padx=6, pady=3)
        checks = ttk.Frame(prop_box)
        checks.grid(row=9, column=0, columnspan=2, sticky="ew", padx=6, pady=3)
        ttk.Checkbutton(checks, text="Bold", variable=self.sel_bold_var).pack(side="left")
        ttk.Checkbutton(checks, text="Italic", variable=self.sel_italic_var).pack(side="left")
        ttk.Checkbutton(prop_box, text="Show selected panel label", variable=self.sel_show_label_var, command=self.apply_selected_label_visibility_quick).grid(row=10, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(prop_box, text="Include in auto labels", variable=self.sel_include_label_var, command=self.apply_selected_label_visibility_quick).grid(row=11, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Button(prop_box, text="Apply changes", command=self.apply_selected_properties).grid(row=12, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        prop_box.columnconfigure(1, weight=1)

        label_box = tk.LabelFrame(right, text="Labels", bg="white", fg="saddle brown", bd=1, font=("Cascadia Code", 12, "bold"))
        label_box.pack(fill="x", pady=(0, 8))
        self.small_button_row(label_box, [("Generate", self.regenerate_labels, 13), ("Toggle", self.toggle_labels, 8)])
        ttk.Label(label_box, text="Default label size").pack(anchor="w", padx=6, pady=(5, 1))
        ttk.Spinbox(label_box, from_=8, to=120, textvariable=self.default_label_size, width=8, command=self.apply_default_label_size).pack(anchor="w", padx=6, pady=(0, 4))
        ttk.Label(label_box, text="Default font family").pack(anchor="w", padx=6, pady=(2, 1))
        label_font_combo = ttk.Combobox(label_box, values=self.font_families, textvariable=self.default_font_family, state="readonly")
        label_font_combo.pack(fill="x", padx=6, pady=(0, 4))
        label_font_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_default_label_font_family())
        ttk.Checkbutton(label_box, text="Keep labels by add order", variable=self.keep_label_order).pack(anchor="w", padx=6, pady=3)
        auto_label_frame = ttk.Frame(label_box)
        auto_label_frame.pack(fill="x", padx=6, pady=4)
        ttk.Label(auto_label_frame, text="Auto position labels").pack(anchor="w")
        self.label_pos_var = tk.StringVar(value="top-in")
        pos_combo = ttk.Combobox(auto_label_frame, values=["top-in", "top-out", "left-in", "left-out", "bottom-in", "bottom-out", "center"], textvariable=self.label_pos_var, state="readonly", width=13)
        pos_combo.pack(side="left", padx=(0, 6))
        ttk.Button(auto_label_frame, text="All", width=6, command=self.apply_auto_label_position_all).pack(side="left")
        ttk.Button(auto_label_frame, text="Selected", width=9, command=self.apply_auto_label_position_selected).pack(side="left", padx=(4, 0))
        offset_row = ttk.Frame(label_box)
        offset_row.pack(fill="x", padx=6, pady=3)
        ttk.Label(offset_row, text="Offset X").pack(side="left")
        ttk.Spinbox(offset_row, from_=-200, to=500, textvariable=self.default_label_offset_x, width=6).pack(side="left", padx=(2, 8))
        ttk.Label(offset_row, text="Y").pack(side="left")
        ttk.Spinbox(offset_row, from_=-200, to=500, textvariable=self.default_label_offset_y, width=6).pack(side="left", padx=(2, 0))
        self.small_button_row(label_box, [("Offsets all", self.apply_default_label_offsets, 12), ("Selected", self.apply_default_label_offsets_selected, 9)])
        self.small_button_row(label_box, [("Exclude", self.exclude_selected_from_labels, 9), ("Include", self.include_selected_in_labels, 9)])
        ttk.Label(label_box, text="Align selected labels to anchor").pack(anchor="w", padx=6, pady=(6, 1))
        self.small_button_row(label_box, [("Same X", lambda: self.align_selected_labels("X"), 9), ("Same Y", lambda: self.align_selected_labels("Y"), 9), ("Same XY", lambda: self.align_selected_labels("Both"), 9)])
        axis_row = ttk.Frame(label_box)
        axis_row.pack(fill="x", padx=6, pady=(3, 0))
        ttk.Label(axis_row, text="Click-align axis").pack(side="left")
        ttk.Combobox(axis_row, values=["X", "Y", "Both"], textvariable=self.label_align_axis_var, state="readonly", width=8).pack(side="left", padx=6)
        self.label_seq_btn = tk.Button(label_box, text="Click label order OFF", command=self.toggle_label_sequence_mode)
        self.label_align_btn = tk.Button(label_box, text="Align labels by click OFF", command=self.toggle_label_align_mode)
        label_mode_row = ttk.Frame(label_box)
        label_mode_row.pack(fill="x", padx=6, pady=4)
        self.label_seq_btn.pack(in_=label_mode_row, side="left", fill="x", expand=True, padx=2)
        self.label_align_btn.pack(in_=label_mode_row, side="left", fill="x", expand=True, padx=2)

        shapes_box = tk.LabelFrame(right, text="Shapes", bg="white", fg="black", bd=1, font=("Cascadia Code", 12, "bold"))
        shapes_box.pack(fill="x", pady=(0, 8))
        self.small_button_row(shapes_box, [("Rect", lambda: self.add_shape("rectangle"), 7), ("Circle", lambda: self.add_shape("circle"), 8), ("Line", lambda: self.add_shape("line"), 7), ("Arrow", lambda: self.add_shape("arrow"), 8)])
        self.small_button_row(shapes_box, [("Arrow head", lambda: self.add_shape("arrow_head"), 12), ("Highlight", lambda: self.add_shape("highlight_bar"), 11)])

        stats_box = tk.LabelFrame(right, text="Statistics ✱", bg="white", fg="darkred", bd=1, font=("Cascadia Code", 12, "bold"))
        stats_box.pack(fill="x", pady=(0, 8))
        stat_row = ttk.Frame(stats_box)
        stat_row.pack(fill="x", padx=6, pady=4)
        ttk.Label(stat_row, text="Symbol").pack(side="left")
        ttk.Entry(stat_row, textvariable=self.stat_symbol_var, width=7).pack(side="left", padx=(4, 10))
        ttk.Label(stat_row, text="Count").pack(side="left")
        ttk.Spinbox(stat_row, from_=1, to=12, textvariable=self.stat_count_var, width=5).pack(side="left", padx=(4, 0))
        stat_row2 = ttk.Frame(stats_box)
        stat_row2.pack(fill="x", padx=6, pady=3)
        ttk.Label(stat_row2, text="Text size").pack(side="left")
        ttk.Spinbox(stat_row2, from_=6, to=200, textvariable=self.stat_text_size_var, width=6).pack(side="left", padx=(4, 10))
        ttk.Label(stat_row2, text="Bracket height").pack(side="left")
        ttk.Spinbox(stat_row2, from_=0, to=300, textvariable=self.stat_bracket_height_var, width=6).pack(side="left", padx=(4, 0))
        stat_row3 = ttk.Frame(stats_box)
        stat_row3.pack(fill="x", padx=6, pady=3)
        ttk.Label(stat_row3, text="Text gap").pack(side="left")
        ttk.Spinbox(stat_row3, from_=0, to=200, textvariable=self.stat_text_gap_var, width=6).pack(side="left", padx=(4, 10))
        ttk.Label(stat_row3, text="Style").pack(side="left")
        ttk.Combobox(stat_row3, values=["Bracket line", "Straight line"], textvariable=self.stat_compare_style_var, state="readonly", width=13).pack(side="left", padx=(4, 0))
        stat_row4 = ttk.Frame(stats_box)
        stat_row4.pack(fill="x", padx=6, pady=3)
        ttk.Label(stat_row4, text="Color").pack(side="left")
        self.stat_color_swatch = tk.Button(stat_row4, text="     ", width=4, bg=self.stat_color_var.get(), command=self.pick_stat_color)
        self.stat_color_swatch.pack(side="left", padx=(4, 10))
        ttk.Checkbutton(stat_row4, text="Auto center text", variable=self.stat_auto_center_var).pack(side="left")
        ttk.Checkbutton(stats_box, text="Snap to data tops", variable=self.stat_snap_tops_var).pack(anchor="w", padx=6, pady=2)
        stat_mode_row = ttk.Frame(stats_box)
        stat_mode_row.pack(fill="x", padx=6, pady=4)
        self.stat_symbol_btn = tk.Button(stat_mode_row, text="Place symbol OFF", command=lambda: self.set_stat_mode("symbol"))
        self.stat_bracket_btn = tk.Button(stat_mode_row, text="Bracket compare OFF", command=lambda: self.set_stat_mode("compare"))
        self.stat_symbol_btn.pack(side="left", fill="x", expand=True, padx=2)
        self.stat_bracket_btn.pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(stats_box, text="Clear last bracket", command=self.clear_last_stat_compare).pack(fill="x", padx=6, pady=(0, 6))

        zoom_box = tk.LabelFrame(right, text="Zoom 🔍", bg="white", fg="dodger blue", bd=1, font=("Cascadia Code", 12, "bold"))
        zoom_box.pack(fill="x", pady=(0, 8))
        self.small_button_row(zoom_box, [("Out", lambda: self.zoom_by(1 / ZOOM_STEP), 8), ("In", lambda: self.zoom_by(ZOOM_STEP), 8), ("Reset", self.reset_zoom, 8)])
        ttk.Label(zoom_box, textvariable=self.zoom_var).pack(anchor="w", padx=6, pady=(0, 6))

        self.bottom_frame = ttk.Frame(self.root)
        self.bottom_frame.pack(fill="x", side="bottom", padx=4, pady=(0, 4))
        ttk.Label(self.bottom_frame, textvariable=self.status_var, anchor="w").pack(side="left", padx=6)
        self.hyperlink_label = tk.Label(self.bottom_frame, text="Developed by Hamid Taghipourbibalan", font=("Cascadia Code", 8, "italic"), cursor="hand2")
        self.hyperlink_label.pack(side="right", padx=8, pady=2)
        self.hyperlink_label.bind("<Button-1>", lambda e: self.open_hyperlink("https://www.linkedin.com/in/hamid-taghipourbibalan-b7239088/"))
        self._build_context_menu()
        self.update_mode_button_styles()

    def _focused_widget_is_text_input(self, event=None):
        widget = getattr(event, "widget", None) if event is not None else None
        if widget is None:
            try:
                widget = self.root.focus_get()
            except Exception:
                widget = None
        if widget is None:
            return False
        text_like = (tk.Entry, tk.Text, ttk.Entry, ttk.Combobox, tk.Spinbox, ttk.Spinbox)
        if isinstance(widget, text_like):
            return True
        cls = widget.winfo_class()
        return cls in ("Entry", "Text", "TEntry", "TCombobox", "Spinbox", "TSpinbox")

    def _shortcut_allowed(self, event=None):
        return not self._focused_widget_is_text_input(event)

    def _setup_traces(self):
        self.gap_x.trace_add("write", lambda *args: self.redraw())
        self.gap_y.trace_add("write", lambda *args: self.redraw())

    def _build_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Undo", command=self.undo)
        self.context_menu.add_command(label="Redo", command=self.redo)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Bring to front", command=self.bring_to_front)
        self.context_menu.add_command(label="Send to back", command=self.send_to_back)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Duplicate", command=self.duplicate_selected)
        self.context_menu.add_command(label="Delete", command=self.delete_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Group selected", command=self.group_selected)
        self.context_menu.add_command(label="Ungroup selected", command=self.ungroup_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Rich text editor", command=self.open_rich_text_editor)
        self.context_menu.add_command(label="Trim selected image", command=self.trim_selected_image)
        self.context_menu.add_command(label="Toggle crop mode", command=self.toggle_crop_mode)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Exclude selected from labels", command=self.exclude_selected_from_labels)
        self.context_menu.add_command(label="Include selected in labels", command=self.include_selected_in_labels)
        self.context_menu.add_command(label="Regenerate labels", command=self.regenerate_labels)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Auto grid", command=self.auto_grid)
        self.context_menu.add_command(label="Distribute horizontally", command=self.distribute_h)
        self.context_menu.add_command(label="Distribute vertically", command=self.distribute_v)

    def _bind_events(self):
        self.canvas.bind("<Button-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Double-Button-1>", self.on_double_click)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self.on_canvas_mousewheel, add="+")
        self.canvas.bind("<Shift-MouseWheel>", self.on_canvas_mousewheel_horizontal, add="+")
        self.canvas.bind("<Control-MouseWheel>", self.on_canvas_ctrl_mousewheel, add="+")
        self.canvas.bind("<Button-4>", self.on_canvas_mousewheel_linux, add="+")
        self.canvas.bind("<Button-5>", self.on_canvas_mousewheel_linux, add="+")
        self.canvas.bind("<Shift-Button-4>", self.on_canvas_mousewheel_linux_horizontal, add="+")
        self.canvas.bind("<Shift-Button-5>", self.on_canvas_mousewheel_linux_horizontal, add="+")
        self.root.bind_all("<Delete>", lambda e: self.delete_selected() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-v>", lambda e: self.paste_from_clipboard() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-o>", lambda e: self.add_files() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-s>", lambda e: self.save_project() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-n>", lambda e: self.new_canvas() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-Shift-s>", lambda e: self.export_image("PNG") if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-g>", lambda e: self.auto_grid() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-a>", lambda e: self.select_all_items() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-z>", lambda e: self.undo() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-y>", lambda e: self.redo() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-plus>", lambda e: self.zoom_by(ZOOM_STEP) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-equal>", lambda e: self.zoom_by(ZOOM_STEP) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-minus>", lambda e: self.zoom_by(1 / ZOOM_STEP) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Control-0>", lambda e: self.reset_zoom() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Escape>", lambda e: self.cancel_modes() if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Left>", lambda e: self.move_selected_with_keyboard(-1, 0) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Right>", lambda e: self.move_selected_with_keyboard(1, 0) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Up>", lambda e: self.move_selected_with_keyboard(0, -1) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Down>", lambda e: self.move_selected_with_keyboard(0, 1) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Shift-Left>", lambda e: self.move_selected_with_keyboard(-10, 0) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Shift-Right>", lambda e: self.move_selected_with_keyboard(10, 0) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Shift-Up>", lambda e: self.move_selected_with_keyboard(0, -10) if self._shortcut_allowed(e) else None)
        self.root.bind_all("<Shift-Down>", lambda e: self.move_selected_with_keyboard(0, 10) if self._shortcut_allowed(e) else None)

    def open_hyperlink(self, url: str):
        try:
            webbrowser.open(url)
        except Exception:
            pass

    def mark_dirty(self):
        self.dirty = True

    def set_status(self, text: str):
        prefix = "● " if self.dirty else ""
        self.status_var.set(prefix + text)

    def update_zoom_label(self):
        self.zoom_var.set(f"{int(round(self.view_zoom * 100))}%")

    def configure_mode_button(self, btn, text_off, text_on, active, color):
        bg = color if active else "#ececec"
        fg = "white" if active else "black"
        btn.configure(text=text_on if active else text_off, bg=bg, fg=fg, activebackground=bg, activeforeground=fg, relief="sunken" if active else "raised")

    def update_mode_button_styles(self):
        if hasattr(self, "erase_btn"):
            self.configure_mode_button(self.erase_btn, "Erase OFF", "Erase ON", self.erase_mode, "#d83a3a")
        if hasattr(self, "crop_btn"):
            self.configure_mode_button(self.crop_btn, "Crop OFF", "Crop ON", self.crop_mode, "#d88922")
        if hasattr(self, "stat_symbol_btn"):
            self.configure_mode_button(self.stat_symbol_btn, "Place symbol OFF", "Place symbol ON", self.stat_mode == "symbol", "#8a2be2")
        if hasattr(self, "stat_bracket_btn"):
            self.configure_mode_button(self.stat_bracket_btn, "Bracket compare OFF", "Bracket compare ON", self.stat_mode == "compare", "#8a2be2")
        if hasattr(self, "label_seq_btn"):
            self.configure_mode_button(self.label_seq_btn, "Click label order OFF", "Click label order ON", self.label_sequence_mode, "#2f7d32")
        if hasattr(self, "label_align_btn"):
            self.configure_mode_button(self.label_align_btn, "Align labels by click OFF", "Align labels by click ON", self.label_align_mode, "#2f7d32")

    def cancel_modes(self):
        self.erase_mode = False
        self.crop_mode = False
        self.crop_start = None
        self.crop_current = None
        self.crop_target_id = None
        self.erase_start = None
        self.erase_current = None
        self.stat_mode = None
        self.stat_first_point = None
        self.label_align_mode = False
        self.label_anchor_panel_id = None
        self.label_sequence_mode = False
        self.update_mode_button_styles()
        self.redraw()
        self.set_status("Modes cancelled")

    def on_gap_change(self):
        self.redraw()
        self.set_status(f"Gaps set to X={self.gap_x.get()} Y={self.gap_y.get()}")

    def zoom_by(self, factor: float):
        new_zoom = clamp(self.view_zoom * factor, MIN_ZOOM, MAX_ZOOM)
        if abs(new_zoom - self.view_zoom) < 1e-9:
            return
        x0 = self.canvas.canvasx(self.canvas.winfo_width() / 2)
        y0 = self.canvas.canvasy(self.canvas.winfo_height() / 2)
        relx = 0.0 if self.canvas.bbox("all") is None else x0 / max(1, self.sx(self.canvas_w))
        rely = 0.0 if self.canvas.bbox("all") is None else y0 / max(1, self.sy(self.canvas_h))
        self.view_zoom = new_zoom
        self.update_zoom_label()
        self.redraw()
        self.canvas.xview_moveto(relx)
        self.canvas.yview_moveto(rely)

    def reset_zoom(self):
        self.view_zoom = 1.0
        self.update_zoom_label()
        self.redraw()

    def snapshot_state(self):
        return {
            "canvas_w": self.canvas_w,
            "canvas_h": self.canvas_h,
            "bg_color": self.bg_color,
            "view_zoom": self.view_zoom,
            "items": self._clone_items(self.items),
            "selected_ids": list(self.selected_ids),
            "anchor_selected_id": self.anchor_selected_id,
            "selected_label_panel_id": self.selected_label_panel_id,
            "next_id": self.next_id,
            "next_group_id": self.next_group_id,
        }

    def save_undo_state(self):
        self.undo_stack.append(self.snapshot_state())
        if len(self.undo_stack) > 100:
            self.undo_stack.pop(0)
        self.redo_stack.clear()
        self.mark_dirty()

    def _clone_items(self, items):
        cloned = []
        for item in items:
            if isinstance(item, PanelItem):
                cloned.append(PanelItem(
                    id=item.id,
                    source_path=item.source_path,
                    pil_image=item.pil_image.copy() if item.pil_image else None,
                    original_size=item.original_size,
                    x=item.x, y=item.y, w=item.w, h=item.h,
                    label=item.label,
                    show_label=item.show_label,
                    include_in_labels=item.include_in_labels,
                    label_font_size=item.label_font_size,
                    label_font_family=item.label_font_family,
                    label_color=item.label_color,
                    label_offset_x=item.label_offset_x,
                    label_offset_y=item.label_offset_y,
                    border_width=item.border_width,
                    border_color=item.border_color,
                    z_index=item.z_index,
                    group_id=item.group_id,
                ))
            elif isinstance(item, ShapeItem):
                cloned.append(ShapeItem(
                    id=item.id,
                    shape_type=item.shape_type,
                    x=item.x, y=item.y, w=item.w, h=item.h,
                    color=item.color,
                    width=item.width,
                    fill=item.fill,
                    fill_alpha=item.fill_alpha,
                    rotation=item.rotation,
                    text=item.text,
                    font_size=item.font_size,
                    font_family=item.font_family,
                    z_index=item.z_index,
                    group_id=item.group_id,
                    compare_style=item.compare_style,
                    bracket_height=item.bracket_height,
                    text_gap=item.text_gap,
                    auto_center_text=item.auto_center_text,
                    snap_to_tops=item.snap_to_tops,
                ))
            else:
                cloned.append(TextItem(
                    id=item.id,
                    text=item.text,
                    x=item.x, y=item.y, w=item.w, h=item.h,
                    font_size=item.font_size,
                    font_family=item.font_family,
                    bold=item.bold,
                    italic=item.italic,
                    rich_runs=json.loads(json.dumps(item.rich_runs)),
                    fill=item.fill,
                    align=item.align,
                    background=item.background,
                    outline=item.outline,
                    rotation=item.rotation,
                    line_spacing=item.line_spacing,
                    z_index=item.z_index,
                    group_id=item.group_id,
                ))
        return cloned

    def restore_state(self, state):
        self.canvas_w = state["canvas_w"]
        self.canvas_h = state["canvas_h"]
        self.bg_color = state.get("bg_color", DEFAULT_BG)
        self.view_zoom = state.get("view_zoom", self.view_zoom)
        self.items = self._clone_items(state["items"])
        self.selected_ids = list(state["selected_ids"])
        self.anchor_selected_id = state["anchor_selected_id"]
        self.selected_label_panel_id = state.get("selected_label_panel_id")
        self.next_id = state["next_id"]
        self.next_group_id = state["next_group_id"]
        self.canvas_w_entry.delete(0, "end")
        self.canvas_w_entry.insert(0, str(self.canvas_w))
        self.canvas_h_entry.delete(0, "end")
        self.canvas_h_entry.insert(0, str(self.canvas_h))
        self.update_zoom_label()
        self.refresh_selected_panel()
        self.redraw()

    def undo(self):
        if not self.undo_stack:
            return
        current = self.snapshot_state()
        self.redo_stack.append(current)
        state = self.undo_stack.pop()
        self.restore_state(state)
        self.mark_dirty()
        self.set_status("Undo")

    def redo(self):
        if not self.redo_stack:
            return
        current = self.snapshot_state()
        self.undo_stack.append(current)
        state = self.redo_stack.pop()
        self.restore_state(state)
        self.mark_dirty()
        self.set_status("Redo")

    def px_from_unit(self, value: float) -> int:
        unit = self.canvas_unit.get()
        dpi = self.export_dpi.get()
        if unit == "px":
            return int(round(value))
        if unit == "in":
            return int(round(value * dpi))
        if unit == "cm":
            return int(round((value / 2.54) * dpi))
        return int(round(value))

    def apply_canvas_size(self):
        try:
            w_in = float(self.canvas_w_entry.get())
            h_in = float(self.canvas_h_entry.get())
            new_w = max(100, self.px_from_unit(w_in))
            new_h = max(100, self.px_from_unit(h_in))
        except ValueError:
            messagebox.showerror("Invalid size", "Canvas width and height must be numeric.")
            return
        self.save_undo_state()
        self.canvas_w = new_w
        self.canvas_h = new_h
        self.redraw()
        self.set_status(f"Canvas set to {self.canvas_w} × {self.canvas_h} px")

    def get_next_id(self) -> int:
        out = self.next_id
        self.next_id += 1
        return out

    def get_next_group_id(self) -> int:
        out = self.next_group_id
        self.next_group_id += 1
        return out

    def get_next_append_position(self, w: int, h: int) -> Tuple[int, int]:
        margin = self.outer_margin.get()
        gap_x = max(self.gap_x.get(), SAFE_GAP)
        gap_y = max(self.gap_y.get(), SAFE_GAP)
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        if not panels:
            return margin, margin
        panels = sorted(panels, key=lambda p: (p.y, p.x))
        last = panels[-1]
        x = last.x + last.w + gap_x
        y = last.y
        if x + w > self.canvas_w - margin:
            x = margin
            y = max(p.y + p.h for p in panels) + gap_y
        return x, y

    def add_files(self):
        filetypes = [
            ("Supported", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.gif *.pdf"),
            ("Images", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.gif"),
            ("PDF files", "*.pdf"),
            ("All files", "*.*"),
        ]
        paths = filedialog.askopenfilenames(title="Select images or PDFs", filetypes=filetypes)
        if not paths:
            return
        self.save_undo_state()
        self.load_paths(list(paths))
        self.redraw()

    def on_drop(self, event):
        paths = parse_dnd_files(event.data)
        if paths:
            self.save_undo_state()
            self.load_paths(paths)
            self.redraw()

    def load_paths(self, paths: List[str]):
        added = 0
        for path in paths:
            path = str(path).strip()
            if not path:
                continue
            ext = Path(path).suffix.lower()
            try:
                if ext == ".pdf":
                    added += self._load_pdf(path)
                else:
                    self._load_image_path(path)
                    added += 1
            except Exception as e:
                messagebox.showerror("Import error", f"Could not open:\n{path}\n\n{e}")
        if added:
            self.set_status(f"Added {added} item(s)")

    def _load_image_path(self, path: str):
        img = Image.open(path).convert("RGBA")
        self._add_panel_from_pil(img, source_path=path)

    def _load_pdf(self, path: str) -> int:
        doc = fitz.open(path)
        count = 0
        try:
            for pno in range(len(doc)):
                page = doc[pno]
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGBA")
                self._add_panel_from_pil(img, source_path=f"{path} [page {pno + 1}]")
                count += 1
        finally:
            doc.close()
        return count

    def paste_from_clipboard(self):
        try:
            data = ImageGrab.grabclipboard()
        except Exception as e:
            messagebox.showerror("Clipboard error", f"Could not access clipboard.\n\n{e}")
            return
        if isinstance(data, Image.Image):
            self.save_undo_state()
            self._add_panel_from_pil(data.convert("RGBA"), source_path="clipboard")
            self.redraw()
            self.set_status("Pasted image from clipboard")
            return
        if isinstance(data, list):
            self.save_undo_state()
            self.load_paths([str(p) for p in data])
            self.redraw()
            return
        messagebox.showinfo("Clipboard", "No image or file list found in clipboard.")

    def _add_panel_from_pil(self, img: Image.Image, source_path: Optional[str] = None):
        ow, oh = img.size
        w, h = fit_size_keep_aspect(ow, oh, 350, 260)
        x, y = self.get_next_append_position(w, h)
        panel = PanelItem(
            id=self.get_next_id(),
            source_path=source_path,
            pil_image=img,
            original_size=(ow, oh),
            x=x,
            y=y,
            w=w,
            h=h,
            label="",
            show_label=True,
            include_in_labels=True,
            label_font_size=self.default_label_size.get(),
            label_font_family=self.default_font_family.get(),
            label_color="black",
            label_offset_x=self.default_label_offset_x.get(),
            label_offset_y=self.default_label_offset_y.get(),
            z_index=max((i.z_index for i in self.items), default=-1) + 1,
        )
        self.items.append(panel)
        self.selected_ids = [panel.id]
        self.anchor_selected_id = panel.id
        self.selected_label_panel_id = None

    def add_text_box(self, default_text="Text"):
        self.save_undo_state()
        item = TextItem(
            id=self.get_next_id(),
            text=default_text,
            rich_runs=make_plain_runs(default_text),
            x=60 + 15 * (len(self.items) % 10),
            y=60 + 15 * (len(self.items) % 10),
            w=320,
            h=90,
            font_size=self.default_text_size.get(),
            font_family=self.default_font_family.get(),
            line_spacing=DEFAULT_LINE_SPACING,
            z_index=max((i.z_index for i in self.items), default=-1) + 1,
        )
        self.items.append(item)
        self.selected_ids = [item.id]
        self.anchor_selected_id = item.id
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Added text box")

    def add_shape(self, shape_type):
        self.save_undo_state()
        base_x = 100 + 20 * (len(self.items) % 10)
        base_y = 100 + 20 * (len(self.items) % 10)
        z = max((i.z_index for i in self.items), default=0) + 1
        if shape_type in ("line", "arrow"):
            item = ShapeItem(id=self.get_next_id(), shape_type=shape_type, x=base_x, y=base_y, w=160, h=0, z_index=z)
        elif shape_type == "arrow_head":
            item = ShapeItem(id=self.get_next_id(), shape_type=shape_type, x=base_x, y=base_y, w=42, h=28, color="black", fill="black", fill_alpha=255, z_index=z)
        elif shape_type == "highlight_bar":
            item = ShapeItem(id=self.get_next_id(), shape_type=shape_type, x=base_x, y=base_y, w=180, h=28, color="#f2c200", width=1, fill="#f2c200", fill_alpha=90, z_index=z)
        else:
            item = ShapeItem(id=self.get_next_id(), shape_type=shape_type, x=base_x, y=base_y, w=150, h=100, z_index=z)
        self.items.append(item)
        self.selected_ids = [item.id]
        self.anchor_selected_id = item.id
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Added {shape_type}")

    def add_caption_below_last_figure(self):
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        if not panels:
            return
        self.save_undo_state()
        x1 = min(p.x for p in panels)
        x2 = max(p.x + p.w for p in panels)
        y2 = max(p.y + p.h for p in panels)
        caption = TextItem(
            id=self.get_next_id(),
            text="Caption",
            rich_runs=make_plain_runs("Caption"),
            x=x1,
            y=y2 + self.gap_y.get(),
            w=max(300, x2 - x1),
            h=120,
            font_size=self.default_text_size.get(),
            font_family=self.default_font_family.get(),
            line_spacing=DEFAULT_LINE_SPACING,
            z_index=max((i.z_index for i in self.items), default=-1) + 1,
        )
        self.items.append(caption)
        needed_h = caption.y + caption.h + self.outer_margin.get()
        if needed_h > self.canvas_h:
            self.canvas_h = needed_h
            self.canvas_h_entry.delete(0, "end")
            self.canvas_h_entry.insert(0, str(self.canvas_h))
        self.selected_ids = [caption.id]
        self.anchor_selected_id = caption.id
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Added caption below panels")

    def redraw(self):
        self.canvas.delete("all")
        self.canvas.config(scrollregion=self.scaled_scrollregion())
        self.canvas.create_rectangle(0, 0, self.sx(self.canvas_w), self.sy(self.canvas_h), fill=self.bg_color, outline="")
        self._draw_usable_margin_frame()
        if self.show_grid.get():
            self._draw_grid()
        for item in sorted(self.items, key=lambda x: x.z_index):
            if isinstance(item, PanelItem):
                self._draw_panel(item)
            elif isinstance(item, ShapeItem):
                self._draw_shape(item)
            else:
                self._draw_text_item(item)
        for item in self.get_selected_items():
            self._draw_selection(item, is_anchor=(item.id == self.anchor_selected_id))
        if self.erase_mode and self.erase_start and self.erase_current:
            x1, y1 = self.erase_start
            x2, y2 = self.erase_current
            self.canvas.create_rectangle(self.sx(x1), self.sy(y1), self.sx(x2), self.sy(y2), outline="red", dash=(6, 3), width=2)
        if self.crop_mode and self.crop_start and self.crop_current:
            x1, y1 = self.crop_start
            x2, y2 = self.crop_current
            self.canvas.create_rectangle(self.sx(x1), self.sy(y1), self.sx(x2), self.sy(y2), outline="#d88922", dash=(6, 3), width=2)
        if self.select_start and self.select_end:
            x1, y1 = self.select_start
            x2, y2 = self.select_end
            self.canvas.create_rectangle(self.sx(min(x1, x2)), self.sy(min(y1, y2)), self.sx(max(x1, x2)), self.sy(max(y1, y2)), outline="#2a73ff", width=2, dash=(4, 4))
        if self.stat_mode == "compare" and self.stat_first_point:
            x, y = self.stat_first_point
            self.canvas.create_oval(self.sx(x) - 4, self.sy(y) - 4, self.sx(x) + 4, self.sy(y) + 4, outline="#8a2be2", width=2)

    def _draw_usable_margin_frame(self):
        m = self.outer_margin.get()
        self.canvas.create_rectangle(self.sx(m), self.sy(m), self.sx(self.canvas_w - m), self.sy(self.canvas_h - m), outline="#a0a0a0", dash=(4, 3), width=1)

    def _draw_grid(self):
        step = 50
        for x in range(0, self.canvas_w + 1, step):
            self.canvas.create_line(self.sx(x), 0, self.sx(x), self.sy(self.canvas_h), fill="#eeeeee")
        for y in range(0, self.canvas_h + 1, step):
            self.canvas.create_line(0, self.sy(y), self.sx(self.canvas_w), self.sy(y), fill="#eeeeee")

    def _panel_preview_photo(self, item: PanelItem):
        key = (item.w, item.h, round(self.view_zoom, 4), id(item.pil_image))
        if item._preview_cache_key == key and item._preview_cache_photo is not None:
            return item._preview_cache_photo
        dw = max(1, self.sx(item.w))
        dh = max(1, self.sy(item.h))
        preview = item.pil_image.resize((dw, dh), RESAMPLE_LANCZOS)
        photo = ImageTk.PhotoImage(preview)
        item._preview_cache_key = key
        item._preview_cache_photo = photo
        return photo

    def panel_label_bbox(self, item: PanelItem) -> Optional[Tuple[int, int, int, int]]:
        if not item.show_label or not item.label or not item.include_in_labels:
            return None
        font = get_font(item.label_font_family, max(6, item.label_font_size), bold=True)
        dummy = Image.new("RGBA", (10, 10), (255, 255, 255, 0))
        draw = ImageDraw.Draw(dummy)
        bbox = draw.textbbox((0, 0), item.label, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        x1 = item.x + item.label_offset_x
        y1 = item.y + item.label_offset_y
        return x1, y1, x1 + tw, y1 + th

    def label_absolute_position(self, panel: PanelItem) -> Tuple[int, int]:
        return panel.x + panel.label_offset_x, panel.y + panel.label_offset_y

    def point_hits_panel_label(self, item: PanelItem, px: int, py: int) -> bool:
        bbox = self.panel_label_bbox(item)
        if bbox is None:
            return False
        x1, y1, x2, y2 = bbox
        pad = 8
        return x1 - pad <= px <= x2 + pad and y1 - pad <= py <= y2 + pad

    def get_panel_label_at(self, x: int, y: int) -> Optional[PanelItem]:
        for item in sorted([i for i in self.items if isinstance(i, PanelItem)], key=lambda z: z.z_index, reverse=True):
            if self.point_hits_panel_label(item, x, y):
                return item
        return None

    def _draw_panel(self, item: PanelItem):
        if item.pil_image is None:
            return
        photo = self._panel_preview_photo(item)
        item.tk_preview = photo
        self.canvas.create_image(self.sx(item.x), self.sy(item.y), image=photo, anchor="nw")
        if item.border_width > 0:
            self.canvas.create_rectangle(self.sx(item.x), self.sy(item.y), self.sx(item.x + item.w), self.sy(item.y + item.h), outline=item.border_color, width=max(1, int(round(item.border_width * self.view_zoom))))
        if item.show_label and item.label and item.include_in_labels:
            font_style = (item.label_font_family, max(6, int(round(item.label_font_size * self.view_zoom))), "bold")
            tx = self.sx(item.x + item.label_offset_x)
            ty = self.sy(item.y + item.label_offset_y)
            self.canvas.create_text(tx, ty, text=item.label, anchor="nw", font=font_style, fill=item.label_color)
        lb = self.panel_label_bbox(item)
        if lb and item.id in self.selected_ids:
            if self.selected_label_panel_id == item.id:
                self.canvas.create_rectangle(self.sx(lb[0]), self.sy(lb[1]), self.sx(lb[2]), self.sy(lb[3]), outline="#ff7a00", width=2)
            else:
                self.canvas.create_rectangle(self.sx(lb[0]), self.sy(lb[1]), self.sx(lb[2]), self.sy(lb[3]), outline="#c84d00", dash=(3, 2))

    def scaled_points(self, pts):
        out = []
        for x, y in pts:
            out.extend([self.sx(x), self.sy(y)])
        return out

    def _draw_shape(self, item: ShapeItem):
        width = max(1, int(round(item.width * self.view_zoom)))
        if item.shape_type == "rectangle":
            pts = item.visual_points()
            if item.rotation:
                fill = item.fill if item.fill_alpha >= 255 else ""
                self.canvas.create_polygon(self.scaled_points(pts), outline=item.color, fill=fill, width=width)
            else:
                self.canvas.create_rectangle(self.sx(item.x), self.sy(item.y), self.sx(item.x + item.w), self.sy(item.y + item.h), outline=item.color, width=width, fill=item.fill if item.fill and item.fill_alpha >= 255 else "")
        elif item.shape_type == "highlight_bar":
            pts = item.visual_points()
            self.canvas.create_polygon(self.scaled_points(pts), outline=item.color, fill=item.fill or item.color, stipple="gray50", width=width)
        elif item.shape_type == "circle":
            self.canvas.create_oval(self.sx(item.x), self.sy(item.y), self.sx(item.x + item.w), self.sy(item.y + item.h), outline=item.color, width=width, fill=item.fill if item.fill and item.fill_alpha >= 255 else "")
        elif item.shape_type in ("line", "arrow"):
            p1, p2 = item.visual_points()[:2]
            kwargs = {"fill": item.color, "width": width}
            if item.shape_type == "arrow":
                kwargs["arrow"] = "last"
                kwargs["arrowshape"] = (max(8, int(12 * self.view_zoom)), max(10, int(14 * self.view_zoom)), max(4, int(5 * self.view_zoom)))
            self.canvas.create_line(self.sx(p1[0]), self.sy(p1[1]), self.sx(p2[0]), self.sy(p2[1]), **kwargs)
        elif item.shape_type == "arrow_head":
            pts = item.visual_points()
            self.canvas.create_polygon(self.scaled_points(pts), outline=item.color, fill=item.fill or item.color, width=width)
        elif item.is_stat_compare():
            pts = item.visual_points()
            coords = self.scaled_points(pts)
            self.canvas.create_line(*coords, fill=item.color, width=width)
            if item.text:
                tx, ty = item.stat_text_position()
                font_style = (item.font_family, max(6, int(round(item.font_size * self.view_zoom))), "bold")
                self.canvas.create_text(self.sx(tx), self.sy(ty), text=item.text, anchor="s", font=font_style, fill=item.color)

    def _compress_line_segments(self, draw, chars, font_family, font_size):
        if not chars:
            return []
        out = []
        current = {"text": chars[0]["text"], "bold": chars[0]["bold"], "italic": chars[0]["italic"]}
        for ch in chars[1:]:
            if ch["bold"] == current["bold"] and ch["italic"] == current["italic"]:
                current["text"] += ch["text"]
            else:
                out.append(current)
                current = {"text": ch["text"], "bold": ch["bold"], "italic": ch["italic"]}
        out.append(current)
        enriched = []
        for seg in out:
            font = get_font(font_family, font_size, seg["bold"], seg["italic"])
            bbox = draw.textbbox((0, 0), seg["text"], font=font)
            enriched.append({"text": seg["text"], "bold": seg["bold"], "italic": seg["italic"], "width": bbox[2] - bbox[0], "height": bbox[3] - bbox[1]})
        return enriched

    def wrap_runs_to_width(self, draw, runs, font_family, font_size, max_width):
        text = runs_to_text(runs)
        if text == "":
            return [[]]
        chars = []
        for run in normalize_runs(runs):
            for ch in run["text"]:
                chars.append({"text": ch, "bold": bool(run.get("bold", False)), "italic": bool(run.get("italic", False))})
        lines = []
        current_line = []
        current_width = 0
        last_space_idx = -1
        for ch in chars:
            if ch["text"] == "\n":
                lines.append(self._compress_line_segments(draw, current_line, font_family, font_size))
                current_line = []
                current_width = 0
                last_space_idx = -1
                continue
            font = get_font(font_family, font_size, ch["bold"], ch["italic"])
            bbox = draw.textbbox((0, 0), ch["text"], font=font)
            cw = bbox[2] - bbox[0]
            if current_line and current_width + cw > max_width:
                if last_space_idx >= 0:
                    left = current_line[:last_space_idx]
                    right = current_line[last_space_idx + 1:]
                    lines.append(self._compress_line_segments(draw, left, font_family, font_size))
                    current_line = right + [ch]
                    current_width = sum(seg.get("width", 0) for seg in self._compress_line_segments(draw, current_line, font_family, font_size))
                    last_space_idx = -1
                else:
                    lines.append(self._compress_line_segments(draw, current_line, font_family, font_size))
                    current_line = [ch]
                    current_width = cw
                    last_space_idx = -1
            else:
                current_line.append(ch)
                current_width += cw
                if ch["text"] == " ":
                    last_space_idx = len(current_line) - 1
        if current_line:
            lines.append(self._compress_line_segments(draw, current_line, font_family, font_size))
        return lines

    def _textbox_cache_key(self, item: TextItem):
        return (item.w, item.h, round(self.view_zoom, 4), item.font_size, item.font_family, item.fill, item.align, item.background, item.outline, item.rotation, round(item.line_spacing, 3), json.dumps(item.rich_runs, sort_keys=True))

    def _draw_justified_line(self, draw, line, x, y, max_width, font_family, fsize, fill):
        chars = []
        for seg in line:
            for ch in seg["text"]:
                chars.append({"text": ch, "bold": seg["bold"], "italic": seg["italic"]})
        total_width = 0
        space_indices = []
        widths = []
        for idx, ch in enumerate(chars):
            font = get_font(font_family, fsize, ch["bold"], ch["italic"])
            bbox = draw.textbbox((0, 0), ch["text"], font=font)
            w = bbox[2] - bbox[0]
            widths.append(w)
            total_width += w
            if ch["text"] == " ":
                space_indices.append(idx)
        extra = max(0, max_width - total_width)
        base_extra = extra // len(space_indices) if space_indices else 0
        remainder = extra % len(space_indices) if space_indices else 0
        cx = x
        max_h = 0
        for idx, ch in enumerate(chars):
            font = get_font(font_family, fsize, ch["bold"], ch["italic"])
            draw.text((cx, y), ch["text"], fill=fill, font=font)
            cx += widths[idx]
            if ch["text"] == " " and space_indices:
                cx += base_extra
                if remainder > 0:
                    cx += 1
                    remainder -= 1
            bbox = draw.textbbox((0, 0), ch["text"], font=font)
            max_h = max(max_h, bbox[3] - bbox[1])
        return max_h

    def _build_textbox_image(self, item: TextItem, scale: float = 1.0) -> Image.Image:
        bw = max(1, int(round(item.w * scale)))
        bh = max(1, int(round(item.h * scale)))
        base = Image.new("RGBA", (bw, bh), (255, 255, 255, 0))
        draw = ImageDraw.Draw(base)
        if item.background:
            draw.rectangle([0, 0, bw - 1, bh - 1], fill=item.background, outline=item.outline if item.outline else None)
        elif item.outline:
            draw.rectangle([0, 0, bw - 1, bh - 1], outline=item.outline)
        pad = max(4, int(round(6 * scale)))
        fsize = max(6, int(round(item.font_size * scale)))
        max_width = max(20, bw - 2 * pad)
        lines = self.wrap_runs_to_width(draw=draw, runs=item.rich_runs, font_family=item.font_family, font_size=fsize, max_width=max_width)
        y = pad
        extra_factor = max(0.0, item.line_spacing - 1.0)
        for line_idx, line in enumerate(lines):
            line_width = sum(seg["width"] for seg in line) if line else 0
            is_last = line_idx == len(lines) - 1
            has_spaces = any(" " in seg["text"] for seg in line)
            if item.align == "center":
                x = (bw - line_width) / 2
                max_h = 0
                for seg in line:
                    font = get_font(item.font_family, fsize, seg["bold"], seg["italic"])
                    draw.text((x, y), seg["text"], fill=item.fill, font=font)
                    x += seg["width"]
                    max_h = max(max_h, seg["height"])
            elif item.align == "right":
                x = bw - pad - line_width
                max_h = 0
                for seg in line:
                    font = get_font(item.font_family, fsize, seg["bold"], seg["italic"])
                    draw.text((x, y), seg["text"], fill=item.fill, font=font)
                    x += seg["width"]
                    max_h = max(max_h, seg["height"])
            elif item.align == "justify" and not is_last and has_spaces:
                x = pad
                max_h = self._draw_justified_line(draw, line, x, y, max_width, item.font_family, fsize, item.fill)
            else:
                x = pad
                max_h = 0
                for seg in line:
                    font = get_font(item.font_family, fsize, seg["bold"], seg["italic"])
                    draw.text((x, y), seg["text"], fill=item.fill, font=font)
                    x += seg["width"]
                    max_h = max(max_h, seg["height"])
            step = max_h + max(2, int(round(4 * scale))) + int(round(max_h * extra_factor))
            y += step
        if item.rotation in (90, 180, 270):
            base = base.rotate(item.rotation, expand=True)
        return base

    def _text_preview_photo(self, item: TextItem):
        key = self._textbox_cache_key(item)
        if item._preview_cache_key == key and item._preview_cache_photo is not None:
            return item._preview_cache_photo
        img = self._build_textbox_image(item, scale=self.view_zoom)
        photo = ImageTk.PhotoImage(img)
        item._preview_cache_key = key
        item._preview_cache_photo = photo
        return photo

    def _draw_text_item(self, item: TextItem):
        photo = self._text_preview_photo(item)
        item.tk_preview = photo
        self.canvas.create_image(self.sx(item.x), self.sy(item.y), image=photo, anchor="nw")

    def _draw_selection(self, item: CanvasItem, is_anchor=False):
        x1, y1, x2, y2 = item.bbox()
        outline = "#ff7a00" if is_anchor else "#2a73ff"
        self.canvas.create_rectangle(self.sx(x1), self.sy(y1), self.sx(x2), self.sy(y2), outline=outline, width=2, dash=(5, 3))
        hs = HANDLE_SIZE
        self.canvas.create_rectangle(self.sx(x2) - hs, self.sy(y2) - hs, self.sx(x2), self.sy(y2), fill=outline, outline=outline)

    def get_item_at(self, x: int, y: int) -> Optional[CanvasItem]:
        for item in sorted(self.items, key=lambda z: z.z_index, reverse=True):
            if item.contains(x, y):
                return item
        return None

    def get_item_by_id(self, item_id: int) -> Optional[CanvasItem]:
        for item in self.items:
            if item.id == item_id:
                return item
        return None

    def get_selected_items(self) -> List[CanvasItem]:
        result = []
        for item_id in self.selected_ids:
            item = self.get_item_by_id(item_id)
            if item is not None:
                result.append(item)
        return result

    def select_all_items(self):
        self.selected_ids = [item.id for item in self.items]
        self.anchor_selected_id = self.selected_ids[-1] if self.selected_ids else None
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Selected all items")

    def on_canvas_motion(self, event):
        x, y = self.model_from_event(event)
        label_item = self.get_panel_label_at(x, y)
        item = self.get_item_at(x, y)
        cursor = ""
        if self.crop_mode or self.erase_mode or self.stat_mode:
            cursor = "crosshair"
        elif self.label_align_mode or self.label_sequence_mode:
            cursor = "hand2"
        elif label_item is not None:
            cursor = "fleur"
        elif item is not None:
            handle_model = self.handle_model_size()
            if item.resize_handle_hit(x, y, handle_model):
                cursor = "sizing"
            else:
                cursor = "fleur"
        self.canvas.config(cursor=cursor)

    def on_canvas_press(self, event):
        self.canvas.focus_set()
        x, y = self.model_from_event(event)
        if self.crop_mode:
            item = self.get_item_at(x, y)
            if isinstance(item, PanelItem):
                self.crop_target_id = item.id
                self.crop_start = (x, y)
                self.crop_current = (x, y)
            return
        if self.erase_mode:
            item = self.get_item_at(x, y)
            if isinstance(item, PanelItem):
                self.erase_start = (x, y)
                self.erase_current = (x, y)
            return
        if self.stat_mode:
            if self.stat_mode == "symbol":
                self.add_stat_symbol_at(x, y)
            elif self.stat_mode == "compare":
                self.handle_stat_compare_click(x, y)
            return
        if self.label_align_mode:
            panel = self.get_panel_label_at(x, y)
            if panel is not None:
                self.handle_label_align_click(panel)
            return
        if self.label_sequence_mode:
            label_panel = self.get_panel_label_at(x, y)
            item = label_panel if label_panel is not None else self.get_item_at(x, y)
            if isinstance(item, PanelItem):
                self.handle_label_sequence_click(item)
            return
        label_hit = self.get_panel_label_at(x, y)
        if label_hit is not None:
            ctrl = bool(event.state & CTRL_MASK)
            if ctrl and label_hit.id in self.selected_ids:
                self.selected_ids = [i for i in self.selected_ids if i != label_hit.id]
                self.selected_label_panel_id = None
                self.anchor_selected_id = self.selected_ids[-1] if self.selected_ids else None
                self.refresh_selected_panel()
                self.redraw()
                return
            if ctrl:
                if label_hit.id not in self.selected_ids:
                    self.selected_ids.append(label_hit.id)
                self.anchor_selected_id = label_hit.id
            else:
                self.selected_ids = [label_hit.id]
                self.anchor_selected_id = label_hit.id
            self.selected_label_panel_id = label_hit.id
            self.drag_start = (x, y)
            self.drag_mode = "move_label"
            self.drag_origin_map = {label_hit.id: (label_hit.label_offset_x, label_hit.label_offset_y, label_hit.w, label_hit.h)}
            self.before_drag_state = self.snapshot_state()
            self.refresh_selected_panel()
            self.redraw()
            return
        item = self.get_item_at(x, y)
        ctrl = bool(event.state & CTRL_MASK)
        self.drag_start = (x, y)
        self.drag_mode = None
        self.drag_origin_map = {}
        self.before_drag_state = None
        if item is None:
            self.drag_mode = "select_rect"
            self.select_start = (x, y)
            self.select_end = (x, y)
            self.rect_add_mode = ctrl
            if not ctrl:
                self.selected_ids = []
                self.anchor_selected_id = None
                self.selected_label_panel_id = None
            self.refresh_selected_panel()
            self.redraw()
            return
        self.selected_label_panel_id = None
        clicked_ids = [item.id]
        if item.group_id is not None:
            clicked_ids = [m.id for m in self.get_group_members(item.group_id)]
        if ctrl:
            if all(clicked_id in self.selected_ids for clicked_id in clicked_ids):
                self.selected_ids = [i for i in self.selected_ids if i not in clicked_ids]
                if self.anchor_selected_id in clicked_ids:
                    self.anchor_selected_id = self.selected_ids[-1] if self.selected_ids else None
                self.refresh_selected_panel()
                self.redraw()
                return
            for clicked_id in clicked_ids:
                if clicked_id not in self.selected_ids:
                    self.selected_ids.append(clicked_id)
            self.anchor_selected_id = item.id
        else:
            self.selected_ids = clicked_ids
            self.anchor_selected_id = item.id
        self.before_drag_state = self.snapshot_state()
        selected_items = self.get_selected_items()
        hit_resize = item.resize_handle_hit(x, y, self.handle_model_size())
        if hit_resize and len(selected_items) == 1:
            self.drag_mode = "resize"
            self.drag_origin_map[item.id] = (item.x, item.y, item.w, item.h)
        else:
            self.drag_mode = "move"
            for sel in selected_items:
                self.drag_origin_map[sel.id] = (sel.x, sel.y, sel.w, sel.h)
        self.refresh_selected_panel()
        self.redraw()

    def proportional_size(self, ow, oh, dx, dy, min_w, min_h):
        new_w = ow + dx
        new_h = oh + dy
        if ow == 0 and oh == 0:
            return max(min_w, new_w), max(min_h, new_h)
        if abs(ow) < 1:
            return 0, new_h
        if abs(oh) < 1:
            return new_w, 0
        ratio = ow / oh
        if abs(dx) >= abs(dy):
            new_h = new_w / ratio
        else:
            new_w = new_h * ratio
        if new_w >= 0:
            new_w = max(min_w, new_w)
        else:
            new_w = min(-min_w, new_w)
        if new_h >= 0:
            new_h = max(min_h, new_h)
        else:
            new_h = min(-min_h, new_h)
        return int(round(new_w)), int(round(new_h))

    def on_canvas_drag(self, event):
        x, y = self.model_from_event(event)
        if self.crop_mode:
            if self.crop_start:
                self.crop_current = (x, y)
                self.redraw()
            return
        if self.erase_mode:
            if self.erase_start:
                self.erase_current = (x, y)
                self.redraw()
            return
        if self.drag_mode == "select_rect":
            self.select_end = (x, y)
            self.redraw()
            return
        if not self.drag_mode or not self.drag_origin_map:
            return
        dx = x - self.drag_start[0]
        dy = y - self.drag_start[1]
        shift = bool(event.state & SHIFT_MASK)
        if self.drag_mode == "move":
            for item_id, geom in self.drag_origin_map.items():
                item = self.get_item_by_id(item_id)
                if item is None:
                    continue
                ox, oy, ow, oh = geom
                item.x = max(0, ox + dx)
                item.y = max(0, oy + dy)
        elif self.drag_mode == "resize":
            item_id = next(iter(self.drag_origin_map.keys()))
            item = self.get_item_by_id(item_id)
            if item is not None:
                ox, oy, ow, oh = self.drag_origin_map[item_id]
                if isinstance(item, ShapeItem) and item.shape_type in ("line", "arrow"):
                    if shift:
                        item.w, item.h = self.proportional_size(ow, oh, dx, dy, 1, 1)
                    else:
                        new_w = ow + dx
                        new_h = oh + dy
                        snap = 12
                        if abs(new_h) < snap:
                            new_h = 0
                        if abs(new_w) < snap:
                            new_w = 0
                        item.w = int(round(new_w))
                        item.h = int(round(new_h))
                elif isinstance(item, ShapeItem) and item.is_stat_compare():
                    item.w = int(round(ow + dx))
                    item.h = int(round(oh + dy))
                else:
                    min_size = 8 if isinstance(item, ShapeItem) else MIN_PANEL_SIZE
                    if shift:
                        new_w, new_h = self.proportional_size(ow, oh, dx, dy, min_size, min_size)
                        item.w = max(min_size, abs(new_w)) if not isinstance(item, ShapeItem) else new_w
                        item.h = max(min_size, abs(new_h)) if not isinstance(item, ShapeItem) else new_h
                    else:
                        item.w = max(min_size, int(round(ow + dx)))
                        item.h = max(min_size, int(round(oh + dy)))
                if isinstance(item, (PanelItem, TextItem)):
                    item._preview_cache_key = None
                    item._preview_cache_photo = None
        elif self.drag_mode == "move_label":
            item_id = next(iter(self.drag_origin_map.keys()))
            item = self.get_item_by_id(item_id)
            if isinstance(item, PanelItem):
                oox, ooy, _, _ = self.drag_origin_map[item_id]
                item.label_offset_x = oox + dx
                item.label_offset_y = ooy + dy
        self.refresh_selected_panel(live=True)
        self.redraw()

    def on_canvas_release(self, event):
        x, y = self.model_from_event(event)
        if self.crop_mode and self.crop_start:
            x1, y1 = self.crop_start
            x2, y2 = x, y
            target = self.get_item_by_id(self.crop_target_id) if self.crop_target_id else None
            self.apply_crop_rectangle(x1, y1, x2, y2, target)
            self.crop_start = None
            self.crop_current = None
            self.crop_target_id = None
            self.redraw()
            return
        if self.erase_mode and self.erase_start:
            x1, y1 = self.erase_start
            x2, y2 = x, y
            self.apply_erase_rectangle(x1, y1, x2, y2)
            self.erase_start = None
            self.erase_current = None
            self.redraw()
            return
        if self.drag_mode == "select_rect":
            if self.select_start and self.select_end:
                sx1, sy1 = self.select_start
                sx2, sy2 = self.select_end
                select_rect = (min(sx1, sx2), min(sy1, sy2), max(sx1, sx2), max(sy1, sy2))
                new_selected = [item.id for item in self.items if rects_overlap(select_rect, item.bbox(), pad=0)]
                if self.rect_add_mode:
                    self.selected_ids = list(dict.fromkeys(self.selected_ids + new_selected))
                else:
                    self.selected_ids = new_selected
                self.anchor_selected_id = self.selected_ids[-1] if self.selected_ids else None
                self.selected_label_panel_id = None
                self.refresh_selected_panel()
            self.select_start = None
            self.select_end = None
            self.drag_mode = None
            self.redraw()
            return
        if self.drag_mode and self.before_drag_state is not None:
            after = self.snapshot_state()
            if self.state_changed(self.before_drag_state, after):
                self.undo_stack.append(self.before_drag_state)
                if len(self.undo_stack) > 100:
                    self.undo_stack.pop(0)
                self.redo_stack.clear()
                self.mark_dirty()
            self.drag_mode = None
            self.drag_origin_map = {}
            self.before_drag_state = None
            self.set_status("Canvas updated")

    def state_changed(self, a, b) -> bool:
        if a["canvas_w"] != b["canvas_w"] or a["canvas_h"] != b["canvas_h"]:
            return True
        if a.get("selected_label_panel_id") != b.get("selected_label_panel_id"):
            return True
        if len(a["items"]) != len(b["items"]):
            return True
        for ia, ib in zip(a["items"], b["items"]):
            if type(ia) != type(ib):
                return True
            if isinstance(ia, PanelItem):
                if (ia.x, ia.y, ia.w, ia.h, ia.group_id, ia.z_index, ia.label, ia.show_label, ia.include_in_labels, ia.label_offset_x, ia.label_offset_y, ia.label_color, ia.label_font_size, ia.label_font_family) != (ib.x, ib.y, ib.w, ib.h, ib.group_id, ib.z_index, ib.label, ib.show_label, ib.include_in_labels, ib.label_offset_x, ib.label_offset_y, ib.label_color, ib.label_font_size, ib.label_font_family):
                    return True
            elif isinstance(ia, ShapeItem):
                if (ia.x, ia.y, ia.w, ia.h, ia.group_id, ia.z_index, ia.shape_type, ia.color, ia.width, ia.fill, ia.fill_alpha, ia.rotation, ia.text, ia.font_size, ia.font_family, ia.compare_style, ia.bracket_height, ia.text_gap, ia.auto_center_text, ia.snap_to_tops) != (ib.x, ib.y, ib.w, ib.h, ib.group_id, ib.z_index, ib.shape_type, ib.color, ib.width, ib.fill, ib.fill_alpha, ib.rotation, ib.text, ib.font_size, ib.font_family, ib.compare_style, ib.bracket_height, ib.text_gap, ib.auto_center_text, ib.snap_to_tops):
                    return True
            else:
                if (ia.x, ia.y, ia.w, ia.h, ia.group_id, ia.z_index, ia.text, ia.font_size, ia.font_family, ia.align, ia.rotation, ia.fill, ia.bold, ia.italic, ia.line_spacing, ia.rich_runs) != (ib.x, ib.y, ib.w, ib.h, ib.group_id, ib.z_index, ib.text, ib.font_size, ib.font_family, ib.align, ib.rotation, ib.fill, ib.bold, ib.italic, ib.line_spacing, ib.rich_runs):
                    return True
        return False

    def on_double_click(self, event):
        x, y = self.model_from_event(event)
        label_panel = self.get_panel_label_at(x, y)
        if label_panel is not None:
            self.selected_ids = [label_panel.id]
            self.anchor_selected_id = label_panel.id
            self.selected_label_panel_id = label_panel.id
            self.refresh_selected_panel()
            new_label = simpledialog.askstring("Edit label", "Panel label:", initialvalue=label_panel.label, parent=self.root)
            if new_label is not None:
                self.save_undo_state()
                label_panel.label = new_label
                label_panel.include_in_labels = True
                label_panel.show_label = True
                self.refresh_selected_panel()
                self.redraw()
            return
        item = self.get_item_at(x, y)
        if item is None:
            return
        if isinstance(item, TextItem):
            self.selected_ids = [item.id]
            self.anchor_selected_id = item.id
            self.selected_label_panel_id = None
            self.refresh_selected_panel()
            self.open_rich_text_editor()
            return
        if isinstance(item, PanelItem):
            self.selected_ids = [item.id]
            self.anchor_selected_id = item.id
            self.selected_label_panel_id = item.id
            self.refresh_selected_panel()
            new_label = simpledialog.askstring("Edit label", "Panel label:", initialvalue=item.label, parent=self.root)
            if new_label is not None:
                self.save_undo_state()
                item.label = new_label
                item.include_in_labels = True
                item.show_label = True
                self.refresh_selected_panel()
                self.redraw()

    def on_right_click(self, event):
        x, y = self.model_from_event(event)
        label_panel = self.get_panel_label_at(x, y)
        item = label_panel if label_panel is not None else self.get_item_at(x, y)
        if item is not None and item.id not in self.selected_ids:
            self.selected_ids = [item.id]
            self.anchor_selected_id = item.id
            self.selected_label_panel_id = item.id if label_panel is not None else None
            self.refresh_selected_panel()
            self.redraw()
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def refresh_selected_panel(self, live=False):
        if not self.selected_ids:
            self.sel_type_var.set("-")
            self.sel_label_var.set("")
            self.sel_text_var.set("")
            self.sel_x_var.set("0")
            self.sel_y_var.set("0")
            self.sel_w_var.set("0")
            self.sel_h_var.set("0")
            self.sel_font_var.set(str(DEFAULT_TEXT_FONT_SIZE))
            self.sel_font_family_var.set(self.default_font_family.get())
            self.sel_show_label_var.set(True)
            self.sel_include_label_var.set(True)
            self.sel_align_var.set("left")
            self.sel_bold_var.set(False)
            self.sel_italic_var.set(False)
            self.sel_rotation_var.set("0")
            self.sel_line_spacing_var.set(str(DEFAULT_LINE_SPACING))
            return
        if len(self.selected_ids) > 1:
            self.sel_type_var.set(f"{len(self.selected_ids)} items")
            anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else None
            if anchor:
                self.sel_x_var.set(str(anchor.x))
                self.sel_y_var.set(str(anchor.y))
                self.sel_w_var.set(str(anchor.w))
                self.sel_h_var.set(str(anchor.h))
            return
        sel = self.get_item_by_id(self.selected_ids[0])
        if sel is None:
            return
        if isinstance(sel, PanelItem) and self.selected_label_panel_id == sel.id:
            bbox = self.panel_label_bbox(sel)
            lx, ly = self.label_absolute_position(sel)
            self.sel_type_var.set("panel label")
            self.sel_label_var.set(sel.label)
            self.sel_text_var.set("")
            self.sel_x_var.set(str(lx))
            self.sel_y_var.set(str(ly))
            if bbox:
                self.sel_w_var.set(str(max(0, bbox[2] - bbox[0])))
                self.sel_h_var.set(str(max(0, bbox[3] - bbox[1])))
            else:
                self.sel_w_var.set("0")
                self.sel_h_var.set("0")
            self.sel_font_var.set(str(sel.label_font_size))
            self.sel_font_family_var.set(sel.label_font_family)
            self.sel_show_label_var.set(sel.show_label)
            self.sel_include_label_var.set(sel.include_in_labels)
            self.sel_align_var.set("left")
            self.sel_bold_var.set(False)
            self.sel_italic_var.set(False)
            self.sel_rotation_var.set("0")
            self.sel_line_spacing_var.set(str(DEFAULT_LINE_SPACING))
            if not live:
                self.set_status(f"Selected label {sel.label or '#'} for panel #{sel.id}")
            return
        self.sel_type_var.set(sel.kind if not isinstance(sel, ShapeItem) else f"shape: {sel.shape_type}")
        self.sel_x_var.set(str(sel.x))
        self.sel_y_var.set(str(sel.y))
        self.sel_w_var.set(str(sel.w))
        self.sel_h_var.set(str(sel.h))
        if isinstance(sel, PanelItem):
            self.sel_label_var.set(sel.label)
            self.sel_text_var.set("")
            self.sel_font_var.set(str(sel.label_font_size))
            self.sel_font_family_var.set(sel.label_font_family)
            self.sel_show_label_var.set(sel.show_label)
            self.sel_include_label_var.set(sel.include_in_labels)
            self.sel_align_var.set("left")
            self.sel_bold_var.set(False)
            self.sel_italic_var.set(False)
            self.sel_rotation_var.set("0")
            self.sel_line_spacing_var.set(str(DEFAULT_LINE_SPACING))
        elif isinstance(sel, ShapeItem):
            self.sel_label_var.set("")
            self.sel_text_var.set(sel.text)
            self.sel_font_var.set(str(sel.width))
            self.sel_font_family_var.set(sel.font_family)
            self.sel_show_label_var.set(True)
            self.sel_include_label_var.set(False)
            self.sel_align_var.set("left")
            self.sel_bold_var.set(False)
            self.sel_italic_var.set(False)
            self.sel_rotation_var.set(str(int(sel.rotation) if float(sel.rotation).is_integer() else sel.rotation))
            self.sel_line_spacing_var.set(str(DEFAULT_LINE_SPACING))
        else:
            sel.sync_text_from_runs()
            self.sel_label_var.set("")
            self.sel_text_var.set(sel.text)
            self.sel_font_var.set(str(sel.font_size))
            self.sel_font_family_var.set(sel.font_family)
            self.sel_show_label_var.set(True)
            self.sel_include_label_var.set(False)
            self.sel_align_var.set(sel.align)
            self.sel_bold_var.set(sel.bold)
            self.sel_italic_var.set(sel.italic)
            self.sel_rotation_var.set(str(sel.rotation))
            self.sel_line_spacing_var.set(str(sel.line_spacing))
        if not live:
            self.set_status(f"Selected {self.sel_type_var.get()} #{sel.id}")

    def rotate_text_item_box(self, item: TextItem, new_rotation: int):
        old_rotation = item.rotation
        old_parity = old_rotation in (90, 270)
        new_parity = new_rotation in (90, 270)
        if old_parity != new_parity:
            item.w, item.h = item.h, item.w
        item.rotation = new_rotation
        item._preview_cache_key = None
        item._preview_cache_photo = None

    def apply_selected_label_visibility_quick(self):
        selected = [i for i in self.get_selected_items() if isinstance(i, PanelItem)]
        if not selected:
            return
        self.save_undo_state()
        for item in selected:
            item.show_label = self.sel_show_label_var.get()
            item.include_in_labels = self.sel_include_label_var.get()
            if item.include_in_labels and not item.label:
                item.label = next_panel_label(len([p for p in self.items if isinstance(p, PanelItem) and p.include_in_labels]))
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Selected label visibility updated")

    def apply_selected_properties(self):
        if not self.selected_ids:
            return
        font_family = self.sel_font_family_var.get()
        align = self.sel_align_var.get()
        try:
            rotation_float = float(self.sel_rotation_var.get())
            rotation_text = int(rotation_float) % 360
            x = int(float(self.sel_x_var.get()))
            y = int(float(self.sel_y_var.get()))
            w_raw = int(float(self.sel_w_var.get()))
            h_raw = int(float(self.sel_h_var.get()))
            w = max(MIN_PANEL_SIZE, abs(w_raw))
            h = max(MIN_PANEL_SIZE, abs(h_raw))
            font_size = max(1, int(float(self.sel_font_var.get())))
            line_spacing = max(0.8, float(self.sel_line_spacing_var.get()))
        except ValueError:
            messagebox.showerror("Invalid input", "Position, size, font size, rotation, and line spacing must be numeric.")
            return
        self.save_undo_state()
        if len(self.selected_ids) == 1:
            sel = self.get_item_by_id(self.selected_ids[0])
            if sel is None:
                return
            if isinstance(sel, PanelItem) and self.selected_label_panel_id == sel.id:
                sel.label = self.sel_label_var.get()
                sel.label_font_size = font_size
                sel.label_font_family = font_family
                sel.show_label = self.sel_show_label_var.get()
                sel.include_in_labels = self.sel_include_label_var.get()
                sel.label_offset_x = x - sel.x
                sel.label_offset_y = y - sel.y
            elif isinstance(sel, PanelItem):
                sel.x, sel.y, sel.w, sel.h = x, y, w, h
                sel.label = self.sel_label_var.get()
                sel.label_font_size = font_size
                sel.label_font_family = font_family
                sel.show_label = self.sel_show_label_var.get()
                sel.include_in_labels = self.sel_include_label_var.get()
                sel._preview_cache_key = None
                sel._preview_cache_photo = None
            elif isinstance(sel, ShapeItem):
                sel.x, sel.y = x, y
                sel.w = w_raw if sel.shape_type in ("line", "arrow") or sel.is_stat_compare() else w
                sel.h = h_raw if sel.shape_type in ("line", "arrow") or sel.is_stat_compare() else h
                sel.width = font_size
                sel.rotation = rotation_float % 360
                sel.text = self.sel_text_var.get()
                sel.font_family = font_family
                if sel.is_stat_compare():
                    sel.font_size = max(6, sel.font_size)
            else:
                old_rotation = sel.rotation
                sel.x, sel.y = x, y
                if rotation_text not in (0, 90, 180, 270):
                    rotation_text = min((0, 90, 180, 270), key=lambda v: abs(v - rotation_text))
                self.rotate_text_item_box(sel, rotation_text)
                if old_rotation != rotation_text and ((old_rotation in (90, 270)) != (rotation_text in (90, 270))):
                    sel.w, sel.h = h, w
                else:
                    sel.w, sel.h = w, h
                old_text = sel.text
                new_text = self.sel_text_var.get()
                sel.font_size = font_size
                sel.font_family = font_family
                sel.align = align
                sel.bold = self.sel_bold_var.get()
                sel.italic = self.sel_italic_var.get()
                sel.line_spacing = line_spacing
                if new_text != old_text:
                    sel.text = new_text
                    sel.rich_runs = make_plain_runs(new_text, sel.bold, sel.italic)
                else:
                    sel.rich_runs = normalize_runs([{"text": r["text"], "bold": r.get("bold", False), "italic": r.get("italic", False)} for r in sel.rich_runs])
                sel.sync_text_from_runs()
                sel._preview_cache_key = None
                sel._preview_cache_photo = None
                if sel.text.strip().lower().startswith("caption"):
                    self.fit_text_box_height(sel, maybe_expand_canvas=True)
        else:
            anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else None
            if anchor:
                anchor.x, anchor.y = x, y
                anchor.w, anchor.h = w, h
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Properties applied")

    def open_rich_text_editor(self):
        if len(self.selected_ids) != 1:
            return
        item = self.get_item_by_id(self.selected_ids[0])
        if not isinstance(item, TextItem):
            return
        def on_save(payload):
            self.save_undo_state()
            old_rotation = item.rotation
            item.rich_runs = payload["rich_runs"]
            item.font_family = payload["font_family"]
            item.font_size = payload["font_size"]
            item.align = payload["align"]
            self.rotate_text_item_box(item, payload["rotation"])
            if old_rotation != payload["rotation"] and ((old_rotation in (90, 270)) != (payload["rotation"] in (90, 270))):
                item.w, item.h = item.h, item.w
            item.line_spacing = payload["line_spacing"]
            item.sync_text_from_runs()
            if len(item.rich_runs) == 1:
                item.bold = bool(item.rich_runs[0].get("bold", False))
                item.italic = bool(item.rich_runs[0].get("italic", False))
            item._preview_cache_key = None
            item._preview_cache_photo = None
            if item.text.strip().lower().startswith("caption"):
                self.fit_text_box_height(item, maybe_expand_canvas=True)
            self.refresh_selected_panel()
            self.redraw()
            self.set_status("Rich text updated")
        RichTextEditor(self.root, item, on_save)

    def pick_selected_text_colour(self):
        if len(self.selected_ids) != 1:
            return
        item = self.get_item_by_id(self.selected_ids[0])
        if item is None:
            return
        initial = item.fill if isinstance(item, TextItem) else item.color if isinstance(item, ShapeItem) else item.label_color if isinstance(item, PanelItem) else "black"
        color = colorchooser.askcolor(color=initial, title="Choose colour")[1]
        if color:
            self.save_undo_state()
            if isinstance(item, TextItem):
                item.fill = color
                item._preview_cache_key = None
                item._preview_cache_photo = None
            elif isinstance(item, ShapeItem):
                item.color = color
                if item.shape_type in ("highlight_bar", "arrow_head"):
                    item.fill = color
            elif isinstance(item, PanelItem):
                item.label_color = color
            self.refresh_selected_panel()
            self.redraw()

    def pick_stat_color(self):
        color = colorchooser.askcolor(color=self.stat_color_var.get(), title="Statistic colour")[1]
        if color:
            self.stat_color_var.set(color)
            self.stat_color_swatch.configure(bg=color, activebackground=color)

    def bring_to_front(self):
        selected = self.get_selected_items()
        if not selected:
            return
        self.save_undo_state()
        max_z = max((i.z_index for i in self.items), default=0)
        for idx, sel in enumerate(selected):
            sel.z_index = max_z + idx + 1
        self.redraw()

    def send_to_back(self):
        selected = self.get_selected_items()
        if not selected:
            return
        self.save_undo_state()
        min_z = min((i.z_index for i in self.items), default=0)
        for idx, sel in enumerate(selected):
            sel.z_index = min_z - len(selected) + idx
        self.redraw()

    def duplicate_selected(self):
        selected = self.get_selected_items()
        if not selected:
            return
        self.save_undo_state()
        new_ids = []
        max_z = max((i.z_index for i in self.items), default=0)
        for idx, sel in enumerate(selected):
            if isinstance(sel, PanelItem):
                dup = PanelItem(
                    id=self.get_next_id(),
                    source_path=sel.source_path,
                    pil_image=sel.pil_image.copy() if sel.pil_image else None,
                    original_size=sel.original_size,
                    x=sel.x + 20,
                    y=sel.y + 20,
                    w=sel.w,
                    h=sel.h,
                    label=sel.label,
                    show_label=sel.show_label,
                    include_in_labels=sel.include_in_labels,
                    label_font_size=sel.label_font_size,
                    label_font_family=sel.label_font_family,
                    label_color=sel.label_color,
                    label_offset_x=sel.label_offset_x,
                    label_offset_y=sel.label_offset_y,
                    border_width=sel.border_width,
                    border_color=sel.border_color,
                    z_index=max_z + idx + 1,
                    group_id=sel.group_id,
                )
            elif isinstance(sel, ShapeItem):
                dup = ShapeItem(
                    id=self.get_next_id(),
                    shape_type=sel.shape_type,
                    x=sel.x + 20,
                    y=sel.y + 20,
                    w=sel.w,
                    h=sel.h,
                    color=sel.color,
                    width=sel.width,
                    fill=sel.fill,
                    fill_alpha=sel.fill_alpha,
                    rotation=sel.rotation,
                    text=sel.text,
                    font_size=sel.font_size,
                    font_family=sel.font_family,
                    z_index=max_z + idx + 1,
                    group_id=sel.group_id,
                    compare_style=sel.compare_style,
                    bracket_height=sel.bracket_height,
                    text_gap=sel.text_gap,
                    auto_center_text=sel.auto_center_text,
                    snap_to_tops=sel.snap_to_tops,
                )
            else:
                dup = TextItem(
                    id=self.get_next_id(),
                    text=sel.text,
                    x=sel.x + 20,
                    y=sel.y + 20,
                    w=sel.w,
                    h=sel.h,
                    font_size=sel.font_size,
                    font_family=sel.font_family,
                    bold=sel.bold,
                    italic=sel.italic,
                    rich_runs=json.loads(json.dumps(sel.rich_runs)),
                    fill=sel.fill,
                    align=sel.align,
                    background=sel.background,
                    outline=sel.outline,
                    rotation=sel.rotation,
                    line_spacing=sel.line_spacing,
                    z_index=max_z + idx + 1,
                    group_id=sel.group_id,
                )
            self.items.append(dup)
            new_ids.append(dup.id)
        self.selected_ids = new_ids
        self.anchor_selected_id = new_ids[-1] if new_ids else None
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()

    def delete_selected(self):
        if not self.selected_ids:
            return
        self.save_undo_state()
        selected_set = set(self.selected_ids)
        self.items = [i for i in self.items if i.id not in selected_set]
        self.selected_ids = []
        self.anchor_selected_id = None
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Deleted selected item(s)")

    def included_panels_for_labels(self):
        panels = [i for i in self.items if isinstance(i, PanelItem) and i.include_in_labels]
        if self.keep_label_order.get():
            return sorted(panels, key=lambda p: p.id)
        return sorted(panels, key=lambda p: (p.y, p.x, p.id))

    def regenerate_labels(self, silent=False):
        panels = self.included_panels_for_labels()
        self.save_undo_state()
        for p in [i for i in self.items if isinstance(i, PanelItem) and not i.include_in_labels]:
            p.label = ""
            p.show_label = False
        for idx, panel in enumerate(panels):
            panel.label = next_panel_label(idx)
            panel.show_label = True
        self.redraw()
        if not silent:
            order = "add order" if self.keep_label_order.get() else "spatial order"
            self.set_status(f"Regenerated panel labels by {order}")

    def toggle_labels(self):
        panels = [i for i in self.items if isinstance(i, PanelItem) and i.include_in_labels]
        if not panels:
            return
        self.save_undo_state()
        turn_on = any(not p.show_label for p in panels)
        for p in panels:
            p.show_label = turn_on
        self.refresh_selected_panel()
        self.redraw()

    def apply_default_label_size(self):
        self.save_undo_state()
        size = self.default_label_size.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                item.label_font_size = size
        self.refresh_selected_panel()
        self.redraw()

    def apply_default_label_offsets(self):
        self.save_undo_state()
        ox = self.default_label_offset_x.get()
        oy = self.default_label_offset_y.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                item.label_offset_x = ox
                item.label_offset_y = oy
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Applied label offsets to all panels")

    def apply_default_label_offsets_selected(self):
        if not self.selected_ids:
            return
        self.save_undo_state()
        ox = self.default_label_offset_x.get()
        oy = self.default_label_offset_y.get()
        for item in self.get_selected_items():
            if isinstance(item, PanelItem):
                item.label_offset_x = ox
                item.label_offset_y = oy
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Applied label offsets to selected panels")

    def exclude_selected_from_labels(self):
        selected = [i for i in self.get_selected_items() if isinstance(i, PanelItem)]
        if not selected:
            return
        self.save_undo_state()
        for p in selected:
            p.include_in_labels = False
            p.show_label = False
            p.label = ""
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Selected panels excluded from auto labels")

    def include_selected_in_labels(self):
        selected = [i for i in self.get_selected_items() if isinstance(i, PanelItem)]
        if not selected:
            return
        self.save_undo_state()
        used = len([i for i in self.items if isinstance(i, PanelItem) and i.include_in_labels and i.id not in {p.id for p in selected}])
        for p in selected:
            p.include_in_labels = True
            p.show_label = True
            if not p.label:
                p.label = next_panel_label(used)
                used += 1
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Selected panels included in auto labels")

    def toggle_label_sequence_mode(self):
        self.label_sequence_mode = not self.label_sequence_mode
        self.label_align_mode = False
        self.stat_mode = None
        self.erase_mode = False
        self.crop_mode = False
        self.label_sequence_ids = []
        if self.label_sequence_mode:
            if messagebox.askyesno("Click label order", "Clear current included labels and then click panels in the order you want?"):
                self.save_undo_state()
                for p in [i for i in self.items if isinstance(i, PanelItem)]:
                    p.label = ""
                    p.show_label = False
                    p.include_in_labels = False
            self.set_status("Click panels in label order. Press Esc or the button again to finish.")
        else:
            self.set_status("Click label order finished")
        self.update_mode_button_styles()
        self.redraw()

    def handle_label_sequence_click(self, panel: PanelItem):
        if panel.id in self.label_sequence_ids:
            self.set_status("Panel already labeled in this sequence")
            return
        self.save_undo_state()
        panel.include_in_labels = True
        panel.show_label = True
        panel.label = next_panel_label(len(self.label_sequence_ids))
        self.label_sequence_ids.append(panel.id)
        self.selected_ids = [panel.id]
        self.anchor_selected_id = panel.id
        self.selected_label_panel_id = panel.id
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Assigned label {panel.label}. Continue clicking panels.")

    def toggle_label_align_mode(self):
        self.label_align_mode = not self.label_align_mode
        self.label_sequence_mode = False
        self.stat_mode = None
        self.erase_mode = False
        self.crop_mode = False
        self.label_anchor_panel_id = None
        if self.label_align_mode:
            self.set_status(f"Click source label, then target labels. Axis: {self.label_align_axis_var.get()}.")
        else:
            self.set_status("Label alignment by click off")
        self.update_mode_button_styles()
        self.redraw()

    def apply_label_alignment_to_panel(self, panel: PanelItem, anchor: PanelItem, axis: str):
        ax, ay = self.label_absolute_position(anchor)
        px, py = self.label_absolute_position(panel)
        if axis in ("X", "Both"):
            panel.label_offset_x = int(round(ax - panel.x))
        if axis in ("Y", "Both"):
            panel.label_offset_y = int(round(ay - panel.y))
        if axis not in ("X", "Y", "Both"):
            panel.label_offset_x = int(round(ax - panel.x))
            panel.label_offset_y = int(round(ay - panel.y))
        panel.include_in_labels = True
        panel.show_label = True

    def handle_label_align_click(self, panel: PanelItem):
        if self.label_anchor_panel_id is None:
            self.label_anchor_panel_id = panel.id
            self.selected_ids = [panel.id]
            self.anchor_selected_id = panel.id
            self.selected_label_panel_id = panel.id
            self.refresh_selected_panel()
            self.redraw()
            self.set_status(f"Anchor label selected. Click target label to align {self.label_align_axis_var.get()}.")
            return
        anchor = self.get_item_by_id(self.label_anchor_panel_id)
        if not isinstance(anchor, PanelItem):
            self.label_anchor_panel_id = None
            return
        self.save_undo_state()
        self.apply_label_alignment_to_panel(panel, anchor, self.label_align_axis_var.get())
        self.selected_ids = [panel.id]
        self.anchor_selected_id = anchor.id
        self.selected_label_panel_id = panel.id
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Aligned label {self.label_align_axis_var.get()} to anchor")

    def align_selected_labels(self, axis: str):
        panels = [i for i in self.get_selected_items() if isinstance(i, PanelItem)]
        if len(panels) < 2:
            messagebox.showinfo("Align labels", "Select at least two panels or labels. The orange/anchor selection is used as the source.")
            return
        anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else panels[0]
        if not isinstance(anchor, PanelItem):
            return
        self.save_undo_state()
        for panel in panels:
            if panel.id == anchor.id:
                continue
            self.apply_label_alignment_to_panel(panel, anchor, axis)
        self.selected_label_panel_id = anchor.id
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Aligned selected labels: {axis}")

    def apply_auto_label_position_all(self):
        self.save_undo_state()
        pos = self.label_pos_var.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                self._apply_auto_label_position(item, pos)
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Auto label position '{pos}' applied to all panels")

    def apply_auto_label_position_selected(self):
        if not self.selected_ids:
            return
        self.save_undo_state()
        pos = self.label_pos_var.get()
        for item in self.get_selected_items():
            if isinstance(item, PanelItem):
                self._apply_auto_label_position(item, pos)
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Auto label position '{pos}' applied to selected panels")

    def _apply_auto_label_position(self, panel: PanelItem, position: str):
        pad = 6
        if position == "top-in":
            panel.label_offset_x = pad
            panel.label_offset_y = pad
        elif position == "top-out":
            panel.label_offset_x = pad
            panel.label_offset_y = -30
        elif position == "left-in":
            panel.label_offset_x = pad
            panel.label_offset_y = panel.h // 2
        elif position == "left-out":
            panel.label_offset_x = -30
            panel.label_offset_y = panel.h // 2
        elif position == "bottom-in":
            panel.label_offset_x = pad
            panel.label_offset_y = panel.h - 30
        elif position == "bottom-out":
            panel.label_offset_x = pad
            panel.label_offset_y = panel.h + 5
        elif position == "center":
            panel.label_offset_x = panel.w // 2
            panel.label_offset_y = panel.h // 2
        panel._preview_cache_key = None
        panel._preview_cache_photo = None

    def auto_grid_if_reasonable(self):
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        if 2 <= len(panels) <= 20:
            self.auto_grid()

    def get_group_members(self, group_id: int) -> List[CanvasItem]:
        return [i for i in self.items if i.group_id == group_id]

    def get_selection_units(self) -> List[List[CanvasItem]]:
        selected = self.get_selected_items()
        if not selected:
            return []
        units = []
        used = set()
        for item in selected:
            if item.id in used:
                continue
            if item.group_id is not None:
                members = [m for m in self.items if m.group_id == item.group_id]
                if any(m.id in self.selected_ids for m in members):
                    units.append(members)
                    used.update(m.id for m in members)
            else:
                units.append([item])
                used.add(item.id)
        return units

    def unit_bbox(self, unit: List[CanvasItem]):
        x1 = min(i.bbox()[0] for i in unit)
        y1 = min(i.bbox()[1] for i in unit)
        x2 = max(i.bbox()[2] for i in unit)
        y2 = max(i.bbox()[3] for i in unit)
        return x1, y1, x2, y2

    def move_unit_to(self, unit: List[CanvasItem], new_x: Optional[int] = None, new_y: Optional[int] = None):
        x1, y1, _, _ = self.unit_bbox(unit)
        dx = 0 if new_x is None else new_x - x1
        dy = 0 if new_y is None else new_y - y1
        for item in unit:
            item.x += dx
            item.y += dy

    def group_selected(self):
        if len(self.selected_ids) < 2:
            return
        self.save_undo_state()
        gid = self.get_next_group_id()
        for item in self.get_selected_items():
            item.group_id = gid
        self.selected_label_panel_id = None
        self.redraw()
        self.set_status("Grouped selected items")

    def ungroup_selected(self):
        selected = self.get_selected_items()
        if not selected:
            return
        self.save_undo_state()
        for item in selected:
            item.group_id = None
        self.redraw()
        self.set_status("Ungrouped selected items")

    def apply_custom_grid(self):
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        if not panels:
            return
        try:
            rows = max(1, int(self.grid_rows_entry.get()))
            cols = max(1, int(self.grid_cols_entry.get()))
        except ValueError:
            messagebox.showerror("Invalid grid", "Rows and columns must be integers.")
            return
        self.save_undo_state()
        self.grid_layout(panels, rows, cols)
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Applied {rows} × {cols} grid")

    def auto_grid(self):
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        if not panels:
            return
        self.save_undo_state()
        n = len(panels)
        cols = math.ceil(math.sqrt(n))
        rows = math.ceil(n / cols)
        self.grid_layout(panels, rows, cols)
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Auto-arranged {n} panel(s)")

    def grid_layout(self, panels: List[PanelItem], rows: int, cols: int):
        margin = self.outer_margin.get()
        gap_x = self.gap_x.get()
        gap_y = self.gap_y.get()
        avail_w = self.canvas_w - 2 * margin - (cols - 1) * gap_x
        avail_h = self.canvas_h - 2 * margin - (rows - 1) * gap_y
        cell_w = max(20, avail_w // cols)
        cell_h = max(20, avail_h // rows)
        panels_sorted = sorted(panels, key=lambda p: p.id if self.keep_label_order.get() else (p.y, p.x, p.id))
        for idx, p in enumerate(panels_sorted):
            r = idx // cols
            c = idx % cols
            x = margin + c * (cell_w + gap_x)
            y = margin + r * (cell_h + gap_y)
            ow, oh = p.original_size
            new_w, new_h = fit_size_keep_aspect(ow, oh, cell_w, cell_h)
            p.x = x + (cell_w - new_w) // 2
            p.y = y + (cell_h - new_h) // 2
            p.w = new_w
            p.h = new_h
            p._preview_cache_key = None
            p._preview_cache_photo = None

    def distribute_h(self):
        units = self.get_selection_units()
        if len(units) < 2:
            return
        self.save_undo_state()
        units = sorted(units, key=lambda u: self.unit_bbox(u)[0])
        left = min(self.unit_bbox(u)[0] for u in units)
        gap = self.gap_x.get()
        cur_x = left
        for u in units:
            self.move_unit_to(u, new_x=cur_x)
            ux1, _, ux2, _ = self.unit_bbox(u)
            cur_x = ux2 + gap
        self.refresh_selected_panel()
        self.redraw()

    def distribute_v(self):
        units = self.get_selection_units()
        if len(units) < 2:
            return
        self.save_undo_state()
        units = sorted(units, key=lambda u: self.unit_bbox(u)[1])
        top = min(self.unit_bbox(u)[1] for u in units)
        gap = self.gap_y.get()
        cur_y = top
        for u in units:
            self.move_unit_to(u, new_y=cur_y)
            _, _, _, uy2 = self.unit_bbox(u)
            cur_y = uy2 + gap
        self.refresh_selected_panel()
        self.redraw()

    def align_to_anchor(self, where: str):
        units = self.get_selection_units()
        if len(units) < 2:
            return
        anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else None
        if anchor is None:
            return
        self.save_undo_state()
        anchor_unit = None
        for u in units:
            if any(i.id == anchor.id for i in u):
                anchor_unit = u
                break
        if anchor_unit is None:
            return
        ax1, ay1, ax2, ay2 = self.unit_bbox(anchor_unit)
        acx = (ax1 + ax2) / 2.0
        acy = (ay1 + ay2) / 2.0
        for u in units:
            if u == anchor_unit:
                continue
            ux1, uy1, ux2, uy2 = self.unit_bbox(u)
            uw = ux2 - ux1
            uh = uy2 - uy1
            if where == "left":
                self.move_unit_to(u, new_x=ax1)
            elif where == "right":
                self.move_unit_to(u, new_x=ax2 - uw)
            elif where == "top":
                self.move_unit_to(u, new_y=ay1)
            elif where == "bottom":
                self.move_unit_to(u, new_y=ay2 - uh)
            elif where == "vcenter":
                self.move_unit_to(u, new_y=int(round(acy - uh / 2)))
            elif where == "hcenter":
                self.move_unit_to(u, new_x=int(round(acx - uw / 2)))
        self.refresh_selected_panel()
        self.redraw()

    def same_widths(self):
        selected = self.get_selected_items()
        if len(selected) < 2:
            return
        anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else selected[0]
        if anchor is None:
            return
        self.save_undo_state()
        w = anchor.w
        for item in selected:
            if item.id == anchor.id:
                continue
            if isinstance(item, PanelItem):
                ow, oh = item.original_size
                new_h = int(round(w * oh / ow)) if ow else item.h
                item.w = max(MIN_PANEL_SIZE, w)
                item.h = max(MIN_PANEL_SIZE, new_h)
                item._preview_cache_key = None
                item._preview_cache_photo = None
            else:
                item.w = max(8, w)
                if isinstance(item, TextItem):
                    item._preview_cache_key = None
                    item._preview_cache_photo = None
        self.refresh_selected_panel()
        self.redraw()

    def same_heights(self):
        selected = self.get_selected_items()
        if len(selected) < 2:
            return
        anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else selected[0]
        if anchor is None:
            return
        self.save_undo_state()
        h = anchor.h
        for item in selected:
            if item.id == anchor.id:
                continue
            if isinstance(item, PanelItem):
                ow, oh = item.original_size
                new_w = int(round(h * ow / oh)) if oh else item.w
                item.w = max(MIN_PANEL_SIZE, new_w)
                item.h = max(MIN_PANEL_SIZE, h)
                item._preview_cache_key = None
                item._preview_cache_photo = None
            else:
                item.h = max(8, h)
                if isinstance(item, TextItem):
                    item._preview_cache_key = None
                    item._preview_cache_photo = None
        self.refresh_selected_panel()
        self.redraw()

    def content_bbox(self) -> Optional[Tuple[int, int, int, int]]:
        if not self.items:
            return None
        x1 = min(i.bbox()[0] for i in self.items)
        y1 = min(i.bbox()[1] for i in self.items)
        x2 = max(i.bbox()[2] for i in self.items)
        y2 = max(i.bbox()[3] for i in self.items)
        return x1, y1, x2, y2

    def crop_canvas_to_content(self):
        bbox = self.content_bbox()
        if bbox is None:
            return
        self.save_undo_state()
        margin = self.outer_margin.get()
        x1, y1, x2, y2 = bbox
        shift_x = margin - x1
        shift_y = margin - y1
        for item in self.items:
            item.x += shift_x
            item.y += shift_y
        self.canvas_w = max(100, x2 - x1 + 2 * margin)
        self.canvas_h = max(100, y2 - y1 + 2 * margin)
        self.canvas_w_entry.delete(0, "end")
        self.canvas_w_entry.insert(0, str(self.canvas_w))
        self.canvas_h_entry.delete(0, "end")
        self.canvas_h_entry.insert(0, str(self.canvas_h))
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Canvas cropped to content")

    def toggle_crop_mode(self):
        self.crop_mode = not self.crop_mode
        if self.crop_mode:
            self.erase_mode = False
            self.stat_mode = None
            self.label_align_mode = False
            self.label_sequence_mode = False
            self.set_status("Crop mode ON. Drag a rectangle over an image panel.")
        else:
            self.crop_start = None
            self.crop_current = None
            self.crop_target_id = None
            self.set_status("Crop mode OFF")
        self.update_mode_button_styles()
        self.redraw()

    def toggle_erase_mode(self):
        self.erase_mode = not self.erase_mode
        if self.erase_mode:
            self.crop_mode = False
            self.stat_mode = None
            self.label_align_mode = False
            self.label_sequence_mode = False
            self.set_status("Erase rectangle mode ON")
        else:
            self.erase_start = None
            self.erase_current = None
            self.set_status("Erase rectangle mode OFF")
        self.update_mode_button_styles()
        self.redraw()

    def apply_crop_rectangle(self, x1, y1, x2, y2, target):
        if not isinstance(target, PanelItem) or target.pil_image is None:
            return
        if abs(x2 - x1) < 4 or abs(y2 - y1) < 4:
            return
        ix1, iy1, ix2, iy2 = target.bbox()
        rx1 = clamp(min(x1, x2), ix1, ix2)
        ry1 = clamp(min(y1, y2), iy1, iy2)
        rx2 = clamp(max(x1, x2), ix1, ix2)
        ry2 = clamp(max(y1, y2), iy1, iy2)
        if rx2 <= rx1 or ry2 <= ry1:
            return
        src_w, src_h = target.pil_image.size
        sx1 = int(round((rx1 - target.x) * src_w / target.w))
        sy1 = int(round((ry1 - target.y) * src_h / target.h))
        sx2 = int(round((rx2 - target.x) * src_w / target.w))
        sy2 = int(round((ry2 - target.y) * src_h / target.h))
        sx1 = clamp(sx1, 0, src_w - 1)
        sy1 = clamp(sy1, 0, src_h - 1)
        sx2 = clamp(sx2, sx1 + 1, src_w)
        sy2 = clamp(sy2, sy1 + 1, src_h)
        old_label_abs = (target.x + target.label_offset_x, target.y + target.label_offset_y)
        self.save_undo_state()
        target.pil_image = target.pil_image.crop((sx1, sy1, sx2, sy2)).convert("RGBA")
        target.original_size = target.pil_image.size
        target.x = int(round(rx1))
        target.y = int(round(ry1))
        target.w = max(MIN_PANEL_SIZE, int(round(rx2 - rx1)))
        target.h = max(MIN_PANEL_SIZE, int(round(ry2 - ry1)))
        target.label_offset_x = int(round(old_label_abs[0] - target.x))
        target.label_offset_y = int(round(old_label_abs[1] - target.y))
        target._preview_cache_key = None
        target._preview_cache_photo = None
        self.selected_ids = [target.id]
        self.anchor_selected_id = target.id
        self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.set_status("Image cropped")

    def trim_selected_image(self):
        if len(self.selected_ids) != 1:
            messagebox.showinfo("Trim image", "Select one image panel first.")
            return
        item = self.get_item_by_id(self.selected_ids[0])
        if not isinstance(item, PanelItem) or item.pil_image is None:
            messagebox.showinfo("Trim image", "Select one image panel first.")
            return
        bbox = trim_bbox_from_image(item.pil_image)
        if bbox is None:
            messagebox.showinfo("Trim image", "No visible content found to trim.")
            return
        sx1, sy1, sx2, sy2 = bbox
        src_w, src_h = item.pil_image.size
        if sx1 <= 0 and sy1 <= 0 and sx2 >= src_w and sy2 >= src_h:
            messagebox.showinfo("Trim image", "This image already appears tightly trimmed.")
            return
        old_label_abs = (item.x + item.label_offset_x, item.y + item.label_offset_y)
        new_x = item.x + (sx1 / src_w) * item.w
        new_y = item.y + (sy1 / src_h) * item.h
        new_w = ((sx2 - sx1) / src_w) * item.w
        new_h = ((sy2 - sy1) / src_h) * item.h
        self.save_undo_state()
        item.pil_image = item.pil_image.crop((sx1, sy1, sx2, sy2)).convert("RGBA")
        item.original_size = item.pil_image.size
        item.x = int(round(new_x))
        item.y = int(round(new_y))
        item.w = max(MIN_PANEL_SIZE, int(round(new_w)))
        item.h = max(MIN_PANEL_SIZE, int(round(new_h)))
        item.label_offset_x = int(round(old_label_abs[0] - item.x))
        item.label_offset_y = int(round(old_label_abs[1] - item.y))
        item._preview_cache_key = None
        item._preview_cache_photo = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Selected image trimmed")

    def apply_erase_rectangle(self, x1, y1, x2, y2):
        if abs(x2 - x1) < 2 or abs(y2 - y1) < 2:
            return
        rect = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        target = None
        for item in sorted(self.items, key=lambda z: z.z_index, reverse=True):
            if isinstance(item, PanelItem) and rects_overlap(rect, item.bbox()):
                target = item
                break
        if target is None or target.pil_image is None:
            return
        ix1, iy1, ix2, iy2 = target.bbox()
        rx1 = clamp(rect[0], ix1, ix2)
        ry1 = clamp(rect[1], iy1, iy2)
        rx2 = clamp(rect[2], ix1, ix2)
        ry2 = clamp(rect[3], iy1, iy2)
        if rx2 <= rx1 or ry2 <= ry1:
            return
        src_w, src_h = target.pil_image.size
        sx1 = int(round((rx1 - target.x) * src_w / target.w))
        sy1 = int(round((ry1 - target.y) * src_h / target.h))
        sx2 = int(round((rx2 - target.x) * src_w / target.w))
        sy2 = int(round((ry2 - target.y) * src_h / target.h))
        sx1 = clamp(sx1, 0, src_w - 1)
        sy1 = clamp(sy1, 0, src_h - 1)
        sx2 = clamp(sx2, sx1 + 1, src_w)
        sy2 = clamp(sy2, sy1 + 1, src_h)
        self.save_undo_state()
        draw = ImageDraw.Draw(target.pil_image)
        draw.rectangle([sx1, sy1, sx2, sy2], fill="white")
        target._preview_cache_key = None
        target._preview_cache_photo = None
        self.set_status("Erased selected rectangle from panel")

    def resolve_all_overlaps(self):
        panels = sorted([i for i in self.items if isinstance(i, PanelItem)], key=lambda p: (p.y, p.x))
        if not panels:
            return
        self.save_undo_state()
        changed = True
        rounds = 0
        while changed and rounds < 60:
            changed = False
            rounds += 1
            for i in range(len(panels)):
                for j in range(i + 1, len(panels)):
                    a = panels[i]
                    b = panels[j]
                    if rects_overlap(a.bbox(), b.bbox(), pad=0):
                        try_x = a.x + a.w + max(self.gap_x.get(), SAFE_GAP)
                        try_y = a.y + a.h + max(self.gap_y.get(), SAFE_GAP)
                        if try_x + b.w <= self.canvas_w - self.outer_margin.get():
                            b.x = try_x
                            b.y = a.y
                        else:
                            b.x = a.x
                            b.y = try_y
                        b._preview_cache_key = None
                        b._preview_cache_photo = None
                        changed = True
        self.refresh_selected_panel()
        self.redraw()

    def fit_text_box_height(self, item: TextItem, maybe_expand_canvas=False):
        img = Image.new("RGBA", (max(100, item.w), 3000), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        lines = self.wrap_runs_to_width(draw=draw, runs=item.rich_runs, font_family=item.font_family, font_size=item.font_size, max_width=max(20, item.w - 12))
        y = 6
        extra_factor = max(0.0, item.line_spacing - 1.0)
        for line in lines:
            line_h = max((seg["height"] for seg in line), default=item.font_size)
            y += line_h + 4 + int(round(line_h * extra_factor))
        item.h = max(item.h, y + 8)
        item._preview_cache_key = None
        item._preview_cache_photo = None
        if maybe_expand_canvas:
            needed_h = item.y + item.h + self.outer_margin.get()
            if needed_h > self.canvas_h:
                self.canvas_h = needed_h
                self.canvas_h_entry.delete(0, "end")
                self.canvas_h_entry.insert(0, str(self.canvas_h))

    def set_stat_mode(self, mode: str):
        self.stat_mode = None if self.stat_mode == mode else mode
        self.stat_first_point = None
        if self.stat_mode:
            self.erase_mode = False
            self.crop_mode = False
            self.label_align_mode = False
            self.label_sequence_mode = False
            if self.stat_mode == "symbol":
                self.set_status("Statistic symbol mode ON. Click where you want the centre of the symbol.")
            else:
                self.set_status("Bracket compare mode ON. Click the first data point, then the second data point.")
        else:
            self.set_status("Statistic mode OFF")
        self.update_mode_button_styles()
        self.redraw()

    def stat_text_value(self):
        sym = self.stat_symbol_var.get() or "*"
        try:
            count = max(1, int(self.stat_count_var.get()))
        except Exception:
            count = 1
        if len(sym) == 1:
            return sym * count
        return sym if count == 1 else sym * count

    def add_stat_symbol_at(self, x: int, y: int):
        text = self.stat_text_value()
        try:
            size = max(6, int(self.stat_text_size_var.get()))
        except Exception:
            size = DEFAULT_STAT_TEXT_SIZE
        width = max(60, int(size * max(2.0, len(text) * 0.9)))
        height = max(40, int(size * 1.7))
        self.save_undo_state()
        item = TextItem(
            id=self.get_next_id(),
            text=text,
            rich_runs=make_plain_runs(text, bold=True),
            x=int(round(x - width / 2)),
            y=int(round(y - height / 2)),
            w=width,
            h=height,
            font_size=size,
            font_family=self.default_font_family.get(),
            bold=True,
            fill=self.stat_color_var.get(),
            align="center",
            line_spacing=DEFAULT_LINE_SPACING,
            z_index=max((i.z_index for i in self.items), default=-1) + 1,
        )
        self.items.append(item)
        self.selected_ids = [item.id]
        self.anchor_selected_id = item.id
        self.selected_label_panel_id = None
        self.stat_mode = None
        self.update_mode_button_styles()
        self.refresh_selected_panel()
        self.redraw()
        self.set_status(f"Added statistic symbol {text}")

    def current_stat_compare_style(self):
        return "line" if self.stat_compare_style_var.get().lower().startswith("straight") else "bracket"

    def handle_stat_compare_click(self, x: int, y: int):
        if self.stat_first_point is None:
            self.stat_first_point = (x, y)
            self.redraw()
            self.set_status("First point set. Click the second comparison point.")
            return
        x1, y1 = self.stat_first_point
        if abs(x - x1) < 4:
            self.set_status("Second point is too close to the first point")
            return
        text = self.stat_text_value()
        try:
            text_size = max(6, int(self.stat_text_size_var.get()))
        except Exception:
            text_size = DEFAULT_STAT_TEXT_SIZE
        try:
            bracket_height = max(0, int(self.stat_bracket_height_var.get()))
        except Exception:
            bracket_height = DEFAULT_STAT_BRACKET_HEIGHT
        try:
            text_gap = max(0, int(self.stat_text_gap_var.get()))
        except Exception:
            text_gap = DEFAULT_STAT_TEXT_GAP
        self.save_undo_state()
        item = ShapeItem(
            id=self.get_next_id(),
            shape_type="stat_compare",
            x=x1,
            y=y1,
            w=x - x1,
            h=y - y1,
            color=self.stat_color_var.get(),
            width=2,
            text=text,
            font_size=text_size,
            font_family=self.default_font_family.get(),
            z_index=max((i.z_index for i in self.items), default=-1) + 1,
            compare_style=self.current_stat_compare_style(),
            bracket_height=bracket_height,
            text_gap=text_gap,
            auto_center_text=self.stat_auto_center_var.get(),
            snap_to_tops=self.stat_snap_tops_var.get(),
        )
        self.items.append(item)
        self.selected_ids = [item.id]
        self.anchor_selected_id = item.id
        self.selected_label_panel_id = None
        self.stat_first_point = None
        self.stat_mode = None
        self.update_mode_button_styles()
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Added statistic comparison")

    def clear_last_stat_compare(self):
        candidates = [i for i in self.items if isinstance(i, ShapeItem) and i.is_stat_compare()]
        if not candidates:
            return
        last = max(candidates, key=lambda i: i.z_index)
        self.save_undo_state()
        self.items = [i for i in self.items if i.id != last.id]
        if last.id in self.selected_ids:
            self.selected_ids = []
            self.anchor_selected_id = None
            self.selected_label_panel_id = None
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Cleared last statistic bracket")

    def render_shape_on_image(self, draw: ImageDraw.ImageDraw, item: ShapeItem, scale: float, out: Image.Image):
        stroke = max(1, int(round(item.width * scale)))
        color = item.color
        if item.shape_type in ("rectangle", "highlight_bar"):
            pts = [(int(round(x * scale)), int(round(y * scale))) for x, y in item.visual_points()]
            if item.fill:
                if item.fill_alpha < 255:
                    overlay = Image.new("RGBA", out.size, (255, 255, 255, 0))
                    od = ImageDraw.Draw(overlay)
                    rgb = color_to_rgb_tuple(item.fill)
                    od.polygon(pts, fill=(rgb[0], rgb[1], rgb[2], item.fill_alpha))
                    out.alpha_composite(overlay)
                else:
                    draw.polygon(pts, fill=item.fill)
            draw.line(pts + [pts[0]], fill=color, width=stroke, joint="curve")
        elif item.shape_type == "circle":
            ix1 = int(round(item.x * scale))
            iy1 = int(round(item.y * scale))
            ix2 = int(round((item.x + item.w) * scale))
            iy2 = int(round((item.y + item.h) * scale))
            if item.fill:
                draw.ellipse([ix1, iy1, ix2, iy2], fill=item.fill)
            draw.ellipse([ix1, iy1, ix2, iy2], outline=item.color, width=stroke)
        elif item.shape_type in ("line", "arrow"):
            p1, p2 = item.visual_points()[:2]
            p1 = (int(round(p1[0] * scale)), int(round(p1[1] * scale)))
            p2 = (int(round(p2[0] * scale)), int(round(p2[1] * scale)))
            draw.line([p1, p2], fill=color, width=stroke)
            if item.shape_type == "arrow":
                angle = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
                head = max(8, int(round(12 * scale)))
                spread = math.pi / 7
                a = p2
                b = (int(round(p2[0] - head * math.cos(angle - spread))), int(round(p2[1] - head * math.sin(angle - spread))))
                c = (int(round(p2[0] - head * math.cos(angle + spread))), int(round(p2[1] - head * math.sin(angle + spread))))
                draw.polygon([a, b, c], fill=color)
        elif item.shape_type == "arrow_head":
            pts = [(int(round(x * scale)), int(round(y * scale))) for x, y in item.visual_points()]
            draw.polygon(pts, fill=item.fill or color, outline=color)
        elif item.is_stat_compare():
            pts = [(int(round(x * scale)), int(round(y * scale))) for x, y in item.visual_points()]
            draw.line(pts, fill=color, width=stroke)
            if item.text:
                tx, ty = item.stat_text_position()
                font = get_font(item.font_family, max(6, int(round(item.font_size * scale))), bold=True)
                bbox = draw.textbbox((0, 0), item.text, font=font)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                draw.text((int(round(tx * scale - tw / 2)), int(round(ty * scale - th))), item.text, fill=color, font=font)

    def render_final_image(self, dpi: int) -> Image.Image:
        scale = max(1.0, dpi / 100.0)
        out_w = max(1, int(round(self.canvas_w * scale)))
        out_h = max(1, int(round(self.canvas_h * scale)))
        out = Image.new("RGBA", (out_w, out_h), self.bg_color)
        draw = ImageDraw.Draw(out)
        for item in sorted(self.items, key=lambda x: x.z_index):
            if isinstance(item, PanelItem):
                if item.pil_image is None:
                    continue
                ix = int(round(item.x * scale))
                iy = int(round(item.y * scale))
                iw = max(1, int(round(item.w * scale)))
                ih = max(1, int(round(item.h * scale)))
                resized = item.pil_image.resize((iw, ih), RESAMPLE_LANCZOS).convert("RGBA")
                out.alpha_composite(resized, (ix, iy))
                if item.border_width > 0:
                    draw.rectangle([ix, iy, ix + iw, iy + ih], outline=item.border_color, width=max(1, int(round(item.border_width * scale))))
                if item.show_label and item.label and item.include_in_labels:
                    font = get_font(item.label_font_family, max(6, int(round(item.label_font_size * scale))), bold=True)
                    draw.text((int(round((item.x + item.label_offset_x) * scale)), int(round((item.y + item.label_offset_y) * scale))), item.label, fill=item.label_color, font=font)
            elif isinstance(item, ShapeItem):
                self.render_shape_on_image(draw, item, scale, out)
            else:
                ix = int(round(item.x * scale))
                iy = int(round(item.y * scale))
                txt_img = self._build_textbox_image(item, scale=scale)
                out.alpha_composite(txt_img, (ix, iy))
        rgb = Image.new("RGB", out.size, self.bg_color)
        rgb.paste(out.convert("RGB"))
        if self.auto_trim_on_export.get():
            rgb = self.trim_image(rgb, pad=max(0, int(round(self.outer_margin.get() * scale))))
        return rgb

    @staticmethod
    def trim_image(img: Image.Image, pad: int = 0) -> Image.Image:
        bg = Image.new(img.mode, img.size, img.getpixel((0, 0)))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        if bbox is None:
            return img
        x1, y1, x2, y2 = bbox
        x1 = max(0, x1 - pad)
        y1 = max(0, y1 - pad)
        x2 = min(img.width, x2 + pad)
        y2 = min(img.height, y2 + pad)
        return img.crop((x1, y1, x2, y2))

    def export_image(self, fmt: str):
        if not self.items:
            messagebox.showinfo("Nothing to export", "Add some panels, shapes, or text first.")
            return
        ext_map = {"PNG": ".png", "TIFF": ".tiff"}
        path = filedialog.asksaveasfilename(title=f"Export {fmt}", defaultextension=ext_map[fmt], filetypes=[(f"{fmt} file", f"*{ext_map[fmt]}")])
        if not path:
            return
        try:
            dpi = max(72, int(self.export_dpi.get()))
            img = self.render_final_image(dpi=dpi)
            save_kwargs = {"dpi": (dpi, dpi)}
            if fmt == "TIFF":
                save_kwargs["compression"] = "tiff_deflate"
            img.save(path, format=fmt, **save_kwargs)
            self.set_status(f"Exported {fmt}: {path}")
        except Exception as e:
            messagebox.showerror("Export error", f"Could not export file.\n\n{e}")

    def export_pdf(self):
        if not self.items:
            messagebox.showinfo("Nothing to export", "Add some panels, shapes, or text first.")
            return
        path = filedialog.asksaveasfilename(title="Export PDF", defaultextension=".pdf", filetypes=[("PDF file", "*.pdf")])
        if not path:
            return
        try:
            dpi = max(72, int(self.export_dpi.get()))
            img = self.render_final_image(dpi=dpi)
            bio = io.BytesIO()
            img.save(bio, format="PNG", dpi=(dpi, dpi))
            png_bytes = bio.getvalue()
            width_pt = img.width / dpi * 72.0
            height_pt = img.height / dpi * 72.0
            doc = fitz.open()
            page = doc.new_page(width=width_pt, height=height_pt)
            page.insert_image(fitz.Rect(0, 0, width_pt, height_pt), stream=png_bytes)
            doc.save(path, deflate=True, garbage=4)
            doc.close()
            self.set_status(f"Exported PDF: {path}")
        except Exception as e:
            messagebox.showerror("Export error", f"Could not export PDF.\n\n{e}")

    def svg_bounds(self):
        if self.auto_trim_on_export.get() and self.items:
            bbox = self.content_bbox()
            if bbox:
                pad = max(0, self.outer_margin.get())
                x1 = max(0, bbox[0] - pad)
                y1 = max(0, bbox[1] - pad)
                x2 = min(self.canvas_w, bbox[2] + pad)
                y2 = min(self.canvas_h, bbox[3] + pad)
                return x1, y1, max(1, x2 - x1), max(1, y2 - y1)
        return 0, 0, self.canvas_w, self.canvas_h

    def svg_points(self, pts, ox, oy):
        return " ".join(f"{x - ox:.3f},{y - oy:.3f}" for x, y in pts)

    def shape_to_svg(self, item: ShapeItem, ox, oy):
        stroke = html.escape(item.color or "black")
        sw = max(1, item.width)
        if item.shape_type == "rectangle":
            pts = self.svg_points(item.visual_points(), ox, oy)
            fill = "none"
            op = 1
            if item.fill:
                fill, op = svg_color_opacity(item.fill, item.fill_alpha)
                fill = html.escape(fill)
            return f'<polygon points="{pts}" fill="{fill}" fill-opacity="{op:.3f}" stroke="{stroke}" stroke-width="{sw}" />'
        if item.shape_type == "highlight_bar":
            pts = self.svg_points(item.visual_points(), ox, oy)
            fill = html.escape(item.fill or item.color or "#f2c200")
            op = max(0.0, min(1.0, item.fill_alpha / 255.0))
            return f'<polygon points="{pts}" fill="{fill}" fill-opacity="{op:.3f}" stroke="{stroke}" stroke-width="{sw}" />'
        if item.shape_type == "circle":
            x = item.x - ox
            y = item.y - oy
            w = item.w
            h = item.h
            fill = "none"
            op = 1
            if item.fill:
                fill, op = svg_color_opacity(item.fill, item.fill_alpha)
                fill = html.escape(fill)
            return f'<ellipse cx="{x + w / 2:.3f}" cy="{y + h / 2:.3f}" rx="{abs(w) / 2:.3f}" ry="{abs(h) / 2:.3f}" fill="{fill}" fill-opacity="{op:.3f}" stroke="{stroke}" stroke-width="{sw}" />'
        if item.shape_type in ("line", "arrow"):
            p1, p2 = item.visual_points()[:2]
            marker = ' marker-end="url(#arrowhead)"' if item.shape_type == "arrow" else ""
            return f'<line x1="{p1[0] - ox:.3f}" y1="{p1[1] - oy:.3f}" x2="{p2[0] - ox:.3f}" y2="{p2[1] - oy:.3f}" stroke="{stroke}" stroke-width="{sw}" stroke-linecap="round"{marker} />'
        if item.shape_type == "arrow_head":
            pts = self.svg_points(item.visual_points(), ox, oy)
            fill = html.escape(item.fill or item.color or "black")
            return f'<polygon points="{pts}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" />'
        if item.is_stat_compare():
            pts = self.svg_points(item.visual_points(), ox, oy)
            s = f'<polyline points="{pts}" fill="none" stroke="{stroke}" stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round" />'
            if item.text:
                tx, ty = item.stat_text_position()
                txt = html.escape(item.text)
                fam = html.escape(item.font_family)
                s += f'<text x="{tx - ox:.3f}" y="{ty - oy:.3f}" text-anchor="middle" font-family="{fam}" font-size="{item.font_size}" font-weight="bold" fill="{stroke}">{txt}</text>'
            return s
        return ""

    def export_svg(self):
        if not self.items:
            messagebox.showinfo("Nothing to export", "Add some panels, shapes, or text first.")
            return
        path = filedialog.asksaveasfilename(title="Export SVG", defaultextension=".svg", filetypes=[("SVG file", "*.svg")])
        if not path:
            return
        try:
            ox, oy, width, height = self.svg_bounds()
            parts = []
            parts.append('<?xml version="1.0" encoding="UTF-8"?>')
            parts.append(f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{width:.3f}" height="{height:.3f}" viewBox="0 0 {width:.3f} {height:.3f}">')
            parts.append('<defs><marker id="arrowhead" markerWidth="10" markerHeight="7" refX="10" refY="3.5" orient="auto" markerUnits="strokeWidth"><polygon points="0 0, 10 3.5, 0 7" fill="context-stroke" /></marker></defs>')
            parts.append(f'<rect x="0" y="0" width="{width:.3f}" height="{height:.3f}" fill="{html.escape(self.bg_color)}" />')
            for item in sorted(self.items, key=lambda x: x.z_index):
                if isinstance(item, PanelItem):
                    if item.pil_image is None:
                        continue
                    data = image_to_base64_png(item.pil_image)
                    parts.append(f'<image x="{item.x - ox:.3f}" y="{item.y - oy:.3f}" width="{item.w:.3f}" height="{item.h:.3f}" href="data:image/png;base64,{data}" xlink:href="data:image/png;base64,{data}" preserveAspectRatio="none" />')
                    if item.border_width > 0:
                        parts.append(f'<rect x="{item.x - ox:.3f}" y="{item.y - oy:.3f}" width="{item.w:.3f}" height="{item.h:.3f}" fill="none" stroke="{html.escape(item.border_color)}" stroke-width="{item.border_width}" />')
                    if item.show_label and item.label and item.include_in_labels:
                        parts.append(f'<text x="{item.x + item.label_offset_x - ox:.3f}" y="{item.y + item.label_offset_y + item.label_font_size - oy:.3f}" font-family="{html.escape(item.label_font_family)}" font-size="{item.label_font_size}" font-weight="bold" fill="{html.escape(item.label_color)}">{html.escape(item.label)}</text>')
                elif isinstance(item, ShapeItem):
                    parts.append(self.shape_to_svg(item, ox, oy))
                else:
                    txt_img = self._build_textbox_image(item, scale=1.0)
                    data = image_to_base64_png(txt_img)
                    parts.append(f'<image x="{item.x - ox:.3f}" y="{item.y - oy:.3f}" width="{txt_img.width:.3f}" height="{txt_img.height:.3f}" href="data:image/png;base64,{data}" xlink:href="data:image/png;base64,{data}" />')
            parts.append('</svg>')
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(parts))
            self.set_status(f"Exported SVG: {path}")
        except Exception as e:
            messagebox.showerror("Export error", f"Could not export SVG.\n\n{e}")

    def item_to_project_dict(self, item: CanvasItem) -> Dict:
        if isinstance(item, PanelItem):
            return {
                "kind": "panel",
                "id": item.id,
                "source_path": item.source_path,
                "image_b64": image_to_base64_png(item.pil_image) if item.pil_image else None,
                "original_size": list(item.original_size),
                "x": item.x,
                "y": item.y,
                "w": item.w,
                "h": item.h,
                "label": item.label,
                "show_label": item.show_label,
                "include_in_labels": item.include_in_labels,
                "label_font_size": item.label_font_size,
                "label_font_family": item.label_font_family,
                "label_color": item.label_color,
                "label_offset_x": item.label_offset_x,
                "label_offset_y": item.label_offset_y,
                "border_width": item.border_width,
                "border_color": item.border_color,
                "z_index": item.z_index,
                "group_id": item.group_id,
            }
        if isinstance(item, ShapeItem):
            return {
                "kind": "shape",
                "id": item.id,
                "shape_type": item.shape_type,
                "x": item.x,
                "y": item.y,
                "w": item.w,
                "h": item.h,
                "color": item.color,
                "width": item.width,
                "fill": item.fill,
                "fill_alpha": item.fill_alpha,
                "rotation": item.rotation,
                "text": item.text,
                "font_size": item.font_size,
                "font_family": item.font_family,
                "z_index": item.z_index,
                "group_id": item.group_id,
                "compare_style": item.compare_style,
                "bracket_height": item.bracket_height,
                "text_gap": item.text_gap,
                "auto_center_text": item.auto_center_text,
                "snap_to_tops": item.snap_to_tops,
            }
        return {
            "kind": "text",
            "id": item.id,
            "text": item.text,
            "rich_runs": item.rich_runs,
            "x": item.x,
            "y": item.y,
            "w": item.w,
            "h": item.h,
            "font_size": item.font_size,
            "font_family": item.font_family,
            "bold": item.bold,
            "italic": item.italic,
            "fill": item.fill,
            "align": item.align,
            "background": item.background,
            "outline": item.outline,
            "rotation": item.rotation,
            "line_spacing": item.line_spacing,
            "z_index": item.z_index,
            "group_id": item.group_id,
        }

    def project_dict_to_item(self, data: Dict) -> CanvasItem:
        if data["kind"] == "panel":
            img = base64_png_to_image(data["image_b64"]) if data.get("image_b64") else None
            return PanelItem(
                id=data["id"],
                source_path=data.get("source_path"),
                pil_image=img,
                original_size=tuple(data.get("original_size", img.size if img else (100, 100))),
                x=data["x"],
                y=data["y"],
                w=data["w"],
                h=data["h"],
                label=data.get("label", ""),
                show_label=data.get("show_label", True),
                include_in_labels=data.get("include_in_labels", True),
                label_font_size=data.get("label_font_size", DEFAULT_LABEL_FONT_SIZE),
                label_font_family=data.get("label_font_family", "Arial"),
                label_color=data.get("label_color", "black"),
                label_offset_x=data.get("label_offset_x", 10),
                label_offset_y=data.get("label_offset_y", 10),
                border_width=data.get("border_width", 0),
                border_color=data.get("border_color", "black"),
                z_index=data.get("z_index", 0),
                group_id=data.get("group_id"),
            )
        if data["kind"] == "shape":
            shape_type = data.get("shape_type", "rectangle")
            if shape_type == "stat_bracket":
                shape_type = "stat_compare"
            return ShapeItem(
                id=data["id"],
                shape_type=shape_type,
                x=data["x"],
                y=data["y"],
                w=data["w"],
                h=data["h"],
                color=data.get("color", "black"),
                width=data.get("width", 2),
                fill=data.get("fill"),
                fill_alpha=data.get("fill_alpha", 0),
                rotation=data.get("rotation", 0),
                text=data.get("text", ""),
                font_size=data.get("font_size", DEFAULT_TEXT_FONT_SIZE),
                font_family=data.get("font_family", "Arial"),
                z_index=data.get("z_index", 0),
                group_id=data.get("group_id"),
                compare_style=data.get("compare_style", "bracket"),
                bracket_height=data.get("bracket_height", DEFAULT_STAT_BRACKET_HEIGHT),
                text_gap=data.get("text_gap", DEFAULT_STAT_TEXT_GAP),
                auto_center_text=data.get("auto_center_text", True),
                snap_to_tops=data.get("snap_to_tops", True),
            )
        return TextItem(
            id=data["id"],
            text=data.get("text", ""),
            rich_runs=data.get("rich_runs", make_plain_runs(data.get("text", ""))),
            x=data["x"],
            y=data["y"],
            w=data["w"],
            h=data["h"],
            font_size=data.get("font_size", DEFAULT_TEXT_FONT_SIZE),
            font_family=data.get("font_family", "Arial"),
            bold=data.get("bold", False),
            italic=data.get("italic", False),
            fill=data.get("fill", "black"),
            align=data.get("align", "left"),
            background=data.get("background"),
            outline=data.get("outline"),
            rotation=data.get("rotation", 0),
            line_spacing=data.get("line_spacing", DEFAULT_LINE_SPACING),
            z_index=data.get("z_index", 0),
            group_id=data.get("group_id"),
        )

    def project_payload(self):
        return {
            "version": PROJECT_VERSION,
            "app": APP_TITLE,
            "canvas_w": self.canvas_w,
            "canvas_h": self.canvas_h,
            "bg_color": self.bg_color,
            "outer_margin": self.outer_margin.get(),
            "gap_x": self.gap_x.get(),
            "gap_y": self.gap_y.get(),
            "default_label_size": self.default_label_size.get(),
            "default_text_size": self.default_text_size.get(),
            "default_font_family": self.default_font_family.get(),
            "default_label_offset_x": self.default_label_offset_x.get(),
            "default_label_offset_y": self.default_label_offset_y.get(),
            "keep_label_order": self.keep_label_order.get(),
            "show_grid": self.show_grid.get(),
            "view_zoom": self.view_zoom,
            "next_id": self.next_id,
            "next_group_id": self.next_group_id,
            "stat_symbol": self.stat_symbol_var.get(),
            "stat_count": self.stat_count_var.get(),
            "stat_text_size": self.stat_text_size_var.get(),
            "stat_bracket_height": self.stat_bracket_height_var.get(),
            "stat_text_gap": self.stat_text_gap_var.get(),
            "stat_color": self.stat_color_var.get(),
            "stat_compare_style": self.stat_compare_style_var.get(),
            "stat_auto_center": self.stat_auto_center_var.get(),
            "stat_snap_tops": self.stat_snap_tops_var.get(),
            "items": [self.item_to_project_dict(i) for i in self.items],
        }

    def save_project(self):
        path = filedialog.asksaveasfilename(title="Save project", defaultextension=".figboard", filetypes=[("Figboard project", "*.figboard"), ("JSON", "*.json")])
        if not path:
            return False
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.project_payload(), f)
            self.project_path = path
            self.dirty = False
            self.set_status(f"Project saved: {path}")
            return True
        except Exception as e:
            messagebox.showerror("Save project error", f"Could not save project.\n\n{e}")
            return False

    def confirm_discard_unsaved(self) -> bool:
        if not self.dirty:
            return True
        resp = messagebox.askyesnocancel("Unsaved changes", "Save changes before continuing?")
        if resp is None:
            return False
        if resp:
            return bool(self.save_project())
        return True

    def open_project(self):
        if not self.confirm_discard_unsaved():
            return
        path = filedialog.askopenfilename(title="Open project", filetypes=[("Figboard project", "*.figboard *.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            self.canvas_w = int(payload.get("canvas_w", DEFAULT_CANVAS_W))
            self.canvas_h = int(payload.get("canvas_h", DEFAULT_CANVAS_H))
            self.bg_color = payload.get("bg_color", DEFAULT_BG)
            self.outer_margin.set(int(payload.get("outer_margin", DEFAULT_OUTER_MARGIN)))
            self.gap_x.set(int(payload.get("gap_x", DEFAULT_GAP_X)))
            self.gap_y.set(int(payload.get("gap_y", DEFAULT_GAP_Y)))
            self.default_label_size.set(int(payload.get("default_label_size", DEFAULT_LABEL_FONT_SIZE)))
            self.default_text_size.set(int(payload.get("default_text_size", DEFAULT_TEXT_FONT_SIZE)))
            self.default_font_family.set(payload.get("default_font_family", "Arial"))
            self.default_label_offset_x.set(int(payload.get("default_label_offset_x", 10)))
            self.default_label_offset_y.set(int(payload.get("default_label_offset_y", 10)))
            self.keep_label_order.set(bool(payload.get("keep_label_order", True)))
            self.show_grid.set(bool(payload.get("show_grid", False)))
            self.view_zoom = float(payload.get("view_zoom", 1.0))
            self.next_id = int(payload.get("next_id", 1))
            self.next_group_id = int(payload.get("next_group_id", 1))
            self.stat_symbol_var.set(payload.get("stat_symbol", "*"))
            self.stat_count_var.set(int(payload.get("stat_count", 3)))
            self.stat_text_size_var.set(int(payload.get("stat_text_size", DEFAULT_STAT_TEXT_SIZE)))
            self.stat_bracket_height_var.set(int(payload.get("stat_bracket_height", DEFAULT_STAT_BRACKET_HEIGHT)))
            self.stat_text_gap_var.set(int(payload.get("stat_text_gap", DEFAULT_STAT_TEXT_GAP)))
            self.stat_color_var.set(payload.get("stat_color", "black"))
            self.stat_compare_style_var.set(payload.get("stat_compare_style", "Bracket line"))
            self.stat_auto_center_var.set(bool(payload.get("stat_auto_center", True)))
            self.stat_snap_tops_var.set(bool(payload.get("stat_snap_tops", True)))
            if hasattr(self, "stat_color_swatch"):
                self.stat_color_swatch.configure(bg=self.stat_color_var.get(), activebackground=self.stat_color_var.get())
            self.items = [self.project_dict_to_item(d) for d in payload.get("items", [])]
            self.selected_ids = []
            self.anchor_selected_id = None
            self.selected_label_panel_id = None
            self.undo_stack.clear()
            self.redo_stack.clear()
            self.project_path = path
            self.dirty = False
            self.cancel_modes()
            self.canvas_w_entry.delete(0, "end")
            self.canvas_w_entry.insert(0, str(self.canvas_w))
            self.canvas_h_entry.delete(0, "end")
            self.canvas_h_entry.insert(0, str(self.canvas_h))
            self.update_zoom_label()
            self.refresh_selected_panel()
            self.redraw()
            self.set_status(f"Project opened: {path}")
        except Exception as e:
            messagebox.showerror("Open project error", f"Could not open project.\n\n{e}")

    def new_canvas(self):
        if not self.confirm_discard_unsaved():
            return
        self.canvas_w = DEFAULT_CANVAS_W
        self.canvas_h = DEFAULT_CANVAS_H
        self.bg_color = DEFAULT_BG
        self.view_zoom = 1.0
        self.items = []
        self.next_id = 1
        self.next_group_id = 1
        self.selected_ids = []
        self.anchor_selected_id = None
        self.selected_label_panel_id = None
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.project_path = None
        self.dirty = False
        self.outer_margin.set(DEFAULT_OUTER_MARGIN)
        self.gap_x.set(DEFAULT_GAP_X)
        self.gap_y.set(DEFAULT_GAP_Y)
        self.canvas_w_entry.delete(0, "end")
        self.canvas_w_entry.insert(0, str(self.canvas_w))
        self.canvas_h_entry.delete(0, "end")
        self.canvas_h_entry.insert(0, str(self.canvas_h))
        self.cancel_modes()
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("New canvas ready")

    def on_close(self):
        if self.confirm_discard_unsaved():
            self.root.destroy()

    def on_canvas_mousewheel(self, event):
        self.canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def on_canvas_mousewheel_horizontal(self, event):
        self.canvas.xview_scroll(-1 * int(event.delta / 120), "units")

    def on_canvas_ctrl_mousewheel(self, event):
        if event.delta > 0:
            self.zoom_by(ZOOM_STEP)
        else:
            self.zoom_by(1 / ZOOM_STEP)

    def on_canvas_mousewheel_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def on_canvas_mousewheel_linux_horizontal(self, event):
        if event.num == 4:
            self.canvas.xview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.xview_scroll(1, "units")

    def move_selected_with_keyboard(self, dx: int, dy: int):
        if not self.selected_ids:
            return
        self.save_undo_state()
        if self.selected_label_panel_id and len(self.selected_ids) == 1:
            item = self.get_item_by_id(self.selected_label_panel_id)
            if isinstance(item, PanelItem):
                item.label_offset_x += dx
                item.label_offset_y += dy
        else:
            for item in self.get_selected_items():
                item.x = max(0, item.x + dx)
                item.y = max(0, item.y + dy)
        self.refresh_selected_panel(live=True)
        self.redraw()
        self.set_status(f"Moved selection by ({dx}, {dy})")


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    root.withdraw()
    splash = tk.Toplevel(root)
    SplashScreen(splash, duration=6200)
    def show_main():
        root.deiconify()
    splash.protocol("WM_DELETE_WINDOW", lambda: None)
    splash.after(6200, show_main)
    splash.wait_window()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    FigureBoardApp(root)
    if not DND_AVAILABLE:
        messagebox.showwarning("Drag-and-drop unavailable", "tkinterdnd2 is not available.\n\nYou can still use the app with the Files button.\n\nInstall with:\npip install tkinterdnd2")
    root.mainloop()


if __name__ == "__main__":
    main()
