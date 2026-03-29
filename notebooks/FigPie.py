import io
import json
import math
import base64
import webbrowser
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

class SplashScreen:
    def __init__(self, root, duration=8000):
        self.root = root
        self.root.overrideredirect(True)
        self.root.attributes("-alpha", 0.0)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")
        self.root.configure(bg="white")
        self.title_label = tk.Label(
            self.root,
            text="FigPie 🍰",
            font=("Cascadia Code", 42, "bold"),
            bg="white",
            fg="black"
        )
        self.title_label.place(relx=0.5, rely=0.43, anchor="center")
        self.subtitle_label = tk.Label(
            self.root,
            text="Make scientific figures as easy as pie 😎",
            font=("Cascadia Code", 20, "italic"),
            bg="white",
            fg="#444444"
        )
        self.subtitle_label.place(relx=0.5, rely=0.50, anchor="center")
        self.dev_label = tk.Label(
            self.root,
            text="Developed by Hamid Taghipour",
            font=("Cascadia Code", 10),
            bg="white",
            fg="navyblue"
        )
        self.dev_label.place(relx=0.5, rely=0.58, anchor="center")
        self.fade_in_out(duration)
    def fade_in_out(self, duration):
        fade_in_ms = 2000
        fade_out_ms = 2000
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

APP_TITLE = "FigPie 🍰"
PROJECT_VERSION = 1
DEFAULT_CANVAS_W = 1600
DEFAULT_CANVAS_H = 1000
DEFAULT_BG = "white"
DEFAULT_OUTER_MARGIN = 30
DEFAULT_GAP_X = 10
DEFAULT_GAP_Y = 10
DEFAULT_LABEL_FONT_SIZE = 28
DEFAULT_TEXT_FONT_SIZE = 22
DEFAULT_LINE_SPACING = 1.15
HANDLE_SIZE = 10
MIN_PANEL_SIZE = 40
SAFE_GAP = 8
DEFAULT_DPI = 600
MIN_ZOOM = 0.25
MAX_ZOOM = 4.0
ZOOM_STEP = 1.15

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
    preferred = [
        "Arial", "Helvetica", "Times New Roman", "Calibri", "DejaVu Sans",
        "Courier New", "Cascadia Code", "Segoe UI"
    ]
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

# @dataclass
# class PanelItem:
#     id: int
#     kind: str = "panel"
#     source_path: Optional[str] = None
#     pil_image: Optional[Image.Image] = None
#     original_size: Tuple[int, int] = (100, 100)
#     x: int = 50
#     y: int = 50
#     w: int = 200
#     h: int = 150
#     label: str = ""
#     show_label: bool = True
#     label_font_size: int = DEFAULT_LABEL_FONT_SIZE
#     label_font_family: str = "Arial"
#     label_offset_x: int = 10
#     label_offset_y: int = 10
#     border_width: int = 0
#     border_color: str = "black"
#     z_index: int = 0
#     group_id: Optional[int] = None
#     tk_preview: Optional[ImageTk.PhotoImage] = None
#     _preview_cache_key: Optional[Tuple] = None
#     _preview_cache_photo: Optional[ImageTk.PhotoImage] = None
#     def bbox(self):
#         return self.x, self.y, self.x + self.w, self.y + self.h
#     def contains(self, px: int, py: int) -> bool:
#         return self.x <= px <= self.x + self.w and self.y <= py <= self.y + self.h
#     def resize_handle_hit(self, px: int, py: int, handle_model: int) -> bool:
#         return (self.x + self.w - handle_model <= px <= self.x + self.w + handle_model and
#                 self.y + self.h - handle_model <= py <= self.y + self.h + handle_model)
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
        return (self.x + self.w - handle_model <= px <= self.x + self.w + handle_model and
                self.y + self.h - handle_model <= py <= self.y + self.h + handle_model)

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
        return (self.x + self.w - handle_model <= px <= self.x + self.w + handle_model and
                self.y + self.h - handle_model <= py <= self.y + self.h + handle_model)
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
    z_index: int = 0
    group_id: Optional[int] = None

    def bbox(self):
        return self.x, self.y, self.x + self.w, self.y + self.h

    def contains(self, px: int, py: int) -> bool:
        return self.x <= px <= self.x + self.w and self.y <= py <= self.y + self.h

    def resize_handle_hit(self, px: int, py: int, handle_model: int) -> bool:
        return (self.x + self.w - handle_model <= px <= self.x + self.w + handle_model and
                self.y + self.h - handle_model <= py <= self.y + self.h + handle_model)

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
    def _on_mousewheel(self, event):
        if self.winfo_ismapped():
            try:
                widget = self.winfo_containing(event.x_root, event.y_root)
            except Exception:
                widget = None
            if widget and str(widget).startswith(str(self)):
                delta = -1 * int(event.delta / 120) if event.delta else 0
                if delta:
                    self.canvas.yview_scroll(delta, "units")
    def _on_mousewheel_linux(self, event):
        if self.winfo_ismapped():
            try:
                widget = self.winfo_containing(event.x_root, event.y_root)
            except Exception:
                widget = None
            if widget and str(widget).startswith(str(self)):
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
        self._setup_tags()
        self._load_item_content()
        self.bind("<Control-b>", lambda e: self.make_bold())
        self.bind("<Control-i>", lambda e: self.make_italic())
    def _setup_tags(self):
        self._reconfigure_tags()
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
        while True:
            next_idx = self.text.index(f"{idx}+1c")
            if self.text.compare(next_idx, ">", "end-1c"):
                ch = self.text.get(idx, "end-1c")
                if not ch:
                    break
                tags = set(self.text.tag_names(idx))
                bold = "bold" in tags or "bolditalic" in tags
                italic = "italic" in tags or "bolditalic" in tags
                if runs and runs[-1]["bold"] == bold and runs[-1]["italic"] == italic:
                    runs[-1]["text"] += ch
                else:
                    runs.append({"text": ch, "bold": bold, "italic": italic})
                break
            ch = self.text.get(idx, next_idx)
            tags = set(self.text.tag_names(idx))
            bold = "bold" in tags or "bolditalic" in tags
            italic = "italic" in tags or "bolditalic" in tags
            if runs and runs[-1]["bold"] == bold and runs[-1]["italic"] == italic:
                runs[-1]["text"] += ch
            else:
                runs.append({"text": ch, "bold": bold, "italic": italic})
            idx = next_idx
            if self.text.compare(idx, ">=", "end-1c"):
                break
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
        self.canvas_w = DEFAULT_CANVAS_W
        self.canvas_h = DEFAULT_CANVAS_H
        self.bg_color = DEFAULT_BG
        self.view_zoom = 1.0
        self.items: List[CanvasItem] = []
        self.next_id = 1
        self.next_group_id = 1
        self.selected_ids: List[int] = []
        self.anchor_selected_id: Optional[int] = None
        self.drag_mode: Optional[str] = None
        self.drag_start = (0, 0)
        self.drag_origin_map: Dict[int, Tuple[int, int, int, int]] = {}
        self.before_drag_state = None
        self.hover_item_id: Optional[int] = None
        self.context_click_item_id: Optional[int] = None
        self.erase_mode = False
        self.erase_start = None
        self.erase_current = None
        self.select_start: Optional[Tuple[int, int]] = None
        self.select_end: Optional[Tuple[int, int]] = None
        self.undo_stack = []
        self.redo_stack = []
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
        self.font_families = list_font_families()
        self._build_ui()
        self._bind_events()
        self._setup_traces()
        self.redraw()
    def apply_default_label_font_family(self):
        family = self.default_font_family.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                item.label_font_family = family
        self.redraw()
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
        left_wrap = ttk.Frame(outer, width=305)
        left_wrap.pack(side="left", fill="y", padx=6, pady=6)
        center = ttk.Frame(outer)
        center.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        right_wrap = ttk.Frame(outer, width=370)
        right_wrap.pack(side="right", fill="y", padx=6, pady=6)
        self.left_scroll = ScrollableFrame(left_wrap, width=305)
        self.left_scroll.pack(fill="both", expand=True)
        left = self.left_scroll.inner
        self.right_scroll = ScrollableFrame(right_wrap, width=370)
        self.right_scroll.pack(fill="both", expand=True)
        right = self.right_scroll.inner
    
        import_box = ttk.LabelFrame(left, text="Import")
        import_box.pack(fill="x", pady=(0, 8))
        ttk.Button(import_box, text="Add image/PDF files 📁", command=self.add_files).pack(fill="x", padx=6, pady=4)
        ttk.Button(import_box, text="Paste from clipboard 📋", command=self.paste_from_clipboard).pack(fill="x", padx=6, pady=4)
        ttk.Button(import_box, text="Add text box 📝", command=self.add_text_box).pack(fill="x", padx=6, pady=4)
        ttk.Button(import_box, text="Add caption box  ", command=lambda: self.add_text_box(default_text="Caption")).pack(fill="x", padx=6, pady=4)
        ttk.Button(import_box, text="Add auto caption below panels", command=self.add_caption_below_last_figure).pack(fill="x", padx=6, pady=4)
        ttk.Button(import_box, text="Add rectangle 🔲", command=lambda: self.add_shape("rectangle")).pack(fill="x", padx=6, pady=2)
        ttk.Button(import_box, text="Add circle ⭕", command=lambda: self.add_shape("circle")).pack(fill="x", padx=6, pady=2)
        ttk.Button(import_box, text="Add line ✏️", command=lambda: self.add_shape("line")).pack(fill="x", padx=6, pady=2)
        ttk.Button(import_box, text="Add arrow 🏹", command=lambda: self.add_shape("arrow")).pack(fill="x", padx=6, pady=2)

        arrange_box = ttk.LabelFrame(left, text="Arrange")
        arrange_box.pack(fill="x", pady=(0, 8))
        grid_row = ttk.Frame(arrange_box)
        grid_row.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Label(grid_row, text="Rows").pack(side="left")
        self.grid_rows_entry = ttk.Entry(grid_row, width=6)
        self.grid_rows_entry.insert(0, "2")
        self.grid_rows_entry.pack(side="left", padx=(4, 12))
        ttk.Label(grid_row, text="Cols").pack(side="left")
        self.grid_cols_entry = ttk.Entry(grid_row, width=6)
        self.grid_cols_entry.insert(0, "2")
        self.grid_cols_entry.pack(side="left", padx=4)
        ttk.Button(arrange_box, text="Apply custom grid 🛠️", command=self.apply_custom_grid).pack(fill="x", padx=6, pady=4)
        ttk.Button(arrange_box, text="Auto grid 🤖", command=self.auto_grid).pack(fill="x", padx=6, pady=4)
        ttk.Button(arrange_box, text="Distribute horizontally ↔️", command=self.distribute_h).pack(fill="x", padx=6, pady=4)
        ttk.Button(arrange_box, text="Distribute vertically ↕️", command=self.distribute_v).pack(fill="x", padx=6, pady=4)
        align_frame = ttk.Frame(arrange_box)
        align_frame.pack(fill="x", padx=6, pady=4)
        ttk.Button(align_frame, text="⬅️", width=3, command=lambda: self.align_to_anchor("left")).pack(side="left", padx=1)
        ttk.Button(align_frame, text="⬆️", width=3, command=lambda: self.align_to_anchor("top")).pack(side="left", padx=1)
        ttk.Button(align_frame, text="⬇️", width=3, command=lambda: self.align_to_anchor("bottom")).pack(side="left", padx=1)
        ttk.Button(align_frame, text="➡️", width=3, command=lambda: self.align_to_anchor("right")).pack(side="left", padx=1)
        ttk.Button(align_frame, text="↔️", width=3, command=lambda: self.align_to_anchor("hcenter")).pack(side="left", padx=1)
        ttk.Button(align_frame, text="↕️", width=3, command=lambda: self.align_to_anchor("vcenter")).pack(side="left", padx=1)
        ttk.Button(arrange_box, text="Same widths 📏", command=self.same_widths).pack(fill="x", padx=6, pady=4)
        ttk.Button(arrange_box, text="Same heights 📐", command=self.same_heights).pack(fill="x", padx=6, pady=4)
        ttk.Button(arrange_box, text="Tight crop to content ✂️", command=self.crop_canvas_to_content).pack(fill="x", padx=6, pady=4)
        ttk.Button(arrange_box, text="Resolve overlaps 🔄", command=self.resolve_all_overlaps).pack(fill="x", padx=6, pady=4)

        group_box = ttk.LabelFrame(left, text="Selection and groups")
        group_box.pack(fill="x", pady=(0, 8))
        ttk.Button(group_box, text="Select all", command=self.select_all_items).pack(fill="x", padx=6, pady=4)
        ttk.Button(group_box, text="Group selected", command=self.group_selected).pack(fill="x", padx=6, pady=4)
        ttk.Button(group_box, text="Ungroup selected", command=self.ungroup_selected).pack(fill="x", padx=6, pady=4)
        ttk.Button(group_box, text="Duplicate selected", command=self.duplicate_selected).pack(fill="x", padx=6, pady=4)
        ttk.Button(group_box, text="Delete selected", command=self.delete_selected).pack(fill="x", padx=6, pady=4)

        edit_box = ttk.LabelFrame(left, text="Edit")
        edit_box.pack(fill="x", pady=(0, 8))
        ttk.Button(edit_box, text="Undo", command=self.undo).pack(fill="x", padx=6, pady=4)
        ttk.Button(edit_box, text="Redo", command=self.redo).pack(fill="x", padx=6, pady=4)
        self.erase_btn = tk.Button(edit_box, text="Toggle erase rectangle mode", command=self.toggle_erase_mode)
        self.erase_btn.pack(fill="x", padx=6, pady=4)
        ttk.Button(edit_box, text="Open rich text editor", command=self.open_rich_text_editor).pack(fill="x", padx=6, pady=4)

        label_box = ttk.LabelFrame(left, text="Labels")
        label_box.pack(fill="x", pady=(0, 8))
        ttk.Button(label_box, text="Regenerate A, B, C... 🔖", command=self.regenerate_labels).pack(fill="x", padx=6, pady=4)
        ttk.Button(label_box, text="Toggle labels on/off 🏷️", command=self.toggle_labels).pack(fill="x", padx=6, pady=4)
        ttk.Label(label_box, text="Default label size").pack(anchor="w", padx=6, pady=(6, 2))
        ttk.Spinbox(label_box, from_=8, to=120, textvariable=self.default_label_size, width=8,
                    command=self.apply_default_label_size).pack(anchor="w", padx=6, pady=(0, 6))
        ttk.Label(label_box, text="Default font family").pack(anchor="w", padx=6, pady=(2, 2))
        #ttk.Combobox(label_box, values=self.font_families, textvariable=self.default_font_family, state="readonly").pack(fill="x", padx=6, pady=(0, 6))
        label_font_combo = ttk.Combobox(label_box, values=self.font_families, textvariable=self.default_font_family, state="readonly")
        label_font_combo.pack(fill="x", padx=6, pady=(0, 6))
        label_font_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_default_label_font_family())
        ttk.Checkbutton(label_box, text="Keep labels by add order", variable=self.keep_label_order).pack(anchor="w", padx=6, pady=4)

        auto_label_frame = ttk.Frame(label_box)
        auto_label_frame.pack(fill="x", padx=6, pady=4)
        ttk.Label(auto_label_frame, text="Auto position labels").pack(anchor="w")
        self.label_pos_var = tk.StringVar(value="top-in")
        pos_combo = ttk.Combobox(auto_label_frame, values=[
            "top-in", "top-out",
            "left-in", "left-out",
            "bottom-in", "bottom-out",
            "center"
        ], textvariable=self.label_pos_var, state="readonly", width=18)
        pos_combo.pack(side="left", padx=(0, 8))
        ttk.Button(auto_label_frame, text="Apply to all", command=self.apply_auto_label_position_all).pack(side="left")
        ttk.Button(auto_label_frame, text="Apply to selected", command=self.apply_auto_label_position_selected).pack(side="left", padx=(4, 0))

        ttk.Label(label_box, text="Default label offset X").pack(anchor="w", padx=6, pady=(6, 2))
        ttk.Spinbox(label_box, from_=-200, to=500, textvariable=self.default_label_offset_x, width=8).pack(anchor="w", padx=6, pady=(0, 4))
        ttk.Label(label_box, text="Default label offset Y").pack(anchor="w", padx=6, pady=(2, 2))
        ttk.Spinbox(label_box, from_=-200, to=500, textvariable=self.default_label_offset_y, width=8).pack(anchor="w", padx=6, pady=(0, 6))
        ttk.Button(label_box, text="Apply offsets to all panels", command=self.apply_default_label_offsets).pack(fill="x", padx=6, pady=4)

        canvas_box = ttk.LabelFrame(left, text="Canvas 🖌️")
        canvas_box.pack(fill="x", pady=(0, 8))
        row = ttk.Frame(canvas_box)
        row.pack(fill="x", padx=6, pady=4)
        ttk.Label(row, text="Unit").pack(side="left")
        ttk.Combobox(row, values=["px", "in", "cm"], textvariable=self.canvas_unit, width=6, state="readonly").pack(side="left", padx=(4, 10))
        ttk.Label(row, text="W").pack(side="left")
        self.canvas_w_entry = ttk.Entry(row, width=8)
        self.canvas_w_entry.insert(0, str(self.canvas_w))
        self.canvas_w_entry.pack(side="left", padx=(4, 10))
        ttk.Label(row, text="H").pack(side="left")
        self.canvas_h_entry = ttk.Entry(row, width=8)
        self.canvas_h_entry.insert(0, str(self.canvas_h))
        self.canvas_h_entry.pack(side="left", padx=4)
        ttk.Button(canvas_box, text="Apply canvas size", command=self.apply_canvas_size).pack(fill="x", padx=6, pady=4)
        ttk.Label(canvas_box, text="Outer margin").pack(anchor="w", padx=6, pady=(6, 2))
        ttk.Spinbox(canvas_box, from_=-500, to=1000, textvariable=self.outer_margin, width=8, command=self.redraw).pack(anchor="w", padx=6, pady=(0, 6))
        ttk.Label(canvas_box, text="Gap X").pack(anchor="w", padx=6, pady=(2, 2))
        ttk.Spinbox(canvas_box, from_=-100, to=1000, textvariable=self.gap_x, width=8, command=self.on_gap_change).pack(anchor="w", padx=6, pady=(0, 6))
        ttk.Label(canvas_box, text="Gap Y").pack(anchor="w", padx=6, pady=(2, 2))
        ttk.Spinbox(canvas_box, from_=-100, to=1000, textvariable=self.gap_y, width=8, command=self.on_gap_change).pack(anchor="w", padx=6, pady=(0, 6))
        ttk.Checkbutton(canvas_box, text="Show grid", variable=self.show_grid, command=self.redraw).pack(anchor="w", padx=6, pady=4)

        zoom_box = ttk.LabelFrame(left, text="Zoom 🔍")
        zoom_box.pack(fill="x", pady=(0, 8))
        zrow = ttk.Frame(zoom_box)
        zrow.pack(fill="x", padx=6, pady=4)
        ttk.Button(zrow, text="Zoom out 🔍➖", command=lambda: self.zoom_by(1 / ZOOM_STEP)).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(zrow, text="Zoom in 🔍➕", command=lambda: self.zoom_by(ZOOM_STEP)).pack(side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Button(zoom_box, text="Reset zoom 🔄", command=self.reset_zoom).pack(fill="x", padx=6, pady=4)
        ttk.Label(zoom_box, textvariable=self.zoom_var).pack(anchor="w", padx=6, pady=(0, 6))

        export_box = ttk.LabelFrame(left, text="Project and export 🎛️")
        export_box.pack(fill="x", pady=(0, 8))
        ttk.Button(export_box, text="Save project 💾", command=self.save_project).pack(fill="x", padx=6, pady=4)
        ttk.Button(export_box, text="Open project 📂", command=self.open_project).pack(fill="x", padx=6, pady=4)
        ttk.Checkbutton(export_box, text="Trim whitespace on export ✂️", variable=self.auto_trim_on_export).pack(anchor="w", padx=6, pady=4)
        dpi_row = ttk.Frame(export_box)
        dpi_row.pack(fill="x", padx=6, pady=4)
        ttk.Label(dpi_row, text="DPI").pack(side="left")
        ttk.Spinbox(dpi_row, from_=72, to=2400, textvariable=self.export_dpi, width=8).pack(side="left", padx=6)
        ttk.Button(export_box, text="Export PNG 🖼️", command=lambda: self.export_image("PNG")).pack(fill="x", padx=6, pady=4)
        ttk.Button(export_box, text="Export TIFF 🖼️", command=lambda: self.export_image("TIFF")).pack(fill="x", padx=6, pady=4)
        ttk.Button(export_box, text="Export PDF 📄", command=self.export_pdf).pack(fill="x", padx=6, pady=4)

        canvas_frame = ttk.Frame(center)
        canvas_frame.pack(fill="both", expand=True)
        self.hbar = ttk.Scrollbar(canvas_frame, orient="horizontal")
        self.vbar = ttk.Scrollbar(canvas_frame, orient="vertical")
        self.canvas = tk.Canvas(
            canvas_frame,
            bg=self.bg_color,
            xscrollcommand=self.hbar.set,
            yscrollcommand=self.vbar.set,
            highlightthickness=1,
            highlightbackground="#888"
        )
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

        prop_box = ttk.LabelFrame(right, text="Selected item")
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
        self.sel_align_var = tk.StringVar(value="left")
        self.sel_bold_var = tk.BooleanVar(value=False)
        self.sel_italic_var = tk.BooleanVar(value=False)
        self.sel_rotation_var = tk.StringVar(value="0")
        self.sel_line_spacing_var = tk.StringVar(value=str(DEFAULT_LINE_SPACING))
        ttk.Label(prop_box, text="Type").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, textvariable=self.sel_type_var).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, text="Label").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_label_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(prop_box, text="Text").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_text_var).grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(prop_box, text="X").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_x_var, width=10).grid(row=3, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, text="Y").grid(row=4, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_y_var, width=10).grid(row=4, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, text="W").grid(row=5, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_w_var, width=10).grid(row=5, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, text="H").grid(row=6, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_h_var, width=10).grid(row=6, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, text="Font size").grid(row=7, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_font_var, width=10).grid(row=7, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(prop_box, text="Font family").grid(row=8, column=0, sticky="w", padx=6, pady=4)
        ttk.Combobox(prop_box, values=self.font_families, textvariable=self.sel_font_family_var, state="readonly").grid(row=8, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(prop_box, text="Align").grid(row=9, column=0, sticky="w", padx=6, pady=4)
        ttk.Combobox(prop_box, values=["left", "center", "right", "justify"], textvariable=self.sel_align_var, state="readonly").grid(row=9, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(prop_box, text="Rotation").grid(row=10, column=0, sticky="w", padx=6, pady=4)
        ttk.Combobox(prop_box, values=["0", "90", "180", "270"], textvariable=self.sel_rotation_var, state="readonly").grid(row=10, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(prop_box, text="Line spacing").grid(row=11, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(prop_box, textvariable=self.sel_line_spacing_var, width=10).grid(row=11, column=1, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(prop_box, text="Bold", variable=self.sel_bold_var).grid(row=12, column=0, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(prop_box, text="Italic", variable=self.sel_italic_var).grid(row=12, column=1, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(prop_box, text="Show panel label", variable=self.sel_show_label_var).grid(
            row=13, column=0, columnspan=2, sticky="w", padx=6, pady=4
        )
        ttk.Button(prop_box, text="Apply changes", command=self.apply_selected_properties).grid(
            row=14, column=0, columnspan=2, sticky="ew", padx=6, pady=6
        )
        ttk.Button(prop_box, text="Rich text editor", command=self.open_rich_text_editor).grid(
            row=15, column=0, columnspan=2, sticky="ew", padx=6, pady=4
        )
        ttk.Button(prop_box, text="Bring to front", command=self.bring_to_front).grid(
            row=16, column=0, columnspan=2, sticky="ew", padx=6, pady=4
        )
        ttk.Button(prop_box, text="Send to back", command=self.send_to_back).grid(
            row=17, column=0, columnspan=2, sticky="ew", padx=6, pady=4
        )
        ttk.Button(prop_box, text="Pick colour", command=self.pick_selected_text_colour).grid(
            row=18, column=0, columnspan=2, sticky="ew", padx=6, pady=4
        )
        prop_box.columnconfigure(1, weight=1)

        help_box = ttk.LabelFrame(right, text="Shortcuts")
        help_box.pack(fill="x", pady=(0, 8))
        help_text = (
            "Ctrl+A select all\n"
            "Ctrl+Z undo\n"
            "Ctrl+Y redo\n"
            "Ctrl+Click multi-select\n"
            "Delete remove selected\n"
            "Ctrl+V paste image\n"
            "Ctrl+O open files\n"
            "Ctrl+S save project\n"
            "Ctrl+Shift+S export PNG\n"
            "Ctrl+G auto grid\n"
            "Ctrl+Mouse wheel zoom\n"
            "Mouse wheel scroll\n"
            "Shift+wheel horizontal scroll\n"
            "Double click text rich editor\n"
            "Drag panel label move label manually\n"
            "Right click context menu\n"
            "Drag empty area = rectangle select\n"
            "Arrow keys move selected (with Ctrl/Shift = faster)"
        )
        ttk.Label(help_box, text=help_text, justify="left").pack(anchor="w", padx=8, pady=8)

        self.bottom_frame = ttk.Frame(self.root)
        self.bottom_frame.pack(fill="x", side="bottom", padx=4, pady=(0, 4))
        ttk.Label(self.bottom_frame, textvariable=self.status_var, anchor="w").pack(side="left", padx=6)
        self.hyperlink_label = tk.Label(
            self.bottom_frame,
            text="Developed by Hamid Taghipourbibalan",
            font=("Cascadia Code", 8, "italic"),
            cursor="hand2"
        )
        self.hyperlink_label.pack(side="right", padx=8, pady=2)
        self.hyperlink_label.bind(
            "<Button-1>",
            lambda e: self.open_hyperlink("https://www.linkedin.com/in/hamid-taghipourbibalan-b7239088/")
        )
        self._build_context_menu()
        self.update_erase_button_style()

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
        self.context_menu.add_command(label="Toggle labels", command=self.toggle_labels)
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

        # self.root.bind("<Delete>", lambda e: self.delete_selected())
        # self.root.bind("<Control-v>", lambda e: self.paste_from_clipboard())
        # self.root.bind("<Control-o>", lambda e: self.add_files())
        # self.root.bind("<Control-s>", lambda e: self.save_project())
        # self.root.bind("<Control-S>", lambda e: self.export_image("PNG"))
        # self.root.bind("<Control-g>", lambda e: self.auto_grid())
        # self.root.bind("<Control-a>", lambda e: self.select_all_items())
        # self.root.bind("<Control-z>", lambda e: self.undo())
        # self.root.bind("<Control-y>", lambda e: self.redo())

        # self.root.bind("<Control-plus>", lambda e: self.zoom_by(ZOOM_STEP))
        # self.root.bind("<Control-equal>", lambda e: self.zoom_by(ZOOM_STEP))
        # self.root.bind("<Control-minus>", lambda e: self.zoom_by(1 / ZOOM_STEP))
        # self.root.bind("<Control-0>", lambda e: self.reset_zoom())

        # self.root.bind("<Left>", lambda e: self.move_selected_with_keyboard(-1, 0))
        # self.root.bind("<Right>", lambda e: self.move_selected_with_keyboard(1, 0))
        # self.root.bind("<Up>", lambda e: self.move_selected_with_keyboard(0, -1))
        # self.root.bind("<Down>", lambda e: self.move_selected_with_keyboard(0, 1))
        # self.root.bind("<Shift-Left>", lambda e: self.move_selected_with_keyboard(-10, 0))
        # self.root.bind("<Shift-Right>", lambda e: self.move_selected_with_keyboard(10, 0))
        # self.root.bind("<Shift-Up>", lambda e: self.move_selected_with_keyboard(0, -10))
        # self.root.bind("<Shift-Down>", lambda e: self.move_selected_with_keyboard(0, 10))
        self.root.bind_all("<Delete>", lambda e: self.delete_selected())
        self.root.bind_all("<Control-v>", lambda e: self.paste_from_clipboard())
        self.root.bind_all("<Control-o>", lambda e: self.add_files())
        self.root.bind_all("<Control-s>", lambda e: self.save_project())
        self.root.bind_all("<Control-Shift-s>", lambda e: self.export_image("PNG"))
        self.root.bind_all("<Control-g>", lambda e: self.auto_grid())
        self.root.bind_all("<Control-a>", lambda e: self.select_all_items())
        self.root.bind_all("<Control-z>", lambda e: self.undo())
        self.root.bind_all("<Control-y>", lambda e: self.redo())

        self.root.bind_all("<Control-plus>", lambda e: self.zoom_by(ZOOM_STEP))
        self.root.bind_all("<Control-equal>", lambda e: self.zoom_by(ZOOM_STEP))
        self.root.bind_all("<Control-minus>", lambda e: self.zoom_by(1 / ZOOM_STEP))
        self.root.bind_all("<Control-0>", lambda e: self.reset_zoom())

        self.root.bind_all("<Left>", lambda e: self.move_selected_with_keyboard(-1, 0))
        self.root.bind_all("<Right>", lambda e: self.move_selected_with_keyboard(1, 0))
        self.root.bind_all("<Up>", lambda e: self.move_selected_with_keyboard(0, -1))
        self.root.bind_all("<Down>", lambda e: self.move_selected_with_keyboard(0, 1))
        self.root.bind_all("<Shift-Left>", lambda e: self.move_selected_with_keyboard(-10, 0))
        self.root.bind_all("<Shift-Right>", lambda e: self.move_selected_with_keyboard(10, 0))
        self.root.bind_all("<Shift-Up>", lambda e: self.move_selected_with_keyboard(0, -10))
        self.root.bind_all("<Shift-Down>", lambda e: self.move_selected_with_keyboard(0, 10))

    def open_hyperlink(self, url: str):
        try:
            webbrowser.open(url)
        except Exception:
            pass

    def set_status(self, text: str):
        self.status_var.set(text)

    def update_zoom_label(self):
        self.zoom_var.set(f"{int(round(self.view_zoom * 100))}%")

    def update_erase_button_style(self):
        if self.erase_mode:
            self.erase_btn.configure(bg="#d83a3a", fg="white", relief="sunken", activebackground="#b82f2f", activeforeground="white")
        else:
            self.erase_btn.configure(bg=self.root.cget("bg"), fg="black", relief="raised", activebackground=self.root.cget("bg"), activeforeground="black")

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
            "next_id": self.next_id,
            "next_group_id": self.next_group_id,
        }

    def save_undo_state(self):
        self.undo_stack.append(self.snapshot_state())
        if len(self.undo_stack) > 100:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _clone_items(self, items):
        cloned = []
        for item in items:
            if isinstance(item, PanelItem):
                # cloned.append(PanelItem(
                #     id=item.id,
                #     source_path=item.source_path,
                #     pil_image=item.pil_image.copy() if item.pil_image else None,
                #     original_size=item.original_size,
                #     x=item.x, y=item.y, w=item.w, h=item.h,
                #     label=item.label, show_label=item.show_label,
                #     label_font_size=item.label_font_size,
                #     label_font_family=item.label_font_family,
                #     label_offset_x=item.label_offset_x,
                #     label_offset_y=item.label_offset_y,
                #     border_width=item.border_width,
                #     border_color=item.border_color,
                #     z_index=item.z_index,
                #     group_id=item.group_id
                # ))
                cloned.append(PanelItem(
                    id=item.id,
                    source_path=item.source_path,
                    pil_image=item.pil_image.copy() if item.pil_image else None,
                    original_size=item.original_size,
                    x=item.x, y=item.y, w=item.w, h=item.h,
                    label=item.label, show_label=item.show_label,
                    label_font_size=item.label_font_size,
                    label_font_family=item.label_font_family,
                    label_color=item.label_color,
                    label_offset_x=item.label_offset_x,
                    label_offset_y=item.label_offset_y,
                    border_width=item.border_width,
                    border_color=item.border_color,
                    z_index=item.z_index,
                    group_id=item.group_id
                ))
            elif isinstance(item, ShapeItem):
                cloned.append(ShapeItem(
                    id=item.id,
                    shape_type=item.shape_type,
                    x=item.x, y=item.y, w=item.w, h=item.h,
                    color=item.color,
                    width=item.width,
                    z_index=item.z_index,
                    group_id=item.group_id
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
                    group_id=item.group_id
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
        self.set_status("Undo")

    def redo(self):
        if not self.redo_stack:
            return
        current = self.snapshot_state()
        self.undo_stack.append(current)
        state = self.redo_stack.pop()
        self.restore_state(state)
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
            self.regenerate_labels(silent=True)
            self.auto_grid_if_reasonable()
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
            self.regenerate_labels(silent=True)
            self.auto_grid_if_reasonable()
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
        offset = 20 * (len(self.items) % 10)
        # panel = PanelItem(
        #     id=self.get_next_id(),
        #     source_path=source_path,
        #     pil_image=img,
        #     original_size=(ow, oh),
        #     x=50 + offset,
        #     y=50 + offset,
        #     w=w,
        #     h=h,
        #     label="",
        #     show_label=True,
        #     label_font_size=self.default_label_size.get(),
        #     label_font_family=self.default_font_family.get(),
        #     label_offset_x=self.default_label_offset_x.get(),
        #     label_offset_y=self.default_label_offset_y.get(),
        #     z_index=max((i.z_index for i in self.items), default=-1) + 1,
        # )
        panel = PanelItem(
            id=self.get_next_id(),
            source_path=source_path,
            pil_image=img,
            original_size=(ow, oh),
            x=50 + offset,
            y=50 + offset,
            w=w,
            h=h,
            label="",
            show_label=True,
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
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Added text box")

    # def add_shape(self, shape_type):
    #     self.save_undo_state()
    #     item = ShapeItem(
    #         id=self.get_next_id(),
    #         shape_type=shape_type,
    #         x=100 + 20 * (len(self.items) % 10),
    #         y=100 + 20 * (len(self.items) % 10),
    #         z_index=max((i.z_index for i in self.items), default=0) + 1
    #     )
    #     self.items.append(item)
    #     self.selected_ids = [item.id]
    #     self.anchor_selected_id = item.id
    #     self.refresh_selected_panel()
    #     self.redraw()
    #     self.set_status(f"Added {shape_type}")
    def add_shape(self, shape_type):
        self.save_undo_state()
        base_x = 100 + 20 * (len(self.items) % 10)
        base_y = 100 + 20 * (len(self.items) % 10)

        if shape_type in ("line", "arrow"):
            item = ShapeItem(
                id=self.get_next_id(),
                shape_type=shape_type,
                x=base_x,
                y=base_y,
                w=160,
                h=0,
                z_index=max((i.z_index for i in self.items), default=0) + 1
            )
        else:
            item = ShapeItem(
                id=self.get_next_id(),
                shape_type=shape_type,
                x=base_x,
                y=base_y,
                w=150,
                h=100,
                z_index=max((i.z_index for i in self.items), default=0) + 1
            )

        self.items.append(item)
        self.selected_ids = [item.id]
        self.anchor_selected_id = item.id
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
        if self.select_start and self.select_end:
            x1, y1 = self.select_start
            x2, y2 = self.select_end
            self.canvas.create_rectangle(
                self.sx(min(x1, x2)), self.sy(min(y1, y2)),
                self.sx(max(x1, x2)), self.sy(max(y1, y2)),
                outline="#2a73ff", width=2, dash=(4, 4)
            )

    def _draw_usable_margin_frame(self):
        m = self.outer_margin.get()
        self.canvas.create_rectangle(
            self.sx(m), self.sy(m), self.sx(self.canvas_w - m), self.sy(self.canvas_h - m),
            outline="#a0a0a0", dash=(4, 3), width=1
        )

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
        preview = item.pil_image.resize((dw, dh), Image.LANCZOS)
        photo = ImageTk.PhotoImage(preview)
        item._preview_cache_key = key
        item._preview_cache_photo = photo
        return photo

    def panel_label_bbox(self, item: PanelItem) -> Optional[Tuple[int, int, int, int]]:
        if not item.show_label or not item.label:
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

    def point_hits_panel_label(self, item: PanelItem, px: int, py: int) -> bool:
        bbox = self.panel_label_bbox(item)
        if bbox is None:
            return False
        x1, y1, x2, y2 = bbox
        pad = 4
        return x1 - pad <= px <= x2 + pad and y1 - pad <= py <= y2 + pad

    def _draw_panel(self, item: PanelItem):
        if item.pil_image is None:
            return
        photo = self._panel_preview_photo(item)
        item.tk_preview = photo
        self.canvas.create_image(self.sx(item.x), self.sy(item.y), image=photo, anchor="nw")
        if item.border_width > 0:
            self.canvas.create_rectangle(
                self.sx(item.x), self.sy(item.y), self.sx(item.x + item.w), self.sy(item.y + item.h),
                outline=item.border_color, width=max(1, int(round(item.border_width * self.view_zoom)))
            )
        if item.show_label and item.label:
            font_style = (
                item.label_font_family,
                max(6, int(round(item.label_font_size * self.view_zoom))),
                "bold"
            )
            tx = self.sx(item.x + item.label_offset_x)
            ty = self.sy(item.y + item.label_offset_y)
            # self.canvas.create_text(
            #     tx,
            #     ty,
            #     text=item.label,
            #     anchor="nw",
            #     font=font_style,
            #     fill="black"
            # )
            self.canvas.create_text(
                tx,
                ty,
                text=item.label,
                anchor="nw",
                font=font_style,
                fill=item.label_color
            )



        if item.id in self.selected_ids:
            lb = self.panel_label_bbox(item)
            if lb:
                self.canvas.create_rectangle(
                    self.sx(lb[0]), self.sy(lb[1]), self.sx(lb[2]), self.sy(lb[3]),
                    outline="#c84d00", dash=(3, 2)
                )

    def _draw_shape(self, item: ShapeItem):
        x1 = self.sx(item.x)
        y1 = self.sy(item.y)
        x2 = self.sx(item.x + item.w)
        y2 = self.sy(item.y + item.h)
        width = max(1, int(round(item.width * self.view_zoom)))
        if item.shape_type == "rectangle":
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=item.color, width=width)
        elif item.shape_type == "circle":
            self.canvas.create_oval(x1, y1, x2, y2, outline=item.color, width=width)
        elif item.shape_type == "line":
            self.canvas.create_line(x1, y1, x2, y2, fill=item.color, width=width)
        elif item.shape_type == "arrow":
            self.canvas.create_line(x1, y1, x2, y2, fill=item.color, width=width, arrow="last")

    def _compress_line_segments(self, draw, chars, font_family, font_size):
        if not chars:
            return []
        out = []
        current = {
            "text": chars[0]["text"],
            "bold": chars[0]["bold"],
            "italic": chars[0]["italic"]
        }
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
            enriched.append({
                "text": seg["text"],
                "bold": seg["bold"],
                "italic": seg["italic"],
                "width": bbox[2] - bbox[0],
                "height": bbox[3] - bbox[1]
            })
        return enriched

    def wrap_runs_to_width(self, draw, runs, font_family, font_size, max_width):
        text = runs_to_text(runs)
        if text == "":
            return [[]]
        chars = []
        for run in normalize_runs(runs):
            for ch in run["text"]:
                chars.append({
                    "text": ch,
                    "bold": bool(run.get("bold", False)),
                    "italic": bool(run.get("italic", False))
                })
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
        return (
            item.w, item.h, round(self.view_zoom, 4), item.font_size, item.font_family,
            item.fill, item.align, item.background, item.outline, item.rotation,
            round(item.line_spacing, 3), json.dumps(item.rich_runs, sort_keys=True)
        )

    def _draw_justified_line(self, draw, line, x, y, max_width, font_family, fsize, fill):
        chars = []
        for seg in line:
            for ch in seg["text"]:
                chars.append({
                    "text": ch,
                    "bold": seg["bold"],
                    "italic": seg["italic"]
                })
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
        lines = self.wrap_runs_to_width(
            draw=draw,
            runs=item.rich_runs,
            font_family=item.font_family,
            font_size=fsize,
            max_width=max_width
        )
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
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Selected all items")

    def on_canvas_motion(self, event):
        x, y = self.model_from_event(event)
        item = self.get_item_at(x, y)
        cursor = ""
        if item is not None:
            handle_model = self.handle_model_size()
            if isinstance(item, PanelItem) and self.point_hits_panel_label(item, x, y):
                cursor = "fleur"
            elif item.resize_handle_hit(x, y, handle_model):
                cursor = "sizing"
            else:
                cursor = "fleur"
        if self.erase_mode:
            cursor = "crosshair"
        self.canvas.config(cursor=cursor)

    def on_canvas_press(self, event):
        self.canvas.focus_set()
        x, y = self.model_from_event(event)
        if self.erase_mode:
            item = self.get_item_at(x, y)
            if isinstance(item, PanelItem):
                self.save_undo_state()
                self.erase_start = (x, y)
                self.erase_current = (x, y)
                return
        item = self.get_item_at(x, y)
        ctrl = bool(event.state & 0x0004)
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
            self.refresh_selected_panel()
            self.redraw()
            return
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
            else:
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
        if isinstance(item, PanelItem) and self.point_hits_panel_label(item, x, y):
            self.drag_mode = "move_label"
            self.drag_origin_map[item.id] = (item.label_offset_x, item.label_offset_y, item.w, item.h)
        elif hit_resize and len(selected_items) == 1:
            self.drag_mode = "resize"
            self.drag_origin_map[item.id] = (item.x, item.y, item.w, item.h)
        else:
            self.drag_mode = "move"
            for sel in selected_items:
                self.drag_origin_map[sel.id] = (sel.x, sel.y, sel.w, sel.h)
        self.refresh_selected_panel()
        self.redraw()

    def on_canvas_drag(self, event):
        x, y = self.model_from_event(event)
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
        if self.drag_mode == "move":
            for item_id, geom in self.drag_origin_map.items():
                item = self.get_item_by_id(item_id)
                if item is None:
                    continue
                ox, oy, ow, oh = geom
                item.x = max(0, ox + dx)
                item.y = max(0, oy + dy)
        # elif self.drag_mode == "resize":
        #     item_id = next(iter(self.drag_origin_map.keys()))
        #     item = self.get_item_by_id(item_id)
        #     if item is not None:
        #         ox, oy, ow, oh = self.drag_origin_map[item_id]
        #         item.w = max(MIN_PANEL_SIZE, ow + dx)
        #         item.h = max(MIN_PANEL_SIZE, oh + dy)
        #         if isinstance(item, (PanelItem, TextItem)):
        #             item._preview_cache_key = None
        #             item._preview_cache_photo = None

        elif self.drag_mode == "resize":
            item_id = next(iter(self.drag_origin_map.keys()))
            item = self.get_item_by_id(item_id)
            if item is not None:
                ox, oy, ow, oh = self.drag_origin_map[item_id]
                if isinstance(item, ShapeItem) and item.shape_type in ("line", "arrow"):
                    new_w = ow + dx
                    new_h = oh + dy
                    snap = 12
                    if abs(new_h) < snap:
                        new_h = 0
                    if abs(new_w) < snap:
                        new_w = 0
                    item.w = new_w
                    item.h = new_h
                else:
                    item.w = max(MIN_PANEL_SIZE, ow + dx)
                    item.h = max(MIN_PANEL_SIZE, oh + dy)
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
                    self.selected_ids = list(set(self.selected_ids + new_selected))
                else:
                    self.selected_ids = new_selected
                self.anchor_selected_id = self.selected_ids[-1] if self.selected_ids else None
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
            self.drag_mode = None
            self.drag_origin_map = {}
            self.before_drag_state = None

    def state_changed(self, a, b) -> bool:
        if a["canvas_w"] != b["canvas_w"] or a["canvas_h"] != b["canvas_h"]:
            return True
        if len(a["items"]) != len(b["items"]):
            return True
        for ia, ib in zip(a["items"], b["items"]):
            if type(ia) != type(ib):
                return True
            if isinstance(ia, PanelItem):
                if (ia.x, ia.y, ia.w, ia.h, ia.group_id, ia.z_index, ia.label, ia.show_label, ia.label_offset_x, ia.label_offset_y) != \
                   (ib.x, ib.y, ib.w, ib.h, ib.group_id, ib.z_index, ib.label, ib.show_label, ib.label_offset_x, ib.label_offset_y):
                    return True
            elif isinstance(ia, ShapeItem):
                if (ia.x, ia.y, ia.w, ia.h, ia.group_id, ia.z_index, ia.shape_type, ia.color, ia.width) != \
                   (ib.x, ib.y, ib.w, ib.h, ib.group_id, ib.z_index, ib.shape_type, ib.color, ib.width):
                    return True
            else:
                if (ia.x, ia.y, ia.w, ia.h, ia.group_id, ia.z_index, ia.text, ia.font_size, ia.font_family,
                    ia.align, ia.rotation, ia.fill, ia.bold, ia.italic, ia.line_spacing, ia.rich_runs) != \
                   (ib.x, ib.y, ib.w, ib.h, ib.group_id, ib.z_index, ib.text, ib.font_size, ib.font_family,
                    ib.align, ib.rotation, ib.fill, ib.bold, ib.italic, ib.line_spacing, ib.rich_runs):
                    return True
        return False

    def on_double_click(self, event):
        x, y = self.model_from_event(event)
        item = self.get_item_at(x, y)
        if item is None:
            return
        if isinstance(item, TextItem):
            self.selected_ids = [item.id]
            self.anchor_selected_id = item.id
            self.refresh_selected_panel()
            self.open_rich_text_editor()
            return
        if isinstance(item, PanelItem):
            self.selected_ids = [item.id]
            self.anchor_selected_id = item.id
            self.refresh_selected_panel()
            self.save_undo_state()
            new_label = simpledialog.askstring("Edit label", "Panel label:", initialvalue=item.label, parent=self.root)
            if new_label is not None:
                item.label = new_label
                self.refresh_selected_panel()
                self.redraw()

    def on_right_click(self, event):
        x, y = self.model_from_event(event)
        item = self.get_item_at(x, y)
        if item is not None and item.id not in self.selected_ids:
            self.selected_ids = [item.id]
            self.anchor_selected_id = item.id
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
        self.sel_type_var.set(sel.kind)
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
            self.sel_align_var.set("left")
            self.sel_bold_var.set(False)
            self.sel_italic_var.set(False)
            self.sel_rotation_var.set("0")
            self.sel_line_spacing_var.set(str(DEFAULT_LINE_SPACING))
        elif isinstance(sel, ShapeItem):
            self.sel_label_var.set("")
            self.sel_text_var.set("")
            self.sel_font_var.set(str(sel.width))
            self.sel_font_family_var.set(self.default_font_family.get())
            self.sel_show_label_var.set(True)
            self.sel_align_var.set("left")
            self.sel_bold_var.set(False)
            self.sel_italic_var.set(False)
            self.sel_rotation_var.set("0")
            self.sel_line_spacing_var.set(str(DEFAULT_LINE_SPACING))
        else:
            sel.sync_text_from_runs()
            self.sel_label_var.set("")
            self.sel_text_var.set(sel.text)
            self.sel_font_var.set(str(sel.font_size))
            self.sel_font_family_var.set(sel.font_family)
            self.sel_show_label_var.set(True)
            self.sel_align_var.set(sel.align)
            self.sel_bold_var.set(sel.bold)
            self.sel_italic_var.set(sel.italic)
            self.sel_rotation_var.set(str(sel.rotation))
            self.sel_line_spacing_var.set(str(sel.line_spacing))
        if not live:
            self.set_status(f"Selected {sel.kind} #{sel.id}")

    def rotate_text_item_box(self, item: TextItem, new_rotation: int):
        old_rotation = item.rotation
        old_parity = old_rotation in (90, 270)
        new_parity = new_rotation in (90, 270)
        if old_parity != new_parity:
            item.w, item.h = item.h, item.w
        item.rotation = new_rotation
        item._preview_cache_key = None
        item._preview_cache_photo = None

    def apply_selected_properties(self):
        if not self.selected_ids:
            return
        self.save_undo_state()
        font_family = self.sel_font_family_var.get()
        align = self.sel_align_var.get()
        rotation = int(self.sel_rotation_var.get())
        try:
            x = int(float(self.sel_x_var.get()))
            y = int(float(self.sel_y_var.get()))
            w = max(MIN_PANEL_SIZE, int(float(self.sel_w_var.get())))
            h = max(MIN_PANEL_SIZE, int(float(self.sel_h_var.get())))
            font_size = max(6, int(float(self.sel_font_var.get())))
            line_spacing = max(0.8, float(self.sel_line_spacing_var.get()))
        except ValueError:
            messagebox.showerror("Invalid input", "Position, size, font size, and line spacing must be numeric.")
            return
        if len(self.selected_ids) == 1:
            sel = self.get_item_by_id(self.selected_ids[0])
            if sel is None:
                return
            if isinstance(sel, PanelItem):
                sel.x, sel.y, sel.w, sel.h = x, y, w, h
                sel.label = self.sel_label_var.get()
                sel.label_font_size = font_size
                sel.label_font_family = font_family
                sel.show_label = self.sel_show_label_var.get()
                sel._preview_cache_key = None
                sel._preview_cache_photo = None
            # elif isinstance(sel, ShapeItem):
            #     sel.x, sel.y, sel.w, sel.h = x, y, w, h
            #     sel.width = font_size

            elif isinstance(sel, ShapeItem):
                old_w = sel.w
                old_h = sel.h

                sel.x, sel.y = x, y
                sel.width = font_size

                if sel.shape_type in ("line", "arrow"):
                    if old_h == 0:
                        sel.w = w
                        sel.h = 0
                    elif old_w == 0:
                        sel.w = 0
                        sel.h = h
                    else:
                        sel.w = w
                        sel.h = h
                else:
                    sel.w, sel.h = w, h



            else:
                old_rotation = sel.rotation
                sel.x, sel.y = x, y
                self.rotate_text_item_box(sel, rotation)
                if old_rotation != rotation and ((old_rotation in (90, 270)) != (rotation in (90, 270))):
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
                    sel.rich_runs = normalize_runs([
                        {"text": r["text"], "bold": r.get("bold", False), "italic": r.get("italic", False)}
                        for r in sel.rich_runs
                    ])
                sel.sync_text_from_runs()
                sel._preview_cache_key = None
                sel._preview_cache_photo = None
                if sel.text.strip().lower().startswith("caption"):
                    self.fit_text_box_height(sel, maybe_expand_canvas=True)
        else:
            anchor = self.get_item_by_id(self.anchor_selected_id) if self.anchor_selected_id else None
            if anchor:
                anchor.x, anchor.y, anchor.w, anchor.h = x, y, w, h
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

    # def pick_selected_text_colour(self):
    #     if len(self.selected_ids) != 1:
    #         return
    #     item = self.get_item_by_id(self.selected_ids[0])
    #     if item is None:
    #         return
    #     initial = item.fill if isinstance(item, TextItem) else item.color if isinstance(item, ShapeItem) else item.border_color if isinstance(item, PanelItem) else "black"
    #     color = colorchooser.askcolor(color=initial, title="Choose colour")[1]
    #     if color:
    #         self.save_undo_state()
    #         if isinstance(item, TextItem):
    #             item.fill = color
    #             item._preview_cache_key = None
    #             item._preview_cache_photo = None
    #         elif isinstance(item, ShapeItem):
    #             item.color = color
    #         elif isinstance(item, PanelItem):
    #             item.border_color = color
    #         self.redraw()


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
            elif isinstance(item, PanelItem):
                item.label_color = color
            self.redraw()

    def pick_selected_label_colour(self):
        if len(self.selected_ids) != 1:
            return
        item = self.get_item_by_id(self.selected_ids[0])
        if not isinstance(item, PanelItem):
            return
        color = colorchooser.askcolor(color=item.label_color, title="Choose label colour")[1]
        if color:
            self.save_undo_state()
            item.label_color = color
            self.redraw() 

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
        for sel in selected:
            if isinstance(sel, PanelItem):
                # dup = PanelItem(
                #     id=self.get_next_id(),
                #     source_path=sel.source_path,
                #     pil_image=sel.pil_image.copy() if sel.pil_image else None,
                #     original_size=sel.original_size,
                #     x=sel.x + 20, y=sel.y + 20, w=sel.w, h=sel.h,
                #     label=sel.label, show_label=sel.show_label,
                #     label_font_size=sel.label_font_size,
                #     label_font_family=sel.label_font_family,
                #     label_offset_x=sel.label_offset_x,
                #     label_offset_y=sel.label_offset_y,
                #     border_width=sel.border_width,
                #     border_color=sel.border_color,
                #     z_index=max((i.z_index for i in self.items), default=0) + 1,
                #     group_id=sel.group_id
                # )
                dup = PanelItem(
                    id=self.get_next_id(),
                    source_path=sel.source_path,
                    pil_image=sel.pil_image.copy() if sel.pil_image else None,
                    original_size=sel.original_size,
                    x=sel.x + 20, y=sel.y + 20, w=sel.w, h=sel.h,
                    label=sel.label, show_label=sel.show_label,
                    label_font_size=sel.label_font_size,
                    label_font_family=sel.label_font_family,
                    label_color=sel.label_color,
                    label_offset_x=sel.label_offset_x,
                    label_offset_y=sel.label_offset_y,
                    border_width=sel.border_width,
                    border_color=sel.border_color,
                    z_index=max((i.z_index for i in self.items), default=0) + 1,
                    group_id=sel.group_id
                )
            elif isinstance(sel, ShapeItem):
                dup = ShapeItem(
                    id=self.get_next_id(),
                    shape_type=sel.shape_type,
                    x=sel.x + 20, y=sel.y + 20, w=sel.w, h=sel.h,
                    color=sel.color,
                    width=sel.width,
                    z_index=max((i.z_index for i in self.items), default=0) + 1,
                    group_id=sel.group_id
                )
            else:
                dup = TextItem(
                    id=self.get_next_id(),
                    text=sel.text,
                    x=sel.x + 20, y=sel.y + 20, w=sel.w, h=sel.h,
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
                    z_index=max((i.z_index for i in self.items), default=0) + 1,
                    group_id=sel.group_id
                )
            self.items.append(dup)
            new_ids.append(dup.id)
        self.selected_ids = new_ids
        self.anchor_selected_id = new_ids[-1] if new_ids else None
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
        self.refresh_selected_panel()
        self.redraw()
        self.set_status("Deleted selected item(s)")

    def regenerate_labels(self, silent=False):
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        panels = sorted(panels, key=lambda p: (p.y, p.x))
        for idx, panel in enumerate(panels):
            panel.label = next_panel_label(idx)
        self.redraw()
        if not silent:
            self.set_status("Regenerated panel labels")

    def toggle_labels(self):
        panels = [i for i in self.items if isinstance(i, PanelItem)]
        if not panels:
            return
        self.save_undo_state()
        turn_on = any(not p.show_label for p in panels)
        for p in panels:
            p.show_label = turn_on
        self.redraw()

    def apply_default_label_size(self):
        size = self.default_label_size.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                item.label_font_size = size
        self.redraw()

    def apply_default_label_offsets(self):
        self.save_undo_state()
        ox = self.default_label_offset_x.get()
        oy = self.default_label_offset_y.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                item.label_offset_x = ox
                item.label_offset_y = oy
        self.redraw()
        self.set_status("Applied custom label offsets to all panels")

    def apply_auto_label_position_all(self):
        self.save_undo_state()
        pos = self.label_pos_var.get()
        for item in self.items:
            if isinstance(item, PanelItem):
                self._apply_auto_label_position(item, pos)
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
        x1 = min(i.x for i in unit)
        y1 = min(i.y for i in unit)
        x2 = max(i.x + i.w for i in unit)
        y2 = max(i.y + i.h for i in unit)
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
        panels_sorted = sorted(panels, key=lambda p: (p.y, p.x))
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
                item.w = max(MIN_PANEL_SIZE, w)
                item._preview_cache_key = None
                item._preview_cache_photo = None
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
                item.h = max(MIN_PANEL_SIZE, h)
                item._preview_cache_key = None
                item._preview_cache_photo = None
        self.redraw()

    def content_bbox(self) -> Optional[Tuple[int, int, int, int]]:
        if not self.items:
            return None
        x1 = min(i.x for i in self.items)
        y1 = min(i.y for i in self.items)
        x2 = max(i.x + i.w for i in self.items)
        y2 = max(i.y + i.h for i in self.items)
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
        self.redraw()
        self.set_status("Canvas cropped to content")

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
        self.redraw()

    def toggle_erase_mode(self):
        self.erase_mode = not self.erase_mode
        self.erase_start = None
        self.erase_current = None
        self.update_erase_button_style()
        self.set_status("Erase rectangle mode ON" if self.erase_mode else "Erase rectangle mode OFF")
        self.redraw()

    def apply_erase_rectangle(self, x1, y1, x2, y2):
        if abs(x2 - x1) < 2 or abs(y2 - y1) < 2:
            return
        rect = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        target = None
        for item in sorted(self.items, key=lambda z: z.z_index, reverse=True):
            if isinstance(item, PanelItem):
                if rects_overlap(rect, item.bbox()):
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
        draw = ImageDraw.Draw(target.pil_image)
        draw.rectangle([sx1, sy1, sx2, sy2], fill="white")
        target._preview_cache_key = None
        target._preview_cache_photo = None
        self.set_status("Erased selected rectangle from panel")

    def fit_text_box_height(self, item: TextItem, maybe_expand_canvas=False):
        img = Image.new("RGBA", (max(100, item.w), 3000), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        lines = self.wrap_runs_to_width(
            draw=draw,
            runs=item.rich_runs,
            font_family=item.font_family,
            font_size=item.font_size,
            max_width=max(20, item.w - 12)
        )
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

    def render_final_image(self, dpi: int) -> Image.Image:
        scale = max(1.0, dpi / 100.0)
        out_w = max(1, int(round(self.canvas_w * scale)))
        out_h = max(1, int(round(self.canvas_h * scale)))
        out = Image.new("RGB", (out_w, out_h), "white")
        draw = ImageDraw.Draw(out)
        for item in sorted(self.items, key=lambda x: x.z_index):
            if isinstance(item, PanelItem):
                if item.pil_image is None:
                    continue
                ix = int(round(item.x * scale))
                iy = int(round(item.y * scale))
                iw = max(1, int(round(item.w * scale)))
                ih = max(1, int(round(item.h * scale)))
                resized = item.pil_image.resize((iw, ih), Image.LANCZOS).convert("RGB")
                out.paste(resized, (ix, iy))
                if item.border_width > 0:
                    draw.rectangle([ix, iy, ix + iw, iy + ih],
                                   outline=item.border_color,
                                   width=max(1, int(round(item.border_width * scale))))
                if item.show_label and item.label:
                    font = get_font(item.label_font_family, max(6, int(round(item.label_font_size * scale))), bold=True)
                    # draw.text(
                    #     (int(round((item.x + item.label_offset_x) * scale)), int(round((item.y + item.label_offset_y) * scale))),
                    #     item.label, fill="black", font=font
                    # )

                    draw.text(
                        (int(round((item.x + item.label_offset_x) * scale)), int(round((item.y + item.label_offset_y) * scale))),
                        item.label, fill=item.label_color, font=font
                    )
            elif isinstance(item, ShapeItem):
                ix1 = int(round(item.x * scale))
                iy1 = int(round(item.y * scale))
                ix2 = int(round((item.x + item.w) * scale))
                iy2 = int(round((item.y + item.h) * scale))
                stroke = max(1, int(round(item.width * scale)))
                if item.shape_type == "rectangle":
                    draw.rectangle([ix1, iy1, ix2, iy2], outline=item.color, width=stroke)
                elif item.shape_type == "circle":
                    draw.ellipse([ix1, iy1, ix2, iy2], outline=item.color, width=stroke)
                elif item.shape_type == "line":
                    draw.line([ix1, iy1, ix2, iy2], fill=item.color, width=stroke)
                elif item.shape_type == "arrow":
                    draw.line([ix1, iy1, ix2, iy2], fill=item.color, width=stroke)
                    angle = math.atan2(iy2 - iy1, ix2 - ix1)
                    head = max(8, int(round(12 * scale)))
                    spread = math.pi / 7
                    p1 = (ix2, iy2)
                    p2 = (int(round(ix2 - head * math.cos(angle - spread))), int(round(iy2 - head * math.sin(angle - spread))))
                    p3 = (int(round(ix2 - head * math.cos(angle + spread))), int(round(iy2 - head * math.sin(angle + spread))))
                    draw.polygon([p1, p2, p3], fill=item.color)
            else:
                ix = int(round(item.x * scale))
                iy = int(round(item.y * scale))
                txt_img = self._build_textbox_image(item, scale=scale)
                out.paste(txt_img.convert("RGB"), (ix, iy), txt_img)
        if self.auto_trim_on_export.get():
            out = self.trim_image(out, pad=max(0, int(round(self.outer_margin.get() * scale))))
        return out

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
            messagebox.showinfo("Nothing to export", "Add some panels or text first.")
            return
        ext_map = {"PNG": ".png", "TIFF": ".tiff"}
        path = filedialog.asksaveasfilename(
            title=f"Export {fmt}",
            defaultextension=ext_map[fmt],
            filetypes=[(f"{fmt} file", f"*{ext_map[fmt]}")]
        )
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
            messagebox.showinfo("Nothing to export", "Add some panels or text first.")
            return
        path = filedialog.asksaveasfilename(
            title="Export PDF",
            defaultextension=".pdf",
            filetypes=[("PDF file", "*.pdf")]
        )
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

    def item_to_project_dict(self, item: CanvasItem) -> Dict:
        if isinstance(item, PanelItem):
            return {
                "kind": "panel",
                "id": item.id,
                "source_path": item.source_path,
                "image_b64": image_to_base64_png(item.pil_image) if item.pil_image else None,
                "original_size": list(item.original_size),
                "x": item.x, "y": item.y, "w": item.w, "h": item.h,
                "label": item.label,
                "show_label": item.show_label,
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
                "x": item.x, "y": item.y, "w": item.w, "h": item.h,
                "color": item.color,
                "width": item.width,
                "z_index": item.z_index,
                "group_id": item.group_id,
            }
        return {
            "kind": "text",
            "id": item.id,
            "text": item.text,
            "rich_runs": item.rich_runs,
            "x": item.x, "y": item.y, "w": item.w, "h": item.h,
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
            # return PanelItem(
            #     id=data["id"],
            #     source_path=data.get("source_path"),
            #     pil_image=img,
            #     original_size=tuple(data.get("original_size", img.size if img else (100, 100))),
            #     x=data["x"], y=data["y"], w=data["w"], h=data["h"],
            #     label=data.get("label", ""),
            #     show_label=data.get("show_label", True),
            #     label_font_size=data.get("label_font_size", DEFAULT_LABEL_FONT_SIZE),
            #     label_font_family=data.get("label_font_family", "Arial"),
            #     label_offset_x=data.get("label_offset_x", 10),
            #     label_offset_y=data.get("label_offset_y", 10),
            #     border_width=data.get("border_width", 0),
            #     border_color=data.get("border_color", "black"),
            #     z_index=data.get("z_index", 0),
            #     group_id=data.get("group_id"),
            # )
        
            return PanelItem(
                id=data["id"],
                source_path=data.get("source_path"),
                pil_image=img,
                original_size=tuple(data.get("original_size", img.size if img else (100, 100))),
                x=data["x"], y=data["y"], w=data["w"], h=data["h"],
                label=data.get("label", ""),
                show_label=data.get("show_label", True),
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
            return ShapeItem(
                id=data["id"],
                shape_type=data.get("shape_type", "rectangle"),
                x=data["x"], y=data["y"], w=data["w"], h=data["h"],
                color=data.get("color", "black"),
                width=data.get("width", 2),
                z_index=data.get("z_index", 0),
                group_id=data.get("group_id"),
            )
        return TextItem(
            id=data["id"],
            text=data.get("text", ""),
            rich_runs=data.get("rich_runs", make_plain_runs(data.get("text", ""))),
            x=data["x"], y=data["y"], w=data["w"], h=data["h"],
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

    def save_project(self):
        path = filedialog.asksaveasfilename(
            title="Save project",
            defaultextension=".figboard",
            filetypes=[("Figboard project", "*.figboard"), ("JSON", "*.json")]
        )
        if not path:
            return
        try:
            payload = {
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
                "items": [self.item_to_project_dict(i) for i in self.items],
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f)
            self.set_status(f"Project saved: {path}")
        except Exception as e:
            messagebox.showerror("Save project error", f"Could not save project.\n\n{e}")

    def open_project(self):
        path = filedialog.askopenfilename(
            title="Open project",
            filetypes=[("Figboard project", "*.figboard *.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            self.save_undo_state()
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
            self.items = [self.project_dict_to_item(d) for d in payload.get("items", [])]
            self.selected_ids = []
            self.anchor_selected_id = None
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
    SplashScreen(splash, duration=8000)

    def show_main():
        root.deiconify()

    splash.protocol("WM_DELETE_WINDOW", lambda: None)
    splash.after(8000, show_main)
    splash.wait_window()

    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    FigureBoardApp(root)

    if not DND_AVAILABLE:
        messagebox.showwarning(
            "Drag-and-drop unavailable",
            "tkinterdnd2 is not available.\n\nYou can still use the app with the 'Add image/PDF files' button.\n\nInstall with:\npip install tkinterdnd2"
        )

    root.mainloop()

if __name__ == "__main__":
    main()