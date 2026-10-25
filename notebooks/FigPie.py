from __future__ import annotations

import asyncio
import base64
import copy
import colorsys
import ctypes
import difflib
import functools
import hashlib
import html
import inspect
import io
import json
import logging
import math
import multiprocessing
import os
import queue
import re
import shutil
import tempfile
import threading
import time
import warnings
import zipfile
from collections import OrderedDict
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from dataclasses import asdict, dataclass, field, fields
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import unquote, urlparse

try:
    from PIL import Image, ImageChops, ImageColor, ImageDraw, ImageFont, ImageOps
    import flet as ft
    import flet.canvas as cv
    import pymupdf as fitz
except ImportError as exc:
    raise SystemExit(
        f"Missing dependency: {exc}\n"
        "Install the application packages in your active Python environment:\n"
        'python -m pip install "flet[all]==1.0.0" "Pillow>=11.3,<13" "PyMuPDF>=1.26,<2" "psutil>=6,<8"'
    ) from exc


__version__ = "0.1.7"


Point = tuple[float, float]
Box = tuple[float, float, float, float]


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def rotate(p: Point, center: Point, degrees: float) -> Point:
    a = math.radians(degrees)
    c, s = math.cos(a), math.sin(a)
    x, y = p[0] - center[0], p[1] - center[1]
    return center[0] + x * c - y * s, center[1] + x * s + y * c


def bounds(points: Iterable[Point]) -> Box:
    points = list(points)
    if not points:
        return 0.0, 0.0, 0.0, 0.0
    return min(p[0] for p in points), min(p[1] for p in points), max(p[0] for p in points), max(p[1] for p in points)


def union(boxes: Iterable[Box]) -> Box:
    boxes = list(boxes)
    return bounds([(b[0], b[1]) for b in boxes] + [(b[2], b[3]) for b in boxes])


def overlaps(a: Box, b: Box, pad: float = 0) -> bool:
    return a[0] < b[2] + pad and a[2] + pad > b[0] and a[1] < b[3] + pad and a[3] + pad > b[1]


def distance_to_segment(p: Point, a: Point, b: Point) -> float:
    dx, dy = b[0] - a[0], b[1] - a[1]
    t = clamp(((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / max(1e-12, dx * dx + dy * dy), 0, 1)
    return math.hypot(p[0] - a[0] - t * dx, p[1] - a[1] - t * dy)


def polygon_contains(p: Point, pts: list[Point]) -> bool:
    inside = False
    j = len(pts) - 1
    for i in range(len(pts)):
        a, b = pts[i], pts[j]
        if (a[1] > p[1]) != (b[1] > p[1]) and p[0] < (b[0] - a[0]) * (p[1] - a[1]) / (b[1] - a[1]) + a[0]:
            inside = not inside
        j = i
    return inside


def canonical(value):
    if isinstance(value, dict):
        return {k: canonical(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [canonical(v) for v in value]
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def label_name(index: int) -> str:
    result = ""
    while True:
        result = chr(65 + index % 26) + result
        index = index // 26 - 1
        if index < 0:
            return result


def normalize_runs(runs: list[dict]) -> list[dict]:
    result: list[dict] = []
    for raw in runs:
        run = {"text": str(raw.get("text", "")), "bold": bool(raw.get("bold", False)), "italic": bool(raw.get("italic", False))}
        if not run["text"]:
            continue
        if result and result[-1]["bold"] == run["bold"] and result[-1]["italic"] == run["italic"]:
            result[-1]["text"] += run["text"]
        else:
            result.append(run)
    return result or [{"text": "", "bold": False, "italic": False}]


@dataclass
class Item:
    id: int
    kind: str = "panel"
    name: str = ""
    x: float = 20
    y: float = 20
    w: float = 320
    h: float = 240
    rotation: float = 0
    z_index: int = 0
    group_id: int | None = None
    locked: bool = False
    visible: bool = True
    asset: str = ""
    source_path: str = ""
    source_page: int = 0
    original_size: list[int] = field(default_factory=lambda: [100, 100])
    crop: list[float] = field(default_factory=lambda: [0, 0, 1, 1])
    erasures: list[list[float]] = field(default_factory=list)
    image_fit: str = "contain"
    lock_aspect: bool = True
    label: str = ""
    show_label: bool = True
    include_in_labels: bool = True
    label_font_size: float = 28
    label_font_family: str = "Arial"
    label_color: str = "#171923"
    label_offset_x: float = 10
    label_offset_y: float = 10
    border_width: float = 0
    border_color: str = "#171923"
    text: str = ""
    rich_runs: list[dict] = field(default_factory=list)
    font_size: float = 22
    font_family: str = "Arial"
    bold: bool = False
    italic: bool = False
    color: str = "#171923"
    align: str = "left"
    line_spacing: float = 1.15
    background: str = ""
    outline: str = ""
    shape_type: str = "rectangle"
    stroke_width: float = 2
    fill: str = ""
    fill_alpha: int = 0
    compare_style: str = "bracket"
    bracket_height: float = 28
    text_gap: float = 10
    auto_center_text: bool = True
    snap_to_tops: bool = True

    def __post_init__(self):
        if not self.rich_runs:
            self.rich_runs = [{"text": self.text, "bold": self.bold, "italic": self.italic}]
        self.validate()

    @property
    def center(self) -> Point:
        return self.x + self.w / 2, self.y + self.h / 2

    @property
    def is_line(self) -> bool:
        return self.kind == "shape" and self.shape_type in ("line", "arrow")

    @property
    def is_stat(self) -> bool:
        return self.kind == "shape" and self.shape_type in ("stat_compare", "stat_bracket")

    def corners(self) -> list[Point]:
        return [rotate(p, self.center, self.rotation) for p in [(self.x, self.y), (self.x + self.w, self.y), (self.x + self.w, self.y + self.h), (self.x, self.y + self.h)]]

    def base_points(self) -> list[Point]:
        if self.is_line:
            return [(self.x, self.y), (self.x + self.w, self.y + self.h)]
        if self.is_stat:
            top = (min(self.y, self.y + self.h) if self.snap_to_tops else self.y) - self.bracket_height
            if self.compare_style == "line":
                return [(self.x, top), (self.x + self.w, top)]
            return [(self.x, self.y), (self.x, top), (self.x + self.w, top), (self.x + self.w, self.y + self.h)]
        if self.shape_type in ("triangle", "arrow_head") and self.kind == "shape":
            return [(self.x + self.w, self.y + self.h / 2), (self.x, self.y), (self.x, self.y + self.h)]
        if self.kind == "shape" and self.shape_type in ("circle", "ellipse"):
            return [(self.center[0] + self.w / 2 * math.cos(i * math.tau / 96), self.center[1] + self.h / 2 * math.sin(i * math.tau / 96)) for i in range(96)]
        return [(self.x, self.y), (self.x + self.w, self.y), (self.x + self.w, self.y + self.h), (self.x, self.y + self.h)]

    def points(self) -> list[Point]:
        return [rotate(p, self.center, self.rotation) for p in self.base_points()]

    def stat_text_point(self) -> Point:
        p = self.base_points()
        return (self.x + self.w / 2 if self.auto_center_text else self.x, min(v[1] for v in p) - self.text_gap)

    def bbox(self) -> Box:
        if self.kind == "shape" and self.shape_type in ("circle", "ellipse"):
            a = math.radians(self.rotation)
            rx = math.hypot(self.w / 2 * math.cos(a), self.h / 2 * math.sin(a))
            ry = math.hypot(self.w / 2 * math.sin(a), self.h / 2 * math.cos(a))
            b = (self.center[0] - rx, self.center[1] - ry, self.center[0] + rx, self.center[1] + ry)
        else:
            b = bounds(self.points())
        if self.is_stat and self.text:
            tx, ty = self.stat_text_point()
            tw = max(10, len(self.text) * self.font_size)
            text_corners = [(tx - tw / 2, ty - self.font_size * 1.3), (tx + tw / 2, ty - self.font_size * 1.3), (tx + tw / 2, ty), (tx - tw / 2, ty)]
            b = union([b, bounds(rotate(p, self.center, self.rotation) for p in text_corners)])
        return b

    def hit(self, p: Point, tolerance: float = 6) -> bool:
        local = rotate(p, self.center, -self.rotation)
        if self.is_line or self.is_stat:
            points = self.points()
            if any(distance_to_segment(p, a, b) <= max(tolerance, self.stroke_width / 2) for a, b in zip(points, points[1:])):
                return True
            if self.is_stat and self.text:
                tx, ty = self.stat_text_point()
                return abs(local[0] - tx) <= len(self.text) * self.font_size / 2 + tolerance and ty - self.font_size * 1.3 - tolerance <= local[1] <= ty + tolerance
            return False
        if self.kind == "shape" and self.shape_type in ("circle", "ellipse"):
            return ((local[0] - self.center[0]) / max(1e-6, abs(self.w) / 2 + tolerance)) ** 2 + ((local[1] - self.center[1]) / max(1e-6, abs(self.h) / 2 + tolerance)) ** 2 <= 1
        if self.kind == "shape" and self.shape_type in ("triangle", "arrow_head"):
            return polygon_contains(local, self.base_points())
        return self.x - tolerance <= local[0] <= self.x + self.w + tolerance and self.y - tolerance <= local[1] <= self.y + self.h + tolerance

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> Item:
        data = copy.deepcopy(data)
        if "stroke_width" not in data and "width" in data:
            data["stroke_width"] = data["width"]
        if data.get("kind") == "text":
            data["color"] = data.get("color", data.get("fill", "#171923"))
            data["fill"] = ""
        if not data.get("rich_runs"):
            data["rich_runs"] = [{"text": data.get("text", ""), "bold": data.get("bold", False), "italic": data.get("italic", False)}]
        allowed = {f.name for f in fields(cls)}
        item = cls(**{k: v for k, v in data.items() if k in allowed and v is not None})
        item.validate()
        return item

    def validate(self) -> None:
        if self.kind not in ("panel", "text", "shape"):
            raise ValueError(f"Unknown item kind: {self.kind}")
        for name in ("x", "y", "w", "h", "rotation", "font_size", "stroke_width", "label_font_size", "line_spacing", "bracket_height", "text_gap", "label_offset_x", "label_offset_y", "border_width"):
            val = float(getattr(self, name))
            if not math.isfinite(val) or abs(val) > 1_000_000:
                raise ValueError(f"Invalid {name}")
            setattr(self, name, val)
        if not (self.is_line or self.is_stat) and (self.w <= 0 or self.h <= 0):
            raise ValueError("Width and height must be positive")
        if not 1 <= self.font_size <= 1000 or not 1 <= self.label_font_size <= 1000:
            raise ValueError("Font size must be between 1 and 1000")
        if not 0 <= self.stroke_width <= 200 or not 0.5 <= self.line_spacing <= 5:
            raise ValueError("Invalid stroke width or line spacing")
        if len(self.crop) != 4 or not 0 <= self.crop[0] < self.crop[2] <= 1 or not 0 <= self.crop[1] < self.crop[3] <= 1:
            raise ValueError("Invalid image crop")
        self.rotation %= 360
        self.rich_runs = normalize_runs(self.rich_runs)
        if self.kind == "text":
            self.text = "".join(r["text"] for r in self.rich_runs)


def resized_corner(item: Item, corner: int, pointer: Point, proportional: bool = False) -> Item:
    result = copy.deepcopy(item)
    opposite = item.corners()[(corner + 2) % 4]
    a = math.radians(item.rotation)
    ux, uy = (math.cos(a), math.sin(a)), (-math.sin(a), math.cos(a))
    dx, dy = pointer[0] - opposite[0], pointer[1] - opposite[1]
    sx, sy = ((-1, -1), (1, -1), (1, 1), (-1, 1))[corner]
    w = max(4, sx * (dx * ux[0] + dy * ux[1]))
    h = max(4, sy * (dx * uy[0] + dy * uy[1]))
    if proportional:
        ratio = max(0.001, abs(item.w / max(0.001, item.h)))
        if w / max(0.001, abs(item.w)) >= h / max(0.001, abs(item.h)):
            h = w / ratio
        else:
            w = h * ratio
    center = (opposite[0] + sx * w / 2 * ux[0] + sy * h / 2 * uy[0], opposite[1] + sx * w / 2 * ux[1] + sy * h / 2 * uy[1])
    result.x, result.y, result.w, result.h = center[0] - w / 2, center[1] - h / 2, w, h
    return result


@dataclass
class Settings:
    canvas_w: float = 1600
    canvas_h: float = 1200
    bg_color: str = "#ffffff"
    margin: float = 20
    gap_x: float = 12
    gap_y: float = 12
    grid_step: float = 50
    show_grid: bool = False
    snap_x: bool = True
    snap_y: bool = True
    snap_grid: bool = False
    snap_gap: bool = True
    default_label_size: float = 28
    default_font_family: str = "Arial"
    default_label_offset_x: float = 10
    default_label_offset_y: float = 10
    keep_label_order: bool = True
    export_dpi: int = 600
    export_scale: float = 1
    export_trim: bool = False


class Document:
    def __init__(self) -> None:
        self.settings = Settings()
        self.items: list[Item] = []
        self.selected: list[int] = []
        self.anchor: int | None = None
        self.label_selected: int | None = None
        self.next_id = 1
        self.next_group_id = 1
        self.undo_stack: list[str] = []
        self.redo_stack: list[str] = []
        self.saved_state = self.snapshot()
        self.history_limit = 100
        self.history_bytes_limit = 24 * 1024 * 1024

    def snapshot(self) -> str:
        return json.dumps(canonical({"version": 2, "settings": asdict(self.settings), "items": [i.to_dict() for i in self.items], "next_id": self.next_id, "next_group_id": self.next_group_id}), ensure_ascii=True, separators=(",", ":"))

    @property
    def dirty(self) -> bool:
        return self.snapshot() != self.saved_state

    def checkpoint(self, before: str | None = None) -> None:
        state = before if before is not None else self.snapshot()
        if not self.undo_stack or self.undo_stack[-1] != state:
            self.undo_stack.append(state)
        while len(self.undo_stack) > self.history_limit or sum(len(x) for x in self.undo_stack) > self.history_bytes_limit:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def restore(self, state: str) -> None:
        data = json.loads(state)
        self.settings = Settings(**data["settings"])
        self.items = [Item.from_dict(d) for d in data["items"]]
        self.next_id = data["next_id"]
        self.next_group_id = data["next_group_id"]
        self.selected = [i for i in self.selected if self.get(i)]
        self.anchor = self.anchor if self.anchor in self.selected else (self.selected[-1] if self.selected else None)
        self.label_selected = self.label_selected if self.label_selected in self.selected else None

    def undo(self) -> bool:
        if not self.undo_stack:
            return False
        self.redo_stack.append(self.snapshot())
        self.restore(self.undo_stack.pop())
        return True

    def redo(self) -> bool:
        if not self.redo_stack:
            return False
        self.undo_stack.append(self.snapshot())
        self.restore(self.redo_stack.pop())
        return True

    def get(self, identifier: int | None) -> Item | None:
        return next((i for i in self.items if i.id == identifier), None)

    def new_item(self, **kwargs: Any) -> Item:
        item = Item(id=self.next_id, z_index=max((i.z_index for i in self.items), default=-1) + 1, **kwargs)
        self.next_id += 1
        self.items.append(item)
        self.select(item.id)
        return item

    def members(self, item: Item) -> list[Item]:
        return [i for i in self.items if i.group_id == item.group_id] if item.group_id is not None else [item]

    def editable(self, item: Item) -> bool:
        return not any(i.locked for i in self.members(item))

    def select(self, identifier: int, add: bool = False, label: bool = False, toggle: bool = False) -> None:
        item = self.get(identifier)
        if not item:
            return
        ids = [identifier] if label else [i.id for i in self.members(item)]
        if not add:
            self.selected = []
        if add and toggle and all(i in self.selected for i in ids):
            self.selected = [i for i in self.selected if i not in ids]
            self.anchor = self.selected[-1] if self.selected else None
        else:
            self.selected = list(dict.fromkeys(self.selected + ids))
            self.anchor = identifier
        self.label_selected = identifier if label else None

    def units(self, all_if_empty: bool = False) -> list[list[Item]]:
        source = [self.get(i) for i in self.selected] if self.selected else (list(self.items) if all_if_empty else [])
        done: set[int] = set()
        result = []
        for item in source:
            if not item or item.id in done or not item.visible:
                continue
            group = self.members(item)
            done.update(i.id for i in group)
            if not any(i.locked or not i.visible for i in group):
                result.append(group)
        return result

    @staticmethod
    def unit_box(unit: list[Item]) -> Box:
        return union(i.bbox() for i in unit)

    @staticmethod
    def translate(unit: list[Item], dx: float, dy: float) -> None:
        for i in unit:
            i.x += dx
            i.y += dy

    def move_selected(self, dx: float, dy: float) -> None:
        if self.label_selected and len(self.selected) == 1:
            item = self.get(self.label_selected)
            if item and self.editable(item):
                item.label_offset_x += dx
                item.label_offset_y += dy
        else:
            for unit in self.units():
                self.translate(unit, dx, dy)

    def append_panel(self, asset: str, size: list[int], source: str, page: int = 0) -> Item:
        scale = min(350 / max(1, size[0]), 260 / max(1, size[1]), 1)
        w, h = max(4, size[0] * scale), max(4, size[1] * scale)
        margin = self.settings.margin
        panels = [i for i in self.items if i.kind == "panel"]
        x, y = margin, margin
        if panels:
            last = panels[-1]
            x, y = last.bbox()[2] + max(8, self.settings.gap_x), last.bbox()[1]
            if x + w > self.settings.canvas_w - margin:
                x, y = margin, max(i.bbox()[3] for i in self.items) + max(8, self.settings.gap_y)
        for _ in range(len(self.items) + 2):
            collisions = [i for i in self.items if overlaps((x, y, x + w, y + h), i.bbox(), 2)]
            if not collisions:
                break
            x, y = margin, max(i.bbox()[3] for i in collisions) + max(8, self.settings.gap_y)
        self.settings.canvas_h = max(self.settings.canvas_h, y + h + margin)
        return self.new_item(asset=asset, original_size=size, source_path=source, source_page=page, name=source.rsplit("/", 1)[-1].rsplit("\\", 1)[-1] + (f" · p{page + 1}" if page else ""), x=x, y=y, w=w, h=h, label_font_size=self.settings.default_label_size, label_font_family=self.settings.default_font_family, label_offset_x=self.settings.default_label_offset_x, label_offset_y=self.settings.default_label_offset_y)

    def replace_image(self, identifier: int, asset: str, size: list[int], source: str, page: int = 0) -> None:
        item = self.get(identifier)
        if not item or item.kind != "panel":
            raise ValueError("Select an image panel to replace")
        if not self.editable(item):
            raise ValueError("Unlock this panel or its group first")
        item.asset, item.original_size, item.source_path, item.source_page = asset, list(size), source, page
        item.crop = [0, 0, 1, 1]
        item.erasures = []

    def group(self) -> None:
        units = self.units()
        members = [i for u in units for i in u]
        if len(members) < 2:
            return
        for i in members:
            i.group_id = self.next_group_id
        self.next_group_id += 1
        self.label_selected = None

    def ungroup(self) -> None:
        for unit in self.units():
            for item in unit:
                item.group_id = None

    def delete(self) -> None:
        if self.label_selected and len(self.selected) == 1:
            item = self.get(self.label_selected)
            if item and self.editable(item):
                item.show_label = False
            self.label_selected = None
            return
        doomed = {i.id for unit in self.units() for i in unit}
        self.items = [i for i in self.items if i.id not in doomed]
        self.selected, self.anchor, self.label_selected = [], None, None

    def duplicate(self) -> None:
        clones = [copy.deepcopy(i) for u in self.units() for i in u]
        group_map: dict[int, int] = {}
        result = []
        for item in clones:
            item.id = self.next_id
            self.next_id += 1
            item.x += 20
            item.y += 20
            if item.group_id is not None:
                if item.group_id not in group_map:
                    group_map[item.group_id] = self.next_group_id
                    self.next_group_id += 1
                item.group_id = group_map[item.group_id]
            item.z_index = max((i.z_index for i in self.items), default=0) + 1
            self.items.append(item)
            result.append(item.id)
        self.selected = result
        self.anchor = result[-1] if result else None
        self.label_selected = None

    def layer(self, front: bool) -> None:
        selected = sorted([i for u in self.units() for i in u], key=lambda i: i.z_index)
        base = max((i.z_index for i in self.items), default=0) + 1 if front else min((i.z_index for i in self.items), default=0) - len(selected)
        for n, item in enumerate(selected):
            item.z_index = base + n

    def align(self, where: str) -> None:
        units = self.units()
        anchor = next((u for u in units if any(i.id == self.anchor for i in u)), None)
        if len(units) < 2 or anchor is None:
            raise ValueError("Select at least two independent panels or groups. The last selected is the anchor.")
        a = self.unit_box(anchor)
        for unit in units:
            if unit is anchor:
                continue
            b = self.unit_box(unit)
            dx = a[0] - b[0] if where == "left" else a[2] - b[2] if where == "right" else (a[0] + a[2] - b[0] - b[2]) / 2 if where == "center_x" else 0
            dy = a[1] - b[1] if where == "top" else a[3] - b[3] if where == "bottom" else (a[1] + a[3] - b[1] - b[3]) / 2 if where == "center_y" else 0
            self.translate(unit, dx, dy)

    def distribute(self, axis: str, automatic: bool = False, equal: bool = False) -> None:
        units = self.units(all_if_empty=automatic)
        if len(units) < 2:
            raise ValueError("Select at least two independent panels or groups")
        dim = 0 if axis == "x" else 1
        units.sort(key=lambda u: self.unit_box(u)[dim])
        boxes = [self.unit_box(u) for u in units]
        sizes = [b[dim + 2] - b[dim] for b in boxes]
        gap = self.settings.gap_x if dim == 0 else self.settings.gap_y
        if equal:
            gap = (boxes[-1][dim + 2] - boxes[0][dim] - sum(sizes)) / (len(units) - 1)
            if gap < 0:
                raise ValueError("The selection is too narrow for equal spacing. Use Auto H/V or increase the space first.")
        other = 1 - dim
        anchor = next((b for u, b in zip(units, boxes) if any(i.id == self.anchor for i in u)), boxes[0])
        center_other = (anchor[other] + anchor[other + 2]) / 2
        cursor = boxes[0][dim]
        for unit, b, size in zip(units, boxes, sizes):
            delta = [0.0, 0.0]
            delta[dim] = cursor - b[dim]
            if automatic:
                delta[other] = center_other - (b[other] + b[other + 2]) / 2
            self.translate(unit, *delta)
            cursor += size + gap
        self.expand_to_content()

    def scale_unit(self, unit: list[Item], factor: float, origin: Point, scale_styles: bool | None = None) -> None:
        for item in unit:
            item.x = origin[0] + (item.x - origin[0]) * factor
            item.y = origin[1] + (item.y - origin[1]) * factor
            item.w *= factor
            item.h *= factor
            if scale_styles if scale_styles is not None else len(unit) > 1:
                item.label_offset_x *= factor
                item.label_offset_y *= factor
                item.label_font_size *= factor
                item.font_size *= factor
                item.stroke_width *= factor
                item.bracket_height *= factor
                item.text_gap *= factor

    def same_size(self, axis: str) -> None:
        units = self.units()
        anchor = next((u for u in units if any(i.id == self.anchor for i in u)), None)
        if len(units) < 2 or anchor is None:
            raise ValueError("Select at least two panels or groups")
        dim = 0 if axis == "x" else 1
        a = self.unit_box(anchor)
        target = a[dim + 2] - a[dim]
        for unit in units:
            if unit is anchor:
                continue
            b = self.unit_box(unit)
            size = b[dim + 2] - b[dim]
            if size > 0:
                self.scale_unit(unit, target / size, (b[0], b[1]))

    def grid(self, rows: int = 0, cols: int = 0) -> None:
        units = self.units(all_if_empty=True)
        if not self.selected:
            units = [u for u in units if any(i.kind == "panel" for i in u)]
        if not units:
            raise ValueError("There are no unlocked panels to arrange")
        units.sort(key=lambda u: min(i.id for i in u))
        n = len(units)
        s = self.settings
        if rows <= 0 or cols <= 0:
            options = []
            for c in range(1, n + 1):
                r = math.ceil(n / c)
                cw = (s.canvas_w - 2 * s.margin - (c - 1) * s.gap_x) / c
                ch = (s.canvas_h - 2 * s.margin - (r - 1) * s.gap_y) / r
                if min(cw, ch) > 4:
                    occupancy = 0.0
                    for u in units:
                        b = self.unit_box(u)
                        bw, bh = max(1, b[2] - b[0]), max(1, b[3] - b[1])
                        f = min(cw / bw, ch / bh)
                        occupancy += bw * bh * f * f
                    options.append((occupancy, r, c))
            if not options:
                raise ValueError("Canvas is too small for these margins and gaps")
            _, rows, cols = max(options)
        if rows * cols < n:
            raise ValueError(f"{n} objects need more than {rows} × {cols} cells")
        cw = (s.canvas_w - 2 * s.margin - (cols - 1) * s.gap_x) / cols
        ch = (s.canvas_h - 2 * s.margin - (rows - 1) * s.gap_y) / rows
        if min(cw, ch) < 4:
            raise ValueError("Margins and gaps leave no room for this grid")
        for index, unit in enumerate(units):
            b = self.unit_box(unit)
            f = min(cw / max(1, b[2] - b[0]), ch / max(1, b[3] - b[1]))
            self.scale_unit(unit, f, (b[0], b[1]))
            scaled = self.unit_box(unit)
            x = s.margin + index % cols * (cw + s.gap_x) + (cw - scaled[2] + scaled[0]) / 2
            y = s.margin + index // cols * (ch + s.gap_y) + (ch - scaled[3] + scaled[1]) / 2
            self.translate(unit, x - scaled[0], y - scaled[1])

    def resolve_overlaps(self) -> None:
        units = self.units(all_if_empty=True)
        moving = {i.id for u in units for i in u}
        fixed = [i.bbox() for i in self.items if i.visible and i.id not in moving]
        units.sort(key=lambda u: (self.unit_box(u)[1], self.unit_box(u)[0]))
        for u in units:
            for _ in range(len(self.items) + 3):
                b = self.unit_box(u)
                hits = [a for a in fixed if overlaps(a, b)]
                if not hits:
                    break
                x = max(a[2] for a in hits) + max(8, self.settings.gap_x)
                y = b[1]
                if x + b[2] - b[0] > self.settings.canvas_w - self.settings.margin:
                    x, y = self.settings.margin, max(a[3] for a in hits) + max(8, self.settings.gap_y)
                self.translate(u, x - b[0], y - b[1])
            fixed.append(self.unit_box(u))
        self.expand_to_content()

    def generate_labels(self) -> None:
        panels = [i for i in self.items if i.kind == "panel" and i.include_in_labels and self.editable(i)]
        panels.sort(key=(lambda i: i.id) if self.settings.keep_label_order else (lambda i: (round(i.y / 20), i.x, i.id)))
        used = {i.label for i in self.items if i.kind == "panel" and not self.editable(i) and i.label}
        index = 0
        for p in panels:
            while label_name(index) in used:
                index += 1
            p.label, p.show_label = label_name(index), True
            index += 1

    def align_labels(self, axis: str) -> None:
        panels = [self.get(i) for i in self.selected]
        panels = [i for i in panels if i and i.kind == "panel" and self.editable(i)]
        anchor = self.get(self.anchor)
        if len(panels) < 2 or not anchor or anchor.kind != "panel":
            raise ValueError("Select at least two panel labels or panels; the last selected is the anchor")
        ax, ay = anchor.x + anchor.label_offset_x, anchor.y + anchor.label_offset_y
        for item in panels:
            if item.id == anchor.id:
                continue
            if axis in ("x", "xy"):
                item.label_offset_x = ax - item.x
            if axis in ("y", "xy"):
                item.label_offset_y = ay - item.y

    def expand_to_content(self) -> None:
        if self.items:
            b = union(i.bbox() for i in self.items)
            self.settings.canvas_h = max(self.settings.canvas_h, b[3] + self.settings.margin)
            self.settings.canvas_w = max(self.settings.canvas_w, b[2] + self.settings.margin)


def snap_translation(box: Box, targets: list[Box], dx: float, dy: float, settings: Settings, zoom: float, disabled: bool = False) -> tuple[float, float, list[tuple[str, float]]]:
    if disabled:
        return dx, dy, []
    threshold = 7 / max(0.05, zoom)
    moved = (box[0] + dx, box[1] + dy, box[2] + dx, box[3] + dy)
    guides: list[tuple[str, float]] = []
    delta = [dx, dy]
    for axis, enabled in enumerate((settings.snap_x, settings.snap_y)):
        if not enabled:
            continue
        starts = [moved[axis], (moved[axis] + moved[axis + 2]) / 2, moved[axis + 2]]
        maximum = settings.canvas_w if axis == 0 else settings.canvas_h
        candidate_values = [(v, 1) for v in [0, settings.margin, maximum / 2, maximum - settings.margin, maximum]]
        for target in targets:
            candidate_values.extend((v, 0) for v in [target[axis], (target[axis] + target[axis + 2]) / 2, target[axis + 2]])
        candidates: list[tuple[float, int, float, float]] = []
        for start in starts:
            for value, priority in candidate_values:
                if abs(value - start) <= threshold:
                    candidates.append((abs(value - start), priority, value - start, value))
        if settings.snap_gap:
            other = 1 - axis
            gap = settings.gap_x if axis == 0 else settings.gap_y
            neighbors = [b for b in targets if b[other] <= moved[other + 2] and b[other + 2] >= moved[other]]
            for b in neighbors:
                for offset, guide in [(b[axis + 2] + gap - starts[0], b[axis + 2] + gap), (b[axis] - gap - starts[2], b[axis] - gap)]:
                    if abs(offset) <= threshold:
                        candidates.append((abs(offset), 2, offset, guide))
            left = [b for b in neighbors if b[axis + 2] <= moved[axis] + threshold]
            right = [b for b in neighbors if b[axis] >= moved[axis + 2] - threshold]
            if left and right:
                a, b = max(v[axis + 2] for v in left), min(v[axis] for v in right)
                start = (a + b - (box[axis + 2] - box[axis])) / 2
                if b - a >= box[axis + 2] - box[axis] and abs(start - starts[0]) <= threshold:
                    candidates.append((abs(start - starts[0]), 2, start - starts[0], start))
        if candidates:
            _, _, offset, guide = min(candidates)
            delta[axis] += offset
            guides.append(("x" if axis == 0 else "y", guide))
        elif settings.snap_grid and settings.grid_step > 0:
            target = round(starts[0] / settings.grid_step) * settings.grid_step
            if abs(target - starts[0]) <= threshold:
                delta[axis] += target - starts[0]
                guides.append(("x" if axis == 0 else "y", target))
    return delta[0], delta[1], guides


def image_area(item: Item) -> Box:
    sw = item.original_size[0] * (item.crop[2] - item.crop[0])
    sh = item.original_size[1] * (item.crop[3] - item.crop[1])
    if item.image_fit == "stretch":
        return item.x, item.y, item.x + item.w, item.y + item.h
    factor = min(item.w / max(1, sw), item.h / max(1, sh)) if item.image_fit == "contain" else max(item.w / max(1, sw), item.h / max(1, sh))
    w, h = sw * factor, sh * factor
    return item.x + (item.w - w) / 2, item.y + (item.h - h) / 2, item.x + (item.w + w) / 2, item.y + (item.h + h) / 2


def crop_or_erase(item: Item, start: Point, end: Point, erase: bool = False) -> bool:
    a, b = rotate(start, item.center, -item.rotation), rotate(end, item.center, -item.rotation)
    area = image_area(item)
    box = (max(item.x, area[0], min(a[0], b[0])), max(item.y, area[1], min(a[1], b[1])), min(item.x + item.w, area[2], max(a[0], b[0])), min(item.y + item.h, area[3], max(a[1], b[1])))
    if box[2] - box[0] < 2 or box[3] - box[1] < 2:
        return False
    old = list(item.crop)
    src = [old[0] + (box[0] - area[0]) / max(1e-9, area[2] - area[0]) * (old[2] - old[0]), old[1] + (box[1] - area[1]) / max(1e-9, area[3] - area[1]) * (old[3] - old[1]), old[0] + (box[2] - area[0]) / max(1e-9, area[2] - area[0]) * (old[2] - old[0]), old[1] + (box[3] - area[1]) / max(1e-9, area[3] - area[1]) * (old[3] - old[1])]
    src = [clamp(v, 0, 1) for v in src]
    if erase:
        item.erasures.append(src)
    else:
        label = item.x + item.label_offset_x, item.y + item.label_offset_y
        center = rotate(((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), item.center, item.rotation)
        item.w, item.h = box[2] - box[0], box[3] - box[1]
        item.x, item.y = center[0] - item.w / 2, center[1] - item.h / 2
        item.crop = src
        item.label_offset_x, item.label_offset_y = label[0] - item.x, label[1] - item.y
    return True


MAX_SOURCE_PIXELS = int(os.getenv("FIGPIE_MAX_SOURCE_PIXELS", "60000000"))
MAX_OUTPUT_PIXELS = int(os.getenv("FIGPIE_MAX_OUTPUT_PIXELS", "80000000"))
PREVIEW_EDGE = int(os.getenv("FIGPIE_PREVIEW_EDGE", "1024"))
MAX_SOURCE_BYTES = 1024 * 1024 * 1024
MAX_PROJECT_BYTES = 4 * 1024 * 1024 * 1024
Image.MAX_IMAGE_PIXELS = MAX_SOURCE_PIXELS
warnings.simplefilter("error", Image.DecompressionBombWarning)


def asset_path(root: str, asset: str) -> Path:
    if not re.fullmatch(r"[0-9a-f]{64}\.[a-z0-9]{1,8}", asset):
        raise ValueError("Invalid asset reference")
    return Path(root) / "assets" / asset


def ensure_root(root: str) -> None:
    (Path(root) / "assets").mkdir(parents=True, exist_ok=True)
    (Path(root) / "previews").mkdir(parents=True, exist_ok=True)


def check_pixels(w: int, h: int, limit: int = MAX_SOURCE_PIXELS) -> None:
    if min(w, h) < 1 or w * h > limit:
        raise ValueError(f"Image size {w:,} × {h:,} exceeds the safety limit of {limit / 1_000_000:g} megapixels. Resize the source first.")


def store_bytes(data: bytes, extension: str, root: str) -> str:
    ensure_root(root)
    if len(data) > MAX_SOURCE_BYTES:
        raise ValueError("The source file exceeds the 1 GB safety limit")
    extension = extension.lower().lstrip(".")
    if not re.fullmatch(r"[a-z0-9]{1,8}", extension):
        raise ValueError("Invalid asset extension")
    asset = hashlib.sha256(data).hexdigest() + "." + extension
    path = asset_path(root, asset)
    if not path.exists():
        path.write_bytes(data)
    return asset


def copy_asset(source: str, root: str) -> str:
    ensure_root(root)
    p = Path(source)
    if not p.is_file():
        raise ValueError(f"File not found: {source}")
    if p.stat().st_size > MAX_SOURCE_BYTES:
        raise ValueError("The source file exceeds the 1 GB safety limit")
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    ext = p.suffix.lower().lstrip(".")
    if ext not in {"png", "jpg", "jpeg", "tif", "tiff", "bmp", "gif", "webp", "pdf"}:
        raise ValueError(f"Unsupported file type: {p.suffix}")
    asset = h.hexdigest() + "." + ext
    out = asset_path(root, asset)
    if not out.exists():
        if shutil.disk_usage(out.parent).free < p.stat().st_size + 256 * 1024 * 1024:
            raise ValueError("Not enough temporary disk space. Free disk space before importing.")
        fd, name = tempfile.mkstemp(dir=out.parent)
        os.close(fd)
        try:
            shutil.copyfile(p, name)
            os.replace(name, out)
        finally:
            Path(name).unlink(missing_ok=True)
    return asset


def inspect_source(path: Path) -> list[list[int]]:
    if path.suffix.lower() == ".pdf":
        with fitz.open(path) as doc:
            if doc.needs_pass:
                raise ValueError("Password-protected PDFs must be unlocked before importing")
            if len(doc) > 500:
                raise ValueError("This PDF has more than 500 pages. Split it before importing.")
            return [[max(1, round(p.rect.width * 3)), max(1, round(p.rect.height * 3))] for p in doc]
    with Image.open(path) as im:
        check_pixels(*im.size)
        w, h = im.size
        if im.getexif().get(274) in (5, 6, 7, 8):
            w, h = h, w
        return [[w, h]]


def stage_source(source: str, root: str) -> dict:
    asset = copy_asset(source, root)
    sizes = inspect_source(asset_path(root, asset))
    return {"asset": asset, "sizes": sizes, "source": source}


def stage_clipboard(root: str) -> dict:
    from PIL import ImageGrab
    data = ImageGrab.grabclipboard()
    if isinstance(data, list):
        return {"paths": [str(p) for p in data]}
    if not isinstance(data, Image.Image):
        raise ValueError("The clipboard does not contain an image or image files")
    try:
        check_pixels(*data.size)
        stream = io.BytesIO()
        data.save(stream, "PNG")
        asset = store_bytes(stream.getvalue(), "png", root)
        return {"asset": asset, "sizes": [list(data.size)], "source": "Clipboard"}
    finally:
        data.close()


def decode_item(data: dict, root: str, target_edge: int | None = None) -> Image.Image:
    item = Item.from_dict(data)
    path = asset_path(root, item.asset)
    if path.suffix == ".pdf":
        with fitz.open(path) as doc:
            if not 0 <= item.source_page < len(doc):
                raise ValueError("This PDF page does not exist")
            page = doc[item.source_page]
            full = page.rect
            crop = fitz.Rect(full.x0 + item.crop[0] * full.width, full.y0 + item.crop[1] * full.height, full.x0 + item.crop[2] * full.width, full.y0 + item.crop[3] * full.height)
            scale = min(8.0, (target_edge or max(item.original_size)) / max(1, crop.width, crop.height))
            check_pixels(max(1, math.ceil(crop.width * scale)), max(1, math.ceil(crop.height * scale)))
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=crop, alpha=False)
            out = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("RGBA")
            del pix
            draw = ImageDraw.Draw(out)
            for e in item.erasures:
                box = [(e[0] - item.crop[0]) / (item.crop[2] - item.crop[0]) * out.width, (e[1] - item.crop[1]) / (item.crop[3] - item.crop[1]) * out.height, (e[2] - item.crop[0]) / (item.crop[2] - item.crop[0]) * out.width, (e[3] - item.crop[1]) / (item.crop[3] - item.crop[1]) * out.height]
                draw.rectangle(box, fill="white")
            return out
    with Image.open(path) as source:
        check_pixels(*source.size)
        if target_edge:
            fraction = max(1e-4, min(item.crop[2] - item.crop[0], item.crop[3] - item.crop[1]))
            draft_edge = min(max(source.size), math.ceil(target_edge / fraction))
            source.draft("RGB", (draft_edge, draft_edge))
        oriented = ImageOps.exif_transpose(source)
        try:
            sw, sh = oriented.size
            box = (max(0, min(sw - 1, int(item.crop[0] * sw))), max(0, min(sh - 1, int(item.crop[1] * sh))), min(sw, max(1, math.ceil(item.crop[2] * sw))), min(sh, max(1, math.ceil(item.crop[3] * sh))))
            cropped = oriented.crop(box)
            if target_edge:
                cropped.thumbnail((target_edge, target_edge), Image.Resampling.LANCZOS)
            out = cropped.convert("RGBA")
            cropped.close()
        finally:
            if oriented is not source:
                oriented.close()
    draw = ImageDraw.Draw(out)
    for e in item.erasures:
        x1 = (e[0] - item.crop[0]) / (item.crop[2] - item.crop[0]) * out.width
        y1 = (e[1] - item.crop[1]) / (item.crop[3] - item.crop[1]) * out.height
        x2 = (e[2] - item.crop[0]) / (item.crop[2] - item.crop[0]) * out.width
        y2 = (e[3] - item.crop[1]) / (item.crop[3] - item.crop[1]) * out.height
        draw.rectangle((x1, y1, x2, y2), fill="white")
    return out


def preview_key(data: dict) -> str:
    values = {k: data[k] for k in ("asset", "source_page", "crop", "erasures")}
    values["edge"] = PREVIEW_EDGE
    return hashlib.sha256(json.dumps(values, sort_keys=True).encode()).hexdigest()


def make_preview(data: dict, root: str) -> str:
    ensure_root(root)
    path = Path(root) / "previews" / (preview_key(data) + ".png")
    if not path.exists():
        with decode_item(data, root, PREVIEW_EDGE) as image:
            image.thumbnail((PREVIEW_EDGE, PREVIEW_EDGE), Image.Resampling.LANCZOS)
            image.save(path, "PNG", compress_level=3)
            for edge in (128, 256, 512):
                if edge < PREVIEW_EDGE:
                    small = image.copy()
                    small.thumbnail((edge, edge), Image.Resampling.LANCZOS)
                    small.save(path.with_name(path.stem + f".{edge}.png"), "PNG", compress_level=3)
                    small.close()
    return str(path)


def trim_source(data: dict, root: str) -> list[float] | None:
    with decode_item(data, root) as image:
        alpha = image.getchannel("A")
        if alpha.getextrema()[0] < 255:
            mask = alpha.point(lambda p: 255 if p > 12 else 0)
        else:
            rgb = image.convert("RGB")
            corners = [rgb.getpixel(p) for p in [(0, 0), (rgb.width - 1, 0), (0, rgb.height - 1), (rgb.width - 1, rgb.height - 1)]]
            bg = max(corners, key=corners.count)
            flat = Image.new("RGB", rgb.size, bg)
            diff = ImageChops.difference(rgb, flat)
            bands = diff.split()
            mask = ImageChops.lighter(ImageChops.lighter(bands[0], bands[1]), bands[2]).point(lambda p: 255 if p > 12 else 0)
            rgb.close()
            flat.close()
            diff.close()
            for band in bands:
                band.close()
        bbox = mask.getbbox()
        mask.close()
        alpha.close()
        return [bbox[0] / image.width, bbox[1] / image.height, bbox[2] / image.width, bbox[3] / image.height] if bbox else None


def apply_trim(item: Item, fraction: list[float]) -> None:
    area = image_area(item)
    box = (area[0] + fraction[0] * (area[2] - area[0]), area[1] + fraction[1] * (area[3] - area[1]), area[0] + fraction[2] * (area[2] - area[0]), area[1] + fraction[3] * (area[3] - area[1]))
    old = item.crop
    item.crop = [old[0] + fraction[0] * (old[2] - old[0]), old[1] + fraction[1] * (old[3] - old[1]), old[0] + fraction[2] * (old[2] - old[0]), old[1] + fraction[3] * (old[3] - old[1])]
    label = item.x + item.label_offset_x, item.y + item.label_offset_y
    c = rotate(((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), item.center, item.rotation)
    item.w, item.h = max(1, box[2] - box[0]), max(1, box[3] - box[1])
    item.x, item.y = c[0] - item.w / 2, c[1] - item.h / 2
    item.label_offset_x, item.label_offset_y = label[0] - item.x, label[1] - item.y


def atomic_path(destination: str) -> tuple[Path, Path]:
    path = Path(destination).expanduser().resolve()
    if not path.parent.is_dir():
        raise ValueError("The destination folder does not exist")
    fd, name = tempfile.mkstemp(prefix="." + path.name + ".", suffix=".tmp", dir=path.parent)
    os.close(fd)
    return path, Path(name)


def save_project(snapshot: str, root: str, destination: str) -> None:
    document = json.loads(snapshot)
    target, temporary = atomic_path(destination)
    try:
        assets = sorted({i["asset"] for i in document["items"] if i["kind"] == "panel"})
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as archive:
            archive.writestr("manifest.json", snapshot)
            for asset in assets:
                path = asset_path(root, asset)
                if not path.is_file():
                    raise ValueError(f"Missing original image: {asset}")
                archive.write(path, "assets/" + asset)
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)


def load_project(source: str, root: str) -> str:
    ensure_root(root)
    path = Path(source)
    if path.stat().st_size > MAX_PROJECT_BYTES:
        raise ValueError("Project exceeds the 4 GB safety limit")
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 10001 or sum(i.file_size for i in infos) > MAX_PROJECT_BYTES:
                raise ValueError("Project exceeds the extraction safety limit")
            manifest_info = archive.getinfo("manifest.json")
            if manifest_info.file_size > 32 * 1024 * 1024:
                raise ValueError("Project metadata is too large")
            data = json.loads(archive.read(manifest_info))
            if data.get("version") != 2:
                raise ValueError("Unsupported FigPie project version")
            assets = {i.get("asset", "") for i in data.get("items", []) if i.get("kind") == "panel"}
            for asset in assets:
                out = asset_path(root, asset)
                info = archive.getinfo("assets/" + asset)
                if info.file_size > MAX_SOURCE_BYTES:
                    raise ValueError("An embedded asset exceeds the 1 GB safety limit")
                digest = hashlib.sha256()
                fd, name = tempfile.mkstemp(dir=out.parent)
                try:
                    with os.fdopen(fd, "wb") as target, archive.open(info) as src:
                        for block in iter(lambda: src.read(1024 * 1024), b""):
                            digest.update(block)
                            target.write(block)
                    if digest.hexdigest() != asset.split(".")[0]:
                        raise ValueError("An embedded image has a corrupt checksum")
                    os.replace(name, out)
                finally:
                    Path(name).unlink(missing_ok=True)
                inspect_source(out)
    else:
        if path.stat().st_size > 512 * 1024 * 1024:
            raise ValueError("Legacy JSON projects above 512 MB must be reduced before migrating")
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(data, dict) or not isinstance(data.get("items"), list):
            raise ValueError("This is not a FigPie project")
        settings = asdict(Settings())
        aliases = {"outer_margin": "margin", "auto_trim_on_export": "export_trim"}
        for key in list(settings):
            if key in data:
                settings[key] = data[key]
        for old, new in aliases.items():
            if old in data:
                settings[new] = data[old]
        items = []
        for raw in data["items"]:
            raw = dict(raw)
            if raw.get("kind") == "panel":
                encoded = raw.pop("image_b64", None)
                if encoded:
                    image_bytes = base64.b64decode(encoded, validate=True)
                    raw["asset"] = store_bytes(image_bytes, "png", root)
                    raw["original_size"] = inspect_source(asset_path(root, raw["asset"]))[0]
                    del image_bytes
                else:
                    original = raw.get("source_path", "")
                    if not original or not Path(original).is_file():
                        raise ValueError(f"Legacy panel {raw.get('id')} has no embedded image or usable source")
                    staged = stage_source(original, root)
                    raw["asset"] = staged["asset"]
                raw["image_fit"] = "stretch"
            items.append(Item.from_dict(raw).to_dict())
        data = {"version": 2, "settings": settings, "items": items, "next_id": max((i["id"] for i in items), default=0) + 1, "next_group_id": max((i.get("group_id") or 0 for i in items), default=0) + 1}
    if len(data.get("items", [])) > 5000:
        raise ValueError("Projects with more than 5000 objects are not supported")
    settings_keys = {f.name for f in fields(Settings)}
    settings = Settings(**{k: v for k, v in data["settings"].items() if k in settings_keys})
    if not all(math.isfinite(float(v)) and 1 <= float(v) <= 100000 for v in [settings.canvas_w, settings.canvas_h]):
        raise ValueError("Invalid canvas dimensions")
    items = [Item.from_dict(raw) for raw in data["items"]]
    if len({i.id for i in items}) != len(items):
        raise ValueError("Duplicate item identifiers in project")
    doc = Document()
    doc.items, doc.settings = items, settings
    doc.next_id = max(max((i.id for i in items), default=0) + 1, int(data.get("next_id", 1)))
    doc.next_group_id = max(max((i.group_id or 0 for i in items), default=0) + 1, int(data.get("next_group_id", 1)))
    return doc.snapshot()


FONT_FAMILIES = ["Arial", "Helvetica", "Times New Roman", "Calibri", "Courier New", "DejaVu Sans", "Segoe UI", "Cascadia Code"]


@lru_cache(maxsize=256)
def font(family: str, size: int, bold: bool = False, italic: bool = False):
    size = max(1, min(4000, int(size)))
    suffix = "bi" if bold and italic else "bd" if bold else "i" if italic else ""
    windows = {"Arial": "arial", "Helvetica": "arial", "Times New Roman": "times", "Calibri": "calibri", "Courier New": "cour", "Segoe UI": "segoeui"}
    serif = family == "Times New Roman"
    mono = family in ("Courier New", "Cascadia Code")
    generic = "LiberationSerif" if serif else "LiberationMono" if mono else "LiberationSans"
    style = "BoldItalic" if bold and italic else "Bold" if bold else "Italic" if italic else "Regular"
    deja = "DejaVuSerif" if serif else "DejaVuSansMono" if mono else "DejaVuSans"
    ds = "-BoldOblique" if bold and italic else "-Bold" if bold else "-Oblique" if italic else ""
    names = []
    if family in windows:
        names.append(windows[family] + suffix + ".ttf")
        if family == "Segoe UI":
            names = ["segoeuiz.ttf" if bold and italic else "segoeuib.ttf" if bold else "segoeuii.ttf" if italic else "segoeui.ttf"]
    names.extend([f"{family}.ttf", f"{family}.otf", f"{generic}-{style}.ttf", deja + ds + ".ttf"])
    roots = [Path(os.getenv("WINDIR", "C:/Windows")) / "Fonts", Path("/usr/share/fonts/truetype/liberation2"), Path("/usr/share/fonts/truetype/liberation"), Path("/usr/share/fonts/truetype/dejavu"), Path("/System/Library/Fonts"), Path("/Library/Fonts")]
    for name in names:
        for candidate in [name] + [str(root / name) for root in roots]:
            try:
                return ImageFont.truetype(candidate, size=size)
            except (OSError, ValueError):
                pass
    return ImageFont.load_default(size=size)


def color(value: str, alpha: int = 255):
    rgb = ImageColor.getrgb(value or "#000000")
    return tuple(rgb[:3]) + (max(0, min(255, int(alpha))),)


def text_extent(text: str, family: str, size: float, bold: bool = False, italic: bool = False) -> tuple[float, float]:
    f = font(family, round(size), bold, italic)
    b = f.getbbox(text or " ")
    return max(1, b[2] - b[0]), max(1, b[3] - b[1])


def label_bbox(item: Item) -> Box | None:
    if item.kind != "panel" or not item.show_label or not item.label or not item.include_in_labels or not item.visible:
        return None
    w, h = text_extent(item.label, item.label_font_family, item.label_font_size, True)
    x, y = item.x + item.label_offset_x, item.y + item.label_offset_y
    return x, y, x + w, y + h


def arrow_head(item: Item) -> list[tuple[float, float]]:
    a, b = item.points()[:2]
    ang = math.atan2(b[1] - a[1], b[0] - a[0])
    length = max(10, item.stroke_width * 4)
    return [b, (b[0] - length * math.cos(ang - math.pi / 7), b[1] - length * math.sin(ang - math.pi / 7)), (b[0] - length * math.cos(ang + math.pi / 7), b[1] - length * math.sin(ang + math.pi / 7))]


def visual_bbox(item: Item, include_label: bool = False) -> Box:
    b = item.bbox()
    if item.kind == "shape":
        if item.shape_type == "arrow":
            b = union([b, bounds(arrow_head(item))])
        pad = item.stroke_width / 2 + 2
        b = b[0] - pad, b[1] - pad, b[2] + pad, b[3] + pad
    elif item.kind == "panel":
        pad = item.border_width / 2
        b = b[0] - pad, b[1] - pad, b[2] + pad, b[3] + pad
    if include_label and label_bbox(item):
        b = union([b, label_bbox(item)])
    return b


def content_bbox(items: list[Item]) -> Box:
    return union(visual_bbox(i, True) for i in items if i.visible)


def styled_lines(item: Item) -> list[list[dict]]:
    chars = []
    for run in item.rich_runs or [{"text": item.text, "bold": item.bold, "italic": item.italic}]:
        for ch in run["text"]:
            f = font(item.font_family, round(item.font_size), run.get("bold", False), run.get("italic", False))
            chars.append({"text": ch, "bold": run.get("bold", False), "italic": run.get("italic", False), "width": f.getlength(ch)})
    lines, current, current_width = [], [], 0.0
    maximum = max(1, item.w - 12)
    for ch in chars:
        if ch["text"] == "\n":
            lines.append(current)
            current, current_width = [], 0.0
            continue
        if current and current_width + ch["width"] > maximum:
            spaces = [i for i, c in enumerate(current) if c["text"] == " "]
            if spaces:
                at = spaces[-1]
                lines.append(current[:at])
                current = current[at + 1:]
            else:
                lines.append(current)
                current = []
            current_width = sum(c["width"] for c in current)
        current.append(ch)
        current_width += ch["width"]
    lines.append(current)
    return lines


def text_height(item: Item) -> float:
    return 12 + len(styled_lines(item)) * item.font_size * item.line_spacing * 1.2


def draw_text_top(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: str, f, fill) -> None:
    b = f.getbbox(text or " ")
    draw.text((xy[0] - b[0], xy[1] - b[1]), text, font=f, fill=fill)


def text_bitmap(item: Item, scale: float) -> Image.Image:
    width, height = max(1, round(item.w * scale)), max(1, round(item.h * scale))
    check_pixels(width, height, MAX_OUTPUT_PIXELS)
    out = Image.new("RGBA", (width, height), color(item.background, 255) if item.background else (0, 0, 0, 0))
    draw = ImageDraw.Draw(out)
    if item.outline:
        draw.rectangle((0, 0, width - 1, height - 1), outline=color(item.outline), width=max(1, round(scale)))
    lines = styled_lines(item)
    y = 6.0
    for index, line in enumerate(lines):
        line_w = sum(ch["width"] for ch in line)
        x = (item.w - line_w) / 2 if item.align == "center" else item.w - 6 - line_w if item.align == "right" else 6.0
        spaces = sum(ch["text"] == " " for ch in line)
        extra = max(0, item.w - 12 - line_w) / spaces if item.align == "justify" and spaces and index < len(lines) - 1 else 0
        fonts = {(ch["bold"], ch["italic"]): font(item.font_family, round(item.font_size * scale), ch["bold"], ch["italic"]) for ch in line}
        ascent = max((f.getmetrics()[0] for f in fonts.values()), default=item.font_size * scale)
        baseline = y * scale + ascent
        for ch in line:
            f = fonts[(ch["bold"], ch["italic"])]
            draw.text((x * scale, baseline), ch["text"], font=f, fill=color(item.color), anchor="ls")
            x += ch["width"] + (extra if ch["text"] == " " else 0)
        y += item.font_size * item.line_spacing * 1.2
        if y > item.h:
            break
    return out


def place_rotated(out: Image.Image, bitmap: Image.Image, center: tuple[float, float], degrees: float, origin: tuple[float, float], scale: float) -> None:
    rotated = bitmap.rotate(-degrees, resample=Image.Resampling.BICUBIC, expand=True) if degrees % 360 else bitmap
    x = round((center[0] - origin[0]) * scale - rotated.width / 2)
    y = round((center[1] - origin[1]) * scale - rotated.height / 2)
    out.alpha_composite(rotated, (x, y))
    if rotated is not bitmap:
        rotated.close()


def render_object(item: Item, root: str, scale: float = 1, preview: bool = False) -> tuple[Image.Image, Box]:
    b = visual_bbox(item)
    width, height = max(1, math.ceil((b[2] - b[0]) * scale)), max(1, math.ceil((b[3] - b[1]) * scale))
    check_pixels(width, height, MAX_OUTPUT_PIXELS)
    out = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    origin = b[0], b[1]
    pt = lambda p: ((p[0] - b[0]) * scale, (p[1] - b[1]) * scale)
    if item.kind == "panel":
        w, h = max(1, round(item.w * scale)), max(1, round(item.h * scale))
        check_pixels(w, h, MAX_OUTPUT_PIXELS)
        with decode_item(item.to_dict(), root, max(w, h) * (2 if not preview else 1)) as source:
            frame = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            if item.image_fit == "cover":
                fitted = ImageOps.fit(source, (w, h), method=Image.Resampling.LANCZOS)
            elif item.image_fit == "stretch":
                fitted = source.resize((w, h), Image.Resampling.LANCZOS)
            else:
                fitted = ImageOps.contain(source, (w, h), Image.Resampling.LANCZOS)
            frame.alpha_composite(fitted, ((w - fitted.width) // 2, (h - fitted.height) // 2))
            fitted.close()
            if item.border_width:
                ImageDraw.Draw(frame).rectangle((0, 0, w - 1, h - 1), outline=color(item.border_color), width=max(1, round(item.border_width * scale)))
            place_rotated(out, frame, item.center, item.rotation, origin, scale)
            frame.close()
    elif item.kind == "text":
        frame = text_bitmap(item, scale)
        place_rotated(out, frame, item.center, item.rotation, origin, scale)
        frame.close()
    else:
        draw = ImageDraw.Draw(out)
        points = [pt(p) for p in item.points()]
        stroke = max(1, round(item.stroke_width * scale))
        ink = color(item.color)
        if item.is_line or item.is_stat:
            draw.line(points, fill=ink, width=stroke, joint="curve")
            if item.shape_type == "arrow":
                draw.polygon([pt(p) for p in arrow_head(item)], fill=ink)
        else:
            fill = color(item.fill or item.color, item.fill_alpha) if item.fill_alpha else None
            if fill:
                draw.polygon(points, fill=fill)
            if item.stroke_width:
                draw.line(points + [points[0]], fill=ink, width=stroke, joint="curve")
        if item.is_stat and item.text:
            tx, ty = item.stat_text_point()
            fw, fh = text_extent(item.text, item.font_family, item.font_size, True)
            f = font(item.font_family, round(item.font_size * scale), True)
            text = Image.new("RGBA", (max(1, math.ceil(fw * scale + 4)), max(1, math.ceil(fh * scale + 4))), (0, 0, 0, 0))
            draw_text_top(ImageDraw.Draw(text), (2, 2), item.text, f, ink)
            center = rotate((tx, ty - fh / 2), item.center, item.rotation)
            place_rotated(out, text, center, item.rotation, origin, scale)
            text.close()
    return out, b


def label_bitmap(item: Item, scale: float = 2) -> bytes:
    w, h = text_extent(item.label, item.label_font_family, item.label_font_size, True)
    image = Image.new("RGBA", (max(1, math.ceil(w * scale + 2)), max(1, math.ceil(h * scale + 2))), (0, 0, 0, 0))
    draw_text_top(ImageDraw.Draw(image), (0, 0), item.label, font(item.label_font_family, round(item.label_font_size * scale), True), color(item.label_color))
    buffer = io.BytesIO()
    image.save(buffer, "PNG")
    image.close()
    return buffer.getvalue()


def export_bounds(items: list[Item], settings: Settings) -> Box:
    if settings.export_trim and items:
        b = content_bbox(items)
        m = settings.margin
        return b[0] - m, b[1] - m, b[2] + m, b[3] + m
    return 0, 0, settings.canvas_w, settings.canvas_h


def render_document(snapshot: str, root: str) -> Image.Image:
    data = json.loads(snapshot)
    settings = Settings(**data["settings"])
    items = [Item.from_dict(d) for d in data["items"] if d.get("visible", True)]
    b = export_bounds(items, settings)
    scale = settings.export_scale
    w, h = max(1, math.ceil((b[2] - b[0]) * scale)), max(1, math.ceil((b[3] - b[1]) * scale))
    check_pixels(w, h, MAX_OUTPUT_PIXELS)
    try:
        import psutil
        if w * h * 20 > psutil.virtual_memory().available * 0.6:
            raise ValueError("This export is too large for currently available RAM. Reduce the export scale or canvas dimensions.")
    except ImportError:
        pass
    out = Image.new("RGBA", (w, h), color(settings.bg_color))
    for item in sorted(items, key=lambda i: i.z_index):
        if (Path(root) / "cancel-export").exists():
            out.close()
            raise InterruptedError("Export cancelled")
        vb = visual_bbox(item)
        if vb[2] >= b[0] and vb[0] <= b[2] and vb[3] >= b[1] and vb[1] <= b[3]:
            image, ib = render_object(item, root, scale)
            out.alpha_composite(image, (round((ib[0] - b[0]) * scale), round((ib[1] - b[1]) * scale)))
            image.close()
        lb = label_bbox(item)
        if lb:
            draw_text_top(ImageDraw.Draw(out), ((lb[0] - b[0]) * scale, (lb[1] - b[1]) * scale), item.label, font(item.label_font_family, round(item.label_font_size * scale), True), color(item.label_color))
    return out


def svg_shape(item: Item) -> str:
    points = " ".join(f"{p[0]:.4f},{p[1]:.4f}" for p in item.points())
    stroke = html.escape(item.color, quote=True)
    result = ""
    if item.is_line or item.is_stat:
        result = f'<polyline points="{points}" fill="none" stroke="{stroke}" stroke-width="{item.stroke_width}" stroke-linejoin="round" />'
        if item.shape_type == "arrow":
            hp = " ".join(f"{p[0]:.4f},{p[1]:.4f}" for p in arrow_head(item))
            result += f'<polygon points="{hp}" fill="{stroke}" />'
    else:
        fill = html.escape(item.fill or item.color, quote=True) if item.fill_alpha else "none"
        result = f'<polygon points="{points}" fill="{fill}" fill-opacity="{item.fill_alpha / 255:.5f}" stroke="{stroke}" stroke-width="{item.stroke_width}" />'
    if item.is_stat and item.text:
        tx, ty = item.stat_text_point()
        result += f'<text x="{tx}" y="{ty}" text-anchor="middle" font-family="{html.escape(item.font_family, quote=True)}" font-size="{item.font_size}" font-weight="bold" fill="{stroke}" transform="rotate({item.rotation} {item.center[0]} {item.center[1]})">{html.escape(item.text)}</text>'
    return result


def write_svg(snapshot: str, root: str, destination: str) -> None:
    data = json.loads(snapshot)
    settings = Settings(**data["settings"])
    items = [Item.from_dict(d) for d in data["items"] if d.get("visible", True)]
    b = export_bounds(items, settings)
    w, h = b[2] - b[0], b[3] - b[1]
    scale = settings.export_scale
    with open(destination, "w", encoding="utf-8") as out:
        out.write(f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{w * scale / settings.export_dpi * 25.4}mm" height="{h * scale / settings.export_dpi * 25.4}mm" viewBox="{b[0]} {b[1]} {w} {h}">\n')
        out.write(f'<rect x="{b[0]}" y="{b[1]}" width="{w}" height="{h}" fill="{html.escape(settings.bg_color, quote=True)}"/>\n')
        for item in sorted(items, key=lambda i: i.z_index):
            if (Path(root) / "cancel-export").exists():
                raise InterruptedError("Export cancelled")
            if item.kind == "shape":
                out.write(svg_shape(item))
            else:
                unrotated = copy.deepcopy(item)
                unrotated.rotation = 0
                bitmap, local = render_object(unrotated, root, scale)
                buffer = io.BytesIO()
                bitmap.save(buffer, "PNG")
                bitmap.close()
                encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
                out.write(f'<image x="{local[0]}" y="{local[1]}" width="{local[2] - local[0]}" height="{local[3] - local[1]}" preserveAspectRatio="none" transform="rotate({item.rotation} {item.center[0]} {item.center[1]})" xlink:href="data:image/png;base64,{encoded}"/>\n')
                del encoded
            lb = label_bbox(item)
            if lb:
                out.write(f'<text x="{lb[0]}" y="{lb[1] + lb[3] - lb[1]}" font-family="{html.escape(item.label_font_family, quote=True)}" font-size="{item.label_font_size}" font-weight="bold" fill="{html.escape(item.label_color, quote=True)}">{html.escape(item.label)}</text>\n')
        out.write('</svg>')


def export_file(snapshot: str, root: str, destination: str, fmt: str) -> None:
    data = json.loads(snapshot)
    settings = Settings(**data["settings"])
    target, temporary = atomic_path(destination)
    try:
        if fmt == "SVG":
            write_svg(snapshot, root, str(temporary))
        else:
            image = render_document(snapshot, root)
            try:
                if fmt == "PDF":
                    buffer = io.BytesIO()
                    image.save(buffer, "PNG")
                    with fitz.open() as doc:
                        p = doc.new_page(width=image.width / settings.export_dpi * 72, height=image.height / settings.export_dpi * 72)
                        p.insert_image(p.rect, stream=buffer.getvalue())
                        doc.save(str(temporary), deflate=True)
                elif fmt == "TIFF":
                    image.save(temporary, "TIFF", compression="tiff_deflate", dpi=(settings.export_dpi, settings.export_dpi))
                elif fmt == "PNG":
                    image.save(temporary, "PNG", dpi=(settings.export_dpi, settings.export_dpi))
                else:
                    raise ValueError("Unknown export format")
            finally:
                image.close()
        if (Path(root) / "cancel-export").exists():
            raise InterruptedError("Export cancelled")
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)


PALETTE = (
    ("Black", "#000000"), ("Charcoal", "#171923"), ("Grey", "#808080"),
    ("Silver", "#C0C0C0"), ("Light grey", "#E5E5E5"), ("White", "#FFFFFF"),
    ("Red", "#D73027"), ("Coral", "#F46D43"), ("Orange", "#E87818"),
    ("Amber", "#FFB300"), ("Yellow", "#F2C94C"), ("Cream", "#FFF4DD"),
    ("Dark green", "#1B7837"), ("Green", "#33A02C"), ("Mint", "#80CDC1"),
    ("Teal", "#008080"), ("Cyan", "#00A9CF"), ("Sky blue", "#56B4E9"),
    ("Navy", "#203864"), ("Blue", "#0072B2"), ("Purple", "#7B3294"),
    ("Pink", "#CC79A7"), ("Brown", "#8C510A"), ("Beige", "#D8B18A"),
)


def normalize_colour(value: str | None, optional: bool = False) -> str:
    if value is None or not str(value).strip():
        return "" if optional else "#000000"
    rgb = ImageColor.getrgb(str(value).strip())
    return "#%02X%02X%02X" % rgb[:3]


def colour_name(value: str | None, empty_label: str = "None") -> str:
    normalized = normalize_colour(value, optional=True)
    if not normalized:
        return empty_label
    return next((name for name, code in PALETTE if code == normalized), "Custom colour")


def rgb_colour(red: float, green: float, blue: float) -> str:
    values = [max(0, min(255, round(float(v)))) for v in (red, green, blue)]
    return "#%02X%02X%02X" % tuple(values)



def colour_hls(value: str | None) -> tuple[float, float, float]:
    rgb = ImageColor.getrgb(normalize_colour(value))[:3]
    hue, lightness, saturation = colorsys.rgb_to_hls(*(channel / 255 for channel in rgb))
    return hue * 360, lightness * 100, saturation * 100


def hls_colour(hue: float, lightness: float, saturation: float) -> str:
    numbers = tuple(float(value) for value in (hue, lightness, saturation))
    if not all(math.isfinite(value) for value in numbers):
        raise ValueError("Colour controls require finite values")
    hue, lightness, saturation = numbers
    rgb = colorsys.hls_to_rgb((hue % 360) / 360, clamp(lightness, 0, 100) / 100, clamp(saturation, 0, 100) / 100)
    return rgb_colour(*(channel * 255 for channel in rgb))


def readable_ink(value: str) -> str:
    r, g, b = ImageColor.getrgb(normalize_colour(value))[:3]
    return "#FFFFFF" if r * 0.299 + g * 0.587 + b * 0.114 < 145 else "#252018"


IMAGE_EXTENSIONS = frozenset({".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp", ".pdf"})
CLICK_TOOLS = frozenset({"compare", "symbol", "label-order", "label-align"})


def event_position(event, fallback=(0.0, 0.0)):
    point = getattr(event, "local_position", None)
    if point is not None:
        if isinstance(point, dict):
            result = float(point.get("x", 0)), float(point.get("y", 0))
        else:
            result = float(point.x), float(point.y)
    elif getattr(event, "local_x", None) is not None:
        result = float(event.local_x), float(event.local_y)
    else:
        result = fallback
    return result if all(math.isfinite(v) for v in result) else fallback


def wheel_delta(event):
    delta = getattr(event, "scroll_delta", None)
    if isinstance(delta, dict):
        result = float(delta.get("x", 0)), float(delta.get("y", 0))
    elif delta is not None:
        result = float(delta.x), float(delta.y)
    else:
        result = float(getattr(event, "scroll_delta_x", 0)), float(getattr(event, "scroll_delta_y", 0))
    return result if all(math.isfinite(v) for v in result) else (0.0, 0.0)


def wheel_zoom_factor(delta):
    return math.exp(max(-1.0, min(1.0, -float(delta) * 0.0018)))


def normalize_key(key):
    key = str(key).strip().lower().replace("arrow ", "arrow")
    return "space" if not key else key


def windows_modifiers():
    if os.name != "nt":
        return None
    try:
        import ctypes
        get = ctypes.WinDLL("user32", use_last_error=True).GetAsyncKeyState
        get.argtypes = [ctypes.c_int]
        get.restype = ctypes.c_short
        return {"control": bool(get(0x11) & 0x8000), "shift": bool(get(0x10) & 0x8000), "alt": bool(get(0x12) & 0x8000), "meta": bool((get(0x5B) | get(0x5C)) & 0x8000), "space": bool(get(0x20) & 0x8000)}
    except (OSError, AttributeError):
        return None


def dropped_paths(values):
    accepted, rejected, seen = [], [], set()
    for value in values:
        value = os.fspath(value).strip().strip('"')
        if value.startswith("file:"):
            parsed = urlparse(value)
            value = unquote(parsed.path)
            if parsed.netloc:
                value = "//" + parsed.netloc + value
            elif len(value) > 2 and value[0] == "/" and value[2] == ":":
                value = value[1:]
        path = Path(value).expanduser()
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTENSIONS:
            rejected.append(str(path))
            continue
        key = os.path.normcase(str(path.resolve()))
        if key not in seen:
            seen.add(key)
            accepted.append(str(path.resolve()))
    return accepted, rejected


class WindowsFileDrop:
    def __init__(self):
        self.events = queue.SimpleQueue()
        self._stop = threading.Event()
        self._thread = None
        self._region = (248.0, 116.0, 700.0, 650.0, 0.0)
        self._enabled = True
        self._hwnd = None
        self._host = None
        self._shown = False
        self._dropped = False
        self._scale = 1.0
        self._screen_box = None

    def set_region(self, left, top, width, height, page_width):
        self._region = tuple(float(v or 0) for v in (left, top, width, height, page_width))

    def set_enabled(self, enabled):
        self._enabled = bool(enabled)

    def start(self):
        if os.name != "nt" or self._thread is not None:
            return False
        self._thread = threading.Thread(target=self._run, name="FigPieFileDrop", daemon=True)
        self._thread.start()
        return True

    def stop(self):
        self._stop.set()

    def _run(self):
        try:
            self._native_loop()
        except Exception as exc:
            logging.exception("Windows file-drop receiver failed")
            self.events.put(("error", str(exc)))
        finally:
            self.events.put(("stopped", None))

    def _native_loop(self):
        from ctypes import wintypes as wt

        user = ctypes.WinDLL("user32", use_last_error=True)
        shell = ctypes.WinDLL("shell32", use_last_error=True)
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        gdi = ctypes.WinDLL("gdi32", use_last_error=True)
        pointer = ctypes.c_void_p
        result_t = ctypes.c_ssize_t
        wparam_t = ctypes.c_size_t
        lparam_t = ctypes.c_ssize_t
        wndproc_t = ctypes.WINFUNCTYPE(result_t, wt.HWND, wt.UINT, wparam_t, lparam_t)
        enum_t = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, lparam_t)

        class WindowClass(ctypes.Structure):
            _fields_ = [("style", wt.UINT), ("lpfnWndProc", wndproc_t), ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int), ("hInstance", wt.HINSTANCE), ("hIcon", wt.HICON), ("hCursor", wt.HANDLE), ("hbrBackground", wt.HBRUSH), ("lpszMenuName", wt.LPCWSTR), ("lpszClassName", wt.LPCWSTR)]

        class PaintStruct(ctypes.Structure):
            _fields_ = [("hdc", wt.HDC), ("fErase", wt.BOOL), ("rcPaint", wt.RECT), ("fRestore", wt.BOOL), ("fIncUpdate", wt.BOOL), ("rgbReserved", ctypes.c_byte * 32)]

        def bind(lib, name, arguments, result):
            function = getattr(lib, name)
            function.argtypes = arguments
            function.restype = result
            return function

        bind(user, "DefWindowProcW", [wt.HWND, wt.UINT, wparam_t, lparam_t], result_t)
        bind(user, "EnumWindows", [enum_t, lparam_t], wt.BOOL)
        bind(user, "GetWindowThreadProcessId", [wt.HWND, ctypes.POINTER(wt.DWORD)], wt.DWORD)
        bind(user, "GetWindowTextW", [wt.HWND, wt.LPWSTR, ctypes.c_int], ctypes.c_int)
        bind(user, "GetClassNameW", [wt.HWND, wt.LPWSTR, ctypes.c_int], ctypes.c_int)
        bind(user, "IsWindow", [wt.HWND], wt.BOOL)
        bind(user, "IsWindowVisible", [wt.HWND], wt.BOOL)
        bind(user, "IsIconic", [wt.HWND], wt.BOOL)
        bind(user, "GetAncestor", [wt.HWND, wt.UINT], wt.HWND)
        bind(user, "WindowFromPoint", [wt.POINT], wt.HWND)
        bind(user, "GetCursorPos", [ctypes.POINTER(wt.POINT)], wt.BOOL)
        bind(user, "GetClientRect", [wt.HWND, ctypes.POINTER(wt.RECT)], wt.BOOL)
        bind(user, "ClientToScreen", [wt.HWND, ctypes.POINTER(wt.POINT)], wt.BOOL)
        bind(user, "GetAsyncKeyState", [ctypes.c_int], ctypes.c_short)
        bind(user, "RegisterClassW", [ctypes.POINTER(WindowClass)], wt.ATOM)
        bind(user, "UnregisterClassW", [wt.LPCWSTR, wt.HINSTANCE], wt.BOOL)
        bind(user, "CreateWindowExW", [wt.DWORD, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.HWND, wt.HMENU, wt.HINSTANCE, pointer], wt.HWND)
        bind(user, "DestroyWindow", [wt.HWND], wt.BOOL)
        bind(user, "ShowWindow", [wt.HWND, ctypes.c_int], wt.BOOL)
        bind(user, "SetWindowPos", [wt.HWND, wt.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.UINT], wt.BOOL)
        bind(user, "SetLayeredWindowAttributes", [wt.HWND, wt.DWORD, ctypes.c_ubyte, wt.DWORD], wt.BOOL)
        bind(user, "PeekMessageW", [ctypes.POINTER(wt.MSG), wt.HWND, wt.UINT, wt.UINT, wt.UINT], wt.BOOL)
        bind(user, "TranslateMessage", [ctypes.POINTER(wt.MSG)], wt.BOOL)
        bind(user, "DispatchMessageW", [ctypes.POINTER(wt.MSG)], result_t)
        bind(user, "BeginPaint", [wt.HWND, ctypes.POINTER(PaintStruct)], wt.HDC)
        bind(user, "EndPaint", [wt.HWND, ctypes.POINTER(PaintStruct)], wt.BOOL)
        bind(user, "DrawTextW", [wt.HDC, wt.LPCWSTR, ctypes.c_int, ctypes.POINTER(wt.RECT), wt.UINT], ctypes.c_int)
        bind(user, "FillRect", [wt.HDC, ctypes.POINTER(wt.RECT), wt.HBRUSH], ctypes.c_int)
        bind(user, "InvalidateRect", [wt.HWND, pointer, wt.BOOL], wt.BOOL)
        bind(gdi, "CreateSolidBrush", [wt.DWORD], wt.HBRUSH)
        bind(gdi, "DeleteObject", [wt.HANDLE], wt.BOOL)
        bind(gdi, "SetTextColor", [wt.HDC, wt.DWORD], wt.DWORD)
        bind(gdi, "SetBkMode", [wt.HDC, ctypes.c_int], ctypes.c_int)
        bind(kernel, "GetModuleHandleW", [wt.LPCWSTR], wt.HMODULE)
        bind(shell, "DragAcceptFiles", [wt.HWND, wt.BOOL], None)
        bind(shell, "DragQueryFileW", [wt.HANDLE, wt.UINT, wt.LPWSTR, wt.UINT], wt.UINT)
        bind(shell, "DragQueryPoint", [wt.HANDLE, ctypes.POINTER(wt.POINT)], wt.BOOL)
        bind(shell, "DragFinish", [wt.HANDLE], None)
        old_dpi = None
        if hasattr(user, "SetThreadDpiAwarenessContext"):
            bind(user, "SetThreadDpiAwarenessContext", [pointer], pointer)
            old_dpi = user.SetThreadDpiAwarenessContext(pointer(-4))
        if hasattr(user, "GetDpiForWindow"):
            bind(user, "GetDpiForWindow", [wt.HWND], wt.UINT)
        instance = kernel.GetModuleHandleW(None)
        class_name = f"FigPieDrop_{os.getpid()}_{id(self)}"
        brush = gdi.CreateSolidBrush(0x001878E8)
        registered = False
        process_ids = {os.getpid()}
        last_scan = 0.0

        def find_host():
            nonlocal process_ids, last_scan
            if time.monotonic() - last_scan > 1:
                last_scan = time.monotonic()
                try:
                    import psutil
                    process_ids = {os.getpid(), *(p.pid for p in psutil.Process().children(recursive=True))}
                except Exception:
                    process_ids = {os.getpid()}
            found = []

            @enum_t
            def enum(hwnd, unused):
                if not user.IsWindowVisible(hwnd):
                    return True
                pid = wt.DWORD()
                user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value not in process_ids:
                    return True
                title = ctypes.create_unicode_buffer(512)
                name = ctypes.create_unicode_buffer(128)
                user.GetWindowTextW(hwnd, title, len(title))
                user.GetClassNameW(hwnd, name, len(name))
                if title.value.startswith("FigPie") and "FLUTTER" in name.value.upper():
                    found.append(hwnd)
                return True

            user.EnumWindows(enum, 0)
            return found[0] if found else None

        def hide():
            if self._shown and self._hwnd:
                user.ShowWindow(self._hwnd, 0)
            self._shown = False

        @wndproc_t
        def window_proc(hwnd, message, wparam, lparam):
            try:
                if message == 0x0233:
                    try:
                        count = shell.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)
                        if count > 4096:
                            raise ValueError("Drop fewer than 4,096 files at once.")
                        paths = []
                        for index in range(count):
                            length = shell.DragQueryFileW(wparam, index, None, 0)
                            buffer = ctypes.create_unicode_buffer(length + 1)
                            shell.DragQueryFileW(wparam, index, buffer, len(buffer))
                            paths.append(buffer.value)
                        position = wt.POINT()
                        shell.DragQueryPoint(wparam, ctypes.byref(position))
                        self._dropped = True
                        hide()
                        self.events.put(("files", paths, (position.x / self._scale, position.y / self._scale)))
                    finally:
                        shell.DragFinish(wparam)
                    return 0
                if message == 0x0021:
                    return 3
                if message == 0x0014:
                    return 1
                if message == 0x000F:
                    paint = PaintStruct()
                    dc = user.BeginPaint(hwnd, ctypes.byref(paint))
                    try:
                        rect = wt.RECT()
                        user.GetClientRect(hwnd, ctypes.byref(rect))
                        user.FillRect(dc, ctypes.byref(rect), brush)
                        gdi.SetTextColor(dc, 0x00FFFFFF)
                        gdi.SetBkMode(dc, 1)
                        user.DrawTextW(dc, "Drop images or PDFs onto your figure", -1, ctypes.byref(rect), 0x0001 | 0x0004 | 0x0020)
                    finally:
                        user.EndPaint(hwnd, ctypes.byref(paint))
                    return 0
            except Exception as exc:
                logging.exception("Windows file-drop message failed")
                self.events.put(("error", str(exc)))
                return 0
            return user.DefWindowProcW(hwnd, message, wparam, lparam)

        started = time.monotonic()
        warned = False
        last_button = False
        external_drag = False
        released_at = None
        try:
            while not self._stop.is_set():
                if not self._host or not user.IsWindow(self._host):
                    self._host = find_host()
                    if not self._host:
                        if not warned and time.monotonic() - started > 30:
                            warned = True
                            self.events.put(("error", "Could not locate this application's native window. Add plots remains available."))
                        self._stop.wait(0.2)
                        continue
                    if not registered:
                        wc = WindowClass(0, window_proc, 0, 0, instance, None, None, brush, None, class_name)
                        if not user.RegisterClassW(ctypes.byref(wc)):
                            raise ctypes.WinError(ctypes.get_last_error())
                        registered = True
                    ex_style = 0x00000080 | 0x08000000 | 0x00080000 | 0x00000010
                    self._hwnd = user.CreateWindowExW(ex_style, class_name, "FigPie file drop", 0x80000000, 0, 0, 1, 1, self._host, None, instance, None)
                    if not self._hwnd:
                        raise ctypes.WinError(ctypes.get_last_error())
                    if not user.SetLayeredWindowAttributes(self._hwnd, 0, 150, 2):
                        raise ctypes.WinError(ctypes.get_last_error())
                    shell.DragAcceptFiles(self._hwnd, True)
                    self._shown = False
                    self._screen_box = None
                    self.events.put(("ready", None))
                msg = wt.MSG()
                while user.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                    user.TranslateMessage(ctypes.byref(msg))
                    user.DispatchMessageW(ctypes.byref(msg))
                point = wt.POINT()
                user.GetCursorPos(ctypes.byref(point))
                at = user.WindowFromPoint(point)
                root_at = user.GetAncestor(at, 2) if at else None
                button = bool(user.GetAsyncKeyState(0x01) & 0x8000)
                if button and not last_button:
                    external_drag = root_at not in (self._host, self._hwnd)
                    self._dropped = False
                    released_at = None
                elif last_button and not button:
                    released_at = time.monotonic()
                last_button = button
                if not button and (released_at is None or time.monotonic() - released_at > 0.25):
                    hide()
                    external_drag = False
                eligible = button and external_drag and self._enabled and not self._dropped and not user.IsIconic(self._host) and user.IsWindowVisible(self._host)
                if eligible:
                    rect = wt.RECT()
                    origin = wt.POINT(0, 0)
                    user.GetClientRect(self._host, ctypes.byref(rect))
                    user.ClientToScreen(self._host, ctypes.byref(origin))
                    left, top, width, height, page_width = self._region
                    self._scale = (rect.right - rect.left) / page_width if page_width > 0 else (user.GetDpiForWindow(self._host) / 96 if hasattr(user, "GetDpiForWindow") else 1.0)
                    self._scale = max(0.5, min(5.0, self._scale))
                    box = (origin.x + round(left * self._scale), origin.y + round(top * self._scale), max(1, round(width * self._scale)), max(1, round(height * self._scale)))
                    inside = box[0] <= point.x < box[0] + box[2] and box[1] <= point.y < box[1] + box[3]
                    if inside and root_at in (self._host, self._hwnd):
                        if not self._shown or box != self._screen_box:
                            self._screen_box = box
                            user.SetWindowPos(self._hwnd, None, *box, 0x0010 | 0x0040 | 0x0200)
                            user.InvalidateRect(self._hwnd, None, True)
                            self._shown = True
                    else:
                        hide()
                elif not self._enabled or self._dropped or not external_drag:
                    hide()
                self._stop.wait(0.012)
        finally:
            hide()
            if self._hwnd and user.IsWindow(self._hwnd):
                shell.DragAcceptFiles(self._hwnd, False)
                user.DestroyWindow(self._hwnd)
            if registered:
                user.UnregisterClassW(class_name, instance)
            if brush:
                gdi.DeleteObject(brush)
            if old_dpi:
                user.SetThreadDpiAwarenessContext(old_dpi)


class ColourField(ft.Container):
    def __init__(self, app, label, value, width=None, optional=False, empty_label="None", on_change=None):
        super().__init__(width=width, padding=0)
        self._app = app
        self._label = label
        self._optional = optional
        self._empty_label = empty_label
        self._callback = on_change
        self._value = normalize_colour(value, optional)
        self._swatch = ft.Container(width=27, height=27, border_radius=6, border=ft.Border.all(1, "#CABBAD"))
        self._name = ft.Text(size=12, expand=True, max_lines=1)
        self._button = ft.Container(
            content=ft.Row([self._swatch, self._name, ft.Icon(ft.Icons.PALETTE_OUTLINED, size=19, color="#B6530A")], spacing=9),
            padding=10, border=ft.Border.all(1, "#D8C9BD"), border_radius=8,
            bgcolor="#FFFCF8", on_click=self._open, tooltip="Choose a colour",
        )
        self.content = ft.Column([ft.Text(label, size=11, color="#7C6B5B"), self._button], spacing=4)
        self._refresh()

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        self._value = normalize_colour(value, self._optional)
        self._refresh()

    @property
    def label(self):
        return self._label

    def _refresh(self):
        self._swatch.bgcolor = self._value or "#FFFFFF"
        self._swatch.content = None if self._value else ft.Icon(ft.Icons.BLOCK, size=18, color="#9B8170")
        self._name.value = colour_name(self._value, self._empty_label)

    async def _open(self, e):
        app = self._app
        if app.busy:
            app.notice("Finish or cancel the current operation before changing colours.")
            return
        original = self.value
        candidate = [original]
        components = list(colour_hls(original or "#808080"))
        if components[2] < 1e-9:
            components[0] = 30.0
        preview_label = ft.Text(colour_name(original, self._empty_label), size=16, weight=ft.FontWeight.BOLD)
        preview = ft.Container(content=preview_label, height=48, alignment=ft.Alignment.CENTER, border_radius=10, border=ft.Border.all(1, "#D8C9BD"))
        shade_value = ft.Text(size=12, text_align=ft.TextAlign.RIGHT)
        hue_value = ft.Text(width=48, size=12, text_align=ft.TextAlign.RIGHT)
        intensity_value = ft.Text(width=48, size=12, text_align=ft.TextAlign.RIGHT)
        palette_swatches = {}
        shade_slider = ft.Slider(min=0, max=100, divisions=1000, value=components[1], label="Lightness: {value}%", round=1, height=36, padding=ft.Padding.symmetric(horizontal=12), tooltip="Make the selected colour darker or lighter")
        hue_slider = ft.Slider(min=0, max=360, divisions=360, value=components[0], label="Hue: {value}°", height=40, expand=True, tooltip="Choose the colour family")
        intensity_slider = ft.Slider(min=0, max=100, divisions=1000, value=components[2], label="Colour intensity: {value}%", round=1, height=40, expand=True, tooltip="Grey at 0%; full colour at 100%")

        def update_preview(sync_values=False):
            selected = candidate[0]
            visible_colour = selected or "#808080"
            preview.bgcolor = selected or "#FFFFFF"
            preview_label.value = colour_name(selected, self._empty_label)
            preview_label.color = readable_ink(selected or "#FFFFFF")
            shade_value.value = f"{components[1]:.1f}% light" if selected else self._empty_label
            hue_value.value = f"{components[0]:.0f}°"
            intensity_value.value = f"{components[2]:.1f}%"
            if sync_values:
                shade_slider.value = components[1]
                hue_slider.value = components[0]
                intensity_slider.value = components[2]
            for slider in (shade_slider, hue_slider, intensity_slider):
                slider.active_color = visible_colour
                slider.thumb_color = visible_colour
                slider.inactive_color = "#B5A79B"
            for code, swatch in palette_swatches.items():
                swatch.border = ft.Border.all(3 if code == selected else 1, "#B6530A" if code == selected else "#D8C9BD")
                swatch.content = ft.Icon(ft.Icons.CHECK, size=18, color=readable_ink(code)) if code == selected else None

        def adjust(index, slider):
            async def callback(event):
                value = getattr(getattr(event, "control", None), "value", slider.value)
                if value is None:
                    return
                number = float(value)
                if not math.isfinite(number):
                    return
                components[index] = clamp(number, 0, 360 if index == 0 else 100)
                slider.value = components[index]
                candidate[0] = hls_colour(*components)
                update_preview()
                app.page.update()
            return callback

        shade_slider.on_change = adjust(1, shade_slider)
        shade_slider.on_change_end = shade_slider.on_change
        hue_slider.on_change = adjust(0, hue_slider)
        hue_slider.on_change_end = hue_slider.on_change
        intensity_slider.on_change = adjust(2, intensity_slider)
        intensity_slider.on_change_end = intensity_slider.on_change

        def pick(code):
            async def callback(event):
                candidate[0] = code
                if code:
                    hue, lightness, saturation = colour_hls(code)
                    components[:] = (hue if saturation > 1e-9 else components[0], lightness, saturation)
                update_preview(sync_values=True)
                app.page.update()
            return callback

        swatches = []
        for name, code in PALETTE:
            swatch = ft.Container(width=46, height=30, bgcolor=code, border_radius=6, alignment=ft.Alignment.CENTER, tooltip=name, on_click=pick(code))
            palette_swatches[code] = swatch
            swatches.append(ft.Column([swatch, ft.Text(name, size=10, width=62, text_align=ft.TextAlign.CENTER)], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=3, width=62))
        common_panel = ft.Row(swatches, wrap=True, spacing=6, run_spacing=8)
        custom_panel = ft.Column([
            ft.Text("Choose a colour family and its intensity. The bars follow your selected colour.", size=12, color="#7C6B5B"),
            ft.Row([ft.Text("Hue", width=62, size=12), hue_slider, hue_value], spacing=4),
            ft.Row([ft.Text("Intensity", width=62, size=12), intensity_slider, intensity_value], spacing=4),
            ft.Text("0% intensity is grey. Move Shade above towards Darker or Lighter to adjust the result.", size=12, color="#7C6B5B"),
        ], spacing=10, visible=False)
        common_button = ft.TextButton(content="Common colours", expand=True)
        custom_button = ft.TextButton(content="Custom colour", expand=True)

        def set_panel(custom):
            common_panel.visible = not custom
            custom_panel.visible = custom
            common_button.style = ft.ButtonStyle(bgcolor="#FFFCF8" if custom else "#FFE3C6", color="#7C6B5B" if custom else "#B6530A")
            custom_button.style = ft.ButtonStyle(bgcolor="#FFE3C6" if custom else "#FFFCF8", color="#B6530A" if custom else "#7C6B5B")

        async def show_common(event):
            set_panel(False)
            app.page.update()

        async def show_custom(event):
            set_panel(True)
            app.page.update()

        common_button.on_click = show_common
        custom_button.on_click = show_custom
        set_panel(False)
        shade_box = ft.Container(content=ft.Column([
            ft.Row([ft.Text("Shade", size=13, weight=ft.FontWeight.BOLD), shade_value], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            shade_slider,
            ft.Row([ft.Text("Darker", size=11, color="#7C6B5B"), ft.Text("Lighter", size=11, color="#7C6B5B")], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        ], spacing=0), padding=ft.Padding.symmetric(horizontal=10, vertical=8), border_radius=8, bgcolor="#F5EEE6", border=ft.Border.all(1, "#D8C9BD"))
        controls = [preview, shade_box, ft.Row([common_button, custom_button], spacing=8), common_panel, custom_panel]
        if self._optional:
            controls.append(app.button(self._empty_label, pick(""), "BLOCK"))
        page_height = float(getattr(app.page, "height", None) or 720)
        page_width = float(getattr(app.page, "width", None) or 1080)
        body = ft.Container(width=min(430, max(200, page_width - 96)), height=min(480 if self._optional else 440, max(180, page_height - 210)), content=ft.Column(controls, spacing=10, expand=True, scroll=ft.ScrollMode.ALWAYS))
        update_preview(sync_values=True)
        result = await app.choose(self.label, body, [(False, "Cancel", False), (True, "Use colour", True)])
        if result is True:
            self.value = candidate[0]
            if self._callback:
                outcome = self._callback(self.value)
                if inspect.isawaitable(outcome):
                    await outcome
            app.page.update()


ACCENT = "#710EC2"
INK = "#30251F"
MUTED = "#806E60"
BORDER = "#E9DCCD"
WORKSPACE = "#F3EEE7"
ANCHOR = "#F19022"
GUIDE = "#D2691E"


class ByteCache:
    def __init__(self, maximum: int = 32 * 1024 * 1024) -> None:
        self.maximum = maximum
        self.size = 0
        self.values = OrderedDict()

    def obtain(self, key, build):
        if key in self.values:
            result = self.values.pop(key)
            self.values[key] = result
            return result
        result = build()
        self.values[key] = result
        self.size += len(result)
        while (self.size > self.maximum or len(self.values) > 512) and len(self.values) > 1:
            _, removed = self.values.popitem(last=False)
            self.size -= len(removed)
        return result

    def clear(self):
        self.values.clear()
        self.size = 0


class FigPieApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.doc = Document()
        self.root = tempfile.mkdtemp(prefix="figpie-")
        ensure_root(self.root)
        self.pool = self.create_pool()
        self.worker_lock = asyncio.Lock()
        self.busy = False
        self.cancel_requested = False
        self.closed = False
        self.path: str | None = None
        self.generation = 0
        self.dialog_depth = 0
        self.typing = False
        self.keys: set[str] = set()
        self.mode = "select"
        self.tool_buttons = {}
        self.previews: dict[str, str] = {}
        self.byte_cache = ByteCache()
        self.visual_controls = {}
        self.label_controls = {}
        self.zoom = 0.6
        self.pan = [30.0, 30.0]
        self.view_w, self.view_h = 700.0, 650.0
        self.fitted = False
        self.left_tab = "Build"
        self.right_tab = "Properties"
        self.drag = None
        self.gesture_last = None
        self.guides = []
        self.marquee = None
        self.stat_first = None
        self.stat_hover = None
        self.stat_controls = None
        self._stat_session = None
        self.file_drop = WindowsFileDrop()
        self.drop_status = "Starting Windows file drop…" if os.name == "nt" else "Add plots or paste files to import."
        self._canvas_origin = (248.0, 116.0)
        self.sequence_ids = []
        self.click_align_anchor = None
        self.click_align_axis = "y"
        self.last_render = 0.0
        self.last_memory = 0.0
        self.fit_on_resize = False
        self._focus_nonce = 0
        self.build()

    @staticmethod
    def create_pool():
        return ProcessPoolExecutor(max_workers=1, mp_context=multiprocessing.get_context("spawn"))

    async def worker(self, function, *args):
        async with self.worker_lock:
            try:
                return await asyncio.get_running_loop().run_in_executor(self.pool, functools.partial(function, *args))
            except BrokenProcessPool as exc:
                self.pool.shutdown(wait=False, cancel_futures=True)
                self.pool = self.create_pool()
                raise RuntimeError("The image worker stopped unexpectedly. Your canvas is intact. The worker has been restarted; try a smaller source image.") from exc

    def icon(self, name):
        return getattr(ft.Icons, name, ft.Icons.CIRCLE)

    def handler(self, function, *args, edit=False):
        async def callback(event=None):
            if self.closed:
                return
            if edit and self.busy:
                self.notice("Finish or cancel the current operation before editing.")
                return
            before = self.doc.snapshot() if edit else None
            try:
                target = getattr(self.doc, function.__name__) if isinstance(getattr(function, "__self__", None), Document) else function
                result = target(*args)
                if inspect.isawaitable(result):
                    await result
                if edit and before != self.doc.snapshot():
                    self.doc.checkpoint(before)
                self.refresh()
            except Exception as exc:
                if edit and before is not None:
                    self.doc.restore(before)
                logging.exception("FigPie action failed")
                self.notice(str(exc), error=True)
                self.refresh()
        return callback

    def button(self, label, callback, icon=None, primary=False, width=None, tooltip=None):
        cls = ft.Button if primary else ft.TextButton
        return cls(content=label, icon=self.icon(icon) if icon else None, on_click=callback, width=width, tooltip=tooltip)

    def ibutton(self, name, tooltip, callback):
        return ft.IconButton(icon=self.icon(name), tooltip=tooltip, on_click=callback, icon_size=20)

    def field(self, label, value, width=None, multiline=False):
        return ft.TextField(label=label, value=str(value), width=width, text_size=13, dense=True, multiline=multiline, min_lines=3 if multiline else 1, max_lines=5 if multiline else 1, on_focus=self.on_input_focus, on_blur=self.on_input_blur)

    def colour_field(self, label, value, width=None, optional=False, empty_label="None"):
        return ColourField(self, label, value, width=width, optional=optional, empty_label=empty_label)

    def dropdown(self, label, value, options, width=None):
        if value not in {v for v, _ in options}:
            options = list(options) + [(value, value)]
        return ft.Dropdown(label=label, value=value, options=[ft.DropdownOption(key=v, text=t) for v, t in options], width=width, text_size=13, dense=True, on_focus=self.on_input_focus, on_blur=self.on_input_blur)

    def heading(self, text, subtitle=None):
        controls = [ft.Text(text, size=12, weight=ft.FontWeight.BOLD, color=INK)]
        if subtitle:
            controls.append(ft.Text(subtitle, size=11, color=MUTED))
        return ft.Container(content=ft.Column(controls, spacing=3), padding=ft.Padding.only(top=12, bottom=5))

    def row(self, *controls):
        return ft.Row(list(controls), spacing=4, wrap=True)

    def build(self):
        p = self.page
        p.title = "FigPie"
        p.theme_mode = ft.ThemeMode.LIGHT
        p.theme = ft.Theme(color_scheme_seed=ACCENT, font_family="Segoe UI", color_scheme=ft.ColorScheme(
            primary=ACCENT, on_primary="#FFFFFF", primary_container="#FFE3C6", on_primary_container="#522500",
            secondary="#A56B28", on_secondary="#FFFFFF", secondary_container="#FFEACC", on_secondary_container="#442C10",
            tertiary="#93623A", on_tertiary="#FFFFFF", tertiary_container="#FBE3CB", on_tertiary_container="#3D2717",
            surface="#FFFCF8", on_surface=INK, on_surface_variant=MUTED, outline="#AE9580", outline_variant=BORDER,
            surface_tint=ACCENT, surface_container_low="#FFF6EC", surface_container="#FAEFE3", surface_container_high="#F2E4D5",
        ))
        p.padding = 0
        p.bgcolor = WORKSPACE
        p.window.width, p.window.height = 1560, 980
        p.window.min_width, p.window.min_height = 1080, 720
        p.window.prevent_close = True
        p.window.on_event = self.on_window_event
        self.picker = ft.FilePicker()
        p.services.append(self.picker)
        self.title_text = ft.Text("Untitled figure", size=13, color=MUTED, max_lines=1)
        self.status = ft.Text("Import plots to start. Nothing is rearranged automatically.", size=12, color=MUTED, expand=True, max_lines=1)
        self.memory_text = ft.Text("", size=11, color=MUTED)
        self.zoom_text = ft.Text("60%", size=12, width=48, text_align=ft.TextAlign.CENTER)
        self.mode_text = ft.Text("SELECT", size=11, color=ACCENT, weight=ft.FontWeight.BOLD)
        self.progress = ft.ProgressBar(value=0, color=ACCENT, bgcolor=BORDER, height=3, visible=False)
        self.cancel_button = self.button("Cancel operation", self.handler(self.cancel_operation), "CLOSE")
        self.cancel_button.visible = False
        self.header = ft.Container(height=68, padding=ft.Padding.symmetric(horizontal=18, vertical=8), bgcolor="#FFFFFF", border=ft.Border.only(bottom=ft.BorderSide(1, BORDER)), content=ft.Row([
            ft.Icon(self.icon("PIE_CHART_ROUNDED"), color=ACCENT, size=30),
            ft.Column([ft.Text("FigPie", size=22, weight=ft.FontWeight.BOLD, color=INK), ft.Text("FIGURE STUDIO", size=9, color=MUTED)], spacing=0),
            ft.Container(width=12), self.title_text, ft.Container(expand=True),
            self.button("New", self.handler(self.new_document), "ADD"),
            self.button("Open", self.handler(self.open_document), "FOLDER_OPEN"),
            self.button("Save", self.handler(self.save_document), "SAVE_OUTLINED"),
            self.button("Export", self.handler(self.export_dialog), "IOS_SHARE", primary=True),
            self.ibutton("HELP_OUTLINE", "Help and keyboard shortcuts", self.handler(self.help_dialog))
        ], spacing=5))
        self.undo_button = self.ibutton("UNDO", "Undo · Ctrl+Z", self.handler(self.undo))
        self.redo_button = self.ibutton("REDO", "Redo · Ctrl+Y", self.handler(self.redo))
        self.toolbar = ft.Container(height=48, bgcolor="#FFFFFF", padding=ft.Padding.symmetric(horizontal=8, vertical=2), border=ft.Border.only(bottom=ft.BorderSide(1, BORDER)), content=ft.Row([
            self.button("Add plots", self.handler(self.pick_import), "ADD_PHOTO_ALTERNATE_OUTLINED", primary=True),
            self.ibutton("CONTENT_PASTE", "Paste image or files · Ctrl+V", self.handler(self.paste)),
            self.undo_button, self.redo_button,
            self.ibutton("CONTENT_COPY", "Duplicate · Ctrl+D", self.handler(self.doc.duplicate, edit=True)),
            self.ibutton("DELETE_OUTLINE", "Delete selection", self.handler(self.doc.delete, edit=True)),
            self.ibutton("GROUP_WORK_OUTLINED", "Group · Ctrl+G", self.handler(self.doc.group, edit=True)),
            self.ibutton("GROUP_OFF", "Ungroup · Ctrl+Shift+G", self.handler(self.doc.ungroup, edit=True)),
            ft.Container(expand=True), self.mode_text,
            self.ibutton("REMOVE", "Zoom out · Ctrl+−", self.handler(self.zoom_by, 1 / 1.15)), self.zoom_text,
            self.ibutton("ADD", "Zoom in · Ctrl++", self.handler(self.zoom_by, 1.15)),
            self.button("100%", self.handler(self.actual_size)),
            self.button("Fit", self.handler(self.fit_view), "FIT_SCREEN")
        ], spacing=2))
        self.left_content = ft.Column(spacing=5, scroll=ft.ScrollMode.AUTO, expand=True)
        self.right_content = ft.Column(spacing=7, scroll=ft.ScrollMode.AUTO, expand=True)
        left_tabs = self.row(*[self.button(t, self.handler(self.set_left_tab, t)) for t in ("Build", "Arrange", "Labels")])
        right_tabs = self.row(*[self.button(t, self.handler(self.set_right_tab, t)) for t in ("Properties", "Layers")])
        self.left_panel = ft.Container(width=248, bgcolor="#FFFFFF", padding=12, border=ft.Border.only(right=ft.BorderSide(1, BORDER)), content=ft.Column([left_tabs, ft.Divider(height=1, color=BORDER), self.left_content], spacing=5, expand=True))
        self.right_panel = ft.Container(width=296, bgcolor="#FFFFFF", padding=12, border=ft.Border.only(left=ft.BorderSide(1, BORDER)), content=ft.Column([right_tabs, ft.Divider(height=1, color=BORDER), self.right_content], spacing=5, expand=True))
        self.background_canvas = cv.Canvas(shapes=[], expand=True, on_resize=self.on_view_resize)
        self.overlay_canvas = cv.Canvas(shapes=[], expand=True)
        self.stack = ft.Stack(controls=[self.background_canvas, self.overlay_canvas], expand=True, fit=ft.StackFit.EXPAND, clip_behavior=ft.ClipBehavior.HARD_EDGE)
        self.gesture = ft.GestureDetector(content=self.stack, expand=True, drag_interval=25, on_tap_down=self.on_tap_down, on_tap_up=self.on_tap_up, on_pan_start=self.on_pan_start, on_pan_update=self.on_pan_update, on_pan_end=self.on_pan_end, on_pan_cancel=self.on_pan_cancel, on_secondary_tap_down=self.on_right_click, on_double_tap_down=self.on_double_click, on_scroll=self.on_scroll, on_hover=self.on_hover, on_enter=self.on_canvas_enter, on_exit=self.on_canvas_exit, hover_interval=35)
        self.context = ft.ContextMenu(content=self.gesture, expand=True, items=[])
        self.center_panel = ft.Container(content=self.context, expand=True, bgcolor=WORKSPACE)
        bottom = ft.Container(height=36, bgcolor="#FFFFFF", padding=ft.Padding.symmetric(horizontal=14, vertical=5), border=ft.Border.only(top=ft.BorderSide(1, BORDER)), content=ft.Row([self.status, self.cancel_button, self.memory_text, ft.Text("Hamid Taghipour · UiT", size=10, color=MUTED)], spacing=12))
        content = ft.Column([self.header, self.toolbar, self.progress, ft.Row([self.left_panel, self.center_panel, self.right_panel], spacing=0, expand=True, vertical_alignment=ft.CrossAxisAlignment.STRETCH), bottom], spacing=0, expand=True)
        self.keyboard = ft.KeyboardListener(content=content, autofocus=True, on_key_down=self.on_key_down, on_key_up=self.on_key_up, on_key_repeat=self.on_key_repeat)
        p.add(ft.Container(content=self.keyboard, expand=True))
        self.build_left()
        self.build_right()
        self.refresh()

    async def on_input_focus(self, e):
        self.typing = True
        self.keys.clear()

    async def on_input_blur(self, e):
        self.typing = False
        self.keys.clear()

    async def focus_canvas(self):
        self.typing = False
        await self.keyboard.focus()

    def set_left_tab(self, tab):
        self.left_tab = tab
        self.build_left()

    def set_right_tab(self, tab):
        self.right_tab = tab
        self.build_right()

    def build_left(self):
        s = self.doc.settings
        self.tool_buttons = {}
        controls = []
        if self.left_tab == "Build":
            controls += [self.heading("TOOLS", "Click a tool, then draw on the canvas.")]
            specifications = [("select", "Select", "NEAR_ME_OUTLINED"), ("pan", "Pan", "PAN_TOOL_OUTLINED"), ("crop", "Crop", "CROP"), ("erase", "Erase", "AUTO_FIX_NORMAL")]
            tool_controls = []
            for mode, label, icon in specifications:
                button = self.button(label, self.handler(self.set_mode, mode), icon, width=106)
                self.tool_buttons[mode] = button
                tool_controls.append(button)
            controls += [self.row(*tool_controls), self.row(self.button("Trim image", self.handler(self.trim_selected), "CONTENT_CUT"), self.button("Reset edits", self.handler(self.reset_image_edits, edit=True))), self.heading("INSERT")]
            controls += [self.row(self.button("Text", self.handler(self.add_text), "TEXT_FIELDS"), self.button("Caption", self.handler(self.add_text, True), "NOTES"))]
            shape_buttons = []
            for kind, label, icon in [("rectangle", "Rectangle", "CROP_SQUARE"), ("ellipse", "Ellipse", "RADIO_BUTTON_UNCHECKED"), ("line", "Line", "REMOVE"), ("arrow", "Arrow", "ARROW_RIGHT_ALT"), ("arrow_head", "Arrowhead", "PLAY_ARROW_OUTLINED"), ("highlight_bar", "Highlight", "HIGHLIGHT_ALT")]:
                button = self.button(label, self.handler(self.set_mode, kind), icon, width=106)
                self.tool_buttons[kind] = button
                shape_buttons.append(button)
            controls += [self.row(*shape_buttons), self.heading("STATISTICS", "Bracket: click the two data positions.")]
            if self.stat_controls is None:
                self.stat_symbol = self.field("Symbol / text", "***", 100)
                self.stat_size = self.field("Text size", 28, 105)
                self.stat_height = self.field("Height", 28, 100)
                self.stat_gap = self.field("Text gap", 10, 105)
                self.stat_width = self.field("Line width", 2, 100)
                self.stat_style = self.dropdown("Comparison", "bracket", [("bracket", "Bracket line"), ("line", "Straight line")], 217)
                self.stat_color = self.colour_field("Colour", "#000000", 217)
                self.stat_center = ft.Checkbox(label="Center symbols above line", value=True)
                self.stat_top = ft.Checkbox(label="Level above both clicks", value=True)
                self.stat_repeat = ft.Checkbox(label="Keep tool active", value=False)
                self.stat_hint = ft.Text("Choose Bracket, then click two data points.", size=11, color=ACCENT)
                self.stat_controls = [self.row(*[self.button(t, self.handler(self.set_stat_symbol, t), width=49) for t in ("*", "**", "***", "ns")]), self.row(self.stat_symbol, self.stat_size), self.row(self.stat_height, self.stat_gap), self.stat_width, self.stat_style, self.stat_color, self.stat_center, self.stat_top, self.stat_repeat, ft.Container(content=self.stat_hint, bgcolor="#FFF1E1", padding=9, border_radius=7)]
            controls += self.stat_controls
            for mode, label in [("symbol", "Symbol"), ("compare", "Bracket")]:
                self.tool_buttons[mode] = self.button(label, self.handler(self.set_mode, mode), width=106)
            controls += [self.row(self.tool_buttons["symbol"], self.tool_buttons["compare"]), self.button("Remove last comparison", self.handler(self.remove_last_comparison, edit=True)), self.heading("CANVAS")]
            self.canvas_unit = self.dropdown("Units", "px", [("px", "Pixels"), ("cm", "Centimetres"), ("in", "Inches")], 217)
            self.canvas_unit.on_select = self.on_canvas_unit_change
            self.canvas_width = self.field("Width", self.number(s.canvas_w), 104)
            self.canvas_height = self.field("Height", self.number(s.canvas_h), 104)
            self.canvas_color = self.colour_field("Background", s.bg_color, 217)
            self.drop_status_text = ft.Text(self.drop_status, size=11, color=MUTED)
            controls += [self.heading("FILE DROP"), self.drop_status_text, self.canvas_unit, self.row(self.canvas_width, self.canvas_height), self.canvas_color, self.button("Apply canvas", self.handler(self.apply_canvas, edit=True)), self.button("Fit canvas to content", self.handler(self.crop_canvas, edit=True))]
        elif self.left_tab == "Arrange":
            controls += [self.heading("SNAPPING", "V guides snap X; H guides snap Y.")]
            for attr, label in [("snap_x", "Vertical guides · snap X"), ("snap_y", "Horizontal guides · snap Y"), ("snap_gap", "Match gaps / equal spacing"), ("snap_grid", "Snap to grid"), ("show_grid", "Show grid")]:
                c = ft.Checkbox(label=label, value=getattr(s, attr))
                c.on_change = self.setting_event(attr)
                controls.append(c)
            controls += [ft.Text("Hold Alt to bypass snapping. Shift constrains movement.", size=11, color=MUTED), self.heading("SPACING")]
            self.margin_field = self.field("Margin", s.margin, 67)
            self.gap_x_field = self.field("Gap X", s.gap_x, 67)
            self.gap_y_field = self.field("Gap Y", s.gap_y, 67)
            self.grid_step_field = self.field("Grid step", s.grid_step, 100)
            controls += [self.row(self.margin_field, self.gap_x_field, self.gap_y_field), self.grid_step_field, self.button("Apply spacing", self.handler(self.apply_spacing, edit=True)), self.heading("ALIGN TO ANCHOR", "Last selected = orange reference.")]
            specs = [("left", "FORMAT_ALIGN_LEFT", "Left edges"), ("center_x", "ALIGN_HORIZONTAL_CENTER", "Horizontal centers"), ("right", "FORMAT_ALIGN_RIGHT", "Right edges"), ("top", "VERTICAL_ALIGN_TOP", "Top edges"), ("center_y", "VERTICAL_ALIGN_CENTER", "Vertical centers"), ("bottom", "VERTICAL_ALIGN_BOTTOM", "Bottom edges")]
            controls += [self.row(*[self.ibutton(icon, tooltip, self.handler(self.doc.align, where, edit=True)) for where, icon, tooltip in specs]), self.row(self.button("Same W", self.handler(self.doc.same_size, "x", edit=True)), self.button("Same H", self.handler(self.doc.same_size, "y", edit=True))), self.heading("AUTO H / V", "Preserves sizes, labels and group structure.")]
            controls += [self.row(self.button("Auto H", self.handler(self.doc.distribute, "x", True, False, edit=True), "HORIZONTAL_DISTRIBUTE"), self.button("Auto V", self.handler(self.doc.distribute, "y", True, False, edit=True), "VERTICAL_DISTRIBUTE")), self.row(self.button("Gap H", self.handler(self.doc.distribute, "x", False, False, edit=True)), self.button("Gap V", self.handler(self.doc.distribute, "y", False, False, edit=True))), self.row(self.button("Equal H", self.handler(self.doc.distribute, "x", False, True, edit=True)), self.button("Equal V", self.handler(self.doc.distribute, "y", False, True, edit=True))), self.heading("GRID LAYOUT", "Selection, or all panels when none selected.")]
            self.rows_field = self.field("Rows", 2, 104)
            self.cols_field = self.field("Columns", 2, 104)
            controls += [self.row(self.rows_field, self.cols_field), self.row(self.button("Custom grid", self.handler(self.custom_grid, edit=True)), self.button("Auto grid", self.handler(self.doc.grid, edit=True))), self.button("Resolve overlaps", self.handler(self.doc.resolve_overlaps, edit=True)), self.heading("GROUPS"), self.row(self.button("Group", self.handler(self.doc.group, edit=True)), self.button("Ungroup", self.handler(self.doc.ungroup, edit=True))), self.button("Select all", self.handler(self.select_all))]
        else:
            controls += [self.heading("PANEL LABELS", "Edit or drag labels")]
            self.label_size_field = self.field("Size", s.default_label_size, 104)
            self.label_family_field = self.dropdown("Font", s.default_font_family, [(f, f) for f in FONT_FAMILIES], 217)
            self.label_x_field = self.field("Offset X", s.default_label_offset_x, 104)
            self.label_y_field = self.field("Offset Y", s.default_label_offset_y, 104)
            self.label_order_check = ft.Checkbox(label="Generate in import order", value=s.keep_label_order, on_change=self.setting_event("keep_label_order"))
            controls += [self.row(self.label_size_field), self.label_family_field, self.row(self.label_x_field, self.label_y_field), self.label_order_check, self.row(self.button("Generate A–Z", self.handler(self.doc.generate_labels, edit=True)), self.button("Show / hide", self.handler(self.toggle_labels, edit=True))), self.row(self.button("Style all", self.handler(self.apply_label_style, False, edit=True)), self.button("Style selected", self.handler(self.apply_label_style, True, edit=True)))]
            self.label_position = self.dropdown("Position", "top-in", [(x, x.replace("-", " ").title()) for x in ["top-in", "top-out", "left-in", "left-out", "bottom-in", "bottom-out", "center"]], 217)
            controls += [self.label_position, self.row(self.button("Position all", self.handler(self.position_labels, False, edit=True)), self.button("Selected", self.handler(self.position_labels, True, edit=True))), self.heading("ALIGN LABELS ONLY", "Select labels/panels; choose the anchor last."), self.row(self.button("Same X", self.handler(self.doc.align_labels, "x", edit=True)), self.button("Same Y", self.handler(self.doc.align_labels, "y", edit=True)), self.button("XY", self.handler(self.doc.align_labels, "xy", edit=True))), self.row(self.button("Include", self.handler(self.include_labels, True, edit=True)), self.button("Exclude", self.handler(self.include_labels, False, edit=True))), self.heading("CLICK TOOLS")]
            self.label_axis_field = self.dropdown("Click-align axis", "y", [("x", "Same X"), ("y", "Same Y"), ("xy", "Same X and Y")], 217)
            controls += [self.label_axis_field]
            for mode, label in [("label-order", "Click label order"), ("label-align", "Click-align labels")]:
                button = self.button(label, self.handler(self.set_mode, mode))
                self.tool_buttons[mode] = button
                controls.append(button)
            controls += [ft.Text("Click order: click panels A, B, C… then Esc. Click-align: first label is the reference; subsequent labels move to it.", size=11, color=MUTED)]
        self.left_content.controls = controls
        self.update_modes()

    def setting_event(self, attr):
        async def event(e):
            if self.busy:
                e.control.value = getattr(self.doc.settings, attr)
            else:
                before = self.doc.snapshot()
                setattr(self.doc.settings, attr, bool(e.control.value))
                self.doc.checkpoint(before)
            self.refresh(inspector=False)
        return event

    def build_right(self):
        if self.right_tab == "Layers":
            self.build_layers()
            return
        item = self.doc.get(self.doc.anchor)
        if not item or not self.doc.selected:
            self.right_content.controls = [self.heading("YOUR FIGURE", "Import plots, then select an object to edit it."), ft.Container(height=12), ft.Icon(self.icon("TOUCH_APP_OUTLINED"), size=42, color=ACCENT), ft.Text("Select a plot, text, shape or label", size=16, weight=ft.FontWeight.BOLD), ft.Text("Shift/Ctrl-click to select more. The last object selected becomes the orange alignment anchor.", size=12, color=MUTED), self.heading(""), ft.Text("", size=12, color=MUTED)]
            return
        if len(self.doc.selected) > 1:
            units = self.doc.units()
            self.right_content.controls = [self.heading(f"{len(self.doc.selected)} OBJECTS SELECTED", f"Anchor: {item.label or item.name or item.kind} · #{item.id}"), ft.Text("Move, rotate or arrange this selection as a whole.", size=12, color=MUTED), self.row(self.button("Group", self.handler(self.doc.group, edit=True)), self.button("Ungroup", self.handler(self.doc.ungroup, edit=True))), self.row(self.button("Rotate −15°", self.handler(self.rotate_selection, -15, edit=True)), self.button("Rotate +15°", self.handler(self.rotate_selection, 15, edit=True))), self.row(self.button("Front", self.handler(self.doc.layer, True, edit=True)), self.button("Back", self.handler(self.doc.layer, False, edit=True))), self.row(self.button("Duplicate", self.handler(self.doc.duplicate, edit=True)), self.button("Delete", self.handler(self.doc.delete, edit=True))), self.button("Lock / unlock selection", self.handler(self.toggle_lock, edit=True))]
            if units:
                b = union(self.doc.unit_box(u) for u in units)
                self.right_content.controls.insert(1, ft.Text(f"X {b[0]:.1f}   Y {b[1]:.1f}\nW {b[2] - b[0]:.1f}   H {b[3] - b[1]:.1f}", size=13, color=MUTED))
            return
        label_only = self.doc.label_selected == item.id and item.kind == "panel"
        label = "PANEL LABEL" if label_only else (item.shape_type.replace("_", " ").upper() if item.kind == "shape" else item.kind.upper())
        controls = [self.heading(label, f"#{item.id} · {item.name or item.label or item.kind}"), self.row(self.button("Front", self.handler(self.doc.layer, True, edit=True)), self.button("Back", self.handler(self.doc.layer, False, edit=True)), self.ibutton("LOCK_OPEN" if self.doc.editable(item) else "LOCK", "Lock/unlock", self.handler(self.toggle_lock, edit=True)))]
        self.prop = {}
        self.prop["x"] = self.field("X", self.number(item.x + item.label_offset_x if label_only else item.x), 128)
        self.prop["y"] = self.field("Y", self.number(item.y + item.label_offset_y if label_only else item.y), 128)
        self.prop["w"] = self.field("Width", self.number(item.w), 128)
        self.prop["h"] = self.field("Height", self.number(item.h), 128)
        self.prop["rotation"] = self.field("Rotation °", self.number(item.rotation), 128)
        self.prop["name"] = self.field("Object name", item.name)
        controls += [self.prop["name"], self.row(self.prop["x"], self.prop["y"])]
        if not label_only:
            controls += [self.row(self.prop["w"], self.prop["h"]), self.prop["rotation"]]
        if item.kind == "panel":
            self.prop["label"] = self.field("Panel label", item.label)
            self.prop["label_font_size"] = self.field("Label size", item.label_font_size, 128)
            self.prop["label_font_family"] = self.dropdown("Label font", item.label_font_family, [(f, f) for f in FONT_FAMILIES])
            self.prop["label_color"] = self.colour_field("Label colour", item.label_color)
            self.prop["show_label"] = ft.Checkbox(label="Show label", value=item.show_label)
            self.prop["include_in_labels"] = ft.Checkbox(label="Include in auto labels", value=item.include_in_labels)
            controls += [self.prop[k] for k in ["label", "label_font_size", "label_font_family", "label_color", "show_label", "include_in_labels"]]
            if not label_only:
                self.prop["image_fit"] = self.dropdown("Image fitting", item.image_fit, [("contain", "Contain · no cropping"), ("cover", "Cover · fill frame"), ("stretch", "Stretch · distort to fit")])
                self.prop["lock_aspect"] = ft.Checkbox(label="Keep ratio when resizing", value=item.lock_aspect)
                self.prop["border_width"] = self.field("Border width", item.border_width, 128)
                self.prop["border_color"] = self.colour_field("Border colour", item.border_color)
                controls += [self.prop[k] for k in ["image_fit", "lock_aspect", "border_width", "border_color"]]
                controls += [self.button("Replace image…", self.handler(self.replace_selected, item.id), "FIND_REPLACE", primary=True), ft.Text("", size=11, color=MUTED), ft.Text(f"Source: {item.original_size[0]:,} × {item.original_size[1]:,}\n{item.source_path}", size=10, color=MUTED, selectable=True)]
        else:
            self.prop["color"] = self.colour_field("Stroke / text colour", item.color)
            self.prop["font_size"] = self.field("Text size", item.font_size, 128)
            self.prop["font_family"] = self.dropdown("Font", item.font_family, [(f, f) for f in FONT_FAMILIES])
            controls += [self.prop["color"]]
            if item.kind == "text":
                self.prop["align"] = self.dropdown("Alignment", item.align, [(a, a.title()) for a in ["left", "center", "right", "justify"]])
                self.prop["line_spacing"] = self.field("Line spacing", item.line_spacing, 128)
                self.prop["background"] = self.colour_field("Text background", item.background, optional=True, empty_label="No background")
                self.prop["outline"] = self.colour_field("Text outline", item.outline, optional=True, empty_label="No outline")
                controls += [self.prop[k] for k in ["font_size", "font_family", "align", "line_spacing", "background", "outline"]]
                controls += [self.button("Edit rich text…", self.handler(self.edit_text, item.id), "EDIT_NOTE", primary=True), self.button("Fit box height to text", self.handler(self.fit_text_height, item.id, edit=True))]
            else:
                self.prop["stroke_width"] = self.field("Stroke width", item.stroke_width, 128)
                self.prop["fill"] = self.colour_field("Fill colour", item.fill, optional=True, empty_label="Use stroke colour")
                self.prop["fill_alpha"] = ft.Slider(min=0, max=100, divisions=100, value=round(item.fill_alpha / 255 * 100), label="{value}%", active_color=ACCENT, data=round(item.fill_alpha / 255 * 100))
                controls += [self.prop["stroke_width"], self.prop["fill"], ft.Text("Fill opacity · 0% transparent / 100% solid", size=11, color=MUTED), self.prop["fill_alpha"]]
                if item.is_stat:
                    self.prop["text"] = self.field("Comparison text", item.text)
                    self.prop["compare_style"] = self.dropdown("Style", item.compare_style, [("bracket", "Bracket line"), ("line", "Straight line")])
                    self.prop["bracket_height"] = self.field("Bracket height", item.bracket_height, 128)
                    self.prop["text_gap"] = self.field("Text gap", item.text_gap, 128)
                    self.prop["auto_center_text"] = ft.Checkbox(label="Center text", value=item.auto_center_text)
                    self.prop["snap_to_tops"] = ft.Checkbox(label="Level above both endpoints", value=item.snap_to_tops)
                    controls += [self.prop[k] for k in ["text", "font_size", "font_family", "compare_style", "bracket_height", "text_gap", "auto_center_text", "snap_to_tops"]]
        controls += [self.button("Apply properties", self.handler(self.apply_properties, item.id, label_only, edit=True), "CHECK", primary=True), ft.Text("Click a colour swatch to choose a common colour or adjust the sliders. Apply properties to save changes.", size=10, color=MUTED)]
        self.right_content.controls = controls

    def build_layers(self):
        controls = [self.heading("LAYERS", "Topmost objects appear first.")]
        for item in sorted(self.doc.items, key=lambda i: i.z_index, reverse=True):
            selected = item.id in self.doc.selected
            title = item.name or item.label or (item.text[:28] if item.kind == "text" else item.shape_type if item.kind == "shape" else "Image")
            icon = "IMAGE_OUTLINED" if item.kind == "panel" else "TEXT_FIELDS" if item.kind == "text" else "CATEGORY_OUTLINED"
            row = ft.Row([self.ibutton(icon, "Select object", self.handler(self.select_layer, item.id)), ft.Text(f"{item.label + ' · ' if item.label else ''}{title}", size=12, max_lines=1, expand=True), self.ibutton("LOCK" if item.locked else "LOCK_OPEN", "Lock/unlock", self.handler(self.toggle_lock, item.id, edit=True)), self.ibutton("VISIBILITY" if item.visible else "VISIBILITY_OFF", "Show/hide", self.handler(self.toggle_visible, item.id, edit=True))], spacing=0)
            controls.append(ft.Container(content=row, bgcolor="#FFF0DE" if selected else "#FFFFFF", border_radius=8, border=ft.Border.all(1, ANCHOR if self.doc.anchor == item.id else BORDER)))
        controls += [self.heading("GROUP SAFETY"), ft.Text(" Duplicate creates a separate group.", size=11, color=MUTED)]
        self.right_content.controls = controls

    @staticmethod
    def number(value):
        return f"{float(value):.2f}".rstrip("0").rstrip(".") if float(value) else "0"

    def parse(self, control, minimum=-1e6, maximum=1e6):
        result = float(str(control.value).strip())
        if not math.isfinite(result) or not minimum <= result <= maximum:
            raise ValueError(f"{control.label}: enter a number between {minimum:g} and {maximum:g}")
        return result

    def notice(self, text, error=False):
        self.status.value = str(text).replace("\n", " ")[:260]
        self.status.color = "#AD3545" if error else MUTED
        if error:
            self.page.show_dialog(ft.SnackBar(content=ft.Text(str(text)[:1800]), bgcolor="#99374A"))
        self.page.update()

    def refresh(self, inspector=True):
        if self.closed:
            return
        if inspector:
            self.build_right()
        self.title_text.value = ("● " if self.doc.dirty else "") + (Path(self.path).name if self.path else "Untitled figure")
        self.undo_button.disabled = self.busy or not self.doc.undo_stack
        self.redo_button.disabled = self.busy or not self.doc.redo_stack
        self.cancel_button.visible = self.busy
        self.progress.visible = self.busy
        self.zoom_text.value = f"{round(self.zoom * 100)}%"
        self.update_modes()
        if time.monotonic() - self.last_memory > 2:
            self.last_memory = time.monotonic()
            history = sum(len(s) for s in self.doc.undo_stack + self.doc.redo_stack) / 1048576
            try:
                import psutil
                proc = psutil.Process()
                rss = proc.memory_info().rss + sum(p.memory_info().rss for p in proc.children() if p.is_running())
                self.memory_text.value = f"{len(self.doc.items)} objects · RAM {rss / 1048576:.0f} MB · history {history:.1f} MB"
            except Exception:
                self.memory_text.value = f"{len(self.doc.items)} objects · history {history:.1f} MB"
        self.paint()
        self.update_drop_region()
        self.page.update()

    def update_modes(self):
        self.mode_text.value = self.mode.replace("_", " ").upper() + (" · Esc to exit" if self.mode not in ("select", "pan") else "")
        for mode, button in self.tool_buttons.items():
            active = self.mode == mode
            bg = "#FBE7E9" if mode == "erase" else "#FFF0DA" if mode == "crop" else "#FFE5C8"
            button.style = ft.ButtonStyle(bgcolor=bg if active else "#FFFFFF", color="#B33245" if active and mode == "erase" else ACCENT if active else MUTED)

    def configure_gestures(self):
        click_only = self.mode in CLICK_TOOLS
        self.gesture.on_pan_start = None if click_only else self.on_pan_start
        self.gesture.on_pan_update = None if click_only else self.on_pan_update
        self.gesture.on_pan_end = None if click_only else self.on_pan_end
        self.gesture.on_pan_cancel = None if click_only else self.on_pan_cancel
        self.gesture.on_double_tap_down = self.on_double_click if self.mode == "select" else None
        self.gesture.mouse_cursor = ft.MouseCursor.PRECISE if self.mode not in ("select", "pan") else ft.MouseCursor.GRAB if self.mode == "pan" else ft.MouseCursor.BASIC

    def set_stat_symbol(self, value):
        self.stat_symbol.value = value

    def set_mode(self, mode):
        if self.busy and mode not in ("select", "pan"):
            self.notice("The current operation is still running.")
            return
        next_mode = "select" if self.mode == mode and mode != "select" else mode
        values = self.stat_values() if next_mode in ("compare", "symbol") else None
        self.mode = next_mode
        self._stat_session = values
        self.drag, self.stat_first, self.stat_hover, self.marquee, self.guides = None, None, None, None, []
        self.configure_gestures()
        if self.mode == "label-order":
            self.sequence_ids = []
        if self.mode == "label-align":
            self.click_align_anchor = None
            self.click_align_axis = self.label_axis_field.value
        tips = {"select": "Select and drag. Shift/Ctrl-click adds objects; last selected is the anchor.", "pan": "Drag the workspace to pan. Mouse wheel scrolls; Ctrl+wheel zooms.", "crop": "Drag inside a panel to crop. Crops are non-destructive; labels retain their positions.", "erase": "Drag a white erase rectangle inside a panel. Undo or Reset edits restores it.", "compare": "Click the first data position, then the second.", "symbol": "Click to place the statistic symbol.", "label-order": "Click panels in A, B, C order. Existing labels outside the sequence are not changed.", "label-align": "Click the reference label first, then click labels to align."}
        self.status.value = tips.get(self.mode, "Drag to draw. Click for a default-size shape. Shift constrains proportions / angle.")
        if hasattr(self, "stat_hint"):
            self.stat_hint.value = "1 of 2 · Click the first data point." if self.mode == "compare" else "Click once to place the symbol." if self.mode == "symbol" else "Choose Bracket, then click two data points."

    def screen(self, p):
        return self.pan[0] + p[0] * self.zoom, self.pan[1] + p[1] * self.zoom

    def model(self, p):
        return (p[0] - self.pan[0]) / self.zoom, (p[1] - self.pan[1]) / self.zoom

    def event_position(self, e):
        return event_position(e, self.gesture_last or (0.0, 0.0))

    def has_key(self, key):
        key = key.lower()
        native = windows_modifiers()
        if native is not None and key in native:
            return native[key]
        return any(key in k for k in self.keys)

    def actual_size(self):
        self.zoom_by(1 / self.zoom)

    def fit_view(self):
        s = self.doc.settings
        self.zoom = clamp(min((self.view_w - 64) / s.canvas_w, (self.view_h - 64) / s.canvas_h), 0.05, 8)
        self.pan = [(self.view_w - s.canvas_w * self.zoom) / 2, (self.view_h - s.canvas_h * self.zoom) / 2]
        self.fitted = True

    def zoom_by(self, factor, position=None):
        pivot = position or (self.view_w / 2, self.view_h / 2)
        at = self.model(pivot)
        self.zoom = clamp(self.zoom * factor, 0.05, 8)
        self.pan = [pivot[0] - at[0] * self.zoom, pivot[1] - at[1] * self.zoom]

    async def on_view_resize(self, e):
        if self.closed:
            return
        self.view_w, self.view_h = max(1, e.width), max(1, e.height)
        if not self.fitted:
            self.fit_view()
        self.refresh(inspector=False)

    def line_shape(self, a, b, color=GUIDE, width=1):
        return cv.Line(a[0], a[1], b[0], b[1], paint=ft.Paint(color=color, stroke_width=width))

    def polyline(self, points, color, width=1, closed=False):
        if closed:
            points = list(points) + [points[0]]
        return [self.line_shape(a, b, color, width) for a, b in zip(points, points[1:])]

    def paint(self):
        s = self.doc.settings
        x, y = self.screen((0, 0))
        w, h = s.canvas_w * self.zoom, s.canvas_h * self.zoom
        shapes = [cv.Rect(x + 3, y + 4, w, h, paint=ft.Paint(color="#DED3C6")), cv.Rect(x, y, w, h, paint=ft.Paint(color=s.bg_color))]
        m = s.margin * self.zoom
        if 2 * m < min(w, h):
            shapes.append(cv.Rect(x + m, y + m, w - 2 * m, h - 2 * m, paint=ft.Paint(color="#DED4C8", style=ft.PaintingStyle.STROKE, stroke_width=1)))
        if s.show_grid and s.grid_step > 0:
            step = s.grid_step * self.zoom
            while step < 12:
                step *= 2
            startx = max(x, x + math.ceil(-x / step) * step)
            starty = max(y, y + math.ceil(-y / step) * step)
            for n in range(min(300, max(0, int((min(x + w, self.view_w) - startx) / step) + 1))):
                gx = startx + n * step
                shapes.append(self.line_shape((gx, max(0, y)), (gx, min(self.view_h, y + h)), "#F1EAE0"))
            for n in range(min(300, max(0, int((min(y + h, self.view_h) - starty) / step) + 1))):
                gy = starty + n * step
                shapes.append(self.line_shape((max(0, x), gy), (min(self.view_w, x + w), gy), "#F1EAE0"))
        self.background_canvas.shapes = shapes
        controls = [self.background_canvas]
        visible_ids, label_ids = set(), set()
        view = (*self.model((-100, -100)), *self.model((self.view_w + 100, self.view_h + 100)))
        for item in sorted(self.doc.items, key=lambda i: i.z_index):
            if not item.visible:
                continue
            if overlaps(visual_bbox(item), view):
                control = self.item_control(item)
                controls.append(control)
                visible_ids.add(item.id)
            lb = label_bbox(item)
            if lb and overlaps(lb, view):
                if item.id not in self.label_controls:
                    self.label_controls[item.id] = ft.Image(src=b"", fit=ft.BoxFit.FILL, gapless_playback=True)
                label_control = self.label_controls[item.id]
                key = ("label", item.label, item.label_font_family, item.label_font_size, item.label_color)
                label_control.src = self.byte_cache.obtain(key, lambda i=item: label_bitmap(i))
                label_control.left, label_control.top = self.screen((lb[0], lb[1]))
                label_control.width, label_control.height = max(1, (lb[2] - lb[0]) * self.zoom), max(1, (lb[3] - lb[1]) * self.zoom)
                controls.append(label_control)
                label_ids.add(item.id)
        self.visual_controls = {k: v for k, v in self.visual_controls.items() if k in visible_ids}
        self.label_controls = {k: v for k, v in self.label_controls.items() if k in label_ids}
        overlay = []
        selected = [self.doc.get(i) for i in self.doc.selected]
        selected = [i for i in selected if i and i.visible]
        if self.doc.label_selected and len(selected) == 1:
            lb = label_bbox(selected[0])
            if lb:
                a, b = self.screen((lb[0], lb[1])), self.screen((lb[2], lb[3]))
                overlay.append(cv.Rect(a[0] - 3, a[1] - 3, b[0] - a[0] + 6, b[1] - a[1] + 6, paint=ft.Paint(color=ANCHOR, style=ft.PaintingStyle.STROKE, stroke_width=2)))
        elif selected:
            for item in selected:
                c = ANCHOR if item.id == self.doc.anchor else ACCENT
                overlay += self.polyline([self.screen(p) for p in item.corners()], c, 1.3, True)
            handles = self.handles()
            for name, p in handles.items():
                px, py = self.screen(p)
                if name == "rotate":
                    center = self.selection_center()
                    overlay.append(self.line_shape(self.screen(center), (px, py), ACCENT, 1))
                    overlay.append(cv.Circle(px, py, 6, paint=ft.Paint(color=ACCENT)))
                    overlay.append(cv.Circle(px, py, 3, paint=ft.Paint(color="#FFFFFF")))
                else:
                    overlay.append(cv.Rect(px - 4, py - 4, 8, 8, paint=ft.Paint(color="#FFFFFF")))
                    overlay.append(cv.Rect(px - 4, py - 4, 8, 8, paint=ft.Paint(color=ACCENT, style=ft.PaintingStyle.STROKE, stroke_width=1.5)))
        for axis, value in self.guides:
            if axis == "x":
                gx = self.screen((value, 0))[0]
                overlay.append(self.line_shape((gx, 0), (gx, self.view_h), GUIDE, 1))
            else:
                gy = self.screen((0, value))[1]
                overlay.append(self.line_shape((0, gy), (self.view_w, gy), GUIDE, 1))
        if self.marquee:
            a, b = self.screen(self.marquee[0]), self.screen(self.marquee[1])
            c = "#D34D5F" if self.mode == "erase" else ANCHOR if self.mode == "crop" else ACCENT
            overlay.append(cv.Rect(min(a[0], b[0]), min(a[1], b[1]), max(1, abs(b[0] - a[0])), max(1, abs(b[1] - a[1])), paint=ft.Paint(color=c, stroke_width=1.5, style=ft.PaintingStyle.STROKE)))
        if self.mode == "compare" and self.stat_first is not None:
            px, py = self.screen(self.stat_first)
            overlay.append(cv.Circle(px, py, 6, paint=ft.Paint(color=ACCENT)))
            overlay.append(cv.Circle(px, py, 3, paint=ft.Paint(color="#FFFFFF")))
            if self.stat_hover is not None and abs(self.stat_first[0] - self.stat_hover[0]) * self.zoom >= 3:
                values = dict(self._stat_session or self.stat_values())
                first, second = self.stat_first, self.stat_hover
                preview = Item(id=-1, kind="shape", shape_type="stat_compare", x=first[0], y=first[1], w=second[0] - first[0], h=second[1] - first[1], font_family=s.default_font_family, **values)
                overlay += self.polyline([self.screen(point) for point in preview.base_points()], values["color"], max(1, values["stroke_width"] * self.zoom))
                tx, ty = preview.stat_text_point()
                tw, th = text_extent(preview.text, preview.font_family, preview.font_size, True)
                label = Item(id=-2, label=preview.text, label_font_family=preview.font_family, label_font_size=preview.font_size, label_color=preview.color)
                key = ("label", label.label, label.label_font_family, label.label_font_size, label.label_color)
                raw = self.byte_cache.obtain(key, lambda: label_bitmap(label))
                lx, ly = self.screen((tx - tw / 2, ty - th))
                controls.append(ft.Image(src=raw, left=lx, top=ly, width=max(1, tw * self.zoom), height=max(1, th * self.zoom), fit=ft.BoxFit.FILL, gapless_playback=True))
                endx, endy = self.screen(second)
                overlay.append(cv.Circle(endx, endy, 4, paint=ft.Paint(color=ACCENT)))
        self.overlay_canvas.shapes = overlay
        controls.append(self.overlay_canvas)
        self.stack.controls = controls

    def item_control(self, item):
        if item.id not in self.visual_controls:
            self.visual_controls[item.id] = ft.Container(content=ft.Image(src=b"", gapless_playback=True, fit=ft.BoxFit.FILL), clip_behavior=ft.ClipBehavior.HARD_EDGE)
        wrapper = self.visual_controls[item.id]
        image = wrapper.content
        if item.kind == "panel":
            key = preview_key(item.to_dict())
            path = self.previews.get(key)
            if path and Path(path).exists():
                edge = max(item.w, item.h) * self.zoom * 1.5
                level = next((n for n in (128, 256, 512) if edge <= n), None)
                small_path = str(Path(path).with_name(Path(path).stem + f".{level}.png")) if level else path
                source_path = small_path if Path(small_path).exists() else path
                image.src = self.byte_cache.obtain(("file", source_path), lambda: Path(source_path).read_bytes())
            else:
                image.src = '<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80"><rect width="80" height="80" fill="#F3EEE7"/><path d="M10 60L30 35L45 50L58 22L70 60Z" fill="#D3C6B7"/></svg>'
            image.fit = {"contain": ft.BoxFit.CONTAIN, "cover": ft.BoxFit.COVER, "stretch": ft.BoxFit.FILL}[item.image_fit]
            wrapper.left, wrapper.top = self.screen((item.x, item.y))
            wrapper.width, wrapper.height = max(1, item.w * self.zoom), max(1, item.h * self.zoom)
            image.width, image.height = wrapper.width, wrapper.height
            wrapper.rotate = ft.Rotate(angle=math.radians(item.rotation), alignment=ft.Alignment.CENTER)
            wrapper.border = ft.Border.all(item.border_width * self.zoom, item.border_color) if item.border_width > 0 else None
        else:
            data = item.to_dict()
            data["x"], data["y"] = 0, 0
            for key in ("id", "name", "z_index", "group_id", "locked", "visible"):
                data.pop(key, None)
            small = copy.deepcopy(item)
            small.x, small.y = 0, 0
            b = visual_bbox(small)
            density = min(2.0, 1600 / max(1, b[2] - b[0], b[3] - b[1]))
            key = ("object", json.dumps(data, sort_keys=True), round(density, 3))
            def build():
                bitmap, _ = render_object(small, self.root, density, True)
                stream = io.BytesIO()
                bitmap.save(stream, "PNG", compress_level=2)
                bitmap.close()
                return stream.getvalue()
            image.src = self.byte_cache.obtain(key, build)
            b = visual_bbox(item)
            wrapper.left, wrapper.top = self.screen((b[0], b[1]))
            wrapper.width, wrapper.height = max(1, (b[2] - b[0]) * self.zoom), max(1, (b[3] - b[1]) * self.zoom)
            image.width, image.height = wrapper.width, wrapper.height
            image.fit = ft.BoxFit.FILL
            wrapper.rotate, wrapper.border = None, None
        return wrapper

    def selection_center(self):
        items = [self.doc.get(i) for i in self.doc.selected]
        items = [i for i in items if i and i.visible]
        if len(items) == 1:
            return items[0].center
        b = union(i.bbox() for i in items)
        return (b[0] + b[2]) / 2, (b[1] + b[3]) / 2

    def handles(self):
        items = [self.doc.get(i) for i in self.doc.selected]
        items = [i for i in items if i and i.visible]
        if not items or self.doc.label_selected or any(not self.doc.editable(i) for i in items):
            return {}
        if len(items) == 1:
            item = items[0]
            if item.is_line or item.is_stat:
                ends = [rotate((item.x, item.y), item.center, item.rotation), rotate((item.x + item.w, item.y + item.h), item.center, item.rotation)]
                handles = {"end0": ends[0], "end1": ends[1]}
                b = item.bbox()
                handles["rotate"] = (item.center[0], b[1] - 28 / self.zoom)
                return handles
            handles = {f"c{n}": p for n, p in enumerate(item.corners())}
            handles["rotate"] = rotate((item.x + item.w / 2, item.y - 28 / self.zoom), item.center, item.rotation)
            return handles
        b = union(i.bbox() for i in items)
        return {"g0": (b[0], b[1]), "g1": (b[2], b[1]), "g2": (b[2], b[3]), "g3": (b[0], b[3]), "rotate": ((b[0] + b[2]) / 2, b[1] - 28 / self.zoom)}

    def hit_label(self, p):
        tolerance = 4 / self.zoom
        for item in sorted(self.doc.items, key=lambda i: i.z_index, reverse=True):
            b = label_bbox(item)
            if b and b[0] - tolerance <= p[0] <= b[2] + tolerance and b[1] - tolerance <= p[1] <= b[3] + tolerance:
                return item
        return None

    def hit_item(self, p):
        return next((i for i in sorted(self.doc.items, key=lambda i: i.z_index, reverse=True) if i.visible and i.hit(p, 4 / self.zoom)), None)

    def schedule_canvas_focus(self):
        self.typing = False
        self.page.run_task(self.focus_canvas)

    async def on_tap_down(self, e):
        if self.dialog_depth or self.closed:
            return
        self.gesture_last = self.event_position(e)
        if self.mode in CLICK_TOOLS:
            self.drag = None
            self.schedule_canvas_focus()
            return
        await self.begin_pointer(self.gesture_last)

    async def on_tap_up(self, e):
        if self.dialog_depth or self.closed:
            return
        self.gesture_last = self.event_position(e)
        if self.mode in CLICK_TOOLS:
            try:
                await self.perform_click_tool(self.model(self.gesture_last))
            except Exception as exc:
                logging.exception("Click tool failed")
                self.notice(str(exc), error=True)
            self.refresh()
        elif self.drag and not self.drag.get("panning"):
            await self.finish_pointer(self.gesture_last)

    async def on_pan_start(self, e):
        if self.dialog_depth or self.closed or self.mode in CLICK_TOOLS:
            return
        p = self.event_position(e)
        if self.drag is None:
            await self.begin_pointer(p)
        if self.drag:
            self.drag["panning"] = True
        self.gesture_last = p

    def update_drop_region(self):
        self._canvas_origin = (248.0, 116.0 + (3.0 if self.busy else 0.0))
        self.file_drop.set_region(*self._canvas_origin, self.view_w, self.view_h, self.page.width or 0)
        self.file_drop.set_enabled(not self.busy and not self.dialog_depth and self.drag is None and not self.closed)

    def remember_canvas_origin(self, event):
        point = getattr(event, "global_position", None)
        if point is not None:
            local = self.event_position(event)
            self._canvas_origin = (point.x - local[0], point.y - local[1])
            self.file_drop.set_region(*self._canvas_origin, self.view_w, self.view_h, self.page.width or 0)

    async def on_canvas_enter(self, e):
        self.remember_canvas_origin(e)

    async def on_canvas_exit(self, e):
        if self.stat_hover is not None:
            self.stat_hover = None
            self.paint()
            self.stack.update()

    async def on_hover(self, e):
        self.remember_canvas_origin(e)
        if self.mode != "compare" or self.stat_first is None or self.dialog_depth or self.busy:
            return
        self.stat_hover = self.model(self.event_position(e))
        self.paint()
        self.stack.update()

    async def desktop_drop_loop(self):
        if not self.file_drop.start():
            return
        try:
            while not self.closed:
                self.update_drop_region()
                try:
                    event = self.file_drop.events.get_nowait()
                except queue.Empty:
                    await asyncio.sleep(0.04)
                    continue
                if event[0] == "ready":
                    self.drop_status = "Drag images or PDFs from Explorer into the workspace."
                elif event[0] == "error":
                    self.drop_status = "File drop unavailable: " + str(event[1])
                    logging.warning(self.drop_status)
                elif event[0] == "files":
                    await self.accept_file_drop(event[1], event[2])
                elif event[0] == "stopped":
                    break
                if hasattr(self, "drop_status_text"):
                    self.drop_status_text.value = self.drop_status
                self.page.update()
        finally:
            self.file_drop.stop()

    async def accept_file_drop(self, paths, position):
        if self.busy or self.dialog_depth or self.closed:
            self.notice("Finish the current operation, then drop the files again.")
            return
        accepted, rejected = dropped_paths(paths)
        if not accepted:
            self.notice("Drop PNG, JPEG, TIFF, BMP, GIF, WebP or PDF files from Windows Explorer.")
            return
        self.set_mode("select")
        self.keys.clear()
        await self.import_paths(accepted, drop_position=self.model(position))
        if rejected:
            self.notice(f"Imported supported files. Skipped {len(rejected)} unsupported file(s) or folder(s).")

    async def begin_pointer(self, position):
        self.schedule_canvas_focus()
        p = self.model(position)
        self.drag = {"screen_start": position, "start": p, "current": p, "before": self.doc.snapshot(), "origins": {}, "action": None, "panning": False, "pan": list(self.pan), "selected_before": list(self.doc.selected)}
        if self.mode == "pan" or self.has_key("space") or self.busy:
            self.drag["action"] = "pan"
            return
        if self.mode in ("crop", "erase"):
            item = self.hit_item(p)
            if item and item.kind == "panel" and self.doc.editable(item):
                self.drag.update(action=self.mode, target=item.id)
                self.doc.select(item.id)
                self.marquee = (p, p)
            return
        if self.mode in ("symbol", "compare", "label-order", "label-align"):
            self.drag["action"] = "click-tool"
            return
        if self.mode not in ("select", "pan"):
            shape = self.mode
            item = self.doc.new_item(kind="shape", shape_type=shape, x=p[0], y=p[1], w=4, h=4, lock_aspect=False, fill="#F2C94C" if shape == "highlight_bar" else "#171923" if shape == "arrow_head" else "", fill_alpha=80 if shape == "highlight_bar" else 255 if shape == "arrow_head" else 0, color="#DDAE30" if shape == "highlight_bar" else "#171923")
            self.drag.update(action="draw", target=item.id)
            self.refresh()
            return
        for name, hp in self.handles().items():
            if math.hypot(p[0] - hp[0], p[1] - hp[1]) <= 9 / self.zoom:
                self.drag["action"] = name
                self.drag["center"] = self.selection_center()
                self.drag["angle"] = math.degrees(math.atan2(p[1] - self.drag["center"][1], p[0] - self.drag["center"][0]))
                self.drag["origins"] = {i.id: copy.deepcopy(i) for u in self.doc.units() for i in u}
                self.drag["box"] = union(i.bbox() for i in self.drag["origins"].values())
                return
        label = self.hit_label(p)
        item = label or self.hit_item(p)
        additive = self.has_key("shift") or self.has_key("control") or self.has_key("meta")
        if item is None:
            self.drag["action"] = "marquee"
            self.drag["additive"] = additive
            self.marquee = (p, p)
            if not additive:
                self.doc.selected, self.doc.anchor, self.doc.label_selected = [], None, None
            self.refresh()
            return
        if additive:
            self.doc.select(item.id, True, bool(label), toggle=self.has_key("control") or self.has_key("meta"))
            if item.id not in self.doc.selected:
                self.drag = None
                self.refresh()
                return
        elif item.id not in self.doc.selected or bool(label) != bool(self.doc.label_selected):
            self.doc.select(item.id, label=bool(label))
        else:
            self.doc.anchor = item.id
        if not self.doc.editable(item):
            self.drag = None
            self.status.value = "This object or its group is locked. Unlock it in Properties or Layers."
            self.refresh()
            return
        if label:
            self.drag["action"] = "label"
            self.drag["target"] = item.id
            self.drag["origins"] = {item.id: copy.deepcopy(item)}
            self.drag["box"] = label_bbox(item)
        else:
            self.drag["action"] = "move"
            self.drag["origins"] = {i.id: copy.deepcopy(i) for u in self.doc.units() for i in u}
            self.drag["box"] = union(i.bbox() for i in self.drag["origins"].values())
        self.refresh()

    async def on_pan_update(self, e):
        if not self.drag or self.dialog_depth:
            return
        position = self.event_position(e)
        self.gesture_last = position
        self.update_pointer(position)
        if time.monotonic() - self.last_render >= 1 / 35:
            self.last_render = time.monotonic()
            self.paint()
            self.stack.update()

    def update_pointer(self, position):
        if not self.drag:
            return
        d = self.drag
        p = self.model(position)
        d["current"] = p
        action = d["action"]
        dx, dy = p[0] - d["start"][0], p[1] - d["start"][1]
        if action == "pan":
            self.pan = [d["pan"][0] + position[0] - d["screen_start"][0], d["pan"][1] + position[1] - d["screen_start"][1]]
            return
        if action in ("marquee", "crop", "erase"):
            self.marquee = (d["start"], p)
            return
        if action == "draw":
            item = self.doc.get(d["target"])
            if not item:
                return
            if item.is_line:
                if self.has_key("shift"):
                    angle = round(math.atan2(dy, dx) / (math.pi / 4)) * math.pi / 4
                    length = math.hypot(dx, dy)
                    dx, dy = length * math.cos(angle), length * math.sin(angle)
                item.x, item.y, item.w, item.h = d["start"][0], d["start"][1], dx, dy
            else:
                if self.has_key("shift"):
                    side = max(abs(dx), abs(dy))
                    dx, dy = math.copysign(side, dx), math.copysign(side, dy)
                item.x, item.y = min(d["start"][0], d["start"][0] + dx), min(d["start"][1], d["start"][1] + dy)
                item.w, item.h = max(4, abs(dx)), max(4, abs(dy))
            return
        if action == "click-tool" or action is None:
            return
        origins = d["origins"]
        if not origins:
            return
        if action in ("move", "label"):
            settings = copy.copy(self.doc.settings)
            if self.has_key("shift"):
                if abs(dx) >= abs(dy):
                    dy, settings.snap_y = 0, False
                else:
                    dx, settings.snap_x = 0, False
            if action == "label":
                targets = [label_bbox(i) for i in self.doc.items if i.id not in origins and label_bbox(i)]
            else:
                targets = [i.bbox() for i in self.doc.items if i.visible and i.id not in origins]
            dx, dy, self.guides = snap_translation(d["box"], targets, dx, dy, settings, self.zoom, self.has_key("alt"))
            for identifier, original in origins.items():
                item = self.doc.get(identifier)
                if action == "label":
                    item.label_offset_x = original.label_offset_x + dx
                    item.label_offset_y = original.label_offset_y + dy
                else:
                    item.x, item.y = original.x + dx, original.y + dy
        elif action == "rotate":
            center = d["center"]
            angle = math.degrees(math.atan2(p[1] - center[1], p[0] - center[0])) - d["angle"]
            if self.has_key("shift"):
                first = next(iter(origins.values())).rotation if len(origins) == 1 else 0
                angle = round((angle + first) / 15) * 15 - first
            self.apply_rotation_origins(origins, center, angle)
        elif action.startswith("c") and action[1:].isdigit():
            original = next(iter(origins.values()))
            result = resized_corner(original, int(action[1]), p, self.has_key("shift") or (original.kind == "panel" and original.lock_aspect))
            item = self.doc.get(original.id)
            item.x, item.y, item.w, item.h = result.x, result.y, result.w, result.h
        elif action.startswith("g"):
            box = d["box"]
            corner = int(action[1])
            corners = [(box[0], box[1]), (box[2], box[1]), (box[2], box[3]), (box[0], box[3])]
            opposite = corners[(corner + 2) % 4]
            base = corners[corner]
            vx, vy = base[0] - opposite[0], base[1] - opposite[1]
            factor = max(0.02, ((p[0] - opposite[0]) * vx + (p[1] - opposite[1]) * vy) / max(1e-9, vx * vx + vy * vy))
            for identifier, original in origins.items():
                replacement = copy.deepcopy(original)
                self.doc.scale_unit([replacement], factor, opposite, scale_styles=True)
                target = self.doc.get(identifier)
                target.__dict__.update(replacement.__dict__)
        elif action.startswith("end"):
            original = next(iter(origins.values()))
            a = rotate((original.x, original.y), original.center, original.rotation)
            b = rotate((original.x + original.w, original.y + original.h), original.center, original.rotation)
            if action == "end0":
                a = p
            else:
                b = p
            if self.has_key("shift") and original.is_line:
                fixed = b if action == "end0" else a
                moving = a if action == "end0" else b
                angle = round(math.atan2(moving[1] - fixed[1], moving[0] - fixed[0]) / (math.pi / 4)) * math.pi / 4
                length = math.dist(fixed, moving)
                adjusted = fixed[0] + length * math.cos(angle), fixed[1] + length * math.sin(angle)
                if action == "end0":
                    a = adjusted
                else:
                    b = adjusted
            center = (a[0] + b[0]) / 2, (a[1] + b[1]) / 2
            delta = rotate((b[0] - a[0], b[1] - a[1]), (0, 0), -original.rotation)
            target = self.doc.get(original.id)
            target.x, target.y, target.w, target.h = center[0] - delta[0] / 2, center[1] - delta[1] / 2, delta[0], delta[1]

    def apply_rotation_origins(self, origins, center, angle):
        for identifier, original in origins.items():
            item = self.doc.get(identifier)
            if len(origins) > 1:
                rotated = rotate(original.center, center, angle)
                item.x, item.y = rotated[0] - original.w / 2, rotated[1] - original.h / 2
                label = rotate((original.x + original.label_offset_x, original.y + original.label_offset_y), center, angle)
                item.label_offset_x, item.label_offset_y = label[0] - item.x, label[1] - item.y
            item.rotation = (original.rotation + angle) % 360

    async def on_pan_end(self, e):
        if self.drag and self.drag.get("panning"):
            await self.finish_pointer(event_position(e, self.gesture_last or self.drag["screen_start"]))

    async def on_pan_cancel(self, e):
        if self.drag is None or (e is not None and not self.drag.get("panning")):
            return
        if self.drag["action"] not in ("pan", "marquee"):
            self.doc.restore(self.drag["before"])
        elif self.drag["action"] == "pan":
            self.pan = self.drag["pan"]
        self.drag, self.marquee, self.guides = None, None, []
        self.refresh()

    async def finish_pointer(self, position):
        if not self.drag:
            return
        d = self.drag
        self.update_pointer(position)
        action = d["action"]
        p = self.model(position)
        self.drag = None
        try:
            if action == "marquee":
                box = bounds([d["start"], p])
                ids = []
                if math.dist(d["start"], p) * self.zoom > 4:
                    for i in self.doc.items:
                        if i.visible and overlaps(box, i.bbox()):
                            ids.extend(m.id for m in self.doc.members(i))
                self.doc.selected = list(dict.fromkeys((d["selected_before"] if d["additive"] else []) + ids))
                self.doc.anchor = self.doc.selected[-1] if self.doc.selected else None
                self.doc.label_selected = None
            elif action in ("crop", "erase"):
                item = self.doc.get(d["target"])
                if item and crop_or_erase(item, d["start"], p, action == "erase"):
                    self.busy = True
                    self.status.value = "Preparing edited image preview…"
                    self.refresh()
                    path = await self.worker(make_preview, item.to_dict(), self.root)
                    self.previews[preview_key(item.to_dict())] = path
            elif action == "draw":
                item = self.doc.get(d["target"])
                if math.dist(d["start"], p) * self.zoom < 5:
                    item.w = 160 if item.shape_type != "arrow_head" else 40
                    item.h = 0 if item.is_line else 28 if item.shape_type in ("highlight_bar", "arrow_head") else 100
                self.set_mode("select")
            elif action == "click-tool":
                await self.perform_click_tool(p)
            if action not in ("pan", "marquee", "click-tool", None) and self.doc.snapshot() != d["before"]:
                self.doc.checkpoint(d["before"])
        except Exception as exc:
            if action not in ("pan", "marquee", "click-tool", None):
                self.doc.restore(d["before"])
            logging.exception("Pointer action failed")
            self.notice(str(exc), error=True)
        finally:
            if action in ("crop", "erase"):
                self.busy = False
            self.marquee, self.guides = None, []
            self.refresh()

    async def perform_click_tool(self, p):
        if self.busy:
            return
        before = self.doc.snapshot()
        if self.mode in ("symbol", "compare"):
            values = dict(self._stat_session) if self.stat_first is not None and self._stat_session else self.stat_values()
            text, size = values["text"], values["font_size"]
            if self.mode == "symbol":
                width = max(40, text_extent(text, self.doc.settings.default_font_family, size, True)[0] + 14)
                item = self.doc.new_item(kind="text", x=p[0] - width / 2, y=p[1] - size, w=width, h=size * 2, text=text, rich_runs=[{"text": text, "bold": True, "italic": False}], font_size=size, font_family=self.doc.settings.default_font_family, bold=True, align="center", color=values["color"])
                if not self.stat_repeat.value:
                    self.set_mode("select")
                self.status.value = "Statistic symbol placed."
            elif self.stat_first is None:
                self.stat_first = p
                self._stat_session = dict(values)
                self.stat_hover = None
                self.stat_hint.value = "2 of 2 · Click the second data point."
                self.status.value = "First point marked. Move the pointer to preview the bracket, then click the second data point."
                return
            else:
                a = self.stat_first
                if abs(a[0] - p[0]) * self.zoom < 3:
                    self.notice("Place the second point farther away horizontally.")
                    return
                self.doc.new_item(kind="shape", shape_type="stat_compare", x=a[0], y=a[1], w=p[0] - a[0], h=p[1] - a[1], font_family=self.doc.settings.default_font_family, **values)
                self.stat_first, self.stat_hover = None, None
                if self.stat_repeat.value:
                    self._stat_session = self.stat_values()
                    self.stat_hint.value = "1 of 2 · Click the next first point."
                else:
                    self.set_mode("select")
                self.status.value = "Comparison added. Symbols are centered above the line." if values["auto_center_text"] else "Comparison added."
        elif self.mode == "label-order":
            item = self.hit_label(p) or self.hit_item(p)
            if not item or item.kind != "panel" or not self.doc.editable(item) or item.id in self.sequence_ids:
                return
            item.label = label_name(len(self.sequence_ids))
            item.include_in_labels, item.show_label = True, True
            self.sequence_ids.append(item.id)
            self.doc.select(item.id, label=True)
        elif self.mode == "label-align":
            item = self.hit_label(p)
            if not item or not self.doc.editable(item):
                return
            if self.click_align_anchor is None:
                self.click_align_anchor = item.id
                self.doc.select(item.id, label=True)
                self.status.value = "Reference selected. Click another label to align it."
                return
            anchor = self.doc.get(self.click_align_anchor)
            if not anchor:
                self.click_align_anchor = None
                return
            if "x" in self.click_align_axis:
                item.label_offset_x = anchor.x + anchor.label_offset_x - item.x
            if "y" in self.click_align_axis:
                item.label_offset_y = anchor.y + anchor.label_offset_y - item.y
            self.doc.select(item.id, label=True)
        if self.doc.snapshot() != before:
            self.doc.checkpoint(before)

    def stat_values(self):
        if not hasattr(self, "stat_symbol"):
            return {"text": "***", "font_size": 28, "bracket_height": 28, "text_gap": 10, "compare_style": "bracket", "color": "#171923", "auto_center_text": True, "snap_to_tops": True, "stroke_width": 2}
        c = self.stat_color.value.strip()
        ImageColor.getrgb(c)
        return {"text": self.stat_symbol.value or "*", "font_size": self.parse(self.stat_size, 6, 500), "bracket_height": self.parse(self.stat_height, 0, 1000), "text_gap": self.parse(self.stat_gap, 0, 1000), "compare_style": self.stat_style.value, "color": c, "auto_center_text": bool(self.stat_center.value), "snap_to_tops": bool(self.stat_top.value), "stroke_width": self.parse(self.stat_width, 0.1, 50)}

    async def on_right_click(self, e):
        if self.dialog_depth:
            return
        self.drag = None
        self.keys.clear()
        p = self.model(self.event_position(e))
        label = self.hit_label(p)
        item = label or self.hit_item(p)
        if item and item.id not in self.doc.selected:
            self.doc.select(item.id, label=bool(label))
        elif item:
            self.doc.anchor = item.id
        entries = []
        if item and item.kind == "panel":
            entries.append(ft.PopupMenuItem(content="Replace image…", icon=self.icon("FIND_REPLACE"), on_click=self.handler(self.replace_selected, item.id)))
            entries.append(ft.PopupMenuItem(content="Edit panel label…", on_click=self.handler(self.edit_label, item.id)))
            entries.append(ft.PopupMenuItem(content="Trim white / transparent margins", on_click=self.handler(self.trim_selected, item.id)))
            entries.append(ft.PopupMenuItem(content="Crop image", on_click=self.handler(self.set_mode, "crop")))
            entries.append(ft.PopupMenuItem(content="Reset image edits", on_click=self.handler(self.reset_image_edits, item.id, edit=True)))
        if item and item.kind == "text":
            entries.append(ft.PopupMenuItem(content="Edit rich text…", on_click=self.handler(self.edit_text, item.id)))
        entries += [ft.PopupMenuItem(content="Bring to front", on_click=self.handler(self.doc.layer, True, edit=True)), ft.PopupMenuItem(content="Send to back", on_click=self.handler(self.doc.layer, False, edit=True)), ft.PopupMenuItem(content="Duplicate", on_click=self.handler(self.doc.duplicate, edit=True)), ft.PopupMenuItem(content="Delete", on_click=self.handler(self.doc.delete, edit=True)), ft.PopupMenuItem(content="Group", on_click=self.handler(self.doc.group, edit=True)), ft.PopupMenuItem(content="Ungroup", on_click=self.handler(self.doc.ungroup, edit=True)), ft.PopupMenuItem(content="Lock / unlock", on_click=self.handler(self.toggle_lock, edit=True)), ft.PopupMenuItem(content="Undo", on_click=self.handler(self.undo)), ft.PopupMenuItem(content="Redo", on_click=self.handler(self.redo))]
        self.context.items = entries
        self.refresh()
        await self.context.open(local_position=e.local_position, global_position=e.global_position)

    async def on_double_click(self, e):
        if self.busy or self.dialog_depth or self.mode != "select":
            return
        self.drag = None
        p = self.model(self.event_position(e))
        label = self.hit_label(p)
        item = label or self.hit_item(p)
        if not item:
            return
        self.doc.select(item.id, label=bool(label))
        if item.kind == "text":
            await self.edit_text(item.id)
        elif item.kind == "panel":
            await self.edit_label(item.id)
        self.refresh()

    async def on_scroll(self, e):
        if self.dialog_depth:
            return
        dx, dy = wheel_delta(e)
        if self.has_key("control") or self.has_key("meta"):
            self.zoom_by(wheel_zoom_factor(dy), self.event_position(e))
        elif self.has_key("shift"):
            self.pan[0] -= dy if dy else dx
        else:
            self.pan[0] -= dx
            self.pan[1] -= dy
        self.refresh(inspector=False)

    async def on_key_down(self, e):
        key = normalize_key(e.key)
        self.keys.add(key)
        if self.typing or self.dialog_depth:
            return
        await self.keyboard_action(key)

    async def on_key_repeat(self, e):
        key = normalize_key(e.key)
        if not self.typing and not self.dialog_depth and any(x in key for x in ["arrow", "left", "right", "up", "down"]):
            await self.keyboard_action(key)

    async def on_key_up(self, e):
        self.keys.discard(normalize_key(e.key))

    async def keyboard_action(self, key):
        ctrl = self.has_key("control") or self.has_key("meta")
        shift = self.has_key("shift")
        if key in ("escape", "esc"):
            if self.drag:
                await self.on_pan_cancel(None)
            self.set_mode("select")
            self.keys.clear()
            self.refresh()
            return
        if ctrl and key in ("+", "=", "add", "equal"):
            self.zoom_by(1.15)
        elif ctrl and key in ("-", "subtract", "minus"):
            self.zoom_by(1 / 1.15)
        elif ctrl and key == "0":
            self.fit_view()
        elif self.busy:
            return
        elif ctrl:
            mapping = {"z": self.redo if shift else self.undo, "y": self.redo, "a": self.select_all, "d": self.doc.duplicate, "g": self.doc.ungroup if shift else self.doc.group, "s": self.save_document, "o": self.pick_import if shift else self.open_document, "n": self.new_document, "v": self.paste, "e": self.export_dialog}
            if key in mapping:
                await self.handler(mapping[key], edit=key in ("d", "g"))()
                return
        elif key in ("delete", "backspace"):
            await self.handler(self.doc.delete, edit=True)()
            return
        elif key in ("left", "arrow left", "arrowleft", "right", "arrow right", "arrowright", "up", "arrow up", "arrowup", "down", "arrow down", "arrowdown"):
            amount = 10 if shift else 1
            dx = -amount if "left" in key else amount if "right" in key else 0
            dy = -amount if "up" in key else amount if "down" in key else 0
            await self.handler(self.doc.move_selected, dx, dy, edit=True)()
            return
        elif key == "v":
            self.set_mode("select")
        elif key == "h":
            self.set_mode("pan")
        self.refresh(inspector=False)

    def select_layer(self, identifier):
        additive = self.has_key("control") or self.has_key("shift")
        self.doc.select(identifier, add=additive, toggle=additive)

    def select_all(self):
        self.doc.selected = [i.id for i in self.doc.items if i.visible]
        self.doc.anchor = self.doc.selected[-1] if self.doc.selected else None
        self.doc.label_selected = None

    def toggle_lock(self, identifier=None):
        targets = [self.doc.get(identifier)] if identifier is not None else [self.doc.get(i) for i in self.doc.selected]
        targets = [i for i in targets if i]
        if not targets:
            return
        value = not any(i.locked for i in targets)
        for item in targets:
            item.locked = value

    def toggle_visible(self, identifier):
        item = self.doc.get(identifier)
        if item:
            item.visible = not item.visible

    def rotate_selection(self, angle):
        origins = {i.id: copy.deepcopy(i) for u in self.doc.units() for i in u}
        if origins:
            self.apply_rotation_origins(origins, self.selection_center(), angle)

    async def undo(self):
        if self.busy:
            return
        if self.doc.undo():
            self.generation += 1
            self.set_mode("select")
            await self.ensure_previews()
            self.build_left()

    async def redo(self):
        if self.busy:
            return
        if self.doc.redo():
            self.generation += 1
            self.set_mode("select")
            await self.ensure_previews()
            self.build_left()

    async def on_canvas_unit_change(self, e):
        unit = self.canvas_unit.value
        factor = 1 if unit == "px" else self.doc.settings.export_dpi / (2.54 if unit == "cm" else 1)
        self.canvas_width.value = self.number(self.doc.settings.canvas_w / factor)
        self.canvas_height.value = self.number(self.doc.settings.canvas_h / factor)
        self.page.update()

    def apply_canvas(self):
        s = self.doc.settings
        factor = 1 if self.canvas_unit.value == "px" else s.export_dpi / (2.54 if self.canvas_unit.value == "cm" else 1)
        w, h = self.parse(self.canvas_width, 0.01, 100000) * factor, self.parse(self.canvas_height, 0.01, 100000) * factor
        if not 10 <= w <= 100000 or not 10 <= h <= 100000:
            raise ValueError("Canvas must be between 10 and 100,000 pixels on each side")
        ImageColor.getrgb(self.canvas_color.value)
        rgb = ImageColor.getrgb(self.canvas_color.value)
        s.canvas_w, s.canvas_h, s.bg_color = round(w), round(h), "#%02x%02x%02x" % rgb[:3]
        self.fit_view()

    def apply_spacing(self):
        s = self.doc.settings
        s.margin = self.parse(self.margin_field, 0, 5000)
        s.gap_x = self.parse(self.gap_x_field, 0, 5000)
        s.gap_y = self.parse(self.gap_y_field, 0, 5000)
        s.grid_step = self.parse(self.grid_step_field, 1, 5000)
        self.status.value = "Spacing updated. Existing panels move only when you choose an arrangement action."

    def custom_grid(self):
        rows, cols = self.parse(self.rows_field, 1, 1000), self.parse(self.cols_field, 1, 1000)
        if rows != int(rows) or cols != int(cols):
            raise ValueError("Rows and columns must be whole numbers")
        self.doc.grid(int(rows), int(cols))

    def crop_canvas(self):
        if not self.doc.items:
            return
        b = content_bbox(self.doc.items)
        m = self.doc.settings.margin
        for i in self.doc.items:
            i.x += m - b[0]
            i.y += m - b[1]
        self.doc.settings.canvas_w = max(10, math.ceil(b[2] - b[0] + 2 * m))
        self.doc.settings.canvas_h = max(10, math.ceil(b[3] - b[1] + 2 * m))
        self.fit_view()
        self.build_left()

    def apply_properties(self, identifier, label_only):
        item = self.doc.get(identifier)
        if not item or not self.doc.editable(item):
            raise ValueError("Unlock the object or its group before editing")
        p = self.prop
        candidate = copy.deepcopy(item)
        candidate.name = p["name"].value
        x, y = self.parse(p["x"]), self.parse(p["y"])
        if label_only:
            candidate.label_offset_x, candidate.label_offset_y = x - item.x, y - item.y
        else:
            candidate.x, candidate.y = x, y
            candidate.w = self.parse(p["w"], -100000 if item.is_line or item.is_stat else 1, 100000)
            candidate.h = self.parse(p["h"], -100000 if item.is_line or item.is_stat else 1, 100000)
            candidate.rotation = self.parse(p["rotation"]) % 360
        strings = ["label", "label_font_family", "label_color", "image_fit", "border_color", "font_family", "color", "fill", "align", "background", "outline", "text", "compare_style"]
        numbers = {"label_font_size": (1, 1000), "border_width": (0, 200), "font_size": (1, 1000), "stroke_width": (0, 200), "line_spacing": (0.5, 5), "bracket_height": (0, 10000), "text_gap": (0, 10000)}
        for key in strings:
            if key in p:
                value = str(p[key].value or "")
                if key in ("label_color", "border_color", "color", "fill", "background", "outline") and value:
                    ImageColor.getrgb(value)
                setattr(candidate, key, value)
        for key, limits in numbers.items():
            if key in p:
                setattr(candidate, key, self.parse(p[key], *limits))
        for key in ["show_label", "include_in_labels", "lock_aspect", "auto_center_text", "snap_to_tops"]:
            if key in p:
                setattr(candidate, key, bool(p[key].value))
        if "fill_alpha" in p:
            opacity = self.parse(p["fill_alpha"], 0, 100)
            if opacity != p["fill_alpha"].data:
                candidate.fill_alpha = round(opacity * 255 / 100)
        candidate.fill_alpha = int(candidate.fill_alpha)
        candidate.validate()
        item.__dict__.update(candidate.__dict__)
        self.status.value = "Properties applied."

    async def add_text(self, caption=False):
        if self.busy:
            return
        before = self.doc.snapshot()
        center = self.model((self.view_w / 2, self.view_h / 2))
        x, y, w, h = center[0] - 160, center[1] - 45, 320, 90
        text = "Caption" if caption else "Text"
        panels = [i for i in self.doc.items if i.kind == "panel"]
        if caption and panels:
            b = union(visual_bbox(i, True) for i in panels)
            x, y, w, h = b[0], b[3] + self.doc.settings.gap_y, max(200, b[2] - b[0]), 120
        item = self.doc.new_item(kind="text", x=x, y=y, w=w, h=h, text=text, rich_runs=[{"text": text, "bold": False, "italic": False}], font_family=self.doc.settings.default_font_family, name=text, lock_aspect=False)
        self.doc.expand_to_content()
        self.doc.checkpoint(before)
        self.refresh()
        await self.edit_text(item.id)

    def fit_text_height(self, identifier):
        item = self.doc.get(identifier)
        if item and item.kind == "text" and self.doc.editable(item):
            item.h = max(20, text_height(item))
            self.doc.expand_to_content()

    def remove_last_comparison(self):
        items = [i for i in self.doc.items if i.is_stat and self.doc.editable(i)]
        if items:
            last = max(items, key=lambda i: i.id)
            self.doc.items.remove(last)
            self.doc.selected = [i for i in self.doc.selected if i != last.id]
            self.doc.anchor = self.doc.selected[-1] if self.doc.selected else None

    def label_targets(self, selected=False):
        ids = set(self.doc.selected)
        return [i for i in self.doc.items if i.kind == "panel" and self.doc.editable(i) and (not selected or i.id in ids)]

    def toggle_labels(self):
        panels = self.label_targets(bool(self.doc.selected))
        show = any(not p.show_label for p in panels)
        for p in panels:
            p.show_label = show

    def include_labels(self, included):
        for p in self.label_targets(True):
            p.include_in_labels, p.show_label = included, included

    def apply_label_style(self, selected):
        s = self.doc.settings
        s.default_label_size = self.parse(self.label_size_field, 1, 1000)
        s.default_font_family = self.label_family_field.value
        s.default_label_offset_x = self.parse(self.label_x_field)
        s.default_label_offset_y = self.parse(self.label_y_field)
        for p in self.label_targets(selected):
            p.label_font_size, p.label_font_family = s.default_label_size, s.default_font_family
            p.label_offset_x, p.label_offset_y = s.default_label_offset_x, s.default_label_offset_y

    def position_labels(self, selected):
        position = self.label_position.value
        for p in self.label_targets(selected):
            w, h = text_extent(p.label or "A", p.label_font_family, p.label_font_size, True)
            locations = {"top-in": (6, 6), "top-out": (6, -h - 6), "left-in": (6, (p.h - h) / 2), "left-out": (-w - 6, (p.h - h) / 2), "bottom-in": (6, p.h - h - 6), "bottom-out": (6, p.h + 6), "center": ((p.w - w) / 2, (p.h - h) / 2)}
            p.label_offset_x, p.label_offset_y = locations[position]

    async def native_picker(self, method, **options):
        self.dialog_depth += 1
        self.keys.clear()
        self.update_drop_region()
        try:
            return await method(**options)
        finally:
            self.dialog_depth -= 1
            self.typing = False
            self.keys.clear()
            self.update_drop_region()

    async def pick_files(self, **options):
        return await self.native_picker(self.picker.pick_files, **options)

    async def save_file(self, **options):
        return await self.native_picker(self.picker.save_file, **options)

    async def choose(self, title, content, options):
        result = asyncio.get_running_loop().create_future()
        async def dismiss(e=None):
            if not result.done():
                result.set_result(None)
        def close_with(value):
            async def close(e):
                if not result.done():
                    result.set_result(value)
                self.page.pop_dialog()
                self.page.update()
            return close
        dialog = ft.AlertDialog(title=ft.Text(title), content=content, modal=True, actions=[self.button(label, close_with(value), primary=primary) for value, label, primary in options], on_dismiss=dismiss)
        self.dialog_depth += 1
        self.keys.clear()
        self.update_drop_region()
        self.page.show_dialog(dialog)
        try:
            return await result
        finally:
            self.dialog_depth -= 1
            self.typing = False
            self.keys.clear()
            self.update_drop_region()

    async def confirm_unsaved(self):
        if not self.doc.dirty:
            return True
        choice = await self.choose("Save this figure?", ft.Text("Your figure contains unsaved changes."), [("cancel", "Cancel", False), ("discard", "Discard", False), ("save", "Save", True)])
        if choice == "save":
            return await self.save_document()
        return choice == "discard"

    async def ensure_previews(self):
        missing = [i for i in self.doc.items if i.kind == "panel" and preview_key(i.to_dict()) not in self.previews]
        if not missing:
            return
        was_busy = self.busy
        self.busy = True
        try:
            for index, item in enumerate(missing):
                self.status.value = f"Preparing preview {index + 1}/{len(missing)}…"
                self.progress.value = index / len(missing)
                self.refresh(inspector=False)
                self.previews[preview_key(item.to_dict())] = await self.worker(make_preview, item.to_dict(), self.root)
        finally:
            self.busy = was_busy
            self.refresh()

    def cancel_operation(self):
        self.cancel_requested = True
        (Path(self.root) / "cancel-export").touch()
        self.status.value = "Stopping after the current file/page finishes…"

    async def pick_import(self):
        if self.busy:
            self.notice("Finish or cancel the current operation before importing more files.")
            return
        self.keys.clear()
        files = await self.pick_files(dialog_title="Add plots · images or PDFs", allow_multiple=True, file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["png", "jpg", "jpeg", "tif", "tiff", "bmp", "gif", "webp", "pdf"])
        paths = [f.path for f in files or [] if f.path]
        if paths:
            await self.import_paths(paths)

    async def import_paths(self, paths, drop_position=None):
        if self.busy:
            return
        self.busy, self.cancel_requested = True, False
        before = self.doc.snapshot()
        added, errors, new_ids = 0, [], []
        cursor = [max(0.0, drop_position[0]), max(0.0, drop_position[1])] if drop_position is not None else None
        row_start = cursor[0] if cursor else 0
        row_height = 0
        try:
            for index, path in enumerate(paths):
                if self.cancel_requested:
                    break
                self.status.value = f"Importing {index + 1}/{len(paths)} · {Path(path).name}"
                self.progress.value = index / max(1, len(paths))
                self.refresh(inspector=False)
                try:
                    staged = await self.worker(stage_source, str(path), self.root)
                    for page_number, size in enumerate(staged["sizes"]):
                        if self.cancel_requested:
                            break
                        self.status.value = f"Preparing {Path(path).name}" + (f" · page {page_number + 1}/{len(staged['sizes'])}" if len(staged["sizes"]) > 1 else "")
                        self.progress.value = (index + page_number / max(1, len(staged["sizes"]))) / max(1, len(paths))
                        self.refresh(inspector=False)
                        proxy = Item(id=0, asset=staged["asset"], original_size=size, source_page=page_number)
                        preview = await self.worker(make_preview, proxy.to_dict(), self.root)
                        canvas_size = self.doc.settings.canvas_w, self.doc.settings.canvas_h
                        item = self.doc.append_panel(staged["asset"], size, str(path), page_number)
                        if cursor is not None:
                            self.doc.settings.canvas_w, self.doc.settings.canvas_h = canvas_size
                            item.x, item.y = cursor
                            row_height = max(row_height, item.h)
                            cursor[0] += item.w + max(8, self.doc.settings.gap_x)
                            if cursor[0] + item.w > self.doc.settings.canvas_w - self.doc.settings.margin:
                                cursor = [row_start, cursor[1] + row_height + max(8, self.doc.settings.gap_y)]
                                row_height = 0
                            self.doc.expand_to_content()
                        self.previews[preview_key(item.to_dict())] = preview
                        added += 1
                        new_ids.append(item.id)
                        self.refresh()
                except Exception as exc:
                    logging.exception("Import failed: %s", path)
                    errors.append(f"{Path(path).name}: {exc}")
        finally:
            self.busy = False
            if before != self.doc.snapshot():
                self.doc.checkpoint(before)
            if new_ids:
                self.doc.selected, self.doc.anchor, self.doc.label_selected = new_ids, new_ids[-1], None
            self.status.value = f"{'Stopped' if self.cancel_requested else 'Imported'} · {added} panel(s). Existing layout and labels were preserved."
            self.refresh()
        if errors:
            self.notice(f"{added} panels imported; {len(errors)} file(s) could not be read.\n" + "\n".join(errors[:8]), error=True)

    async def paste(self):
        if self.busy:
            return
        self.busy = True
        self.status.value = "Reading clipboard…"
        self.refresh(inspector=False)
        try:
            staged = await self.worker(stage_clipboard, self.root)
            if "paths" in staged:
                self.busy = False
                await self.import_paths(staged["paths"])
                return
            before = self.doc.snapshot()
            proxy = Item(id=0, asset=staged["asset"], original_size=staged["sizes"][0])
            preview = await self.worker(make_preview, proxy.to_dict(), self.root)
            item = self.doc.append_panel(staged["asset"], staged["sizes"][0], "Clipboard")
            self.previews[preview_key(item.to_dict())] = preview
            self.doc.checkpoint(before)
            self.status.value = "Clipboard image added."
        finally:
            self.busy = False
            self.refresh()

    async def replace_selected(self, identifier=None):
        if self.busy:
            self.notice("Finish or cancel the current operation before replacing an image.")
            return
        identifier = identifier if identifier is not None else self.doc.anchor
        item = self.doc.get(identifier)
        if not item or item.kind != "panel" or not self.doc.editable(item):
            raise ValueError("Select an unlocked image panel")
        generation = self.generation
        files = await self.pick_files(dialog_title="Replace this image · keep frame and label", allow_multiple=False, file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["png", "jpg", "jpeg", "tif", "tiff", "bmp", "gif", "webp", "pdf"])
        if not files or not files[0].path:
            return
        path = files[0].path
        self.busy, self.cancel_requested = True, False
        self.status.value = "Preparing replacement without changing the figure…"
        self.refresh(inspector=False)
        try:
            staged = await self.worker(stage_source, path, self.root)
            page_number = 0
            if len(staged["sizes"]) > 1:
                field = self.field(f"Page · 1–{len(staged['sizes'])}", 1, 250)
                result = await self.choose("Choose replacement PDF page", field, [(None, "Cancel", False), (True, "Use page", True)])
                if not result:
                    return
                page_value = self.parse(field, 1, len(staged["sizes"]))
                if page_value != int(page_value):
                    raise ValueError("PDF page must be a whole number")
                page_number = int(page_value) - 1
            candidate = copy.deepcopy(item)
            candidate.asset, candidate.original_size, candidate.source_page = staged["asset"], staged["sizes"][page_number], page_number
            candidate.crop, candidate.erasures = [0, 0, 1, 1], []
            preview = await self.worker(make_preview, candidate.to_dict(), self.root)
            if self.cancel_requested or generation != self.generation or not self.doc.get(identifier):
                return
            before = self.doc.snapshot()
            self.doc.replace_image(identifier, candidate.asset, candidate.original_size, path, page_number)
            self.previews[preview_key(candidate.to_dict())] = preview
            self.doc.checkpoint(before)
            self.status.value = "Image replaced. Frame, label, layer, group and other objects are unchanged."
        finally:
            self.busy = False
            self.refresh()

    async def trim_selected(self, identifier=None):
        if self.busy:
            return
        identifier = identifier if identifier is not None else self.doc.anchor
        item = self.doc.get(identifier)
        if not item or item.kind != "panel" or not self.doc.editable(item):
            raise ValueError("Select an unlocked image panel to trim")
        self.busy = True
        self.status.value = "Finding white / transparent margins…"
        self.refresh(inspector=False)
        before = self.doc.snapshot()
        try:
            fraction = await self.worker(trim_source, item.to_dict(), self.root)
            if not fraction or all(abs(a - b) < 0.0001 for a, b in zip(fraction, [0, 0, 1, 1])):
                self.status.value = "No removable margins were found."
                return
            candidate = copy.deepcopy(item)
            apply_trim(candidate, fraction)
            preview = await self.worker(make_preview, candidate.to_dict(), self.root)
            item.__dict__.update(candidate.__dict__)
            self.previews[preview_key(item.to_dict())] = preview
            self.doc.checkpoint(before)
            self.status.value = "Margins trimmed non-destructively. The label stayed in place."
        finally:
            self.busy = False
            self.refresh()

    async def reset_image_edits(self, identifier=None):
        item = self.doc.get(identifier if identifier is not None else self.doc.anchor)
        if not item or item.kind != "panel" or not self.doc.editable(item):
            raise ValueError("Select an unlocked image panel")
        item.crop, item.erasures = [0, 0, 1, 1], []
        await self.ensure_previews()
        self.status.value = "Image pixels restored inside the current frame. Use Undo to restore an earlier frame size."

    async def new_document(self):
        if self.busy:
            self.notice("Finish or cancel the current operation before starting a new figure.")
            return
        if not await self.confirm_unsaved():
            return
        self.doc = Document()
        self.path = None
        self.generation += 1
        self.visual_controls.clear()
        self.label_controls.clear()
        self.byte_cache.clear()
        self.set_mode("select")
        self.build_left()
        self.fit_view()
        self.status.value = "New figure ready."

    async def save_document(self):
        if self.busy:
            self.notice("Finish the current operation before saving.")
            return False
        path = self.path if self.path and Path(self.path).suffix.lower() == ".figpie" else None
        if not path:
            path = await self.save_file(dialog_title="Save portable FigPie project", file_name="Figure.figpie", file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["figpie"])
        if not path:
            return False
        if Path(path).suffix.lower() != ".figpie":
            path += ".figpie"
        snapshot = self.doc.snapshot()
        self.busy = True
        self.status.value = "Saving original assets and layout…"
        self.refresh(inspector=False)
        try:
            await self.worker(save_project, snapshot, self.root, path)
            self.path = path
            self.doc.saved_state = snapshot
            self.status.value = f"Saved · {Path(path).name}"
            return True
        finally:
            self.busy = False
            self.refresh()

    async def open_document(self):
        if self.busy:
            self.notice("Finish or cancel the current operation before opening another figure.")
            return
        if not await self.confirm_unsaved():
            return
        files = await self.pick_files(dialog_title="Open FigPie or legacy Figboard project", allow_multiple=False, file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["figpie", "figboard", "json"])
        if not files or not files[0].path:
            return
        path = files[0].path
        self.busy, self.cancel_requested = True, False
        self.status.value = "Reading project and checking embedded assets…"
        self.refresh(inspector=False)
        try:
            snapshot = await self.worker(load_project, path, self.root)
            staged = Document()
            staged.restore(snapshot)
            panels = [i for i in staged.items if i.kind == "panel"]
            for index, item in enumerate(panels):
                if self.cancel_requested:
                    self.status.value = "Open cancelled. The previous figure is unchanged."
                    return
                key = preview_key(item.to_dict())
                if key not in self.previews:
                    self.status.value = f"Opening project · preview {index + 1}/{len(panels)}"
                    self.progress.value = index / max(1, len(panels))
                    self.refresh(inspector=False)
                    self.previews[key] = await self.worker(make_preview, item.to_dict(), self.root)
            if self.cancel_requested:
                return
            staged.saved_state = staged.snapshot()
            self.doc = staged
            self.path = path if Path(path).suffix.lower() == ".figpie" else None
            self.generation += 1
            self.visual_controls.clear()
            self.label_controls.clear()
            self.byte_cache.clear()
            self.set_mode("select")
            self.build_left()
            self.fit_view()
            self.status.value = f"Opened · {Path(path).name}" + (" · Save as .figpie to keep a migrated copy." if self.path is None else "")
        finally:
            self.busy = False
            self.refresh()

    async def export_dialog(self):
        if self.busy:
            return
        if not any(i.visible for i in self.doc.items):
            raise ValueError("Add a plot, text or shape before exporting")
        s = self.doc.settings
        fmt = self.dropdown("Format", "PNG", [(v, v) for v in ["PNG", "TIFF", "PDF", "SVG"]], 270)
        dpi = self.field("DPI · physical print density", s.export_dpi, 270)
        scale = self.field("Pixel scale · 1 = canvas pixel size", s.export_scale, 270)
        trim = ft.Checkbox(label="Fit output to content + margin", value=s.export_trim)
        size_note = ft.Text("", size=13, color=INK)
        async def update_note(e=None):
            try:
                local = copy.copy(s)
                local.export_scale = self.parse(scale, 0.1, 12)
                local.export_dpi = int(self.parse(dpi, 72, 2400))
                local.export_trim = bool(trim.value)
                b = export_bounds(self.doc.items, local)
                w, h = math.ceil((b[2] - b[0]) * local.export_scale), math.ceil((b[3] - b[1]) * local.export_scale)
                size_note.value = f"{w:,} × {h:,} pixels · {w * h / 1e6:.1f} MP\n{w / local.export_dpi * 2.54:.2f} × {h / local.export_dpi * 2.54:.2f} cm at {local.export_dpi} DPI"
            except ValueError:
                size_note.value = "Enter valid DPI and scale values."
            self.page.update()
        async def legacy(e):
            try:
                scale.value = self.number(max(1, self.parse(dpi, 72, 2400) / 100))
                await update_note()
            except ValueError as exc:
                size_note.value = str(exc)
                self.page.update()
        dpi.on_change, scale.on_change, trim.on_change = update_note, update_note, update_note
        b = export_bounds(self.doc.items, s)
        size_note.value = f"{math.ceil((b[2] - b[0]) * s.export_scale):,} × {math.ceil((b[3] - b[1]) * s.export_scale):,} pixels"
        body = ft.Container(width=440, content=ft.Column([fmt, dpi, scale, trim, size_note, self.button("", legacy), ft.Text("Increase Pixel scale for a larger export.", size=12, color=MUTED), ft.Text("", size=11, color=MUTED)], spacing=12, tight=True, scroll=ft.ScrollMode.AUTO))
        choice = await self.choose("Export figure", body, [(None, "Cancel", False), (True, "Choose file…", True)])
        if not choice:
            return
        export_settings = copy.copy(s)
        export_settings.export_scale = self.parse(scale, 0.1, 12)
        export_settings.export_dpi = int(self.parse(dpi, 72, 2400))
        export_settings.export_trim = bool(trim.value)
        b = export_bounds(self.doc.items, export_settings)
        check_pixels(math.ceil((b[2] - b[0]) * export_settings.export_scale), math.ceil((b[3] - b[1]) * export_settings.export_scale), MAX_OUTPUT_PIXELS)
        ext = {"PNG": "png", "TIFF": "tiff", "PDF": "pdf", "SVG": "svg"}[fmt.value]
        path = await self.save_file(dialog_title=f"Export {fmt.value}", file_name=f"Figure.{ext}", file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=[ext])
        if not path:
            return
        if Path(path).suffix.lower() != "." + ext:
            path += "." + ext
        before = self.doc.snapshot()
        self.doc.settings = export_settings
        if before != self.doc.snapshot():
            self.doc.checkpoint(before)
        snapshot = self.doc.snapshot()
        self.busy = True
        self.cancel_requested = False
        (Path(self.root) / "cancel-export").unlink(missing_ok=True)
        self.status.value = f"Rendering {fmt.value} from original assets…"
        self.progress.value = None
        self.refresh(inspector=False)
        try:
            await self.worker(export_file, snapshot, self.root, path, fmt.value)
            self.status.value = f"Exported · {Path(path).name}"
        except InterruptedError:
            self.status.value = "Export cancelled. Existing files were not overwritten."
        finally:
            self.busy = False
            self.progress.value = 0
            self.refresh()

    async def edit_label(self, identifier):
        if self.busy:
            return
        item = self.doc.get(identifier)
        if not item or item.kind != "panel" or not self.doc.editable(item):
            return
        value = self.field("Panel label", item.label, 320)
        choice = await self.choose("Edit panel label", value, [(None, "Cancel", False), (True, "Apply", True)])
        if choice:
            if len(value.value or "") > 64:
                raise ValueError("Panel labels may contain at most 64 characters")
            before = self.doc.snapshot()
            item.label, item.show_label, item.include_in_labels = value.value or "", True, True
            self.doc.checkpoint(before)
            self.doc.select(item.id, label=True)

    async def edit_text(self, identifier):
        if self.busy:
            return
        item = self.doc.get(identifier)
        if not item or item.kind != "text" or not self.doc.editable(item):
            return
        chars = [(ch, bool(r.get("bold")), bool(r.get("italic"))) for r in item.rich_runs for ch in r["text"]]
        selected = [0, 0]
        editor = self.field("Text", "".join(c[0] for c in chars), multiline=True)
        editor.min_lines, editor.max_lines = 7, 10
        editor.width = 710
        preview = ft.Text(spans=[], size=17, selectable=True)
        font_field = self.dropdown("Font", item.font_family, [(f, f) for f in FONT_FAMILIES], 215)
        size_field = self.field("Size", item.font_size, 104)
        align_field = self.dropdown("Align", item.align, [(a, a.title()) for a in ["left", "center", "right", "justify"]], 130)
        spacing_field = self.field("Line spacing", item.line_spacing, 120)
        def get_runs():
            return normalize_runs([{"text": c, "bold": b, "italic": i} for c, b, i in chars])
        def update_preview():
            preview.spans = [ft.TextSpan(text=r["text"], style=ft.TextStyle(weight=ft.FontWeight.BOLD if r["bold"] else ft.FontWeight.NORMAL, italic=r["italic"])) for r in get_runs()]
        def reconcile(new):
            old = "".join(c[0] for c in chars)
            if old == new:
                return
            updated = []
            for tag, a, b, c, d in difflib.SequenceMatcher(a=old, b=new, autojunk=False).get_opcodes():
                if tag == "equal":
                    updated.extend(chars[a:b])
                elif tag in ("replace", "insert"):
                    style = chars[min(max(0, a - 1), len(chars) - 1)][1:] if chars else (False, False)
                    updated.extend((ch, *style) for ch in new[c:d])
            chars[:] = updated
        async def changed(e):
            new = editor.value or ""
            if len(new) > 50000:
                editor.error_text = "Text boxes are limited to 50,000 characters."
            else:
                editor.error_text = None
                reconcile(new)
                update_preview()
            self.page.update()
        def utf16_index(text, offset):
            count = 0
            for index, ch in enumerate(text):
                if count >= offset:
                    return index
                count += 2 if ord(ch) > 0xFFFF else 1
            return len(text)
        async def selection_changed(e):
            if e.selection:
                text = editor.value or ""
                selected[:] = [utf16_index(text, e.selection.start), utf16_index(text, e.selection.end)]
        def style_action(bold, italic):
            async def apply(e):
                reconcile(editor.value or "")
                a, b = sorted(selected)
                if a == b:
                    a, b = 0, len(chars)
                for index in range(a, min(b, len(chars))):
                    chars[index] = chars[index][0], bold, italic
                update_preview()
                self.page.update()
            return apply
        editor.on_change = changed
        editor.on_selection_change = selection_changed
        update_preview()
        body = ft.Container(width=730, content=ft.Column([self.row(font_field, size_field, align_field, spacing_field), self.row(self.button("Bold", style_action(True, False), "FORMAT_BOLD"), self.button("Italic", style_action(False, True), "FORMAT_ITALIC"), self.button("Bold + italic", style_action(True, True)), self.button("Normal", style_action(False, False))), ft.Text("Select text to format it; no selection applies the style to all text. The formatted preview is below the editor.", size=11, color=MUTED), editor, ft.Text("FORMATTED PREVIEW", size=10, color=MUTED), ft.Container(content=preview, padding=12, border=ft.Border.all(1, BORDER), border_radius=8)], tight=True, spacing=10, scroll=ft.ScrollMode.AUTO))
        result = await self.choose("Rich text editor", body, [(None, "Cancel", False), (True, "Save text", True)])
        if result:
            if len(editor.value or "") > 50000:
                raise ValueError("Text boxes are limited to 50,000 characters")
            reconcile(editor.value or "")
            candidate = copy.deepcopy(item)
            candidate.rich_runs = get_runs()
            candidate.font_family = font_field.value
            candidate.font_size = self.parse(size_field, 1, 1000)
            candidate.align = align_field.value
            candidate.line_spacing = self.parse(spacing_field, 0.5, 5)
            candidate.validate()
            candidate.h = max(candidate.h, text_height(candidate))
            before = self.doc.snapshot()
            item.__dict__.update(candidate.__dict__)
            self.doc.expand_to_content()
            self.doc.checkpoint(before)

    async def help_dialog(self):
        text = """Import with Add plots, or paste an image with Ctrl+V. 

Right-click a panel and choose Replace image.

Drag objects to move them. Shift/Ctrl click adds to the selection. The last selected object is the alignment anchor. Drag a corner to resize or the round handle to rotate. Shift snaps rotation to 15°; Alt bypasses snapping.

Auto H makes a row; Auto V makes a column. Gap H/V applies the chosen gap without cross-axis alignment. Equal H/V keeps the outer endpoints and shares the intervening space.

Click or drag labels independently. Double-click a label or text to edit.

Choose a shape tool, then drag to draw or click for a default size. Statistics: choose Bracket and click exactly two data positions. The first click places  marker; move the mouse to preview, then click the second endpoint. Keep tool active allows consecutive comparisons. Esc cancels an unfinished bracket. Crop and Erase stay visibly active until Esc; both are reversible.
 

Developed by Hamid Taghipour | UiT The Arctic University of Norway."""
        await self.choose(f"FigPie {__version__} · Help", ft.Container(width=650, height=530, content=ft.Column([ft.Text(text, size=13, selectable=True)], scroll=ft.ScrollMode.AUTO)), [(True, "Close", True)])

    async def on_window_event(self, e):
        event = str(getattr(e, "type", getattr(e, "data", ""))).lower()
        if "blur" in event or "minimize" in event:
            self.keys.clear()
        if not (event.endswith("close") or event == "close"):
            return
        if self.dialog_depth:
            return
        if self.busy:
            self.cancel_requested = True
            self.notice("The worker is finishing the current operation. Close again once it has stopped.")
            return
        if not await self.confirm_unsaved():
            return
        self.closed = True
        self.file_drop.stop()
        self.pool.shutdown(wait=False, cancel_futures=True)
        shutil.rmtree(self.root, ignore_errors=True)
        self.page.window.prevent_close = False
        await self.page.window.close()


async def main(page: ft.Page):
    if page.web:
        page.add(ft.Text("FigPie requires a native desktop window. Run this Python file directly from your activated environment."))
        return
    directory = Path.home() / ".figpie"
    directory.mkdir(exist_ok=True)
    logging.basicConfig(filename=directory / "figpie.log", level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    try:
        app = FigPieApp(page)
        page.data = app
        page.run_task(app.desktop_drop_loop)
    except Exception:
        logging.exception("FigPie could not start")
        raise


if __name__ == "__main__":
    multiprocessing.freeze_support()
    ft.run(main)
