#!/usr/bin/env python3
"""
Generate a lightweight PNG preview from a .vsdx file.
"""

import struct
import zlib
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"v": "http://schemas.microsoft.com/office/visio/2012/main"}


def _float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _get_page_size(pages_root):
    page = pages_root.find(".//v:Page", NS)
    if page is None:
        return 11.0, 8.5
    sheet = page.find("v:PageSheet", NS)
    width = height = None
    if sheet is not None:
        for cell in sheet.findall(".//v:Cell", NS):
            if cell.get("N") == "PageWidth":
                width = _float(cell.get("V"))
            if cell.get("N") == "PageHeight":
                height = _float(cell.get("V"))
    return width or 11.0, height or 8.5


def _get_master_names(masters_root):
    master_names = {}
    for master in masters_root.findall(".//v:Master", NS):
        master_id = master.get("ID")
        name = master.get("NameU") or master.get("Name")
        if master_id and name:
            master_names[master_id] = name
    return master_names


def _extract_text(shape):
    text_el = shape.find("v:Text", NS)
    if text_el is None:
        return ""
    return "".join(text_el.itertext()).strip()


class Canvas:
    def __init__(self, width, height, bg=(255, 255, 255)):
        self.width = int(width)
        self.height = int(height)
        self.pixels = bytearray(self.width * self.height * 3)
        self.fill(bg)

    def fill(self, color):
        r, g, b = color
        for i in range(0, len(self.pixels), 3):
            self.pixels[i] = r
            self.pixels[i + 1] = g
            self.pixels[i + 2] = b

    def set_pixel(self, x, y, color):
        if x < 0 or y < 0 or x >= self.width or y >= self.height:
            return
        idx = (y * self.width + x) * 3
        self.pixels[idx:idx + 3] = bytes(color)

    def draw_line(self, x0, y0, x1, y1, color):
        x0 = int(round(x0))
        y0 = int(round(y0))
        x1 = int(round(x1))
        y1 = int(round(y1))
        dx = abs(x1 - x0)
        dy = -abs(y1 - y0)
        sx = 1 if x0 < x1 else -1
        sy = 1 if y0 < y1 else -1
        err = dx + dy
        while True:
            self.set_pixel(x0, y0, color)
            if x0 == x1 and y0 == y1:
                break
            e2 = 2 * err
            if e2 >= dy:
                err += dy
                x0 += sx
            if e2 <= dx:
                err += dx
                y0 += sy

    def draw_rect(self, x, y, w, h, color):
        x0 = int(round(x))
        y0 = int(round(y))
        x1 = int(round(x + w))
        y1 = int(round(y + h))
        for xi in range(x0, x1 + 1):
            self.set_pixel(xi, y0, color)
            self.set_pixel(xi, y1, color)
        for yi in range(y0, y1 + 1):
            self.set_pixel(x0, yi, color)
            self.set_pixel(x1, yi, color)

    def fill_rect(self, x, y, w, h, color):
        x0 = int(round(x))
        y0 = int(round(y))
        x1 = int(round(x + w))
        y1 = int(round(y + h))
        for yi in range(y0, y1):
            for xi in range(x0, x1):
                self.set_pixel(xi, yi, color)

    def draw_polygon(self, points, color):
        for i in range(len(points)):
            x0, y0 = points[i]
            x1, y1 = points[(i + 1) % len(points)]
            self.draw_line(x0, y0, x1, y1, color)

    def to_png(self):
        raw = bytearray()
        for y in range(self.height):
            raw.append(0)
            start = y * self.width * 3
            raw.extend(self.pixels[start:start + self.width * 3])

        compressed = zlib.compress(bytes(raw), level=9)

        def chunk(tag, data):
            return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

        header = struct.pack(">IIBBBBB", self.width, self.height, 8, 2, 0, 0, 0)
        return b"".join([
            b"\x89PNG\r\n\x1a\n",
            chunk(b"IHDR", header),
            chunk(b"IDAT", compressed),
            chunk(b"IEND", b""),
        ])


def render_preview(vsdx_path: Path, output_path: Path, scale: float = 40.0) -> None:
    with zipfile.ZipFile(vsdx_path, "r") as z:
        page_xml = z.read("visio/pages/page1.xml")
        pages_xml = z.read("visio/pages/pages.xml")
        masters_xml = z.read("visio/masters/masters.xml")

    page_root = ET.fromstring(page_xml)
    pages_root = ET.fromstring(pages_xml)
    masters_root = ET.fromstring(masters_xml)

    page_width, page_height = _get_page_size(pages_root)
    master_names = _get_master_names(masters_root)

    canvas = Canvas(page_width * scale, page_height * scale, bg=(255, 255, 255))

    for shape in page_root.findall(".//v:Shape", NS):
        cells = {cell.get("N"): cell.get("V") for cell in shape.findall("v:Cell", NS)}
        pin_x = _float(cells.get("PinX"))
        pin_y = _float(cells.get("PinY"))
        width = _float(cells.get("Width"))
        height = _float(cells.get("Height"))

        begin_x = _float(cells.get("BeginX"))
        begin_y = _float(cells.get("BeginY"))
        end_x = _float(cells.get("EndX"))
        end_y = _float(cells.get("EndY"))
        is_connector = "BeginX" in cells and "EndX" in cells

        if is_connector:
            x0 = begin_x * scale
            y0 = (page_height - begin_y) * scale
            x1 = end_x * scale
            y1 = (page_height - end_y) * scale
            canvas.draw_line(x0, y0, x1, y1, (120, 120, 120))
            continue

        if not width or not height:
            continue

        x = (pin_x - width / 2.0) * scale
        y = (page_height - pin_y - height / 2.0) * scale

        master_name = master_names.get(shape.get("Master"), "")

        if "Decision" in master_name or "Diamond" in master_name:
            cx = x + (width * scale) / 2.0
            cy = y + (height * scale) / 2.0
            points = [
                (cx, y),
                (x + width * scale, cy),
                (cx, y + height * scale),
                (x, cy),
            ]
            canvas.draw_polygon(points, (160, 160, 160))
        else:
            canvas.draw_rect(x, y, width * scale, height * scale, (170, 170, 170))

        text = _extract_text(shape)
        if text:
            # Mark center point for visibility.
            center_x = int(round(x + (width * scale) / 2.0))
            center_y = int(round(y + (height * scale) / 2.0))
            for dx in range(-1, 2):
                for dy in range(-1, 2):
                    canvas.set_pixel(center_x + dx, center_y + dy, (80, 80, 80))

    output_path.write_bytes(canvas.to_png())


if __name__ == "__main__":
    render_preview(Path("output.vsdx"), Path("output_preview.png"))
