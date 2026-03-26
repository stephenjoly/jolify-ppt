#!/usr/bin/env python3
"""
Generate Jolify's custom semantic icons from oversized master artwork.

This script intentionally targets only the weak custom icon families that are
shared between the taskpane and, in some cases, the ribbon.

Usage:
  python3 scripts/generate-custom-icons.py
  python3 scripts/generate-custom-icons.py --contact-sheet /tmp/jolify-custom-icons.png
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Callable

from PIL import Image, ImageDraw

MASTER = 256
SIZES = (16, 32, 80)
OUT_DIR = Path(__file__).resolve().parents[1] / "assets" / "icons"

STROKE = "#757575"
BLUE = "#3a7cf2"
RED = "#d05643"
PANEL = "#ffffff"
GUIDE = "#d9d9d9"
TRANSPARENT = (0, 0, 0, 0)


def px(value: float) -> int:
    return round(value * MASTER / 100)


def make_canvas() -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGBA", (MASTER, MASTER), TRANSPARENT)
    return image, ImageDraw.Draw(image)


def save_sizes(master: Image.Image, name: str) -> None:
    for size in SIZES:
        out = master.resize((size, size), resample=Image.Resampling.LANCZOS)
        out.save(OUT_DIR / f"{name}_{size}.png")


def rounded_box(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    *,
    fill: str | None = None,
    outline: str | None = STROKE,
    width: int = 10,
    radius: int = 18,
) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def left_arrow(draw: ImageDraw.ImageDraw, x1: int, x2: int, y: int, *, color: str = BLUE, width: int = 14) -> None:
    draw.line((x1, y, x2, y), fill=color, width=width)
    head = px(7)
    draw.line((x2, y, x2 - head, y - head + 2), fill=color, width=width)
    draw.line((x2, y, x2 - head, y + head - 2), fill=color, width=width)


def right_arrow(draw: ImageDraw.ImageDraw, x1: int, x2: int, y: int, *, color: str = BLUE, width: int = 14) -> None:
    draw.line((x1, y, x2, y), fill=color, width=width)
    head = px(7)
    draw.line((x2, y, x2 - head, y - head + 2), fill=color, width=width)
    draw.line((x2, y, x2 - head, y + head - 2), fill=color, width=width)


def slash(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int]) -> None:
    x0, y0, x1, y1 = box
    draw.line((x0 + px(4), y1 - px(4), x1 - px(4), y0 + px(4)), fill=RED, width=16)


def guide_corner(draw: ImageDraw.ImageDraw, x: int, y: int, *, flip_x: bool = False, flip_y: bool = False) -> None:
    arm = px(12)
    width = 10
    x2 = x - arm if flip_x else x + arm
    y2 = y - arm if flip_y else y + arm
    draw.line((x, y, x2, y), fill=GUIDE, width=width)
    draw.line((x, y, x, y2), fill=GUIDE, width=width)


def grouped_frame(draw: ImageDraw.ImageDraw) -> None:
    x0, y0, x1, y1 = px(18), px(18), px(82), px(82)
    width = 10
    dash = px(8)
    gap = px(5)
    for start in range(x0, x1, dash + gap):
        draw.line((start, y0, min(start + dash, x1), y0), fill=GUIDE, width=width)
        draw.line((start, y1, min(start + dash, x1), y1), fill=GUIDE, width=width)
    for start in range(y0, y1, dash + gap):
        draw.line((x0, start, x0, min(start + dash, y1)), fill=GUIDE, width=width)
        draw.line((x1, start, x1, min(start + dash, y1)), fill=GUIDE, width=width)


def draw_shape(
    draw: ImageDraw.ImageDraw,
    *,
    box: tuple[int, int, int, int],
    fill: bool,
    outline: bool = True,
    blue_outline: bool = False,
) -> None:
    rounded_box(
        draw,
        box,
        fill=BLUE if fill else PANEL,
        outline=BLUE if blue_outline else (STROKE if outline else None),
        width=12,
        radius=22,
    )


def copy_fill() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(16), px(38), px(40), px(62)), fill=True)
    rounded_box(draw, (px(55), px(20), px(78), px(44)), fill=PANEL, outline=STROKE, width=10, radius=16)
    right_arrow(draw, px(43), px(56), px(50))
    return image


def paste_fill() -> Image.Image:
    image, draw = make_canvas()
    rounded_box(draw, (px(18), px(20), px(41), px(44)), fill=PANEL, outline=STROKE, width=10, radius=16)
    draw_shape(draw, box=(px(58), px(38), px(82), px(62)), fill=True)
    right_arrow(draw, px(43), px(56), px(50))
    return image


def copy_outline() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(16), px(38), px(40), px(62)), fill=False)
    rounded_box(draw, (px(55), px(20), px(78), px(44)), fill=PANEL, outline=STROKE, width=10, radius=16)
    right_arrow(draw, px(43), px(56), px(50))
    return image


def paste_outline() -> Image.Image:
    image, draw = make_canvas()
    rounded_box(draw, (px(18), px(20), px(41), px(44)), fill=PANEL, outline=STROKE, width=10, radius=16)
    draw_shape(draw, box=(px(58), px(38), px(82), px(62)), fill=False)
    right_arrow(draw, px(43), px(56), px(50))
    return image


def match_fill() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(12), px(38), px(33), px(59)), fill=True)
    draw_shape(draw, box=(px(67), px(38), px(88), px(59)), fill=True)
    right_arrow(draw, px(39), px(60), px(49))
    return image


def match_outline() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(12), px(38), px(33), px(59)), fill=False)
    draw_shape(draw, box=(px(67), px(38), px(88), px(59)), fill=False)
    right_arrow(draw, px(39), px(60), px(49))
    return image


def match_style() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(12), px(34), px(33), px(58)), fill=True)
    draw.line((px(15), px(65), px(31), px(65)), fill=STROKE, width=10)
    draw_shape(draw, box=(px(67), px(34), px(88), px(58)), fill=True)
    draw.line((px(70), px(65), px(86), px(65)), fill=STROKE, width=10)
    right_arrow(draw, px(39), px(60), px(47))
    return image


def no_fill() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(22), px(22), px(78), px(78)), fill=True)
    slash(draw, (px(22), px(22), px(78), px(78)))
    return image


def no_outline() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(22), px(22), px(78), px(78)), fill=False)
    slash(draw, (px(22), px(22), px(78), px(78)))
    return image


def copy_position_only() -> Image.Image:
    image, draw = make_canvas()
    guide_corner(draw, px(12), px(22))
    guide_corner(draw, px(40), px(62), flip_x=True, flip_y=True)
    draw_shape(draw, box=(px(16), px(26), px(40), px(50)), fill=True)
    rounded_box(draw, (px(58), px(34), px(82), px(58)), fill=PANEL, outline=STROKE, width=10, radius=18)
    right_arrow(draw, px(43), px(56), px(46))
    return image


def paste_position_only() -> Image.Image:
    image, draw = make_canvas()
    rounded_box(draw, (px(18), px(22), px(42), px(46)), fill=PANEL, outline=STROKE, width=10, radius=18)
    guide_corner(draw, px(58), px(30))
    guide_corner(draw, px(86), px(70), flip_x=True, flip_y=True)
    draw_shape(draw, box=(px(58), px(34), px(82), px(58)), fill=True)
    right_arrow(draw, px(45), px(56), px(46))
    return image


def copy_size_only() -> Image.Image:
    image, draw = make_canvas()
    draw_shape(draw, box=(px(14), px(40), px(34), px(60)), fill=True)
    rounded_box(draw, (px(56), px(30), px(86), px(70)), fill=PANEL, outline=STROKE, width=10, radius=18)
    right_arrow(draw, px(38), px(52), px(50))
    draw.line((px(56), px(76), px(86), px(76)), fill=GUIDE, width=10)
    draw.line((px(56), px(76), px(63), px(69)), fill=GUIDE, width=10)
    draw.line((px(56), px(76), px(63), px(83)), fill=GUIDE, width=10)
    draw.line((px(86), px(76), px(79), px(69)), fill=GUIDE, width=10)
    draw.line((px(86), px(76), px(79), px(83)), fill=GUIDE, width=10)
    return image


def paste_size_only() -> Image.Image:
    image, draw = make_canvas()
    rounded_box(draw, (px(14), px(30), px(44), px(70)), fill=PANEL, outline=STROKE, width=10, radius=18)
    draw.line((px(14), px(76), px(44), px(76)), fill=GUIDE, width=10)
    draw.line((px(14), px(76), px(21), px(69)), fill=GUIDE, width=10)
    draw.line((px(14), px(76), px(21), px(83)), fill=GUIDE, width=10)
    draw.line((px(44), px(76), px(37), px(69)), fill=GUIDE, width=10)
    draw.line((px(44), px(76), px(37), px(83)), fill=GUIDE, width=10)
    draw_shape(draw, box=(px(62), px(40), px(82), px(60)), fill=True)
    right_arrow(draw, px(48), px(58), px(50))
    return image


def copy_position_and_size() -> Image.Image:
    image, draw = make_canvas()
    guide_corner(draw, px(10), px(18))
    guide_corner(draw, px(42), px(66), flip_x=True, flip_y=True)
    draw_shape(draw, box=(px(14), px(24), px(42), px(52)), fill=True)
    rounded_box(draw, (px(58), px(28), px(86), px(60)), fill=PANEL, outline=STROKE, width=10, radius=18)
    right_arrow(draw, px(45), px(56), px(46))
    draw.line((px(58), px(68), px(86), px(68)), fill=GUIDE, width=10)
    draw.line((px(58), px(68), px(65), px(61)), fill=GUIDE, width=10)
    draw.line((px(58), px(68), px(65), px(75)), fill=GUIDE, width=10)
    draw.line((px(86), px(68), px(79), px(61)), fill=GUIDE, width=10)
    draw.line((px(86), px(68), px(79), px(75)), fill=GUIDE, width=10)
    return image


def paste_position_and_size() -> Image.Image:
    image, draw = make_canvas()
    rounded_box(draw, (px(14), px(22), px(42), px(54)), fill=PANEL, outline=STROKE, width=10, radius=18)
    draw.line((px(14), px(64), px(42), px(64)), fill=GUIDE, width=10)
    draw.line((px(14), px(64), px(21), px(57)), fill=GUIDE, width=10)
    draw.line((px(14), px(64), px(21), px(71)), fill=GUIDE, width=10)
    draw.line((px(42), px(64), px(35), px(57)), fill=GUIDE, width=10)
    draw.line((px(42), px(64), px(35), px(71)), fill=GUIDE, width=10)
    guide_corner(draw, px(58), px(30))
    guide_corner(draw, px(90), px(78), flip_x=True, flip_y=True)
    draw_shape(draw, box=(px(58), px(36), px(86), px(64)), fill=True)
    right_arrow(draw, px(45), px(56), px(44))
    return image


def layout_icon(orientation: str, both: bool = False, distribute: bool = False) -> Image.Image:
    image, draw = make_canvas()
    grouped_frame(draw)
    boxes = [
        (px(26), px(28), px(42), px(44)),
        (px(54), px(44), px(70), px(60)),
        (px(72), px(26), px(88), px(42)),
    ]
    for box in boxes:
        draw_shape(draw, box=box, fill=True)
    if both:
        draw.line((px(50), px(12), px(50), px(88)), fill=BLUE, width=12)
        draw.line((px(12), px(50), px(88), px(50)), fill=BLUE, width=12)
    elif orientation == "horizontal":
        y = px(50)
        draw.line((px(12), y, px(88), y), fill=BLUE, width=12)
        if distribute:
            draw.line((px(32), px(26), px(32), px(74)), fill=GUIDE, width=10)
            draw.line((px(60), px(26), px(60), px(74)), fill=GUIDE, width=10)
    else:
        x = px(50)
        draw.line((x, px(12), x, px(88)), fill=BLUE, width=12)
        if distribute:
            draw.line((px(26), px(32), px(74), px(32)), fill=GUIDE, width=10)
            draw.line((px(26), px(60), px(74), px(60)), fill=GUIDE, width=10)
    return image


def align_center_h_and_group() -> Image.Image:
    return layout_icon("horizontal", distribute=False)


def align_middle_v_and_group() -> Image.Image:
    return layout_icon("vertical", distribute=False)


def center_middle_and_group() -> Image.Image:
    return layout_icon("horizontal", both=True)


def distribute_h_and_group() -> Image.Image:
    return layout_icon("horizontal", distribute=True)


def distribute_v_and_group() -> Image.Image:
    return layout_icon("vertical", distribute=True)


def distribute_hv_and_group() -> Image.Image:
    return layout_icon("horizontal", both=True, distribute=False)


def textbox(draw: ImageDraw.ImageDraw) -> tuple[int, int, int, int]:
    box = (px(18), px(18), px(82), px(82))
    rounded_box(draw, box, fill=PANEL, outline=STROKE, width=10, radius=18)
    return box


def text_lines(draw: ImageDraw.ImageDraw, left: int, right: int, ys: list[int]) -> None:
    for y in ys:
        draw.line((left, y, right, y), fill=STROKE, width=10)


def text_margins_none() -> Image.Image:
    image, draw = make_canvas()
    x0, y0, x1, y1 = textbox(draw)
    text_lines(draw, x0 + px(4), x1 - px(4), [px(34), px(48), px(62)])
    return image


def text_margins_tight() -> Image.Image:
    image, draw = make_canvas()
    x0, y0, x1, y1 = textbox(draw)
    text_lines(draw, x0 + px(10), x1 - px(10), [px(34), px(48), px(62)])
    return image


def text_margins_roomy() -> Image.Image:
    image, draw = make_canvas()
    x0, y0, x1, y1 = textbox(draw)
    text_lines(draw, x0 + px(18), x1 - px(18), [px(34), px(48), px(62)])
    return image


def autofit_off() -> Image.Image:
    image, draw = make_canvas()
    textbox(draw)
    draw.line((px(30), px(38), px(70), px(38)), fill=STROKE, width=10)
    draw.line((px(30), px(54), px(60), px(54)), fill=STROKE, width=10)
    slash(draw, (px(18), px(18), px(82), px(82)))
    return image


def autofit_text() -> Image.Image:
    image, draw = make_canvas()
    textbox(draw)
    draw.line((px(30), px(38), px(70), px(38)), fill=STROKE, width=10)
    draw.line((px(30), px(54), px(60), px(54)), fill=STROKE, width=10)
    draw.line((px(76), px(30), px(76), px(70)), fill=BLUE, width=10)
    draw.line((px(76), px(30), px(70), px(36)), fill=BLUE, width=10)
    draw.line((px(76), px(30), px(82), px(36)), fill=BLUE, width=10)
    draw.line((px(76), px(70), px(70), px(64)), fill=BLUE, width=10)
    draw.line((px(76), px(70), px(82), px(64)), fill=BLUE, width=10)
    return image


def autofit_shape() -> Image.Image:
    image, draw = make_canvas()
    textbox(draw)
    draw.line((px(32), px(40), px(68), px(40)), fill=STROKE, width=10)
    draw.line((px(32), px(56), px(56), px(56)), fill=STROKE, width=10)
    draw.line((px(18), px(90), px(82), px(90)), fill=BLUE, width=10)
    draw.line((px(18), px(90), px(26), px(82)), fill=BLUE, width=10)
    draw.line((px(18), px(90), px(26), px(98)), fill=BLUE, width=10)
    draw.line((px(82), px(90), px(74), px(82)), fill=BLUE, width=10)
    draw.line((px(82), px(90), px(74), px(98)), fill=BLUE, width=10)
    return image


GENERATORS: dict[str, Callable[[], Image.Image]] = {
    "copyFillSemantic": copy_fill,
    "pasteFillSemantic": paste_fill,
    "copyOutlineSemantic": copy_outline,
    "pasteOutlineSemantic": paste_outline,
    "matchFillSemantic": match_fill,
    "matchOutlineSemantic": match_outline,
    "matchStyleSemantic": match_style,
    "noFillSemantic": no_fill,
    "noOutlineSemantic": no_outline,
    "copyPositionOnly": copy_position_only,
    "pastePositionOnly": paste_position_only,
    "copySizeOnly": copy_size_only,
    "pasteSizeOnly": paste_size_only,
    "copyPositionAndSize": copy_position_and_size,
    "pastePositionAndSize": paste_position_and_size,
    "alignCenterHAndGroup": align_center_h_and_group,
    "alignMiddleVAndGroup": align_middle_v_and_group,
    "centerMiddleAndGroup": center_middle_and_group,
    "distributeHAndGroup": distribute_h_and_group,
    "distributeVAndGroup": distribute_v_and_group,
    "distributeHVAndGroupSemantic": distribute_hv_and_group,
    "textMarginsNoneSemantic": text_margins_none,
    "textMarginsTightSemantic": text_margins_tight,
    "textMarginsRoomySemantic": text_margins_roomy,
    "autofitOffSemantic": autofit_off,
    "autofitTextSemantic": autofit_text,
    "autofitShapeSemantic": autofit_shape,
}


def build_contact_sheet(path: Path) -> None:
    names = list(GENERATORS)
    cell = 140
    cols = 4
    rows = (len(names) + cols - 1) // cols
    sheet = Image.new("RGBA", (cols * cell + 24, rows * cell + 24), "white")
    draw = ImageDraw.Draw(sheet)
    for idx, name in enumerate(names):
        icon = Image.open(OUT_DIR / f"{name}_80.png").convert("RGBA")
        x = 12 + (idx % cols) * cell + 30
        y = 12 + (idx // cols) * cell + 8
        sheet.alpha_composite(icon, (x, y))
        draw.text((12 + (idx % cols) * cell + 8, 12 + (idx // cols) * cell + 104), name, fill="black")
    sheet.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--contact-sheet", type=Path, help="Optional output path for an 80px contact sheet preview.")
    args = parser.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for name, generator in GENERATORS.items():
        save_sizes(generator(), name)
        print(f"generated {name}")
    if args.contact_sheet:
        build_contact_sheet(args.contact_sheet)
        print(f"wrote contact sheet to {args.contact_sheet}")


if __name__ == "__main__":
    main()
