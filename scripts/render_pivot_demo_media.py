from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math


ROOT = Path(__file__).resolve().parents[1]
MEDIA_DIR = ROOT / "media"
MEDIA_DIR.mkdir(exist_ok=True)

W, H = 1280, 720
BG = "#f6f1e6"
SURFACE = "#fffdf8"
LINE = "#d8d0bf"
TEXT = "#2c302a"
MUTED = "#6f7669"
ACCENT = "#5a7d58"
BLUE = "#6aa9ff"
AMBER = "#efb74f"
GREEN = "#51b36f"
PURPLE = "#a482ff"
SHADOW = (44, 52, 46, 24)
DIM = (42, 48, 42)
DURATION_SCALE = 1.4

FONT_SANS = "/System/Library/Fonts/STHeiti Medium.ttc"
FONT_SANS_LIGHT = "/System/Library/Fonts/STHeiti Light.ttc"
FONT_MONO = "/System/Library/Fonts/SFNSMono.ttf"

SHEET_BOX = (92, 166, 850, 626)
RAW_GRID_BOX = (114, 240, 850, 574)
PIVOT_GRID_BOX = (114, 240, 828, 574)
SIDEBAR_BOX = (892, 182, 1184, 616)


def f(path, size):
    return ImageFont.truetype(path, size)


FONT_TITLE = f(FONT_SANS, 34)
FONT_BODY = f(FONT_SANS_LIGHT, 20)
FONT_SMALL = f(FONT_SANS_LIGHT, 16)
FONT_LABEL = f(FONT_SANS, 18)
FONT_MONO_16 = f(FONT_MONO, 16)
FONT_MONO_18 = f(FONT_MONO, 18)


RAW_HIGHLIGHTS = {
    (0, 0): {"text": "部門", "fill": "#edf6eb", "color": "#365c3d"},
    (1, 0): {"text": "月份", "fill": "#edf6eb", "color": "#365c3d"},
    (2, 0): {"text": "業績", "fill": "#edf6eb", "color": "#365c3d"},
    (0, 1): {"text": "電子", "fill": "#f3f8ff", "color": "#2f4c78"},
    (1, 1): {"text": "1月", "fill": "#f3f8ff", "color": TEXT},
    (2, 1): {"text": "120,000", "fill": "#edf9ef", "color": "#2f6b46"},
    (0, 2): {"text": "電子", "fill": "#f3f8ff", "color": "#2f4c78"},
    (1, 2): {"text": "2月", "fill": "#f3f8ff", "color": TEXT},
    (2, 2): {"text": "95,000", "fill": "#f3f8ff", "color": TEXT},
    (0, 3): {"text": "家電", "fill": "#f3f8ff", "color": "#2f4c78"},
    (1, 3): {"text": "1月", "fill": "#f3f8ff", "color": TEXT},
    (2, 3): {"text": "80,000", "fill": "#f3f8ff", "color": TEXT},
    (0, 4): {"text": "家電", "fill": "#f3f8ff", "color": "#2f4c78"},
    (1, 4): {"text": "2月", "fill": "#f3f8ff", "color": TEXT},
    (2, 4): {"text": "110,000", "fill": "#edf9ef", "color": "#2f6b46"},
    (0, 5): {"text": "電子", "fill": "#f3f8ff", "color": "#2f4c78"},
    (1, 5): {"text": "1月", "fill": "#f3f8ff", "color": TEXT},
    (2, 5): {"text": "60,000", "fill": "#edf9ef", "color": "#2f6b46"},
}

PIVOT_EMPTY = {
    (0, 2): {"text": "樞紐分析表", "fill": "#fff7e6", "color": "#915e12"},
    (0, 3): {"text": "在這裡建立報表", "fill": "#eef7ff", "color": "#375787"},
}

ROWS_ADD = {
    (0, 2): {"text": "部門", "fill": "#edf6eb", "color": "#365c3d"},
    (0, 3): {"text": "電子", "fill": "#f3f8ff", "color": "#2f4c78"},
    (0, 4): {"text": "家電", "fill": "#f3f8ff", "color": "#2f4c78"},
}

COLS_ADD = {
    (1, 2): {"text": "1月", "fill": "#fff5e3", "color": "#a16a11"},
    (2, 2): {"text": "2月", "fill": "#fff5e3", "color": "#a16a11"},
}

VALUES_ADD = {
    (1, 3): {"text": "180,000", "fill": "#edf9ef", "color": "#2f6b46"},
    (2, 3): {"text": "95,000", "fill": "#eef7ff", "color": TEXT},
    (1, 4): {"text": "80,000", "fill": "#eef7ff", "color": TEXT},
    (2, 4): {"text": "110,000", "fill": "#edf9ef", "color": "#2f6b46"},
}


def lerp(a, b, t):
    return a + (b - a) * t


def clamp(v, low=0.0, high=1.0):
    return max(low, min(high, v))


def ease_out_cubic(t):
    return 1 - (1 - clamp(t)) ** 3


def ease_in_out_cubic(t):
    t = clamp(t)
    if t < 0.5:
        return 4 * t * t * t
    return 1 - ((-2 * t + 2) ** 3) / 2


def bezier(p0, p1, p2, t):
    mt = 1 - t
    return (
        mt * mt * p0[0] + 2 * mt * t * p1[0] + t * t * p2[0],
        mt * mt * p0[1] + 2 * mt * t * p1[1] + t * t * p2[1],
    )


def hex_to_rgb(color):
    value = color.lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def rgba(color, alpha):
    if isinstance(color, tuple):
        return color[:3] + (alpha,)
    return hex_to_rgb(color) + (alpha,)


def clone_highlights(base):
    return {coord: dict(cell) for coord, cell in base.items()}


def expand_box(box, pad):
    return (box[0] - pad, box[1] - pad, box[2] + pad, box[3] + pad)


def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def shadowed_card(base, box, radius=22, fill=SURFACE, outline=LINE):
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    rounded(sdraw, (box[0], box[1] + 8, box[2], box[3] + 8), radius, SHADOW)
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))
    base.alpha_composite(shadow)
    draw = ImageDraw.Draw(base)
    rounded(draw, box, radius, fill, outline=outline, width=1)


def soft_glow(base, box, color, alpha=90, blur=20, radius=20, spread=16):
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    rounded(draw, expand_box(box, spread), radius, rgba(color, alpha))
    layer = layer.filter(ImageFilter.GaussianBlur(blur))
    base.alpha_composite(layer)


def spotlight(base, focus_boxes, dim_alpha=92, blur=34):
    if not focus_boxes:
        return
    mask = Image.new("L", base.size, 255)
    draw = ImageDraw.Draw(mask)
    for box in focus_boxes:
        rounded(draw, expand_box(box, 18), 28, 0)
    mask = mask.filter(ImageFilter.GaussianBlur(blur))
    mask = mask.point(lambda px: int(px * dim_alpha / 255))
    overlay = Image.new("RGBA", base.size, DIM + (0,))
    overlay.putalpha(mask)
    base.alpha_composite(overlay)


def text(draw, xy, s, font, fill=TEXT, anchor="la"):
    draw.text(xy, s, font=font, fill=fill, anchor=anchor)


def field_style(label):
    if label == "部門":
        return "#eef6ff", BLUE
    if label == "月份":
        return "#fff5e3", AMBER
    if label == "業績":
        return "#ecf7ef", GREEN
    return "#edf3ea", ACCENT


def zone_color(name):
    return {"篩選": PURPLE, "欄": AMBER, "列": BLUE, "值": GREEN}.get(name, ACCENT)


def chip(draw, box, label, fill="#edf3ea", stroke="#d4dfd0", color=TEXT, alpha=255):
    rounded(draw, box, 16, rgba(fill, alpha), rgba(stroke, alpha))
    cx = (box[0] + box[2]) / 2
    cy = (box[1] + box[3]) / 2 + 1
    text(draw, (cx, cy), label, FONT_SMALL, fill=rgba(color, alpha), anchor="mm")


def floating_chip(base, box, label, alpha=255):
    fill, color = field_style(label)
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    rounded(sdraw, (box[0], box[1] + 10, box[2], box[3] + 10), 18, (26, 32, 28, 58))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    base.alpha_composite(shadow)
    draw = ImageDraw.Draw(base)
    chip(draw, box, label, fill=fill, stroke="#cfd9cf", color=color, alpha=alpha)


def pointer(draw, x, y, scale=1.0, fill="#1f2937"):
    pts = [(0, 0), (36, 22), (23, 24), (34, 50), (25, 54), (16, 30), (0, 48)]
    pts = [(x + px * scale, y + py * scale) for px, py in pts]
    draw.polygon(pts, fill=fill)
    hi = [
        (x + 4 * scale, y + 4 * scale),
        (x + 25 * scale, y + 18 * scale),
        (x + 17 * scale, y + 20 * scale),
    ]
    draw.polygon(hi, fill=(255, 255, 255, 90))


def progress_track(draw, stage, pulse=0.0):
    steps = ["資料", "插入", "列", "欄", "值"]
    start_x, y = 836, 88
    spacing = 82
    for idx, label in enumerate(steps):
        cx = start_x + idx * spacing
        is_done = idx < stage - 1
        is_active = idx == stage - 1
        if idx < len(steps) - 1:
            line_color = "#d7dfd1" if not is_done else "#b3caa9"
            draw.line((cx + 16, y, cx + spacing - 16, y), fill=line_color, width=4)
        radius = 13 if is_active else 11
        if is_active:
            radius += int(2 * pulse)
        fill = "#ffffff"
        outline = "#d3dbcf"
        if is_done:
            fill = "#ecf4e8"
            outline = "#b3caa9"
        if is_active:
            fill = "#f2f8ef"
            outline = ACCENT
        draw.ellipse((cx - radius, y - radius, cx + radius, y + radius), fill=fill, outline=outline, width=3)
        text(draw, (cx, y), str(idx + 1), FONT_SMALL, fill=ACCENT if is_active else MUTED, anchor="mm")
        text(draw, (cx, y + 32), label, FONT_SMALL, fill=TEXT if is_active else MUTED, anchor="mm")


def ribbon(draw, box, active="插入"):
    rounded(draw, box, 18, "#eef4ea", outline="#dde6d9")
    x = box[0] + 18
    for tab in ["常用", "插入", "公式", "資料", "檢視"]:
        width = 54 if len(tab) == 2 else 62
        fill = "#ffffff" if tab == active else "#eef4ea"
        outline = "#d7e0d4" if tab == active else None
        rounded(draw, (x, box[1] + 10, x + width, box[1] + 40), 14, fill, outline=outline)
        text(draw, (x + width / 2, box[1] + 26), tab, FONT_SMALL, fill=ACCENT if tab == active else MUTED, anchor="mm")
        x += width + 10


def workbook_shell(draw, box, sheet_name, formula, active_ref, ribbon_tab="插入", formula_progress=1.0):
    rounded(draw, box, 20, "#fbfbf8", outline="#dad4c8")
    tx0, ty0, tx1, ty1 = box
    ribbon(draw, (tx0 + 18, ty0 + 18, tx1 - 18, ty0 + 62), active=ribbon_tab)
    rounded(draw, (tx0 + 18, ty0 + 78, tx0 + 98, ty0 + 116), 12, "#f2f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 58, ty0 + 97), active_ref, FONT_MONO_16, fill="#57716b", anchor="mm")
    rounded(draw, (tx0 + 108, ty0 + 78, tx0 + 146, ty0 + 116), 12, "#f2f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 127, ty0 + 97), "fx", FONT_MONO_16, fill=MUTED, anchor="mm")
    rounded(draw, (tx0 + 156, ty0 + 78, tx1 - 18, ty0 + 116), 12, "#ffffff", outline="#d7dbd0")
    length = max(1, min(len(formula), int(math.ceil(len(formula) * clamp(formula_progress)))))
    typed = formula[:length]
    if formula_progress < 0.999:
        typed += "▏"
    text(draw, (tx0 + 176, ty0 + 97), typed, FONT_MONO_16, fill=TEXT, anchor="lm")
    rounded(draw, (tx0 + 18, ty1 - 52, tx0 + 116, ty1 - 16), 12, "#ecf3e8", outline="#d7dbd0")
    text(draw, (tx0 + 67, ty1 - 34), sheet_name, FONT_SMALL, fill=ACCENT, anchor="mm")


def grid_cell_box(box, cols, rows, col, row):
    x0, y0, x1, y1 = box
    row_h = (y1 - y0) / (rows + 1)
    col_w = (x1 - x0) / (cols + 1)
    return (
        x0 + (col + 1) * col_w,
        y0 + (row + 1) * row_h,
        x0 + (col + 2) * col_w,
        y0 + (row + 2) * row_h,
    )


def group_box(box, cols, rows, cells, pad=10):
    xs, ys = [], []
    for col, row in cells:
        cb = grid_cell_box(box, cols, rows, col, row)
        xs.extend([cb[0], cb[2]])
        ys.extend([cb[1], cb[3]])
    return (min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)


def draw_grid(base, box, cols, rows, active=None, highlights=None, active_pulse=0.0):
    highlights = highlights or {}
    draw = ImageDraw.Draw(base)
    x0, y0, x1, y1 = box
    row_h = (y1 - y0) / (rows + 1)
    col_w = (x1 - x0) / (cols + 1)
    rounded(draw, box, 16, "#ffffff", outline="#d9d4c9")
    for c in range(cols + 1):
        xx = x0 + c * col_w
        draw.line((xx, y0, xx, y1), fill="#ded8cb", width=1)
    for r in range(rows + 1):
        yy = y0 + r * row_h
        draw.line((x0, yy, x1, yy), fill="#ded8cb", width=1)
    for c in range(cols + 1):
        rounded(draw, (x0 + c * col_w, y0, x0 + (c + 1) * col_w, y0 + row_h), 0, "#eef3ea")
    for r in range(rows + 1):
        rounded(draw, (x0, y0 + r * row_h, x0 + col_w, y0 + (r + 1) * row_h), 0, "#eef3ea")
    for c in range(cols):
        text(draw, (x0 + (c + 1.5) * col_w, y0 + row_h / 2), chr(65 + c), FONT_MONO_16, fill="#6f7a70", anchor="mm")
    for r in range(rows):
        text(draw, (x0 + col_w / 2, y0 + (r + 1.5) * row_h), str(r + 1), FONT_MONO_16, fill="#6f7a70", anchor="mm")
    for (col, row), cell in highlights.items():
        cb = grid_cell_box(box, cols, rows, col, row)
        alpha = clamp(cell.get("alpha", 1.0))
        if alpha > 0:
            fill = cell.get("fill", "#ffffff")
            layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
            ldraw = ImageDraw.Draw(layer)
            rounded(ldraw, cb, 0, rgba(fill, int(255 * alpha)))
            base.alpha_composite(layer)
            draw = ImageDraw.Draw(base)
            text(
                draw,
                ((cb[0] + cb[2]) / 2, (cb[1] + cb[3]) / 2),
                cell.get("text", ""),
                FONT_MONO_18,
                fill=rgba(cell.get("color", TEXT), int(255 * alpha)),
                anchor="mm",
            )
    if active:
        cb = grid_cell_box(box, cols, rows, active[0], active[1])
        soft_glow(base, cb, ACCENT, alpha=int(44 + 36 * active_pulse), blur=16, spread=8)
        draw = ImageDraw.Draw(base)
        width = 4 + int(2 * active_pulse)
        draw.rounded_rectangle((cb[0] + 2, cb[1] + 2, cb[2] - 2, cb[3] - 2), radius=2, outline="#2f855a", width=width)
        handle = 7 + int(1 * active_pulse)
        draw.rectangle((cb[2] - handle - 2, cb[3] - handle - 2, cb[2] - 2, cb[3] - 2), fill="#2f855a")


def sidebar_layout(box):
    field_y = box[1] + 54
    field_boxes = {}
    for label in ["部門", "月份", "業績"]:
        field_boxes[label] = (box[0] + 18, field_y, box[0] + 138, field_y + 42)
        field_y += 52
    zones = {
        "篩選": (box[0] + 18, box[1] + 252, box[0] + 168, box[1] + 294),
        "欄": (box[0] + 18, box[1] + 306, box[0] + 168, box[1] + 348),
        "列": (box[0] + 18, box[1] + 360, box[0] + 168, box[1] + 402),
        "值": (box[0] + 18, box[1] + 414, box[0] + 168, box[1] + 456),
    }
    return {"fields": field_boxes, "zones": zones}


def draw_field_sidebar(base, box, placed=None, placed_alpha=None, drop_target=None):
    placed = placed or {}
    placed_alpha = placed_alpha or {}
    draw = ImageDraw.Draw(base)
    rounded(draw, box, 18, "#fbfaf5", outline="#dad4c8")
    text(draw, (box[0] + 18, box[1] + 20), "樞紐分析欄位", FONT_LABEL, fill=TEXT)
    layout = sidebar_layout(box)
    for label, fb in layout["fields"].items():
        fill, color = field_style(label)
        chip(draw, fb, label, fill=fill, stroke="#d7dfd3", color=color)
    text(draw, (box[0] + 18, box[1] + 228), "拖到這裡", FONT_SMALL, fill=MUTED)
    for name, zone_box in layout["zones"].items():
        accent = zone_color(name)
        fill = "#f8f6ef"
        outline = "#ddd6c9"
        width = 1
        if drop_target == name:
            fill = "#fffaf2" if name == "欄" else "#eef7ff" if name == "列" else "#edf7ef" if name == "值" else "#f8f4ff"
            outline = accent
            width = 3
        rounded(draw, zone_box, 14, fill, outline=outline, width=width)
        text(draw, (zone_box[0] + 16, (zone_box[1] + zone_box[3]) / 2), name, FONT_SMALL, fill=MUTED, anchor="lm")
        if name in placed:
            fill_color, text_color = field_style(placed[name])
            alpha = int(255 * clamp(placed_alpha.get(name, 1.0)))
            chip(
                draw,
                (zone_box[0] + 52, zone_box[1] + 6, zone_box[0] + 140, zone_box[3] - 6),
                placed[name],
                fill=fill_color,
                stroke="#d8dfd2",
                color=text_color,
                alpha=alpha,
            )
    return layout


def draw_note_card(draw, box, title, lines, tint="#eef7ef", title_color=ACCENT):
    rounded(draw, box, 18, tint, outline="#d2e3d6")
    text(draw, (box[0] + 20, box[1] + 26), title, FONT_LABEL, fill=title_color)
    for idx, line in enumerate(lines):
        text(draw, (box[0] + 20, box[1] + 60 + idx * 30), "• " + line, FONT_BODY, fill="#4e5b50")


def merge_with_reveal(base, additions, progress, order, step=0.18, window=0.34):
    merged = clone_highlights(base)
    for idx, coord in enumerate(order):
        local = clamp((progress - idx * step) / window)
        if local <= 0:
            continue
        cell = dict(additions[coord])
        cell["alpha"] = ease_out_cubic(local)
        merged[coord] = cell
    return merged


def rows_transition(progress):
    return merge_with_reveal(PIVOT_EMPTY, ROWS_ADD, progress, [(0, 2), (0, 3), (0, 4)])


def cols_transition(progress):
    base = clone_highlights(ROWS_ADD)
    return merge_with_reveal(base, COLS_ADD, progress, [(1, 2), (2, 2)], step=0.22, window=0.42)


def values_transition(progress):
    base = clone_highlights(ROWS_ADD)
    base.update(clone_highlights(COLS_ADD))
    return merge_with_reveal(base, VALUES_ADD, progress, [(1, 3), (2, 3), (1, 4), (2, 4)], step=0.12, window=0.3)


def render_raw_scene(pulse=0.0, pointer_shift=0.0):
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    text(draw, (88, 68), "樞紐分析操作短片", FONT_TITLE)
    text(draw, (88, 108), "從原始資料 → 插入樞紐 → 拖欄位 → 產生報表", FONT_BODY, fill=MUTED)
    progress_track(draw, 1, pulse)
    shadowed_card(img, (72, 144, 892, 646), 24)
    draw = ImageDraw.Draw(img)
    workbook_shell(draw, SHEET_BOX, "Sales", "站在資料區內任一格，準備插入樞紐分析表", "B3")
    draw_grid(img, RAW_GRID_BOX, 3, 6, active=(1, 2), highlights=RAW_HIGHLIGHTS, active_pulse=pulse)
    check_box = (926, 190, 1198, 362)
    draw_note_card(draw, check_box, "這一步先檢查", ["第一列是欄位名", "中間不要有空白欄", "資料最好先轉成 Table"], tint="#fffdf8", title_color=TEXT)
    spotlight(img, [expand_box(RAW_GRID_BOX, 20), expand_box(check_box, 14)], dim_alpha=74)
    draw = ImageDraw.Draw(img)
    pointer(draw, 362 + pointer_shift, 398 - pointer_shift * 0.4, 1.08)
    return img


def render_pivot_scene(
    stage,
    title,
    subtitle,
    formula,
    active_ref,
    highlights,
    active_cell,
    placed=None,
    placed_alpha=None,
    drop_target=None,
    drop_glow=0.0,
    source_focus=None,
    source_glow=0.0,
    floating=None,
    pointer_pos=None,
    focus_cells=None,
    formula_progress=1.0,
    note=None,
    active_pulse=0.0,
):
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    text(draw, (88, 68), title, FONT_TITLE)
    text(draw, (88, 108), subtitle, FONT_BODY, fill=MUTED)
    progress_track(draw, stage, active_pulse)
    shadowed_card(img, (72, 144, 1208, 646), 24)
    draw = ImageDraw.Draw(img)
    workbook_shell(draw, SHEET_BOX, "Pivot", formula, active_ref, formula_progress=formula_progress)

    sidebar = sidebar_layout(SIDEBAR_BOX)
    focus_boxes = []
    if source_focus:
        field_box = sidebar["fields"][source_focus]
        soft_glow(img, field_box, zone_color(drop_target or "列"), alpha=int(58 + 68 * source_glow), blur=18, spread=10)
        focus_boxes.append(field_box)
    if drop_target:
        zone_box = sidebar["zones"][drop_target]
        soft_glow(img, zone_box, zone_color(drop_target), alpha=int(52 + 90 * drop_glow), blur=20, spread=16)
        focus_boxes.append(zone_box)
    if focus_cells:
        grid_focus = group_box(PIVOT_GRID_BOX, 4, 6, focus_cells, pad=12)
        soft_glow(img, grid_focus, ACCENT, alpha=62, blur=18, spread=10)
        focus_boxes.append(grid_focus)
    if floating:
        focus_boxes.append(expand_box(floating["box"], 14))

    draw_grid(img, PIVOT_GRID_BOX, 4, 6, active=active_cell, highlights=highlights, active_pulse=active_pulse)
    draw_field_sidebar(img, SIDEBAR_BOX, placed=placed, placed_alpha=placed_alpha, drop_target=drop_target)

    note_box = None
    if note:
        note_box = (900, 520, 1184, 612)
        draw = ImageDraw.Draw(img)
        draw_note_card(draw, note_box, note["title"], note["lines"], tint=note.get("tint", "#eef7ef"), title_color=note.get("color", "#2f6b46"))
        focus_boxes.append(note_box)

    spotlight(img, focus_boxes, dim_alpha=80)

    if floating:
        floating_chip(img, floating["box"], floating["label"], alpha=int(255 * clamp(floating.get("alpha", 1.0))))
    if pointer_pos:
        draw = ImageDraw.Draw(img)
        pointer(draw, pointer_pos[0], pointer_pos[1], 1.02)
    return img


def drag_path(field, target):
    layout = sidebar_layout(SIDEBAR_BOX)
    src = layout["fields"][field]
    dst = layout["zones"][target]
    start = ((src[0] + src[2]) / 2, (src[1] + src[3]) / 2)
    end = (dst[0] + 94, (dst[1] + dst[3]) / 2)
    ctrl = (max(start[0], end[0]) + 54, min(start[1], end[1]) - 40)
    return start, ctrl, end


def drag_frame(
    stage,
    field,
    target,
    title,
    subtitle,
    formula,
    active_ref,
    placed_before,
    placed_after,
    highlight_builder,
    reveal,
    move_progress,
    settle=False,
    settle_pointer=None,
    focus_cells=None,
):
    placed = dict(placed_before)
    placed_alpha = {zone: 1.0 for zone in placed_before}
    if reveal > 0:
        placed[target] = field
        placed_alpha[target] = reveal

    floating = None
    pointer_pos = settle_pointer
    source_glow = 0.0 if settle else (1.0 - move_progress) * 0.9 + 0.1
    drop_glow = 0.25 + 0.65 * reveal
    if not settle:
        start, ctrl, end = drag_path(field, target)
        cx, cy = bezier(start, ctrl, end, ease_in_out_cubic(move_progress))
        box = (cx - 54, cy - 20, cx + 54, cy + 20)
        floating = {"label": field, "box": box, "alpha": 1.0}
        pointer_pos = (box[0] + 72, box[1] - 8)
    return render_pivot_scene(
        stage=stage,
        title=title,
        subtitle=subtitle,
        formula=formula,
        active_ref=active_ref,
        highlights=highlight_builder(reveal),
        active_cell=(0, 3) if target == "列" and reveal > 0.5 else (1, 2) if target == "欄" and reveal > 0.5 else (1, 3) if target == "值" and reveal > 0.6 else (0, 3),
        placed=placed,
        placed_alpha=placed_alpha,
        drop_target=target,
        drop_glow=drop_glow,
        source_focus=None if settle else field,
        source_glow=source_glow,
        floating=floating,
        pointer_pos=pointer_pos,
        focus_cells=focus_cells,
        formula_progress=0.38 + 0.62 * ease_out_cubic(move_progress) if not settle else 1.0,
        active_pulse=0.38 + 0.62 * reveal,
    )


def append_intro(frames, durations):
    for idx in range(6):
        t = idx / 5
        pulse = 0.35 + 0.65 * math.sin(t * math.pi)
        frames.append(render_raw_scene(pulse=pulse, pointer_shift=8 * math.sin(t * math.pi)))
        durations.append(90)
    frames.append(render_raw_scene(pulse=0.9, pointer_shift=0))
    durations.append(180)


def append_insert(frames, durations):
    start = render_raw_scene(pulse=0.9, pointer_shift=0)
    for idx in range(6):
        t = idx / 5
        pivot = render_pivot_scene(
            stage=2,
            title="插入後，Excel 會開一張空的樞紐畫布",
            subtitle="先看到空白很正常，真正內容要等你拖欄位",
            formula="PivotTable from Table1",
            active_ref="A4",
            highlights=PIVOT_EMPTY,
            active_cell=(0, 3),
            focus_cells=[(0, 2), (0, 3)],
            formula_progress=0.25 + 0.75 * t,
            active_pulse=0.4 + 0.4 * t,
            placed={},
            pointer_pos=(330 + 10 * t, 440 - 4 * t),
        )
        frames.append(Image.blend(start, pivot, ease_in_out_cubic(t)))
        durations.append(86)
    for idx in range(3):
        pulse = 0.45 + idx * 0.15
        frames.append(
            render_pivot_scene(
                stage=2,
                title="插入後，Excel 會開一張空的樞紐畫布",
                subtitle="先看到空白很正常，真正內容要等你拖欄位",
                formula="PivotTable from Table1",
                active_ref="A4",
                highlights=PIVOT_EMPTY,
                active_cell=(0, 3),
                focus_cells=[(0, 2), (0, 3)],
                formula_progress=1.0,
                active_pulse=pulse,
                placed={},
                pointer_pos=(338, 436),
            )
        )
        durations.append(110 if idx < 2 else 170)


def append_drag(frames, durations, field, target, stage, title, subtitle, formula, active_ref, placed_before, placed_after, highlight_builder, focus_cells, settle_pointer):
    for idx in range(2):
        pulse = 0.48 + 0.32 * idx
        frames.append(
            render_pivot_scene(
                stage=stage,
                title=title,
                subtitle=subtitle,
                formula=formula,
                active_ref=active_ref,
                highlights=highlight_builder(0),
                active_cell=(0, 3),
                placed=placed_before,
                drop_target=target,
                drop_glow=0.18,
                source_focus=field,
                source_glow=0.7 + 0.15 * idx,
                focus_cells=None,
                formula_progress=0.36,
                active_pulse=pulse,
                pointer_pos=(972, sidebar_layout(SIDEBAR_BOX)["fields"][field][1] - 14),
            )
        )
        durations.append(90)
    for idx in range(10):
        t = idx / 9
        reveal = ease_out_cubic(clamp((t - 0.68) / 0.32))
        frames.append(
            drag_frame(
                stage=stage,
                field=field,
                target=target,
                title=title,
                subtitle=subtitle,
                formula=formula,
                active_ref=active_ref,
                placed_before=placed_before,
                placed_after=placed_after,
                highlight_builder=highlight_builder,
                reveal=reveal,
                move_progress=t,
                focus_cells=focus_cells if reveal > 0.2 else None,
            )
        )
        durations.append(78)
    for idx in range(3):
        pulse = 0.5 + 0.4 * math.sin((idx + 1) / 3 * math.pi)
        frames.append(
            drag_frame(
                stage=stage,
                field=field,
                target=target,
                title=title,
                subtitle=subtitle,
                formula=formula,
                active_ref=active_ref,
                placed_before=placed_before,
                placed_after=placed_after,
                highlight_builder=highlight_builder,
                reveal=1.0,
                move_progress=1.0,
                settle=True,
                settle_pointer=settle_pointer,
                focus_cells=focus_cells,
            )
        )
        durations.append(108 if idx < 2 else 165)


def append_final(frames, durations):
    note = {
        "title": "你現在得到：",
        "lines": ["電子 1 月自動聚合為 180,000", "之後只要換欄位位置，就能切別的角度"],
    }
    for idx in range(4):
        pulse = 0.45 + 0.45 * math.sin((idx + 1) / 5 * math.pi)
        frames.append(
            render_pivot_scene(
                stage=5,
                title="完成：原始 5 筆交易，變成可讀的樞紐報表",
                subtitle="之後只要換欄位位置或加篩選器，就能快速切出其他角度",
                formula="Rows = 部門 / Columns = 月份 / Values = SUM(業績)",
                active_ref="B4",
                highlights=values_transition(1),
                active_cell=(1, 3),
                placed={"列": "部門", "欄": "月份", "值": "業績"},
                drop_target="值",
                drop_glow=0.45 + 0.22 * pulse,
                focus_cells=[(1, 3), (2, 3), (1, 4), (2, 4)],
                formula_progress=1.0,
                note=note,
                pointer_pos=(448, 438),
                active_pulse=pulse,
            )
        )
        durations.append(128 if idx < 3 else 280)


def build_animation():
    frames, durations = [], []
    append_intro(frames, durations)
    append_insert(frames, durations)
    append_drag(
        frames,
        durations,
        field="部門",
        target="列",
        stage=3,
        title="把『部門』拖進列，先決定左邊怎麼分組",
        subtitle="先做出報表骨架，讓每個部門有自己的列標籤",
        formula="Rows = 部門",
        active_ref="A4",
        placed_before={},
        placed_after={"列": "部門"},
        highlight_builder=rows_transition,
        focus_cells=[(0, 2), (0, 3), (0, 4)],
        settle_pointer=(264, 438),
    )
    append_drag(
        frames,
        durations,
        field="月份",
        target="欄",
        stage=4,
        title="再把『月份』拖進欄，讓上方按月份展開",
        subtitle="同一張報表裡，時間軸現在會沿著上方橫向打開",
        formula="Rows = 部門 / Columns = 月份",
        active_ref="B3",
        placed_before={"列": "部門"},
        placed_after={"列": "部門", "欄": "月份"},
        highlight_builder=cols_transition,
        focus_cells=[(0, 2), (1, 2), (2, 2)],
        settle_pointer=(410, 366),
    )
    append_drag(
        frames,
        durations,
        field="業績",
        target="值",
        stage=5,
        title="最後把『業績』拖進值，Excel 就會自動加總",
        subtitle="這一步完成後，原始交易會立刻被彙總成可讀的數字報表",
        formula="Rows = 部門 / Columns = 月份 / Values = SUM(業績)",
        active_ref="B4",
        placed_before={"列": "部門", "欄": "月份"},
        placed_after={"列": "部門", "欄": "月份", "值": "業績"},
        highlight_builder=values_transition,
        focus_cells=[(1, 3), (2, 3), (1, 4), (2, 4)],
        settle_pointer=(448, 438),
    )
    append_final(frames, durations)
    return frames, durations


def main():
    frames, durations = build_animation()
    durations = [int(round(d * DURATION_SCALE)) for d in durations]

    gif_path = MEDIA_DIR / "p2-03-pivot-demo.gif"
    webp_path = MEDIA_DIR / "p2-03-pivot-demo.webp"
    poster_path = MEDIA_DIR / "p2-03-pivot-demo-poster.png"

    frames[0].save(
        gif_path,
        save_all=True,
        append_images=frames[1:],
        duration=durations,
        loop=0,
        disposal=2,
        optimize=True,
    )
    frames[0].save(
        webp_path,
        format="WEBP",
        save_all=True,
        append_images=frames[1:],
        duration=durations,
        loop=0,
        quality=88,
        method=6,
    )
    frames[-1].save(poster_path)
    print(gif_path)
    print(webp_path)
    print(poster_path)


if __name__ == "__main__":
    main()
