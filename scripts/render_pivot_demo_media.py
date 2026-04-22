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
ACCENT_SOFT = "#e7f1e4"
BLUE = "#6aa9ff"
AMBER = "#efb74f"
GREEN = "#51b36f"
PURPLE = "#a482ff"
SHADOW = (44, 52, 46, 26)


FONT_SANS = "/System/Library/Fonts/STHeiti Medium.ttc"
FONT_SANS_LIGHT = "/System/Library/Fonts/STHeiti Light.ttc"
FONT_MONO = "/System/Library/Fonts/SFNSMono.ttf"


def f(path, size):
    return ImageFont.truetype(path, size)


FONT_TITLE = f(FONT_SANS, 34)
FONT_BODY = f(FONT_SANS_LIGHT, 20)
FONT_SMALL = f(FONT_SANS_LIGHT, 16)
FONT_LABEL = f(FONT_SANS, 18)
FONT_MONO_16 = f(FONT_MONO, 16)
FONT_MONO_18 = f(FONT_MONO, 18)
FONT_MONO_22 = f(FONT_MONO, 22)


def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def shadowed_card(base, box, radius=22, fill=SURFACE, outline=LINE):
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    offset_box = (box[0], box[1] + 8, box[2], box[3] + 8)
    rounded(sdraw, offset_box, radius, SHADOW)
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))
    base.alpha_composite(shadow)
    draw = ImageDraw.Draw(base)
    rounded(draw, box, radius, fill, outline=outline, width=1)


def text(draw, xy, s, font, fill=TEXT, anchor="la"):
    draw.text(xy, s, font=font, fill=fill, anchor=anchor)


def chip(draw, box, label, fill="#edf3ea", stroke="#d4dfd0", color=TEXT):
    rounded(draw, box, 16, fill, stroke)
    cx = (box[0] + box[2]) / 2
    cy = (box[1] + box[3]) / 2 + 1
    text(draw, (cx, cy), label, FONT_SMALL, fill=color, anchor="mm")


def pointer(draw, x, y, scale=1.0, fill="#1f2937"):
    pts = [
        (0, 0), (36, 22), (23, 24), (34, 50), (25, 54), (16, 30), (0, 48)
    ]
    pts = [(x + px * scale, y + py * scale) for px, py in pts]
    draw.polygon(pts, fill=fill)
    hi = [(x + 4 * scale, y + 4 * scale), (x + 25 * scale, y + 18 * scale), (x + 17 * scale, y + 20 * scale)]
    draw.polygon(hi, fill=(255, 255, 255, 90))


def arrow(draw, start, end, color="#7c8d79", width=4):
    draw.line([start, end], fill=color, width=width)
    ang = math.atan2(end[1] - start[1], end[0] - start[0])
    for da in (math.pi * 0.82, -math.pi * 0.82):
        p = (
            end[0] + math.cos(ang + da) * 16,
            end[1] + math.sin(ang + da) * 16,
        )
        draw.line([end, p], fill=color, width=width)


def ribbon(draw, box, active="插入"):
    rounded(draw, box, 18, "#eef4ea", outline="#dde6d9")
    x = box[0] + 18
    tabs = ["常用", "插入", "公式", "資料", "檢視"]
    for t in tabs:
        w = 54 if len(t) == 2 else 62
        fill = "#ffffff" if t == active else "#eef4ea"
        outline = "#d7e0d4" if t == active else None
        rounded(draw, (x, box[1] + 10, x + w, box[1] + 40), 14, fill, outline=outline)
        text(draw, (x + w / 2, box[1] + 26), t, FONT_SMALL, fill=ACCENT if t == active else MUTED, anchor="mm")
        x += w + 10


def workbook_shell(draw, box, sheet_name, formula, active_ref, active_tab=True):
    rounded(draw, box, 20, "#fbfbf8", outline="#dad4c8")
    tx0, ty0, tx1, ty1 = box
    ribbon(draw, (tx0 + 18, ty0 + 18, tx1 - 18, ty0 + 62))
    rounded(draw, (tx0 + 18, ty0 + 78, tx0 + 98, ty0 + 116), 12, "#f2f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 58, ty0 + 97), active_ref, FONT_MONO_16, fill="#57716b", anchor="mm")
    rounded(draw, (tx0 + 108, ty0 + 78, tx0 + 146, ty0 + 116), 12, "#f2f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 127, ty0 + 97), "fx", FONT_MONO_16, fill=MUTED, anchor="mm")
    rounded(draw, (tx0 + 156, ty0 + 78, tx1 - 18, ty0 + 116), 12, "#ffffff", outline="#d7dbd0")
    text(draw, (tx0 + 176, ty0 + 97), formula, FONT_MONO_16, fill=TEXT, anchor="lm")
    rounded(draw, (tx0 + 18, ty1 - 52, tx0 + 116, ty1 - 16), 12, "#ecf3e8" if active_tab else "#f5f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 67, ty1 - 34), sheet_name, FONT_SMALL, fill=ACCENT if active_tab else MUTED, anchor="mm")


def draw_grid(draw, box, cols, rows, active=None, highlights=None):
    highlights = highlights or {}
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
        fill = "#eef3ea"
        rounded(draw, (x0 + c * col_w, y0, x0 + (c + 1) * col_w, y0 + row_h), 0, fill)
    for r in range(rows + 1):
        rounded(draw, (x0, y0 + r * row_h, x0 + col_w, y0 + (r + 1) * row_h), 0, "#eef3ea")
    for c in range(cols):
        label = chr(65 + c)
        text(draw, (x0 + (c + 1.5) * col_w, y0 + row_h / 2), label, FONT_MONO_16, fill="#6f7a70", anchor="mm")
    for r in range(rows):
        text(draw, (x0 + col_w / 2, y0 + (r + 1.5) * row_h), str(r + 1), FONT_MONO_16, fill="#6f7a70", anchor="mm")
    for (col, row), cell in highlights.items():
        cx0 = x0 + (col + 1) * col_w
        cy0 = y0 + (row + 1) * row_h
        cx1 = x0 + (col + 2) * col_w
        cy1 = y0 + (row + 2) * row_h
        fill = cell.get("fill", "#ffffff")
        rounded(draw, (cx0, cy0, cx1, cy1), 0, fill)
        text(draw, ((cx0 + cx1) / 2, (cy0 + cy1) / 2), cell.get("text", ""), FONT_MONO_18, fill=cell.get("color", TEXT), anchor="mm")
    if active:
        col, row = active
        cx0 = x0 + (col + 1) * col_w
        cy0 = y0 + (row + 1) * row_h
        cx1 = x0 + (col + 2) * col_w
        cy1 = y0 + (row + 2) * row_h
        draw.rounded_rectangle((cx0 + 2, cy0 + 2, cx1 - 2, cy1 - 2), radius=2, outline="#2f855a", width=4)
        draw.rectangle((cx1 - 9, cy1 - 9, cx1 - 2, cy1 - 2), fill="#2f855a")


def fields_sidebar(draw, box, drag=None, placed=None):
    placed = placed or {}
    shadow_img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    pass


def draw_field_sidebar(draw, box, dragged=None, dragged_pos=None, drop_target=None, placed=None):
    placed = placed or {}
    rounded(draw, box, 18, "#fbfaf5", outline="#dad4c8")
    text(draw, (box[0] + 18, box[1] + 20), "樞紐分析欄位", FONT_LABEL, fill=TEXT)
    field_y = box[1] + 54
    fields = [("部門", BLUE), ("月份", AMBER), ("業績", GREEN)]
    field_boxes = {}
    for label, color in fields:
      b = (box[0] + 18, field_y, box[0] + 138, field_y + 42)
      field_boxes[label] = b
      fill = {
          BLUE: "#eef6ff",
          AMBER: "#fff5e3",
          GREEN: "#ecf7ef"
      }[color]
      chip(draw, b, label, fill=fill, stroke="#d7dfd3", color=color)
      field_y += 52

    text(draw, (box[0] + 18, box[1] + 228), "拖到這裡", FONT_SMALL, fill=MUTED)
    zones = [
        ("篩選", (box[0] + 18, box[1] + 252, box[0] + 168, box[1] + 294), PURPLE),
        ("欄", (box[0] + 18, box[1] + 306, box[0] + 168, box[1] + 348), AMBER),
        ("列", (box[0] + 18, box[1] + 360, box[0] + 168, box[1] + 402), BLUE),
        ("值", (box[0] + 18, box[1] + 414, box[0] + 168, box[1] + 456), GREEN),
    ]
    zone_map = {}
    for name, zb, color in zones:
        zone_map[name] = zb
        rounded(draw, zb, 14, "#f8f6ef", outline="#ddd6c9")
        if drop_target == name:
            rounded(draw, zb, 14, "#fffaf2" if color == AMBER else "#eef7ff" if color == BLUE else "#edf7ef", outline=color, width=3)
        text(draw, (zb[0] + 16, (zb[1] + zb[3]) / 2), name, FONT_SMALL, fill=MUTED, anchor="lm")
        if name in placed:
            label, fill, tcolor = placed[name]
            chip(draw, (zb[0] + 52, zb[1] + 6, zb[0] + 140, zb[3] - 6), label, fill=fill, stroke="#d8dfd2", color=tcolor)

    if dragged and dragged_pos:
        base_fill = "#eef6ff" if dragged == "部門" else "#fff5e3" if dragged == "月份" else "#ecf7ef"
        base_color = BLUE if dragged == "部門" else AMBER if dragged == "月份" else GREEN
        b = (dragged_pos[0], dragged_pos[1], dragged_pos[0] + 108, dragged_pos[1] + 40)
        chip(draw, b, dragged, fill=base_fill, stroke="#cfdad0", color=base_color)
        src = field_boxes[dragged]
        arrow(draw, (src[2] + 8, (src[1] + src[3]) / 2), (dragged_pos[0], dragged_pos[1] + 18), color="#9ba99a", width=3)
        if drop_target in zone_map:
            zb = zone_map[drop_target]
            arrow(draw, (dragged_pos[0] + 108, dragged_pos[1] + 18), (zb[0] + 20, (zb[1] + zb[3]) / 2), color="#9ba99a", width=3)


def raw_sheet_frame():
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    text(draw, (88, 68), "樞紐分析操作短片", FONT_TITLE)
    text(draw, (88, 108), "從原始資料 → 插入樞紐 → 拖欄位 → 產生報表", FONT_BODY, fill=MUTED)
    shadowed_card(img, (72, 144, 892, 646), 24)
    draw = ImageDraw.Draw(img)
    workbook_shell(draw, (92, 166, 872, 626), "Sales", "站在資料區內任一格，準備插入樞紐分析表", "B3")
    highlights = {
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
    draw_grid(draw, (114, 240, 850, 574), 3, 6, active=(1, 2), highlights=highlights)
    rounded(draw, (926, 190, 1198, 356), 22, "#fffdf8", outline="#d8d0bf")
    text(draw, (948, 218), "這一步先檢查", FONT_LABEL)
    for i, line in enumerate(["第一列是欄位名", "中間不要有空白欄", "資料最好先轉成 Table"]):
        text(draw, (954, 258 + i * 34), "• " + line, FONT_BODY, fill=MUTED)
    pointer(draw, 362, 398, 1.1)
    return img


def pivot_canvas_frame():
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    text(draw, (88, 68), "插入後，Excel 會開一張空的樞紐畫布", FONT_TITLE)
    text(draw, (88, 108), "先看到空白很正常，真正內容要等你拖欄位", FONT_BODY, fill=MUTED)
    shadowed_card(img, (72, 144, 1208, 646), 24)
    draw = ImageDraw.Draw(img)
    workbook_shell(draw, (92, 166, 850, 626), "Pivot", "PivotTable from Table1", "A4")
    highlights = {
        (0, 2): {"text": "樞紐分析表", "fill": "#fff7e6", "color": "#915e12"},
        (0, 3): {"text": "在這裡建立報表", "fill": "#eef7ff", "color": "#375787"},
    }
    draw_grid(draw, (114, 240, 828, 574), 4, 6, active=(0, 3), highlights=highlights)
    draw_field_sidebar(draw, (892, 182, 1184, 616))
    pointer(draw, 330, 440, 1.05)
    return img


def drag_frame(field, drop_target, placed):
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    title = "把欄位拖進四大區域"
    sub = {
        "列": "先拖『部門』到列，決定左邊怎麼分組",
        "欄": "再拖『月份』到欄，決定上方怎麼展開",
        "值": "最後拖『業績』到值，讓 Excel 幫你自動加總",
    }[drop_target]
    text(draw, (88, 68), title, FONT_TITLE)
    text(draw, (88, 108), sub, FONT_BODY, fill=MUTED)
    shadowed_card(img, (72, 144, 1208, 646), 24)
    draw = ImageDraw.Draw(img)
    workbook_shell(draw, (92, 166, 850, 626), "Pivot", f"{field} → {drop_target}", "A3")
    highlights = {}
    if "列" in placed:
        highlights[(0, 2)] = {"text": "部門", "fill": "#edf6eb", "color": "#365c3d"}
        highlights[(0, 3)] = {"text": "電子", "fill": "#f3f8ff", "color": "#2f4c78"}
        highlights[(0, 4)] = {"text": "家電", "fill": "#f3f8ff", "color": "#2f4c78"}
    if "欄" in placed:
        highlights[(1, 2)] = {"text": "1月", "fill": "#fff5e3", "color": "#a16a11"}
        highlights[(2, 2)] = {"text": "2月", "fill": "#fff5e3", "color": "#a16a11"}
    if "值" in placed:
        highlights[(1, 3)] = {"text": "180,000", "fill": "#edf9ef", "color": "#2f6b46"}
        highlights[(2, 3)] = {"text": "95,000", "fill": "#eef7ff", "color": TEXT}
        highlights[(1, 4)] = {"text": "80,000", "fill": "#eef7ff", "color": TEXT}
        highlights[(2, 4)] = {"text": "110,000", "fill": "#edf9ef", "color": "#2f6b46"}
    draw_grid(draw, (114, 240, 828, 574), 4, 6, active=(0 if drop_target == "列" else 1 if drop_target == "欄" else 1, 2 if drop_target != "值" else 3), highlights=highlights)
    drag_positions = {
        "列": (930, 364),
        "欄": (930, 310),
        "值": (930, 418),
    }
    color_map = {
        "部門": ("#eef6ff", BLUE),
        "月份": ("#fff5e3", AMBER),
        "業績": ("#ecf7ef", GREEN),
    }
    placed_visual = {}
    for zone, label in placed.items():
        fill, tcolor = color_map[label]
        placed_visual[zone] = (label, fill, tcolor)
    draw_field_sidebar(draw, (892, 182, 1184, 616), dragged=field, dragged_pos=drag_positions[drop_target], drop_target=drop_target, placed=placed_visual)
    pointer(draw, drag_positions[drop_target][0] + 92, drag_positions[drop_target][1] + 14, 1.0)
    return img


def final_frame():
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    text(draw, (88, 68), "完成：原始 5 筆交易，變成可讀的樞紐報表", FONT_TITLE)
    text(draw, (88, 108), "之後只要換欄位位置或加篩選器，就能快速切出其他角度", FONT_BODY, fill=MUTED)
    shadowed_card(img, (72, 144, 1208, 646), 24)
    draw = ImageDraw.Draw(img)
    workbook_shell(draw, (92, 166, 850, 626), "Pivot", "Rows = 部門 / Columns = 月份 / Values = SUM(業績)", "B4")
    highlights = {
        (0, 2): {"text": "部門", "fill": "#edf6eb", "color": "#365c3d"},
        (1, 2): {"text": "1月", "fill": "#fff5e3", "color": "#a16a11"},
        (2, 2): {"text": "2月", "fill": "#fff5e3", "color": "#a16a11"},
        (0, 3): {"text": "電子", "fill": "#f3f8ff", "color": "#2f4c78"},
        (1, 3): {"text": "180,000", "fill": "#ecf7ef", "color": "#2f6b46"},
        (2, 3): {"text": "95,000", "fill": "#eef7ff", "color": TEXT},
        (0, 4): {"text": "家電", "fill": "#f3f8ff", "color": "#2f4c78"},
        (1, 4): {"text": "80,000", "fill": "#eef7ff", "color": TEXT},
        (2, 4): {"text": "110,000", "fill": "#ecf7ef", "color": "#2f6b46"},
    }
    draw_grid(draw, (114, 240, 828, 574), 4, 6, active=(1, 3), highlights=highlights)
    draw_field_sidebar(
        draw, (892, 182, 1184, 616),
        placed={
            "列": ("部門", "#eef6ff", BLUE),
            "欄": ("月份", "#fff5e3", AMBER),
            "值": ("業績", "#ecf7ef", GREEN),
        },
    )
    rounded(draw, (900, 520, 1184, 610), 18, "#eef7ef", outline="#d2e3d6")
    text(draw, (922, 545), "你現在得到：", FONT_LABEL, fill="#2f6b46")
    text(draw, (922, 580), "• 電子 1 月自動聚合為 180,000", FONT_BODY, fill="#4e5b50")
    pointer(draw, 446, 440, 1.05)
    return img


def main():
    frames = [
        raw_sheet_frame(),
        pivot_canvas_frame(),
        drag_frame("部門", "列", {}),
        drag_frame("月份", "欄", {"列": "部門"}),
        drag_frame("業績", "值", {"列": "部門", "欄": "月份"}),
        final_frame(),
    ]
    durations = [900, 800, 760, 760, 760, 1300]

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
        optimize=False,
    )
    frames[0].save(
        webp_path,
        format="WEBP",
        save_all=True,
        append_images=frames[1:],
        duration=durations,
        loop=0,
        quality=90,
        method=6,
    )
    frames[-1].save(poster_path)
    print(gif_path)
    print(webp_path)
    print(poster_path)


if __name__ == "__main__":
    main()
