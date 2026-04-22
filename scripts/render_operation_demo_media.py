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
RED = "#d26a6a"
PURPLE = "#8d72d6"
SHADOW = (44, 52, 46, 22)

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


def clamp(v, low=0.0, high=1.0):
    return max(low, min(high, v))


def rgba(hex_color, alpha):
    value = hex_color.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4)) + (alpha,)


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


def soft_glow(base, box, color, alpha=90, blur=18, radius=22, spread=18):
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    rounded(
        draw,
        (box[0] - spread, box[1] - spread, box[2] + spread, box[3] + spread),
        radius,
        rgba(color, alpha),
    )
    layer = layer.filter(ImageFilter.GaussianBlur(blur))
    base.alpha_composite(layer)


def text(draw, xy, s, font, fill=TEXT, anchor="la"):
    draw.text(xy, s, font=font, fill=fill, anchor=anchor)


def pointer(draw, x, y, scale=1.0, fill="#1f2937"):
    pts = [(0, 0), (36, 22), (23, 24), (34, 50), (25, 54), (16, 30), (0, 48)]
    pts = [(x + px * scale, y + py * scale) for px, py in pts]
    draw.polygon(pts, fill=fill)
    hi = [(x + 4 * scale, y + 4 * scale), (x + 24 * scale, y + 18 * scale), (x + 16 * scale, y + 20 * scale)]
    draw.polygon(hi, fill=(255, 255, 255, 90))


def progress_track(draw, labels, active_idx):
    start_x, y, spacing = 820, 88, 92
    for idx, label in enumerate(labels):
        cx = start_x + idx * spacing
        if idx < len(labels) - 1:
            draw.line((cx + 16, y, cx + spacing - 16, y), fill="#cfd7cb" if idx >= active_idx else "#aec9a7", width=4)
        fill = "#ffffff" if idx > active_idx else "#ecf4e8"
        outline = "#d2d9d0" if idx > active_idx else ACCENT
        draw.ellipse((cx - 13, y - 13, cx + 13, y + 13), fill=fill, outline=outline, width=3)
        text(draw, (cx, y), str(idx + 1), FONT_SMALL, fill=ACCENT if idx <= active_idx else MUTED, anchor="mm")
        text(draw, (cx, y + 32), label, FONT_SMALL, fill=TEXT if idx == active_idx else MUTED, anchor="mm")


def base_canvas(title, subtitle, labels, active_idx):
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)
    text(draw, (88, 68), title, FONT_TITLE)
    text(draw, (88, 108), subtitle, FONT_BODY, fill=MUTED)
    progress_track(draw, labels, active_idx)
    shadowed_card(img, (72, 144, 1208, 646), 24)
    return img


def workbook_shell(draw, box, sheet_name, formula, active_ref, ribbon_tab="常用"):
    rounded(draw, box, 20, "#fbfbf8", outline="#dad4c8")
    tx0, ty0, tx1, ty1 = box
    rounded(draw, (tx0 + 18, ty0 + 18, tx1 - 18, ty0 + 62), 18, "#eef4ea", outline="#dde6d9")
    x = tx0 + 36
    for tab in ["常用", "插入", "資料", "設計"]:
        width = 58
        fill = "#ffffff" if tab == ribbon_tab else "#eef4ea"
        rounded(draw, (x, ty0 + 28, x + width, ty0 + 52), 12, fill, outline="#d7e0d4" if tab == ribbon_tab else None)
        text(draw, (x + width / 2, ty0 + 40), tab, FONT_SMALL, fill=ACCENT if tab == ribbon_tab else MUTED, anchor="mm")
        x += width + 10
    rounded(draw, (tx0 + 18, ty0 + 78, tx0 + 98, ty0 + 116), 12, "#f2f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 58, ty0 + 97), active_ref, FONT_MONO_16, fill="#57716b", anchor="mm")
    rounded(draw, (tx0 + 108, ty0 + 78, tx0 + 146, ty0 + 116), 12, "#f2f4ef", outline="#d7dbd0")
    text(draw, (tx0 + 127, ty0 + 97), "fx", FONT_MONO_16, fill=MUTED, anchor="mm")
    rounded(draw, (tx0 + 156, ty0 + 78, tx1 - 18, ty0 + 116), 12, "#ffffff", outline="#d7dbd0")
    text(draw, (tx0 + 176, ty0 + 97), formula, FONT_MONO_16, fill=TEXT, anchor="lm")
    rounded(draw, (tx0 + 18, ty1 - 52, tx0 + 126, ty1 - 16), 12, "#ecf3e8", outline="#d7dbd0")
    text(draw, (tx0 + 72, ty1 - 34), sheet_name, FONT_SMALL, fill=ACCENT, anchor="mm")


def draw_table(base, box, columns, rows, col_tones=None, cell_tones=None, bars=None):
    col_tones = col_tones or {}
    cell_tones = cell_tones or {}
    bars = bars or {}
    draw = ImageDraw.Draw(base)
    x0, y0, x1, y1 = box
    cols = len(columns)
    rows_count = len(rows)
    row_h = (y1 - y0) / (rows_count + 1)
    col_w = (x1 - x0) / cols
    rounded(draw, box, 16, "#ffffff", outline="#d9d4c9")
    for c in range(cols + 1):
        xx = x0 + c * col_w
        draw.line((xx, y0, xx, y1), fill="#ded8cb", width=1)
    for r in range(rows_count + 2):
        yy = y0 + r * row_h
        draw.line((x0, yy, x1, yy), fill="#ded8cb", width=1)

    for c, label in enumerate(columns):
        head_box = (x0 + c * col_w, y0, x0 + (c + 1) * col_w, y0 + row_h)
        fill = "#eef3ea"
        if col_tones.get(c) == "select":
            fill = "#eef6ff"
        elif col_tones.get(c) == "result":
            fill = "#eef7ef"
        elif col_tones.get(c) == "focus":
            fill = "#fff4dd"
        rounded(draw, head_box, 0, fill)
        text(draw, ((head_box[0] + head_box[2]) / 2, (head_box[1] + head_box[3]) / 2), label, FONT_MONO_16, fill="#5b6358", anchor="mm")

    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cb = (x0 + c * col_w, y0 + (r + 1) * row_h, x0 + (c + 1) * col_w, y0 + (r + 2) * row_h)
            tone = cell_tones.get((r, c))
            fill = "#ffffff"
            color = TEXT
            if tone == "select":
                fill = "#f3f8ff"
                color = "#2f4c78"
            elif tone == "result":
                fill = "#edf9ef"
                color = "#2f6b46"
            elif tone == "focus":
                fill = "#fff5e3"
                color = "#915e12"
            elif tone == "danger":
                fill = "#fceaea"
                color = "#9b3a3a"
            elif tone == "warn":
                fill = "#fff7d9"
                color = "#8a6611"
            elif tone == "muted":
                fill = "#f5f3ee"
                color = MUTED
            rounded(draw, cb, 0, fill)
            if (r, c) in bars:
                width = clamp(bars[(r, c)]) * (cb[2] - cb[0] - 22)
                rounded(draw, (cb[0] + 10, cb[1] + 10, cb[0] + 10 + width, cb[3] - 10), 10, rgba("#7abf7b", 120))
            anchor = "lm" if c < 2 else "rm"
            tx = cb[0] + 12 if c < 2 else cb[2] - 12
            text(draw, (tx, (cb[1] + cb[3]) / 2), str(value), FONT_MONO_16, fill=color, anchor=anchor)


def draw_side_panel(draw, box, title, lines, tint="#fffdf8", title_color=TEXT):
    rounded(draw, box, 18, tint, outline="#d7d0c4")
    text(draw, (box[0] + 18, box[1] + 24), title, FONT_LABEL, fill=title_color)
    for idx, line in enumerate(lines):
        text(draw, (box[0] + 18, box[1] + 60 + idx * 30), "• " + line, FONT_BODY, fill="#4e5b50")


def draw_dialog(draw, box, title, rows, footer=None, accent=ACCENT):
    rounded(draw, box, 18, "#fffdf8", outline="#d9d1c4")
    rounded(draw, (box[0], box[1], box[2], box[1] + 44), 18, "#f5f1e6")
    text(draw, (box[0] + 18, box[1] + 24), title, FONT_LABEL, fill=TEXT)
    y = box[1] + 62
    for label, value, tone in rows:
        text(draw, (box[0] + 18, y), label, FONT_SMALL, fill=MUTED)
        rounded(draw, (box[0] + 132, y - 12, box[2] - 18, y + 14), 10, "#ffffff", outline="#d7d0c4")
        if tone == "accent":
            soft_glow_img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            soft_glow(soft_glow_img, (box[0] + 132, y - 12, box[2] - 18, y + 14), accent, alpha=70, blur=14, spread=8)
            draw.bitmap((0, 0), soft_glow_img)
        text(draw, (box[0] + 148, y + 1), value, FONT_SMALL, fill=TEXT, anchor="lm")
        y += 42
    if footer:
        rounded(draw, (box[0] + 18, box[3] - 54, box[2] - 18, box[3] - 18), 12, ACCENT_SOFT, outline="#d2dbc8")
        text(draw, (box[0] + 36, box[3] - 36), footer, FONT_SMALL, fill=ACCENT, anchor="lm")


def draw_dropdown(draw, box, options, active_idx=0):
    rounded(draw, box, 12, "#ffffff", outline="#d7d0c4")
    row_h = (box[3] - box[1]) / len(options)
    for idx, option in enumerate(options):
        row_box = (box[0], box[1] + idx * row_h, box[2], box[1] + (idx + 1) * row_h)
        if idx == active_idx:
            rounded(draw, row_box, 0, "#eef6ff")
        text(draw, (box[0] + 14, (row_box[1] + row_box[3]) / 2), option, FONT_SMALL, fill=TEXT, anchor="lm")


def draw_chart(draw, box, months, bars, line, polished=False):
    rounded(draw, box, 18, "#ffffff", outline="#d9d4c9")
    x0, y0, x1, y1 = box
    plot = (x0 + 48, y0 + 46, x1 - 24, y1 - 44)
    for idx in range(4):
        yy = plot[1] + idx * (plot[3] - plot[1]) / 3
        draw.line((plot[0], yy, plot[2], yy), fill="#ebe5d7", width=1)
    step = (plot[2] - plot[0]) / len(months)
    max_bar = max(bars)
    points = []
    for idx, month in enumerate(months):
        cx = plot[0] + step * idx + step / 2
        bar_h = (bars[idx] / max_bar) * (plot[3] - plot[1] - 16)
        bar_box = (cx - 18, plot[3] - bar_h, cx + 18, plot[3])
        rounded(draw, bar_box, 10, "#cfe7cf" if polished else "#dfe7f1")
        line_y = plot[3] - (line[idx] / max(line)) * (plot[3] - plot[1] - 24)
        points.append((cx, line_y))
        text(draw, (cx, plot[3] + 20), month, FONT_SMALL, fill=MUTED, anchor="mm")
        if polished:
            text(draw, (cx, bar_box[1] - 10), f"{bars[idx]:,}", FONT_SMALL, fill="#2f6b46", anchor="mb")
    draw.line(points, fill=PURPLE if polished else BLUE, width=4, joint="curve")
    for idx, point in enumerate(points):
        draw.ellipse((point[0] - 6, point[1] - 6, point[0] + 6, point[1] + 6), fill="#ffffff", outline=PURPLE if polished else BLUE, width=3)
        if polished:
            text(draw, (point[0], point[1] - 14), f"{line[idx]}%", FONT_SMALL, fill=PURPLE, anchor="mb")


def draw_relation(draw, start, end, label):
    draw.line((start, end), fill=PURPLE, width=4)
    draw.ellipse((start[0] - 6, start[1] - 6, start[0] + 6, start[1] + 6), fill="#ffffff", outline=PURPLE, width=3)
    draw.ellipse((end[0] - 6, end[1] - 6, end[0] + 6, end[1] + 6), fill="#ffffff", outline=PURPLE, width=3)
    mid = ((start[0] + end[0]) / 2, (start[1] + end[1]) / 2)
    rounded(draw, (mid[0] - 44, mid[1] - 16, mid[0] + 44, mid[1] + 12), 12, "#f4ecff", outline="#d5c4f4")
    text(draw, mid, label, FONT_SMALL, fill=PURPLE, anchor="mm")


def blend_sequence(scenes, hold=540, trans_frames=8, trans_ms=115, final_hold=760):
    frames = []
    durations = []
    rendered = [scene() for scene in scenes]
    for idx, img in enumerate(rendered[:-1]):
        frames.append(img)
        durations.append(hold)
        for step in range(1, trans_frames + 1):
            t = step / (trans_frames + 1)
            frames.append(Image.blend(img, rendered[idx + 1], t))
            durations.append(trans_ms)
    frames.append(rendered[-1])
    durations.append(final_hold)
    return frames, durations


def save_animation(name, scenes):
    frames, durations = blend_sequence(scenes)
    path = MEDIA_DIR / name
    frames[0].save(
        path,
        format="WEBP",
        save_all=True,
        append_images=frames[1:],
        duration=durations,
        loop=0,
        quality=86,
        method=6,
    )
    print(path)


def conditional_formatting_scenes():
    labels = ["選欄位", "設規則", "上色", "看重點"]
    months = ["Amy", "Ben", "Cara", "Dora", "Evan"]
    revenue = ["82,000", "61,000", "104,000", "57,000", "91,000"]
    growth = ["12%", "-6%", "18%", "-3%", "8%"]
    complaints = ["1", "4", "0", "3", "5"]

    def scene1():
        img = base_canvas("條件式格式化：讓重點自己跳出來", "先選欄位，再用資料橫條和醒目規則把主管要看的東西浮出來", labels, 0)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Sales_KPI", "先選取要加格式的欄位：業績 / 成長率 / 客訴數", "B2")
        draw_table(
            img,
            (120, 240, 844, 562),
            ["業務員", "業績", "成長率", "客訴數"],
            list(zip(months, revenue, growth, complaints)),
            col_tones={1: "select"},
            cell_tones={(idx, 1): "select" for idx in range(5)},
        )
        draw_side_panel(draw, (892, 188, 1182, 360), "這一步先做", ["先選會影響判讀的欄位", "業績適合橫條，成長率適合紅綠", "客訴數適合警示色"], tint="#fffdf8")
        pointer(draw, 532, 308, 1.04)
        return img

    def scene2():
        img = base_canvas("條件式格式化：先決定規則，不要亂上色", "同一張表可以疊多層規則，但每一層都要有明確目的", labels, 1)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Sales_KPI", "常用 → 條件式格式化", "B2")
        draw_table(
            img,
            (120, 240, 844, 562),
            ["業務員", "業績", "成長率", "客訴數"],
            list(zip(months, revenue, growth, complaints)),
            col_tones={1: "select"},
            cell_tones={(idx, 1): "select" for idx in range(5)},
        )
        soft_glow(img, (886, 212, 1186, 520), AMBER, alpha=76)
        draw_dialog(
            draw,
            (910, 208, 1178, 522),
            "條件式格式化",
            [
                ("規則 1", "資料橫條：業績", "accent"),
                ("規則 2", "小於 0 = 紅色", None),
                ("規則 3", "客訴數 > 3 = 黃底", None),
            ],
            footer="先讓數字能比較，再讓異常被提醒",
            accent=AMBER,
        )
        pointer(draw, 1088, 306, 1.02)
        return img

    def scene3():
        img = base_canvas("條件式格式化：資料自己開始說故事", "資料橫條先建立相對大小感，主管不用看完整數字也知道誰高誰低", labels, 2)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Sales_KPI", "業績欄套用資料橫條；成長率和客訴數套用警示規則", "C3")
        bars = {(idx, 1): val for idx, val in enumerate([0.79, 0.58, 1.0, 0.52, 0.87])}
        cell_tones = {(idx, 1): "result" for idx in range(5)}
        cell_tones[(1, 2)] = "danger"
        cell_tones[(3, 2)] = "danger"
        cell_tones[(1, 3)] = "warn"
        cell_tones[(4, 3)] = "warn"
        draw_table(
            img,
            (120, 240, 844, 562),
            ["業務員", "業績", "成長率", "客訴數"],
            list(zip(months, revenue, growth, complaints)),
            col_tones={1: "result", 2: "focus", 3: "focus"},
            cell_tones=cell_tones,
            bars=bars,
        )
        draw_side_panel(draw, (892, 188, 1182, 360), "現在一眼能看到", ["Cara 業績最高", "Ben、Dora 成長率轉負", "Ben、Evan 客訴數過高"], tint="#eef7ef", title_color="#2f6b46")
        pointer(draw, 726, 362, 1.03)
        return img

    def scene4():
        img = base_canvas("條件式格式化完成：主管 3 秒就知道誰要追", "格式不是裝飾，而是把要追蹤、要表揚的人直接浮上來", labels, 3)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Sales_KPI", "先比較，再提醒，最後才考慮美化", "D6")
        bars = {(idx, 1): val for idx, val in enumerate([0.79, 0.58, 1.0, 0.52, 0.87])}
        cell_tones = {(idx, 1): "result" for idx in range(5)}
        cell_tones[(1, 2)] = "danger"
        cell_tones[(3, 2)] = "danger"
        cell_tones[(1, 3)] = "warn"
        cell_tones[(4, 3)] = "warn"
        draw_table(img, (120, 240, 844, 562), ["業務員", "業績", "成長率", "客訴數"], list(zip(months, revenue, growth, complaints)), col_tones={1: "result", 2: "focus", 3: "focus"}, cell_tones=cell_tones, bars=bars)
        draw_side_panel(draw, (892, 188, 1182, 392), "報表完成後", ["高績效：Cara、Evan", "優先追蹤：Ben（負成長 + 客訴偏高）", "只靠黑白表格時，這些重點很難一眼看出"], tint="#fff7e6", title_color="#8a5c12")
        soft_glow(img, (120, 240, 844, 562), GREEN, alpha=54)
        pointer(draw, 728, 470, 1.02)
        return img

    return [scene1, scene2, scene3, scene4]


def data_validation_scenes():
    labels = ["建規則", "開下拉", "擋錯誤", "交表"]
    rows = [
        ("部門", "", "", ""),
        ("到職日", "", "", ""),
        ("手機", "", "", ""),
        ("班別", "", "", ""),
    ]

    def draw_form(img, filled=None, tone_map=None):
        filled = filled or {}
        tone_map = tone_map or {}
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Onboarding_Form", "資料 → 資料驗證", "B2", ribbon_tab="資料")
        table_rows = []
        cell_tones = {}
        for idx, row in enumerate(rows):
            dept = filled.get(idx, "")
            table_rows.append([row[0], dept, row[2], row[3]])
            tone = tone_map.get(idx)
            if tone:
                cell_tones[(idx, 1)] = tone
        draw_table(img, (120, 240, 844, 530), ["欄位", "輸入值", "", ""], table_rows, col_tones={1: "select"}, cell_tones=cell_tones)

    def scene1():
        img = base_canvas("資料驗證：把錯誤擋在入口，而不是事後補救", "先為部門欄建立清單，之後手機、日期和班別都能用同一套思路", labels, 0)
        draw_form(img)
        draw = ImageDraw.Draw(img)
        draw_side_panel(draw, (892, 188, 1182, 360), "這支短片示範", ["部門：下拉選單", "手機：10 碼限制", "錯誤提醒：不允許亂填"], tint="#fffdf8")
        pointer(draw, 512, 308, 1.03)
        return img

    def scene2():
        img = base_canvas("資料驗證：先把合法值定義好", "同一欄只允許選『業務、財務、資訊、人事』，就不會再有人自由發揮", labels, 1)
        draw_form(img)
        draw = ImageDraw.Draw(img)
        soft_glow(img, (906, 212, 1176, 512), BLUE, alpha=76)
        draw_dialog(draw, (912, 208, 1178, 516), "資料驗證", [("允許", "清單", "accent"), ("來源", "業務, 財務, 資訊, 人事", None), ("錯誤提醒", "停止", None)], footer="合法值先定好，之後只准從清單選", accent=BLUE)
        pointer(draw, 1092, 292, 1.02)
        return img

    def scene3():
        img = base_canvas("資料驗證：使用者看到的是很簡單的下拉選單", "對填表的人來說，只是多了一個下拉；對你來說，錯誤格式已經被擋住", labels, 1)
        draw_form(img, filled={0: "資訊"}, tone_map={0: "result"})
        draw = ImageDraw.Draw(img)
        draw_dropdown(draw, (440, 290, 660, 430), ["業務", "財務", "資訊", "人事"], active_idx=2)
        pointer(draw, 602, 360, 1.0)
        draw_side_panel(draw, (892, 188, 1182, 360), "這裡發生的事", ["使用者只需要選，不需要背格式", "後端資料也會一致", "之後做統計和樞紐都更穩"], tint="#eef7ef", title_color="#2f6b46")
        return img

    def scene4():
        img = base_canvas("資料驗證：亂填時就當場阻止", "錯誤提醒比事後清資料便宜太多，尤其是表單回收量一大時", labels, 2)
        draw_form(img, filled={0: "資訊", 1: "2026/05/01", 2: "09123"}, tone_map={0: "result", 1: "result", 2: "danger"})
        draw = ImageDraw.Draw(img)
        soft_glow(img, (908, 248, 1172, 430), RED, alpha=76)
        draw_dialog(draw, (920, 248, 1170, 430), "錯誤提醒", [("原因", "手機欄必須是 10 碼", None), ("目前輸入", "09123", None)], footer="停止：不允許錯誤值進來", accent=RED)
        pointer(draw, 572, 412, 1.0)
        return img

    def scene5():
        img = base_canvas("資料驗證完成：表單看起來沒變，但資料品質差很多", "真正的專業不是花很多時間修資料，而是讓錯誤根本進不來", labels, 3)
        draw_form(img, filled={0: "資訊", 1: "2026/05/01", 2: "0912345678", 3: "晚班"}, tone_map={0: "result", 1: "result", 2: "result", 3: "result"})
        draw = ImageDraw.Draw(img)
        draw_side_panel(draw, (892, 188, 1182, 392), "交表前已經做到", ["部門只能選合法值", "日期格式一致", "手機長度正確", "後面不用再人工清洗"], tint="#fff7e6", title_color="#8a5c12")
        soft_glow(img, (120, 240, 844, 530), GREEN, alpha=52)
        pointer(draw, 744, 452, 1.02)
        return img

    return [scene1, scene2, scene3, scene4, scene5]


def chart_design_scenes():
    labels = ["選資料", "插圖表", "去雜訊", "講故事"]
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    revenue = [120, 132, 141, 155, 162, 176]
    growth = [6, 8, 7, 10, 4, 9]

    def scene1():
        img = base_canvas("圖表設計：先決定要回答什麼，再決定畫什麼", "這裡要同時看營收成長和成長率，所以最適合的是組合圖，而不是兩張分開的小圖", labels, 0)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Monthly_Report", "先選月份、營收、成長率三欄", "A1", ribbon_tab="插入")
        draw_table(
            img,
            (120, 240, 844, 560),
            ["月份", "營收", "成長率"],
            list(zip(months, [f"{n:,}" for n in revenue], [f"{n}%" for n in growth])),
            col_tones={0: "select", 1: "result", 2: "focus"},
            cell_tones={(idx, 1): "result" for idx in range(6)} | {(idx, 2): "focus" for idx in range(6)},
        )
        draw_side_panel(draw, (892, 188, 1182, 364), "這張圖要回答", ["營收是否持續走高", "成長率在哪個月份放慢", "主管能不能一眼看出趨勢"], tint="#fffdf8")
        pointer(draw, 514, 304, 1.02)
        return img

    def scene2():
        img = base_canvas("圖表設計：原始圖先出來，再慢慢修成專業版本", "先快速插入組合圖，之後再刪掉多餘元素、補標題和標示", labels, 1)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 640, 626), "Monthly_Report", "插入 → 組合圖（營收=直條，成長率=折線）", "A1", ribbon_tab="插入")
        draw_table(img, (120, 240, 616, 560), ["月份", "營收", "成長率"], list(zip(months, [f"{n:,}" for n in revenue], [f"{n}%" for n in growth])), col_tones={0: "select", 1: "result", 2: "focus"})
        soft_glow(img, (684, 236, 1182, 580), BLUE, alpha=48)
        draw_chart(draw, (692, 244, 1172, 574), months, revenue, growth, polished=False)
        pointer(draw, 934, 420, 1.0)
        return img

    def scene3():
        img = base_canvas("圖表設計：刪掉讀者不需要的東西", "格線、重複圖例和無觀點標題，通常都是先刪掉的對象", labels, 2)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 640, 626), "Monthly_Report", "刪除格線、保留必要標示", "E8", ribbon_tab="設計")
        draw_table(img, (120, 240, 616, 560), ["月份", "營收", "成長率"], list(zip(months, [f"{n:,}" for n in revenue], [f"{n}%" for n in growth])), col_tones={0: "select", 1: "result", 2: "focus"})
        draw_chart(draw, (692, 244, 1172, 574), months, revenue, growth, polished=True)
        draw_side_panel(draw, (894, 188, 1178, 226), "這時還在修", ["拿掉不必要元素，讓眼睛只看重點"], tint="#eef7ef", title_color="#2f6b46")
        pointer(draw, 1086, 274, 0.98)
        return img

    def scene4():
        img = base_canvas("圖表設計完成：現在這張圖真的在講一個結論", "同樣的資料，只要換成正確圖型和乾淨標示，整張月報就像升級一個檔次", labels, 3)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 640, 626), "Monthly_Report", "標題改成有觀點：營收連 6 月走高，4 月成長率短暫放慢", "H10", ribbon_tab="設計")
        draw_table(img, (120, 240, 616, 560), ["月份", "營收", "成長率"], list(zip(months, [f"{n:,}" for n in revenue], [f"{n}%" for n in growth])), col_tones={0: "select", 1: "result", 2: "focus"})
        draw_chart(draw, (692, 244, 1172, 574), months, revenue, growth, polished=True)
        rounded(draw, (724, 214, 1138, 248), 14, "#fff7e6", outline="#e6d3aa")
        text(draw, (744, 231), "營收連 6 月走高，4 月成長率短暫放慢", FONT_LABEL, fill="#8a5c12")
        pointer(draw, 978, 378, 1.0)
        return img

    return [scene1, scene2, scene3, scene4]


def power_query_scenes():
    labels = ["接來源", "記步驟", "附加月份", "重新整理"]

    def scene1():
        img = base_canvas("Power Query：第一次是在建流程，不是在趕快手修", "先把很醜的來源接進來，後面每一步整理才會被記住", labels, 0)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "March.csv", "來源很醜沒關係，先接進來", "A1", ribbon_tab="資料")
        draw_table(
            img,
            (120, 240, 844, 558),
            ["訂單日期", "部門", "金額", "備註"],
            [["2026/03/01", "電子", "32,000", "北區"], ["2026/03/04", "家電", "18,500", "南區"], ["2026/03/05", "電子", "29,000", "北區"], ["2026/03/08", "家電", "21,300", "急件"]],
            col_tones={3: "focus"},
            cell_tones={(idx, 3): "focus" for idx in range(4)},
        )
        draw_side_panel(draw, (892, 188, 1182, 360), "這一步的心態", ["不要在工作表手改", "Power Query 專門接這種髒資料", "先建立連線再開始整理"], tint="#fffdf8")
        pointer(draw, 532, 304, 1.02)
        return img

    def scene2():
        img = base_canvas("Power Query：整理值不稀奇，記住整理步驟才值錢", "改型別、移除備註欄、重新命名，這些都要變成可重跑的步驟", labels, 1)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 820, 626), "Query1", "已記住：變更型別 / 移除欄位 / 重新命名", "A2", ribbon_tab="資料")
        draw_table(
            img,
            (120, 240, 796, 558),
            ["訂單日期", "部門", "金額"],
            [["2026-03-01", "電子", "32,000"], ["2026-03-04", "家電", "18,500"], ["2026-03-05", "電子", "29,000"], ["2026-03-08", "家電", "21,300"]],
            col_tones={2: "result"},
            cell_tones={(idx, 2): "result" for idx in range(4)},
        )
        draw_side_panel(draw, (848, 188, 1182, 422), "已記住的步驟", ["變更型別：日期 → 日期", "變更型別：金額 → 數值", "移除欄位：備註", "重新命名欄位"], tint="#eef7ef", title_color="#2f6b46")
        pointer(draw, 954, 346, 1.0)
        return img

    def scene3():
        img = base_canvas("Power Query：每月資料不要手貼，用附加查詢接成長表", "一旦月份結構一致，就該做成可以持續吃新資料的正式分析表", labels, 2)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Fact_Sales", "把 1 月、2 月、3 月資料附加成同一張表", "A2", ribbon_tab="資料")
        draw_table(
            img,
            (120, 240, 844, 558),
            ["月份", "部門", "金額"],
            [["1月", "電子", "120,000"], ["1月", "家電", "80,000"], ["2月", "電子", "95,000"], ["2月", "家電", "110,000"], ["3月", "電子", "132,000"], ["3月", "家電", "98,000"]],
            col_tones={2: "result"},
            cell_tones={(idx, 2): "result" for idx in range(6)},
        )
        draw_side_panel(draw, (892, 188, 1182, 390), "現在這張表可拿去", ["樞紐分析", "圖表與儀表板", "資料模型和 Power Pivot"], tint="#fff7e6", title_color="#8a5c12")
        pointer(draw, 734, 472, 1.0)
        return img

    def scene4():
        img = base_canvas("Power Query完成：下個月只是按重新整理，不是重做一次", "真正省時間的，是四月、五月、六月都能沿用這條流程", labels, 3)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Monthly_Report", "新增 2026-04.csv 後按重新整理，報表自動更新", "A5", ribbon_tab="資料")
        draw_table(
            img,
            (120, 240, 844, 558),
            ["月份", "電子", "家電", "總計"],
            [["1月", "120,000", "80,000", "200,000"], ["2月", "95,000", "110,000", "205,000"], ["3月", "132,000", "98,000", "230,000"], ["4月", "141,000", "105,000", "246,000"]],
            col_tones={3: "result"},
            cell_tones={(3, 0): "focus", (3, 1): "focus", (3, 2): "focus", (3, 3): "result"},
        )
        draw_side_panel(draw, (892, 188, 1182, 392), "四月只做這些", ["把新檔放進資料夾", "按重新整理", "既有步驟自動套用", "報表和樞紐同步更新"], tint="#eef7ef", title_color="#2f6b46")
        soft_glow(img, (120, 458, 844, 558), GREEN, alpha=56)
        pointer(draw, 622, 512, 1.0)
        return img

    return [scene1, scene2, scene3, scene4]


def data_model_scenes():
    labels = ["分開表", "建關聯", "寫量值", "交給樞紐"]

    def scene1():
        img = base_canvas("資料模型：先分清楚誰是交易，誰是主檔", "不要把所有欄位硬攤成一張超大表，先讓每張表保留自己的角色", labels, 0)
        draw = ImageDraw.Draw(img)
        rounded(draw, (114, 218, 598, 566), 18, "#fffdf8", outline="#d7cfbf")
        text(draw, (138, 246), "Orders（事實表）", FONT_LABEL, fill=TEXT)
        draw_table(img, (136, 274, 576, 536), ["訂單ID", "產品ID", "金額"], [["SO-001", "P001", "32,000"], ["SO-002", "P003", "18,500"], ["SO-003", "P001", "29,000"]], col_tones={2: "result"}, cell_tones={(0, 2): "result", (2, 2): "result"})
        rounded(draw, (656, 218, 1140, 566), 18, "#fffdf8", outline="#d7cfbf")
        text(draw, (680, 246), "Products（維度表）", FONT_LABEL, fill=TEXT)
        draw_table(img, (678, 274, 1118, 536), ["產品ID", "產品名稱", "類別"], [["P001", "無線滑鼠", "周邊"], ["P003", "27 吋螢幕", "顯示"], ["P005", "網路攝影機", "周邊"]], col_tones={0: "focus"}, cell_tones={(0, 0): "focus", (1, 0): "focus"})
        pointer(draw, 790, 322, 1.02)
        return img

    def scene2():
        img = base_canvas("資料模型：真正的核心不是公式，而是關聯", "只要產品 ID 對得上，訂單表就能借用產品表的類別與名稱", labels, 1)
        draw = ImageDraw.Draw(img)
        rounded(draw, (114, 218, 598, 566), 18, "#fffdf8", outline="#d7cfbf")
        text(draw, (138, 246), "Orders", FONT_LABEL, fill=TEXT)
        draw_table(img, (136, 274, 576, 536), ["訂單ID", "產品ID", "金額"], [["SO-001", "P001", "32,000"], ["SO-002", "P003", "18,500"], ["SO-003", "P001", "29,000"]], col_tones={1: "focus"}, cell_tones={(0, 1): "focus", (1, 1): "focus", (2, 1): "focus"})
        rounded(draw, (656, 218, 1140, 566), 18, "#fffdf8", outline="#d7cfbf")
        text(draw, (680, 246), "Products", FONT_LABEL, fill=TEXT)
        draw_table(img, (678, 274, 1118, 536), ["產品ID", "產品名稱", "類別"], [["P001", "無線滑鼠", "周邊"], ["P003", "27 吋螢幕", "顯示"], ["P005", "網路攝影機", "周邊"]], col_tones={0: "focus"}, cell_tones={(0, 0): "focus", (1, 0): "focus"})
        draw_relation(draw, (516, 366), (678, 366), "產品ID")
        pointer(draw, 620, 340, 1.0)
        return img

    def scene3():
        img = base_canvas("資料模型：指標要寫在模型裡，不要散落在每張樞紐裡", "量值只定義一次，之後任何報表都共用同一套商業邏輯", labels, 2)
        draw = ImageDraw.Draw(img)
        rounded(draw, (140, 230, 1140, 546), 22, "#fffdf8", outline="#d8d0bf")
        text(draw, (170, 264), "Measure", FONT_LABEL, fill=ACCENT)
        rounded(draw, (170, 294, 1110, 356), 16, "#171a1f")
        text(draw, (198, 326), "總業績 := SUM(Orders[金額])", FONT_MONO_18, fill="#eef2f7", anchor="lm")
        rounded(draw, (170, 382, 1110, 444), 16, "#171a1f")
        text(draw, (198, 414), "周邊業績 := CALCULATE([總業績], Products[類別] = \"周邊\")", FONT_MONO_16, fill="#eef2f7", anchor="lm")
        draw_side_panel(draw, (170, 468, 1110, 526), "這裡的觀念", ["量值寫一次，多張樞紐共用", "要改商業邏輯時，只改模型，不是每張報表都改"], tint="#eef7ef", title_color="#2f6b46")
        pointer(draw, 458, 316, 1.0)
        return img

    def scene4():
        img = base_canvas("資料模型完成：樞紐只負責展示，不再背負計算邏輯", "當關聯和量值都在模型裡定好，樞紐分析表只是出口，而不是算式本體", labels, 3)
        draw = ImageDraw.Draw(img)
        workbook_shell(draw, (92, 166, 868, 626), "Model_Pivot", "類別 × 總業績（由模型量值提供）", "B4", ribbon_tab="插入")
        draw_table(img, (120, 240, 844, 518), ["類別", "總業績"], [["周邊", "61,000"], ["顯示", "18,500"]], col_tones={1: "result"}, cell_tones={(0, 1): "result", (1, 1): "result"})
        draw_side_panel(draw, (892, 188, 1182, 392), "現在的差別", ["樞紐只是切片器與展示層", "邏輯集中在模型裡", "之後加新樞紐也沿用同一個量值"], tint="#fff7e6", title_color="#8a5c12")
        soft_glow(img, (120, 240, 844, 518), GREEN, alpha=54)
        pointer(draw, 548, 352, 1.0)
        return img

    return [scene1, scene2, scene3, scene4]


def main():
    save_animation("p2-04-conditional-demo.webp", conditional_formatting_scenes())
    save_animation("p3-01-validation-demo.webp", data_validation_scenes())
    save_animation("p3-03-chart-demo.webp", chart_design_scenes())
    save_animation("p4-04-power-query-demo.webp", power_query_scenes())
    save_animation("p4-05-data-model-demo.webp", data_model_scenes())


if __name__ == "__main__":
    main()
