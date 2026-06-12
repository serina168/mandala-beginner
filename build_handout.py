# -*- coding: utf-8 -*-
"""一級講師培訓課程・學員講義 (A4 直式 21 頁)
   Run:  python build_handout.py
   Out:  mandala_l1_handout.pptx
"""
import os, math, base64

def _fix_img(path):
    try:
        with open(path, 'rb') as f:
            b = f.read(1)
        if b and b[0] < 0x80:
            with open(path, 'r') as f:
                data = f.read().strip()
            missing = len(data) % 4
            if missing:
                data += '=' * (4 - missing)
            decoded = base64.b64decode(data)
            with open(path, 'wb') as f:
                f.write(decoded)
    except Exception:
        pass

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image, ImageDraw

L1   = "/home/user/mandala/l1"
DIAG = f"{L1}/diagrams"

for _img in [
    f"{L1}/bg_ending.jpg", f"{L1}/bg_watercolor.jpg", f"{L1}/bg_cover.jpg",
    f"{L1}/art1.jpg", f"{L1}/art2.jpg", f"{L1}/art3.jpg",
    f"{L1}/art4.jpg", f"{L1}/art5.jpg", f"{L1}/art6.jpg",
    f"{L1}/art_octagon4.jpg",
    f"{DIAG}/el_dots.png", f"{DIAG}/el_geometry.png", f"{DIAG}/el_leaves.png",
    f"{DIAG}/el_lines.png", f"{DIAG}/el_petals.png", f"{DIAG}/el_teardrop.png",
    f"{DIAG}/structure.png",
]:
    if os.path.exists(_img):
        _fix_img(_img)

def _make_div_png(n, path, size=400):
    img = Image.new("RGB", (size, size), (0xFB, 0xF6, 0xEF))
    dr = ImageDraw.Draw(img)
    cx = cy = size // 2; r = size // 2 - 18
    dr.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(0xF3, 0xE7, 0xD8),
               outline=(0xC9, 0xA9, 0x6E), width=6)
    for k in range(n):
        ang = math.radians(k * 360.0 / n - 90)
        ex = int(cx + r * math.cos(ang)); ey = int(cy + r * math.sin(ang))
        dr.line([cx, cy, ex, ey], fill=(0xA9, 0x5C, 0x66), width=4)
    dr.ellipse([cx-r, cy-r, cx+r, cy+r], fill=None, outline=(0xC9, 0xA9, 0x6E), width=6)
    dr.ellipse([cx-13, cy-13, cx+13, cy+13], fill=(0xC9, 0x7B, 0x84))
    img.save(path)

_DIV = {}
for _n in [4, 6, 8, 12]:
    _DIV[_n] = f"/tmp/_ho_div_{_n}.png"
    _make_div_png(_n, _DIV[_n])

# ── 色系 ─────────────────────────────────────────────────────
CREAM      = RGBColor(0xFB, 0xF6, 0xEF)
CREAM_DEEP = RGBColor(0xF3, 0xE7, 0xD8)
ROSE       = RGBColor(0xC9, 0x7B, 0x84)
ROSE_DEEP  = RGBColor(0xA9, 0x5C, 0x66)
GOLD       = RGBColor(0xC9, 0xA9, 0x6E)
TAUPE      = RGBColor(0x8A, 0x7A, 0x6A)
INK        = RGBColor(0x4A, 0x40, 0x3A)
SAGE       = RGBColor(0x9C, 0xA9, 0x8C)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

WHEEL = [
    ("紅",   RGBColor(0xE8, 0x43, 0x4B)),
    ("紅橙", RGBColor(0xE8, 0x6A, 0x3A)),
    ("橙",   RGBColor(0xE8, 0x9A, 0x3A)),
    ("黃橙", RGBColor(0xE8, 0xC2, 0x4A)),
    ("黃",   RGBColor(0xE6, 0xDA, 0x55)),
    ("黃綠", RGBColor(0xA8, 0xC8, 0x4A)),
    ("綠",   RGBColor(0x5A, 0xA8, 0x5A)),
    ("藍綠", RGBColor(0x4A, 0xA8, 0x9A)),
    ("藍",   RGBColor(0x4A, 0x7A, 0xC8)),
    ("藍紫", RGBColor(0x6A, 0x5A, 0xB8)),
    ("紫",   RGBColor(0x9A, 0x5A, 0xB8)),
    ("紅紫", RGBColor(0xC8, 0x4A, 0x8A)),
]

EMU       = 914400
PW, PH    = 8.27, 11.69   # A4 portrait inches
ML = MR   = 0.55
CW        = PW - ML - MR  # 7.17"
FONT      = "微軟正黑體"
FS        = 1.20           # global font scale (1.20 = all text 20% larger)

prs = Presentation()
prs.slide_width  = Emu(int(PW * EMU))
prs.slide_height = Emu(int(PH * EMU))
BLANK = prs.slide_layouts[6]

BG_IMG  = f"{L1}/bg_watercolor.jpg"
_HAS_BG = os.path.exists(BG_IMG)

def slide():
    return prs.slides.add_slide(BLANK)

def bg(s, color=CREAM):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color
    if color == CREAM and _HAS_BG:
        p = s.shapes.add_picture(BG_IMG, Inches(0), Inches(0), Inches(PW), Inches(PH))
        p.line.fill.background()
        ov = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(PW), Inches(PH))
        ov.fill.solid(); ov.fill.fore_color.rgb = CREAM; ov.line.fill.background()
        ov.shadow.inherit = False
        set_alpha(ov, 80)

def rect(s, x, y, w, h, color, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def set_alpha(shape, alpha_pct):
    sp = shape.fill._xPr.find(qn('a:solidFill'))
    srgb = sp.find(qn('a:srgbClr'))
    a = srgb.makeelement(qn('a:alpha'), {'val': str(int(alpha_pct * 1000))})
    srgb.append(a)

def txt(s, x, y, w, h, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, font=FONT, spacing=1.0, italic=False):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = spacing
        r = p.add_run(); r.text = ln
        f = r.font; f.size = Pt(size * FS); f.bold = bold; f.italic = italic
        f.name = font; f.color.rgb = color
    return tb

def pic_cover(s, path, x, y, w, h, border=True):
    iw, ih = Image.open(path).size
    tw, th  = w * EMU, h * EMU
    sr, ir  = tw / th, iw / ih
    if ir > sr:
        cw_img = int(ih * sr); off = (iw - cw_img) // 2
        crop = (off / iw, 0, (off + cw_img) / iw, 1)
    else:
        ch_img = int(iw / sr); off = (ih - ch_img) // 2
        crop = (0, off / ih, 1, (off + ch_img) / ih)
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    p.crop_left, p.crop_top   = crop[0], crop[1]
    p.crop_right, p.crop_bottom = 1 - crop[2], 1 - crop[3]
    if border:
        p.line.color.rgb = WHITE; p.line.width = Pt(2.0)
    else:
        p.line.fill.background()
    return p

def color_wheel(s, cx, cy, r, dot=0.42):
    for i, (name, col) in enumerate(WHEEL):
        ang = math.radians(-90 + i * 30)
        px = cx + r * math.cos(ang) - dot / 2
        py = cy + r * math.sin(ang) - dot / 2
        rect(s, px, py, dot, dot, col, line=WHITE, lw=1.2, shape=MSO_SHAPE.OVAL)
        lx = cx + (r + 0.52) * math.cos(ang) - 0.33
        ly = cy + (r + 0.52) * math.sin(ang) - 0.14
        txt(s, lx, ly, 0.66, 0.28, name, 9, color=TAUPE, align=PP_ALIGN.CENTER)

def header(s, eb, title, size=22):
    rect(s, 0, 0, PW, 1.08, CREAM_DEEP)
    txt(s, ML, 0.20, CW, 0.32, eb, 10, color=GOLD, bold=True)
    txt(s, ML, 0.52, CW, 0.54, title, size, color=INK, bold=True)

def pgnum(s, n):
    txt(s, 0, PH - 0.34, PW, 0.28, f"— {n} —", 9, color=TAUPE, align=PP_ALIGN.CENTER)

# ============================================================
# P1  封面
# ============================================================
s = slide()
s.background.fill.solid(); s.background.fill.fore_color.rgb = CREAM
BG_COVER = f"{L1}/bg_cover.jpg"
if os.path.exists(BG_COVER):
    p = s.shapes.add_picture(BG_COVER, Inches(0), Inches(0), Inches(PW), Inches(PH))
    p.line.fill.background()
txt(s, ML, 2.10, CW, 0.44,
    "財團法人中華綜合發展研究院／文創藝術研究所", 14, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, ML, 2.56, CW, 0.40,
    "心靈藝術美學中心", 15, color=TAUPE, bold=True, align=PP_ALIGN.CENTER)
txt(s, ML, 3.12, CW, 0.96, "一級講師培訓課程", 48, color=INK,      bold=True, align=PP_ALIGN.CENTER)
txt(s, ML, 4.20, CW, 0.68, "曼陀羅色彩藝術",   28, color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
txt(s, ML, 5.02, CW, 0.72, "學  員  講  義",   34, color=INK,      bold=True, align=PP_ALIGN.CENTER)
txt(s, ML, PH - 1.0, CW, 0.42,
    "請妥善保存，作為日後教學參考", 13, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, ML, PH - 0.50, CW, 0.34,
    "版權所有／翻印必究", 10, color=TAUPE, align=PP_ALIGN.CENTER)

# ============================================================
# P2  課程六大收穫
# ============================================================
s = slide(); bg(s)
header(s, "COURSE  HIGHLIGHTS", "一級講師課程・六大收穫")
pgnum(s, 2)
goals = [
    ("01", "進階色彩學",    "色相、明度、彩度、色彩心理學，從理解到精準運用"),
    ("02", "深化理解與應用","作品解析力，洞察色彩背後的情感與內在狀態"),
    ("03", "大型木器創作",  "在大型八角板木器上完成一件精緻曼陀羅作品"),
    ("04", "高級鑽飾裝飾",  "施華洛世奇水晶鑽飾貼附，提升作品至展覽水準"),
    ("05", "立體畫法",      "進階筆法與層次技巧，讓圖案呈現真實立體質感"),
    ("06", "教學引導訓練",  "實際演練教學流程，從「會畫」躍升為「能教」"),
]
cw_g = (CW - 0.22) / 2; ch_g = 1.58; gy_g = 0.22
for i, (n, t, d) in enumerate(goals):
    r, c = divmod(i, 2)
    x = ML + c * (cw_g + 0.22); y = 1.25 + r * (ch_g + gy_g)
    rect(s, x, y, cw_g, ch_g, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, cw_g, 0.42, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + 0.15, y + 0.04, cw_g - 0.15, 0.38, CREAM_DEEP)
    txt(s, x + 0.14, y + 0.07, 0.4, 0.3, n, 11, color=ROSE_DEEP, bold=True)
    txt(s, x + 0.6,  y + 0.06, cw_g - 0.75, 0.34, t, 13.5, color=ROSE_DEEP, bold=True)
    txt(s, x + 0.22, y + 0.56, cw_g - 0.38, 0.92, d, 12, color=TAUPE, spacing=1.25)

# ============================================================
# P3  課程提供材料
# ============================================================
s = slide(); bg(s)
header(s, "MATERIALS  PROVIDED", "課程提供材料")
pgnum(s, 3)
mats = [
    ("考核用大型八角板木器 ×1", "符合條件者可升級為八角折疊桌"),
    ("上課用木器 ×1",          "全程使用，課後帶回"),
    ("施華洛世奇鑽飾套組",      "珠寶膠、沾珠筆等專業工具"),
    ("高級貂毛筆套組",          "橢圓筆、斜筆、勾線筆各一支"),
    ("筆記本與書寫工具",         "課程講義一份"),
]
cw_m = (CW - 0.22) / 2; ch_m = 1.22; gy_m = 0.22
for i, (title, sub) in enumerate(mats):
    r, c = divmod(i, 2)
    x = ML + c * (cw_m + 0.22); y = 1.25 + r * (ch_m + gy_m)
    rect(s, x, y, cw_m, ch_m, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, 0.1, ch_m, GOLD)
    txt(s, x + 0.26, y + 0.12, cw_m - 0.4,  0.48, title, 14,   color=ROSE_DEEP, bold=True)
    txt(s, x + 0.28, y + 0.64, cw_m - 0.42, 0.48, sub,   12.5, color=TAUPE)

# ============================================================
# P4  色彩三大屬性
# ============================================================
s = slide(); bg(s)
header(s, "COLOR  FUNDAMENTALS", "色彩的三大屬性")
pgnum(s, 4)
txt(s, ML, 1.18, CW, 0.36,
    "任何顏色都能用三個維度描述。掌握它們，就能精準調出腦中想像的每一種色。",
    12, color=TAUPE, spacing=1.2)
attrs = [
    ("色相 Hue",    "顏色的「相貌」",
     "紅橙黃綠藍紫——\n區分顏色種類的名稱，\n對應色相環上的位置。",
     [WHEEL[0][1], WHEEL[2][1], WHEEL[4][1], WHEEL[6][1], WHEEL[8][1], WHEEL[10][1]], None),
    ("明度 Value",  "顏色的明暗",
     "加白色→顏色漸亮；\n加黑色→顏色漸暗。\n明度差創造畫面的層次感。",
     [RGBColor(0xFF,0xFF,0xFF), RGBColor(0xCC,0xCC,0xCC), RGBColor(0x99,0x99,0x99),
      RGBColor(0x66,0x66,0x66), RGBColor(0x33,0x33,0x33), RGBColor(0x00,0x00,0x00)], None),
    ("彩度 Chroma", "顏色的鮮濁",
     "加灰或互補色後，\n色彩從鮮豔→混濁，\n整體氛圍更柔和耐看。",
     [RGBColor(0xE8,0x43,0x4B), RGBColor(0xE8,0x9A,0x3A), RGBColor(0xE6,0xDA,0x55),
      RGBColor(0x5A,0xA8,0x5A), RGBColor(0x4A,0x7A,0xC8), RGBColor(0x9A,0x5A,0xB8)],
     [RGBColor(0xBC,0x6A,0x6D), RGBColor(0xBC,0x95,0x65), RGBColor(0xBB,0xB5,0x73),
      RGBColor(0x75,0x9C,0x75), RGBColor(0x6D,0x85,0xAC), RGBColor(0x95,0x75,0xA4)]),
]
cw_a = (CW - 0.3) / 3  # ~2.29"
for i, (t, sub, d, sw1, sw2) in enumerate(attrs):
    x = ML + i * (cw_a + 0.15)
    card_h = 5.5 if sw2 is not None else 5.1
    rect(s, x, 1.62, cw_a, card_h, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x + 0.16, 1.80, cw_a - 0.32, 0.44, t,   14,   color=ROSE_DEEP, bold=True)
    txt(s, x + 0.16, 2.28, cw_a - 0.32, 0.30, sub, 10.5, color=GOLD, bold=True)
    sww = (cw_a - 0.32) / 6
    sw_line = RGBColor(0xBB, 0xBB, 0xBB) if t.startswith("明度") else None
    if sw2 is None:
        for j, c in enumerate(sw1):
            rect(s, x + 0.16 + j * sww, 2.68, sww, 0.40, c, line=sw_line, lw=0.5)
        txt(s, x + 0.16, 3.18, cw_a - 0.32, 3.8, d, 12, color=TAUPE, spacing=1.38)
    else:
        txt(s, x + 0.16, 2.64, cw_a - 0.32, 0.25, "高彩度  鮮豔", 8.5, color=ROSE_DEEP, bold=True)
        for j, c in enumerate(sw1):
            rect(s, x + 0.16 + j * sww, 2.92, sww, 0.38, c)
        txt(s, x + 0.16, 3.34, cw_a - 0.32, 0.25, "低彩度  混濁", 8.5, color=TAUPE, bold=True)
        for j, c in enumerate(sw2):
            rect(s, x + 0.16 + j * sww, 3.62, sww, 0.38, c)
        txt(s, x + 0.16, 4.10, cw_a - 0.32, 3.3, d, 12, color=TAUPE, spacing=1.38)

# ============================================================
# P5  十二色相環
# ============================================================
s = slide(); bg(s)
header(s, "COLOR  WHEEL", "十二色相環")
pgnum(s, 5)
items_cw = [
    ("三原色（原色）",
     "紅・黃・藍——無法由其他顏色調出，是一切色彩的源頭。"),
    ("三間色（二次色）",
     "橙（紅+黃）、綠（黃+藍）、紫（藍+紅）——兩原色等量相混。"),
    ("複色（三次色）",
     "原色與相鄰間色相混，如紅橙、黃綠，使色相環更細緻。"),
    ("冷暖之分",
     "紅橙黃為暖色、藍綠紫為冷色；暖色前進、冷色後退，影響空間感。"),
]
y = 1.25
for t, d in items_cw:
    rect(s, ML, y + 0.08, 0.12, 0.12, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, ML + 0.28, y - 0.02, CW - 0.3, 0.38, t, 14.5, color=ROSE_DEEP, bold=True)
    txt(s, ML + 0.30, y + 0.38, CW - 0.32, 0.6,  d, 12.5, color=TAUPE, spacing=1.2)
    y += 1.1
txt(s, ML, 5.68, CW, 0.36, "十二色相環示意圖", 14, color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
color_wheel(s, PW / 2, 7.88, 1.15)

# ============================================================
# P6  五種配色法
# ============================================================
s = slide(); bg(s)
header(s, "COLOR  SCHEMES", "五種經典配色法")
pgnum(s, 6)
schemes = [
    ("同類配色", "同色相・不同明彩度", [0, 0, 0],  "和諧內斂，統一氛圍",
     "整體感強，適合寧靜冥想主題；以不同明度製造層次，避免過於單調。"),
    ("鄰近配色", "相鄰 2–3 色",       [8, 9, 10], "自然柔和，過渡流暢",
     "視覺舒適，是最常見的配色方式；顏色過渡自然，適合初學者嘗試。"),
    ("互補配色", "色環對角兩色",       [0, 6],     "強烈對比，最具張力",
     "視覺衝擊力強；一色為主、互補色為點綴，避免兩色等量使畫面緊張。"),
    ("分裂互補", "一色＋互補兩側",     [0, 7, 5],  "有對比又不刺眼",
     "保有互補的張力，卻因主色兩側分裂而柔和許多，是最安全的對比配色。"),
    ("三角配色", "等距三色",           [0, 4, 8],  "活潑均衡，色彩豐富",
     "三色均衡搭配，畫面活潑；以一色為主調，其餘二色輔助，保持視覺平衡。"),
]
ch_sch = 1.72; gy_sch = 0.2
for i, (t, sub, idx, note, desc) in enumerate(schemes):
    y = 1.22 + i * (ch_sch + gy_sch)
    rect(s, ML, y, CW, ch_sch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if t == "同類配色":
        cols = [RGBColor(0xF0, 0xC8, 0xCC), RGBColor(0xC9, 0x7B, 0x84), RGBColor(0x8E, 0x44, 0x4D)]
    else:
        cols = [WHEEL[k][1] for k in idx]
    sw_w = 0.48
    for j, col in enumerate(cols):
        rect(s, ML + 0.18 + j * (sw_w + 0.1), y + 0.6, sw_w, sw_w,
             col, line=WHITE, lw=1.2, shape=MSO_SHAPE.OVAL)
    txt(s, ML + 0.22, y + 0.1,  2.6,      0.42, t,    15,   color=ROSE_DEEP, bold=True)
    txt(s, ML + 2.92, y + 0.12, CW - 3.1, 0.36, sub,  10.5, color=TAUPE)
    txt(s, ML + 2.3,  y + 0.55, CW - 2.5, 1.08, desc, 11.5, color=INK, spacing=1.2)
    txt(s, ML + 0.22, y + 1.35, 2.2,      0.3,  note, 11,   color=GOLD, bold=True)

# ============================================================
# P7  實用調色技巧
# ============================================================
s = slide(); bg(s)
header(s, "MIXING  SKILLS", "實用調色技巧")
pgnum(s, 7)
tips = [
    ("提高明度", "加入白色——顏色變亮、變粉嫩，適合花瓣高光與漸層。"),
    ("降低明度", "加入少量黑或深褐——加深陰影，黑色易濁，宜少量多次。"),
    ("降低彩度", "加灰或加一點互補色——讓過於鮮豔的色變柔和耐看。"),
    ("調出高級灰", "互補色相混可中和成有層次的「高級灰」，勝過直接用黑灰。"),
    ("漸層練習",   "同色相由淺到深排出 5 階，是曼陀羅層次感的關鍵基本功。"),
    ("先淺後深",   "由淺色鋪底、再疊深色——好修正，畫面也更通透乾淨。"),
]
ch_tip = 1.35; gy_tip = 0.22
for i, (t, d) in enumerate(tips):
    y = 1.22 + i * (ch_tip + gy_tip)
    rect(s, ML, y, CW, ch_tip, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML, y, 0.1, ch_tip, GOLD)
    txt(s, ML + 0.28, y + 0.14, 2.1,      0.45, t, 15,   color=ROSE_DEEP, bold=True)
    txt(s, ML + 2.45, y + 0.04, CW - 2.6, ch_tip, d, 13, color=TAUPE, spacing=1.1,
        anchor=MSO_ANCHOR.MIDDLE)

# ============================================================
# P8  色彩心理學
# ============================================================
s = slide(); bg(s)
header(s, "COLOR  PSYCHOLOGY", "色彩心理學")
pgnum(s, 8)
txt(s, ML, 1.18, CW, 0.36,
    "顏色會說話。理解色彩傳遞的情緒，讓曼陀羅不只好看，更能呼應內在。",
    12, color=TAUPE, spacing=1.2)
psy = [
    ("紅",  WHEEL[0][1],           "熱情・能量・行動力",
     "激發動能、提振精神。象徵勇氣與強烈的生命力，使人感到興奮與躍躍欲試。"),
    ("橙",  WHEEL[2][1],           "溫暖・喜悅・社交",
     "帶來歡快、開朗的氛圍。有助增進人際溝通，讓人感受到溫度與活力。"),
    ("黃",  WHEEL[4][1],           "陽光・希望・自信",
     "象徵光明與樂觀的心境。激發創意與好奇心，讓人充滿希望感與正能量。"),
    ("綠",  WHEEL[6][1],           "療癒・平衡・安定",
     "自然界最和諧的色彩。緩解壓力、平衡身心，帶來放鬆、平靜的療癒感。"),
    ("藍",  WHEEL[8][1],           "冷靜・信任・沉澱",
     "讓思緒沉澱、回歸理性。傳遞誠信與穩重，適合需要靜心專注的狀態。"),
    ("紫",  WHEEL[10][1],          "靈性・直覺・想像",
     "連結內在深層意識。象徵智慧與神秘，喚起靈感、直覺與靈性感受。"),
    ("粉",  RGBColor(0xE6,0xA9,0xC0), "溫柔・愛・包容",
     "傳遞溫柔與無條件的愛。讓人放下防備、敞開心房，感受被呵護與接納的溫暖。"),
    ("白",  RGBColor(0xF0,0xEC,0xE4), "純淨・開始・留白",
     "象徵清白與全新的起點。在畫面中製造「留白」，讓其他色彩更顯呼吸與空間感。"),
]
cw_p = (CW - 0.22) / 2; ch_p = 1.52; gy_p = 0.18
for i, (t, col, kw, d) in enumerate(psy):
    r, c = divmod(i, 2)
    x = ML + c * (cw_p + 0.22); y = 1.62 + r * (ch_p + gy_p)
    rect(s, x, y, cw_p, ch_p, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, 0.46, ch_p, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + 0.23, y, 0.23, ch_p, col)
    txt(s, x + 0.04, y + (ch_p - 0.42) / 2, 0.38, 0.42, t,  16, color=WHITE,    bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.60, y + 0.10, cw_p - 0.74, 0.36, kw, 11,   color=ROSE_DEEP,   bold=True)
    txt(s, x + 0.60, y + 0.52, cw_p - 0.74, 0.90, d,  10.5, color=TAUPE, spacing=1.22)

# ── 色彩心理學深析資料（P10–P17 用）────────────────────────
_AXES_HO = ["能量 · 活力", "熱情 · 溫度", "平靜 · 穩定", "社交 · 外向", "靈性 · 直覺"]
_psy_full = [
    ("紅", WHEEL[0][1], "RED", ["熱情","能量","行動","勇氣"],
     [95, 98, 15, 80, 30],
     ["行動力強，敢於挑戰、突破現狀", "充滿熱情，能感染身邊每一個人",
      "天生領導氣質，果斷而積極主動", "對生命充滿渴望與旺盛的活力"],
     "紅色刺激腎上腺素分泌，使心跳加速、呼吸加快，帶來興奮、緊迫與高度專注；\n是波長最長、最能凝聚行動能量的顏色。",
     "在心靈彩繪中，紅色點燃內在火種，適合需要勇氣、突破與自我肯定的階段，幫助重拾生命熱度。",
     "以紅作花心或主花瓣，搭配金色輪廓，展現強烈生命力；少量點綴即有畫龍點睛之效。"),
    ("橙", WHEEL[2][1], "ORANGE", ["溫暖","喜悅","社交","創意"],
     [85, 88, 35, 95, 35],
     ["外向開朗，善於人際溝通互動", "感染力強，能帶動整體氛圍",
      "思維活潑，富有創意與巧思", "情感豐沛，自然地溫暖他人"],
     "橙色結合紅的能量與黃的溫暖，刺激食慾、提振情緒，\n帶來歡快輕鬆的心情，對低落、退縮的狀態特別有提振作用。",
     "橙色開啟表達與連結的能量，適合需要打開心房、增進互動與重燃熱情的時刻。",
     "橙色用於中圈花瓣搭配黃色漸層，溫暖明亮；與藍色互補可大幅增強視覺活力。"),
    ("黃", WHEEL[4][1], "YELLOW", ["陽光","智慧","希望","自信"],
     [90, 70, 40, 85, 45],
     ["思維清晰，邏輯與分析力強", "積極向上，充滿好奇心",
      "自信從容，樂於表達自我", "散發正能量，帶動希望感"],
     "黃色是視覺上最明亮的顏色，刺激神經系統、提升思考清晰度與注意力，\n與陽光、希望、輕鬆愉快的感受緊密相連。",
     "黃色是「永遠向陽」的生命力顏色，適合需要信心、樂觀與重新點亮希望的心靈狀態。",
     "黃色與金色同頻，常用於花心發光效果；搭配紫色形成最強力的互補配色。"),
    ("綠", WHEEL[6][1], "GREEN", ["療癒","平衡","安定","自然"],
     [50, 45, 95, 55, 60],
     ["穩重踏實，重視內外和諧", "善解人意，體貼而包容",
      "親近自然，喜愛寧靜的環境", "身心容易取得平衡與安定"],
     "綠色位於色相環中央，是眼睛最不費力辨識的顏色；\n能降低血壓、放鬆眼部肌肉與神經，是最具療癒能量的波長之一。",
     "綠色帶來休息與重整的力量，適合疲憊、需要被安撫與回到平衡的階段。",
     "大面積綠色搭配白色與金色細節，呈現森林療癒感；加入粉紅花朵帶來春天生機。"),
    ("藍", WHEEL[8][1], "BLUE", ["冷靜","信任","深度","專注"],
     [45, 30, 90, 40, 70],
     ["理性分析，思維邏輯清晰", "值得信賴，誠實而穩重",
      "追求深度，不甘流於表面", "靜心沉澱後能爆發創意"],
     "藍色降低心跳速率與血壓，幫助神經進入放鬆狀態；\n研究顯示藍色環境能提升專注、效率與被信任感。",
     "藍色是沉澱心靈最有效的顏色，適合需要冷靜、沉思與回歸理性秩序的時刻。",
     "深淺藍漸層表現寧靜與深度；搭配橙色互補，讓冷靜中帶有活力的點綴。"),
    ("紫", WHEEL[10][1], "PURPLE", ["靈性","直覺","神秘","智慧"],
     [55, 50, 65, 45, 98],
     ["靈性敏感，直覺力豐富", "想像力強，思維獨特深邃",
      "追求生命意義，喜愛自省", "散發神秘魅力，引人入勝"],
     "紫色融合紅的熱情與藍的冷靜，形成獨特而高頻的靈性波長；\n自古與高貴、神秘、智慧相連，能誘發冥想與深層感受。",
     "紫色連結高層意識與靈感泉源，是冥想與心靈藝術中能量最強的顏色。",
     "紫色搭配金色展現高貴靈性感；淺到深紫漸層在花心創造神秘而深邃的效果。"),
    ("粉", RGBColor(0xE6,0xA9,0xC0), "PINK", ["溫柔","愛","包容","呵護"],
     [55, 75, 70, 70, 60],
     ["溫柔細膩，善於關懷他人", "情感豐富，易感受他人情緒",
      "包容體諒，不輕易批判他人", "傳遞溫暖，讓人卸下防備"],
     "粉色是紅色柔化後的版本，保留愛的能量卻去除攻擊性；\n研究顯示粉色環境能平息激動情緒、降低敵意，帶來安全感。",
     "粉色開啟自我疼惜與被愛的能量，適合需要溫柔對待自己、修復情感的階段。",
     "粉色搭配白色與玫瑰金，呈現浪漫溫柔基調；加入少許紫色，可提升靈性層次。"),
    ("白", TAUPE, "WHITE", ["純淨","開始","留白","整合"],
     [40, 35, 92, 45, 80],
     ["追求完美，注重每一個細節", "心靈通透，格局開闊包容",
      "懂得留白，擅於取捨化簡", "有整合能力，化繁為簡"],
     "白色涵蓋所有光的波長，象徵完整與純粹；\n在視覺上製造空間與呼吸感，心理上帶來清晰與重新開始的期許。",
     "白色是「重置」心靈最有效的顏色，適合需要放下、清空與迎接全新開始的時刻。",
     "白色作底色或高光點，讓其他顏色更透亮；中心留白的設計，給觀者呼吸與冥想的空間。"),
]

def _ho_color_page(pg, name, col, en, kws, vals, traits, physio, spirit, tip):
    s = slide(); bg(s)
    header(s, f"COLOR  PSYCHOLOGY  ·  {en}", f"色彩心理學深度解析・{name}色", size=22)
    pgnum(s, pg)
    # Full-width color block
    rect(s, ML, 1.18, CW, 1.12, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, ML + 0.20, 1.22, 1.6, 1.0, name, 52, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, ML + 1.90, 1.28, CW - 2.10, 0.46, en, 18, color=WHITE, bold=True)
    txt(s, ML + 1.90, 1.76, CW - 2.10, 0.40, "  ·  ".join(kws), 12, color=WHITE, bold=True)
    # Left: bar chart  Right: traits
    LW = CW * 0.455   # ≈ 3.26"
    RX = ML + LW + 0.24
    RW = CW - LW - 0.24
    txt(s, ML, 2.44, LW, 0.34, "能量 · 個性指數", 12, color=ROSE_DEEP, bold=True)
    rect(s, ML, 2.78, LW, 0.04, CREAM_DEEP)
    cy0 = 2.88
    for i, (ax, v) in enumerate(zip(_AXES_HO, vals)):
        by = cy0 + i * 0.53
        txt(s, ML, by - 0.04, 1.26, 0.32, ax, 10, color=TAUPE)
        bx = ML + 1.30; bmax = LW - 1.30
        rect(s, bx, by, bmax, 0.26, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, bx, by, max(bmax * v / 100.0, 0.18), 0.26, col,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        txt(s, bx + bmax - 0.54, by - 0.03, 0.50, 0.28, str(v), 10,
            color=TAUPE, bold=True, align=PP_ALIGN.RIGHT)
    txt(s, RX, 2.44, RW, 0.34, "個 性 特 質", 12, color=ROSE_DEEP, bold=True)
    rect(s, RX, 2.78, RW, 0.04, CREAM_DEEP)
    ty = 2.88
    for tr in traits:
        rect(s, RX + 0.04, ty + 0.09, 0.12, 0.12, col, shape=MSO_SHAPE.OVAL)
        txt(s, RX + 0.24, ty, RW - 0.28, 0.34, tr, 11, color=INK)
        ty += 0.44
    # Full-width sections (bars end at cy0 + 5*0.53 = 5.53)
    y = 5.62
    rect(s, ML, y, 0.10, 0.50, col)
    txt(s, ML + 0.24, y + 0.04, CW - 0.30, 0.36, "生理 · 心理影響", 13, color=col, bold=True)
    rect(s, ML, y + 0.54, CW, 0.04, CREAM_DEEP)
    txt(s, ML + 0.06, y + 0.64, CW - 0.12, 1.10, physio, 12, color=TAUPE, spacing=1.30)
    y += 1.84
    rect(s, ML, y, 0.10, 0.50, col)
    txt(s, ML + 0.24, y + 0.04, CW - 0.30, 0.36, "心 靈 意 義", 13, color=col, bold=True)
    rect(s, ML, y + 0.54, CW, 0.04, CREAM_DEEP)
    txt(s, ML + 0.06, y + 0.64, CW - 0.12, 0.84, spirit, 12, color=TAUPE, spacing=1.30)
    y += 1.58
    rect(s, ML, y, CW, 1.30, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, ML + 0.18, y + 0.08, CW - 0.36, 0.30, "✦  曼陀羅應用建議", 11, color=GOLD, bold=True)
    txt(s, ML + 0.18, y + 0.42, CW - 0.36, 0.80, tip, 11, color=TAUPE, spacing=1.2)

# ============================================================
# P9  色彩能量・個性對照
# ============================================================
s = slide(); bg(s)
header(s, "COLOR  ENERGY  &  PERSONALITY", "色彩能量・個性對照表")
pgnum(s, 9)
energy = [
    ("紅", WHEEL[0][1],           "行動 · 能量 · 熱情", "行動力強  充滿幹勁\n勇氣十足  領導氣質"),
    ("橙", WHEEL[2][1],           "溫暖 · 社交 · 喜悅", "外向開朗  善於溝通\n活潑感染  暖化人心"),
    ("黃", WHEEL[4][1],           "陽光 · 智慧 · 希望", "積極樂觀  創意豐富\n自信滿滿  活力四射"),
    ("綠", WHEEL[6][1],           "療癒 · 平衡 · 安定", "穩重踏實  善解人意\n身心放鬆  平靜舒緩"),
    ("藍", WHEEL[8][1],           "冷靜 · 信任 · 深度", "理性分析  邏輯清晰\n值得信賴  追求深度"),
    ("紫", WHEEL[10][1],          "靈性 · 直覺 · 神秘", "富想象力  靈性敏感\n神秘獨特  直覺敏銳"),
    ("粉", RGBColor(0xE6,0xA9,0xC0), "溫柔 · 愛 · 包容",  "溫柔細膩  善於關懷\n情感豐富  包容體貼"),
    ("白", RGBColor(0xCC,0xCC,0xC0), "純淨 · 開始 · 留白", "簡約純粹  追求完美\n心靈通透  全新起點"),
]
cw_e = (CW - 0.22) / 2; ch_e = 2.32; gy_e = 0.18
for i, (name, col, energy_kw, personality) in enumerate(energy):
    r_e, c_e = divmod(i, 2)
    x = ML + c_e * (cw_e + 0.22); y = 1.22 + r_e * (ch_e + gy_e)
    rect(s, x, y, cw_e, ch_e, WHITE, line=col, lw=2.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, 0.78, ch_e, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + 0.39, y, 0.39, ch_e, col)
    txt(s, x + 0.02, y + (ch_e - 0.5) / 2, 0.74, 0.5,
        name, 20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.96, y + 0.16, cw_e - 1.12, 0.46, energy_kw,  11, color=col, bold=True, spacing=1.1)
    rect(s, x + 0.96, y + 0.66, cw_e - 1.22, 0.03, CREAM_DEEP)
    txt(s, x + 0.96, y + 0.76, cw_e - 1.12, ch_e - 0.95, personality, 11, color=TAUPE, spacing=1.35)

# ============================================================
# P10–P17  色彩心理學深度解析（每色一頁）
# ============================================================
for _psy_i, _psy_c in enumerate(_psy_full):
    _ho_color_page(10 + _psy_i, *_psy_c)

# ============================================================
# P18  曼陀羅意涵起源 + 四大結構
# ============================================================
s = slide(); bg(s)
header(s, "MANDALA  ·  MEANING  &  STRUCTURE", "曼陀羅的意涵・起源與結構")
pgnum(s, 18)
txt(s, ML, 1.18, CW, 1.35,
    "「Mandala」源自古印度梵語，原義為「圓」與「中心」，是宇宙、圓滿與內在完整的象徵。"
    "在藏傳佛教中，曼陀羅是宇宙地圖，也是修行者凝神入定的工具。"
    "在現代心靈藝術中，繪製曼陀羅是「專注當下」的靜心歷程——"
    "由圓心出發，一筆一畫安放自己的心，每一件作品都是創作者內在的鏡子。",
    12.5, color=TAUPE, spacing=1.4)
origins = [
    ("圓滿・完整", "從圓心向外，象徵生命由核心向外擴展的力量。"),
    ("秩序・和諧", "等分對稱讓視覺平衡，呼應內在對秩序的渴望。"),
    ("靜心・療癒", "創作過程即冥想，讓雜念沉澱、專注回歸當下。"),
]
y = 2.65
for t, d in origins:
    rect(s, ML, y, 0.1, 0.52, ROSE)
    txt(s, ML + 0.24, y + 0.02, 2.4,       0.34, t, 13,   color=ROSE_DEEP, bold=True)
    txt(s, ML + 0.26, y + 0.36, CW - 0.3,  0.28, d, 11.5, color=TAUPE)
    y += 0.65
txt(s, ML, 4.70, CW, 0.40, "曼陀羅的四大結構", 16, color=ROSE_DEEP, bold=True)
rect(s, ML, 5.08, CW, 0.055, CREAM_DEEP)
struct = [
    ("圓心 Center",   "一切的起點，視覺與能量的核心。所有圖案由此向外生長。"),
    ("放射 Radial",   "圖案由圓心向外發散，如光芒、花瓣，引導視線流動。"),
    ("對稱 Symmetry", "以對稱軸重複，產生穩定有秩序之美；常見 8／12 等分。"),
    ("層次 Layers",   "由內而外一圈圈堆疊，節奏感與豐富度同步提升。"),
]
y = 5.24
for t, d in struct:
    rect(s, ML, y, CW, 1.08, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML, y, 0.1, 1.08, ROSE)
    txt(s, ML + 0.24, y + 0.12, 2.2,      0.46, t, 13.5, color=ROSE_DEEP, bold=True)
    txt(s, ML + 2.55, y + 0.04, CW - 2.72, 1.0, d, 12.5,
        color=TAUPE, anchor=MSO_ANCHOR.MIDDLE, spacing=1.2)
    y += 1.22

# ============================================================
# P19  幾何基礎：對稱與分割（含示意圖）
# ============================================================
s = slide(); bg(s)
header(s, "GEOMETRY  &  SYMMETRY", "幾何基礎：對稱與分割")
pgnum(s, 19)
txt(s, ML, 1.18, CW, 0.54,
    "曼陀羅之美，建立在「等分」之上。先畫出輔助線，在一個扇形設計好圖案，對稱重複即完成。",
    12, color=TAUPE, spacing=1.2)
_divs_ho = [
    (4,  "4 等分",  "簡潔十字構圖"),
    (6,  "6 等分",  "柔和花型"),
    (8,  "8 等分",  "最常用，與八角板契合"),
    (12, "12 等分", "細緻繁複，進階挑戰"),
]
_cw_div4 = (CW - 0.3) / 4   # ≈1.72"
for _i4, (_n4, _lab4, _note4) in enumerate(_divs_ho):
    _x4 = ML + _i4 * (_cw_div4 + 0.1)
    if os.path.exists(_DIV[_n4]):
        pic_cover(s, _DIV[_n4], _x4, 1.84, _cw_div4, _cw_div4, border=False)
    txt(s, _x4, 1.84 + _cw_div4 + 0.06, _cw_div4, 0.34, _lab4, 13,
        color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
    txt(s, _x4, 1.84 + _cw_div4 + 0.40, _cw_div4, 0.28, _note4, 9.5,
        color=TAUPE, align=PP_ALIGN.CENTER)
_bullet_y = 1.84 + _cw_div4 + 0.76
txt(s, ML, _bullet_y, CW, 1.28,
    "・分割數越多，圖案越繁複細緻。初學常用 8 等分，與八角板的造型相呼應。\n"
    "・先以鉛筆淡淡畫出同心圓與放射線當「輔助線」，再沿線設計，完成後可擦除或覆蓋。\n"
    "・只要在一個扇形區塊設計好圖案，再依對稱重複到每一等分，整體就會自然和諧。",
    12.5, color=INK, spacing=1.4)

# ============================================================
# P20  常見圖案元素（含示意圖）
# ============================================================
s = slide(); bg(s)
header(s, "DESIGN  ELEMENTS", "常見圖案元素")
pgnum(s, 20)
txt(s, ML, 1.18, CW, 0.54,
    "曼陀羅由幾種基本「語彙」組合而成。熟悉每種元素的畫法，就能靈活搭配、自由創作。",
    12, color=TAUPE, spacing=1.2)
_elems_ho = [
    ("圓點 Dots",     "由大到小排列出律動感，是最基本也最萬用的元素。",   "el_dots.png"),
    ("花瓣 Petals",   "水滴形、橢圓形組合成花朵，是最常見的主視覺元素。", "el_petals.png"),
    ("葉形 Leaves",   "尖葉、羽葉穿插花朵之間，增添自然生氣與流動感。",   "el_leaves.png"),
    ("水滴 Teardrop", "一頭圓一頭尖，可放射、可串連，變化萬千。",         "el_teardrop.png"),
    ("線條 Lines",    "直線、波浪線、卷草串起各層，引導視線流動。",        "el_lines.png"),
    ("幾何 Shapes",   "三角、菱形、弧形構成骨架，穩定整體結構。",          "el_geometry.png"),
]
_cw_el2 = (CW - 0.22) / 2   # ≈3.475"
_ch_el2 = 1.42; _gy_el2 = 0.22; _img_sq = 1.0
for _i_el, (_t_el, _d_el, _img_el) in enumerate(_elems_ho):
    _row_el, _col_el = divmod(_i_el, 2)
    _x_el = ML + _col_el * (_cw_el2 + 0.22)
    _y_el = 2.00 + _row_el * (_ch_el2 + _gy_el2)
    rect(s, _x_el, _y_el, _cw_el2, _ch_el2, WHITE, line=CREAM_DEEP, lw=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    _img_path = f"{DIAG}/{_img_el}"
    if os.path.exists(_img_path):
        _img_offset_y = _y_el + (_ch_el2 - _img_sq) / 2
        pic_cover(s, _img_path, _x_el + 0.06, _img_offset_y, _img_sq, _img_sq, border=False)
    _tx_el = _x_el + _img_sq + 0.18
    _tw_el = _cw_el2 - _img_sq - 0.24
    txt(s, _tx_el, _y_el + 0.10, _tw_el, 0.42, _t_el, 13.5, color=ROSE_DEEP, bold=True)
    txt(s, _tx_el, _y_el + 0.58, _tw_el, 0.76, _d_el, 11.5, color=TAUPE, spacing=1.2)

# ============================================================
# P21  曼陀羅設計六步驟
# ============================================================
s = slide(); bg(s)
header(s, "STEP  BY  STEP", "曼陀羅設計六步驟")
pgnum(s, 21)
steps_six = [
    ("定圓心",   "在板面正中央定出圓心，這是整個曼陀羅的核心。"),
    ("畫輔助線", "以鉛筆淡淡畫出同心圓與放射對稱線，建立骨架。"),
    ("設計核心", "從圓心開始設計第一層主圖（如花心），奠定主題。"),
    ("由內而外", "一層一層向外擴展，注意每層的大小與間距節奏。"),
    ("對稱重複", "在一個扇形設計好，再對稱複製到每一等分。"),
    ("配色點綴", "依色彩學配色上色，最後以圓點與鑽飾點睛收尾。"),
]
cw_st = (CW - 0.22) / 2; ch_st = 1.45; gy_st = 0.22
for i, (t, d) in enumerate(steps_six):
    r, c = divmod(i, 2)
    x = ML + c * (cw_st + 0.22); y = 1.22 + r * (ch_st + gy_st)
    rect(s, x, y, cw_st, ch_st, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + 0.2, y + 0.34, 0.62, 0.62, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, x + 0.2, y + 0.40, 0.62, 0.52, str(i + 1), 20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 1.02, y + 0.18, cw_st - 1.18, 0.44, t, 14.5, color=ROSE_DEEP, bold=True)
    txt(s, x + 1.04, y + 0.65, cw_st - 1.22, 0.68, d, 12.5, color=TAUPE, spacing=1.1)
# art image row at bottom
arts_row = ["art2.jpg", "art5.jpg", "art6.jpg"]
cw_art = (CW - 0.3) / 3
for i, a in enumerate(arts_row):
    path = f"{L1}/{a}"
    if os.path.exists(path):
        pic_cover(s, path, ML + i * (cw_art + 0.15), 6.28, cw_art, 1.88)

# ============================================================
# P22  課程流程總覽 + 第一天
# ============================================================
s = slide(); bg(s)
header(s, "COURSE  FLOW  ·  DAY 1", "課程流程總覽・第一天")
pgnum(s, 22)
txt(s, ML, 1.18, CW, 0.58,
    "連續兩天密集班——從進階色彩學、設計規劃，到大型八角板創作與貼鑽完成，"
    "最後進入作品解析與教學技巧引導。從「會畫」到「能教」，一次到位。",
    12.5, color=TAUPE, spacing=1.3)
cards = [
    ("22", "小時", "兩日密集實作"),
    ("2",  "天",   "色彩×設計×教學"),
    ("≤10","人",   "小班精緻教學"),
    ("3",  "年",   "免費無限複訓"),
]
cw_c = (CW - 0.3) / 4
for i, (n, u, d) in enumerate(cards):
    x = ML + i * (cw_c + 0.1)
    rect(s, x, 1.88, cw_c, 2.05, WHITE, line=GOLD, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x, 2.18, cw_c, 0.78, n, 34, color=ROSE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x, 3.02, cw_c, 0.36, u, 13, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x, 3.42, cw_c, 0.42, d, 10, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, ML, 4.20, CW, 0.40, "第一天・打底與色彩", 16, color=ROSE_DEEP, bold=True)
rect(s, ML, 4.58, CW, 0.055, CREAM_DEEP)
steps1 = [
    ("進階色彩學", "複習並深化調色、配色原理，導入色彩心理學，建立自己的用色語言。"),
    ("構圖與設計", "規劃曼陀羅對稱結構與層次，於大型八角板上完成精準底稿。"),
    ("分層上色",   "由底色到主視覺逐層堆疊，掌握漸層、暈染與層次處理。"),
]
y = 4.74
for i, (t, d) in enumerate(steps1):
    rect(s, ML, y, 0.6, 0.6, ROSE, shape=MSO_SHAPE.OVAL)
    txt(s, ML, y + 0.06, 0.6, 0.5, str(i + 1), 20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, ML + 0.82, y - 0.02, CW - 0.9, 0.42, t, 15, color=ROSE_DEEP, bold=True)
    txt(s, ML + 0.84, y + 0.42, CW - 0.92, 0.7, d, 12.5, color=TAUPE, spacing=1.15)
    y += 1.28

# ============================================================
# P23  第二天 + 大型八角板設計應用（完整說明）
# ============================================================
s = slide(); bg(s)
header(s, "DAY 2  ·  FINISHING  &  TEACHING", "第二天・完成細節・貼鑽・作品解析・教學引導", size=18)
pgnum(s, 23)
txt(s, ML, 1.18, CW, 0.38, "第二天・完成與教學力", 16, color=ROSE_DEEP, bold=True)
rect(s, ML, 1.54, CW, 0.055, CREAM_DEEP)
steps2_full = [
    ("立體畫法與完成",
     "運用光影原理，讓平面圖案呈現真實的立體質感。同一色系由淺至深分三層堆疊——"
     "最亮處用白色或同色淺化提出高光，最深處以同色加深或疊一點互補色強化陰影，"
     "中間過渡層自然連接。細修所有邊緣線條，確保每一筆都乾淨落位，整件作品達到"
     "細緻、飽滿、富張力的最終呈現水準。"),
    ("施華洛世奇貼鑽",
     "使用透明珠寶膠搭配沾珠筆，逐顆將施華洛世奇水晶鑽飾精準黏貼於圓點中心或"
     "圖案焦點處。鑽飾分佈應有節奏感——不宜全面鋪蓋，以每隔幾個圓點點綴一顆為宜，"
     "讓光澤在畫面中自然閃爍。貼附後輕壓定位，等膠乾燥後作品即提升至展覽販售水準。"),
    ("作品解析",
     "老師帶領逐一解析 15 件精選示範作品，培養你從配色、圖案佈局與整體氛圍解讀"
     "學員內在狀態的洞察力。學習在教學現場如何給予有建設性、有溫度的作品回饋——"
     "這是讓你不只「會畫曼陀羅」，更能「讀懂學員作品」的關鍵教學能力。"),
    ("教學引導訓練",
     "實際演練課堂教學語言：如何清晰示範每個步驟、如何觀察學員的專注狀態與困難點、"
     "如何調整現場節奏與氣氛。採小組輪替方式，讓每位學員同時體驗「學員」與「講師」"
     "兩種角色，在真實演練中建立面對學生時的自信心與教學節奏感。"),
]
y = 1.68
for i, (t, d) in enumerate(steps2_full):
    ch = 2.00
    rect(s, ML, y, CW, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML, y, 0.58, ch, GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML + 0.29, y, 0.29, ch, GOLD)
    txt(s, ML + 0.02, y + (ch - 0.44) / 2, 0.54, 0.44,
        str(i + 1), 20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, ML + 0.74, y + 0.14, CW - 0.88, 0.42, t, 14.5, color=ROSE_DEEP, bold=True)
    txt(s, ML + 0.74, y + 0.58, CW - 0.90, ch - 0.66, d, 11.5, color=TAUPE, spacing=1.22)
    y += ch + 0.13

# ============================================================
# P24  大型八角板・設計應用（獨立頁）
# ============================================================
s = slide(); bg(s)
header(s, "OCTAGON  BOARD  APPLICATION", "大型八角板・設計應用")
pgnum(s, 24)
txt(s, ML, 1.18, CW, 0.90,
    "一級課程的考核作品，是在「大型八角板」上完成的曼陀羅。"
    "八角的外形與 8 等分的放射結構天然契合，讓設計更顯大器而穩定。"
    "以下是在八角板上設計曼陀羅的五大要點：",
    12.5, color=TAUPE, spacing=1.3)
octpts_detail = [
    ("採 8 或 16 等分配置主結構",
     "八角板的造型與 8 等分放射結構天然呼應。初學建議採 8 等分，"
     "進階可用 16 等分製造更細緻繁複的圖案效果。先以鉛筆輕描放射輔助線，"
     "確保每個扇形單位完全等分對稱。"),
    ("主視覺集中圓心，四角與邊緣以小元素呼應",
     "圓心是整件作品最重要的焦點——主花或主圖案放在中心，"
     "由內向外層層擴展。八角板的四個角落與八個邊緣，以較小的圖案元素（如葉形、點點）"
     "做呼應，讓整體設計飽滿而有向心力。"),
    ("善用明度漸層，製造立體感與層次深度",
     "從圓心向外，顏色由深到淺（或由淺到深）自然遞變，"
     "使畫面在視覺上產生立體凸起或向內凹陷的空間感。"
     "每一圈顏色深淺差需清晰可見，避免各圈顏色過於相近而失去層次。"),
    ("鄰近色系鋪底，互補色點睛",
     "主體以鄰近色系（如粉橙黃、藍紫綠）鋪設大面積底色，"
     "再以少量互補色（如主色為藍則點綴橙色）作為圓點或花心，"
     "瞬間突顯中心焦點，活化整體色彩張力而不顯雜亂。"),
    ("最後鑲嵌施華洛世奇鑽飾，大幅提升質感",
     "作品完成、保護漆乾透後，在圓點中心或圖案焦點逐顆貼附施華洛世奇水晶。"
     "鑽飾不宜過密，以每隔幾個圓點點綴為宜，讓光澤自然閃爍。"
     "貼附後作品質感立即提升至展覽及商業販售水準。"),
]
cw_oct = CW; ch_oct = 1.68; gy_oct = 0.18
for i, (t, d) in enumerate(octpts_detail):
    yo = 2.06 + i * (ch_oct + gy_oct)
    rect(s, ML, yo, cw_oct, ch_oct, WHITE, line=CREAM_DEEP, lw=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML, yo, 0.10, ch_oct, GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML + 0.05, yo, 0.05, ch_oct, GOLD)
    rect(s, ML + 0.20, yo + 0.14, 0.56, 0.56, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, ML + 0.20, yo + 0.18, 0.56, 0.48,
        str(i + 1), 14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, ML + 0.94, yo + 0.14, CW - 1.10, 0.38, t, 13.5, color=ROSE_DEEP, bold=True)
    txt(s, ML + 0.94, yo + 0.55, CW - 1.10, ch_oct - 0.62, d, 11.5, color=TAUPE, spacing=1.22)

# ============================================================
# P16–P20  考核必備技巧（每頁兩項，完整說明）
# ============================================================
techniques = [
    ("01", "素材全平面打底上色技法",
     "在木器素材的整個表面（含側面與每個邊緣角落）以平塗方式均勻上色，"
     "確保色層飽和、不露白底或原木色。底色品質直接影響後續圖案的顯色效果，"
     "是整件作品的第一道關鍵工序。",
     "① 選用中型圓頭筆或平頭筆，筆毛飽含顏料但不過量\n"
     "② 以均一力道沿橫向或縱向平行塗佈，避免來回反覆塗抹\n"
     "③ 第一層自然風乾後，再疊第二層補強至顏色飽和無底色透出\n"
     "④ 側面與邊緣角落一同處理，360° 不遺漏任何部位\n"
     "⑤ 整體乾燥後輕輕確認有無漏塗或筆痕，局部補強修平",
     "整體無筆痕、色澤均一飽和；側面與邊緣完整覆蓋，無任何白底或素材原色露出"),
    ("02", "原創設計構圖",
     "自行構思整件曼陀羅作品的圖案設計，展現個人創作力與審美觀。"
     "需體現獨立的設計思維——包含中心主題、圖層安排、元素選擇與整體比例，"
     "而非複製老師的示範範例。",
     "① 確定中心圖案主題（花形、幾何、水滴組合等）與整體風格\n"
     "② 選擇等分數（8 或 16 等分），規劃各圈層的圖案與節奏\n"
     "③ 先在紙上打草稿，確認圓心、各層比例與圖案疏密平衡\n"
     "④ 將構圖轉至木器，用鉛筆輕描同心圓與放射輔助線\n"
     "⑤ 在最小扇形單位設計好後，以對稱重複完成整圈",
     "設計具完整性與個人特色，可見清晰的設計邏輯；非範本複製，展現個人原創構思"),
    ("03", "圖案設計平塗打底技法（位置不限，但一定要有）",
     "在設計好的圖案色塊內——花瓣、幾何形或其他圖案均可——以平塗技法完整填色打底。"
     "作品中需至少出現一處，位置由創作者自行選擇，但要確保技法清晰可辨。",
     "① 選定要進行平塗的圖案色塊（花瓣、菱形、葉形等均可）\n"
     "② 使用小至中型圓頭筆，筆毛飽含顏料但不滴落\n"
     "③ 先描邊界，再從邊緣向內平塗填滿，確保邊界清晰整齊\n"
     "④ 確保內部顏料均勻飽和，無空白或底色透出\n"
     "⑤ 第一層乾透後視需要疊第二層，提升顏色飽和度",
     "圖案內部平塗均勻飽和、邊緣整齊清晰，底色完全遮蓋；作品中明確可見此技法至少一處"),
    ("04", "大的平頭筆圓點技法",
     "使用平頭筆（Flat Brush）的平面端，以輕壓方式製作大圓點。"
     "重點在每個圓點大小一致、邊緣圓潤飽滿，是曼陀羅作品中最醒目的視覺焦點元素之一，"
     "也是考核中最容易被看出是否紮實的基本功。",
     "① 選用與所需圓點大小相符的平頭筆（直徑 0.8–2 cm 均常見）\n"
     "② 筆頭沾取適量顏料，在調色紙上輕刮去多餘顏料確認大小\n"
     "③ 將筆頭平面垂直輕放至表面，輕壓後迅速提起——不推移、不旋轉\n"
     "④ 每次取量一致、按壓力道均一，確保每個圓點大小完全相同\n"
     "⑤ 圓點間距保持均勻，沿同心圓或放射方向有節奏地排列",
     "圓點圓潤飽滿、邊緣清晰，大小高度一致；間距均勻規律，無毛邊、壓痕或大小不均"),
    ("05", "大圓點上方有其他圖案或多色重疊技法",
     "在已完成的大圓點上方疊加其他圖案元素（如小花心、星形或短線），"
     "或在大圓點中心再疊一個不同顏色的小圓點，製造「圓中有圓」的豐富層次效果，"
     "展現色彩深度與設計細緻度。",
     "① 確認大圓點底色已完全乾燥（輕觸不黏手為準）\n"
     "② 選擇疊加圖案：圓珠小點、線條圖案或小型幾何圖形均可\n"
     "③ 疊加色以互補色或高對比色為佳，使其在底色上清晰可辨\n"
     "④ 以圓珠工具或細線筆在大圓點中心精準點按疊加元素\n"
     "⑤ 可在大圓點外圍再加輻射小點或短線，增加整體層次豐富感",
     "大圓點上的疊加圖案居中清晰、與底色對比明顯；展現層次感而非模糊混色或覆蓋失敗"),
    ("06", "圓珠筆由大到小圓點技法",
     "使用多支不同尺寸的圓珠工具（Dotting Tool），由大到小依序排列漸變圓點，"
     "形成視覺上自然收縮的韻律感，常見於放射線兩側、花瓣邊緣或同心圓的過渡帶，"
     "是增加動感與精緻度的重要技法。",
     "① 準備至少三個尺寸的圓珠工具（大、中、小各一支）\n"
     "② 先以最大圓珠工具點出第一顆，再換中號繼續，最後換小號\n"
     "③ 每次點按前確認顏料取量充足，力道垂直均勻，不拖拉\n"
     "④ 漸變圓點沿直線、弧線或放射線延伸，維持間距一致\n"
     "⑤ 一組漸變完成後清潔筆尖再繼續，避免混色或筆尖大小改變",
     "漸變過渡自然、大小比例遞進明顯；間距整齊，整組方向一致；無突然跳號或大小顛倒"),
    ("07", "圓珠筆均勻圓點技法",
     "以單一尺寸的圓珠工具，在整排或整圈中點出大小完全一致、間距規律的均勻圓點陣列，"
     "是曼陀羅中製造整齊節奏感的基礎骨幹技法。"
     "考核中對一致性要求極嚴格，是最能看出平日練習量的技法之一。",
     "① 選定單一尺寸圓珠工具並確保筆尖乾淨圓潤\n"
     "② 每次取量一致（沾顏料後在調色紙輕點一下先確認大小）\n"
     "③ 按壓時力道完全相同、停留時間一致，提起速度一致\n"
     "④ 沿放射線、同心圓或格狀排列，用輔助線確保間距相等\n"
     "⑤ 每完成 5–8 個圓點清潔筆尖一次，保持顏料新鮮穩定",
     "所有圓點大小完全一致，無一大一小情形；間距均勻規律，整排整圈對齊整齊"),
    ("08", "圓珠筆逗點技法",
     "以圓珠工具落筆後迅速向外拉尾，形成「頭圓尾尖」的逗點形狀。"
     "方向性強，常用於放射線兩側、花瓣輪廓或環狀排列，"
     "製造流動飛揚的動態感，是技法中最具生命力的元素之一。",
     "① 顏料取量略多於均勻圓點（需有足夠形成尾巴的顏料量）\n"
     "② 圓珠工具尖端垂直輕觸表面——此為逗點圓頭，力道稍重定住\n"
     "③ 迅速以弧形或直線向外拉尾，力道漸減、收細收尖\n"
     "④ 全排逗點的拉伸方向須完全一致（朝圓心或朝外同向）\n"
     "⑤ 先在練習紙上熟悉「按→拉」連貫動作，再正式上木器",
     "逗點形狀清晰（頭圓尾尖）、頭尾比例自然；全排方向完全一致，長度大致相同"),
    ("09", "長線筆逗點技法",
     "以長線筆（長毛勾線筆）拉出比圓珠逗點更修長優雅的逗點線條，"
     "筆毛的天然彈性讓尾巴自然收細，呈現精緻的書法感流線美，"
     "是展現手部控制力與細膩技巧的最高難度技法之一。",
     "① 使用長線筆，確保筆毛飽含顏料且整齊聚合成細尖\n"
     "② 筆腹輕接表面開始落筆——此為逗點圓頭（稍停一瞬定形）\n"
     "③ 以緩慢均勻的速度向外拉出線條，過程中漸漸提筆減壓\n"
     "④ 線條需連貫流暢，不可斷開、顫抖或出現鋸齒邊緣\n"
     "⑤ 長度通常是圓珠逗點的 3–5 倍，收尾至極細點自然離筆",
     "線條修長流暢不斷、起頭圓潤飽滿、收尾自然細尖；整排長度一致，無抖動或斷線"),
    ("10", "保護漆平塗技法",
     "作品全部完成後，均勻塗抹水性保護漆，使顏料層不脫落、不褪色，"
     "同時提升作品光澤與耐久性。此為最後一道工序，若操作不當容易留下氣泡或刷痕，"
     "將影響作品整體質感，不可輕忽。",
     "① 確認所有顏料層完全乾透——建議靜置至少 24 小時後再上漆\n"
     "② 選用大型平頭刷或專用海綿刷，刷面寬度適合木器大小\n"
     "③ 保護漆倒少量於調色紙，刷子輕蘸後刮去多餘液體\n"
     "④ 以單一方向（橫向或縱向）均勻刷塗全面，不回刷以避免氣泡\n"
     "⑤ 第一層乾透後視需要輕磨砂再上第二層；側面與底部同步完整處理",
     "表面均勻有光澤，無刷痕、氣泡、結塊或白霧未覆蓋區域；側面與底部也完整覆蓋"),
]

def skill_card(s, tech, y_start, card_h=5.04):
    num, title, desc, steps, review = tech
    rect(s, ML, y_start, CW, card_h, WHITE,
         line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML, y_start, 0.10, card_h, ROSE_DEEP,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML + 0.05, y_start, 0.05, card_h, ROSE_DEEP)
    rect(s, ML + 0.20, y_start + 0.15, 0.68, 0.68, ROSE_DEEP, shape=MSO_SHAPE.OVAL)
    txt(s, ML + 0.20, y_start + 0.19, 0.68, 0.58,
        num, 14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, ML + 1.06, y_start + 0.16, CW - 1.22, 0.52,
        title, 14.5, color=ROSE_DEEP, bold=True)
    rect(s, ML + 0.20, y_start + 0.92, CW - 0.38, 0.04, CREAM_DEEP)
    txt(s, ML + 0.22, y_start + 1.02, CW - 0.44, 1.05,
        desc, 12, color=TAUPE, spacing=1.28)
    txt(s, ML + 0.22, y_start + 2.14, 1.5, 0.30,
        "執行步驟", 9.5, color=GOLD, bold=True)
    rect(s, ML + 1.72, y_start + 2.26, CW - 1.90, 0.02, CREAM_DEEP)
    txt(s, ML + 0.22, y_start + 2.50, CW - 0.44, 1.52,
        steps, 11.5, color=INK, spacing=1.22)
    txt(s, ML + 0.22, y_start + 4.08, 1.5, 0.28,
        "審核重點", 9.5, color=GOLD, bold=True)
    rect(s, ML + 1.72, y_start + 4.18, CW - 1.90, 0.02, CREAM_DEEP)
    txt(s, ML + 0.22, y_start + 4.30, CW - 0.44, 0.66,
        review, 11, color=ROSE_DEEP, spacing=1.2)

CARD_H = 5.04; CARD_GAP = 0.16

# P25  考核技法 01–02
s = slide(); bg(s)
header(s, "CERTIFICATION  SKILLS  01–02", "一級講師考核必備技巧")
pgnum(s, 25)
skill_card(s, techniques[0], 1.10, CARD_H)
skill_card(s, techniques[1], 1.10 + CARD_H + CARD_GAP, CARD_H)

# P26–P29  技法03–10
for _order_i, page_i in enumerate([1, 2, 3, 4]):
    s = slide(); bg(s)
    pg_num = 26 + _order_i
    eb_suffix = f"  {page_i * 2 + 1:02d}–{page_i * 2 + 2:02d}"
    header(s, f"CERTIFICATION  SKILLS{eb_suffix}", "一級講師考核必備技巧")
    pgnum(s, pg_num)
    t1 = techniques[page_i * 2]
    t2 = techniques[page_i * 2 + 1]
    skill_card(s, t1, 1.10, CARD_H)
    skill_card(s, t2, 1.10 + CARD_H + CARD_GAP, CARD_H)

# ============================================================
# P30  證書考核 + 結業後支持
# ============================================================
s = slide(); bg(s)
header(s, "CERTIFICATION  &  AFTER  GRADUATION", "證書考核・結業後持續陪伴")
pgnum(s, 30)
txt(s, ML, 1.18, CW, 0.38, "一級講師・證書考核流程", 16, color=ROSE_DEEP, bold=True)
rect(s, ML, 1.54, CW, 0.055, CREAM_DEEP)
reqs = [
    ("完成課程", "完成 22 小時完整課程訓練"),
    ("提交作品", "提交一幅考核作品（大型八角板，需包含 10 大必備技法）"),
    ("通過考核", "作品經評核通過，取得一級講師資格證書"),
]
y_req = 1.68
for i, (t, d) in enumerate(reqs):
    rect(s, ML, y_req, CW, 0.98, WHITE, line=GOLD, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML + 0.2, y_req + 0.26, 0.52, 0.52, ROSE, shape=MSO_SHAPE.OVAL)
    txt(s, ML + 0.2, y_req + 0.31, 0.52, 0.44, str(i + 1), 17, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, ML + 0.95, y_req + 0.14, CW - 1.12, 0.42, t, 15,   color=ROSE_DEEP, bold=True)
    txt(s, ML + 0.95, y_req + 0.56, CW - 1.12, 0.35, d, 11.5, color=TAUPE)
    y_req += 1.12
rect(s, ML, 5.12, CW, 1.35, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s, ML + 0.2, 5.17, CW - 0.4, 1.25,
    "考核細則\n"
    "・上課結束當天起三個月內完成考核作品\n"
    "・照片：素色背景，提供正面、側面、中央側面特寫三個角度\n"
    "・三個月內繳交：NT$3,500　　逾期繳交：NT$4,200",
    12, color=INK, spacing=1.38)
txt(s, ML, 6.66, CW, 0.40, "結業後・持續陪伴", 16, color=ROSE_DEEP, bold=True)
rect(s, ML, 7.04, CW, 0.055, CREAM_DEEP)
bens = [
    "免費教學輔導，陪你站穩講台",
    "以講師價進貨材料，適用教學販售",
    "三年內免學費複訓（僅收場地材料餐費）",
    "教學助教機會，累積實戰經驗",
    "專屬師資群組即時支援",
    "一級升級考核通過，可教授二級講師培訓課程",
    "課程已納入公務人員終身學習認證",
    "年度師生成果展，增加曝光與資歷",
]
cw_bn = (CW - 0.22) / 2
y_bn = 7.22
for i, b in enumerate(bens):
    c_bn = i % 2; r_bn = i // 2
    x = ML + c_bn * (cw_bn + 0.22); y = y_bn + r_bn * 0.92
    rect(s, x, y, cw_bn, 0.78, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + 0.2, y + 0.29, 0.22, 0.22, SAGE, shape=MSO_SHAPE.OVAL)
    txt(s, x + 0.56, y, cw_bn - 0.70, 0.78, b, 11.5, color=INK, anchor=MSO_ANCHOR.MIDDLE)

OUT = "/home/user/mandala/mandala_l1_handout.pptx"
prs.save(OUT)
print(f"Saved  {OUT}")
print(f"Slides {len(prs.slides)}")
