# -*- coding: utf-8 -*-
"""一級講師課程教學過程簡報 — 溫暖療癒柔色系 v2
   結構：封面 → 講師 → 課程重點 → 材料 → Part1 色彩學 → Part2 曼陀羅設計 → Part3 課程流程/照片/成果/報名
"""
import os, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

L1   = "/home/user/mandala/l1"
DIAG = f"{L1}/diagrams"

# ── 溫暖療癒柔色系 ──────────────────────────────────────
CREAM      = RGBColor(0xFB, 0xF6, 0xEF)
CREAM_DEEP = RGBColor(0xF3, 0xE7, 0xD8)
ROSE       = RGBColor(0xC9, 0x7B, 0x84)
ROSE_DEEP  = RGBColor(0xA9, 0x5C, 0x66)
GOLD       = RGBColor(0xC9, 0xA9, 0x6E)
TAUPE      = RGBColor(0x8A, 0x7A, 0x6A)
INK        = RGBColor(0x4A, 0x40, 0x3A)
SAGE       = RGBColor(0x9C, 0xA9, 0x8C)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# ── 12 色相環 ────────────────────────────────────────────
WHEEL = [
    ("紅",   RGBColor(0xE8,0x43,0x4B)),
    ("紅橙", RGBColor(0xE8,0x6A,0x3A)),
    ("橙",   RGBColor(0xE8,0x9A,0x3A)),
    ("黃橙", RGBColor(0xE8,0xC2,0x4A)),
    ("黃",   RGBColor(0xE6,0xDA,0x55)),
    ("黃綠", RGBColor(0xA8,0xC8,0x4A)),
    ("綠",   RGBColor(0x5A,0xA8,0x5A)),
    ("藍綠", RGBColor(0x4A,0xA8,0x9A)),
    ("藍",   RGBColor(0x4A,0x7A,0xC8)),
    ("藍紫", RGBColor(0x6A,0x5A,0xB8)),
    ("紫",   RGBColor(0x9A,0x5A,0xB8)),
    ("紅紫", RGBColor(0xC8,0x4A,0x8A)),
]

EMU  = 914400
SW, SH = 13.333, 7.5
FONT = "微軟正黑體"

prs = Presentation()
prs.slide_width  = Emu(int(SW*EMU))
prs.slide_height = Emu(int(SH*EMU))
BLANK = prs.slide_layouts[6]

# ── 基本繪圖工具 ─────────────────────────────────────────
def slide(): return prs.slides.add_slide(BLANK)
BG_IMG  = f"{L1}/bg_watercolor.jpg"
_HAS_BG = os.path.exists(BG_IMG)
def bg(s, color=CREAM):
    s.background.fill.solid(); s.background.fill.fore_color.rgb = color
    # 統一水彩底圖（僅淺色頁；深色頁保留滿版照片戲劇感）
    if color == CREAM and _HAS_BG:
        p = s.shapes.add_picture(BG_IMG, Inches(0), Inches(0), Inches(SW), Inches(SH))
        p.line.fill.background()
        ov = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SW), Inches(SH))
        ov.fill.solid(); ov.fill.fore_color.rgb = CREAM; ov.line.fill.background()
        ov.shadow.inherit = False
        set_alpha(ov, 80)

def rect(s, x, y, w, h, color, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def set_alpha(shape, alpha_pct):
    sp = shape.fill._xPr.find(qn('a:solidFill'))
    srgb = sp.find(qn('a:srgbClr'))
    a = srgb.makeelement(qn('a:alpha'), {'val': str(int(alpha_pct*1000))})
    srgb.append(a)

def txt(s, x, y, w, h, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, font=FONT, spacing=1.0, italic=False):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = spacing
        r = p.add_run(); r.text = ln
        f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
        f.name = font; f.color.rgb = color
    return tb

def pic_cover(s, path, x, y, w, h, border=True):
    iw, ih = Image.open(path).size
    tw, th = w*EMU, h*EMU; sr, ir = tw/th, iw/ih
    if ir > sr:
        cw = int(ih*sr); off = (iw-cw)//2; crop = (off/iw,0,(off+cw)/iw,1)
    else:
        ch = int(iw/sr); off = (ih-ch)//2; crop = (0,off/ih,1,(off+ch)/ih)
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    p.crop_left, p.crop_top = crop[0], crop[1]
    p.crop_right, p.crop_bottom = 1-crop[2], 1-crop[3]
    if border:
        p.line.color.rgb = WHITE; p.line.width = Pt(2.0)
    else:
        p.line.fill.background()
    return p

def pic_top(s, path, x, y, w, h, border=True):
    """Crop to fit box but anchor to top — keeps heads/tops of artwork visible."""
    iw, ih = Image.open(path).size
    tw, th = w*EMU, h*EMU; sr, ir = tw/th, iw/ih
    if ir > sr:
        cw = int(ih*sr); off = (iw-cw)//2; crop = (off/iw, 0, (off+cw)/iw, 1)
    else:
        ch = int(iw/sr); crop = (0, 0, 1, ch/ih)
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    p.crop_left, p.crop_top = crop[0], crop[1]
    p.crop_right, p.crop_bottom = 1-crop[2], 1-crop[3]
    if border:
        p.line.color.rgb = WHITE; p.line.width = Pt(2.0)
    else:
        p.line.fill.background()
    return p

def pic_contain(s, path, x, y, w, h, border=True):
    """Show full image with no cropping — letterbox with CREAM_DEEP background."""
    iw, ih = Image.open(path).size
    ir = iw/ih; sr = w/h
    if ir > sr:           # image wider: fit width, pad top/bottom
        dw, dh = w, w/ir; dx, dy = x, y + (h-dh)/2
    else:                 # image taller: fit height, pad left/right
        dw, dh = h*ir, h; dx, dy = x + (w-dw)/2, y
    rect(s, x, y, w, h, CREAM_DEEP)
    p = s.shapes.add_picture(path, Inches(dx), Inches(dy), Inches(dw), Inches(dh))
    p.crop_left = p.crop_top = p.crop_right = p.crop_bottom = 0
    p.line.fill.background()
    if border:
        p.line.color.rgb = WHITE; p.line.width = Pt(1.5)
    return p

def eyebrow(s, x, y, text, color=GOLD):
    txt(s, x, y, 8, 0.4, text, 13, color=color, bold=True)
def sec_title(s, x, y, text, color=INK, size=34):
    txt(s, x, y, 11.8, 0.9, text, size, color=color, bold=True)
def gold_line(s, x, y, w=1.1): rect(s, x, y, w, 0.045, GOLD)

def header(s, eb, title, size=32):
    rect(s, 0, 0, SW, 1.5, CREAM_DEEP)
    eyebrow(s, 0.9, 0.48, eb)
    sec_title(s, 0.85, 0.78, title, size=size)

def divider(num_label, title, subtitle, img, alpha=18, title_size=36):
    s = slide(); bg(s, INK)
    pic_cover(s, img, 6.7, 0, SW-6.7, SH, border=False)
    ov = rect(s, 0, 0, 7.3, SH, INK); set_alpha(ov, alpha)
    rect(s, 0.9, 2.45, 0.09, 2.45, GOLD)
    txt(s, 1.25, 2.35, 6.2, 0.6, num_label, 18, color=GOLD, bold=True)
    txt(s, 1.22, 2.92, 6.2, 1.3, title, title_size, color=WHITE, bold=True, spacing=1.0)
    txt(s, 1.25, 4.55, 6.0, 1.1, subtitle, 15.5, color=CREAM, spacing=1.35)
    return s

def swatch_row(s, x, y, colors, sw=0.62, gap=0.0, labels=None, lab_size=9):
    for i,c in enumerate(colors):
        cx = x + i*(sw+gap)
        rect(s, cx, y, sw, sw, c)
        if labels:
            txt(s, cx-0.1, y+sw+0.02, sw+0.2, 0.3, labels[i], lab_size,
                color=TAUPE, align=PP_ALIGN.CENTER)

def color_wheel(s, cx, cy, r, dot=0.62, label=True):
    for i,(name,col) in enumerate(WHEEL):
        ang = math.radians(-90 + i*30)
        px = cx + r*math.cos(ang) - dot/2
        py = cy + r*math.sin(ang) - dot/2
        rect(s, px, py, dot, dot, col, line=WHITE, lw=1.5, shape=MSO_SHAPE.OVAL)
        if label:
            lx = cx + (r+0.55)*math.cos(ang) - 0.35
            ly = cy + (r+0.55)*math.sin(ang) - 0.15
            txt(s, lx, ly, 0.7, 0.3, name, 10, color=TAUPE, align=PP_ALIGN.CENTER)

def _make_div_circle_png(n, path, size=480):
    """Draw a circle divided into n equal slices (like cutting a cake). Saves PNG."""
    from PIL import ImageDraw as _D
    img = Image.new("RGB", (size, size), (0xFB, 0xF6, 0xEF))
    dr = _D.Draw(img)
    cx = cy = size // 2
    r = size // 2 - 22
    dr.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(0xF3, 0xE7, 0xD8), outline=(0xC9, 0xA9, 0x6E), width=7)
    for k in range(n):
        ang = math.radians(k * 360.0 / n - 90)
        ex = int(cx + r * math.cos(ang))
        ey = int(cy + r * math.sin(ang))
        dr.line([cx, cy, ex, ey], fill=(0xA9, 0x5C, 0x66), width=5)
    dr.ellipse([cx-r, cy-r, cx+r, cy+r], fill=None, outline=(0xC9, 0xA9, 0x6E), width=7)
    dr.ellipse([cx-16, cy-16, cx+16, cy+16], fill=(0xC9, 0x7B, 0x84))
    img.save(path)

_DIV = {}
for _n in [4, 6, 8, 12]:
    _DIV[_n] = f"/tmp/_mandala_div_{_n}.png"
    _make_div_circle_png(_n, _DIV[_n])

def photo_wall(title, sub, imgs):
    s = slide(); bg(s)
    eyebrow(s, 0.9, 0.5, "IN  THE  CLASSROOM")
    sec_title(s, 0.85, 0.85, title, size=28)
    txt(s, 0.9, 1.62, 11.5, 0.5, sub, 14, color=TAUPE)
    cw=3.95; ch=4.55; gx=0.3
    x0=(SW-(cw*3+gx*2))/2
    for i,im in enumerate(imgs[:3]):
        pic_cover(s, im, x0+i*(cw+gx), 2.25, cw, ch)
    return s

# ============================================================
# 1. 封面
# ============================================================
s = slide()
s.background.fill.solid(); s.background.fill.fore_color.rgb = CREAM
_p0 = s.shapes.add_picture(f"{L1}/bg_ending.jpg", Inches(0), Inches(0), Inches(SW), Inches(SH))
_p0.line.fill.background()
_ov0 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(8.2), Inches(SH))
_ov0.fill.solid(); _ov0.fill.fore_color.rgb = WHITE
_ov0.line.fill.background(); _ov0.shadow.inherit = False
set_alpha(_ov0, 55)
rect(s, 1.0, 2.35, 0.09, 2.55, GOLD)
eyebrow(s, 1.35, 2.45, "MANDALA  ADVANCED  INSTRUCTOR", color=ROSE_DEEP)
txt(s, 1.32, 2.85, 7.5, 1.4, "曼陀羅心靈彩繪", 50, color=INK, bold=True)
txt(s, 1.32, 3.95, 7.5, 1.0, "一級講師課程", 30, color=ROSE_DEEP, bold=True)
txt(s, 1.35, 5.15, 7.5, 0.6,
    "色彩學　×　曼陀羅設計　×　22 小時兩日密集　×　黃彥蓁老師親授", 15, color=TAUPE)

# ============================================================
# 2. 講師介紹
# ============================================================
s = slide(); bg(s)
rect(s, 0, 0, 5.0, SH, CREAM_DEEP)
pic_cover(s, f"{L1}/teacher_portrait.jpg", 0.7, 0.9, 3.6, 5.7)
eyebrow(s, 5.55, 0.95, "MASTER  PROFILE")
sec_title(s, 5.5, 1.3, "黃彥蓁 老師")
gold_line(s, 5.55, 2.25)
prof = [
    "1989 年開始彩繪教學，35 年色彩與心靈藝術資歷",
    "美國 SDP 彩繪協會 CDA 國際認證藝術家──台灣第二位",
    "2015 受邀赴美國俄亥俄州彩繪大會授課（首位華人講師）",
    "出版彩繪專書六冊＋教學 DVD＋國際合輯，全球發行",
    "中華民國傢飾彩繪協會 創會理事長",
    "彩繪界首創「3 年內無限次免費複訓」制度",
]
y = 2.6
for p in prof:
    rect(s, 5.55, y+0.14, 0.16, 0.16, ROSE, shape=MSO_SHAPE.OVAL)
    txt(s, 5.95, y, 6.9, 0.7, p, 14.5, color=INK, spacing=1.15)
    y += 0.78

# ============================================================
# 3. 課程重點一覽
# ============================================================
s = slide(); bg(s)
header(s, "COURSE  HIGHLIGHTS", "一級講師課程・六大收穫")
goals = [
    ("01", "進階色彩學", "色相、明度、彩度、色彩心理學，從理解到精準運用"),
    ("02", "深化理解與應用", "作品解析力，洞察色彩背後的情感與內在狀態"),
    ("03", "大型木器創作", "在大型八角板木器上完成一件精緻曼陀羅作品"),
    ("04", "高級鑽飾裝飾", "施華洛世奇水晶鑽飾貼附，提升作品至展覽水準"),
    ("05", "立體畫法", "進階筆法與層次技巧，讓圖案呈現真實立體質感"),
    ("06", "教學引導訓練", "實際演練教學流程，從「會畫」躍升為「能教」"),
]
cw=3.8; ch=1.58; gx=0.25; gy=0.25
x0=(SW-(cw*3+gx*2))/2
for i,(n,t,d) in enumerate(goals):
    r,c = divmod(i,3)
    x=x0+c*(cw+gx); y=1.9+r*(ch+gy)
    rect(s, x, y, cw, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, cw, 0.42, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x+0.22, y+0.06, 0.55, 0.35, n, 13, color=ROSE_DEEP, bold=True)
    txt(s, x+0.78, y+0.06, cw-1.0, 0.38, t, 15, color=ROSE_DEEP, bold=True)
    txt(s, x+0.22, y+0.58, cw-0.44, 0.92, d, 12.5, color=TAUPE, spacing=1.25)

# ============================================================
# 4. 課程提供材料
# ============================================================
s = slide(); bg(s)
header(s, "MATERIALS  PROVIDED", "課程提供材料")
mats = [
    ("考核用大型八角板木器 ×1",  "符合條件者可升級為八角折疊桌"),
    ("上課用木器 ×1",           "全程使用，課後帶回"),
    ("施華洛世奇鑽飾套組",       "珠寶膠、沾珠筆等專業工具"),
    ("高級貂毛筆套組",           "橢圓筆、斜筆、勾線筆各一支"),
    ("個人繪畫工具",             "顏料、畫筆、調色紙、調色刀等"),
    ("筆記本與書寫工具",          "課程講義一份"),
]
cw=5.55; ch=1.2; gx=0.5; gy=0.28
x0=(SW-(cw*2+gx))/2
for i,(title,sub) in enumerate(mats):
    r,c=divmod(i,2); x=x0+c*(cw+gx); y=1.9+r*(ch+gy)
    rect(s, x, y, cw, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, 0.12, ch, GOLD)
    txt(s, x+0.35, y+0.14, cw-0.5, 0.48, title, 16, color=ROSE_DEEP, bold=True)
    txt(s, x+0.37, y+0.64, cw-0.55, 0.42, sub, 12.5, color=TAUPE)

# ============================================================
#  Part 1 — 色彩學
# ============================================================
divider("PART 1  ·  COLOR  THEORY", "色彩學基礎",
        "一切創作，從理解色彩開始——\n色相、明度、彩度，是曼陀羅的語言。",
        f"{L1}/c3/470951_0.jpg")

# A1 色彩三屬性  ── 左側說明卡 + 右側真實作品
s = slide(); bg(s)
header(s, "FUNDAMENTALS", "色彩的三大屬性")
txt(s, 0.85, 1.65, 7.0, 0.5,
    "任何顏色都能用三個維度描述。掌握它們，就能精準調出腦中想像的每一種色。",
    13.5, color=TAUPE, spacing=1.2)
# (名稱, 副標, 描述文字, 色階列1, 色階列2_or_None)
attrs = [
    ("色相 Hue", "顏色的「相貌」",
     "紅橙黃綠藍紫——\n區分顏色種類的名稱，\n對應色相環上的位置。",
     [WHEEL[0][1],WHEEL[2][1],WHEEL[4][1],WHEEL[6][1],WHEEL[8][1],WHEEL[10][1]],
     None),
    ("明度 Value", "顏色的明暗",
     "加白色→顏色漸亮；\n加黑色→顏色漸暗。\n明度差創造畫面的層次感。",
     # 純灰階：白 → 黑
     [RGBColor(0xFF,0xFF,0xFF),RGBColor(0xCC,0xCC,0xCC),RGBColor(0x99,0x99,0x99),
      RGBColor(0x66,0x66,0x66),RGBColor(0x33,0x33,0x33),RGBColor(0x00,0x00,0x00)],
     None),
    ("彩度 Chroma", "顏色的鮮濁",
     "加灰或互補色後，\n色彩從鮮豔→混濁，\n整體氛圍更柔和耐看。",
     # 上排：高彩度（鮮豔）
     [RGBColor(0xE8,0x43,0x4B),RGBColor(0xE8,0x9A,0x3A),RGBColor(0xE6,0xDA,0x55),
      RGBColor(0x5A,0xA8,0x5A),RGBColor(0x4A,0x7A,0xC8),RGBColor(0x9A,0x5A,0xB8)],
     # 下排：低彩度（混濁，同色相加灰）
     [RGBColor(0xBC,0x6A,0x6D),RGBColor(0xBC,0x95,0x65),RGBColor(0xBB,0xB5,0x73),
      RGBColor(0x75,0x9C,0x75),RGBColor(0x6D,0x85,0xAC),RGBColor(0x95,0x75,0xA4)]),
]
cw=2.98; x0=0.7
for i,(t,sub,d,sw1,sw2) in enumerate(attrs):
    x = x0 + i*(cw+0.18)
    card_h = 4.85 if sw2 is not None else 4.5
    rect(s, x, 2.4, cw, card_h, WHITE, line=CREAM_DEEP, lw=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x+0.25, 2.63, cw-0.5, 0.5, t, 17, color=ROSE_DEEP, bold=True)
    txt(s, x+0.25, 3.12, cw-0.5, 0.38, sub, 12, color=GOLD, bold=True)
    sww = (cw-0.5)/6
    # 明度用淺灰框線讓白色格可見
    sw_line = RGBColor(0xBB,0xBB,0xBB) if t.startswith("明度") else None
    if sw2 is None:
        for j,c in enumerate(sw1):
            rect(s, x+0.25+j*sww, 3.65, sww, 0.45, c, line=sw_line, lw=0.5)
        txt(s, x+0.25, 4.3, cw-0.5, 2.35, d, 13, color=TAUPE, spacing=1.38)
    else:
        # 雙排色階（彩度專用）
        txt(s, x+0.25, 3.58, cw-0.5, 0.3, "高彩度  鮮豔", 9.5, color=ROSE_DEEP, bold=True)
        for j,c in enumerate(sw1):
            rect(s, x+0.25+j*sww, 3.86, sww, 0.4, c)
        txt(s, x+0.25, 4.3, cw-0.5, 0.3, "低彩度  混濁", 9.5, color=TAUPE, bold=True)
        for j,c in enumerate(sw2):
            rect(s, x+0.25+j*sww, 4.58, sww, 0.4, c)
        txt(s, x+0.25, 5.1, cw-0.5, 1.95, d, 13, color=TAUPE, spacing=1.38)
# 右側作品範例
pic_cover(s, f"{L1}/art1.jpg", 10.0, 1.65, 2.9, 2.55)
pic_cover(s, f"{L1}/art3.jpg", 10.0, 4.35, 2.9, 2.55)
txt(s, 10.0, 4.98, 2.9, 0.6, "▲ 真實作品中\n色相・明度・彩度的層次", 10.5,
    color=TAUPE, align=PP_ALIGN.CENTER, spacing=1.2)

# A2 色相環  ── 左側色環圖 + 右側說明 + 右下角作品
s = slide(); bg(s)
header(s, "COLOR  WHEEL", "十二色相環")
color_wheel(s, 3.5, 4.35, 2.05)
txt(s, 2.75, 3.98, 1.5, 0.6, "色相環", 14, color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
items = [
    ("三原色（原色）",
     "紅・黃・藍——無法由其他顏色調出，是一切色彩的源頭。"),
    ("三間色（二次色）",
     "橙（紅+黃）、綠（黃+藍）、紫（藍+紅）——兩原色等量相混。"),
    ("複色（三次色）",
     "原色與相鄰間色相混，如紅橙、黃綠，使色相環更細緻。"),
    ("冷暖之分",
     "紅橙黃為暖色、藍綠紫為冷色；暖色前進、冷色後退，影響空間感。"),
]
y = 2.0
for t,d in items:
    rect(s, 6.95, y+0.08, 0.14, 0.14, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, 7.28, y-0.02, 5.7, 0.42, t, 15, color=ROSE_DEEP, bold=True)
    txt(s, 7.3,  y+0.42, 5.7, 0.72, d, 12.5, color=TAUPE, spacing=1.2)
    y += 1.22

# A3 五種配色法  ── 上方 5 張卡（含色點） + 下方過程照 + 作品照
s = slide(); bg(s)
header(s, "COLOR  SCHEMES", "五種經典配色法")
schemes = [
    ("同類配色", "同色相・不同明彩度", [0,0,0],  "和諧內斂，統一氛圍"),
    ("鄰近配色", "相鄰 2–3 色",       [8,9,10], "自然柔和，過渡流暢"),
    ("互補配色", "色環對角兩色",       [0,6],    "強烈對比，最具張力"),
    ("分裂互補", "一色＋互補兩側",     [0,7,5],  "有對比又不刺眼"),
    ("三角配色", "等距三色",           [0,4,8],  "活潑均衡，色彩豐富"),
]
cw=2.35; ch=2.2; gx=0.23
x0=(SW-(cw*5+gx*4))/2
for i,(t,sub,idx,note) in enumerate(schemes):
    x=x0+i*(cw+gx); y=1.85
    rect(s, x, y, cw, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x+0.2, y+0.18, cw-0.4, 0.42, t, 15, color=ROSE_DEEP, bold=True)
    txt(s, x+0.2, y+0.6,  cw-0.4, 0.38, sub, 11, color=TAUPE)
    if t=="同類配色":
        cols=[RGBColor(0xF0,0xC8,0xCC),RGBColor(0xC9,0x7B,0x84),RGBColor(0x8E,0x44,0x4D)]
    else:
        cols=[WHEEL[k][1] for k in idx]
    sw_w=0.55
    for j,col in enumerate(cols):
        rect(s, x+0.2+j*(sw_w+0.1), y+1.08, sw_w, sw_w, col,
             line=WHITE, lw=1.5, shape=MSO_SHAPE.OVAL)
    txt(s, x+0.2, y+1.78, cw-0.4, 0.35, note, 11, color=GOLD, bold=True)
# 下方：過程照 + 成品照（展示配色實際效果）
pic_cover(s, f"{L1}/c2/470937_0.jpg", 0.7,  4.35, 3.8, 2.8)
pic_cover(s, f"{L1}/art2.jpg",         4.75, 4.35, 3.8, 2.8)
pic_cover(s, f"{L1}/art6.jpg",         8.8,  4.35, 3.8, 2.8)
txt(s, 0.7, 7.22, 3.8, 0.3, "▲ 鄰近配色上色過程", 10.5, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, 4.75,7.22, 3.8, 0.3, "▲ 互補配色完成作品", 10.5, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, 8.8, 7.22, 3.8, 0.3, "▲ 三角配色完成作品", 10.5, color=TAUPE, align=PP_ALIGN.CENTER)

# A4 調色技巧  ── 左側 6 技巧卡 + 右側 2 張過程照
s = slide(); bg(s)
header(s, "MIXING  SKILLS", "實用調色技巧")
tips = [
    ("提高明度", "加入白色——顏色變亮、變粉嫩，適合花瓣高光與漸層。"),
    ("降低明度", "加入少量黑或深褐——加深陰影，黑色易濁，宜少量多次。"),
    ("降低彩度", "加灰或加一點互補色——讓過於鮮豔的色變柔和耐看。"),
    ("調出高級灰", "互補色相混可中和成有層次的「高級灰」，勝過直接用黑灰。"),
    ("漸層練習", "同色相由淺到深排出 5 階，是曼陀羅層次感的關鍵基本功。"),
    ("先淺後深", "由淺色鋪底、再疊深色——好修正，畫面也更通透乾淨。"),
]
cw=6.3; ch=1.32; gx=0.0; gy=0.24
for i,(t,d) in enumerate(tips):
    y=1.88+i*(ch+gy)
    rect(s, 0.7, y, cw, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, 0.7, y, 0.12, ch, GOLD)
    txt(s, 1.0,  y+0.16, 1.9, 0.45, t, 15.5, color=ROSE_DEEP, bold=True)
    txt(s, 2.95, y+0.04, cw-2.38, ch, d, 12.5, color=TAUPE, spacing=1.1,
        anchor=MSO_ANCHOR.MIDDLE)
# 右側：實際調色過程照
pic_cover(s, f"{L1}/c1/proc1.jpg",     7.35, 1.88, 2.8, 5.38)
pic_cover(s, f"{L1}/c2/470938_0.jpg",  10.35, 1.88, 2.7, 5.38)
txt(s, 7.35, 7.32, 5.7, 0.3, "▲ 學員實際調色・分層上色過程", 10.5,
    color=TAUPE, align=PP_ALIGN.CENTER)

# A5 色彩心理學  ── 8 色詳細說明卡（2×4） + 下方作品
s = slide(); bg(s)
header(s, "COLOR  PSYCHOLOGY", "色彩心理學")
txt(s, 0.85, 1.65, 11.6, 0.42,
    "顏色會說話。理解色彩傳遞的情緒，讓曼陀羅不只好看，更能呼應內在。",
    13.5, color=TAUPE, spacing=1.2)
psy = [
    ("紅", WHEEL[0][1], "熱情・能量・行動力",
     "激發動能、提振精神。\n象徵勇氣與強烈的生命力，\n使人感到興奮與躍躍欲試。"),
    ("橙", WHEEL[2][1], "溫暖・喜悅・社交",
     "帶來歡快、開朗的氛圍。\n有助增進人際溝通，\n讓人感受到溫度與活力。"),
    ("黃", WHEEL[4][1], "陽光・希望・自信",
     "象徵光明與樂觀的心境。\n激發創意與好奇心，\n讓人充滿希望感與正能量。"),
    ("綠", WHEEL[6][1], "療癒・平衡・安定",
     "自然界最和諧的色彩。\n緩解壓力、平衡身心，\n帶來放鬆、平靜的療癒感。"),
    ("藍", WHEEL[8][1], "冷靜・信任・沉澱",
     "讓思緒沉澱、回歸理性。\n傳遞誠信與穩重，\n適合需要靜心專注的狀態。"),
    ("紫", WHEEL[10][1], "靈性・直覺・想像",
     "連結內在深層意識。\n象徵智慧與神秘，\n喚起靈感、直覺與靈性感受。"),
    ("粉", RGBColor(0xE6,0xA9,0xC0), "溫柔・愛・包容",
     "傳遞溫柔與無條件的愛。\n讓人放下防備、敞開心房，\n感受被呵護與接納的溫暖。"),
    ("白", RGBColor(0xF0,0xEC,0xE4), "純淨・開始・留白",
     "象徵清白與全新的起點。\n在畫面中製造「留白」，\n讓其他色彩更顯呼吸與空間感。"),
]
cw=3.0; ch=1.75; gx=0.2; gy=0.18
x0=(SW-(cw*4+gx*3))/2
for i,(t,col,kw,d) in enumerate(psy):
    r,c=divmod(i,4); x=x0+c*(cw+gx); y=2.08+r*(ch+gy)
    rect(s, x, y, cw, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # 左側色塊
    rect(s, x, y, 0.48, ch, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x+0.24, y, 0.24, ch, col)
    # 關鍵詞
    txt(s, x+0.62, y+0.1, cw-0.75, 0.38, f"{t}  {kw}", 12, color=ROSE_DEEP, bold=True)
    # 詳細說明
    txt(s, x+0.62, y+0.52, cw-0.75, 1.12, d, 11, color=TAUPE, spacing=1.25)
# 下方作品（卡片兩排結束於 y≈5.76，照片從 5.88 開始不重疊）
pic_cover(s, f"{L1}/art4.jpg", 0.7,  5.88, 3.8, 1.38)
pic_cover(s, f"{L1}/art5.jpg", 4.75, 5.88, 3.8, 1.38)
pic_cover(s, f"{L1}/art6.jpg", 8.8,  5.88, 3.8, 1.38)
txt(s, 0.7, 7.3, 3.8, 0.22, "▲ 暖色系：熱情與能量", 9.5, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, 4.75,7.3, 3.8, 0.22, "▲ 冷色系：靈性與沉澱", 9.5, color=TAUPE, align=PP_ALIGN.CENTER)
txt(s, 8.8, 7.3, 3.8, 0.22, "▲ 多色：活力與均衡",   9.5, color=TAUPE, align=PP_ALIGN.CENTER)

# A5b 色彩能量・個性對照表
s = slide(); bg(s)
header(s, "COLOR  ENERGY  &  PERSONALITY", "色彩能量・個性對照表")
energy_chart = [
    ("紅", WHEEL[0][1],  "行動 · 能量 · 熱情",
     "行動力強  充滿幹勁\n勇氣十足  領導氣質"),
    ("橙", WHEEL[2][1],  "溫暖 · 社交 · 喜悅",
     "外向開朗  善於溝通\n活潑感染  暖化人心"),
    ("黃", WHEEL[4][1],  "陽光 · 智慧 · 希望",
     "積極樂觀  創意豐富\n自信滿滿  活力四射"),
    ("綠", WHEEL[6][1],  "療癒 · 平衡 · 安定",
     "穩重踏實  善解人意\n身心放鬆  平靜舒緩"),
    ("藍", WHEEL[8][1],  "冷靜 · 信任 · 深度",
     "理性分析  邏輯清晰\n值得信賴  追求深度"),
    ("紫", WHEEL[10][1], "靈性 · 直覺 · 神秘",
     "富想象力  靈性敏感\n神秘獨特  直覺敏銳"),
    ("粉", RGBColor(0xE6,0xA9,0xC0), "溫柔 · 愛 · 包容",
     "溫柔細膩  善於關懷\n情感豐富  包容體貼"),
    ("白", RGBColor(0xCC,0xCC,0xC0), "純淨 · 開始 · 留白",
     "簡約純粹  追求完美\n心靈通透  全新起點"),
]
cw_e=2.95; ch_e=2.55; gx_e=0.12; gy_e=0.18
x0_e=(SW-(cw_e*4+gx_e*3))/2
for i,(name,col,energy_kw,personality) in enumerate(energy_chart):
    r_e,c_e=divmod(i,4)
    x=x0_e+c_e*(cw_e+gx_e); y=1.72+r_e*(ch_e+gy_e)
    rect(s,x,y,cw_e,ch_e,WHITE,line=col,lw=2.0,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # 左側彩色條
    rect(s,x,y,0.85,ch_e,col,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s,x+0.43,y,0.42,ch_e,col)
    txt(s,x+0.02,y+(ch_e-0.55)/2,0.81,0.55,name,22,color=WHITE,bold=True,align=PP_ALIGN.CENTER)
    # 能量關鍵詞
    txt(s,x+1.0,y+0.18,cw_e-1.15,0.52,energy_kw,12,color=col,bold=True,spacing=1.1)
    rect(s,x+1.0,y+0.73,cw_e-1.25,0.03,CREAM_DEEP)
    # 個性描述
    txt(s,x+1.0,y+0.83,cw_e-1.15,ch_e-1.0,personality,11.5,color=TAUPE,spacing=1.35)

# A5c~ 每個顏色一張・色彩心理學深度解析（含能量・個性指數圖表）
# 五大共同指標，讓每個顏色都能橫向對照、活潑呈現
_AXES = ["能量 · 活力", "熱情 · 溫度", "平靜 · 穩定", "社交 · 外向", "靈性 · 直覺"]
# (名稱, 色, 英文, 關鍵詞×4, 五指標分數, 個性特質×4, 生理心理影響, 心靈意義, 曼陀羅應用)
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

def _color_slide(name, col, en, kws, vals, traits, physio, spirit, tip):
    s = slide(); bg(s)
    header(s, f"COLOR  PSYCHOLOGY  ·  {en}", f"色彩心理學深度解析・{name}色", size=28)
    # ── 左欄：大色卡 + 能量個性指數圖表 ──
    LX, LW = 0.55, 5.45
    rect(s, LX, 1.72, LW, 2.05, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, LX+0.32, 1.86, 2.0, 1.4, name, 60, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, LX+2.05, 2.18, LW-2.25, 0.5, en, 19, color=WHITE, bold=True)
    txt(s, LX+2.05, 2.74, LW-2.25, 0.85, "  ·  ".join(kws), 14, color=WHITE, bold=True, spacing=1.25)
    # 能量・個性指數（橫條圖，活潑生動）
    txt(s, LX, 3.88, LW, 0.34, "能量 · 個性指數", 13, color=ROSE_DEEP, bold=True)
    cy0 = 4.32
    for i,(ax,v) in enumerate(zip(_AXES, vals)):
        by = cy0 + i*0.57
        txt(s, LX, by-0.05, 1.7, 0.34, ax, 11, color=INK)
        bx = LX+1.75; bmax = LW-1.75
        rect(s, bx, by, bmax, 0.27, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, bx, by, max(bmax*v/100.0, 0.3), 0.27, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        txt(s, bx+bmax-0.62, by-0.04, 0.58, 0.3, str(v), 10.5, color=TAUPE, bold=True, align=PP_ALIGN.RIGHT)
    # ── 右欄：個性特質 / 生理心理 / 心靈意義 / 曼陀羅應用 ──
    RX, RW = 6.5, 6.3
    y = 1.78
    rect(s, RX, y+0.02, 0.13, 0.46, col)
    txt(s, RX+0.28, y, RW-0.3, 0.4, "個 性 特 質", 14.5, color=col, bold=True)
    y += 0.52
    for tr in traits:
        rect(s, RX+0.06, y+0.1, 0.13, 0.13, col, shape=MSO_SHAPE.OVAL)
        txt(s, RX+0.34, y-0.02, RW-0.4, 0.36, tr, 12.5, color=INK)
        y += 0.41
    y += 0.10
    rect(s, RX, y+0.02, 0.13, 0.46, col)
    txt(s, RX+0.28, y, RW-0.3, 0.4, "生 理 · 心 理 影 響", 14.5, color=col, bold=True)
    txt(s, RX+0.06, y+0.46, RW-0.1, 0.85, physio, 12, color=TAUPE, spacing=1.28)
    y += 1.42
    rect(s, RX, y+0.02, 0.13, 0.46, col)
    txt(s, RX+0.28, y, RW-0.3, 0.4, "心 靈 意 義", 14.5, color=col, bold=True)
    txt(s, RX+0.06, y+0.46, RW-0.1, 0.7, spirit, 12, color=TAUPE, spacing=1.28)
    y += 1.18
    rect(s, RX, y, RW, 0.78, CREAM_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, RX+0.22, y+0.09, 3.5, 0.3, "✦  曼陀羅應用建議", 11, color=GOLD, bold=True)
    txt(s, RX+0.22, y+0.42, RW-0.44, 0.34, tip, 11, color=TAUPE)

for _c in _psy_full:
    _color_slide(*_c)

# ============================================================
#  Part 2 — 曼陀羅圖案設計
# ============================================================
divider("PART 2  ·  MANDALA  DESIGN", "曼陀羅圖案設計",
        "Mandala，梵語意為「圓」。\n從一個圓心，向外綻放出秩序與美。",
        f"{L1}/c3/470950_0.jpg")

# B1a 意涵與起源  ── 文字 + 作品
s = slide(); bg(s)
header(s, "MEANING  &  ORIGIN", "曼陀羅的意涵與起源")
txt(s, 0.85, 1.65, 6.8, 2.9,
    "「Mandala」源自古印度梵語，原義為「圓」與「中心」，"
    "是宇宙、圓滿與內在完整的象徵。\n\n"
    "在藏傳佛教中，曼陀羅是宇宙地圖，也是修行者凝神入定的工具。"
    "沙畫曼陀羅完成後即行抹去，象徵無常與當下。\n\n"
    "在現代心靈藝術中，繪製曼陀羅是「專注當下」的靜心歷程——"
    "由圓心出發，一筆一畫安放自己的心，每一件作品都是創作者內在的鏡子。",
    14, color=TAUPE, spacing=1.45)
origins = [
    ("圓滿・完整", "從圓心向外，象徵生命由核心向外擴展的力量。"),
    ("秩序・和諧", "等分對稱讓視覺平衡，呼應內在對秩序的渴望。"),
    ("靜心・療癒", "創作過程即冥想，讓雜念沉澱、專注回歸當下。"),
]
y=4.7
for t,d in origins:
    rect(s, 0.85, y, 0.12, 0.56, ROSE)
    txt(s, 1.18, y+0.02, 2.8, 0.36, t, 14.5, color=ROSE_DEEP, bold=True)
    txt(s, 1.18, y+0.38, 5.4, 0.32, d, 12, color=TAUPE)
    y += 0.7
pic_cover(s, f"{L1}/art4.jpg", 7.9, 1.65, 5.0, 5.5)

# B1b 曼陀羅的結構  ── 4 大構成 + 結構示意圖 + 作品
s = slide(); bg(s)
header(s, "STRUCTURE", "曼陀羅的四大結構")
struct = [
    ("圓心 Center", "一切的起點，視覺與能量的核心。所有圖案由此向外生長。"),
    ("放射 Radial",  "圖案由圓心向外發散，如光芒、花瓣，引導視線流動。"),
    ("對稱 Symmetry","以對稱軸重複，產生穩定有秩序之美；常見 8／12 等分。"),
    ("層次 Layers",  "由內而外一圈圈堆疊，節奏感與豐富度同步提升。"),
]
y=2.0
for t,d in struct:
    rect(s, 0.75, y, 5.6, 1.08, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, 0.75, y, 0.12, 1.08, ROSE)
    txt(s, 1.1,  y+0.12, 2.0, 0.48, t, 14.5, color=ROSE_DEEP, bold=True)
    txt(s, 3.18, y+0.04, 3.1, 1.0, d, 12.5, color=TAUPE, anchor=MSO_ANCHOR.MIDDLE, spacing=1.2)
    y += 1.22
# 右側：程式畫出的結構示意圖＋作品
pic_cover(s, f"{DIAG}/structure.png", 6.7, 1.7, 3.0, 3.0, border=False)
pic_cover(s, f"{L1}/art6.jpg",        9.85, 1.7, 3.1, 5.45)
txt(s, 6.7, 4.76, 3.0, 0.38, "▲ 圓心・放射・對稱・層次", 10, color=TAUPE, align=PP_ALIGN.CENTER)

# B2 幾何基礎：對稱與分割
s = slide(); bg(s)
header(s, "GEOMETRY", "幾何基礎：對稱與分割")
txt(s, 0.85, 1.65, 11.6, 0.5,
    "曼陀羅之美，建立在「等分」之上。將圓平均分割，畫出對稱輔助線，圖案自然工整。",
    14, color=TAUPE, spacing=1.2)
divs = [(4,"4 等分"),(6,"6 等分"),(8,"8 等分"),(12,"12 等分")]
x0=2.1; gap=2.95; r_img=1.05
for i,(n,lab) in enumerate(divs):
    cx = x0 + i*gap
    pic_cover(s, _DIV[n], cx-r_img, 2.5, r_img*2, r_img*2, border=False)
    txt(s, cx-1.0, 4.68, 2.0, 0.42, lab, 13.5, color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
txt(s, 0.85, 5.45, 11.6, 1.75,
    "・分割數越多，圖案越繁複細緻。初學常用 8 等分，與八角板的造型相呼應。\n"
    "・先以鉛筆淡淡畫出同心圓與放射線當「輔助線」，再沿線設計，完成後可擦除或覆蓋。\n"
    "・只要在一個扇形區塊設計好圖案，再依對稱重複到每一等分，整體就會自然和諧。",
    13.5, color=INK, spacing=1.4)

# B3 常見圖案元素  ── 每項左側有程式圖示意
s = slide(); bg(s)
header(s, "DESIGN  ELEMENTS", "常見圖案元素")
txt(s, 0.75, 1.62, 11.8, 0.45,
    "曼陀羅由幾種基本「語彙」組合而成。熟悉每種元素的畫法，就能靈活搭配、自由創作。",
    13.5, color=TAUPE)
elems = [
    ("圓點 Dots",     "由大到小排列出律動感，是最基本也最萬用的元素。",   "el_dots.png"),
    ("花瓣 Petals",   "水滴形、橢圓形組合成花朵，是最常見的主視覺元素。", "el_petals.png"),
    ("葉形 Leaves",   "尖葉、羽葉穿插花朵之間，增添自然生氣與流動感。",   "el_leaves.png"),
    ("水滴 Teardrop", "一頭圓一頭尖，可放射、可串連，變化萬千。",         "el_teardrop.png"),
    ("線條 Lines",    "直線、波浪線、卷草串起各層，引導視線流動。",        "el_lines.png"),
    ("幾何 Shapes",   "三角、菱形、弧形構成骨架，穩定整體結構。",          "el_geometry.png"),
]
cw_img=1.5; cw_txt=4.45; ch=1.38; gy=0.22; col_gap=0.55
x_left=0.65; x_right=x_left+(cw_img+cw_txt+col_gap)+0.55
for i,(t,d,img) in enumerate(elems):
    col=i%2; row=i//2
    xi = x_left if col==0 else x_right
    y = 2.2 + row*(ch+gy)
    # 圖示
    pic_cover(s, f"{DIAG}/{img}", xi, y, cw_img, ch, border=False)
    # 文字
    rect(s, xi+cw_img+0.1, y, cw_txt, ch, WHITE, line=CREAM_DEEP, lw=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, xi+cw_img+0.32, y+0.12, cw_txt-0.45, 0.45, t, 15, color=ROSE_DEEP, bold=True)
    txt(s, xi+cw_img+0.32, y+0.6,  cw_txt-0.45, 0.68, d, 12, color=TAUPE, spacing=1.2)

# B4 設計步驟
s = slide(); bg(s)
header(s, "STEP  BY  STEP", "曼陀羅設計六步驟")
steps = [
    ("定圓心","在板面正中央定出圓心，這是整個曼陀羅的核心。"),
    ("畫輔助線","以鉛筆淡淡畫出同心圓與放射對稱線，建立骨架。"),
    ("設計核心","從圓心開始設計第一層主圖（如花心），奠定主題。"),
    ("由內而外","一層一層向外擴展，注意每層的大小與間距節奏。"),
    ("對稱重複","在一個扇形設計好，再對稱複製到每一等分。"),
    ("配色點綴","依色彩學配色上色，最後以圓點與鑽飾點睛收尾。"),
]
cw=5.7; ch=1.35; gx=0.4; gy=0.25
x0=(SW-(cw*2+gx))/2
for i,(t,d) in enumerate(steps):
    r,c = divmod(i,2)
    x=x0+c*(cw+gx); y=1.95+r*(ch+gy)
    rect(s, x, y, cw, ch, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x+0.2, y+0.34, 0.66, 0.66, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, x+0.2, y+0.4, 0.66, 0.55, str(i+1), 22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x+1.05, y+0.2, cw-1.2, 0.45, t, 16, color=ROSE_DEEP, bold=True)
    txt(s, x+1.07, y+0.68, cw-1.25, 0.6, d, 12.5, color=TAUPE, spacing=1.1)

# ============================================================
#  Part 3 — 課程流程
# ============================================================
divider("PART 3  ·  COURSE  FLOW", "課程流程總覽",
        "兩天，把色彩、設計與教學力一次練成。",
        f"{L1}/c2/470937_0.jpg")

# C0 課程總覽數字
s = slide(); bg(s)
header(s, "COURSE  OVERVIEW", "課程總覽")
txt(s, 0.85, 1.85, 11.6, 1.0,
    "連續兩天密集班——從進階色彩學、設計規劃，到大型八角板創作與貼鑽完成，"
    "最後進入作品解析與教學技巧引導。從「會畫」到「能教」，一次到位。",
    15.5, color=TAUPE, spacing=1.3)
cards = [
    ("22","小時","兩日密集實作"),
    ("2","天","色彩×設計×教學"),
    ("≤10","人","小班精緻教學"),
    ("3","年","免費無限複訓"),
]
cw,gap=2.75,0.35; x0=(SW-(cw*4+gap*3))/2
for i,(n,u,d) in enumerate(cards):
    x=x0+i*(cw+gap)
    rect(s, x, 3.45, cw, 2.6, WHITE, line=GOLD, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, x, 3.8,  cw, 1.0, n, 46, color=ROSE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x, 4.88, cw, 0.4, u, 16, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x, 5.35, cw, 0.5, d, 14, color=TAUPE, align=PP_ALIGN.CENTER)

# Day 1 process
s = slide(); bg(s)
header(s, "DAY 1  ·  PROCESS", "第一天：打底與色彩", size=28)
steps1 = [
    ("進階色彩學","複習並深化調色、配色原理，導入色彩心理學，建立自己的用色語言。"),
    ("構圖與設計","規劃曼陀羅對稱結構與層次，於大型八角板上完成精準底稿。"),
    ("分層上色",  "由底色到主視覺逐層堆疊，掌握漸層、暈染與層次處理。"),
]
y=2.0
for i,(t,d) in enumerate(steps1):
    rect(s, 0.9, y, 0.7, 0.7, ROSE, shape=MSO_SHAPE.OVAL)
    txt(s, 0.9, y+0.05, 0.7, 0.6, str(i+1), 24, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, 1.8, y-0.02, 4.8, 0.5, t, 19, color=ROSE_DEEP, bold=True)
    txt(s, 1.8, y+0.5, 5.0, 0.9, d, 14, color=TAUPE, spacing=1.15)
    y += 1.5
pic_cover(s, f"{L1}/c2/470935_0.jpg", 7.0, 2.0, 2.95, 4.0)
pic_cover(s, f"{L1}/c2/470938_0.jpg", 10.1, 2.0, 2.95, 4.0)

# Day 2 divider + process（全版照片設計）
s = slide(); bg(s, INK)
pic_cover(s, f"{L1}/c3/group3.jpg", 0, 0, SW, SH, border=False)
ov = rect(s, 0, 0, SW, SH, INK); set_alpha(ov, 62)
ov2 = rect(s, 0, 0, 9.0, SH, INK); set_alpha(ov2, 20)
rect(s, 0.9, 2.3, 0.09, 2.65, GOLD)
txt(s, 1.25, 2.3,  8.0, 0.62, "DAY 2", 20, color=GOLD, bold=True)
txt(s, 1.22, 2.92, 8.0, 1.3,  "完成・貼鑽・解析・教學", 36, color=WHITE, bold=True, spacing=1.0)
txt(s, 1.25, 4.52, 7.5, 1.2,
    "作品完成  ×  施華洛世奇貼鑽\n15 件作品解析  ×  教學引導訓練",
    15.5, color=CREAM, spacing=1.35)
s = slide(); bg(s)
header(s, "DAY 2  ·  PROCESS", "第二天：完成與教學力", size=28)
steps2 = [
    ("立體畫法與完成","以立體畫法強化層次張力，細修細節，邁向完成。"),
    ("施華洛世奇貼鑽","以珠寶膠、沾珠筆鑲嵌鑽飾，質感提升至參展水準。"),
    ("作品解析",      "解析 15 件作品範例，培養對色彩與內在狀態的洞察力。"),
    ("教學引導訓練",  "團體創作演練與教學引導，為開班授課做足準備。"),
]
y=2.0
for i,(t,d) in enumerate(steps2):
    rect(s, 0.9, y, 0.6, 0.6, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, 0.9, y+0.03, 0.6, 0.5, str(i+1), 20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, 1.7, y-0.02, 5.0, 0.5, t, 18, color=ROSE_DEEP, bold=True)
    txt(s, 1.7, y+0.45, 5.1, 0.7, d, 13.5, color=TAUPE, spacing=1.1)
    y += 1.18
pic_cover(s, f"{L1}/c3/470948_0.jpg", 7.0, 2.0, 2.95, 4.0)
pic_cover(s, f"{L1}/c3/470947_0.jpg", 10.1, 2.0, 2.95, 4.0)

# ============================================================
#  Part D — 創作現場照片
# ============================================================
photo_wall("創作現場・第一屆",
           "團體共學的氛圍裡，每一筆都是與自己對話的過程。",
           [f"{L1}/c1/proc1.jpg", f"{L1}/c1/proc2.jpg", f"{L1}/c1/group1.jpg"])
photo_wall("創作現場・第二屆",
           "從打底、上色到細節堆疊——專注，是最美的姿態。",
           [f"{L1}/c2/470935_0.jpg", f"{L1}/c2/470939_0.jpg", f"{L1}/c2/470943_0.jpg"])
photo_wall("創作現場・第三屆",
           "大型八角板上的曼陀羅，在指尖與色彩間慢慢綻放。",
           [f"{L1}/c3/470944_0.jpg", f"{L1}/c3/470945_0.jpg", f"{L1}/c3/470946_0.jpg"])

# ============================================================
#  Part E — 成果與證書
# ============================================================
s = slide(); bg(s)
header(s, "CERTIFICATION", "一級講師・證書考核")
reqs = [
    ("完成課程", "完成 22 小時完整課程訓練"),
    ("提交作品", "提交一幅曼陀羅作品，展示課程所學技法"),
    ("通過考核", "作品經評核通過，取得一級講師資格"),
]
y=2.1
for i,(t,d) in enumerate(reqs):
    rect(s, 0.9, y, 5.6, 1.25, WHITE, line=GOLD, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, 1.15, y+0.32, 0.6, 0.6, ROSE, shape=MSO_SHAPE.OVAL)
    txt(s, 1.15, y+0.37, 0.6, 0.5, str(i+1), 22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    txt(s, 1.95, y+0.2, 4.3, 0.5, t, 18, color=ROSE_DEEP, bold=True)
    txt(s, 1.95, y+0.68, 4.4, 0.5, d, 12.5, color=TAUPE)
    y += 1.5
pic_cover(s, f"{L1}/c3/470949_0.jpg", 7.1, 2.1, 5.3, 4.3)

# B5 八角板設計應用（移至考核說明之後）
s = slide(); bg(s)
header(s, "APPLICATION", "大型八角板・設計應用")
txt(s, 0.85, 1.7, 6.0, 1.4,
    "一級課程的考核作品，是在「大型八角板」上完成的曼陀羅。"
    "八角的外形與 8 等分的放射結構天然契合，讓設計更顯大器。",
    14.5, color=TAUPE, spacing=1.35)
pts = [
    "依八角造型，採 8 或 16 等分配置主結構",
    "主視覺集中於中心，四角與邊緣以小元素呼應",
    "善用明度漸層，讓中心向外自然遞變、富立體感",
    "鄰近色鋪底＋互補色點睛，凸顯中心焦點",
    "最後鑲嵌施華洛世奇鑽飾，提升質感與光澤",
]
y=3.5
for p in pts:
    rect(s, 0.9, y+0.08, 0.16, 0.16, GOLD, shape=MSO_SHAPE.OVAL)
    txt(s, 1.28, y-0.03, 5.5, 0.55, p, 13.5, color=INK, spacing=1.15)
    y += 0.62
pic_cover(s, f"{L1}/art_octagon4.jpg", 7.3, 1.75, 5.1, 5.0)

s = slide(); bg(s)
eyebrow(s, 0.9, 0.5, "GRADUATION  WORKS")
sec_title(s, 0.85, 0.85, "考核通過作品", size=28)
txt(s, 0.9, 1.62, 11.5, 0.5,
    "完成 22 小時課程並通過考核——八角板木器 × 施華洛世奇鑽飾。", 14, color=TAUPE)
arts=["art1.jpg","art2.jpg","art3.jpg","art4.jpg","art5.jpg","art6.jpg"]
cw=3.9; ch=2.65; gx=0.28; gy=0.12; x0=(SW-(cw*3+gx*2))/2
for i,a in enumerate(arts):
    r,c=divmod(i,3)
    pic_contain(s, f"{L1}/{a}", x0+c*(cw+gx), 1.88+r*(ch+gy), cw, ch)

# 學員成果・第一屆（6 張合照，2×3 格）
s = slide(); bg(s)
eyebrow(s, 0.9, 0.5, "GRADUATION  PRIDE")
sec_title(s, 0.85, 0.85, "學員成果・第一屆驕傲時刻", size=26)
txt(s, 0.9, 1.62, 11.5, 0.45, "從一張底板，到捧在手心的完成作品——這是屬於你的里程碑。", 14, color=TAUPE)
c1_pride = ["ind2.jpg","ind3.jpg","ind4.jpg","inst2.jpg","inst3.jpg","inst4.jpg"]
cw=3.9; ch=2.65; gx=0.28; gy=0.12; x0=(SW-(cw*3+gx*2))/2
for i,fn in enumerate(c1_pride):
    r,c=divmod(i,3)
    pic_contain(s, f"{L1}/c1/{fn}", x0+c*(cw+gx), 1.88+r*(ch+gy), cw, ch)

# 學員成果・第二屆（3 張）
photo_wall("學員成果・第二屆驕傲時刻",
           "專注凝神的每一筆，化為最驕傲的一刻。",
           [f"{L1}/c2/470940_0.jpg", f"{L1}/c2/470941_0.jpg", f"{L1}/c2/470942_0.jpg"])

# 學員成果・第三屆（4 張横排）
s = slide(); bg(s)
eyebrow(s, 0.9, 0.5, "GRADUATION  PRIDE")
sec_title(s, 0.85, 0.85, "學員成果・第三屆驕傲時刻", size=26)
txt(s, 0.9, 1.62, 11.5, 0.45, "用色彩寫下屬於自己的美麗故事。", 14, color=TAUPE)
c3_pride = ["470952_0.jpg","470954_0.jpg","470955_0.jpg","470956_0.jpg"]
cw=2.95; ch=4.6; gx=0.35; x0=(SW-(cw*4+gx*3))/2
for i,fn in enumerate(c3_pride):
    pic_cover(s, f"{L1}/c3/{fn}", x0+i*(cw+gx), 2.2, cw, ch)

# 結業後支持
s = slide(); bg(s)
header(s, "AFTER  GRADUATION", "結業後・持續陪伴")
bens = [
    "免費教學輔導，陪你站穩講台",
    "以講師價進貨材料，適用教學販售",
    "三年內免學費複訓（僅收場地材料餐費）",
    "教學助教機會，累積實戰經驗",
    "專屬師資群組即時支援",
    "一級升級講師培訓考核通過，可教授二級講師培訓課程",
    "課程已納入公務人員終身學習認證",
    "年度師生成果展，增加曝光與資歷",
]
cw=5.7
for i,b in enumerate(bens):
    c = 0 if i<4 else 1; r = i if i<4 else i-4
    x=0.9+c*6.0; y=1.95+r*1.1
    rect(s, x, y, cw, 0.9, WHITE, line=CREAM_DEEP, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x+0.25, y+0.31, 0.28, 0.28, SAGE, shape=MSO_SHAPE.OVAL)
    txt(s, x+0.72, y, cw-0.9, 0.9, b, 13.5, color=INK, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================
#  結尾（bg_ending.jpg 底圖，白色半透明圓角矩形顯文字）
# ============================================================
s = slide()
s.background.fill.solid(); s.background.fill.fore_color.rgb = CREAM
END_IMG = f"{L1}/bg_ending.jpg"
if os.path.exists(END_IMG):
    p = s.shapes.add_picture(END_IMG, Inches(0), Inches(0), Inches(SW), Inches(SH))
    p.line.fill.background()
cx, cw, cy, ch = 2.0, 9.33, 1.55, 4.45
ov_text = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(cw), Inches(ch))
ov_text.fill.solid(); ov_text.fill.fore_color.rgb = WHITE
ov_text.line.fill.background(); ov_text.shadow.inherit = False
set_alpha(ov_text, 72)
rect(s, cx+1.5, cy+0.28, cw-3.0, 0.04, GOLD)
txt(s, cx, cy+0.52, cw, 1.1, "一起，從會畫到能教", 46, color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
rect(s, cx+3.5, cy+1.68, cw-7.0, 0.04, GOLD)
txt(s, cx, cy+1.82, cw, 0.65, "曼陀羅心靈彩繪・第十七屆一級講師課程", 20, color=INK, bold=True, align=PP_ALIGN.CENTER)
txt(s, cx, cy+2.52, cw, 0.52, "台中・平日班　6/15（一）、6/16（二）　09:00–21:00", 16, color=ROSE_DEEP, bold=True, align=PP_ALIGN.CENTER)
txt(s, cx, cy+3.1, cw, 0.42, "上課地點：台中市西屯區中康街 7 號（瑞恩悅琚）・小班制不超過 10 人", 13, color=TAUPE, align=PP_ALIGN.CENTER)
rect(s, cx+1.5, cy+3.72, cw-3.0, 0.04, GOLD)

out = "/home/user/mandala/一級講師課程教學過程.pptx"
prs.save(out)
print(f"SAVED {out}  slides={len(prs.slides._sldIdLst)}")
