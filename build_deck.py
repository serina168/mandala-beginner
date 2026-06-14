# -*- coding: utf-8 -*-
"""Enhance the mandala course deck: cohesive design, backgrounds, framed images.
Text content is preserved verbatim; only styling/layout/imagery is added."""
import copy
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

SRC = "original.pptx"
OUT = "繽紛五感體驗曼陀羅書籤畫_課程大綱.pptx"
A = "_assets/"

# ---- palette (Morandi: muted, warm, dusty tones) ----
INK   = RGBColor(0x4B,0x46,0x41)   # warm charcoal (body)
SUBINK= RGBColor(0x8A,0x81,0x78)   # muted greige
DEEP  = RGBColor(0x57,0x6B,0x67)   # muted teal-grey (titles)
GOLD  = RGBColor(0xAB,0x90,0x55)   # muted ochre / brass (rules)
GOLDL = RGBColor(0xC0,0xA6,0x6B)   # light muted ochre
ROSE  = RGBColor(0xB0,0x89,0x7A)   # dusty rose (secondary accent)
TEAL  = RGBColor(0x7C,0x8C,0x84)   # muted sage-teal
CREAM = RGBColor(0xEF,0xE9,0xDF)   # warm greige background
PANEL = RGBColor(0xFB,0xF8,0xF1)   # soft warm white panel
PBORD = RGBColor(0xD9,0xD0,0xC2)   # panel border
WHITE = PANEL                       # use warm white instead of pure white

prs = Presentation(SRC)
SW, SH = prs.slide_width, prs.slide_height
EMU = 914400

def IN(v): return Emu(int(v*EMU))

# ---------- xml / zorder helpers ----------
def spTree(slide): return slide.shapes._spTree
def send_to_back(slide, shape):
    tree = spTree(slide); el = shape._element; tree.remove(el); tree.insert(2, el)
def send_after_bg(slide, shape):
    # place right after background (index 2) -> behind original content
    tree = spTree(slide); el = shape._element; tree.remove(el); tree.insert(3, el)

def set_subpix_shadow(shape, blur=0.08, dist=0.06, dir=5400000, alpha=62000):
    spPr = shape._element.spPr
    # remove existing effectLst
    for e in spPr.findall(qn('a:effectLst')): spPr.remove(e)
    eff = spPr.makeelement(qn('a:effectLst'), {})
    sh = eff.makeelement(qn('a:outerShdw'), {
        'blurRad': str(int(blur*EMU)), 'dist': str(int(dist*EMU)),
        'dir': str(dir), 'rotWithShape':'0'})
    clr = sh.makeelement(qn('a:srgbClr'), {'val':'3A2E2E'})
    al = clr.makeelement(qn('a:alpha'), {'val':str(alpha)})
    clr.append(al); sh.append(clr); eff.append(sh); spPr.append(eff)

def round_corners(pic, radius=0.06):
    spPr = pic._element.spPr
    for g in spPr.findall(qn('a:prstGeom')): spPr.remove(g)
    geom = spPr.makeelement(qn('a:prstGeom'), {'prst':'roundRect'})
    av = geom.makeelement(qn('a:avLst'), {})
    gd = av.makeelement(qn('a:gd'), {'name':'adj','fmla':'val %d'%int(radius*100000)})
    av.append(gd); geom.append(av)
    # insert prstGeom after xfrm
    xfrm = spPr.find(qn('a:xfrm'))
    xfrm.addnext(geom)

def pic_border(pic, color=WHITE, w=2.2):
    pic.line.color.rgb = color; pic.line.width = Pt(w)

# ---------- background ----------
def add_bg(slide, img):
    pic = slide.shapes.add_picture(A+img, 0, 0, SW, SH)
    send_to_back(slide, pic)
    return pic

# ---------- cover-crop image placement ----------
def place_image(slide, path, x, y, w, h, rounded=True, border=True,
                bcolor=WHITE, bw=2.2, shadow=True, radius=0.05):
    im = Image.open(path); iw, ih = im.size
    box_ar = w/h; src_ar = iw/ih
    cl=cr=ct=cb=0.0
    if src_ar > box_ar:   # too wide -> crop sides
        new = box_ar/src_ar; cl=cr=(1-new)/2
    else:                 # too tall -> crop top/bottom
        new = src_ar/box_ar; ct=cb=(1-new)/2
    pic = slide.shapes.add_picture(path, IN(x), IN(y), IN(w), IN(h))
    pic.crop_left=cl; pic.crop_right=cr; pic.crop_top=ct; pic.crop_bottom=cb
    if rounded: round_corners(pic, radius)
    if border: pic_border(pic, bcolor, bw)
    if shadow: set_subpix_shadow(pic)
    return pic

# ---------- rounded panel ----------
def add_panel(slide, x, y, w, h, fill=WHITE, alpha=None, line=None, lw=1.0,
              radius=0.04, shadow=True, behind=True):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, IN(x),IN(y),IN(w),IN(h))
    sp.adjustments[0] = radius
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if alpha is not None:
        # set fill alpha
        sf = sp.fill.fore_color._xFill.find(qn('a:srgbClr'))
        a = sf.makeelement(qn('a:alpha'), {'val':str(int(alpha*1000))}); sf.append(a)
    if line is not None:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    else:
        sp.line.fill.background()
    if shadow: set_subpix_shadow(sp, blur=0.10, dist=0.05, alpha=40000)
    if behind: send_after_bg(slide, sp)
    sp.shadow.inherit = False
    return sp

# ---------- thin rule ----------
def add_rule(slide, x, y, w, color=GOLD, h=0.035):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, IN(x),IN(y),IN(w),IN(h))
    sp.fill.solid(); sp.fill.fore_color.rgb=color; sp.line.fill.background()
    sp.shadow.inherit=False
    return sp

def add_motif(slide, path, x, y, w, h):
    pic = slide.shapes.add_picture(A+path, IN(x),IN(y),IN(w),IN(h))
    return pic

# ---------- text styling ----------
def style_runs(tf, name="微軟正黑體", color=None, size=None, bold=None):
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = name
            _set_ea_font(r, name)
            if color is not None: r.font.color.rgb = color
            if size is not None: r.font.size = Pt(size)
            if bold is not None: r.font.bold = bold

def _set_ea_font(run, name):
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin','a:ea','a:cs'):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {}); rPr.append(e)
        e.set('typeface', name)

def set_line_spacing(tf, mult=1.12, space_after=6):
    for p in tf.paragraphs:
        p.line_spacing = mult
        p.space_after = Pt(space_after)

def shrink_fit(tf, scale=88, ln=10):
    """write normAutofit so text shrinks to fit shape (content unchanged)."""
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    for t in ('a:normAutofit','a:spAutoFit','a:noAutofit'):
        e=bodyPr.find(qn(t))
        if e is not None: bodyPr.remove(e)
    na = bodyPr.makeelement(qn('a:normAutofit'),
        {'fontScale':str(scale*1000),'lnSpcReduction':str(ln*1000)})
    bodyPr.append(na)

def set_insets(shape, l=0.18,r=0.18,t=0.14,b=0.14):
    tf=shape.text_frame
    tf.margin_left=IN(l); tf.margin_right=IN(r); tf.margin_top=IN(t); tf.margin_bottom=IN(b)
    tf.word_wrap=True

def _text_shapes(slide):
    return [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]

def _pick_title(cands):
    named=[s for s in cands if '標題' in s.name]
    pool=named or cands
    return min(pool, key=lambda s:s.top)

def classify(slide):
    """return (title_shape, body_shape) robustly despite messy original names."""
    ts=_text_shapes(slide)
    if not ts: return None,None
    longest=max(ts, key=lambda s:len(s.text_frame.text))
    if len(longest.text_frame.text) < 70:          # no real body (image/gallery slide)
        return _pick_title(ts), None
    body=longest
    cands=[s for s in ts if s is not body]
    return (_pick_title(cands) if cands else None), body

def find_title(slide):  return classify(slide)[0]
def biggest_text(slide, exclude=None): return classify(slide)[1]

def clear_shape_fill(shape):
    try: shape.fill.background()
    except Exception: pass
    try: shape.line.fill.background()
    except Exception: pass
    try: shape.shadow.inherit=False
    except Exception: pass

def all_pictures(slide):
    return [sh for sh in slide.shapes if sh.shape_type==13]
def empty_placeholders(slide):
    res=[]
    for sh in slide.shapes:
        if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip() \
           and sh.shape_type!=13:
            # skip title placeholders
            if '標題' in sh.name or '副標' in sh.name: continue
            res.append(sh)
    return res

# ---------- title header (content slides) ----------
def style_title_content(slide, color=DEEP, size=None, rulew=None):
    t = find_title(slide)
    if not t: return None
    tf=t.text_frame
    txt=tf.text
    n=len(txt.replace("\n",""))
    sz = size or (30 if n<=14 else 26 if n<=26 else 22 if n<=40 else 19)
    style_runs(tf, color=color, size=sz, bold=True)
    vertical = Emu(t.width).inches < 2.0
    for p in tf.paragraphs:
        p.line_spacing=1.04
        if not vertical: p.alignment=PP_ALIGN.LEFT
    tf.word_wrap=True
    return t

# ---------- footer ----------
def add_footer(slide, idx, dark=False):
    tb = slide.shapes.add_textbox(IN(0.55), IN(7.06), IN(9.5), IN(0.34))
    tf=tb.text_frame; tf.word_wrap=False
    p=tf.paragraphs[0]; r=p.add_run()
    r.text="心語能量美學 ‧ 曼陀羅心靈彩繪"
    r.font.size=Pt(9); r.font.name="微軟正黑體"; _set_ea_font(r,"微軟正黑體")
    r.font.color.rgb = (CREAM if dark else SUBINK)
    # page number
    pb = slide.shapes.add_textbox(IN(12.3), IN(7.06), IN(0.7), IN(0.34))
    pf=pb.text_frame; pp=pf.paragraphs[0]; pp.alignment=PP_ALIGN.RIGHT
    pr=pp.add_run(); pr.text=str(idx); pr.font.size=Pt(10); pr.font.bold=True
    pr.font.name="微軟正黑體"; _set_ea_font(pr,"微軟正黑體")
    pr.font.color.rgb=(GOLDL if dark else GOLD)

# =====================================================================
#  per-slide configuration
# =====================================================================
TWOCOL = {        # slide -> image (dense + medium text, two-column)
 2:'artwork-05.jpg', 3:'artwork-04.jpg', 4:'artwork-06.jpg', 5:'artwork-07.jpg',
 6:'showcase-01.jpg', 7:'showcase-06.jpg', 8:'artwork-03.jpg', 9:'_assets/m_grid.jpg',
 10:'showcase-04.jpg', 14:'artwork-02.jpg', 15:'showcase-08.jpg', 16:'showcase-09.jpg',
 17:'showcase-07.jpg', 18:'showcase-05.jpg', 19:'_assets/m_bookmark.jpg', 26:'showcase-12.jpg',
 28:'_assets/sunset.jpg',
}
SECTION_BG = {2,10}
TITLE_SLIDE = 1
IMAGE_SLIDES = {11,12,13,20,21,22,23,24,25,27,29,30,31,32,33,34,35}

def title_header(slide, dark=False):
    t = find_title(slide)
    if not t: return
    clear_shape_fill(t)
    c = GOLDL if dark else DEEP
    style_title_content(slide, color=c)
    # gold rule under title
    try:
        ty = Emu(t.top).inches; th = Emu(t.height).inches; tx = Emu(t.left).inches
    except: tx,ty,th = 0.6,0.34,1.0
    ruley = min(ty+th-0.06, 1.5)
    if Emu(t.width).inches < 2:   # vertical side-title -> skip rule
        return
    add_rule(slide, max(tx,0.62), ruley, 2.4, color=(GOLDL if dark else GOLD))

def enhance_two_column(slide, idx, img):
    body = biggest_text(slide, exclude=find_title(slide))
    # ---- title spans the top ----
    t = find_title(slide)
    if t:
        t.left=IN(0.62); t.top=IN(0.34); t.width=IN(12.1); t.height=IN(1.05)
    title_header(slide)
    # ---- geometry ----
    bx,by,bw,bh = 0.55,1.66,8.40,5.15
    ix,iy,iw,ih = 9.18,1.66,3.58,5.15
    # body panel
    add_panel(slide, bx,by,bw,bh, fill=WHITE, alpha=96, line=PBORD, lw=1.0)
    # body text shape -> reposition into panel, strip its own fill
    if body is not None:
        body.left=IN(bx+0.10); body.top=IN(by+0.06); body.width=IN(bw-0.20); body.height=IN(bh-0.12)
        try: body.fill.background()
        except: pass
        try: body.line.fill.background()
        except: pass
        body.shadow.inherit=False
        set_insets(body, 0.18,0.18,0.12,0.12)
        tf=body.text_frame
        style_runs(tf, color=INK)
        set_line_spacing(tf, 1.12, 4)
        n=len(tf.text)
        if   n>430: shrink_fit(tf, 78, 12)
        elif n>330: shrink_fit(tf, 84, 10)
        elif n>240: shrink_fit(tf, 90, 8)
        elif n>150: shrink_fit(tf, 96, 6)
    # right image
    p = img if img.startswith('_assets/') else img
    place_image(slide, p, ix,iy,iw,ih, rounded=True, border=True, bcolor=WHITE, bw=3,
                shadow=True, radius=0.05)
    # gold corner accent over image top
    add_rule(slide, ix, iy-0.02, iw, color=GOLD, h=0.06)
    add_footer(slide, idx)

def frame_existing_pictures(slide):
    for pic in all_pictures(slide):
        try:
            round_corners(pic, 0.04); pic_border(pic, WHITE, 2.4); set_subpix_shadow(pic)
        except Exception as e:
            pass

def enhance_image_slide(slide, idx):
    title = find_title(slide)
    title_header(slide)
    frame_existing_pictures(slide)
    # style captions / lift any real body text onto a clean panel
    tel = title._element if title is not None else None
    for sh in _text_shapes(slide):
        if tel is not None and sh._element is tel: continue
        txt = sh.text_frame.text
        if len(txt) >= 70:                       # a real text block (e.g. slide 13)
            clear_shape_fill(sh)
            x=Emu(sh.left).inches; y=Emu(sh.top).inches
            w=Emu(sh.width).inches; h=Emu(sh.height).inches
            add_panel(slide, x-0.14, y-0.14, w+0.28, h+0.28, fill=WHITE, alpha=96,
                      line=PBORD, lw=1.0)
            set_insets(sh, 0.20,0.20,0.14,0.14)
            style_runs(sh.text_frame, color=INK); set_line_spacing(sh.text_frame, 1.16, 6)
        else:                                    # caption
            style_runs(sh.text_frame, color=DEEP, bold=True)
            for p in sh.text_frame.paragraphs: p.alignment=PP_ALIGN.CENTER
    # fill empty placeholders with curated images
    fills = FILL_MAP.get(idx, [])
    eps = empty_placeholders(slide)
    for ph,(imgpath,) in zip(eps, [(f,) for f in fills]):
        x=Emu(ph.left).inches; y=Emu(ph.top).inches
        w=Emu(ph.width).inches; h=Emu(ph.height).inches
        place_image(slide, imgpath, x,y,w,h, rounded=True, border=True, bcolor=WHITE, bw=3)
    add_footer(slide, idx)

# explicit fills for empty placeholders / blank slides
FILL_MAP = {
 11:['showcase-02.jpg','showcase-03.jpg'],
 12:['_assets/m_round1.jpg','_assets/m_tree.jpg'],
 13:['_assets/singingbowl.jpg'],
 29:['_assets/autumn.jpg','_assets/sunset.jpg'],
}

# =====================================================================
#  main loop
# =====================================================================
for i, slide in enumerate(prs.slides, start=1):
    if i == TITLE_SLIDE:
        add_bg(slide, 'bg_title.png')
        CREAMTX = RGBColor(0xF3,0xEE,0xE2)
        # top decorative accent
        add_motif(slide, 'motif_ring_gold.png', 1.02, 0.62, 0.92, 0.92)
        add_rule(slide, 1.05, 1.78, 1.7, color=GOLDL, h=0.045)
        # big title
        t = find_title(slide)
        if t:
            t.left=IN(1.02); t.top=IN(2.0); t.width=IN(7.9); t.height=IN(2.7)
            tf=t.text_frame; tf.word_wrap=True
            style_runs(tf, color=CREAMTX, size=46, bold=True)
            for p in tf.paragraphs: p.line_spacing=1.14; p.alignment=PP_ALIGN.LEFT
        # divider rule + subtitle card with vertical accent bar
        add_rule(slide, 1.05, 4.78, 3.0, color=GOLDL, h=0.05)
        add_rule(slide, 1.05, 5.12, 0.07, color=ROSE, h=1.05)   # vertical accent bar
        sub=None
        for sh in slide.shapes:
            if sh.has_text_frame and ('副標' in sh.name):
                sub=sh; break
        if sub:
            sub.left=IN(1.32); sub.top=IN(5.08); sub.width=IN(7.0); sub.height=IN(1.6)
            sf=sub.text_frame; sf.word_wrap=True
            style_runs(sf, color=CREAMTX, size=22, bold=False)
            for p in sf.paragraphs: p.alignment=PP_ALIGN.LEFT; p.line_spacing=1.25
        continue

    bg = 'bg_section.png' if i in SECTION_BG else 'bg_content.png'
    add_bg(slide, bg)

    if i in TWOCOL:
        enhance_two_column(slide, i, TWOCOL[i])
    elif i in IMAGE_SLIDES:
        enhance_image_slide(slide, i)
    else:
        title_header(slide); add_footer(slide, i)

    # decorative top-right motif on non-title slides
    add_motif(slide, 'motif_ring_gold.png', 12.5, 0.18, 0.6, 0.6)

print("done, slides:", len(prs.slides))
prs.save(OUT)
print("saved", OUT)
