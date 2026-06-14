# -*- coding: utf-8 -*-
"""Enhance the mandala course deck: cohesive design, backgrounds, framed images.
Text content is preserved verbatim; only styling/layout/imagery is added."""
import copy, math, re
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
def set_geom(pic, prst, radius=0.05):
    spPr = pic._element.spPr
    for g in spPr.findall(qn('a:prstGeom')): spPr.remove(g)
    geom = spPr.makeelement(qn('a:prstGeom'), {'prst':prst})
    av = geom.makeelement(qn('a:avLst'), {})
    if prst=='roundRect':
        gd = av.makeelement(qn('a:gd'), {'name':'adj','fmla':'val %d'%int(radius*100000)})
        av.append(gd)
    geom.append(av)
    spPr.find(qn('a:xfrm')).addnext(geom)

def place_image(slide, path, x, y, w, h, rounded=True, border=True,
                bcolor=WHITE, bw=2.2, shadow=True, radius=0.05, circle=False,
                contain=False):
    im = Image.open(path); iw, ih = im.size
    box_ar = w/h; src_ar = iw/ih
    if contain and not circle:
        # show the WHOLE image (no crop): scale to fit, center in the region
        if src_ar > box_ar: nw=w; nh=w/src_ar
        else:               nh=h; nw=h*src_ar
        x=x+(w-nw)/2.0; y=y+(h-nh)/2.0; w=nw; h=nh
        pic = slide.shapes.add_picture(path, IN(x), IN(y), IN(w), IN(h))
    else:
        cl=cr=ct=cb=0.0
        if src_ar > box_ar:   cl=cr=(1-(box_ar/src_ar))/2
        else:                 ct=cb=(1-(src_ar/box_ar))/2
        pic = slide.shapes.add_picture(path, IN(x), IN(y), IN(w), IN(h))
        pic.crop_left=cl; pic.crop_right=cr; pic.crop_top=ct; pic.crop_bottom=cb
    if circle:    set_geom(pic, 'ellipse')
    elif rounded: set_geom(pic, 'roundRect', radius)
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

def _clear_autofit(tf):
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    for t in ('a:normAutofit','a:spAutoFit','a:noAutofit'):
        e=bodyPr.find(qn(t))
        if e is not None: bodyPr.remove(e)
    return bodyPr

def set_autofit(tf, mode='norm'):
    """mode 'norm' = shrink-text-to-fit (safety net); 'none' = no autofit (no grow)."""
    bodyPr=_clear_autofit(tf)
    tag = 'a:normAutofit' if mode=='norm' else 'a:noAutofit'
    bodyPr.append(bodyPr.makeelement(qn(tag), {}))

def _vlen(s):
    """visual length: CJK/full-width = 1.0, latin/space = 0.55."""
    t=0.0
    for ch in s:
        if ch=='\n': continue
        t += 1.0 if ord(ch) > 0x2000 else 0.55
    return t

# 微軟正黑體 real line box ≈ 1.32 em; conservative width/height usage so text
# truly fits inside the box when opened in PowerPoint (not just the LO preview).
FONT_LH = 1.38
def fit_size(shape, max_pt, min_pt, ls, sa_pt=4):
    """largest integer pt (max..min) at which the text fits inside the shape box."""
    tf=shape.text_frame
    w=Emu(shape.width).inches - Emu(tf.margin_left).inches - Emu(tf.margin_right).inches
    h=Emu(shape.height).inches - Emu(tf.margin_top).inches - Emu(tf.margin_bottom).inches
    paras=[p.text for p in tf.paragraphs]
    if w<=0.3 or h<=0.2: return min_pt
    for S in range(int(max_pt), int(min_pt)-1, -1):
        cpl=max(1.0, (w*72.0/S)*0.90)          # assume wider effective advance
        lines=0
        for pt in paras:
            # \x0b / \n inside a paragraph are soft line breaks -> own line(s)
            for seg in re.split('[\x0b\n]', pt):
                lines += max(1, math.ceil(_vlen(seg)/cpl))
        line_h = S*max(ls,1.0)*FONT_LH/72.0    # real CJK line height
        total = lines*line_h + len(paras)*(sa_pt/72.0)
        if total <= h*0.90:
            return S
    return int(min_pt)

def apply_size(tf, size):
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(size)

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
    if not vertical:                       # shrink long titles to stay inside the box
        fs = fit_size(t, sz, 13, 1.04, sa_pt=0)
        if fs < sz: apply_size(tf, fs)
        set_autofit(tf, 'none')            # never auto-grow over the body
    return t

# ---------- footer ----------
def add_footer(slide, idx, dark=False):
    tb = slide.shapes.add_textbox(IN(0.55), IN(7.06), IN(9.5), IN(0.34))
    tf=tb.text_frame; tf.word_wrap=False
    p=tf.paragraphs[0]; r=p.add_run()
    r.text="點鏡藝術工作室　‧　一點一滴，照見自己。"
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
TWOCOL = {        # slide -> image (作者本人創作照 + 純作品圖)
 2:'artwork-05.jpg', 3:'artwork-04.jpg', 4:'artwork-06.jpg', 5:'artwork-07.jpg',
 6:'_assets/teacher_studio2.jpg', 7:'artwork-03.jpg', 8:'_assets/m_grid.jpg', 9:'artwork-01.jpg',
 10:'_assets/teacher_studio1.jpg', 14:'artwork-02.jpg', 15:'_assets/m_owl.jpg', 16:'_assets/m_tree.jpg',
 17:'_assets/m_round1.jpg', 18:'_assets/m_img5.jpg', 19:'_assets/m_bookmark.jpg',
 26:'_assets/m_grid.jpg', 28:'_assets/sunset.jpg',
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
    bx,by,bw,bh = 0.55,1.58,8.40,5.42
    ix,iy,iw,ih = 9.18,1.58,3.58,5.42
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
        set_insets(body, 0.16,0.16,0.10,0.10)
        tf=body.text_frame
        style_runs(tf, color=INK)
        set_line_spacing(tf, 1.05, 3)
        # explicit size that truly fits, + normAutofit as a safety net
        apply_size(tf, fit_size(body, 18, 9, 1.05, sa_pt=3))
        set_autofit(tf, 'norm')
    # right image — show the WHOLE picture (no crop), scaled to fit & centered
    place_image(slide, img, ix,iy,iw,ih, rounded=True, border=True, bcolor=WHITE, bw=3,
                shadow=True, radius=0.05, contain=True)
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
            set_insets(sh, 0.18,0.18,0.12,0.12)
            style_runs(sh.text_frame, color=INK); set_line_spacing(sh.text_frame, 1.08, 4)
            apply_size(sh.text_frame, fit_size(sh, 18, 10, 1.08, sa_pt=4))
            set_autofit(sh.text_frame, 'norm')
        else:                                    # caption
            style_runs(sh.text_frame, color=DEEP, bold=True)
            for p in sh.text_frame.paragraphs: p.alignment=PP_ALIGN.CENTER
    # fill empty placeholders with curated images
    fills = FILL_MAP.get(idx, [])
    eps = empty_placeholders(slide)
    for ph,(imgpath,) in zip(eps, [(f,) for f in fills]):
        x=Emu(ph.left).inches; y=Emu(ph.top).inches
        w=Emu(ph.width).inches; h=Emu(ph.height).inches
        place_image(slide, imgpath, x,y,w,h, rounded=True, border=True, bcolor=WHITE, bw=3, contain=True)
    add_footer(slide, idx)

# explicit fills for empty placeholders / blank slides
FILL_MAP = {
 11:['artwork-02.jpg','artwork-05.jpg'],
 12:['_assets/m_owl.jpg','_assets/m_tree.jpg'],
 13:['_assets/singingbowl.jpg'],
 29:['_assets/autumn.jpg','_assets/sunset.jpg'],
}

# =====================================================================
#  main loop
# =====================================================================
def replace_terms(prs, mapping):
    n=0
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame: continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    for a,b in mapping.items():
                        if a in r.text:
                            r.text=r.text.replace(a,b); n+=1
    return n

# 將內容由「長照導向」改寫為「適用於所有人」的廣泛運用（保留原意與架構）
REWRITE_TITLE = {
 7:  "繪畫創作與曼陀羅的核心應用優點與實務做法",
 16: "4 ╴1、心感的內在覺察～能夠強烈且深刻地引發每個人與身邊夥伴彼此的內在覺察。",
}
REWRITE_BODY = {
 4: """一、藝術輔療的定義、目標與專業角色
二、探討其在日常生活與各種場域的應用範圍
特別著重於如何透過藝術媒介促進每個人的情緒表達、認知刺激與社交互動。
三、曼陀羅從最初的「宗教宇宙觀」，演變為今日在身心調適、心理諮商中廣泛應用的「心靈自癒工具」。""",
 7: """（1）認知刺激與大腦活化
a、激活大腦神經：繪畫時挑選顏色、構圖、抓握畫筆，需要同時動用視覺、空間感與微細肌肉運動，能有效刺激大腦神經網絡。
b、維持專注與協調：透過結構化的繪畫活動（如曼陀羅著色、線條禪繞），能幫助每個人維持專注力與手眼協調能力。
（2）情緒支持與非言語溝通
a、突破語言的限制：許多情緒與感受難以用言語精確表達。繪畫提供非言語的溝通管道，讓人用色彩與線條表達內心感受與需求。
b、安定情緒、減少焦躁：專注於創作能帶來平靜，顯著降低焦慮、煩躁與低落情緒。
（3）連結記憶與自我認同
a、重塑生命故事：引導繪畫家鄉、童年或生命中重要的時刻（懷舊繪畫），能喚醒深層記憶。
b、重拾成就與掌控感：在畫紙上，每個人都能全權決定色彩與圖案，重新獲得成就感與自我價值。
（4）社交互動與減少孤獨
a、團體創作促進交流：透過集體繪畫或作品分享，能讓彼此互相欣賞、讚美，建立人際連結並降低孤獨感。""",
 10: """透過
1、視覺色彩選擇 2、聽覺音樂引導 3、嗅覺精油輔助 以及 4、心感的內在覺察 5、觸覺引領，帶領學員進入專注創作的流動狀態。
學員將完成個人化的曼陀羅書籤作品，並學習如何將此技法應用於生活的各種情境中，提升身心安頓與生活福祉。""",
 14: """嗅覺是唯一不經過視丘、直接刺激大腦邊緣系統（情緒與記憶中心）的感官，能為每個人帶來立竿見影的效益。
對「自己」的好處
1、喚起生命記憶（懷舊療癒）：嗅覺能瞬間喚醒深層記憶。例如檜木或柑橘清香，能讓人聯想到森林或廚房的記憶，誘發繪畫創作的靈感與表達欲。
2、安定情緒、緩解焦躁：薰衣草、佛手柑等精油結合深呼吸，能激活副交感神經，有效緩解煩躁、不安與無故的焦慮。
3、提升呼吸順暢度與含氧量：透過尤加利或薄荷等精油進行呼吸引導，能擴張胸腔、促進血液循環，提升大腦的含氧與灌流。
4、非侵入性的溫和安撫：氣味引導不需要耗費體力，能提供最溫和、無壓力的撫慰與安全感，適合任何人、任何狀態。""",
 15: """對「身邊一同參與的人」的好處
1、即時抽離高壓、快速減壓：長期忙碌的人常處於慢性疲勞與神經緊繃。透過幾分鐘的精油深呼吸，能快速降低皮質醇（壓力荷爾蒙），讓大腦按下暫停鍵。
2、改善睡眠品質：壓力與煩惱常伴隨睡眠中斷。在夜間或休息時進行芳香呼吸，能引導大腦進入放鬆波，緩解神經性失眠與焦慮。
3、轉化負面情緒：生活中難免產生挫折與煩躁。柑橘類（如甜橙、葡萄柚）能帶來陽光般的能量，重塑正向情緒。
雙方共同的隱形好處：
當大家「同時」身處在同一個和諧的香氣與呼吸引導中，原本各自緊繃的狀態會轉化為共同體驗的「陪伴」。這能營造一個高度安心的氛圍，非常適合在呼吸放鬆後，順暢地過渡到曼陀羅的繪畫創作中。""",
 16: """對「長期承擔、忙於照顧他人或工作的人」的內在覺察：
從「外在承擔」回到「自我陪伴」——這樣的人日常往往把焦點放在「別人需要什麼」，完全壓抑了自我。
覺察身體的緊繃：在精油呼吸引導的當下，氣味會把注意力從外界拉回自己身上。他們會在深呼吸中突然驚覺：「原來我的肩膀這麼緊、呼吸這麼淺。」進而覺察到自己累積已久的疲憊。
看見被壓抑的情緒：在隨後的曼陀羅繪畫中，可能會自發地畫出混亂的線條、或選用大面積的灰與黑。這面畫紙就像心靈的鏡子，讓人坦然面對並看見內心深處的挫折、憤怒或愧疚。""",
 17: """2、對「感到無力、失去掌控感的人」的內在覺察：
從「無能為力」到「找回內在價值」——許多人在壓力或低潮中，覺得失去了對生活的掌控與聲音。
喚醒沉睡的自我本質：當特定的植物香氣（如泥土香、草本香）刺激大腦情緒中心時，會瞬間連結到生命中最有活力的歲月。這能幫助人覺察到自己不只是被處境困住，而是一個擁有豐富生命故事的個體。
透過色彩宣洩內在的聲音：難以用言語清楚表達的情緒，會透過色彩展現。當在曼陀羅中央填入強烈的紅色時，可能是在表達內在的痛楚或不滿；填入藍色時，可能代表渴望平靜。這個過程能讓人重新與自己的身體與心理感受產生連結。""",
 18: """3、彼此「共同」的覺察：
看見關係的質變——當大家並肩坐著，聞著同一股香氣、畫著各自的曼陀羅時，會觸發一種「平行共處」的奇妙體驗。
看見彼此都是獨立的靈魂：在作品完成後的分享（或靜默欣賞）中，往往會驚訝地發現：原來對方心裡還渴望著這樣的色彩，遠比我們以為的更立體。也能感受到彼此在畫紙上流露出的壓力與溫柔。
放下對立，轉為同理：這種不帶批判的藝術空間，能讓彼此同時覺察到：「我們都在各自的生活裡努力著。」原本的緊張或對立，會在此刻轉化為深度的生命共鳴與同理心。""",
 19: """心理學上的藝術輔療以及科技領域中，它的意思如下：
Tactile（觸覺的）：指我們透過皮膚、手指、手掌去觸摸、按壓、感覺物體時的身體感官。
Feedback（反饋/回饋）：指物體接觸到我們身體後，回傳給大腦的信號與物理反應。
在藝術輔療情境中，它的意思是：當我們用手去接觸畫筆、紙張或顏料時，身體所得到的「真實觸感」與「物理抵抗力」。
例如：當畫筆按壓紙張的時候，讓大腦知道「我正在做這件事」，進而產生安心感。用重複的蓋印章動作，這種簡單的觸感直接安撫了緊繃的神經。印上顏料時的厚實與變形，手掌感受到的壓力和反彈，讓人宣洩內心的力道與情緒。
簡單來說，Tactile Feedback 就是「手指與皮膚摸到東西時，大腦所接收到的真實感受與回報」。""",
 26: """1、纏枝紋的「生命力」寓意——纏枝紋（又稱「萬壽藤」）的核心特徵是連綿不斷、循環往復。
生生不息：它描繪植物藤蔓纏繞、委婉多姿的樣子，象徵著強大且持久的生命韌性。
循環與轉機：藤蔓無始無終的生長軌跡，寓意生命力的流轉與更新，與「生生不息、重獲生機」的祝福非常契合。
2、「瓶」與「平」的雙重祝福——梅瓶在民間工藝中本就有「平平安安」的諧音。
無論送給自己或所珍視的人，最美好的期盼莫過於「平安」。「瓶（平）」安與「纏枝（生生不息）」結合，構成了「平安健康、生命長青」的完整祝福。""",
}

def _set_text(shape, text):
    tf=shape.text_frame; tf.clear()
    for k,ln in enumerate(text.split('\n')):
        p=tf.paragraphs[0] if k==0 else tf.add_paragraph()
        p.add_run().text=ln

def apply_rewrites(prs):
    for idx,b in REWRITE_BODY.items():
        body=classify(prs.slides[idx-1])[1]
        if body is not None: _set_text(body, b)
    for idx,t in REWRITE_TITLE.items():
        ttl=find_title(prs.slides[idx-1])
        if ttl is not None: _set_text(ttl, t)

# 用語調整：治療 -> 輔療；其餘零星長照用語
print("replaced terms:", replace_terms(prs, {'治療':'輔療', '對長輩而言':'對某些人而言'}))
apply_rewrites(prs)

for i, slide in enumerate(prs.slides, start=1):
    if i == TITLE_SLIDE:
        add_bg(slide, 'bg_title.png')
        CREAMTX = RGBColor(0xF3,0xEE,0xE2)
        def _coverbox(text, x,y,w,h, size, color, bold, align=PP_ALIGN.LEFT, ls=1.12, spc=None):
            tb=slide.shapes.add_textbox(IN(x),IN(y),IN(w),IN(h)); tf=tb.text_frame; tf.word_wrap=True
            pp=tf.paragraphs[0]; pp.alignment=align; pp.line_spacing=ls
            r=pp.add_run(); r.text=text; r.font.size=Pt(size); r.font.bold=bold
            r.font.name="微軟正黑體"; _set_ea_font(r,"微軟正黑體"); r.font.color.rgb=color
            if spc is not None: r._r.get_or_add_rPr().set('spc', str(spc))
            return tb
        # top decorative accent + 工作室名稱
        add_motif(slide, 'motif_ring_gold.png', 1.02, 0.55, 0.82, 0.82)
        _coverbox("點鏡藝術工作室", 1.05, 1.46, 6.5, 0.5, 18, GOLDL, True, spc=300)
        add_rule(slide, 1.05, 2.02, 1.7, color=GOLDL, h=0.045)
        # big title
        t = find_title(slide)
        if t:
            t.left=IN(1.02); t.top=IN(2.22); t.width=IN(7.9); t.height=IN(2.05)
            # clean embedded soft-breaks/blank lines into tidy lines
            lines=[seg for seg in re.split('[\x0b\n]', t.text_frame.text) if seg.strip()]
            tf=t.text_frame; tf.clear(); tf.word_wrap=True
            for k,seg in enumerate(lines):
                pr = tf.paragraphs[0] if k==0 else tf.add_paragraph()
                pr.add_run().text=seg
            style_runs(tf, color=CREAMTX, size=46, bold=True)
            for p in tf.paragraphs: p.line_spacing=1.12; p.alignment=PP_ALIGN.LEFT
            apply_size(tf, fit_size(t, 46, 28, 1.12, sa_pt=0)); set_autofit(tf,'none')
        # 工作室標語
        _coverbox("一點一滴，照見自己。", 1.05, 4.42, 7.2, 0.66, 22, GOLDL, False, spc=200)
        # divider rule + subtitle card with vertical accent bar
        add_rule(slide, 1.05, 5.18, 3.0, color=GOLDL, h=0.05)
        add_rule(slide, 1.05, 5.46, 0.07, color=ROSE, h=0.62)   # vertical accent bar
        sub=None
        for sh in slide.shapes:
            if sh.has_text_frame and ('副標' in sh.name):
                sub=sh; break
        if sub:
            sub.left=IN(1.34); sub.top=IN(5.45); sub.width=IN(7.0); sub.height=IN(0.9)
            sf=sub.text_frame; sf.clear(); sf.word_wrap=True
            sf.paragraphs[0].add_run().text="主講：張婉愉老師"
            style_runs(sf, color=CREAMTX, size=29, bold=True)
            for p in sf.paragraphs: p.alignment=PP_ALIGN.LEFT; p.line_spacing=1.1
            set_autofit(sf,'none')
        # 封面僅使用曼陀羅作品（已在背景呈現），不放人物照
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
