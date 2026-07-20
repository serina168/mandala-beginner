from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math, os, random

SCRATCHPAD = "/tmp/claude-0/-home-user-mandala-beginner/c55d3146-c188-5ff5-8b4c-a49e578fc1a9/scratchpad"
FONTS = "/root/.claude/skills/canvas-design/canvas-fonts"
NOTO_S  = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc"
NOTO_SA = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"

W, H = 1087, 661

CREAM  = (245, 237, 224)
LAV1   = (205, 188, 225)   # medium lavender
LAV2   = (182, 160, 212)   # deeper lavender
BLUSH1 = (232, 180, 192)
BLUSH2 = (215, 155, 172)
PUR    = (148, 118, 180)
GOLD   = (201, 168, 76)
GOLDD  = (145, 112, 35)
GOLDL  = (228, 207, 140)
BROWN  = (70, 52, 36)


def make_bg(w, h, variant="front"):
    # Start with a lavender-tinted cream base
    base_color = (238, 228, 235)   # warm lavender-cream
    img = Image.new("RGBA", (w, h), base_color + (255,))

    washes = []

    if variant == "front":
        # BIG lavender wash top-left bleeding to center
        a1 = Image.new("RGBA",(w,h),(0,0,0,0))
        d1 = ImageDraw.Draw(a1)
        for i in range(25):
            a = int(130*(1-i/25)**1.2)
            d1.ellipse([-120-i*22,-90-i*14,540+i*20,470+i*16], fill=(195,172,222,a))
        a1 = a1.filter(ImageFilter.GaussianBlur(55))
        washes.append(a1)

        # Blush bottom-right
        a2 = Image.new("RGBA",(w,h),(0,0,0,0))
        d2 = ImageDraw.Draw(a2)
        for i in range(20):
            a = int(120*(1-i/20)**1.2)
            d2.ellipse([w+i*10,h+i*12,w+450,h+460], fill=(235,170,190,a))
        a2 = a2.filter(ImageFilter.GaussianBlur(62))
        washes.append(a2)

        # Secondary lavender top-right
        a3 = Image.new("RGBA",(w,h),(0,0,0,0))
        d3 = ImageDraw.Draw(a3)
        for i in range(16):
            a = int(95*(1-i/16)**1.3)
            d3.ellipse([w-260+i*14,-80-i*12,w+200-i*8,310+i*14], fill=(188,162,218,a))
        a3 = a3.filter(ImageFilter.GaussianBlur(48))
        washes.append(a3)

        # Warm cream island center-left to keep text readable
        a4 = Image.new("RGBA",(w,h),(0,0,0,0))
        d4 = ImageDraw.Draw(a4)
        for i in range(12):
            a = int(80*(1-i/12)**1.8)
            cx_,cy_ = int(w*0.26), int(h*0.50)
            r_ = int(185+i*38)
            d4.ellipse([cx_-r_,cy_-r_,cx_+r_,cy_+r_], fill=(252,242,228,a))
        a4 = a4.filter(ImageFilter.GaussianBlur(40))
        washes.append(a4)

    else:  # back
        a1 = Image.new("RGBA",(w,h),(0,0,0,0))
        d1 = ImageDraw.Draw(a1)
        for i in range(22):
            a = int(120*(1-i/22)**1.2)
            d1.ellipse([-100-i*20,-70-i*12,480+i*18,430+i*14], fill=(198,175,224,a))
        a1 = a1.filter(ImageFilter.GaussianBlur(52))
        washes.append(a1)

        a2 = Image.new("RGBA",(w,h),(0,0,0,0))
        d2 = ImageDraw.Draw(a2)
        for i in range(16):
            a = int(105*(1-i/16)**1.3)
            d2.ellipse([int(w*0.48)+i*18,-50+i*10,w+120,int(h*0.72)-i*8], fill=(210,185,225,a))
        a2 = a2.filter(ImageFilter.GaussianBlur(50))
        washes.append(a2)

        a3 = Image.new("RGBA",(w,h),(0,0,0,0))
        d3 = ImageDraw.Draw(a3)
        for i in range(18):
            a = int(110*(1-i/18)**1.2)
            d3.ellipse([w+i*10,h+i*10,w+440,h+460], fill=(232,172,192,a))
        a3 = a3.filter(ImageFilter.GaussianBlur(60))
        washes.append(a3)

        # Readable area for text
        a4 = Image.new("RGBA",(w,h),(0,0,0,0))
        d4 = ImageDraw.Draw(a4)
        for i in range(10):
            a = int(65*(1-i/10)**1.8)
            cx_,cy_ = int(w*0.30), int(h*0.52)
            r_ = int(200+i*35)
            d4.ellipse([cx_-r_,cy_-r_,cx_+r_,cy_+r_], fill=(252,242,226,a))
        a4 = a4.filter(ImageFilter.GaussianBlur(38))
        washes.append(a4)

    for w_ in washes:
        img = Image.alpha_composite(img, w_)
    return img


def draw_mandala(layer, cx, cy, R, opacity=1.0):
    sz = int(R*2.8)
    m = Image.new("RGBA",(sz,sz),(0,0,0,0))
    d = ImageDraw.Draw(m)
    mx=sz//2; my=sz//2

    def o(a): return max(0,min(255,int(a*opacity)))
    def gc(a): return GOLD+(o(a),)
    def glc(a): return GOLDL+(o(a),)
    def gdc(a): return GOLDD+(o(a),)
    def pc(a): return PUR+(o(a),)
    def bc(a): return BLUSH1+(o(a),)
    def lc(a): return LAV1+(o(a),)
    def l2c(a): return LAV2+(o(a),)

    def ell(r, oc, fw=1, fc=None):
        ri=int(r)
        d.ellipse([mx-ri,my-ri,mx+ri,my+ri], outline=oc, fill=fc, width=fw)

    def teardrop_petal(n, r_in, r_out, ang_off, fill_c, oc_a, squeeze=1.0):
        """True teardrop/almond shaped petals."""
        hw = math.pi/n * squeeze
        for i in range(n):
            ang = 2*math.pi*i/n + ang_off
            # Bezier-like approximation using many points
            pts = []
            steps = 32
            for t in range(steps+1):
                frac = t/steps
                # Shape: narrow at base, wide in middle, pointed at tip
                width_frac = math.sin(math.pi * frac)
                r = r_in + (r_out - r_in) * frac
                side = hw * width_frac
                a_point = ang - side + (2*side * (0.5))  # center line
                # Left side
                r_left = r
                a_left = ang - side + 0.01
                pts.append((int(mx+r_left*math.cos(ang-side*width_frac)),
                            int(my+r_left*math.sin(ang-side*width_frac))))
            # Right side (reversed)
            for t in range(steps,-1,-1):
                frac = t/steps
                width_frac = math.sin(math.pi * frac)
                r = r_in + (r_out - r_in) * frac
                side = hw * width_frac
                pts.append((int(mx+r*math.cos(ang+side*width_frac)),
                            int(my+r*math.sin(ang+side*width_frac))))
            if len(pts)>=3:
                d.polygon(pts, fill=fill_c, outline=gc(oc_a))

    def dots(n, r, dr=2.5, a=140):
        for i in range(n):
            ang=2*math.pi*i/n
            px=int(mx+r*math.cos(ang)); py=int(my+r*math.sin(ang))
            ri=int(dr)
            d.ellipse([px-ri,py-ri,px+ri,py+ri], fill=gc(a))

    def spokes(n, r0, r1, a_off=0, alpha=50):
        for i in range(n):
            ang=2*math.pi*i/n+a_off
            x0=int(mx+r0*math.cos(ang)); y0=int(my+r0*math.sin(ang))
            x1=int(mx+r1*math.cos(ang)); y1=int(my+r1*math.sin(ang))
            d.line([(x0,y0),(x1,y1)], fill=gc(alpha), width=1)

    # ===== Build mandala from center out =====

    # Center starburst
    for i in range(24):
        ang=2*math.pi*i/24
        x0=int(mx+R*0.02*math.cos(ang)); y0=int(my+R*0.02*math.sin(ang))
        x1=int(mx+R*0.12*math.cos(ang)); y1=int(my+R*0.12*math.sin(ang))
        d.line([(x0,y0),(x1,y1)], fill=gdc(200), width=1)

    # Center circle
    ell(R*0.08, gdc(240), 2, gdc(200))
    ell(R*0.05, gdc(255), 1, glc(230))

    # Ring 1: 8 gold teardrops
    teardrop_petal(8, R*0.10, R*0.24, math.pi/8, gc(110), 170, 0.90)
    ell(R*0.10, gc(150), 2)
    ell(R*0.26, gc(120), 1)
    dots(8, R*0.26, 2, 145)

    # Ring 2: 12 blush teardrops
    teardrop_petal(12, R*0.28, R*0.46, 0, bc(80), 130, 0.88)
    teardrop_petal(12, R*0.28, R*0.46, math.pi/12, glc(55), 90, 0.80)
    spokes(12, R*0.26, R*0.28, 0, 55)
    ell(R*0.28, gc(130), 2)
    dots(12, R*0.28, 2, 130)
    ell(R*0.48, gc(110), 1)
    dots(12, R*0.48, 2.2, 115)

    # Ring 3: 16 purple/lavender teardrops (larger)
    teardrop_petal(16, R*0.50, R*0.70, math.pi/16, pc(65), 110, 0.86)
    teardrop_petal(16, R*0.50, R*0.70, 0, lc(48), 80, 0.80)
    spokes(16, R*0.48, R*0.50, math.pi/16, 50)
    ell(R*0.50, gc(125), 2)
    dots(16, R*0.50, 2.2, 125)
    ell(R*0.72, gc(108), 2)
    dots(24, R*0.72, 2, 108)

    # Ring 4: 24 delicate outer teardrops
    teardrop_petal(24, R*0.73, R*0.91, 0, lc(32), 72, 0.82)
    teardrop_petal(24, R*0.73, R*0.91, math.pi/24, bc(22), 52, 0.76)
    spokes(24, R*0.72, R*0.73, 0, 42)
    spokes(12, R*0.72, R*0.91, 0, 35)

    # Outer rings
    ell(R*0.93, gc(88), 1)
    ell(R*0.955, gc(60), 1)
    dots(36, R*0.975, 2, 110)

    m = m.filter(ImageFilter.GaussianBlur(0.5))
    layer.paste(m, (int(cx-sz//2), int(cy-sz//2)), m)


def draw_lotus(layer, cx, cy, size):
    sz = int(size*4.0)
    m = Image.new("RGBA",(sz,sz),(0,0,0,0))
    d = ImageDraw.Draw(m)
    lx=sz//2; ly=sz//2

    def petal(n, r_in, r_out, ao, fill_c, oc_a, hw_f=0.88):
        hw=math.pi/n*hw_f
        for i in range(n):
            ang=2*math.pi*i/n+ao
            pts=[]
            for t in range(30):
                frac=t/29
                wf=math.sin(math.pi*frac)
                r=r_in+(r_out-r_in)*frac
                pts.append((int(lx+r*math.cos(ang-hw*wf)), int(ly+r*math.sin(ang-hw*wf))))
            for t in range(29,-1,-1):
                frac=t/29
                wf=math.sin(math.pi*frac)
                r=r_in+(r_out-r_in)*frac
                pts.append((int(lx+r*math.cos(ang+hw*wf)), int(ly+r*math.sin(ang+hw*wf))))
            d.polygon(pts, fill=fill_c, outline=GOLD+(oc_a,))

    petal(8, size*0.25, size*1.05, 0, BLUSH1+(95,), 108, 0.88)
    petal(6, size*0.20, size*0.85, math.pi/8, LAV1+(110,), 130, 0.85)
    petal(5, size*0.14, size*0.62, 0, PUR+(100,), 155, 0.83)
    petal(4, size*0.08, size*0.32, math.pi/4, GOLDL+(145,), 180, 0.85)

    rc=int(size*0.14); d.ellipse([lx-rc,ly-rc,lx+rc,ly+rc], fill=GOLD+(245,))
    rc2=int(size*0.07); d.ellipse([lx-rc2,ly-rc2,lx+rc2,ly+rc2], fill=GOLDD+(255,))

    m=m.filter(ImageFilter.GaussianBlur(0.55))
    layer.paste(m, (int(cx-sz//2), int(cy-sz//2)), m)


def draw_branch(layer, x, y, ang, length, depth, w0):
    if depth<=0 or length<9: return
    d=ImageDraw.Draw(layer)
    ex=int(x+length*math.cos(ang)); ey=int(y+length*math.sin(ang))
    d.line([(int(x),int(y)),(ex,ey)], fill=GOLD+(145,), width=max(1,int(w0)))
    if depth>=2:
        for t_f in [0.30,0.54,0.74]:
            lx_=int(x+length*t_f*math.cos(ang)); ly_=int(y+length*t_f*math.sin(ang))
            for side in [-1,1]:
                la=ang+side*(math.pi/4.1+0.08)
                ll=length*0.30
                lex=int(lx_+ll*math.cos(la)); ley=int(ly_+ll*math.sin(la))
                d.line([(lx_,ly_),(lex,ley)], fill=GOLD+(128,), width=max(1,int(w0-0.8)))
                sz_=max(4,int(ll*0.38))
                sz2=max(2,int(sz_*0.52))
                d.ellipse([lex-sz_,ley-sz2,lex+sz_,ley+sz2],
                          fill=GOLDL+(78,), outline=GOLD+(105,))
    draw_branch(layer, ex, ey, ang-0.38, length*0.64, depth-1, w0-0.52)
    draw_branch(layer, ex, ey, ang+0.40, length*0.60, depth-1, w0-0.52)


def contact_icon(draw, cx, cy, r, kind):
    d=draw; ri=int(r*0.64)
    bg_=CREAM
    d.ellipse([cx-r,cy-r,cx+r,cy+r], fill=GOLD+(218,), outline=GOLDD+(240,), width=1)
    if kind=="phone":
        for ofs in range(3):
            d.arc([cx-int(ri*0.75)+ofs*2,cy-ri+ofs*2,cx+int(ri*0.75)-ofs*2,cy+ri-ofs*2],
                  200,340, fill=bg_, width=2)
        d.ellipse([cx-3,cy-ri+2,cx+3,cy-ri+8],fill=bg_)
        d.ellipse([cx-3,cy+ri-8,cx+3,cy+ri-2],fill=bg_)
    elif kind=="pin":
        d.ellipse([cx-int(ri*0.72),cy-int(ri*1.05),cx+int(ri*0.72),cy+int(ri*0.05)],
                  fill=None, outline=bg_, width=2)
        pts=[(cx-int(ri*0.65),cy-int(ri*0.35)),(cx,cy+int(ri*0.95)),(cx+int(ri*0.65),cy-int(ri*0.35))]
        d.polygon(pts, fill=bg_)
        rc=int(ri*0.32)
        d.ellipse([cx-rc,cy-int(ri*0.82),cx+rc,cy-int(ri*0.82)+rc*2+rc], fill=GOLD+(215,))
    elif kind=="email":
        ew=int(ri*1.15); eh=int(ri*0.85)
        d.rectangle([cx-ew,cy-eh,cx+ew,cy+eh], outline=bg_, width=2)
        d.line([(cx-ew,cy-eh),(cx,cy+5),(cx+ew,cy-eh)], fill=bg_, width=2)
    elif kind=="line":
        d.rounded_rectangle([cx-ri,cy-ri,cx+ri,cy+ri], radius=int(ri*0.28), fill=bg_)
        d.rounded_rectangle([cx-ri+2,cy-ri+2,cx+ri-2,cy+ri-2],
                            radius=int(ri*0.22), fill=GOLD+(218,))
        fw=3; hs=int(ri*0.48); vs=int(ri*0.55)
        d.line([(cx-hs//2,cy-vs),(cx-hs//2,cy+vs)], fill=bg_, width=fw)
        d.line([(cx-hs//2,cy+vs),(cx+hs,cy+vs)], fill=bg_, width=fw)
    elif kind=="fb":
        d.rounded_rectangle([cx-ri,cy-ri,cx+ri,cy+ri], radius=int(ri*0.28), fill=bg_)
        d.rounded_rectangle([cx-ri+2,cy-ri+2,cx+ri-2,cy+ri-2],
                            radius=int(ri*0.22), fill=GOLD+(218,))
        fx=cx-int(ri*0.12); fw=3
        d.line([(fx,cy-int(ri*0.72)),(fx,cy+int(ri*0.72))], fill=bg_, width=fw)
        d.line([(fx,cy-int(ri*0.12)),(fx+int(ri*0.58),cy-int(ri*0.12))], fill=bg_, width=2)
        d.line([(fx-int(ri*0.14),cy-int(ri*0.68)),(fx+int(ri*0.52),cy-int(ri*0.68))], fill=bg_, width=2)
    elif kind=="ig":
        d.rounded_rectangle([cx-ri,cy-ri,cx+ri,cy+ri], radius=int(ri*0.38), fill=bg_)
        d.rounded_rectangle([cx-ri+3,cy-ri+3,cx+ri-3,cy+ri-3],
                            radius=int(ri*0.30), fill=GOLD+(215,))
        inn=int(ri*0.58)
        d.ellipse([cx-inn,cy-inn,cx+inn,cy+inn], outline=bg_, width=2)
        dc=int(ri*0.22); d.ellipse([cx-dc,cy-dc,cx+dc,cy+dc], fill=bg_)
        pr=int(ri*0.42); pa2=int(ri*0.74)
        d.ellipse([cx+pr,cy-pa2,cx+pa2,cy-pr], fill=bg_)


def gold_line(draw, y, x0, x1, diamond=True):
    y=int(y)
    for off,a in [(-1,52),(0,168),(1,52)]:
        draw.line([(int(x0),y+off),(int(x1),y+off)], fill=GOLD+(a,), width=1)
    if diamond:
        xm=int((x0+x1)/2); sz=5
        draw.polygon([(xm,y-sz),(xm+sz,y),(xm,y+sz),(xm-sz,y)], fill=GOLD+(215,))
        for dx in [-90,90]:
            draw.ellipse([xm+dx-2,y-2,xm+dx+2,y+2], fill=GOLDL+(175,))


# ============================================================
# FRONT
# ============================================================
print("Rendering front...")
bg = make_bg(W, H, "front")
canvas = bg.convert("RGBA")

ml = Image.new("RGBA",(W,H),(0,0,0,0))
draw_mandala(ml, int(W*0.735), H//2, 282, 0.88)

sm = Image.new("RGBA",(W,H),(0,0,0,0))
draw_mandala(sm, int(W*0.038), int(H*0.103), 100, 0.52)

bl = Image.new("RGBA",(W,H),(0,0,0,0))
draw_branch(bl, int(W*0.545), H+22, -math.pi/2-0.28, 155, 5, 3.0)
draw_branch(bl, int(W*0.895), H+15, -math.pi/2+0.18, 120, 4, 2.2)
draw_branch(bl, int(W*0.405), H+10, -math.pi/2-0.50, 102, 4, 2.0)

ll = Image.new("RGBA",(W,H),(0,0,0,0))
draw_lotus(ll, int(W*0.080), int(H*0.845), 52)
draw_lotus(ll, int(W*0.963), int(H*0.140), 40)

canvas = Image.alpha_composite(canvas, ml)
canvas = Image.alpha_composite(canvas, sm)
canvas = Image.alpha_composite(canvas, bl)
canvas = Image.alpha_composite(canvas, ll)
front = canvas.convert("RGB")
fd = ImageDraw.Draw(front)

gold_line(fd, H*0.340, W*0.097, W*0.538)
gold_line(fd, H*0.692, W*0.097, W*0.538)

fn  = ImageFont.truetype(NOTO_S,  92, index=0)
ft  = ImageFont.truetype(NOTO_S,  22, index=0)
fb  = ImageFont.truetype(NOTO_S,  28, index=0)
fen = ImageFont.truetype(os.path.join(FONTS,"CrimsonPro-Italic.ttf"), 22)

nx=int(W*0.098); ny=int(H*0.278)
for ox,oy in [(-2,-2),(2,-2),(-2,2),(2,2),(0,-3),(0,3),(-3,0),(3,0)]:
    fd.text((nx+ox,ny+oy),"黃 彥 萼",font=fn,fill=GOLDL+(40,))
fd.text((nx+2,ny+2),"黃 彥 萼",font=fn,fill=(95,70,35,100))
fd.text((nx,ny),"黃 彥 萼",font=fn,fill=GOLDD)

for i,t in enumerate(["藝術療癒引導師","心語能量美學創意總監","財團法人中華綜合發展研究院文創所所長"]):
    fd.text((int(W*0.100),int(H*0.503)+i*34), t, font=ft, fill=BROWN)

fd.text((int(W*0.100),int(H*0.754)),"心語能量美學",font=fb,fill=GOLDD)
fd.text((int(W*0.100),int(H*0.830)),"— INNER LIGHT AESTHETICS —",font=fen,fill=GOLD)

front.save(os.path.join(SCRATCHPAD,"front_final.png"))
print("Front saved.")

# ============================================================
# BACK
# ============================================================
print("Rendering back...")
bg2 = make_bg(W, H, "back")
canvas2 = bg2.convert("RGBA")

ml2 = Image.new("RGBA",(W,H),(0,0,0,0))
draw_mandala(ml2, int(W*0.722), int(H*0.530), 265, 0.36)

sm2 = Image.new("RGBA",(W,H),(0,0,0,0))
draw_mandala(sm2, int(W*0.972), int(H*0.860), 88, 0.48)

bl2 = Image.new("RGBA",(W,H),(0,0,0,0))
draw_branch(bl2, int(W*0.052), H+14, -math.pi/2+0.32, 136, 4, 2.5)

ll2 = Image.new("RGBA",(W,H),(0,0,0,0))
draw_lotus(ll2, int(W*0.874), int(H*0.112), 58)

canvas2 = Image.alpha_composite(canvas2, ml2)
canvas2 = Image.alpha_composite(canvas2, sm2)
canvas2 = Image.alpha_composite(canvas2, bl2)
canvas2 = Image.alpha_composite(canvas2, ll2)
back = canvas2.convert("RGB")
bd = ImageDraw.Draw(back)

fbc  = ImageFont.truetype(NOTO_SA, 21, index=0)
fbb  = ImageFont.truetype(NOTO_S,  24, index=0)
fenb = ImageFont.truetype(os.path.join(FONTS,"CrimsonPro-Italic.ttf"), 18)
flb  = ImageFont.truetype(os.path.join(FONTS,"CrimsonPro-Regular.ttf"), 15)

bd.text((int(W*0.102),int(H*0.078)),"心語能量美學",font=fbb,fill=GOLDD)
bd.text((int(W*0.102),int(H*0.152)),"— INNER LIGHT AESTHETICS —",font=fenb,fill=GOLD)
gold_line(bd, H*0.224, W*0.097, W*0.628)

contacts = [
    ("phone","0909-756-762"),
    ("pin","台中市西屯區中康街7號（瑞恩悦璃館）"),
    ("email","serina168@gmail.com"),
    ("line","Line ID：serina99"),
    ("fb","臉書：心語能量美學館"),
    ("ig","IG：colors_cantell"),
]
icx=int(W*0.113); tx2=int(W*0.167); sy2=int(H*0.258); sp=57

for i,(kind,info) in enumerate(contacts):
    y_=sy2+i*sp; cy_=y_+13
    contact_icon(bd, icx, cy_, 14, kind)
    bd.text((tx2,y_), info, font=fbc, fill=BROWN)

# QR boxes
qr_y=int(H*0.790); qsz=85
qr_en = ["LINE","FB Page","colors_cantell","Website"]
for i,lbl in enumerate(qr_en):
    qx=int(W*0.097)+i*(qsz+18)
    bd.rectangle([qx-2,qr_y-2,qx+qsz+2,qr_y+qsz+2], fill=GOLDL+(50,),outline=GOLD+(145,),width=1)
    bd.rectangle([qx,qr_y,qx+qsz,qr_y+qsz], fill=(255,255,255,245))
    cell=8; mg=7
    pat=[[1,1,1,0,1,1,1],[1,0,1,0,1,0,1],[1,0,1,0,1,0,1],[0,0,0,0,0,0,0],
         [1,0,1,0,1,0,1],[1,0,1,0,1,0,1],[1,1,1,0,1,1,1]]
    for row in range(9):
        for col in range(9):
            dx=qx+mg+col*cell; dy=qr_y+mg+row*cell
            tl=row<3 and col<3; tr=row<3 and col>5; bl3=row>5 and col<3
            if tl or tr or bl3:
                shade=BROWN
            elif row<7 and col<7 and pat[row%7][col%7]:
                shade=BROWN+(150,)
            else: continue
            bd.rectangle([dx+1,dy+1,dx+cell-2,dy+cell-2], fill=shade)
    bd.text((qx+4,qr_y+qsz+5), lbl, font=flb, fill=BROWN)

back.save(os.path.join(SCRATCHPAD,"back_final.png"))
print("Back saved.")
print("Done.")
