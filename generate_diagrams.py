# -*- coding: utf-8 -*-
"""產生曼陀羅圖案元素／結構示意圖（PIL 線稿風，配合簡報柔色系）。"""
import math, os
from PIL import Image, ImageDraw

OUT = "/home/user/mandala/l1/diagrams"
os.makedirs(OUT, exist_ok=True)

# 柔色系（與簡報一致）
CREAM = (0xFB, 0xF6, 0xEF, 255)
ROSE  = (0xC9, 0x7B, 0x84, 255)
ROSE_L= (0xE6, 0xA9, 0xAE, 255)
ROSE_D= (0xA9, 0x5C, 0x66, 255)
GOLD  = (0xC9, 0xA9, 0x6E, 255)
GOLD_L= (0xE2, 0xCD, 0xA0, 255)
SAGE  = (0x9C, 0xA9, 0x8C, 255)
SAGE_L= (0xC3, 0xCE, 0xB6, 255)
TAUPE = (0x8A, 0x7A, 0x6A, 255)
INK   = (0x4A, 0x40, 0x3A, 255)

SS = 4               # 超取樣倍率（抗鋸齒）
SZ = 600             # 最終輸出尺寸

def canvas():
    return Image.new("RGBA", (SZ*SS, SZ*SS), CREAM)

def finish(img, name):
    img = img.resize((SZ, SZ), Image.LANCZOS)
    img.convert("RGB").save(f"{OUT}/{name}.png", quality=95)
    print("  ", name)

def petal(draw, cx, cy, length, width, ang, fill, outline, ow=3):
    """以兩段貝茲弧線構成的花瓣（水滴形），尖端朝 ang 方向。"""
    a = math.radians(ang)
    tx, ty = cx + length*math.cos(a), cy + length*math.sin(a)     # 尖端
    # 左右兩側控制點
    pa = a + math.pi/2
    lx, ly = cx + width*math.cos(pa), cy + width*math.sin(pa)
    rx, ry = cx - width*math.cos(pa), cy - width*math.sin(pa)
    mx, my = cx + length*0.5*math.cos(a), cy + length*0.5*math.sin(a)
    pts = bezier([(cx,cy),(lx,ly),(mx+ (lx-mx)*0.3, my+(ly-my)*0.3),(tx,ty)]) + \
          bezier([(tx,ty),(mx+(rx-mx)*0.3, my+(ry-my)*0.3),(rx,ry),(cx,cy)])
    draw.polygon(pts, fill=fill, outline=outline)
    # 加粗外框
    draw.line(pts+[pts[0]], fill=outline, width=ow, joint="curve")

def bezier(ctrl, n=40):
    res=[]
    for i in range(n+1):
        t=i/n; x=y=0
        m=len(ctrl)-1
        for k,(px,py) in enumerate(ctrl):
            b=math.comb(m,k)*(t**k)*((1-t)**(m-k))
            x+=b*px; y+=b*py
        res.append((x,y))
    return res

def teardrop(draw, cx, cy, length, width, ang, fill, outline, ow=3):
    """水滴：一端圓一端尖。"""
    a=math.radians(ang)
    tx,ty=cx+length*math.cos(a),cy+length*math.sin(a)
    pa=a+math.pi/2
    r=width
    bx,by=cx,cy
    lx,ly=bx+r*math.cos(pa),by+r*math.sin(pa)
    rx,ry=bx-r*math.cos(pa),by-r*math.sin(pa)
    pts=bezier([(lx,ly),(lx+(tx-lx)*0.2,ly+(ty-ly)*0.2),(tx,ty)]) + \
        bezier([(tx,ty),(rx+(tx-rx)*0.2,ry+(ty-ry)*0.2),(rx,ry)])
    # 圓底
    arc=[]
    for i in range(31):
        t=math.pi*i/30
        arc.append((bx+r*math.cos(pa+math.pi-t),by+r*math.sin(pa+math.pi-t)))
    pts=pts+arc
    draw.polygon(pts,fill=fill,outline=outline)
    draw.line(pts+[pts[0]],fill=outline,width=ow,joint="curve")

def C(cx,cy,r): return [cx-r,cy-r,cx+r,cy+r]

# ---------------------------------------------------------------
def d_dots():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    # 同心點環：三圈，外大內小漸層
    rings=[(0.72,18,ROSE),(0.5,14,GOLD),(0.28,10,SAGE)]
    for rr,dot,col in rings:
        R=c*rr; n=int(2*math.pi*R/(dot*SS*2.4))
        for i in range(n):
            a=2*math.pi*i/n
            x,y=c+R*math.cos(a),c+R*math.sin(a)
            dr.ellipse(C(x,y,dot*SS),fill=col,outline=INK,width=2)
    # 中央漸大圓點（律動）
    for i,rr in enumerate([6,11,16,22]):
        dr.ellipse(C(c,c,rr*SS),fill=(ROSE_L if i%2 else GOLD_L),outline=ROSE_D,width=3)
    finish(img,"el_dots")

def d_petals():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    n=8; L=c*0.78; W=c*0.2
    for i in range(n):
        ang=360*i/n
        petal(dr,c,c,L,W,ang,ROSE_L,ROSE_D,ow=4)
    for i in range(n):
        ang=360*i/n+360/(2*n)
        petal(dr,c,c,L*0.6,W*0.7,ang,GOLD_L,GOLD,ow=3)
    dr.ellipse(C(c,c,c*0.16),fill=GOLD,outline=ROSE_D,width=4)
    dr.ellipse(C(c,c,c*0.07),fill=CREAM,outline=ROSE_D,width=3)
    finish(img,"el_petals")

def d_leaves():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    # 中央藤蔓 + 兩側葉
    dr.line([(c,c*1.7),(c,c*0.3)],fill=SAGE,width=5*SS)
    for t in [0.35,0.55,0.75,0.95,1.2,1.4]:
        y=c*0.3+ (c*1.4)*(t/1.4)
        for s in (1,-1):
            petal(dr,c,y,c*0.34,c*0.1,(-60*s if s>0 else 180+60),  # placeholder
                  SAGE_L,SAGE,ow=3)
    finish(img,"el_leaves")

def d_leaves2():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    dr.line([(c,int(c*1.72)),(c,int(c*0.34))],fill=SAGE,width=4*SS)
    ys=[0.55,0.85,1.15,1.45]
    for k,t in enumerate(ys):
        y=int(c*t)
        petal(dr,c,y,c*0.42,c*0.12,-35,SAGE_L,SAGE,ow=3)   # 右上
        petal(dr,c,y,c*0.42,c*0.12,180+35,SAGE_L,SAGE,ow=3) # 左上
    dr.ellipse(C(c,int(c*0.34),c*0.12),fill=GOLD_L,outline=SAGE,width=3)
    finish(img,"el_leaves")

def d_teardrop():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    n=10; L=c*0.55; W=c*0.17
    base=c*0.2                      # 圓底離中心一段距離，露出水滴的圓端
    for i in range(n):
        a=math.radians(360*i/n)
        bx,by=c+base*math.cos(a),c+base*math.sin(a)
        teardrop(dr,bx,by,L,W,360*i/n,ROSE_L,ROSE_D,ow=4)
    # 中央小花
    for i in range(n):
        a=math.radians(360*i/n)
        teardrop(dr,c,c,c*0.18,c*0.07,360*i/n,GOLD_L,GOLD,ow=2)
    dr.ellipse(C(c,c,c*0.08),fill=ROSE,outline=INK,width=3)
    finish(img,"el_teardrop")

def d_lines():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    # 同心波浪線（弧線律動）+ 放射卷草，呈現「線條」流動
    for ri,(rr,col) in enumerate([(0.8,GOLD),(0.6,ROSE),(0.4,SAGE)]):
        R=c*rr; pts=[]; waves=12; amp=c*0.05
        for k in range(241):
            a=2*math.pi*k/240
            rad=R+amp*math.sin(waves*a)
            pts.append((c+rad*math.cos(a),c+rad*math.sin(a)))
        dr.line(pts,fill=col,width=3*SS,joint="curve")
    # 中央放射卷草（S 形）
    n=8
    for i in range(n):
        a=2*math.pi*i/n
        ca,sa=math.cos(a),math.sin(a)
        def rot(x,y): return (c+x*ca-y*sa, c+x*sa+y*ca)
        local=[(c*0.0,0),(c*0.14,c*0.06),(c*0.26,-c*0.04),(c*0.34,0)]
        pts=[rot(x,y) for x,y in bezier(local,30)]
        dr.line(pts,fill=ROSE_D,width=3*SS,joint="curve")
    dr.ellipse(C(c,c,c*0.08),fill=GOLD,outline=ROSE_D,width=3)
    finish(img,"el_lines")

def d_geometry():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    # 外圈：菱形環
    R=c*0.72
    n=8
    for i in range(n):
        a=2*math.pi*i/n
        x,y=c+R*math.cos(a),c+R*math.sin(a)
        s=c*0.13
        dr.polygon([(x,y-s),(x+s,y),(x,y+s),(x-s,y)],fill=SAGE_L,outline=SAGE,width=3)
    # 中圈：三角
    R2=c*0.42
    for i in range(6):
        a=2*math.pi*i/6 - math.pi/2
        x,y=c+R2*math.cos(a),c+R2*math.sin(a)
        s=c*0.14
        a2=a
        p1=(x+s*math.cos(a2),y+s*math.sin(a2))
        p2=(x+s*math.cos(a2+2.5),y+s*math.sin(a2+2.5))
        p3=(x+s*math.cos(a2-2.5),y+s*math.sin(a2-2.5))
        dr.polygon([p1,p2,p3],fill=GOLD_L,outline=GOLD,width=3)
    # 中心圓
    dr.ellipse(C(c,c,c*0.18),fill=ROSE_L,outline=ROSE_D,width=4)
    dr.ellipse(C(c,c,c*0.09),fill=CREAM,outline=ROSE_D,width=3)
    finish(img,"el_geometry")

# 結構示意：標註圓心 / 放射 / 對稱 / 層次
def d_structure():
    img=canvas(); dr=ImageDraw.Draw(img); c=SZ*SS//2
    # 層次同心圓
    for rr,col in [(0.82,SAGE_L),(0.62,GOLD_L),(0.4,ROSE_L)]:
        dr.ellipse(C(c,c,c*rr),outline=col,width=4*SS)
    # 放射線
    n=12
    for i in range(n):
        a=2*math.pi*i/n
        dr.line([(c,c),(c+c*0.82*math.cos(a),c+c*0.82*math.sin(a))],
                fill=TAUPE,width=2*SS)
    # 花瓣層
    for i in range(n):
        petal(dr,c,c,c*0.6,c*0.1,360*i/n,(0,0,0,0),ROSE,ow=3)
    # 圓心
    dr.ellipse(C(c,c,c*0.1),fill=ROSE,outline=INK,width=4)
    finish(img,"structure")

if __name__=="__main__":
    print("generating diagrams...")
    d_dots(); d_petals(); d_leaves(); d_teardrop(); d_lines(); d_geometry()
    d_structure()
    print("done ->", OUT)
