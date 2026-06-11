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

print('SAVED placeholder - see full file in repo')
