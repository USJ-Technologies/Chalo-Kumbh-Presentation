#!/usr/bin/env python3
"""
Chalo Kumbh — Ardh Kumbh Mela 2027  ·  Presentation Generator
Generates a 14-slide widescreen .pptx using python-pptx.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── paths ──────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
IMG  = os.path.join(BASE, "ppt screenshots")
OUT  = os.path.join(BASE, "Chalo Kumbh — Ardh Kumbh Mela 2027.pptx")

def img(name):
    return os.path.join(IMG, name)

# ── colours ────────────────────────────────────────────────────────────
NAVY      = RGBColor(0x1B, 0x2D, 0x5B)
SAFFRON   = RGBColor(0xD4, 0x6A, 0x0A)
GOLD      = RGBColor(0xF5, 0xB9, 0x42)
CREAM     = RGBColor(0xFD, 0xFA, 0xF4)
OFF_WHITE = RGBColor(0xF5, 0xF0, 0xE8)
TEXT_CLR   = RGBColor(0x1A, 0x1A, 0x1A)
MUTED     = RGBColor(0x5A, 0x5A, 0x5A)
GREEN     = RGBColor(0x2D, 0x6A, 0x3F)
GREEN_BG  = RGBColor(0xEB, 0xF5, 0xEF)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
SAFFRON_BG = RGBColor(0xFD, 0xF0, 0xE0)
LIGHT_BLUE_BG = RGBColor(0xE0, 0xEE, 0xF8)

# ── dimensions ─────────────────────────────────────────────────────────
SLD_W = Inches(13.333)
SLD_H = Inches(7.5)
HDR_H = Inches(0.55)
FTR_H = Inches(0.35)

# ── helpers ────────────────────────────────────────────────────────────

def set_slide_bg(slide, colour):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour

def add_rect(slide, left, top, width, height, fill_color, outline=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    shp.line.fill.background()
    if outline:
        shp.line.color.rgb = fill_color
    return shp

def add_textbox(slide, left, top, width, height):
    return slide.shapes.add_textbox(left, top, width, height)

def set_text(tf, text, font_name="Calibri", size=14, bold=False, italic=False,
             color=TEXT_CLR, align=PP_ALIGN.LEFT):
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.alignment = align
    return p

def add_para(tf, text, font_name="Calibri", size=14, bold=False, italic=False,
             color=TEXT_CLR, align=PP_ALIGN.LEFT, space_before=0, space_after=0):
    p = tf.add_paragraph()
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    if space_after:
        p.space_after = Pt(space_after)
    return p

def add_bullet_feature(tf, dot_color, title, desc, size=13):
    """Add a bullet line:  ● Title — Desc"""
    p = tf.add_paragraph()
    p.space_before = Pt(4)
    run_dot = p.add_run()
    run_dot.text = "● "
    run_dot.font.size = Pt(size)
    run_dot.font.color.rgb = dot_color
    run_dot.font.name = "Calibri"
    run_title = p.add_run()
    run_title.text = title
    run_title.font.size = Pt(size)
    run_title.font.bold = True
    run_title.font.color.rgb = TEXT_CLR
    run_title.font.name = "Calibri"
    if desc:
        run_desc = p.add_run()
        run_desc.text = f" — {desc}"
        run_desc.font.size = Pt(size)
        run_desc.font.color.rgb = MUTED
        run_desc.font.name = "Calibri"

def header_bar(slide):
    """Navy top bar with 'Chalo Kumbh' branding."""
    add_rect(slide, 0, 0, SLD_W, HDR_H, NAVY)
    tb = add_textbox(slide, Inches(0.5), 0, Inches(4), HDR_H)
    tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tb.text_frame, "Chalo Kumbh", "Georgia", 16, True, False, WHITE)

def footer_bar(slide, text="Chalo Kumbh  ·  Ardh Kumbh Mela 2027  ·  chalokumbh.com"):
    top = SLD_H - FTR_H
    add_rect(slide, 0, top, SLD_W, FTR_H, NAVY)
    tb = add_textbox(slide, Inches(0.5), top, SLD_W - Inches(1), FTR_H)
    tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tb.text_frame, text, "Calibri", 9, False, False, WHITE, PP_ALIGN.CENTER)

def section_label(slide, left, top, text):
    tb = add_textbox(slide, left, top, Inches(5), Inches(0.35))
    set_text(tb.text_frame, text, "Calibri", 11, True, False, SAFFRON)

def slide_title(slide, left, top, text, width=Inches(5.5)):
    tb = add_textbox(slide, left, top, width, Inches(0.55))
    set_text(tb.text_frame, text, "Georgia", 24, True, False, NAVY)

def accent_bar(slide, left, top, color=SAFFRON, width=Inches(0.8)):
    add_rect(slide, left, top, width, Inches(0.05), color)

def body_text(slide, left, top, text, width=Inches(5.2), size=13):
    tb = add_textbox(slide, left, top, width, Inches(1.2))
    tb.text_frame.word_wrap = True
    set_text(tb.text_frame, text, "Calibri", size, False, False, MUTED)
    return tb

def add_screenshot(slide, filename, left, top, width=None, height=None):
    path = img(filename)
    if not os.path.exists(path):
        print(f"  ⚠ Missing: {path}")
        return None
    kwargs = {"left": left, "top": top}
    if width:
        kwargs["width"] = width
    if height:
        kwargs["height"] = height
    return slide.shapes.add_picture(path, **kwargs)

def stat_box(slide, left, top, value, label, bg=NAVY, val_color=GOLD, lbl_color=WHITE,
             width=Inches(2.6), height=Inches(1.1)):
    add_rect(slide, left, top, width, height, bg)
    tb = add_textbox(slide, left, top + Inches(0.1), width, Inches(0.5))
    set_text(tb.text_frame, value, "Georgia", 26, True, False, val_color, PP_ALIGN.CENTER)
    tb2 = add_textbox(slide, left, top + Inches(0.55), width, Inches(0.35))
    set_text(tb2.text_frame, label, "Calibri", 11, False, False, lbl_color, PP_ALIGN.CENTER)

def pillar_box(slide, left, top, title, width=Inches(2.2), height=Inches(0.9)):
    shp = add_rect(slide, left, top, width, height, GREEN_BG)
    shp.text_frame.word_wrap = True
    shp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(shp.text_frame, title, "Calibri", 11, True, False, GREEN, PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════
# SLIDES
# ═══════════════════════════════════════════════════════════════════════

prs = Presentation()
prs.slide_width  = SLD_W
prs.slide_height = SLD_H
blank = prs.slide_layouts[6]  # blank layout

# ───────────────────────────── SLIDE 1 — COVER ─────────────────────────
s1 = prs.slides.add_slide(blank)
# hero background
add_screenshot(s1, "Homepage header.png", 0, 0, width=SLD_W, height=SLD_H)
# dark navy overlay (solid, no transparency)
add_rect(s1, 0, 0, SLD_W, SLD_H, NAVY)
# Logo
logo_w = Inches(1.8)
logo_left = (SLD_W - logo_w) // 2
add_screenshot(s1, "Kumbh Mela Logo.jpg", logo_left, Inches(1.2), width=logo_w)
# subtitle
tb = add_textbox(s1, 0, Inches(3.1), SLD_W, Inches(0.4))
set_text(tb.text_frame, "ARDH KUMBH MELA 2027  ·  HARIDWAR", "Calibri", 13, True, False,
         GOLD, PP_ALIGN.CENTER)
# title
tb = add_textbox(s1, 0, Inches(3.55), SLD_W, Inches(1.0))
set_text(tb.text_frame, "Chalo Kumbh", "Georgia", 52, True, False, WHITE, PP_ALIGN.CENTER)
# tagline
tb = add_textbox(s1, 0, Inches(4.55), SLD_W, Inches(0.4))
set_text(tb.text_frame, "Official Digital Partner — Ardh Kumbh Mela 2027, Haridwar",
         "Georgia", 15, False, True, RGBColor(0xCC, 0xCC, 0xCC), PP_ALIGN.CENTER)
# saffron accent bar
accent_bar(s1, (SLD_W - Inches(1.5)) // 2, Inches(5.05), SAFFRON, Inches(1.5))
# 3 stats
stats = [("150M+", "Projected Pilgrims"), ("45", "Days of Divinity"), ("Jan–Apr", "2027")]
for i, (val, lbl) in enumerate(stats):
    x = Inches(2.5) + Inches(3.0) * i
    tb = add_textbox(s1, x, Inches(5.4), Inches(2.5), Inches(0.45))
    set_text(tb.text_frame, val, "Georgia", 26, True, False, GOLD, PP_ALIGN.CENTER)
    tb2 = add_textbox(s1, x, Inches(5.85), Inches(2.5), Inches(0.3))
    set_text(tb2.text_frame, lbl, "Calibri", 11, False, False, WHITE, PP_ALIGN.CENTER)

# ───────────────────────────── SLIDE 2 — PLATFORM OVERVIEW ─────────────
s2 = prs.slides.add_slide(blank)
set_slide_bg(s2, CREAM)
header_bar(s2)
footer_bar(s2)

LM = Inches(0.7)  # left margin for content slides
TOP_START = Inches(0.9)

section_label(s2, LM, TOP_START, "WHAT WE BUILT")
slide_title(s2, LM, Inches(1.25), "The Official Digital Platform for Ardh Kumbh 2027")
accent_bar(s2, LM, Inches(1.85))
body_text(s2, LM, Inches(2.0),
          "chalokumbh.com is a comprehensive web platform designed to digitise "
          "and elevate the Ardh Kumbh Mela 2027 experience — from ritual booking "
          "to environmental stewardship.")

tb = add_textbox(s2, LM, Inches(2.9), Inches(5.5), Inches(3.5))
tf = tb.text_frame
tf.word_wrap = True
features = [
    ("Online Puja & E-Seva", "Book sacred rituals and receive Prasad delivery worldwide"),
    ("Live Darshan Streaming", "Watch ceremonies live from Har Ki Pauri via WebRTC"),
    ("Hotel & Stay Booking", "Curated accommodations with real-time availability"),
    ("Green Haridwar", "Environmental monitoring, volunteer drives & gamified challenges"),
    ("Multi-Portal System", "Dedicated dashboards for Pandits, Hotels, and Admin"),
]
set_text(tf, "", "Calibri", 13)
for title, desc in features:
    add_bullet_feature(tf, SAFFRON, title, desc)

add_screenshot(s2, "Homepage Screenshot.png", Inches(6.8), Inches(1.0),
               width=Inches(5.8), height=Inches(5.5))

# ───────────────────────────── SLIDE 3 — ONLINE PUJA ──────────────────
s3 = prs.slides.add_slide(blank)
set_slide_bg(s3, CREAM)
header_bar(s3)
footer_bar(s3)

add_screenshot(s3, "Online Puja Booking Page.png", Inches(0.5), Inches(1.0),
               width=Inches(5.5), height=Inches(5.8))

RX = Inches(6.5)
section_label(s3, RX, TOP_START, "ONLINE PUJA & E-SEVA")
slide_title(s3, RX, Inches(1.25), "Book Sacred Rituals from Anywhere in the World", Inches(6.2))
accent_bar(s3, RX, Inches(1.85))
body_text(s3, RX, Inches(2.0),
          "Devotees can browse, customise, and book authentic Vedic rituals performed "
          "by verified pandits at Haridwar's sacred ghats — with live-streamed ceremonies "
          "and Prasad shipped to their door.", Inches(6.0))

# 4-step flow
flow_steps = ["Choose Ritual", "Select Date", "Pick Pandit", "Join Live"]
for i, step in enumerate(flow_steps):
    x = RX + Inches(1.5) * i
    shp = add_rect(s3, x, Inches(3.3), Inches(1.35), Inches(0.5), SAFFRON_BG)
    shp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(shp.text_frame, f"{i+1}. {step}", "Calibri", 10, True, False, SAFFRON, PP_ALIGN.CENTER)

tb = add_textbox(s3, RX, Inches(4.0), Inches(6.0), Inches(2.5))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [("50+ Rituals Available", "Ganga Aarti, Rudrabhishek, Satyanarayan & more"),
             ("Prasad Delivery", "Sacred offerings shipped across India"),
             ("Personalised Sankalp", "Custom prayer invocations with devotee's name & gotra")]:
    add_bullet_feature(tf, SAFFRON, t, d)

# ───────────────────────────── SLIDE 4 — LIVE DARSHAN ─────────────────
s4 = prs.slides.add_slide(blank)
set_slide_bg(s4, CREAM)
header_bar(s4)
footer_bar(s4)

section_label(s4, LM, TOP_START, "LIVE DARSHAN  ·  REAL-TIME STREAMING")
slide_title(s4, LM, Inches(1.25), "Watch Sacred Ceremonies Live — Worldwide")
accent_bar(s4, LM, Inches(1.85))

tb = add_textbox(s4, LM, Inches(2.05), Inches(5.5), Inches(3.0))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("Homepage Integration", "Live stream embedded directly on the landing page"),
    ("No Login Required", "Open access for millions of concurrent viewers"),
    ("QR Code Access", "Scan-to-watch for on-ground pilgrims"),
    ("Admin Stream Control", "Start, stop, and manage streams from the Company Portal"),
]:
    add_bullet_feature(tf, SAFFRON, t, d)

# Tech box
tech_box = add_rect(s4, LM, Inches(4.4), Inches(5.5), Inches(0.65), OFF_WHITE)
tb_tech = add_textbox(s4, LM + Inches(0.15), Inches(4.45), Inches(5.2), Inches(0.55))
tf_tech = tb_tech.text_frame
tf_tech.word_wrap = True
set_text(tf_tech, "Technology:  ", "Calibri", 11, True, False, NAVY)
r = tf_tech.paragraphs[0].add_run()
r.text = "LiveKit Cloud  ·  WebRTC  ·  Viewer token auth  ·  Multi-stream support"
r.font.size = Pt(11)
r.font.name = "Calibri"
r.font.color.rgb = MUTED

add_screenshot(s4, "Live Darshan Section Screenshot.png", Inches(6.8), Inches(1.0),
               width=Inches(5.8), height=Inches(5.5))

# ───────────────────────────── SLIDE 5 — HOTELS ──────────────────────
s5 = prs.slides.add_slide(blank)
set_slide_bg(s5, CREAM)
header_bar(s5)
footer_bar(s5)

add_screenshot(s5, "Hotels Listing Page.png", Inches(0.5), Inches(1.0),
               width=Inches(5.5), height=Inches(5.8))

section_label(s5, RX, TOP_START, "HOTEL & ACCOMMODATION BOOKING")
slide_title(s5, RX, Inches(1.25), "Handpicked Stays Near Har Ki Pauri", Inches(6.2))
accent_bar(s5, RX, Inches(1.85))

tb = add_textbox(s5, RX, Inches(2.05), Inches(6.0), Inches(4.0))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("Smart Filtering", "Search by location, price, amenities, and star rating"),
    ("Three Price Tiers", "Budget, Standard, and Premium categories for every pilgrim"),
    ("Property Types", "Hotels, Dharamshalas, Ashrams, and Boutique stays"),
    ("Commission Management", "Transparent 10-18% tiered commission model"),
    ("Real-Time Availability", "Live room inventory with instant confirmation"),
]:
    add_bullet_feature(tf, SAFFRON, t, d)

# ───────────────────────────── SLIDE 6 — GREEN HARIDWAR OVERVIEW ──────
s6 = prs.slides.add_slide(blank)
set_slide_bg(s6, CREAM)
header_bar(s6)
footer_bar(s6)

section_label(s6, LM, TOP_START, "ENVIRONMENTAL INITIATIVE")
tb = add_textbox(s6, LM, Inches(1.25), Inches(8), Inches(0.6))
tf = tb.text_frame
set_text(tf, "Green Haridwar ", "Georgia", 26, True, False, GREEN)
r = tf.paragraphs[0].add_run()
r.text = "🌿"
r.font.size = Pt(26)
accent_bar(s6, LM, Inches(1.85), GREEN)

# 5 pillar boxes
pillars = ["Social Media\nCampaigns", "Volunteer\nDrives", "Green\nChallenges",
           "Issue\nReporting", "Event\nAwareness"]
pw = Inches(2.2)
gap = Inches(0.25)
for i, p in enumerate(pillars):
    x = LM + (pw + gap) * i
    pillar_box(s6, x, Inches(2.1), p)

# Screenshot on left
add_screenshot(s6, "Green Haridwar — Hero Section.png", LM, Inches(3.3),
               width=Inches(6.5), height=Inches(3.3))

# Partnering With box
partner_box = add_rect(s6, Inches(7.8), Inches(3.3), Inches(4.8), Inches(3.3), GREEN_BG)
tb = add_textbox(s6, Inches(8.0), Inches(3.45), Inches(4.4), Inches(0.35))
set_text(tb.text_frame, "Partnering With", "Georgia", 16, True, False, GREEN, PP_ALIGN.CENTER)

partners = [
    "Uttarakhand State Government",
    "National Mission for Clean Ganga (NMCG)",
    "Haridwar Municipal Corporation",
    "Local Environmental NGOs",
]
for i, partner in enumerate(partners):
    tb = add_textbox(s6, Inches(8.2), Inches(3.95) + Inches(0.4) * i,
                     Inches(4.2), Inches(0.35))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run_dot = p.add_run()
    run_dot.text = "● "
    run_dot.font.size = Pt(12)
    run_dot.font.color.rgb = GREEN
    run_dot.font.name = "Calibri"
    run_txt = p.add_run()
    run_txt.text = partner
    run_txt.font.size = Pt(12)
    run_txt.font.color.rgb = TEXT_CLR
    run_txt.font.name = "Calibri"

# ───────────────────────────── SLIDE 7 — GREEN REPORTING ─────────────
s7 = prs.slides.add_slide(blank)
set_slide_bg(s7, CREAM)
header_bar(s7)
footer_bar(s7)

# 2 screenshots stacked left
add_screenshot(s7, "Haridwar Interactive Map with Report Pins.png",
               Inches(0.5), Inches(1.0), width=Inches(5.3), height=Inches(3.0))
add_screenshot(s7, "Issue Report Submission Form.png",
               Inches(0.5), Inches(4.15), width=Inches(5.3), height=Inches(2.8))

section_label(s7, RX, TOP_START, "REAL-TIME ISSUE REPORTING")
slide_title(s7, RX, Inches(1.25), "Citizen-Powered Environmental Monitoring", Inches(6.2))
accent_bar(s7, RX, Inches(1.85), GREEN)

tb = add_textbox(s7, RX, Inches(2.05), Inches(6.0), Inches(3.0))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("4 Issue Categories", "Pollution, Waste Dumping, Water Quality, Encroachment"),
    ("Location Picker", "Interactive map with GPS pin-drop for precise reporting"),
    ("Photo Upload", "Visual evidence attached to every report"),
    ("Status Tracking", "Real-time updates: Submitted → Under Review → Resolved"),
]:
    add_bullet_feature(tf, GREEN, t, d)

# Government benefit box
gov_box = add_rect(s7, RX, Inches(4.5), Inches(6.0), Inches(1.0), GREEN_BG)
tb = add_textbox(s7, RX + Inches(0.2), Inches(4.55), Inches(5.6), Inches(0.9))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "For Government & Authorities", "Georgia", 13, True, False, GREEN)
add_para(tf, "Real-time environmental data dashboard enabling faster response times, "
             "resource allocation, and evidence-based policy decisions during the Mela.",
         "Calibri", 11, False, False, MUTED, space_before=4)

# ───────────────────────────── SLIDE 8 — GREEN CHALLENGES ─────────────
s8 = prs.slides.add_slide(blank)
set_slide_bg(s8, CREAM)
header_bar(s8)
footer_bar(s8)

section_label(s8, LM, TOP_START, "GREEN CHALLENGES & VOLUNTEER ENGAGEMENT")
slide_title(s8, LM, Inches(1.25), "Gamified Sustainability for Mass Participation")
accent_bar(s8, LM, Inches(1.85), GREEN)

tb = add_textbox(s8, LM, Inches(2.05), Inches(5.5), Inches(2.5))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("Photo-Verified Challenges", "Participants submit evidence of eco-actions for approval"),
    ("Leaderboard & Rewards", "Top contributors earn certificates and recognition"),
    ("Volunteer Registration", "One-click sign-up for clean-up drives and tree planting"),
]:
    add_bullet_feature(tf, GREEN, t, d)

# stat boxes
stat_box(s8, LM, Inches(4.3), "2,400+", "Registered Volunteers", GREEN_BG, GREEN, TEXT_CLR,
         Inches(2.5), Inches(0.9))
stat_box(s8, LM + Inches(2.8), Inches(4.3), "12", "Active Challenges", GREEN_BG, GREEN, TEXT_CLR,
         Inches(2.5), Inches(0.9))

# screenshots right
add_screenshot(s8, "Green Challenges Listing.png", Inches(6.8), Inches(1.0),
               width=Inches(5.8), height=Inches(3.0))
add_screenshot(s8, "Recent Environmental Reports List.png", Inches(6.8), Inches(4.15),
               width=Inches(5.8), height=Inches(2.8))

# ───────────────────────────── SLIDE 9 — PANDIT DASHBOARD ────────────
s9 = prs.slides.add_slide(blank)
set_slide_bg(s9, CREAM)
header_bar(s9)
footer_bar(s9)

section_label(s9, LM, TOP_START, "PANDIT PORTAL")
slide_title(s9, LM, Inches(1.25), "Dedicated Dashboard for Registered Priests")
accent_bar(s9, LM, Inches(1.85))

tb = add_textbox(s9, LM, Inches(2.05), Inches(5.5), Inches(3.5))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("Booking Management", "Accept, decline, and track incoming puja requests"),
    ("Weekly Availability", "Set available dates and time slots for ceremonies"),
    ("Live Session Access", "One-click join for live-streamed rituals via LiveKit"),
    ("Approval Workflow", "Profile verification by admin before going live on platform"),
]:
    add_bullet_feature(tf, SAFFRON, t, d)

add_screenshot(s9, "Pandit Dashboard — Main Overview.png", Inches(6.8), Inches(1.0),
               width=Inches(5.8), height=Inches(5.5))

# ───────────────────────────── SLIDE 10 — HOTEL DASHBOARD ────────────
s10 = prs.slides.add_slide(blank)
set_slide_bg(s10, CREAM)
header_bar(s10)
footer_bar(s10)

add_screenshot(s10, "Hotel Dashboard — Main Overview.png", Inches(0.5), Inches(1.0),
               width=Inches(5.5), height=Inches(5.8))

section_label(s10, RX, TOP_START, "HOTEL PARTNER PORTAL")
slide_title(s10, RX, Inches(1.25), "Self-Service Portal for Hotel Partners", Inches(6.2))
accent_bar(s10, RX, Inches(1.85))

tb = add_textbox(s10, RX, Inches(2.05), Inches(6.0), Inches(3.5))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("Room Type Management", "Add and configure room categories, pricing, and capacity"),
    ("Availability & Date Blocking", "Calendar-based inventory control with blackout dates"),
    ("Photo Gallery", "Upload property and room images to attract bookings"),
    ("Earnings Dashboard", "Track revenue, commissions, and payout history"),
]:
    add_bullet_feature(tf, SAFFRON, t, d)

# ───────────────────────────── SLIDE 11 — COMPANY PORTAL ─────────────
s11 = prs.slides.add_slide(blank)
set_slide_bg(s11, CREAM)
header_bar(s11)
footer_bar(s11)

add_screenshot(s11, "Company Portal — Main Dashboard.png", Inches(0.5), Inches(1.0),
               width=Inches(5.3), height=Inches(3.0))
add_screenshot(s11, "Bookings Management Table.png", Inches(0.5), Inches(4.15),
               width=Inches(5.3), height=Inches(2.8))

section_label(s11, RX, TOP_START, "ADMIN  ·  COMPANY PORTAL")
slide_title(s11, RX, Inches(1.25), "Complete Platform Administration", Inches(6.2))
accent_bar(s11, RX, Inches(1.85))

tb = add_textbox(s11, RX, Inches(2.05), Inches(6.0), Inches(4.0))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 13)
for t, d in [
    ("Platform Analytics", "Real-time dashboards with booking volume, revenue & user metrics"),
    ("Pandit Approvals", "Review, verify, and approve pandit registrations"),
    ("Livestream Management", "Control live darshan streams — start, stop, multi-camera"),
    ("Marketing & Content", "Manage homepage banners, promotions, and notifications"),
    ("Hotel Earnings", "Commission tracking and payout management for all properties"),
]:
    add_bullet_feature(tf, SAFFRON, t, d)

# ───────────────────────────── SLIDE 12 — SERVICES OVERVIEW ──────────
s12 = prs.slides.add_slide(blank)
set_slide_bg(s12, CREAM)
header_bar(s12)
footer_bar(s12)

section_label(s12, LM, TOP_START, "PILGRIMAGE SERVICES")
slide_title(s12, LM, Inches(1.25),
            "Everything a Pilgrim Needs — In One Platform", Inches(11))
accent_bar(s12, LM, Inches(1.85))

# Table
tbl_data = [
    ("Service", "Description", "Key Features"),
    ("Rituals & Pujas", "Book authentic Vedic ceremonies online",
     "50+ rituals, verified pandits, Prasad delivery"),
    ("VIP Ghat Access", "Reserved spots at prime bathing locations",
     "Priority entry, guided experience, safety escort"),
    ("Premium Stays", "Curated hotels and heritage properties",
     "Budget to luxury, instant booking, real reviews"),
    ("Satvik Meals", "Traditional vegetarian meal plans",
     "Pre-booked thalis, dietary customisation"),
    ("Travel Assistance", "End-to-end logistics and transport",
     "Airport transfers, local cabs, route planning"),
    ("24×7 Concierge", "Round-the-clock pilgrim support",
     "WhatsApp, call centre, multilingual agents"),
    ("Veda Hydration", "Branded purified water for pilgrims",
     "Eco-friendly packaging, ghat distribution"),
]

rows, cols = len(tbl_data), 3
tbl_w = Inches(11.8)
col_widths = [Inches(2.2), Inches(4.8), Inches(4.8)]
tbl = s12.shapes.add_table(rows, cols, LM, Inches(2.1), tbl_w, Inches(4.6)).table

for ci, cw in enumerate(col_widths):
    tbl.columns[ci].width = cw

for ri, row_data in enumerate(tbl_data):
    for ci, cell_text in enumerate(row_data):
        cell = tbl.cell(ri, ci)
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        p.text = cell_text
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE if ri == 0 else TEXT_CLR
        p.font.bold = ri == 0 or ci == 0
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        if ri == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
        elif ri % 2 == 1:
            cell.fill.solid()
            cell.fill.fore_color.rgb = CREAM
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = OFF_WHITE

# ───────────────────────────── SLIDE 13 — IMPACT ─────────────────────
s13 = prs.slides.add_slide(blank)
set_slide_bg(s13, CREAM)
header_bar(s13)
footer_bar(s13)

section_label(s13, LM, TOP_START, "WHY IT MATTERS")
slide_title(s13, LM, Inches(1.25),
            "Digital Infrastructure for India's Largest Gathering", Inches(11))
accent_bar(s13, LM, Inches(1.85))

# 3 stat boxes
stat_labels = [("150M+", "Projected Pilgrims"), ("24×7", "Support Coverage"),
               ("3", "Dedicated Portals")]
for i, (v, l) in enumerate(stat_labels):
    x = LM + Inches(4.0) * i
    stat_box(s13, x, Inches(2.2), v, l, NAVY, GOLD, WHITE, Inches(3.6), Inches(1.3))

# Digital India box
di_box = add_rect(s13, LM, Inches(3.9), Inches(5.8), Inches(2.6), SAFFRON_BG)
tb = add_textbox(s13, LM + Inches(0.2), Inches(4.0), Inches(5.4), Inches(0.35))
set_text(tb.text_frame, "🇮🇳  Digital India Alignment", "Georgia", 15, True, False, SAFFRON)
tb = add_textbox(s13, LM + Inches(0.2), Inches(4.45), Inches(5.4), Inches(1.8))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 12)
for t in [
    "Paperless booking and digital payments for all services",
    "Real-time data dashboards for government authorities",
    "Scalable cloud infrastructure for 150M+ pilgrims",
]:
    add_bullet_feature(tf, SAFFRON, t, "", 11)

# Swachh Bharat box
sb_box = add_rect(s13, Inches(6.9), Inches(3.9), Inches(5.8), Inches(2.6), LIGHT_BLUE_BG)
tb = add_textbox(s13, Inches(7.1), Inches(4.0), Inches(5.4), Inches(0.35))
set_text(tb.text_frame, "🌊  Swachh Bharat Alignment", "Georgia", 15, True, False,
         RGBColor(0x1A, 0x6B, 0x8A))
tb = add_textbox(s13, Inches(7.1), Inches(4.45), Inches(5.4), Inches(1.8))
tf = tb.text_frame
tf.word_wrap = True
set_text(tf, "", "Calibri", 12)
for t in [
    "Green Haridwar initiative for environmental monitoring",
    "Citizen-powered issue reporting with real-time tracking",
    "Gamified volunteer engagement for cleanliness drives",
]:
    add_bullet_feature(tf, RGBColor(0x1A, 0x6B, 0x8A), t, "", 11)

# ───────────────────────────── SLIDE 14 — CLOSING ────────────────────
s14 = prs.slides.add_slide(blank)
# hero background
add_screenshot(s14, "Homepage header.png", 0, 0, width=SLD_W, height=SLD_H)
# dark navy overlay
add_rect(s14, 0, 0, SLD_W, SLD_H, NAVY)

# Logo
add_screenshot(s14, "Kumbh Mela Logo.jpg", logo_left, Inches(1.4), width=logo_w)

# Chant
tb = add_textbox(s14, 0, Inches(3.2), SLD_W, Inches(0.4))
set_text(tb.text_frame, "HAR HAR GANGE  ·  JAI GANGA MAIYA", "Calibri", 13, True, False,
         GOLD, PP_ALIGN.CENTER)

# Title
tb = add_textbox(s14, Inches(1.5), Inches(3.7), SLD_W - Inches(3), Inches(1.0))
set_text(tb.text_frame, "Together, Let Us Make Kumbh 2027 Extraordinary",
         "Georgia", 36, True, False, WHITE, PP_ALIGN.CENTER)

# Saffron bar
accent_bar(s14, (SLD_W - Inches(1.5)) // 2, Inches(4.8), SAFFRON, Inches(1.5))

# Contact
contacts = [
    ("Platform", "chalokumbh.com"),
    ("Built By", "USJ Technologies"),
    ("Contact", "techteam@usjtechnologies.com"),
]
for i, (label, value) in enumerate(contacts):
    x = Inches(2.0) + Inches(3.5) * i
    tb = add_textbox(s14, x, Inches(5.2), Inches(3.2), Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run_lbl = p.add_run()
    run_lbl.text = f"{label}:  "
    run_lbl.font.size = Pt(12)
    run_lbl.font.bold = True
    run_lbl.font.name = "Calibri"
    run_lbl.font.color.rgb = GOLD
    run_val = p.add_run()
    run_val.text = value
    run_val.font.size = Pt(12)
    run_val.font.name = "Calibri"
    run_val.font.color.rgb = WHITE

# ═══════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════
prs.save(OUT)
print(f"✅  Presentation saved → {OUT}")
print(f"   Slides: {len(prs.slides)}")
