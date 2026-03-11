"""Generate the 4-slide executive summary PowerPoint for Apple Patent Explorer."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Colors
BG = RGBColor(0x1A, 0x1B, 0x1E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DIMMED = RGBColor(0x8B, 0x92, 0x9E)
GREEN = RGBColor(0x51, 0xCF, 0x66)
BLUE = RGBColor(0x33, 0x9A, 0xF0)
RED = RGBColor(0xFF, 0x6B, 0x6B)
ORANGE = RGBColor(0xFF, 0x92, 0x2B)
YELLOW = RGBColor(0xFC, 0xC4, 0x19)
CYAN = RGBColor(0x22, 0xB8, 0xCF)
PURPLE = RGBColor(0xCC, 0x5D, 0xE8)
CARD_BG = RGBColor(0x25, 0x26, 0x2B)
BORDER = RGBColor(0x2C, 0x2E, 0x33)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

def set_bg(slide):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG

def add_text(slide, left, top, width, height, text, size=14, color=WHITE, bold=False, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.alignment = align
    return txBox

def add_card(slide, left, top, width, height, label, value, color=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_BG
    shape.line.color.rgb = BORDER
    shape.line.width = Pt(1)
    add_text(slide, left + 0.15, top + 0.1, width - 0.3, 0.3, label, size=9, color=DIMMED, bold=True)
    add_text(slide, left + 0.15, top + 0.4, width - 0.3, 0.5, str(value), size=22, color=color, bold=True)

def add_bar(slide, left, top, width, max_width, height, color, label, value, label_width=4):
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(max_width), Inches(height))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = RGBColor(0x2A, 0x2E, 0x35)
    bg_shape.line.fill.background()
    if width > 0.05:
        fill_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        fill_shape.fill.solid()
        fill_shape.fill.fore_color.rgb = color
        fill_shape.line.fill.background()
    add_text(slide, left + max_width + 0.15, top - 0.03, label_width, height + 0.06, f"{label}: {value}", size=10, color=DIMMED)

def divider(slide):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.2), Inches(12.3), Pt(2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = BLUE
    shape.line.fill.background()


# ============================================================
# SLIDE 1: Executive Summary — concise, let KPIs do the talking
# ============================================================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide1)

add_text(slide1, 0.5, 0.3, 12, 0.6, "Apple Patent Explorer", size=32, color=WHITE, bold=True)
add_text(slide1, 0.5, 0.85, 12, 0.4, "Executive Summary  |  Design Patent Portfolio Analytics", size=16, color=DIMMED)
divider(slide1)

# Short description — two lines max
add_text(slide1, 0.5, 1.6, 12, 0.4,
    "Interactive dashboard for 3,000 Apple design patents pulled from the USPTO PEDS API.",
    size=14, color=WHITE)
add_text(slide1, 0.5, 2.0, 12, 0.3,
    "Raw data cleaned with clean_apple.py (parses names, formats dates, standardizes titles) then served via Flask + React.",
    size=11, color=DIMMED)

# KPI cards — 2 rows of 4, big and centered
row1_y = 2.7
row2_y = 4.1
row3_y = 5.5
cw = 2.7  # card width
ch = 1.1  # card height
gap = 0.3
x1 = 0.6
x2 = x1 + cw + gap
x3 = x2 + cw + gap
x4 = x3 + cw + gap

add_card(slide1, x1, row1_y, cw, ch, "TOTAL APPLICATIONS", "3,000", WHITE)
add_card(slide1, x2, row1_y, cw, ch, "GRANTED", "2,961", GREEN)
add_card(slide1, x3, row1_y, cw, ch, "GRANT RATE", "98.7%", CYAN)
add_card(slide1, x4, row1_y, cw, ch, "AVG TIME TO GRANT", "19.2 mo", YELLOW)

add_card(slide1, x1, row2_y, cw, ch, "UNIQUE INVENTORS", "1,680", PURPLE)
add_card(slide1, x2, row2_y, cw, ch, "ABANDONED", "29", RED)
add_card(slide1, x3, row2_y, cw, ch, "PENDING / ACTIVE", "10", BLUE)
add_card(slide1, x4, row2_y, cw, ch, "PORTFOLIO TYPE", "100% Design", ORANGE)

add_card(slide1, x1, row3_y, cw, ch, "PEAK FILING YEAR", "2022 (502)", WHITE)
add_card(slide1, x2, row3_y, cw, ch, "FILING SPAN", "2018-2025", WHITE)
add_card(slide1, x3, row3_y, cw, ch, "LAST 12 MONTHS", "41 filings", RED)
add_card(slide1, x4, row3_y, cw, ch, "TOP LAW FIRM", "Sterne Kessler", BLUE)

add_text(slide1, 0.5, 7.0, 12, 0.3, "Apple Inc.  |  Design Patent Portfolio  |  3,000 Applications  |  2018-2025", size=10, color=DIMMED, align=PP_ALIGN.CENTER)


# ============================================================
# SLIDE 2: Prosecution — focus on pending/OA/abandonment, NOT granted
# ============================================================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide2)

add_text(slide2, 0.5, 0.3, 12, 0.6, "Prosecution & Status Analysis", size=28, color=WHITE, bold=True)
add_text(slide2, 0.5, 0.8, 12, 0.3, "Office Actions, Abandonments & Active Pipeline", size=14, color=DIMMED)
divider(slide2)

# LEFT TOP: Active Examination Pipeline — this is what matters
add_text(slide2, 0.5, 1.45, 6, 0.3, "ACTIVE EXAMINATION PIPELINE (10 PENDING)", size=11, color=BLUE, bold=True)

pipeline_data = [
    ("Ready for Examination", 2, BLUE),
    ("Office Action (Non-Final)", 3, ORANGE),
    ("Allowed / Notice of Allowance", 3, GREEN),
    ("Other Active", 2, PURPLE),
]

px = 0.5
for label, val, color in pipeline_data:
    add_card(slide2, px, 1.85, 2.8, 0.95, label.upper(), str(val), color)
    px += 3.05

# LEFT: Abandonment breakdown bars (no "Patented Case" bar)
add_text(slide2, 0.5, 3.2, 6.5, 0.3, "ABANDONMENT & OA BREAKDOWN (EXCLUDING GRANTED)", size=11, color=RED, bold=True)

pending_statuses = [
    ("Failed to Respond to OA", 25, RED),
    ("Failed to Pay Issue Fee", 3, RED),
    ("NOA Mailed (in progress)", 3, CYAN),
    ("Docketed / Ready for Exam", 2, BLUE),
    ("Response to Non-Final OA", 2, ORANGE),
    ("After Examiner's Answer", 1, RED),
    ("Non-Final OA Mailed", 1, ORANGE),
]

y = 3.6
max_bar = 3.2
max_val = 25
for label, val, clr in pending_statuses:
    bar_w = max(0.08, (val / max_val) * max_bar)
    add_bar(slide2, 0.5, y, bar_w, max_bar, 0.22, clr, label, str(val), label_width=2.8)
    y += 0.38

# RIGHT: Prosecution KPIs — stacked vertically
add_text(slide2, 7.3, 3.2, 5.5, 0.3, "PROSECUTION METRICS", size=11, color=BLUE, bold=True)

add_card(slide2, 7.3, 3.6, 2.5, 0.85, "TOTAL ABANDONMENTS", "29", RED)
add_card(slide2, 10.1, 3.6, 2.5, 0.85, "ABANDONMENT RATE", "0.97%", RED)
add_card(slide2, 7.3, 4.65, 2.5, 0.85, "OFFICE ACTIONS ISSUED", "~31", ORANGE)
add_card(slide2, 10.1, 4.65, 2.5, 0.85, "STILL PENDING", "10", BLUE)

# Key takeaway
add_text(slide2, 7.3, 5.75, 5.3, 0.3, "KEY TAKEAWAY", size=11, color=GREEN, bold=True)
add_text(slide2, 7.3, 6.05, 5.3, 0.6,
    "29 abandonments out of 3,000 (0.97%). 25 were failure to respond to OAs — "
    "likely strategic. Only 10 apps still in the pipeline.",
    size=11, color=WHITE)

add_text(slide2, 0.5, 7.1, 12, 0.3, "Apple Inc.  |  Prosecution & Status Analysis  |  Data Source: USPTO PEDS", size=10, color=DIMMED, align=PP_ALIGN.CENTER)


# ============================================================
# SLIDE 3: Portfolio Analytics — fixed layout, no overlapping
# ============================================================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide3)

add_text(slide3, 0.5, 0.3, 12, 0.6, "Portfolio Analytics", size=28, color=WHITE, bold=True)
add_text(slide3, 0.5, 0.8, 12, 0.3, "Filing Trends, Top Inventors & Law Firm Distribution", size=14, color=DIMMED)
divider(slide3)

# LEFT: Filing trends bar chart — same as before but contained
add_text(slide3, 0.5, 1.45, 6.5, 0.3, "FILING TRENDS BY YEAR", size=11, color=BLUE, bold=True)

years_data = [
    ("2018", 327), ("2019", 417), ("2020", 484), ("2021", 448),
    ("2022", 502), ("2023", 403), ("2024", 321), ("2025", 98),
]
max_filings = 502
bar_base_y = 4.2
bar_max_h = 2.0
bar_w_each = 0.55
x_start = 0.7

for i, (year, count) in enumerate(years_data):
    bar_h = (count / max_filings) * bar_max_h
    x_pos = x_start + i * 0.72
    bar_shape = slide3.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x_pos), Inches(bar_base_y - bar_h),
        Inches(bar_w_each), Inches(bar_h)
    )
    bar_shape.fill.solid()
    bar_shape.fill.fore_color.rgb = BLUE
    bar_shape.line.fill.background()
    add_text(slide3, x_pos - 0.05, bar_base_y - bar_h - 0.25, 0.65, 0.25, str(count), size=9, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide3, x_pos - 0.05, bar_base_y + 0.05, 0.65, 0.25, year, size=9, color=DIMMED, align=PP_ALIGN.CENTER)

add_text(slide3, 0.5, 4.55, 6, 0.3, "Peak: 2022 (502)  |  Last 12 Months: 41 (-87% YoY)", size=10, color=DIMMED)

# RIGHT: Top 8 inventors (trimmed to fit)
add_text(slide3, 7.3, 1.45, 5.5, 0.3, "TOP INVENTORS BY PATENT COUNT", size=11, color=BLUE, bold=True)

inventors = [
    ("Jonathan P. IVE", 849), ("Jody AKANA", 781), ("Richard P. HOWARTH", 736),
    ("Peter RUSSELL-CLARKE", 728), ("Bartley K. ANDRE", 722),
    ("Duncan Robert KERR", 721), ("M. Evans HANKEY", 704),
    ("Eugene Antony WHANG", 684),
]
max_inv = 849
y = 1.85
for name, count in inventors:
    bar_w = (count / max_inv) * 2.5
    add_bar(slide3, 7.3, y, bar_w, 2.5, 0.2, YELLOW, name, str(count), label_width=3.0)
    y += 0.34

# BOTTOM LEFT: Law firms
add_text(slide3, 0.5, 5.1, 6, 0.3, "LAW FIRM DISTRIBUTION", size=11, color=BLUE, bold=True)

firms = [
    ("Sterne Kessler", 2452, BLUE),
    ("DesignLaw Group", 428, GREEN),
    ("SAIDMAN DesignLaw", 119, YELLOW),
]
y = 5.45
for name, count, clr in firms:
    bar_w = (count / 2452) * 3.5
    add_bar(slide3, 0.5, y, bar_w, 3.5, 0.22, clr, name, f"{count:,}", label_width=2.5)
    y += 0.38

# BOTTOM RIGHT: Filing activity cards
add_text(slide3, 7.3, 5.1, 5.5, 0.3, "FILING ACTIVITY", size=11, color=BLUE, bold=True)
add_card(slide3, 7.3, 5.45, 2.5, 0.85, "LAST 12 MONTHS", "41", RED)
add_card(slide3, 10.1, 5.45, 2.5, 0.85, "PRIOR 12 MONTHS", "308", GREEN)

add_text(slide3, 7.3, 6.4, 5, 0.3, "-87% year-over-year decline in new filings.", size=10, color=DIMMED)

add_text(slide3, 0.5, 7.1, 12, 0.3, "Apple Inc.  |  Portfolio Analytics  |  1,680 Unique Inventors  |  3 Law Firms", size=10, color=DIMMED, align=PP_ALIGN.CENTER)


# ============================================================
# SLIDE 4: How It Was Built — concise, no money talk
# ============================================================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide4)

add_text(slide4, 0.5, 0.3, 12, 0.6, "How It Was Built", size=28, color=WHITE, bold=True)
add_text(slide4, 0.5, 0.8, 12, 0.3, "Technology Stack & Architecture", size=14, color=DIMMED)
divider(slide4)

# Data flow cards — same layout, works well
add_text(slide4, 0.5, 1.45, 12, 0.3, "DATA FLOW", size=11, color=BLUE, bold=True)

flow_items = [
    ("1. DATA SOURCE", "USPTO PEDS API", "Official U.S. Patent\nOffice records", CYAN),
    ("2. DATA CLEANING", "clean_apple.py", "Parses names,\nformats dates,\nstandardizes titles", ORANGE),
    ("3. BACKEND", "Flask (Python)", "Serves cleaned data\nvia REST API", GREEN),
    ("4. FRONTEND", "React 19 + TypeScript", "Interactive dashboard\nwith charts & tables", BLUE),
]

x = 0.5
for title, tech, desc, clr in flow_items:
    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.85), Inches(2.8), Inches(1.6))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = clr
    card.line.width = Pt(2)
    add_text(slide4, x + 0.15, 1.95, 2.5, 0.25, title, size=9, color=clr, bold=True)
    add_text(slide4, x + 0.15, 2.2, 2.5, 0.4, tech, size=13, color=WHITE, bold=True)
    add_text(slide4, x + 0.15, 2.65, 2.5, 0.6, desc, size=10, color=DIMMED)
    if x < 9:
        arrow = slide4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 2.9), Inches(2.45), Inches(0.25), Inches(0.3))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = DIMMED
        arrow.line.fill.background()
    x += 3.15

# Tech stack — two columns of cards instead of wordy list
add_text(slide4, 0.5, 3.7, 12, 0.3, "TECH STACK", size=11, color=BLUE, bold=True)

stack_items = [
    ("React 19", "Web framework", BLUE),
    ("TypeScript", "Type-safe JavaScript", CYAN),
    ("Vite 7", "Build tool & dev server", GREEN),
    ("Mantine v8", "UI component library", PURPLE),
    ("AG Grid v35", "Data table (sort, filter, export)", YELLOW),
    ("Plotly.js", "Charts & visualizations", ORANGE),
    ("Python Flask", "Backend API server", GREEN),
    ("USPTO PEDS", "Patent data source", CYAN),
]

# 2 rows of 4 cards
row1_y = 4.1
row2_y = 5.25
card_w = 2.8
card_h = 0.9
gap = 0.2

for i, (tech, desc, clr) in enumerate(stack_items):
    row = i // 4
    col = i % 4
    cx = 0.5 + col * (card_w + gap)
    cy = row1_y if row == 0 else row2_y

    s = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(card_w), Inches(card_h))
    s.fill.solid()
    s.fill.fore_color.rgb = CARD_BG
    s.line.color.rgb = BORDER
    s.line.width = Pt(1)

    # Color dot
    dot = slide4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + 0.15), Inches(cy + 0.18), Inches(0.12), Inches(0.12))
    dot.fill.solid()
    dot.fill.fore_color.rgb = clr
    dot.line.fill.background()

    add_text(slide4, cx + 0.35, cy + 0.12, 2.3, 0.25, tech, size=12, color=WHITE, bold=True)
    add_text(slide4, cx + 0.35, cy + 0.45, 2.3, 0.3, desc, size=10, color=DIMMED)

# Brief one-liner
add_text(slide4, 0.5, 6.4, 12, 0.3,
    "100% open-source stack. Runs in any browser. Data from official USPTO records.",
    size=12, color=DIMMED, align=PP_ALIGN.CENTER)

add_text(slide4, 0.5, 7.1, 12, 0.3, "Apple Inc.  |  Built with React 19 + Flask + Plotly  |  Open-Source Stack", size=10, color=DIMMED, align=PP_ALIGN.CENTER)


# Save
prs.save("/Users/fgomez/Developer/Apple/Apple_Patent_Explorer_Executive_Summary.pptx")
print("PowerPoint saved.")
