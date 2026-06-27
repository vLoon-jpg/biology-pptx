#!/usr/bin/env python
"""Build a professional 6-slide PPTX on Photosynthesis."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ── Colour palette ──────────────────────────────────────────────
DARK_GREEN   = RGBColor(0x1B, 0x5E, 0x20)   # headers, footer
MID_GREEN    = RGBColor(0x2E, 0x7D, 0x32)   # accents
LIGHT_GREEN  = RGBColor(0x4C, 0xAF, 0x50)   # highlights
PALE_GREEN   = RGBColor(0xE8, 0xF5, 0xE9)   # backgrounds
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
BLACK        = RGBColor(0x00, 0x00, 0x00)
DARK_GREY    = RGBColor(0x33, 0x33, 0x33)
MID_GREY     = RGBColor(0x75, 0x75, 0x75)
ACCENT_GOLD  = RGBColor(0xFF, 0xA7, 0x26)   # bullet highlights
LEAF_YELLOW  = RGBColor(0xCD, 0xDC, 0x39)

W = Inches(13.333)   # 16:9 widescreen
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

# ── Helpers ─────────────────────────────────────────────────────
def add_bg(slide, color):
    """Fill the entire slide background with *color*."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    """Add a rectangle shape."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, left, top, width, height, text, font_size=18,
                color=BLACK, bold=False, alignment=PP_ALIGN.LEFT,
                font_name="Calibri", anchor=MSO_ANCHOR.TOP):
    """Add a text box with a single run."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox

def add_bullet_frame(slide, left, top, width, height, bullets, font_size=16,
                     color=DARK_GREY, bold_first=False, line_spacing=1.3,
                     bullet_char="●", bullet_color=LIGHT_GREEN):
    """Add a text frame with bulleted lines."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"{bullet_char}  {line}"
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.space_after = Pt(line_spacing * font_size - font_size)
        # First word bold option
        # (simple approach — bold the whole line for emphasis bullets)
    return txBox

def add_footer(slide, slide_num, presenter=""):
    """Standard footer bar."""
    bar = add_rect(slide, Inches(0), H - Inches(0.55), W, Inches(0.55), DARK_GREEN)
    add_textbox(slide, Inches(0.5), H - Inches(0.48), Inches(5), Inches(0.4),
                f"Photosynthesis: How Plants Make Their Own Food  |  {presenter}",
                font_size=9, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(12.2), H - Inches(0.48), Inches(0.8), Inches(0.4),
                str(slide_num), font_size=9, color=WHITE, bold=True,
                alignment=PP_ALIGN.RIGHT)

def add_slide_number(slide, num, presenter=""):
    add_footer(slide, num, presenter)

# ── Presenters ──────────────────────────────────────────────────
PRESENTERS = [
    "Dr. Sarah Chen",
    "Prof. Marcus Williams",
    "Dr. Aisha Patel",
    "Dr. James Rodriguez",
    "Prof. Emily Nakamura",
]

# =====================================================================
# SLIDE 1 — Title Slide
# =====================================================================
sl = prs.slides.add_slide(prs.slide_layouts[6])  # blank
add_bg(sl, DARK_GREEN)

# Decorative leaf-shaped accent band
add_rect(sl, Inches(0), Inches(2.8), W, Inches(1.8), MID_GREEN)

# Title
add_textbox(sl, Inches(1), Inches(1.0), Inches(11.3), Inches(1.2),
            "Photosynthesis",
            font_size=52, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER,
            font_name="Calibri Light")
add_textbox(sl, Inches(1), Inches(2.1), Inches(11.3), Inches(0.8),
            "How Plants Make Their Own Food",
            font_size=28, color=LEAF_YELLOW, bold=False, alignment=PP_ALIGN.CENTER,
            font_name="Calibri Light")

# Presenter list inside the band
y_start = Inches(3.15)
add_textbox(sl, Inches(2), y_start, Inches(9.3), Inches(0.35),
            "PRESENTED BY", font_size=11, color=WHITE, bold=True,
            alignment=PP_ALIGN.CENTER)

for i, name in enumerate(PRESENTERS):
    y = y_start + Inches(0.35) + Inches(i * 0.32)
    add_textbox(sl, Inches(2), y, Inches(9.3), Inches(0.3),
                name, font_size=16, color=WHITE, bold=False,
                alignment=PP_ALIGN.CENTER)

# Date / context
add_textbox(sl, Inches(1), Inches(6.2), Inches(11.3), Inches(0.4),
            "Department of Biological Sciences  |  June 2026",
            font_size=11, color=RGBColor(0xA5, 0xD6, 0xA7), bold=False,
            alignment=PP_ALIGN.CENTER)

add_footer(sl, 1, "")

# =====================================================================
# SLIDE 2 — What is Photosynthesis?
# =====================================================================
sl = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(sl, WHITE)

# Header bar
add_rect(sl, Inches(0), Inches(0), W, Inches(1.15), DARK_GREEN)
add_textbox(sl, Inches(0.8), Inches(0.25), Inches(11.5), Inches(0.7),
            "What is Photosynthesis?",
            font_size=36, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT,
            font_name="Calibri Light")

# Left panel — Definition
left_box = add_rect(sl, Inches(0.6), Inches(1.5), Inches(5.8), Inches(2.4), PALE_GREEN, LIGHT_GREEN)
add_textbox(sl, Inches(0.9), Inches(1.65), Inches(5.2), Inches(2.1),
            "Photosynthesis is the biochemical process by which green plants, algae, and "
            "some bacteria convert light energy into chemical energy stored in glucose. "
            "Using sunlight, water, and carbon dioxide, plants produce glucose and release "
            "oxygen as a by-product.",
            font_size=16, color=DARK_GREY, bold=False, alignment=PP_ALIGN.LEFT)

# Right panel — Equation (styled)
eq_box = add_rect(sl, Inches(6.8), Inches(1.5), Inches(5.8), Inches(2.4), DARK_GREEN)
add_textbox(sl, Inches(7.1), Inches(1.65), Inches(5.2), Inches(0.4),
            "THE CHEMICAL EQUATION",
            font_size=12, color=LEAF_YELLOW, bold=True, alignment=PP_ALIGN.CENTER)
add_textbox(sl, Inches(7.1), Inches(2.2), Inches(5.2), Inches(1.5),
            "6 CO₂  +  6 H₂O  +  Light Energy\n"
            "           ═══════════════►\n"
            "C₆H₁₂O₆  +  6 O₂",
            font_size=18, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

# Key facts row
facts = [
    ("📗  Autotrophic", "Plants produce\ntheir own food"),
    ("🌱  Chloroplasts", "Organelles where\nphotosynthesis occurs"),
    ("☀️  Endergonic", "Requires energy\ninput (sunlight)"),
    ("💨  Gas Exchange", "CO₂ in, O₂ out\nvia stomata"),
]
for i, (title, desc) in enumerate(facts):
    x = Inches(0.6 + i * 3.15)
    box = add_rect(sl, x, Inches(4.25), Inches(2.9), Inches(1.7), PALE_GREEN)
    box.line.color.rgb = LIGHT_GREEN
    box.line.width = Pt(1.5)
    add_textbox(sl, Inches(x.inches + 0.15), Inches(4.4), Inches(2.6), Inches(0.5),
                title, font_size=12, color=DARK_GREEN, bold=True,
                alignment=PP_ALIGN.CENTER)
    add_textbox(sl, Inches(x.inches + 0.15), Inches(4.85), Inches(2.6), Inches(0.9),
                desc, font_size=11, color=MID_GREY, bold=False,
                alignment=PP_ALIGN.CENTER)

# Bottom highlight
add_rect(sl, Inches(0.6), Inches(6.2), Inches(12.1), Inches(0.55), DARK_GREEN)
add_textbox(sl, Inches(0.9), Inches(6.28), Inches(11.5), Inches(0.4),
            "💡  Fun Fact:  Photosynthesis produces ~170 billion tonnes of biomass annually — "
            "the foundation of nearly all food chains on Earth!",
            font_size=12, color=WHITE, bold=False, alignment=PP_ALIGN.CENTER)

add_slide_number(sl, 2, PRESENTERS[0])

# =====================================================================
# SLIDE 3 — The Ingredients
# =====================================================================
sl = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(sl, WHITE)

# Header
add_rect(sl, Inches(0), Inches(0), W, Inches(1.15), DARK_GREEN)
add_textbox(sl, Inches(0.8), Inches(0.25), Inches(11.5), Inches(0.7),
            "The Ingredients of Photosynthesis",
            font_size=36, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT,
            font_name="Calibri Light")

# Four ingredient cards
ingredients = [
    ("☀️  Sunlight",
     "• Primary energy source\n• Visible spectrum (400-700 nm)\n• Absorbed by chlorophyll pigments\n• Powers light-dependent reactions"),
    ("💧  Water (H₂O)",
     "• Absorbed by roots from soil\n• Transported via xylem vessels\n• Split during photolysis\n• Provides electrons & H⁺ ions"),
    ("🫧  Carbon Dioxide (CO₂)",
     "• Enters through stomata\n• Diffuses into leaf mesophyll\n• Fixed in the Calvin cycle\n• Provides carbon backbone"),
    ("🌿  Chlorophyll",
     "• Green pigment in chloroplasts\n• Located in thylakoid membranes\n• Absorbs red & blue light best\n• Reflects green light"),
]

for i, (title, desc) in enumerate(ingredients):
    x = Inches(0.4 + i * 3.2)
    card = add_rect(sl, x, Inches(1.45), Inches(2.95), Inches(3.5), PALE_GREEN)
    card.line.color.rgb = LIGHT_GREEN
    card.line.width = Pt(1.5)
    # Title bar on card
    add_rect(sl, x, Inches(1.45), Inches(2.95), Inches(0.65), LIGHT_GREEN)
    add_textbox(sl, Inches(x.inches + 0.1), Inches(1.55), Inches(2.75), Inches(0.5),
                title, font_size=15, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    add_textbox(sl, Inches(x.inches + 0.2), Inches(2.3), Inches(2.55), Inches(2.4),
                desc, font_size=12, color=DARK_GREY, bold=False, alignment=PP_ALIGN.LEFT)

# Bottom summary bar
add_rect(sl, Inches(0.4), Inches(5.2), Inches(12.5), Inches(0.6), DARK_GREEN)
add_textbox(sl, Inches(0.7), Inches(5.28), Inches(12), Inches(0.45),
            "🔑  All four ingredients must be present simultaneously for photosynthesis to proceed efficiently.",
            font_size=13, color=WHITE, bold=False, alignment=PP_ALIGN.CENTER)

# Additional detail: absorption spectrum mini-note
add_rect(sl, Inches(0.4), Inches(6.05), Inches(12.5), Inches(0.75), PALE_GREEN)
add_textbox(sl, Inches(0.7), Inches(6.12), Inches(11.9), Inches(0.6),
            "📊  Key Insight:  Chlorophyll a absorbs maximally at ~430 nm (blue) and ~662 nm (red). "
            "Accessory pigments (chlorophyll b, carotenoids) broaden the absorption spectrum, "
            "capturing energy that chlorophyll a alone would miss.",
            font_size=11, color=DARK_GREY, bold=False, alignment=PP_ALIGN.LEFT)

add_slide_number(sl, 3, PRESENTERS[1])

# =====================================================================
# SLIDE 4 — The Process (Two Stages)
# =====================================================================
sl = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(sl, WHITE)

# Header
add_rect(sl, Inches(0), Inches(0), W, Inches(1.15), DARK_GREEN)
add_textbox(sl, Inches(0.8), Inches(0.25), Inches(11.5), Inches(0.7),
            "The Two Stages of Photosynthesis",
            font_size=36, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT,
            font_name="Calibri Light")

# ---- Stage 1: Light-Dependent Reactions ----
left_box = add_rect(sl, Inches(0.4), Inches(1.45), Inches(6.1), Inches(2.8), PALE_GREEN, LIGHT_GREEN)
# Stage 1 header
add_rect(sl, Inches(0.4), Inches(1.45), Inches(6.1), Inches(0.55), MID_GREEN)
add_textbox(sl, Inches(0.55), Inches(1.52), Inches(5.8), Inches(0.45),
            "⚡  STAGE 1: Light-Dependent Reactions",
            font_size=17, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT)

stage1_bullets = [
    "Occurs in the thylakoid membranes of chloroplasts",
    "Sunlight splits water molecules (photolysis): 2 H₂O → 4 H⁺ + 4 e⁻ + O₂",
    "Electrons move through the electron transport chain (ETC)",
    "ATP is produced via chemiosmosis (photophosphorylation)",
    "NADP⁺ is reduced to NADPH",
    "Oxygen is released as a by-product into the atmosphere",
]
add_bullet_frame(sl, Inches(0.65), Inches(2.2), Inches(5.6), Inches(1.9),
                 stage1_bullets, font_size=12, color=DARK_GREY, bullet_char="▸",
                 bullet_color=MID_GREEN)

# Arrow between stages
add_textbox(sl, Inches(6.35), Inches(2.6), Inches(0.65), Inches(0.5),
            "ATP\nNADPH\n▼", font_size=10, color=MID_GREEN, bold=True,
            alignment=PP_ALIGN.CENTER)

# ---- Stage 2: Calvin Cycle ----
right_box = add_rect(sl, Inches(6.8), Inches(1.45), Inches(6.1), Inches(2.8), PALE_GREEN, LIGHT_GREEN)
# Stage 2 header
add_rect(sl, Inches(6.8), Inches(1.45), Inches(6.1), Inches(0.55), MID_GREEN)
add_textbox(sl, Inches(6.95), Inches(1.52), Inches(5.8), Inches(0.45),
            "🔄  STAGE 2: Calvin Cycle (Light-Independent)",
            font_size=17, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT)

stage2_bullets = [
    "Occurs in the stroma of chloroplasts",
    "Uses ATP and NADPH from Stage 1",
    "CO₂ is fixed by the enzyme RuBisCO",
    "3-phosphoglycerate (3-PGA) is reduced to G3P",
    "G3P is used to synthesize glucose (C₆H₁₂O₆)",
    "RuBP is regenerated to continue the cycle",
]
add_bullet_frame(sl, Inches(7.05), Inches(2.2), Inches(5.6), Inches(1.9),
                 stage2_bullets, font_size=12, color=DARK_GREY, bullet_char="▸",
                 bullet_color=MID_GREEN)

# ---- Summary flow diagram ──────────────────────────────────────
add_textbox(sl, Inches(0.4), Inches(4.5), Inches(12.5), Inches(0.4),
            "PROCESS FLOW", font_size=12, color=MID_GREEN, bold=True,
            alignment=PP_ALIGN.CENTER)

flow_items = ["Sunlight\n+ H₂O", "Light\nReactions", "ATP +\nNADPH", "Calvin\nCycle", "Glucose\nC₆H₁₂O₆"]
flow_width = Inches(2.2)
total_width = len(flow_items) * flow_width.inches + (len(flow_items) - 1) * 0.3
start_x = (W.inches - total_width) / 2

for i, item in enumerate(flow_items):
    x = Inches(start_x + i * (flow_width.inches + 0.3))
    box = add_rect(sl, x, Inches(4.95), flow_width, Inches(0.75),
                   MID_GREEN if i % 2 == 0 else LIGHT_GREEN)
    add_textbox(sl, Inches(x.inches + 0.05), Inches(5.0), Inches(flow_width.inches - 0.1), Inches(0.65),
                item, font_size=11, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    if i < len(flow_items) - 1:
        add_textbox(sl, Inches(x.inches + flow_width.inches + 0.02), Inches(5.1), Inches(0.3), Inches(0.4),
                    "→", font_size=18, color=MID_GREEN, bold=True, alignment=PP_ALIGN.CENTER)

# Bottom note
add_rect(sl, Inches(0.4), Inches(6.05), Inches(12.5), Inches(0.75), DARK_GREEN)
add_textbox(sl, Inches(0.7), Inches(6.12), Inches(11.9), Inches(0.6),
            "🧬  Key Enzyme: RuBisCO (Ribulose-1,5-bisphosphate carboxylase/oxygenase) is the most abundant "
            "protein on Earth and catalyzes the first step of carbon fixation.",
            font_size=11, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT)

add_slide_number(sl, 4, PRESENTERS[2])

# =====================================================================
# SLIDE 5 — Why Photosynthesis Matters
# =====================================================================
sl = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(sl, WHITE)

# Header
add_rect(sl, Inches(0), Inches(0), W, Inches(1.15), DARK_GREEN)
add_textbox(sl, Inches(0.8), Inches(0.25), Inches(11.5), Inches(0.7),
            "Why Photosynthesis Matters",
            font_size=36, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT,
            font_name="Calibri Light")

# Four impact columns
impacts = [
    ("🌍  Oxygen Production",
     "Photosynthesis generates the oxygen that sustains aerobic life. Approximately 50-80% "
     "of Earth's oxygen comes from marine phytoplankton and terrestrial plants combined."),
    ("🍽️  Food Chain Foundation",
     "As primary producers, plants convert inorganic carbon into organic compounds. "
     "Every food chain on Earth — from herbivores to apex predators — depends on this process."),
    ("🌡️  Climate Regulation",
     "Photosynthesis removes ~120 Gt of CO₂ from the atmosphere annually. Forests and oceans "
     "act as carbon sinks, mitigating the greenhouse effect and regulating global temperatures."),
    ("⚡  Renewable Energy",
     "Biofuels, biomass energy, and even fossil fuels (ancient photosynthetic matter) trace "
     "back to photosynthesis. Understanding it inspires artificial photosynthesis research."),
]

for i, (title, desc) in enumerate(impacts):
    y = Inches(1.45 + i * 1.35)
    card = add_rect(sl, Inches(0.4), y, Inches(12.5), Inches(1.2), PALE_GREEN, LIGHT_GREEN)
    # Icon + title column
    add_rect(sl, Inches(0.4), y, Inches(3.0), Inches(1.2), LIGHT_GREEN)
    add_textbox(sl, Inches(0.6), Inches(y.inches + 0.3), Inches(2.6), Inches(0.6),
                title, font_size=16, color=WHITE, bold=True, alignment=PP_ALIGN.LEFT)

    # Description
    add_textbox(sl, Inches(3.7), Inches(y.inches + 0.15), Inches(9.0), Inches(0.9),
                desc, font_size=12, color=DARK_GREY, bold=False, alignment=PP_ALIGN.LEFT)

# Bottom callout
add_rect(sl, Inches(0.4), Inches(6.9), Inches(12.5), Inches(0.45), DARK_GREEN)
add_textbox(sl, Inches(0.7), Inches(6.93), Inches(11.9), Inches(0.38),
            "🌱  \"In every walk with nature, one receives far more than he seeks.\" — John Muir",
            font_size=11, color=WHITE, bold=False, alignment=PP_ALIGN.CENTER)

# Note: the footer is tight here, so adjust
add_textbox(sl, Inches(0.5), Inches(6.93), Inches(5), Inches(0.4),
            f"Photosynthesis: How Plants Make Their Own Food  |  {PRESENTERS[3]}",
            font_size=9, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT)
add_textbox(sl, Inches(12.2), Inches(6.93), Inches(0.8), Inches(0.4),
            "5", font_size=9, color=WHITE, bold=True, alignment=PP_ALIGN.RIGHT)

# =====================================================================
# SLIDE 6 — Summary & Questions
# =====================================================================
sl = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(sl, DARK_GREEN)

# Large decorative element
add_rect(sl, Inches(0), Inches(2.6), W, Inches(2.8), MID_GREEN)

# Title
add_textbox(sl, Inches(1), Inches(0.6), Inches(11.3), Inches(1.0),
            "Summary & Key Takeaways",
            font_size=44, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER,
            font_name="Calibri Light")

# Summary bullets in the band
summary_points = [
    "Photosynthesis converts light energy → chemical energy (glucose)",
    "Requires four key ingredients: Sunlight, Water, CO₂, and Chlorophyll",
    "Two stages: Light-Dependent Reactions + Calvin Cycle",
    "Produces oxygen essential for aerobic life on Earth",
    "Foundation of global food chains and carbon cycle regulation",
]

for i, point in enumerate(summary_points):
    y = Inches(2.9 + i * 0.42)
    add_textbox(sl, Inches(1.5), y, Inches(10.3), Inches(0.4),
                f"✦  {point}", font_size=16, color=WHITE, bold=False,
                alignment=PP_ALIGN.LEFT)

# Questions section
add_textbox(sl, Inches(1), Inches(5.6), Inches(11.3), Inches(0.7),
            "Questions & Discussion",
            font_size=32, color=LEAF_YELLOW, bold=True, alignment=PP_ALIGN.CENTER,
            font_name="Calibri Light")

add_textbox(sl, Inches(1), Inches(6.2), Inches(11.3), Inches(0.5),
            "Thank you for your attention!  We welcome your questions.",
            font_size=16, color=WHITE, bold=False, alignment=PP_ALIGN.CENTER)

# Footer
add_textbox(sl, Inches(0.5), Inches(6.93), Inches(5), Inches(0.4),
            f"Photosynthesis: How Plants Make Their Own Food  |  {PRESENTERS[4]}",
            font_size=9, color=WHITE, bold=False, alignment=PP_ALIGN.LEFT)
add_textbox(sl, Inches(12.2), Inches(6.93), Inches(0.8), Inches(0.4),
            "6", font_size=9, color=WHITE, bold=True, alignment=PP_ALIGN.RIGHT)

# =====================================================================
# Save
# =====================================================================
output_path = r"C:\Users\LENOVO\projects\biology-pptx\Photosynthesis-Claude.pptx"
prs.save(output_path)
print(f"✅ Saved: {output_path}")
print(f"   Slides: {len(prs.slides)}")
print(f"   Presenters: {', '.join(PRESENTERS)}")
