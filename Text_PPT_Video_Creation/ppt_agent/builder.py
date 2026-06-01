

def _text_to_bullets(text: str) -> list[str]:
    """Convert paragraph text to bullet points by splitting on sentences or semicolons."""
    if not text or not isinstance(text, str):
        return []
    text = clean_text(text).strip()
    # Try splitting by sentence (. ! ?)
    bullets = re.split(r'(?<=[.!?])\s+', text)
    bullets = [b.strip() for b in bullets if len(b.strip()) > 10]
    if len(bullets) < 2:
        # Try splitting by semicolons
        bullets = re.split(r';\s*', text)
        bullets = [b.strip() for b in bullets if len(b.strip()) > 10]
    if len(bullets) < 2:
        # Try splitting by "and" or "or"
        bullets = re.split(r'\s+(?:and|or)\s+', text)
        bullets = [b.strip() for b in bullets if len(b.strip()) > 10]
    if len(bullets) < 1:
        bullets = [text]
    return bullets[:6]
"""
builder.py - Build PPTX slides from AI-generated outline + embed audio narration.

Audio embedding strategy:
  - Add each .mp3 as a media Part in the PPTX package
  - Insert a hidden <p:sp> audio shape via raw XML
  - Insert <p:timing> XML so the audio auto-plays when the slide is shown
"""

import os
import copy
import re
import random
from pathlib import Path
from lxml import etree

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

from .themes import THEMES, FONTS, rgb
from .normalize import normalize_slides

# ── Slide dimensions (16:9) ──────────────────────────────────────────────────
W = Inches(13.333)
H = Inches(7.5)
MARGIN = Inches(0.6)

# ── XML Namespaces ───────────────────────────────────────────────────────────
PML  = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML  = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14  = "http://schemas.microsoft.com/office/powerpoint/2010/main"
AUDIO_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media"


# ═══════════════════════════════════════════════════════════════════════════
# Helper utilities
# ═══════════════════════════════════════════════════════════════════════════

def _new_prs() -> Presentation:
    """Create a blank 16:9 presentation."""
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs


def _blank_slide(prs: Presentation):
    """Add a completely blank slide (layout index 6)."""
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)


def _fill_bg(slide, color_tuple: tuple):
    """Fill slide background with a solid color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*color_tuple)


def _add_rect(slide, left, top, width, height, color_tuple):
    """Add a solid filled rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height,
    )
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color_tuple)
    shape.line.color.rgb = RGBColor(*color_tuple)
    return shape


def _add_textbox(slide, left, top, width, height,
                 text, font_size, color_tuple,
                 bold=False, align=PP_ALIGN.LEFT, wrap=True) -> None:
    """Add a text box with given formatting, handling newlines correctly to prevent overlap."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    
    lines = str(text).split('\n')
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        if idx > 0:
            p.space_before = Pt(8)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color_tuple)
    return txBox


def clean_text(text: str) -> str:
    """Strip markdown bold markers (**), nested list bullet markers (*, -, •), and extra spaces."""
    if not text or not isinstance(text, str):
        return text
    # Remove markdown bold/italic/underline characters like **, __, *, _
    text = re.sub(r"\*\*|__|\*|_", "", text)
    # Strip any leading bullets or list dashes/dots, e.g. "- ", "* ", "• ", "+ "
    text = re.sub(r"^[ \t]*[-*•+]+[ \t]*", "", text)
    return text.strip()


def clean_data(val):
    """Recursively clean all text in the data structure."""
    if isinstance(val, str):
        return clean_text(val)
    elif isinstance(val, list):
        return [clean_data(x) for x in val]
    elif isinstance(val, dict):
        return {k: clean_data(v) for k, v in val.items()}
    return val


def get_dynamic_theme(theme_name: str, topic: str = "") -> str:
    """Resolve theme name, matching keywords if 'Auto-Select' is chosen, or selecting randomly."""
    available_themes = list(THEMES.keys())
    
    if theme_name == "Random Theme":
        return random.choice(available_themes)
        
    if theme_name == "Auto-Select (AI Pick)":
        topic_lower = str(topic).lower()
        
        tech_keywords = ["ai", "tech", "data", "future", "algorithm", "digital", "robot", "cyber", "science", "software", "development", "machine learning"]
        business_keywords = ["business", "corporate", "sales", "finance", "strategy", "investment", "marketing", "pitch", "quarterly", "report", "growth", "revenue"]
        education_keywords = ["learn", "teach", "academic", "university", "school", "history", "thesis", "course", "study", "research"]
        nature_keywords = ["nature", "green", "environment", "climate", "garden", "earth", "eco", "energy", "sustainable", "plant", "forest"]
        creative_keywords = ["art", "design", "creative", "fashion", "aesthetic", "music", "style", "film", "portfolio", "photo"]
        
        if any(w in topic_lower for w in tech_keywords):
            return random.choice(["Dark Pro", "Midnight Navy", "Tech Amber", "Amethyst", "Business Steel"])
        elif any(w in topic_lower for w in business_keywords):
            return random.choice(["Light Corporate", "Creatix Dark", "Global Teal", "Corporate Navy"])
        elif any(w in topic_lower for w in education_keywords):
            return random.choice(["Academic Blue", "Academic Dark", "Skylark Blue"])
        elif any(w in topic_lower for w in nature_keywords):
            return random.choice(["Nature Studio", "Earthy Warmth"])
        elif any(w in topic_lower for w in creative_keywords):
            return random.choice(["Vibrant Creative", "Warm Aesthete", "Bold Red", "Bold Vermillion"])
        
        return random.choice(available_themes)
        
    return theme_name


def _add_notes(slide, notes_text: str):
    """Add speaker notes to a slide."""
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


# ═══════════════════════════════════════════════════════════════════════════
# Audio Embedding
# ═══════════════════════════════════════════════════════════════════════════

def _embed_audio(slide, audio_path: str):
    """
    Embed an .mp3 file into a slide using python-pptx's native add_movie.
    This avoids XML corruption issues.
    """
    if not audio_path or not os.path.exists(audio_path):
        return

    try:
        # We place a small speaker icon off-screen or in the corner
        slide.shapes.add_movie(
            audio_path,
            Inches(0.5), Inches(H.inches - 1.0), Inches(0.5), Inches(0.5),
            poster_frame_image=None,
            mime_type='audio/mpeg'
        )
    except Exception as e:
        print(f"Warning: Failed to embed audio with add_movie: {e}")


def _add_transition(slide, transition_type="fade"):
    """
    Add an elegant transition animation to the slide (XML manipulation).
    """
    try:
        sld = slide._element
        # Remove any existing transition element
        for elem in list(sld):
            if elem.tag == qn('p:transition'):
                sld.remove(elem)
                
        transition = etree.Element(qn('p:transition'))
        
        if transition_type == "fade":
            effect = etree.Element(qn('p:fade'))
        elif transition_type == "push":
            effect = etree.Element(qn('p:push'))
            effect.set("dir", "l")  # push from left
        elif transition_type == "wipe":
            effect = etree.Element(qn('p:wipe'))
            effect.set("dir", "r")  # wipe from right
        elif transition_type == "split":
            effect = etree.Element(qn('p:split'))
        elif transition_type == "zoom":
            effect = etree.Element(qn('p:zoom'))
        elif transition_type == "wheel":
            effect = etree.Element(qn('p:wheel'))
            effect.set("spokes", "4")
        elif transition_type == "comb":
            effect = etree.Element(qn('p:comb'))
            effect.set("dir", "horz")
        elif transition_type == "checker":
            effect = etree.Element(qn('p:checker'))
            effect.set("dir", "horz")
        else:
            effect = etree.Element(qn('p:fade'))
            
        transition.append(effect)
        sld.append(transition)
    except Exception as e:
        print(f"Warning: Failed to add slide transition: {e}")


# ═══════════════════════════════════════════════════════════════════════════
# Slide Layout Builders
# ═══════════════════════════════════════════════════════════════════════════

def _build_title_slide(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])

    # Accent bar left
    _add_rect(slide, 0, 0, Inches(0.18), H, t["accent"])

    # Decorative bottom gradient bar
    _add_rect(slide, 0, H - Inches(0.08), W, Inches(0.08), t["accent2"])

    # Centered title
    title = data.get("title", "Presentation Title")
    _add_textbox(
        slide,
        MARGIN + Inches(0.4), Inches(2.2),
        W - MARGIN * 2 - Inches(0.4), Inches(1.6),
        title, FONTS["title"], t["title_color"],
        bold=True, align=PP_ALIGN.LEFT,
    )

    # Subtitle
    subtitle = data.get("subtitle") or ""
    if subtitle:
        _add_textbox(
            slide,
            MARGIN + Inches(0.4), Inches(4.0),
            W - MARGIN * 2 - Inches(0.4), Inches(1.0),
            subtitle, FONTS["subtitle"], t["subtitle_color"],
            bold=False, align=PP_ALIGN.LEFT,
        )

    # Accent dot
    _add_rect(slide, MARGIN + Inches(0.4), Inches(2.0), Inches(0.7), Inches(0.06), t["accent"])


def _build_section_header(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])

    # Full-width accent top bar
    _add_rect(slide, 0, 0, W, Inches(0.5), t["accent"])

    title = data.get("title", "Section")
    _add_textbox(
        slide,
        MARGIN, Inches(2.0),
        W - MARGIN * 2, Inches(1.2),
        title, FONTS["heading"] + 6, t["title_color"],
        bold=True, align=PP_ALIGN.CENTER,
    )

    subtitle = data.get("subtitle") or ""
    if subtitle:
        _add_textbox(
            slide,
            MARGIN, Inches(3.4),
            W - MARGIN * 2, Inches(1.0),
            subtitle, FONTS["body"], t["subtitle_color"],
            bold=False, align=PP_ALIGN.CENTER,
        )

    # Bottom divider
    _add_rect(slide, MARGIN * 3, H - Inches(1.2), W - MARGIN * 6, Inches(0.04), t["accent2"])


def _build_bullet_points(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])

    # Top accent bar
    _add_rect(slide, 0, 0, W, Inches(0.08), t["accent"])

    title = data.get("title", "Key Points")
    _add_textbox(
        slide,
        MARGIN, Inches(0.25),
        W - MARGIN * 2, Inches(0.85),
        title, FONTS["heading"], t["title_color"],
        bold=True,
    )

    # Divider under title
    _add_rect(slide, MARGIN, Inches(1.15), Inches(1.5), Inches(0.05), t["accent"])

    bullets = (data.get("content") or {}).get("bullets") or []
    if bullets:
        # Compact bullet layout with optimal spacing
        # Calculate height based on number of bullets to avoid excessive whitespace
        num_bullets = min(len(bullets), 6)
        line_height = Inches(0.65)  # ~18-20pt text + spacing
        content_height = num_bullets * line_height + Inches(0.2)
        max_content_height = H - Inches(2.0)
        actual_height = min(content_height, max_content_height)
        
        # Center content vertically if few bullets
        start_top = Inches(1.35)
        if actual_height < max_content_height - Inches(0.5):
            start_top = Inches(1.35) + (max_content_height - actual_height) / 2
        
        txBox = slide.shapes.add_textbox(
            MARGIN, start_top,
            W - MARGIN * 2, actual_height
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        
        for i, bullet in enumerate(bullets[:6]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(8)  # Reduced spacing
            p.space_before = Pt(2) if i > 0 else Pt(0)
            
            run_bullet = p.add_run()
            run_bullet.text = "▪  "
            run_bullet.font.size = Pt(FONTS["body"])
            run_bullet.font.bold = True
            run_bullet.font.color.rgb = RGBColor(*t["accent"])
            
            run_text = p.add_run()
            run_text.text = bullet
            run_text.font.size = Pt(FONTS["body"])
            run_text.font.bold = False
            run_text.font.color.rgb = RGBColor(*t["text_color"])


def _build_two_column(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])
    _add_rect(slide, 0, 0, W, Inches(0.08), t["accent"])

    title = data.get("title", "Comparison")
    _add_textbox(
        slide,
        MARGIN, Inches(0.15),
        W - MARGIN * 2, Inches(0.8),
        title, FONTS["heading"], t["title_color"],
        bold=True,
    )

    col_w = (W - MARGIN * 3) / 2
    content = data.get("content") or {}
    card_height = H - Inches(1.3)

    # Convert text to bullet points
    left_text = content.get("left_column") or ""
    right_text = content.get("right_column") or ""
    
    left_bullets = _text_to_bullets(left_text)
    right_bullets = _text_to_bullets(right_text)
    
    # Left column card with bullets
    _add_rect(slide, MARGIN, Inches(1.15), col_w, card_height, t["card_bg"])
    left_box = slide.shapes.add_textbox(
        MARGIN + Inches(0.25), Inches(1.3),
        col_w - Inches(0.5), card_height - Inches(0.25)
    )
    left_tf = left_box.text_frame
    left_tf.word_wrap = True
    for i, bullet in enumerate(left_bullets[:5]):
        p = left_tf.paragraphs[0] if i == 0 else left_tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(6)
        p.space_before = Pt(1) if i > 0 else Pt(0)
        run_bullet = p.add_run()
        run_bullet.text = "\u2022 "
        run_bullet.font.size = Pt(FONTS["body"])
        run_bullet.font.bold = True
        run_bullet.font.color.rgb = RGBColor(*t["accent"])
        run_text = p.add_run()
        run_text.text = bullet
        run_text.font.size = Pt(FONTS["body"])
        run_text.font.color.rgb = RGBColor(*t["text_color"])

    # Right column card with bullets
    right_left = MARGIN * 2 + col_w
    _add_rect(slide, right_left, Inches(1.15), col_w, card_height, t["card_bg"])
    right_box = slide.shapes.add_textbox(
        right_left + Inches(0.25), Inches(1.3),
        col_w - Inches(0.5), card_height - Inches(0.25)
    )
    right_tf = right_box.text_frame
    right_tf.word_wrap = True
    for i, bullet in enumerate(right_bullets[:5]):
        p = right_tf.paragraphs[0] if i == 0 else right_tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(6)
        p.space_before = Pt(1) if i > 0 else Pt(0)
        run_bullet = p.add_run()
        run_bullet.text = "\u2022 "
        run_bullet.font.size = Pt(FONTS["body"])
        run_bullet.font.bold = True
        run_bullet.font.color.rgb = RGBColor(*t["accent2"])
        run_text = p.add_run()
        run_text.text = bullet
        run_text.font.size = Pt(FONTS["body"])
        run_text.font.color.rgb = RGBColor(*t["text_color"])

    # Vertical divider
    _add_rect(
        slide,
        MARGIN * 2 + col_w - Inches(0.02), Inches(1.2),
        Inches(0.04), card_height - Inches(0.1),
        t["accent"],
    )


def _build_quote_slide(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])

    # Large accent quote mark background
    _add_rect(slide, Inches(0.5), Inches(0.8), Inches(1.2), Inches(1.5), t["accent"])
    _add_textbox(
        slide,
        Inches(0.55), Inches(0.6),
        Inches(1.2), Inches(1.5),
        "\u201c", 96, t["background"],
        bold=True, align=PP_ALIGN.CENTER,
    )

    content = data.get("content") or {}
    quote = content.get("quote") or data.get("title") or ""
    attribution = content.get("attribution") or ""

    # Centered quote with adaptive spacing
    quote_height = Inches(2.8) if len(quote) > 200 else Inches(2.4)
    quote_top = Inches(1.8) + (Inches(2.6) - quote_height) / 2
    
    _add_textbox(
        slide,
        MARGIN, quote_top,
        W - MARGIN * 2, quote_height,
        quote, FONTS["quote"], t["title_color"],
        bold=True, align=PP_ALIGN.CENTER,
    )

    if attribution:
        _add_textbox(
            slide,
            MARGIN, Inches(5.2),
            W - MARGIN * 2, Inches(0.8),
            attribution, FONTS["body"], t["accent"],
            bold=False, align=PP_ALIGN.CENTER,
        )

    _add_rect(slide, MARGIN * 4, H - Inches(0.8), W - MARGIN * 8, Inches(0.05), t["accent2"])


def _build_data_summary(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])
    _add_rect(slide, 0, 0, W, Inches(0.08), t["accent"])

    title = data.get("title", "Key Metrics")
    _add_textbox(
        slide,
        MARGIN, Inches(0.15),
        W - MARGIN * 2, Inches(0.8),
        title, FONTS["heading"], t["title_color"],
        bold=True,
    )

    content = data.get("content") or {}
    stats = (content.get("stats") or [])[:3]

    card_w = (W - MARGIN * 4) / 3
    card_height = H - Inches(1.15)
    for i, stat in enumerate(stats):
        left = MARGIN + i * (card_w + MARGIN)
        # Card background
        _add_rect(slide, left, Inches(1.05), card_w, card_height, t["card_bg"])
        # Accent top bar on card
        _add_rect(slide, left, Inches(1.05), card_w, Inches(0.1), t["accent"] if i % 2 == 0 else t["accent2"])
        # Stat value
        _add_textbox(
            slide,
            left + Inches(0.15), Inches(1.5),
            card_w - Inches(0.3), Inches(1.2),
            stat.get("value", "—"), FONTS["stat_val"], t["accent"] if i % 2 == 0 else t["accent2"],
            bold=True, align=PP_ALIGN.CENTER,
        )
        # Label
        _add_textbox(
            slide,
            left + Inches(0.15), Inches(2.85),
            card_w - Inches(0.3), Inches(0.6),
            stat.get("label", ""), FONTS["stat_lbl"] + 2, t["title_color"],
            bold=True, align=PP_ALIGN.CENTER,
        )
        # Description
        _add_textbox(
            slide,
            left + Inches(0.15), Inches(3.6),
            card_w - Inches(0.3), Inches(1.2),
            stat.get("description", ""), FONTS["caption"], t["subtitle_color"],
            align=PP_ALIGN.CENTER,
        )


def _build_closing_slide(slide, data: dict, theme: dict):
    t = theme
    _fill_bg(slide, t["background"])

    # Full accent strip at top
    _add_rect(slide, 0, 0, W, Inches(0.5), t["accent"])
    # Full accent strip at bottom
    _add_rect(slide, 0, H - Inches(0.5), W, Inches(0.5), t["accent2"])

    title = data.get("title", "Thank You!")
    _add_textbox(
        slide,
        MARGIN, Inches(2.0),
        W - MARGIN * 2, Inches(1.5),
        title, FONTS["title"] + 4, t["title_color"],
        bold=True, align=PP_ALIGN.CENTER,
    )

    content = data.get("content") or {}
    cta = content.get("cta") or data.get("subtitle") or ""
    if cta:
        _add_textbox(
            slide,
            MARGIN, Inches(3.7),
            W - MARGIN * 2, Inches(1.2),
            cta, FONTS["body"] + 2, t["subtitle_color"],
            align=PP_ALIGN.CENTER,
        )

    _add_rect(slide, W / 2 - Inches(1.5), Inches(3.0), Inches(3.0), Inches(0.06), t["accent"])


# ── Layout dispatcher ────────────────────────────────────────────────────────

LAYOUT_BUILDERS = {
    "title_slide":    _build_title_slide,
    "section_header": _build_section_header,
    "bullet_points":  _build_bullet_points,
    "two_column":     _build_two_column,
    "quote_slide":    _build_quote_slide,
    "data_summary":   _build_data_summary,
    "closing_slide":  _build_closing_slide,
}


# ═══════════════════════════════════════════════════════════════════════════
# Public: Build PPTX
# ═══════════════════════════════════════════════════════════════════════════

def build_pptx(
    slides_data: list[dict],
    audio_paths: list[str | None],
    theme_name: str,
    out_path: str,
    progress_cb=None,
) -> str:
    """
    Build the final .pptx file from slide data and audio paths.
    Returns the path to the saved file.
    """
    # Clean text and ensure every layout has visible body content
    slides_data = normalize_slides(clean_data(slides_data))

    # Dynamically select theme if Auto-Select or Random is chosen
    topic_str = ""
    if slides_data:
        topic_str = slides_data[0].get("title", "")
    resolved_theme = get_dynamic_theme(theme_name, topic_str)
    theme = THEMES.get(resolved_theme, THEMES["Dark Pro"])
    
    prs = _new_prs()

    for i, slide_data in enumerate(slides_data):
        layout_name = slide_data.get("layout", "bullet_points")
        builder_fn  = LAYOUT_BUILDERS.get(layout_name)
        if builder_fn is None:
            layout_name = "bullet_points"
            builder_fn = _build_bullet_points

        if progress_cb:
            progress_cb(f"Building slide {i+1}/{len(slides_data)}: {slide_data.get('title','')[:40]} ...")

        slide = _blank_slide(prs)

        # Build visual layout
        builder_fn(slide, slide_data, theme)

        # Add speaker notes to notes pane
        _add_notes(slide, slide_data.get("speaker_notes", ""))

        # Add transition animation (alternate different animations)
        transitions = ["fade", "push", "wipe", "split", "zoom", "wheel", "comb", "checker"]
        transition_type = transitions[i % len(transitions)]
        _add_transition(slide, transition_type=transition_type)

        # Embed audio (auto-play narration)
        audio_path = audio_paths[i] if i < len(audio_paths) else None
        if audio_path and os.path.exists(audio_path):
            _embed_audio(slide, audio_path)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)

    if progress_cb:
        progress_cb(f"Saved: {out_path}")

    return out_path
