"""
normalize.py - Validate and repair slide outlines so every layout has visible content.
"""

from __future__ import annotations

import re
from typing import Any

VALID_LAYOUTS = frozenset({
    "title_slide",
    "section_header",
    "bullet_points",
    "two_column",
    "quote_slide",
    "data_summary",
    "closing_slide",
})

LAYOUT_ALIASES = {
    "title": "title_slide",
    "title_slide": "title_slide",
    "opening": "title_slide",
    "section": "section_header",
    "section_header": "section_header",
    "section_divider": "section_header",
    "divider": "section_header",
    "bullets": "bullet_points",
    "bullet": "bullet_points",
    "bullet_points": "bullet_points",
    "bulletpoints": "bullet_points",
    "content": "bullet_points",
    "two_column": "two_column",
    "twocolumn": "two_column",
    "two_columns": "two_column",
    "comparison": "two_column",
    "columns": "two_column",
    "quote": "quote_slide",
    "quote_slide": "quote_slide",
    "pull_quote": "quote_slide",
    "data": "data_summary",
    "data_summary": "data_summary",
    "stats": "data_summary",
    "metrics": "data_summary",
    "kpi": "data_summary",
    "closing": "closing_slide",
    "closing_slide": "closing_slide",
    "thank_you": "closing_slide",
    "end": "closing_slide",
}


def _normalize_layout_key(name: str | None) -> str:
    if not name:
        return "bullet_points"
    key = re.sub(r"[^a-z0-9]+", "_", str(name).strip().lower()).strip("_")
    if key in VALID_LAYOUTS:
        return key
    return LAYOUT_ALIASES.get(key, "bullet_points")


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        parts = [_as_text(v) for v in value]
        return "\n".join(p for p in parts if p)
    if isinstance(value, dict):
        for k in ("text", "body", "content", "value", "label", "title"):
            if k in value and value[k]:
                return _as_text(value[k])
    return str(value).strip()


def _as_bullet_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        lines = re.split(r"[\n;]+", value)
        items = [ln.strip() for ln in lines if ln.strip()]
        if len(items) <= 1 and len(value) > 80:
            items = re.split(r"(?<=[.!?])\s+", value)
            items = [s.strip() for s in items if s.strip()]
        return items[:8]
    if isinstance(value, list):
        out = []
        for item in value:
            t = _as_text(item)
            if t:
                out.append(t)
        return out[:8]
    t = _as_text(value)
    return [t] if t else []


def _sentences_from_notes(notes: str, max_items: int = 6) -> list[str]:
    if not notes:
        return []
    parts = re.split(r"(?<=[.!?])\s+", notes.strip())
    parts = [p.strip() for p in parts if len(p.strip()) > 10]
    return parts[:max_items]


def _is_placeholder_title(text: str | None) -> bool:
    if not text:
        return False
    text = text.strip().lower()
    # Reject obvious placeholders: "slide 5", "section 3", "page 2", etc.
    if re.match(r"^(slide|section|page)[\s_-]*\d+$", text):
        return True
    # Also reject very short generic titles
    if text in ("title", "content", "overview", "slide", "section", "page"):
        return True
    return False


def _derive_title(layout: str, subtitle: str, bullets: list[str], notes: str, slide_index: int, raw_title: str = "") -> str:
    # First check if raw_title from LLM is valid and not a placeholder
    if raw_title and not _is_placeholder_title(raw_title) and len(raw_title.strip()) > 3:
        return raw_title
    # Try subtitle
    if subtitle and not _is_placeholder_title(subtitle):
        return subtitle
    # Try first bullet
    for candidate in bullets:
        if candidate and not _is_placeholder_title(candidate) and len(candidate.strip()) > 3:
            return candidate
    # Try extracting from notes
    note_title = _sentences_from_notes(notes, max_items=1)
    if note_title:
        candidate = note_title[0]
        if candidate and not _is_placeholder_title(candidate) and len(candidate.strip()) > 3:
            return candidate
    # Fallback to layout-appropriate defaults
    defaults = {
        "title_slide": "Presentation Title",
        "section_header": "Section Overview",
        "bullet_points": "Key Points & Insights",
        "two_column": "Side-by-Side Comparison",
        "quote_slide": "Important Quote",
        "data_summary": "Key Metrics & Analytics",
        "closing_slide": "Thank You",
    }
    return defaults.get(layout, f"Key Insight {slide_index + 1}")


def _split_bullets_halves(bullets: list[str]) -> tuple[str, str]:
    if not bullets:
        return "", ""
    mid = max(1, len(bullets) // 2)
    left = "\n\n".join(f"• {b}" for b in bullets[:mid])
    right = "\n\n".join(f"• {b}" for b in bullets[mid:])
    return left, right


def _bullets_from_content(content: dict) -> list[str]:
    for key in ("bullets", "points", "items", "list", "key_points"):
        found = _as_bullet_list(content.get(key))
        if found:
            return found
    text = _as_text(content.get("text") or content.get("body"))
    return _as_bullet_list(text)


def _stats_from_bullets(bullets: list[str]) -> list[dict]:
    stats = []
    for i, bullet in enumerate(bullets[:3]):
        m = re.match(r"^([^:–—-]{1,40})[:–—-]\s*(.+)$", bullet)
        if m:
            label, rest = m.group(1).strip(), m.group(2).strip()
            value = rest.split(".")[0][:24]
            desc = rest if len(rest) > len(value) else ""
        else:
            words = bullet.split()
            value = " ".join(words[:3]) if words else str(i + 1)
            label = f"Insight {i + 1}"
            desc = bullet
        stats.append({
            "label": label[:40],
            "value": value[:28],
            "description": (desc or bullet)[:120],
        })
    while len(stats) < 3:
        n = len(stats) + 1
        stats.append({
            "label": f"Key point {n}",
            "value": "—",
            "description": "See speaker notes for details.",
        })
    return stats[:3]


def _normalize_stat(stat: Any, index: int) -> dict:
    if isinstance(stat, dict):
        return {
            "label": _as_text(stat.get("label")) or f"Metric {index + 1}",
            "value": _as_text(stat.get("value")) or "—",
            "description": _as_text(stat.get("description") or stat.get("desc") or stat.get("text")),
        }
    text = _as_text(stat)
    return {
        "label": f"Metric {index + 1}",
        "value": text[:28] if text else "—",
        "description": text,
    }


def _ensure_content_dict(slide: dict) -> dict:
    content = slide.get("content")
    if content is None:
        slide["content"] = {}
        return slide["content"]
    if isinstance(content, str):
        slide["content"] = {"text": content}
        return slide["content"]
    if not isinstance(content, dict):
        slide["content"] = {"text": _as_text(content)}
        return slide["content"]
    return content


def normalize_slide(slide: dict, slide_index: int, total: int) -> dict:
    """Repair a single slide so its layout has populated visible fields."""
    slide = dict(slide)
    layout = _normalize_layout_key(slide.get("layout"))

    if slide_index == 0:
        layout = "title_slide"
    elif slide_index == total - 1:
        layout = "closing_slide"

    slide["layout"] = layout
    content = _ensure_content_dict(slide)

    raw_title = _as_text(slide.get("title"))
    subtitle = _as_text(slide.get("subtitle"))
    notes = _as_text(slide.get("speaker_notes"))
    bullets = _bullets_from_content(content)

    if not bullets and notes:
        bullets = _sentences_from_notes(notes)
    if not bullets and subtitle:
        bullets = _as_bullet_list(subtitle)
    if not bullets and raw_title and layout not in ("title_slide", "closing_slide") and not _is_placeholder_title(raw_title):
        bullets = [raw_title]
    
    # Derive title with all available info, including raw_title as a fallback candidate
    title = _derive_title(layout, subtitle, bullets, notes, slide_index, raw_title)

    if layout == "title_slide":
        slide["title"] = title
        if not subtitle:
            subtitle = _sentences_from_notes(notes, max_items=1)
            subtitle = subtitle[0] if subtitle else _as_text(content.get("subtitle"))
        slide["subtitle"] = subtitle or None

    elif layout == "section_header":
        slide["title"] = title
        if not subtitle:
            subtitle = bullets[0] if bullets else _sentences_from_notes(notes, max_items=1)
            subtitle = subtitle[0] if isinstance(subtitle, list) else subtitle
        slide["subtitle"] = _as_text(subtitle) or None

    elif layout == "bullet_points":
        slide["title"] = title
        if len(bullets) < 2 and notes:
            bullets = _sentences_from_notes(notes) or bullets
        if len(bullets) < 2:
            bullets = bullets + [f"Further detail on: {title}"]
        content["bullets"] = bullets[:6]

    elif layout == "two_column":
        slide["title"] = title
        left = _as_text(content.get("left_column") or content.get("left") or content.get("column_left"))
        right = _as_text(content.get("right_column") or content.get("right") or content.get("column_right"))

        if not left or not right:
            bl, br = _split_bullets_halves(bullets)
            left = left or bl
            right = right or br

        if not left and notes:
            parts = _sentences_from_notes(notes, max_items=4)
            left, right = _split_bullets_halves(parts)

        if not left:
            left = subtitle or title
        if not right:
            right = notes[:500] if notes else "Additional perspectives and practical implications."

        content["left_column"] = left
        content["right_column"] = right

    elif layout == "quote_slide":
        quote = _as_text(content.get("quote") or content.get("text"))
        if not quote and bullets:
            quote = bullets[0]
        if not quote and notes:
            quote = _sentences_from_notes(notes, max_items=1)
            quote = quote[0] if quote else ""
        if not quote:
            quote = title
        content["quote"] = quote
        if not _as_text(content.get("attribution")):
            content["attribution"] = _as_text(content.get("author")) or subtitle or None
        slide["title"] = title

    elif layout == "data_summary":
        slide["title"] = title
        raw_stats = content.get("stats") or content.get("metrics") or content.get("kpis")
        stats: list[dict] = []
        if isinstance(raw_stats, list):
            for i, s in enumerate(raw_stats[:3]):
                stats.append(_normalize_stat(s, i))
        if len(stats) < 3:
            stats = _stats_from_bullets(bullets or _sentences_from_notes(notes))
        content["stats"] = stats[:3]

    elif layout == "closing_slide":
        slide["title"] = title or "Thank You"
        cta = _as_text(content.get("cta") or content.get("call_to_action"))
        if not cta:
            cta = subtitle or (bullets[0] if bullets else "")
        if not cta and notes:
            sents = _sentences_from_notes(notes, max_items=1)
            cta = sents[0] if sents else ""
        if not cta:
            cta = "Questions? Let's continue the conversation."
        content["cta"] = cta
        slide["subtitle"] = subtitle or None

    slide["content"] = content
    return slide


def normalize_slides(slides_data: list[dict]) -> list[dict]:
    """Normalize an entire deck: layouts, content fields, and fallbacks."""
    if not slides_data:
        return slides_data

    total = len(slides_data)
    normalized = []
    for i, slide in enumerate(slides_data):
        if not isinstance(slide, dict):
            continue
        normalized.append(normalize_slide(slide, i, total))

    if normalized and normalized[0].get("layout") != "title_slide":
        normalized[0]["layout"] = "title_slide"
        normalized[0] = normalize_slide(normalized[0], 0, total)

    if normalized and normalized[-1].get("layout") != "closing_slide":
        normalized[-1]["layout"] = "closing_slide"
        normalized[-1] = normalize_slide(normalized[-1], total - 1, total)

    return normalized
