"""
themes.py - Color palettes and font configurations for PPT slides.

Themes are derived from analysis of 20 professional PPTX templates:
  - CREATIX, ECONIFY, COLLEGIUM (dark corporate)
  - Camelow Red Maroon, Santara, Skylark (hero colors)
  - Global Corporate, Business Strategies (corporate blue/teal)
  - Startup Inspiration Black (deep navy)
  - Management Thesis Light/Dark (academic)
  - Aesthetic, Interior Aesthetics (minimal warm)
  - Macayle (purple/cyan), Tech Startup XL (amber), Robinson (bold red)
"""

from pptx.dml.color import RGBColor

# ── Color helper ─────────────────────────────────────────────────────────────

def _h(hex_str: str) -> tuple:
    """Convert '#RRGGBB' hex string to (R, G, B) int tuple."""
    h = hex_str.lstrip('#')
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ── Color Themes ─────────────────────────────────────────────────────────────
THEMES = {

    # ─── Original built-in themes ─────────────────────────────────────────

    "Dark Pro": {
        "background":     (12,  14,  30),
        "title_color":    (255, 255, 255),
        "text_color":     (210, 215, 235),
        "accent":         (99,  102, 241),   # indigo
        "accent2":        (236, 72,  153),   # pink
        "subtitle_color": (160, 165, 200),
        "card_bg":        (25,  28,  55),
        "divider":        (50,  55,  100),
    },
    "Light Corporate": {
        "background":     (245, 247, 250),
        "title_color":    (15,  23,  42),
        "text_color":     (51,  65,  85),
        "accent":         (14,  165, 233),   # sky blue
        "accent2":        (16,  185, 129),   # emerald
        "subtitle_color": (100, 116, 139),
        "card_bg":        (255, 255, 255),
        "divider":        (203, 213, 225),
    },
    "Vibrant Creative": {
        "background":     (15,  10,  40),
        "title_color":    (255, 255, 255),
        "text_color":     (228, 225, 245),
        "accent":         (168, 85,  247),   # purple
        "accent2":        (251, 146, 60),    # orange
        "subtitle_color": (196, 181, 253),
        "card_bg":        (30,  20,  65),
        "divider":        (80,  40,  120),
    },
    "Minimal White": {
        "background":     (255, 255, 255),
        "title_color":    (17,  24,  39),
        "text_color":     (75,  85,  99),
        "accent":         (59,  130, 246),   # blue
        "accent2":        (239, 68,  68),    # red
        "subtitle_color": (107, 114, 128),
        "card_bg":        (243, 244, 246),
        "divider":        (229, 231, 235),
    },

    # ─── Derived from CREATIX / ECONIFY / COLLEGIUM templates ─────────────
    # Dark slate background, golden-yellow + crimson accents, bold corporate feel

    "Creatix Dark": {
        "background":     _h("#0E1117"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#D4D8E2"),
        "accent":         _h("#E5B31B"),     # golden yellow
        "accent2":        _h("#C82837"),     # crimson
        "subtitle_color": _h("#A0A8B8"),
        "card_bg":        _h("#1A1E2C"),
        "divider":        _h("#2D3347"),
    },

    # ─── Derived from Camelow Red Maroon template ──────────────────────────
    # Deep navy slate bg, dark maroon + bright coral accents

    "Red Maroon": {
        "background":     _h("#1D2530"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#D3D1AA"),
        "accent":         _h("#701524"),     # deep maroon
        "accent2":        _h("#EC5132"),     # vivid coral
        "subtitle_color": _h("#B4B3B2"),
        "card_bg":        _h("#263040"),
        "divider":        _h("#394656"),
    },

    # ─── Derived from Santara template ────────────────────────────────────
    # Warm olive-green dark bg, amber + terracotta accents

    "Earthy Warmth": {
        "background":     _h("#2A2F22"),
        "title_color":    _h("#FFF8EE"),
        "text_color":     _h("#CCDDEA"),
        "accent":         _h("#E48312"),     # amber
        "accent2":        _h("#BD582C"),     # terracotta
        "subtitle_color": _h("#9B8357"),
        "card_bg":        _h("#3B4130"),
        "divider":        _h("#637052"),
    },

    # ─── Derived from Skylark template ────────────────────────────────────
    # Warm neutral light bg, deep navy + sky-blue accents

    "Skylark Blue": {
        "background":     _h("#FFFAED"),
        "title_color":    _h("#1A1915"),
        "text_color":     _h("#3D3830"),
        "accent":         _h("#013A85"),     # deep navy
        "accent2":        _h("#29B8F4"),     # sky blue
        "subtitle_color": _h("#5D564D"),
        "card_bg":        _h("#F0EBD8"),
        "divider":        _h("#DCD6CA"),
    },

    # ─── Derived from Global Corporate template ────────────────────────────
    # Clean white bg, teal + deep blue accents, Roboto/Open Sans

    "Global Teal": {
        "background":     _h("#FFFFFF"),
        "title_color":    _h("#0A2033"),
        "text_color":     _h("#2D4A5E"),
        "accent":         _h("#3CAFA4"),     # teal
        "accent2":        _h("#2474B2"),     # deep blue
        "subtitle_color": _h("#5A7A90"),
        "card_bg":        _h("#F0F8FA"),
        "divider":        _h("#C8DDE8"),
    },

    # ─── Derived from Business Strategies template ─────────────────────────
    # Dark navy bg, steel blue + warm orange + sky cyan accents

    "Business Steel": {
        "background":     _h("#07263A"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#D3D3D3"),
        "accent":         _h("#F7931F"),     # warm orange
        "accent2":        _h("#4CC1EF"),     # sky cyan
        "subtitle_color": _h("#8AA7BB"),
        "card_bg":        _h("#0E3A50"),
        "divider":        _h("#1B5070"),
    },

    # ─── Derived from Startup Inspiration Black template ──────────────────
    # Ultra-deep midnight navy, dark accent tones with white text

    "Midnight Navy": {
        "background":     _h("#00072D"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#C8D4E8"),
        "accent":         _h("#2255CC"),     # bright navy blue
        "accent2":        _h("#0A9AF0"),     # electric blue
        "subtitle_color": _h("#7A9ACA"),
        "card_bg":        _h("#0A1240"),
        "divider":        _h("#0C2E3A"),
    },

    # ─── Derived from Management Thesis Light template ────────────────────
    # Clean light lavender bg, deep blue + cyan accents, Inter/Roboto

    "Academic Blue": {
        "background":     _h("#E4E8F6"),
        "title_color":    _h("#2D3847"),
        "text_color":     _h("#363D6E"),
        "accent":         _h("#15579D"),     # deep blue
        "accent2":        _h("#00A4E6"),     # bright cyan
        "subtitle_color": _h("#4A6080"),
        "card_bg":        _h("#FFFFFF"),
        "divider":        _h("#B8C8E8"),
    },

    # ─── Derived from Management Thesis Dark template ─────────────────────
    # Deep indigo bg, blue + green accent tones

    "Academic Dark": {
        "background":     _h("#363D6E"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#D0D8F0"),
        "accent":         _h("#00A4E6"),     # bright cyan
        "accent2":        _h("#45B653"),     # green
        "subtitle_color": _h("#90A4D8"),
        "card_bg":        _h("#2D3565"),
        "divider":        _h("#1E2650"),
    },

    # ─── Derived from Aesthetic Google Slides template ────────────────────
    # Warm cream/beige bg, dusty rose + mauve accents, Open Sans / Antic Didone

    "Warm Aesthete": {
        "background":     _h("#FCF8E8"),
        "title_color":    _h("#3D2020"),
        "text_color":     _h("#5A3A3A"),
        "accent":         _h("#815B5B"),     # dusty rose
        "accent2":        _h("#C4956A"),     # warm tan
        "subtitle_color": _h("#9C7070"),
        "card_bg":        _h("#F3EEEA"),
        "divider":        _h("#E0D0C0"),
    },

    # ─── Derived from Interior Aesthetics template ────────────────────────
    # Soft off-white bg, sage green + golden accents, Inter/Helvetica

    "Nature Studio": {
        "background":     _h("#F5F2EC"),
        "title_color":    _h("#1E2820"),
        "text_color":     _h("#3A4840"),
        "accent":         _h("#597460"),     # sage green
        "accent2":        _h("#E3D27D"),     # warm gold
        "subtitle_color": _h("#71937A"),
        "card_bg":        _h("#EDEAE0"),
        "divider":        _h("#BEC8BC"),
    },

    # ─── Derived from Macayle template ────────────────────────────────────
    # Deep purple bg, magenta + electric cyan, dramatic

    "Amethyst": {
        "background":     _h("#1A051A"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#EAE5EB"),
        "accent":         _h("#6D1D6B"),     # deep purple/magenta
        "accent2":        _h("#00B0F0"),     # electric cyan
        "subtitle_color": _h("#C0A0C0"),
        "card_bg":        _h("#2C0E2C"),
        "divider":        _h("#632E62"),
    },

    # ─── Derived from Tech Startup XL template ────────────────────────────
    # Dark charcoal bg, warm amber accent — tech startup feel

    "Tech Amber": {
        "background":     _h("#121212"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#EEEEEE"),
        "accent":         _h("#FFAB40"),     # warm amber
        "accent2":        _h("#78909C"),     # steel grey
        "subtitle_color": _h("#AAAAAA"),
        "card_bg":        _h("#1E1E1E"),
        "divider":        _h("#333333"),
    },

    # ─── Derived from Robinson template ───────────────────────────────────
    # Stark black/white bg, bold red/magenta accent — editorial bold

    "Bold Red": {
        "background":     _h("#0A0A0A"),
        "title_color":    _h("#FFFFFF"),
        "text_color":     _h("#E8E8E8"),
        "accent":         _h("#E91A4D"),     # bold magenta-red
        "accent2":        _h("#B6153D"),     # deep crimson
        "subtitle_color": _h("#C3C3C3"),
        "card_bg":        _h("#1A1A1A"),
        "divider":        _h("#434343"),
    },

    # ─── Derived from Blue and White template ─────────────────────────────
    # Professional corporate blue, Poppins font feel

    "Corporate Navy": {
        "background":     _h("#F5F8FF"),
        "title_color":    _h("#0A225A"),
        "text_color":     _h("#1E3A6E"),
        "accent":         _h("#00569E"),     # corporate blue
        "accent2":        _h("#2E66AF"),     # medium blue
        "subtitle_color": _h("#4378BD"),
        "card_bg":        _h("#FFFFFF"),
        "divider":        _h("#C0D0E8"),
    },

    # ─── Derived from Red and White template ──────────────────────────────
    # Clean white bg, bold red, Poppins / League Spartan feel

    "Bold Vermillion": {
        "background":     _h("#FFFFFF"),
        "title_color":    _h("#0A0A0A"),
        "text_color":     _h("#2A2A2A"),
        "accent":         _h("#FF1616"),     # vivid red
        "accent2":        _h("#222222"),     # near black
        "subtitle_color": _h("#555555"),
        "card_bg":        _h("#F5F5F5"),
        "divider":        _h("#DDDDDD"),
    },
}


# ── Font Sizes (in points) ───────────────────────────────────────────────────
FONTS = {
    "title":    48,
    "subtitle": 26,
    "heading":  36,
    "body":     22,
    "caption": 18,
    "stat_val": 48,
    "stat_lbl": 18,
    "quote":    36,
}


def rgb(theme: dict, key: str) -> RGBColor:
    """Return an RGBColor from a theme tuple."""
    return RGBColor(*theme[key])
