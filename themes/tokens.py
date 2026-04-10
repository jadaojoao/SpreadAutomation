# themes/tokens.py
# Design System tokens translated from OkLch → hex for CustomTkinter.
# Source: DESIGN_SYSTEM_EXPORT.md (cvm_repots_capture, 2026-04-10)
# Format: DS["token"] = (light_hex, dark_hex)

DS = {
    # Surfaces
    "background":          ("#F5F3ED", "#1A1A1A"),
    "foreground":          ("#2C2920", "#FBFBFB"),
    "card":                ("#FDFCF9", "#282828"),
    "card_foreground":     ("#2C2920", "#FBFBFB"),
    "popover":             ("#FDFCF9", "#282828"),
    "popover_foreground":  ("#2C2920", "#FBFBFB"),

    # Brand
    "primary":             ("#2E6B52", "#E8E8E8"),
    "primary_foreground":  ("#FAFAF5", "#282828"),

    # Neutral tones
    "secondary":           ("#E6DBBA", "#353535"),
    "secondary_foreground":("#302B1E", "#FBFBFB"),
    "muted":               ("#EDEAE2", "#353535"),
    "muted_foreground":    ("#737060", "#999999"),
    "accent":              ("#EBE0CD", "#353535"),
    "accent_foreground":   ("#2F2B1F", "#FBFBFB"),

    # Semantic
    "destructive":         ("#D93B28", "#E8583A"),
    "border":              ("#DEDAD2", "#3A3A3A"),
    "input":               ("#DEDAD2", "#3A3A3A"),
    "ring":                ("#4F8E74", "#6E6E6E"),

    # Hover / active states (derived)
    "primary_hover":       ("#24583F", "#C8C8C8"),
    "secondary_hover":     ("#D8CFA8", "#454545"),
    "ghost_hover":         ("#EDEAE2", "#2A2A2A"),
    "destructive_hover":   ("#B82F1E", "#CC4A30"),

    # Chart palette (single value, same light/dark)
    "chart_1":             ("#4F9B78", "#5BB88F"),   # teal — primary/brand
    "chart_2":             ("#D4823A", "#E8953D"),   # amber — complementary
    "chart_3":             ("#6B7ED4", "#7A90E0"),   # blue-violet
    "chart_4":             ("#8CB84A", "#A2CE55"),   # lime
    "chart_5":             ("#C45AA0", "#D46AB5"),   # rose

    # Scrollbar
    "scrollbar":           ("#C8C4BC", "#4A4A4A"),
    "scrollbar_hover":     ("#A8A49C", "#5A5A5A"),
}

# Geometry
RADIUS = {
    "sm":  6,
    "md":  8,
    "base": 10,
    "lg":  14,
    "xl":  16,
    "2xl": 20,
}

BORDER_WIDTH = 1

# Font helpers — CTkFont uses system fonts; map DS roles to weight/size pairs.
# DS heading: Space Grotesk bold → CTkFont bold
# DS body: Manrope → CTkFont normal
# DS mono: IBM Plex Mono → CTkFont Courier
FONT = {
    "h1":      {"size": 22, "weight": "bold"},
    "h2":      {"size": 17, "weight": "bold"},
    "h3":      {"size": 14, "weight": "bold"},
    "body":    {"size": 13, "weight": "normal"},
    "body_sm": {"size": 12, "weight": "normal"},
    "caption": {"size": 11, "weight": "normal"},
    "mono":    {"size": 11, "weight": "normal", "family": "Courier"},
    "label":   {"size": 12, "weight": "bold"},
}


def light(token: str) -> str:
    """Return the light-mode hex for a DS token."""
    return DS[token][0]


def dark(token: str) -> str:
    """Return the dark-mode hex for a DS token."""
    return DS[token][1]


def pair(token: str) -> tuple[str, str]:
    """Return (light, dark) tuple for a DS token."""
    return DS[token]
