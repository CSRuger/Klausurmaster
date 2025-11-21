"""Color helpers to keep UI styling logic centralized."""

import colorsys
import random


# Generates a pastel color that keeps the UI subtle and readable.
def random_pastel_color() -> str:
    """Return a soft tone while covering ~200% more hue/saturation range."""

    hue = random.random()  # full 360° hue range
    saturation = random.uniform(0.35, 0.75)  # widened saturation band
    value = random.uniform(0.75, 0.95)  # keep it bright for readability
    r, g, b = colorsys.hsv_to_rgb(hue, saturation, value)
    return f"#{int(r * 255):02x}{int(g * 255):02x}{int(b * 255):02x}"


# Converts a hexadecimal color string into an RGB tuple.
def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


# Converts RGB values back to a hexadecimal color string.
def rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"#{r:02x}{g:02x}{b:02x}"


# Calculates the intermediate color between two colors based on the factor (0..1).
def interpolate_color(start_color: str, end_color: str, factor: float) -> str:
    r1, g1, b1 = hex_to_rgb(start_color)
    r2, g2, b2 = hex_to_rgb(end_color)

    r = int(r1 + (r2 - r1) * factor)
    g = int(g1 + (g2 - g1) * factor)
    b = int(b1 + (b2 - b1) * factor)

    return rgb_to_hex(r, g, b)
