"""
Generates assets/logo.ico: a single bold document icon used for the app's
title bar, taskbar icon, and the built exe's file icon (everywhere except
the in-app header, which intentionally keeps the plain EDMS "ED" mark -
see assets/logo.png).

An earlier version showed JSON braces -> arrow -> document as two cards,
but that much detail blurred into an unreadable blob once actually shrunk
to real icon sizes (16-32px, i.e. the taskbar/title bar). This single bold
document shape stays legible even at 16px.

Colors are sampled from assets/source/edms-icon-source.jpg so this stays
visually consistent with the official EDMS red. Re-run this script (from
the repo root: python assets/source/generate_databridge_icon.py) any time
the design needs tweaking.
"""

from pathlib import Path

from PIL import Image, ImageDraw

REPO_ROOT = Path(__file__).parent.parent.parent
TOP_RED = (227, 46, 39)
BOTTOM_RED = (181, 17, 24)


def render(size: int) -> Image.Image:
    ss = 4  # supersample for anti-aliasing, downscaled at the end
    px = size * ss
    im = Image.new("RGB", (px, px))
    draw = ImageDraw.Draw(im)

    for y in range(px):
        t = y / px
        r = int(TOP_RED[0] + (BOTTOM_RED[0] - TOP_RED[0]) * t)
        g = int(TOP_RED[1] + (BOTTOM_RED[1] - TOP_RED[1]) * t)
        b = int(TOP_RED[2] + (BOTTOM_RED[2] - TOP_RED[2]) * t)
        draw.line([(0, y), (px, y)], fill=(r, g, b))

    cx = cy = px / 2
    ring_r = px * 0.44
    draw.ellipse(
        [cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r],
        outline="white", width=int(px * 0.026),
    )

    # A single bold document: folded corner + a few thick text lines
    card_w, card_h = px * 0.40, px * 0.50
    x0, y0 = cx - card_w / 2, cy - card_h / 2
    x1, y1 = cx + card_w / 2, cy + card_h / 2
    fold = card_w * 0.28
    draw.polygon(
        [(x0, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1), (x0, y1)],
        fill="white",
    )
    draw.polygon(
        [(x1 - fold, y0), (x1, y0 + fold), (x1 - fold, y0 + fold)],
        fill=TOP_RED,
    )

    line_w = int(px * 0.028)
    pad = card_w * 0.16
    ly0 = y0 + card_h * 0.48
    for i in range(3):
        y = ly0 + i * card_h * 0.16
        x1b = x1 - pad if i != 2 else x1 - pad - card_w * 0.14
        draw.line([(x0 + pad, y), (x1b, y)], fill=TOP_RED, width=line_w)

    return im.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    full = render(1024)
    full.save(REPO_ROOT / "assets" / "source" / "databridge-icon-generated.png")
    full.save(
        REPO_ROOT / "assets" / "logo.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    print("Wrote assets/logo.ico and assets/source/databridge-icon-generated.png")
