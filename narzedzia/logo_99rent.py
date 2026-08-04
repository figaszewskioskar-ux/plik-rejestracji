# -*- coding: utf-8 -*-
"""Generuje logo 99rent wg przesłanego wzoru: czerwony kwadrat, dwie białe
'9' w stylu pinezek (koło + prosty ścięty ogon) i gruby, zwarty napis RENT."""
from PIL import Image, ImageDraw, ImageFont

RED = (237, 28, 36)
S = 1024
SS = 6  # supersampling

FONT = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"


def nine(dr, cx, cy, R, tw, tail_y, color, grow=0):
    """Pinezka '9': koło + prosty ogon przy prawej krawędzi, ścięty ukośnie."""
    Rg = R + grow
    dr.ellipse([cx - Rg, cy - Rg, cx + Rg, cy + Rg], fill=color)
    x2 = cx + R + grow
    x1 = x2 - tw - 2 * grow
    y2 = tail_y + grow
    dr.polygon([(x1, cy), (x2, cy),
                (x2, y2), (x1, y2 - int(tw * 0.85))], fill=color)


def hole(dr, cx, cy, r):
    dr.ellipse([cx - r, cy - r, cx + r, cy + r], fill=RED)


def main(out="logo99rent.png", out_small="logo99rent_small.png"):
    W = S * SS
    img = Image.new("RGB", (W, W), RED)
    d = ImageDraw.Draw(img)

    R = int(W * 0.165)
    hr = int(R * 0.42)
    tw = int(R * 0.50)
    cy = int(W * 0.285)
    tail_y = int(W * 0.60)
    c1 = int(W * 0.315)
    c2 = c1 + int(R * 1.86)
    gap = int(W * 0.010)

    nine(d, c1, cy, R, tw, tail_y, "white")
    hole(d, c1, cy, hr)
    nine(d, c2, cy, R, tw, tail_y, RED, grow=gap)   # szczelina
    nine(d, c2, cy, R, tw, tail_y, "white")
    hole(d, c2, cy, hr)

    # RENT — zwarty, bardzo ciężki
    f = ImageFont.truetype(FONT, int(W * 0.235))
    sw = int(W * 0.013)
    text = "RENT"
    widths = [d.textlength(ch, font=f) + 2 * sw for ch in text]
    span = W * 0.72
    gap_l = (span - sum(widths)) / (len(text) - 1)
    x = (W - span) / 2 + sw
    y = int(W * 0.815)
    for ch, cw in zip(text, widths):
        d.text((x, y), ch, font=f, fill="white", anchor="lm",
               stroke_width=sw, stroke_fill="white")
        x += cw + gap_l

    img = img.resize((S, S), Image.LANCZOS)
    img.save(out)
    img.resize((256, 256), Image.LANCZOS).save(out_small)
    print("logo zapisane:", out, out_small)


if __name__ == "__main__":
    main()
