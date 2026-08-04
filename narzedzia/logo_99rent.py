# -*- coding: utf-8 -*-
"""Generuje logo 99rent (czerwony kwadrat, białe '99' w stylu pinezek + RENT)."""
from PIL import Image, ImageDraw, ImageFont

RED = (227, 6, 19)
S = 512


def font(sz):
    return ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", sz)


def nine_shapes(dr, cx, cy, R, tw, color, grow=0):
    Rg = R + grow
    tail_len = int(R * 1.85) + grow
    dr.ellipse([cx - Rg, cy - Rg, cx + Rg, cy + Rg], fill=color)
    x1 = cx + R - tw - grow
    x2 = cx + R + grow
    dr.polygon([(x1, cy - 10), (x2, cy - 10), (x2, cy + tail_len),
                (x1, cy + tail_len - int(tw * 0.75))], fill=color)


def punch_hole(dr, cx, cy, hole):
    dr.ellipse([cx - hole, cy - hole, cx + hole, cy + hole], fill=RED)


def main(out="logo99rent.png", out_small="logo99rent_small.png"):
    img = Image.new("RGB", (S, S), RED)
    d = ImageDraw.Draw(img)
    R, hole, tw = 84, 35, 33
    cy, c1, c2 = 150, 152, 296
    nine_shapes(d, c1, cy, R, tw, "white")
    punch_hole(d, c1, cy, hole)
    nine_shapes(d, c2, cy, R, tw, RED, grow=8)   # czerwony kontur oddzielający
    nine_shapes(d, c2, cy, R, tw, "white")
    punch_hole(d, c2, cy, hole)
    d.text((S / 2 + 2, 442), "RENT", font=font(100), fill="white", anchor="mm")
    img.save(out)
    img.resize((128, 128), Image.LANCZOS).save(out_small)


if __name__ == "__main__":
    main()
