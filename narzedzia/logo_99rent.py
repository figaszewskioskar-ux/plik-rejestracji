# -*- coding: utf-8 -*-
"""Generuje logo 99rent — czerwony kwadrat, dwie białe pinezki-'9' + napis RENT
(odwzorowanie przesłanego logo)."""
from PIL import Image, ImageDraw, ImageFont

RED = (228, 13, 24)
S = 1024
SS = 4  # supersampling


def font(sz):
    return ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", sz)


def nine(dr, cx, cy, R, tw, tail, color, grow=0):
    """Pinezka '9': koło + zwężający się ogon w dół po prawej stronie."""
    Rg = R + grow
    dr.ellipse([cx - Rg, cy - Rg, cx + Rg, cy + Rg], fill=color)
    x2 = cx + R + grow            # prawa krawędź ogona = prawa krawędź koła
    x1 = x2 - tw - 2 * grow
    y2 = cy + tail + grow
    dr.polygon([(x1, cy - R // 3), (x2, cy - R // 3),
                (x2, y2), (x1, y2 - int(tw * 0.9))], fill=color)


def hole(dr, cx, cy, r):
    dr.ellipse([cx - r, cy - r, cx + r, cy + r], fill=RED)


def main(out="logo99rent.png", out_small="logo99rent_small.png"):
    W = S * SS
    img = Image.new("RGB", (W, W), RED)
    d = ImageDraw.Draw(img)
    R = int(W * 0.16)             # promień koła '9'
    hr = int(R * 0.44)            # promień otworu
    tw = int(R * 0.44)            # szerokość ogona
    tail = int(R * 1.70)          # długość ogona pod środkiem koła
    cy = int(W * 0.295)
    c1 = int(W * 0.325)
    c2 = c1 + int(R * 1.82)       # lekkie nachodzenie na pierwszą '9'
    gap = int(W * 0.011)          # czerwona szczelina między dziewiątkami

    nine(d, c1, cy, R, tw, tail, "white")
    hole(d, c1, cy, hr)
    nine(d, c2, cy, R, tw, tail, RED, grow=gap)   # kontur oddzielający
    nine(d, c2, cy, R, tw, tail, "white")
    hole(d, c2, cy, hr)

    # RENT — bardzo gruby, szeroki napis
    f = font(int(W * 0.20))
    d.text((W / 2, int(W * 0.815)), "RENT", font=f, fill="white",
           anchor="mm", stroke_width=int(W * 0.005), stroke_fill="white")

    img = img.resize((S, S), Image.LANCZOS)
    img.save(out)
    img.resize((128, 128), Image.LANCZOS).save(out_small)
    print("logo zapisane:", out, out_small)


if __name__ == "__main__":
    main()
