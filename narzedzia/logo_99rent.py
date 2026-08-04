# -*- coding: utf-8 -*-
"""Generuje logo 99rent — czerwony kwadrat, dwie białe pinezki-'9' + napis RENT
(odwzorowanie przesłanego logo)."""
from PIL import Image, ImageDraw, ImageFont

RED = (228, 13, 24)
S = 1024
SS = 4  # supersampling


def font(sz):
    return ImageFont.truetype(
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", sz)


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
    tail = int(R * 1.78)          # długość ogona pod środkiem koła
    cy = int(W * 0.295)
    c1 = int(W * 0.325)
    c2 = c1 + int(R * 1.82)       # lekkie nachodzenie na pierwszą '9'
    gap = int(W * 0.011)          # czerwona szczelina między dziewiątkami

    nine(d, c1, cy, R, tw, tail, "white")
    hole(d, c1, cy, hr)
    nine(d, c2, cy, R, tw, tail, RED, grow=gap)   # kontur oddzielający
    nine(d, c2, cy, R, tw, tail, "white")
    hole(d, c2, cy, hr)

    # RENT — bardzo gruby napis, litery rozstrzelone na szerokość dziewiątek
    f = font(int(W * 0.225))
    letters = "RENT"
    widths = [d.textlength(ch, font=f) for ch in letters]
    span = W * 0.64
    gap_l = (span - sum(widths)) / (len(letters) - 1)
    x = (W - span) / 2
    y = int(W * 0.815)
    sw = int(W * 0.006)
    for ch, cw in zip(letters, widths):
        d.text((x, y), ch, font=f, fill="white", anchor="lm",
               stroke_width=sw, stroke_fill="white")
        x += cw + gap_l

    img = img.resize((S, S), Image.LANCZOS)
    img.save(out)
    img.resize((128, 128), Image.LANCZOS).save(out_small)
    print("logo zapisane:", out, out_small)


if __name__ == "__main__":
    main()
