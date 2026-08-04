# -*- coding: utf-8 -*-
"""Generuje logo 99rent — czerwony kwadrat, grube białe '99' + napis RENT
(odwzorowanie przesłanego logo; cyfry i litery z kroju Liberation Sans Bold
pogrubione obrysem)."""
from PIL import Image, ImageDraw, ImageFont

RED = (237, 28, 36)
S = 1024
SS = 6  # supersampling

FONT = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"


def spaced_text(d, text, f, cx, y, span, stroke):
    widths = [d.textlength(ch, font=f) + 2 * stroke for ch in text]
    gap = (span - sum(widths)) / (len(text) - 1) if len(text) > 1 else 0
    x = cx - span / 2 + stroke
    for ch, cw in zip(text, widths):
        d.text((x, y), ch, font=f, fill="white", anchor="lm",
               stroke_width=stroke, stroke_fill="white")
        x += cw + gap


def main(out="logo99rent.png", out_small="logo99rent_small.png"):
    W = S * SS
    img = Image.new("RGB", (W, W), RED)
    d = ImageDraw.Draw(img)

    # "99" — duże, grube, lekko zachodzące na siebie
    f99 = ImageFont.truetype(FONT, int(W * 0.62))
    sw99 = int(W * 0.005)
    spaced_text(d, "99", f99, W / 2, int(W * 0.35), int(W * 0.63), sw99)

    # "RENT" — szeroki, bardzo gruby
    fr = ImageFont.truetype(FONT, int(W * 0.20))
    swr = int(W * 0.007)
    spaced_text(d, "RENT", fr, W / 2, int(W * 0.82), int(W * 0.78), swr)

    img = img.resize((S, S), Image.LANCZOS)
    img.save(out)
    img.resize((256, 256), Image.LANCZOS).save(out_small)
    print("logo zapisane:", out, out_small)


if __name__ == "__main__":
    main()
