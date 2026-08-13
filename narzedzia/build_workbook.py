# -*- coding: utf-8 -*-
"""Build the 99rent vehicle-registration workbook (.xlsm with macros/buttons,
or a formula-identical .xlsx copy for LibreOffice recalc verification)."""
import calendar
import collections
import datetime
import re
import openpyxl
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

SOURCE = "Raport_rejestracji.xlsx"
RED = "#E30613"
DARK = "#3F3F3F"
GRAY_BG = "#F2F2F2"
YELLOW = "#FFF2CC"
BAND = "#FAFAFA"
LAST = 3000  # formula range horizon

FORM_ROW = 5   # 1-indexed sheet row of form inputs
HDR_ROW = 8    # 1-indexed sheet row of table headers
DATA_ROW = 9   # first data row
ARCH_DATA = 15  # 1-indexed first data row in Archiwum sheets (stats above)


GREEN_FILLS = {"FF00B050", "FF92D050"}   # odebrane / zarejestrowane
YELLOW_FILLS = {"FFFFFF00"}              # złożone, w rejestracji
RED_FILLS = {"FFFF0000"}                 # oznaczone na czerwono w źródle


def _row_fill(cells):
    fills = []
    for c in cells:
        f = c.fill
        rgb = None
        if f is not None and f.patternType == "solid":
            v = f.fgColor.rgb if f.fgColor is not None else None
            if isinstance(v, str):
                rgb = v
        fills.append(rgb)
    return collections.Counter(fills).most_common(1)[0][0]


def _uwagi_data_rej(u):
    """Uwagi typu '12.08.2026 DATA REJESTRACJI' / 'zarejestrowane 12.08':
    zwraca (data_rejestracji, uwagi bez tej frazy)."""
    if not u or not re.search(r"data\s+rejestracji|zarejestrowan", u, re.I):
        return None, u
    m = re.search(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", u)
    if not m:
        return None, u
    dd, mm = int(m.group(1)), int(m.group(2))
    yy = int(m.group(3)) if m.group(3) else datetime.date.today().year
    if yy < 100:
        yy += 2000
    try:
        dt = datetime.datetime(yy, mm, dd)
    except ValueError:
        return None, u
    rest = u.replace(m.group(0), "")
    rest = re.sub(r"data\s+rejestracji|zarejestrowan\w*", "", rest, flags=re.I)
    rest = re.sub(r"\s+", " ", rest).strip(" -–,.:;") or None
    return dt, rest


def _uwagi_odbior(u):
    """Uwagi typu 'ODBIOR 13.08' / 'DAWID ODBIERA 13.08.2026':
    zwraca (planowany odbiór, osoba); uwagi zostają bez zmian."""
    if not u or not re.search(r"odbi[oó]r|odbierz|odbiera", u, re.I):
        return None, None
    dt = None
    m = re.search(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", u)
    if m:
        dd, mm = int(m.group(1)), int(m.group(2))
        yy = int(m.group(3)) if m.group(3) else datetime.date.today().year
        if yy < 100:
            yy += 2000
        try:
            dt = datetime.datetime(yy, mm, dd)
        except ValueError:
            dt = None
    osoba = None
    mo = re.search(r"([A-ZĄĆĘŁŃÓŚŹŻ][\w]*)\s+ODBIERA", u, re.I)
    if mo:
        osoba = mo.group(1).capitalize()
    return dt, osoba


def load_data():
    wb = openpyxl.load_workbook(SOURCE, data_only=True)
    def s(v):
        return str(v).strip() if v is not None and str(v).strip() else None

    # Arkusz2: szczegóły rejestracji (nr rej, data rej, cena) wg VIN
    detail = {}
    for r in wb["Arkusz2"].iter_rows(min_row=2, values_only=True):
        # Marka, Model, data rej, nr rej, VIN, cena
        if s(r[4]):
            detail[s(r[4])] = (s(r[3]), r[2], r[5])

    # Arkusz1: kolor wiersza decyduje o arkuszu docelowym
    wrej = []    # (marka, model, vin, dealer, wsp, urzad, dzl, uwagi, is_red)
    zarej = []   # (marka, model, vin, nrrej, dealer, urzad, dzl, datarej, cena, uwagi)
    for cells in wb["Arkusz1"].iter_rows(min_row=2):
        r = [c.value for c in cells[:8]]
        row = (s(r[0]), s(r[1]), s(r[2]), s(r[3]), s(r[4]), s(r[5]), r[6], s(r[7]))
        if not any(row[:7]):
            continue
        fill = _row_fill(cells[:7])
        if fill in GREEN_FILLS:
            nrrej, datarej, cena = detail.get(row[2], (None, None, None))
            zarej.append((row[0], row[1], row[2], nrrej, row[3], row[5],
                          row[6], datarej, cena, row[7]))
        else:
            wrej.append(row + (fill in RED_FILLS,))

    # Arkusz2 VIN-y spoza zielonych wierszy (nie powinno ich być, ale nie gubimy)
    seen = {z[2] for z in zarej}
    for vin, (nrrej, datarej, cena) in detail.items():
        if vin not in seen:
            zarej.append((None, None, vin, nrrej, None, None, None,
                          datarej, cena, None))

    # deduplikacja wariantów pisowni (Euro-Kas / euro kas, SRÓDMIEŚCIE /
    # ŚRÓDMIEŚCIE itp.) — wszystkie wiersze dostają ujednoliconą nazwę
    m_marka = canonical_map([r[0] for r in wrej] + [z[0] for z in zarej])
    m_dealer = canonical_map([r[3] for r in wrej] + [z[4] for z in zarej])
    for var in list(m_dealer):
        alias = DEALER_ALIASES.get(_norm_key(m_dealer[var]))
        if alias:
            m_dealer[var] = alias
    m_wsp = canonical_map([r[4] for r in wrej])
    m_urzad = {u: u for u in URZEDY_CANON}
    m_urzad.update({v: canonical_map(
        URZEDY_CANON + [r[5] for r in wrej] + [z[5] for z in zarej]).get(v, v)
        for v in set([r[5] for r in wrej] + [z[5] for z in zarej]) if v})
    # urzędy mapuj na kanoniczną listę (typo SRÓDMIEŚCIE -> ŚRÓDMIEŚCIE)
    urz_canon_by_key = {_norm_key(u): u for u in URZEDY_CANON}
    def fix_urzad(v):
        if not v:
            return v
        return urz_canon_by_key.get(_norm_key(v), v)

    # wrej: (marka, model, vin, dealer, wsp, urzad, dzl,
    #        odbior, osoba, uwagi, is_red)
    wrej = [(m_marka.get(r[0], r[0]), r[1], r[2], m_dealer.get(r[3], r[3]),
             m_wsp.get(r[4], r[4]), fix_urzad(r[5]), r[6])
            + _uwagi_odbior(r[7]) + (r[7], r[8])
            for r in wrej]
    zarej = [(m_marka.get(z[0], z[0]), z[1], z[2], z[3],
              m_dealer.get(z[4], z[4]), fix_urzad(z[5]), z[6], z[7], z[8], z[9])
             for z in zarej]

    # data rejestracji zapisana w Uwagach ("12.08.2026 DATA REJESTRACJI",
    # "zarejestrowane 12.08") -> kolumna Data rejestracji, Uwagi czyszczone
    fixed = []
    for z in zarej:
        if z[7] is None:
            dt, rest = _uwagi_data_rej(z[9])
            if dt is not None:
                z = z[:7] + (dt, z[8], rest)
        fixed.append(z)
    zarej = fixed

    # DO REJESTRACJI startuje pusta — wzór wypełnia użytkownik
    # krotka: (marka, model, vin, dealer, wsp, urzad, dzl, komplet, brakuje, uwagi)
    dorej = []
    return wrej, zarej, dorej


import unicodedata


def _norm_key(v):
    """Klucz porównywania nazw: bez wielkości liter, ogonków i interpunkcji,
    żeby 'Euro-Kas', 'euro kas' i 'EURO KAS' trafiały do jednej grupy."""
    v = unicodedata.normalize("NFKD", v)
    v = "".join(ch for ch in v if not unicodedata.combining(ch))
    return "".join(ch for ch in v.casefold() if ch.isalnum())


def canonical_map(values):
    """Mapa: oryginalna pisownia -> najczęstsza pisownia w grupie."""
    groups = collections.defaultdict(collections.Counter)
    for v in values:
        if v:
            groups[_norm_key(v)][v] += 1
    out = {}
    for cnt in groups.values():
        best = cnt.most_common(1)[0][0]
        for variant in cnt:
            out[variant] = best
    return out


def canonical(values):
    """Ujednolicone, posortowane nazwy (po deduplikacji wariantów)."""
    m = canonical_map(values)
    return sorted(set(m.values()), key=lambda x: x.casefold())


URZEDY_CANON = ["BEMOWO", "BIAŁOŁĘKA", "OCHOTA", "ŚRÓDMIEŚCIE", "WAWER", "WILANÓW"]

# ręczne scalenia wariantów, których nie łapie normalizacja (literówki itp.)
DEALER_ALIASES = {
    "inchape": "Inchcape",
    "nord": "NORD AUTO",
    "mbmotors": "MB Motors Poznań",
    "saga": "Inter Saga",
    "vw": "VW GROUP",
}


# plik główny użytkownika — źródło aktualnego stanu przy przebudowie
STATE_CANDIDATES = ["Raport_rejestracji_99rent-GLOWNYYY].xlsm",
                    "Raport_rejestracji_99rent-GLOWNYYY.xlsm",
                    "user_file.xlsm"]
import os as _os
STATE = next((p for p in STATE_CANDIDATES if _os.path.exists(p)),
             STATE_CANDIDATES[-1])


def load_state(path):
    """Wczytuje aktualny stan z pliku roboczego użytkownika:
    W rejestracji (z kolorami), katalog, zakładki Archiwum RRRR-MM."""
    wb = openpyxl.load_workbook(path, data_only=True)
    def s(v):
        return str(v).strip() if v is not None and str(v).strip() else None
    synth_state = set()
    wrej = []
    ws = wb["W rejestracji"]
    # nowy układ ma kolumny Planowany odbiór (I) i Osoba prowadząca (J)
    nowy_w = str(ws.cell(row=HDR_ROW, column=9).value or "").startswith("Planowany")
    for r in range(DATA_ROW, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 12)]
        if not any(s(v) for v in vals[:7]):
            continue
        f = ws.cell(row=r, column=1).fill
        rgb = str(f.fgColor.rgb) if f.patternType == "solid" else ""
        if nowy_w:
            odb, osoba, uw = vals[8], s(vals[9]), s(vals[10])
        else:
            uw = s(vals[8])
            odb, osoba = _uwagi_odbior(uw)
        wrej.append((s(vals[0]), s(vals[1]), s(vals[2]), s(vals[3]),
                     s(vals[4]), s(vals[5]), vals[6], odb, osoba, uw,
                     rgb == "FFFFC7CE"))
    def rows_of(ws, first):
        out = []
        for r in range(first, ws.max_row + 1):
            v = [ws.cell(row=r, column=c).value for c in range(1, 11)]
            if not (s(v[0]) or s(v[2])):   # wiersze bez VIN też zachowujemy
                continue
            t = (s(v[0]), s(v[1]), s(v[2]), s(v[3]), s(v[4]),
                 s(v[5]), v[6], v[7], None, s(v[9]))
            if v[7] is not None and v[8] is None:
                # data rejestracji przybliżona (czas w pliku pusty) — nie
                # odtwarzamy czasu rejestracji przy przebudowie
                synth_state.add(id(t))
            out.append(t)
        return out
    zarej = rows_of(wb["Zarejestrowane"], DATA_ROW)
    archiwa = {}
    for name in wb.sheetnames:
        if name.startswith("Archiwum "):
            # nowy układ ma statystyki na górze i nagłówki w wierszu 14
            first = ARCH_DATA if str(
                wb[name].cell(row=ARCH_DATA - 1, column=1).value or ""
            ) == "Marka" else 2
            archiwa[name.split()[1]] = rows_of(wb[name], first)
    # DO REJESTRACJI: (marka, model, vin, dealer, wsp, urzad, dzl, komplet,
    # brakuje, uwagi) — czytamy nowy układ, starsze warianty konwertujemy
    dorej = []
    if "Do rejestracji" in wb.sheetnames:
        d = wb["Do rejestracji"]
        if str(d.cell(row=8, column=8).value or "").startswith("Komplet"):
            # nowy układ: nagłówki w wierszu 8, dane od 9, 10 kolumn
            for r in d.iter_rows(min_row=9, max_col=10, values_only=True):
                vals = tuple(v if hasattr(v, "year") else s(v) for v in r[:10])
                if any(vals[:7]):
                    dorej.append(vals)
        else:
            # stary układ: nagłówki w wierszu 4, dane od 5, 6 kolumn
            for r in d.iter_rows(min_row=5, max_col=6, values_only=True):
                if any(s(v) for v in r[:6]):
                    m, mo, vin, urz, kom, brak = (s(v) for v in r[:6])
                    dorej.append((m, mo, vin, None, None, urz, None,
                                  kom, brak, None))
    elif "Pilne" in wb.sheetnames:
        for r in wb["Pilne"].iter_rows(min_row=4, max_col=2, values_only=True):
            nm, vin = s(r[0]), s(r[1])
            if not (nm or vin) or (vin or "").upper() == "VIN":
                continue
            marka, _, model = (nm or "").partition(" ")
            dorej.append((marka or None, model or None, vin) + (None,) * 7)
    if "Import hurtowy" in wb.sheetnames:
        # niedokończone wiersze importu nie giną — trafiają do DO REJESTRACJI
        for r in wb["Import hurtowy"].iter_rows(min_row=9, max_col=8,
                                                values_only=True):
            if any(s(v) for v in r[:6]) or r[6] is not None:
                m, mo, vin, de, wsp, urz = (s(v) for v in r[:6])
                dorej.append((m, mo, vin, de, wsp, urz, r[6],
                              None, None, s(r[7])))
    return wrej, zarej, archiwa, dorej, synth_state


def build(path, with_vba, vba_bin=None, logo="logo99rent.png",
          logo_small="logo99rent_small.png"):
    import os
    synth = set()  # wiersze z przybliżoną datą rejestracji (koniec miesiąca)
    if os.path.exists(STATE):
        wrej, zarej, stare_by_month, dorej, synth = load_state(STATE)
    else:
        wrej, zarej, dorej = load_data()
        prog = (datetime.date.today().year, datetime.date.today().month)
        stare_by_month = {}
        kat = []
        for z in zarej:
            d = z[7]
            if d is None and z[6] is not None and hasattr(z[6], "year") \
                    and (z[6].year, z[6].month) < prog:
                # zarejestrowany bez daty rejestracji, złożony w starym
                # miesiącu -> archiwum tego miesiąca z datą przybliżoną
                # (koniec miesiąca); czas rejestracji zostaje pusty
                last = calendar.monthrange(z[6].year, z[6].month)[1]
                d = datetime.datetime(z[6].year, z[6].month, last)
                z = z[:7] + (d,) + z[8:]
                synth.add(id(z))
            if d is not None and (d.year, d.month) < prog:
                stare_by_month.setdefault("%04d-%02d" % (d.year, d.month),
                                          []).append(z)
            else:
                kat.append(z)
        zarej = kat

    # arkusze, po których liczą się "zarejestrowane": katalog + archiwa
    REJ_SHEETS = [("Zarejestrowane", DATA_ROW, LAST)] + \
        [("'Archiwum %s'" % k, ARCH_DATA, 5000) for k in sorted(stare_by_month)]

    def zsum(tmpl):
        """Suma formuły po katalogu i wszystkich zakładkach Archiwum.
        tmpl używa %(s)s (arkusz), %(a)d (pierwszy wiersz), %(b)d (ostatni)."""
        return "+".join(tmpl % {"s": s0, "a": a0, "b": b0}
                        for s0, a0, b0 in REJ_SHEETS)
    # listy obejmują też pojazdy z zakładek Archiwum
    zarch = zarej + [z for rows in stare_by_month.values() for z in rows]
    marki = canonical([r[0] for r in wrej] + [z[0] for z in zarch])
    dealerzy = canonical([r[3] for r in wrej] + [z[4] for z in zarch])
    wspolwl = canonical([r[4] for r in wrej])

    def by_count(items, *value_lists):
        """Kolejność malejąco wg łącznej liczby wystąpień w danych."""
        cnt = collections.Counter()
        for vals in value_lists:
            for v in vals:
                if v:
                    cnt[v] += 1
        return sorted(items, key=lambda x: (-cnt.get(x, 0), x.casefold()))

    wb = xlsxwriter.Workbook(path, {"remove_timezone": True})
    if with_vba:
        wb.add_vba_project(vba_bin)
        wb.set_vba_name("ThisWorkbook")

    A = {"font_name": "Arial", "font_size": 10}

    def fmt(**kw):
        d = dict(A)
        d.update(kw)
        return wb.add_format(d)

    f_band = fmt(bg_color=RED, font_color="white", bold=True, font_size=16,
                 align="left", valign="vcenter", indent=7)
    f_band_sub = fmt(bg_color=RED, font_color="white", font_size=10,
                     align="right", valign="vcenter")
    f_hdr = fmt(bold=True, font_color="white", bg_color=DARK, border=1,
                bottom=2, bottom_color="#E30613",
                align="center", valign="vcenter", text_wrap=True)
    f_form_label = fmt(font_size=8, font_color="#7F7F7F", bold=True,
                       align="center")
    f_form_title = fmt(bold=True, font_color=RED, font_size=10,
                       align="left", valign="vcenter")
    f_input = fmt(bg_color=YELLOW, border=1, border_color="#BFBFBF",
                  valign="vcenter")
    f_input_date = fmt(bg_color=YELLOW, border=1, border_color="#BFBFBF",
                       num_format="yyyy-mm-dd", valign="vcenter")
    f_auto = fmt(bg_color="#E7E6E6", border=1, border_color="#BFBFBF",
                 italic=True, font_color="#7F7F7F", align="center",
                 valign="vcenter")
    f_input_free = fmt(bg_color=YELLOW, border=1, border_color="#BFBFBF",
                       valign="vcenter", locked=False)
    f_text = fmt()
    f_date = fmt(num_format="yyyy-mm-dd")
    f_int = fmt(num_format="0", align="center")
    f_money = fmt(num_format='#,##0.00 "zł"')
    f_kpi_num = fmt(bold=True, font_color=RED, font_size=26,
                    align="center", valign="vcenter", bg_color="white",
                    top=1, left=1, right=1, border_color="#D9D9D9")
    f_kpi_lbl = fmt(font_size=9, font_color="#595959", align="center",
                    valign="top", text_wrap=True, bold=True, bg_color="white",
                    bottom=1, left=1, right=1, border_color="#D9D9D9")
    f_sec = fmt(bold=True, font_size=11, font_color=RED)
    f_sec_band = fmt(bold=True, font_size=11, font_color="white",
                     bg_color=RED, valign="vcenter", indent=1)
    f_tbl_text = fmt(bg_color="white", border=1, border_color="#D9D9D9")
    f_tbl_int = fmt(bg_color="white", border=1, border_color="#D9D9D9",
                    num_format="0", align="center")
    f_val_box = fmt(bg_color="white", border=1, border_color="#D9D9D9",
                    bold=True, font_color=RED, align="center")
    f_total_box = fmt(bold=True, border=1, border_color="#BFBFBF",
                      bg_color="#F2F2F2", num_format="0", align="center")
    f_total_lbl = fmt(bold=True, border=1, border_color="#BFBFBF",
                      bg_color="#F2F2F2")
    f_lbl = fmt(bold=True)
    f_val = fmt(align="center", bold=True)
    f_stat_lbl = fmt()
    f_note = fmt(font_size=9, font_color="#595959", italic=True,
                 text_wrap=True, valign="top")
    f_instr = fmt(font_size=10, text_wrap=True, valign="top",
                  bg_color=GRAY_BG, border=1, border_color="#D9D9D9")
    f_total = fmt(bold=True, top=1)

    f_red = fmt(bg_color="#FFC7CE", font_color="#9C0006")
    f_amber = fmt(bg_color="#FFEB9C", font_color="#9C6500")
    f_ok = fmt(bg_color="#C6EFCE", font_color="#006100")
    B = {"border": 1, "border_color": "#9E9E9E"}
    f_text_red = fmt(bg_color="#FFC7CE", **B)
    f_date_red = fmt(bg_color="#FFC7CE", num_format="yyyy-mm-dd", **B)
    f_int_red = fmt(bg_color="#FFC7CE", num_format="0", align="center", **B)
    # żółte wiersze = pojazdy w rejestracji (konwencja z pliku źródłowego)
    f_text_y = fmt(bg_color="#FFF9C4", **B)
    f_date_y = fmt(bg_color="#FFF9C4", num_format="yyyy-mm-dd", **B)
    f_int_y = fmt(bg_color="#FFF9C4", num_format="0", align="center", **B)
    # zielone wiersze = pojazdy zarejestrowane (konwencja z pliku źródłowego)
    f_text_g = fmt(bg_color="#C6EFCE", **B)
    f_date_g = fmt(bg_color="#C6EFCE", num_format="yyyy-mm-dd", **B)
    f_int_g = fmt(bg_color="#C6EFCE", num_format="0", align="center", **B)
    f_import = fmt(bg_color="#FFFDE7", border=1, border_color="#E0D9A0")
    f_import_date = fmt(bg_color="#FFFDE7", border=1, border_color="#E0D9A0",
                        num_format="yyyy-mm-dd")
    f_bandrow = fmt(bg_color=BAND)
    f_cnt_lbl = fmt(bold=True, font_size=9, font_color="#595959",
                    align="center", valign="vcenter", bg_color="white",
                    top=1, left=1, bottom=1, border_color="#D9D9D9")
    f_cnt_num = fmt(bold=True, font_color=RED, font_size=16, align="center",
                    valign="vcenter", bg_color="white",
                    top=1, right=1, bottom=1, border_color="#D9D9D9")

    def header_band(ws, title, ncols):
        ws.set_row(0, 40)
        ws.merge_range(0, 0, 0, ncols - 2, title, f_band)
        ws.write(0, ncols - 1, "99rent", f_band_sub)
        ws.insert_image(0, 0, logo_small,
                        {"x_scale": 0.14, "y_scale": 0.14,
                         "x_offset": 4, "y_offset": 4,
                         "object_position": 3})

    def nav_button(ws, col):
        ws.set_row(1, 40)
        if with_vba:
            ws.insert_button(1, col, {"macro": "IdzPulpit", "caption": "◀ PULPIT",
                                      "width": 96, "height": 28,
                                      "x_offset": 2, "y_offset": 5})

    sheet_order = []

    # ================================================================ PULPIT
    ws = wb.add_worksheet("PULPIT")
    sheet_order.append(ws)
    ws.set_tab_color(RED)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:C", 14)
    ws.set_column("D:O", 12)
    header_band(ws, "  RAPORT REJESTRACJI POJAZDÓW", 15)
    f_clock = fmt(bold=True, font_size=12, font_color=DARK, align="right",
                  num_format="yyyy-mm-dd  hh:mm:ss")
    ws.write(1, 8, "stan na:", fmt(font_size=9, font_color="#7F7F7F",
                                   align="right", valign="vcenter"))
    ws.merge_range(1, 9, 1, 10, "", f_clock)
    ws.write_formula(1, 9, "=NOW()", f_clock)
    ws.insert_image("B3", logo, {"x_scale": 0.24, "y_scale": 0.24,
                                 "object_position": 3})

    cnt_all = zsum('SUMPRODUCT(--((%(s)s!$A$%(a)d:$A$%(b)d&%(s)s!$C$%(a)d:$C$%(b)d)<>""))')
    cnt_i = zsum('COUNT(%(s)s!$I$%(a)d:$I$%(b)d)')
    sum_i = zsum('SUM(%(s)s!$I$%(a)d:$I$%(b)d)')
    cnt_mies = zsum('COUNTIFS(%(s)s!$H$%(a)d:$H$%(b)d,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))')
    kpis = [
        ("=SUMPRODUCT(--(('W rejestracji'!$A$%d:$A$%d&'W rejestracji'!$C$%d:$C$%d)<>\"\"))"
         % (DATA_ROW, LAST, DATA_ROW, LAST), "POJAZDY\nW REJESTRACJI"),
        ("=" + cnt_all, "POJAZDY\nZAREJESTROWANE"),
        ("=IF((%s)=0,\"—\",ROUND((%s)/(%s),1))" % (cnt_i, sum_i, cnt_i),
         "ŚREDNI CZAS\nREJESTRACJI (DNI)"),
        ("=" + cnt_mies, "ZAREJESTROWANE\nW TYM MIESIĄCU"),
        ("=COUNTIFS('W rejestracji'!$I$%d:$I$%d,\">=\"&TODAY(),"
         "'W rejestracji'!$I$%d:$I$%d,\"<=\"&(TODAY()+1))"
         % (DATA_ROW, LAST, DATA_ROW, LAST),
         "ODBIORY\nDZIŚ / JUTRO"),
        ("=COUNTIF('W rejestracji'!$G$%d:$G$%d,\"<=\"&(TODAY()-25))"
         % (DATA_ROW, LAST),
         "TERMIN 30 DNI\nZOSTAŁO ≤ 5 DNI"),
    ]
    col = 3
    for formula, label in kpis:
        ws.merge_range(3, col, 4, col + 1, "", f_kpi_num)
        ws.write_formula(3, col, formula, f_kpi_num)
        ws.merge_range(5, col, 5, col + 1, label, f_kpi_lbl)
        ws.set_row(5, 26)
        col += 2

    if with_vba:
        navs = [("W REJESTRACJI", "IdzWRejestracji"),
                ("DO REJESTRACJI", "IdzImport"),
                ("KATALOG ZAREJESTR.", "IdzZarejestrowane"),
                ("PODSUMOWANIE", "IdzPodsumowanie")]
        col = 1
        for caption, macro in navs:
            ws.insert_button(7, col, {"macro": macro, "caption": caption,
                                      "width": 140, "height": 34})
            col += 2
    ws.set_row(7, 30)

    instr = (
        "JAK KORZYSTAĆ Z PLIKU\n"
        "1.  Przy otwarciu kliknij „Włącz zawartość” — przyciski wymagają włączonych makr. Jeśli Excel blokuje makra: zamknij plik, "
        "kliknij go prawym przyciskiem → Właściwości → zaznacz „Odblokuj” → OK i otwórz ponownie.\n"
        "2.  W REJESTRACJI — pojazdy oczekujące (żółte wiersze). Nowy wniosek wpisujesz w formularzu (wiersz 5) i klikasz DODAJ WNIOSEK. "
        "Kolumna „Dni od złożenia” liczy się sama i podświetla pojazdy czekające zbyt długo (pomarańczowy > 10 dni, czerwony > 21 dni); "
        "w kolumnach „Planowany odbiór” i „Osoba prowadząca” pilnujesz odbiorów — karty ODBIORY DZIŚ/JUTRO i TERMIN 30 DNI na PULPICIE liczą się z nich same. "
        "Wyszukiwarka VIN na dole PULPITU pokazuje, w którym arkuszu jest pojazd.\n"
        "3.  DO REJESTRACJI — wzór pojazdów przed złożeniem: wklejasz wiele naraz, oznaczasz komplet dokumentów (TAK/NIE, przy NIE wpisujesz czego brakuje) "
        "i przyciskiem IMPORTUJ DO REJESTRU dodajesz je do rejestru (zaznaczenie wierszy = import tylko wybranych; przenoszą się tylko wiersze z kompletem).\n"
        "4.  Po odebraniu rejestracji: zaznacz pojazdy w W REJESTRACJI i kliknij ZAREJESTRUJ ZAZNACZONE — przechodzą do katalogu "
        "ZAREJESTROWANE z licznikiem dni. Numer rejestracyjny wpisujesz wprost w kolumnie „Nr rejestracyjny” katalogu.\n"
        "5.  ZAREJESTROWANE — pojazd zarejestrowany wcześniej (poza rejestrem) dodasz bezpośrednio przyciskiem DODAJ DO KATALOGU; "
        "pomyłkę cofniesz przyciskiem PRZYWRÓĆ DO REJESTRACJI (zaznacz wiersze).\n"
        "6.  PODSUMOWANIE — statystyki wg urzędu, dealera i marki oraz czasy rejestracji liczą się automatycznie.\n"
        "7.  ARCHIWUM — na początku każdego miesiąca plik proponuje przeniesienie pojazdów zarejestrowanych w starych miesiącach "
        "do zakładki „Archiwum RRRR-MM” na dole pliku (przycisk ARCHIWIZUJ STARE MIES. w katalogu robi to na żądanie).\n"
        "Żółte pola = pola do wypełnienia.  Duplikaty VIN są blokowane przez przyciski i podświetlane na czerwono w tabelach."
    )
    ws.merge_range(9, 1, 17, 10, instr, f_instr)

    zasady = (
        "ZASADY DLA NOWYCH UŻYTKOWNIKÓW — JAK NIE POPSUĆ PLIKU\n"
        "•  Pojazdy przenoś WYŁĄCZNIE przyciskami (IMPORTUJ DO REJESTRU, ZAREJESTRUJ ZAZNACZONE, PRZYWRÓĆ, ARCHIWIZUJ) — "
        "nie wycinaj i nie przeklejaj wierszy ręcznie między zakładkami.\n"
        "•  Nie zmieniaj nazw zakładek ani nagłówków tabel — liczniki i przyciski szukają ich po nazwach.\n"
        "•  Nie wpisuj nic w kolumny liczone automatycznie („Dni od złożenia”, „Czas rejestracji”) — PULPIT i PODSUMOWANIE "
        "są chronione przed edycją; wypełniasz tylko żółte pola, tabele danych i pola filtrów.\n"
        "•  Dane z zewnątrz wklejaj jako wartości (prawy przycisk → Wklej specjalnie → Wartości), żeby nie nadpisać kolorów i list rozwijanych.\n"
        "•  Nie usuwaj zakładek Archiwum ani ukrytej zakładki „Listy” — zasilają statystyki i listy rozwijane.\n"
        "•  Zapisuj zawsze jako .xlsm (skoroszyt z obsługą makr) i pracujcie w jednej kopii pliku naraz.\n"
        "•  Daty wpisuj jako RRRR-MM-DD albo wybieraj z listy; VIN ma 17 znaków — duplikaty podświetlają się na czerwono.\n"
        "•  Coś poszło nie tak? Zamknij plik BEZ zapisywania i otwórz ponownie — wróci ostatni zapisany stan."
    )
    f_rules = fmt(font_size=10, text_wrap=True, valign="top",
                  bg_color="#FDE9E9", border=1, border_color=RED)
    ws.merge_range(19, 1, 27, 10, zasady, f_rules)
    f_leg_y = fmt(bg_color="#FFF9C4", border=1, border_color="#9E9E9E",
                  align="center", font_size=9)
    f_leg_g = fmt(bg_color="#C6EFCE", border=1, border_color="#9E9E9E",
                  align="center", font_size=9)
    f_leg_r = fmt(bg_color="#FFC7CE", border=1, border_color="#9E9E9E",
                  align="center", font_size=9)
    ws.write(29, 1, "LEGENDA KOLORÓW:", f_lbl)
    ws.merge_range(29, 3, 29, 4, "w rejestracji (złożone)", f_leg_y)
    ws.merge_range(29, 5, 29, 6, "zarejestrowany", f_leg_g)
    ws.merge_range(29, 7, 29, 8, "wymaga uwagi / zaległy", f_leg_r)

    # --- SZUKAJ VIN --------------------------------------------------------
    ws.write(31, 1, "SZUKAJ VIN:", f_lbl)
    ws.merge_range(31, 3, 31, 4, "", f_input_free)
    ws.write(32, 3, "wpisz pełny VIN i Enter", f_note)
    SV = "$D$32"
    szuk = [("'W rejestracji'", DATA_ROW, LAST, "W REJESTRACJI"),
            ("Zarejestrowane", DATA_ROW, LAST, "ZAREJESTROWANY (katalog)"),
            ("'Do rejestracji'", DATA_ROW, DATA_ROW + 299, "DO REJESTRACJI")] + \
           [("'Archiwum %s'" % k, ARCH_DATA, 5000,
             "ZAREJESTROWANY (Archiwum %s)" % k) for k in sorted(stare_by_month)]
    st_f = '"NIE ZNALEZIONO"'
    poj_f = '""'
    for sh, a, b, lab in reversed(szuk):
        cnt = "COUNTIF(%s!$C$%d:$C$%d,%s)>0" % (sh, a, b, SV)
        st_f = 'IF(%s,"%s",%s)' % (cnt, lab, st_f)
        idx = ('T(INDEX(%s!$A$%d:$A$%d,MATCH(%s,%s!$C$%d:$C$%d,0)))&" "&'
               'T(INDEX(%s!$B$%d:$B$%d,MATCH(%s,%s!$C$%d:$C$%d,0)))'
               % (sh, a, b, SV, sh, a, b, sh, a, b, SV, sh, a, b))
        poj_f = 'IF(%s,%s,%s)' % (cnt, idx, poj_f)
    ws.merge_range(31, 5, 31, 7, "", f_val_box)
    ws.write_formula(31, 5, '=IF(%s="","",%s)' % (SV, st_f), f_val_box)
    ws.merge_range(31, 8, 31, 10, "", f_tbl_text)
    ws.write_formula(31, 8, '=IF(%s="","",TRIM(%s))' % (SV, poj_f), f_tbl_text)

    ws.write(34, 1, "Stworzone przez: Oskar Figaszewski", f_note)

    # ========================================================= W REJESTRACJI
    ws = wb.add_worksheet("W rejestracji")
    sheet_order.append(ws)
    ws.set_tab_color("#F4A100")
    headers = ["Marka", "Model", "VIN", "Dealer", "Współwłaściciel", "Urząd",
               "Data złożenia", "Dni od złożenia", "Planowany odbiór",
               "Osoba prowadząca", "Uwagi"]
    widths = [14, 20, 23, 17, 17, 15, 14, 14, 16, 17, 32]
    colfmts = [f_text, f_text, f_text, f_text, f_text, f_text, f_date, f_int,
               f_date, f_text, f_text]
    for c, (w, cf) in enumerate(zip(widths, colfmts)):
        ws.set_column(c, c, w, cf)
    ws.set_column(11, 11, 2)
    ws.set_column(12, 12, 22)
    header_band(ws, "  POJAZDY W TRAKCIE REJESTRACJI", 11)
    nav_button(ws, 12)
    ws.merge_range(1, 0, 1, 1, "POJAZDY W REJESTRACJI", f_cnt_lbl)
    ws.write_formula(
        1, 2, "=SUMPRODUCT(--(($A$%d:$A$%d&$C$%d:$C$%d)<>\"\"))"
        % (DATA_ROW, LAST, DATA_ROW, LAST), f_cnt_num)

    ws.merge_range(2, 0, 2, 10,
                   "FORMULARZ — NOWY WNIOSEK:  wypełnij żółte pola i dodaj przez DO REJESTRACJI lub wpisz bezpośrednio w tabeli",
                   f_form_title)
    for c, h in enumerate(headers):
        ws.write(3, c, h, f_form_label)
    for c in range(11):
        if c in (6, 8):
            ws.write_blank(FORM_ROW - 1, c, None, f_input_date)
        elif c == 7:
            ws.write(FORM_ROW - 1, c, "auto", f_auto)
        else:
            ws.write_blank(FORM_ROW - 1, c, None, f_input)
    ws.set_row(FORM_ROW - 1, 22)
    if with_vba:
        ws.insert_button(6, 12, {"macro": "ZarejestrujZaznaczone",
                                 "caption": "ZAREJESTRUJ ZAZNACZONE ▶",
                                 "width": 170, "height": 34})
        ws.write(9, 12, "Zaznacz wiersze pojazdów i kliknij, aby przenieść "
                 "je do katalogu ZAREJESTROWANE (nr rej. wpiszesz tam w kolumnie D).",
                 f_note)

    for c, h in enumerate(headers):
        ws.write(HDR_ROW - 1, c, h, f_hdr)
    ws.set_row(HDR_ROW - 1, 28)

    r = DATA_ROW - 1  # 0-indexed
    for row in wrej:
        (marka, model, vin, dealer, wsp, urzad, dzl,
         odbior, osoba, uwagi, is_red) = row
        ftxt = f_text_red if is_red else f_text_y
        fdat = f_date_red if is_red else f_date_y
        fint = f_int_red if is_red else f_int_y
        for c, v in ((0, marka), (1, model), (2, vin), (3, dealer),
                     (4, wsp), (5, urzad), (9, osoba), (10, uwagi)):
            if v is not None:
                ws.write_string(r, c, v, ftxt)
            else:
                ws.write_blank(r, c, None, ftxt)
        for c, v in ((6, dzl), (8, odbior)):
            if v is not None:
                ws.write_datetime(r, c, v, fdat)
            else:
                ws.write_blank(r, c, None, fdat)
        ws.write_formula(
            r, 7, '=IF($G%d="","",TODAY()-$G%d)' % (r + 1, r + 1), fint)
        r += 1
    last_data = r  # 0-indexed row after last

    ws.autofilter(HDR_ROW - 1, 0, LAST - 1, 10)
    ws.freeze_panes(HDR_ROW, 0)
    ws.conditional_format(DATA_ROW - 1, 7, LAST - 1, 7,
                          {"type": "cell", "criteria": ">", "value": 21,
                           "format": f_red})
    ws.conditional_format(DATA_ROW - 1, 7, LAST - 1, 7,
                          {"type": "cell", "criteria": "between",
                           "minimum": 11, "maximum": 21, "format": f_amber})
    ws.conditional_format(DATA_ROW - 1, 2, LAST - 1, 2,
                          {"type": "duplicate", "format": f_red})
    # planowany odbiór dziś/jutro — podświetl na zielono
    ws.conditional_format(DATA_ROW - 1, 8, LAST - 1, 8,
                          {"type": "formula",
                           "criteria": '=AND($I%d<>"",$I%d<=TODAY()+1)'
                           % (DATA_ROW, DATA_ROW),
                           "format": f_ok})
    nl = len(marki)
    ws.data_validation(FORM_ROW - 1, 0, FORM_ROW - 1, 0,
                       {"validate": "list", "source": "=Listy!$A$2:$A$%d" % (nl + 1),
                        "show_error": False})
    ws.data_validation(FORM_ROW - 1, 3, FORM_ROW - 1, 3,
                       {"validate": "list",
                        "source": "=Listy!$B$2:$B$%d" % (len(dealerzy) + 1),
                        "show_error": False})
    ws.data_validation(FORM_ROW - 1, 4, FORM_ROW - 1, 4,
                       {"validate": "list",
                        "source": "=Listy!$D$2:$D$%d" % (len(wspolwl) + 1),
                        "show_error": False})
    ws.data_validation(FORM_ROW - 1, 5, FORM_ROW - 1, 5,
                       {"validate": "list",
                        "source": "=Listy!$C$2:$C$%d" % (len(URZEDY_CANON) + 1),
                        "show_error": False})
    ws.data_validation(DATA_ROW - 1, 5, LAST - 1, 5,
                       {"validate": "list",
                        "source": "=Listy!$C$2:$C$%d" % (len(URZEDY_CANON) + 1),
                        "show_error": False})
    ws.data_validation(FORM_ROW - 1, 6, FORM_ROW - 1, 6,
                       {"validate": "list", "source": "=Listy!$E$2:$E$32",
                        "show_error": False, "show_input": True,
                        "input_title": "Data złożenia",
                        "input_message": "Wybierz datę z listy (ostatnie 31 dni) "
                                         "albo wpisz RRRR-MM-DD. Puste pole = dziś."})
    ws.data_validation(FORM_ROW - 1, 8, FORM_ROW - 1, 8,
                       {"validate": "list", "source": "=Listy!$J$2:$J$16",
                        "show_error": False, "show_input": True,
                        "input_title": "Planowany odbiór",
                        "input_message": "Wybierz datę z listy (14 dni w przód) "
                                         "albo wpisz RRRR-MM-DD."})
    ws.data_validation(DATA_ROW - 1, 8, LAST - 1, 8,
                       {"validate": "list", "source": "=Listy!$J$2:$J$16",
                        "show_error": False})

    # ===================== DO REJESTRACJI (wzór + import hurtowy) ==========
    ws = wb.add_worksheet("Do rejestracji")
    sheet_order.append(ws)
    ws.set_tab_color("#E65100")
    dhdr = ["Marka", "Model", "VIN", "Dealer", "Współwłaściciel", "Urząd",
            "Data złożenia", "Komplet dokumentów?", "Czego brakuje", "Uwagi"]
    dw = [14, 20, 23, 17, 17, 15, 14, 20, 30, 24]
    for c, w in enumerate(dw):
        ws.set_column(c, c, w)
    ws.set_column(10, 10, 2)
    ws.set_column(11, 11, 26)
    header_band(ws, "  DO REJESTRACJI — WZÓR I IMPORT HURTOWY", 10)
    nav_button(ws, 11)
    DR_END = DATA_ROW - 1 + 299  # 0-indeksowany ostatni wiersz tabeli
    ws.merge_range(1, 0, 1, 1, "POJAZDY DO REJESTRACJI", f_cnt_lbl)
    ws.write_formula(
        1, 2, "=SUMPRODUCT(--(($A$%d:$A$%d&$C$%d:$C$%d)<>\"\"))"
        % (DATA_ROW, DR_END + 1, DATA_ROW, DR_END + 1), f_cnt_num)

    ws.merge_range(2, 0, 2, 9,
                   "Wklej 10, 20, 30… pojazdów do tabeli (od wiersza 9) i kliknij "
                   "IMPORTUJ DO REJESTRU (możesz też zaznaczyć tylko wybrane wiersze)", f_form_title)
    ws.merge_range(3, 0, 5, 9,
                   "Wzór: w kolumnie „Komplet dokumentów?” wybierz TAK / NIE — przy "
                   "NIE wpisz, czego brakuje (wiersz podświetli się na czerwono, "
                   "komplet na zielono). Import przenosi do W REJESTRACJI tylko "
                   "wiersze z kompletem (TAK) i wypełnionymi kolumnami A–G; braki "
                   "zostają podświetlone na czerwono. Bez zaznaczenia przenosi się "
                   "wszystko, duplikaty VIN są pomijane.", f_instr)
    if with_vba:
        ws.insert_button(2, 11, {"macro": "PrzeniesWnioski",
                                 "caption": "IMPORTUJ DO REJESTRU",
                                 "width": 180, "height": 44})

    for c, h in enumerate(dhdr):
        ws.write(HDR_ROW - 1, c, h, f_hdr)
    ws.set_row(HDR_ROW - 1, 28)
    for rr in range(DATA_ROW - 1, DR_END + 1):
        for c in range(10):
            ws.write_blank(rr, c, None,
                           f_import_date if c == 6 else f_import)
    r = DATA_ROW - 1
    for row10 in dorej:
        for c, v in enumerate(row10[:10]):
            if v is None:
                continue
            if hasattr(v, "year"):
                ws.write_datetime(r, c, v, f_import_date)
            else:
                ws.write_string(r, c, str(v), f_import)
        r += 1
    ws.data_validation(DATA_ROW - 1, 3, DR_END, 3,
                       {"validate": "list",
                        "source": "=Listy!$B$2:$B$%d" % (len(dealerzy) + 1),
                        "show_error": False})
    ws.data_validation(DATA_ROW - 1, 4, DR_END, 4,
                       {"validate": "list",
                        "source": "=Listy!$D$2:$D$%d" % (len(wspolwl) + 1),
                        "show_error": False})
    ws.data_validation(DATA_ROW - 1, 5, DR_END, 5,
                       {"validate": "list",
                        "source": "=Listy!$C$2:$C$%d" % (len(URZEDY_CANON) + 1),
                        "show_error": False})
    ws.data_validation(DATA_ROW - 1, 6, DR_END, 6,
                       {"validate": "list", "source": "=Listy!$E$2:$E$32",
                        "show_error": False})
    ws.data_validation(DATA_ROW - 1, 7, DR_END, 7,
                       {"validate": "list", "source": ["TAK", "NIE"],
                        "show_error": False})
    ws.conditional_format(DATA_ROW - 1, 0, DR_END, 9,
                          {"type": "formula",
                           "criteria": '=$H%d="NIE"' % DATA_ROW,
                           "format": f_red})
    ws.conditional_format(DATA_ROW - 1, 0, DR_END, 9,
                          {"type": "formula",
                           "criteria": '=$H%d="TAK"' % DATA_ROW,
                           "format": f_ok})
    ws.freeze_panes(HDR_ROW, 0)

    # ========================================================= ZAREJESTROWANE
    ws = wb.add_worksheet("Zarejestrowane")
    sheet_order.append(ws)
    ws.set_tab_color("#2E7D32")
    zhdr = ["Marka", "Model", "VIN", "Nr rejestracyjny", "Dealer", "Urząd",
            "Data złożenia", "Data rejestracji", "Czas rejestracji (dni)",
            "Uwagi"]
    zw = [14, 20, 23, 17, 17, 15, 14, 15, 15, 32]
    zfmts = [f_text, f_text, f_text, f_text, f_text, f_text, f_date, f_date,
             f_int, f_text]
    for c, (w, cf) in enumerate(zip(zw, zfmts)):
        ws.set_column(c, c, w, cf)
    ws.set_column(10, 10, 2)
    ws.set_column(11, 11, 24)
    header_band(ws, "  KATALOG POJAZDÓW ZAREJESTROWANYCH", 10)
    nav_button(ws, 11)
    ws.merge_range(1, 0, 1, 1, "W KATALOGU (BIEŻĄCY MIES.)", f_cnt_lbl)
    ws.write_formula(
        1, 2, "=SUMPRODUCT(--(($A$%d:$A$%d&$C$%d:$C$%d)<>\"\"))"
        % (DATA_ROW, LAST, DATA_ROW, LAST), f_cnt_num)
    ws.merge_range(1, 4, 1, 5, "ZAREJESTROWANE ŁĄCZNIE Z ARCHIWUM", f_cnt_lbl)
    ws.write_formula(1, 6, "=" + cnt_all, f_cnt_num)

    ws.merge_range(2, 0, 2, 9,
                   "FORMULARZ — POJAZD JUŻ ZAREJESTROWANY:  wypełnij żółte pola i kliknij DODAJ DO KATALOGU",
                   f_form_title)
    for c, h in enumerate(zhdr):
        ws.write(3, c, h, f_form_label)
    for c in range(10):
        if c in (6, 7):
            ws.write_blank(FORM_ROW - 1, c, None, f_input_date)
        elif c == 8:
            ws.write(FORM_ROW - 1, c, "auto", f_auto)
        else:
            ws.write_blank(FORM_ROW - 1, c, None, f_input)
    ws.set_row(FORM_ROW - 1, 22)
    if with_vba:
        ws.insert_button(3, 11, {"macro": "DodajZarejestrowany",
                                 "caption": "DODAJ DO KATALOGU",
                                 "width": 165, "height": 40})
        ws.insert_button(6, 11, {"macro": "ArchiwizujStareMiesiace",
                                 "caption": "ARCHIWIZUJ STARE MIES.",
                                 "width": 165, "height": 34})
        ws.insert_button(13, 11, {"macro": "PrzywrocZaznaczone",
                                  "caption": "◀ PRZYWRÓĆ DO REJESTRACJI",
                                  "width": 165, "height": 34})
        ws.write(16, 11, "Zaznacz wiersze i kliknij, aby cofnąć pojazdy "
                 "do arkusza W REJESTRACJI (nr rej i data rejestracji "
                 "zostaną usunięte).", f_note)
        ws.write(9, 11, "Przenosi pojazdy zarejestrowane w poprzednich "
                 "miesiącach do zakładki „Archiwum RRRR-MM” na dole pliku. "
                 "Przy pierwszym otwarciu w nowym miesiącu plik sam o to "
                 "zapyta.", f_note)

    for c, h in enumerate(zhdr):
        ws.write(HDR_ROW - 1, c, h, f_hdr)
    ws.set_row(HDR_ROW - 1, 28)
    ws.write_comment(
        HDR_ROW - 1, 6,
        "Dla pojazdów dodanych ręcznie uzupełnij datę złożenia wniosku, "
        "aby licznik czasu rejestracji mógł się policzyć.",
        {"author": "99rent"})

    r = DATA_ROW - 1
    for marka, model, vin, nrrej, dealer, urzad, dzl, datarej, cena, uwagi in zarej:
        for c, v in ((0, marka), (1, model), (2, vin), (3, nrrej),
                     (4, dealer), (5, urzad), (9, uwagi)):
            if v:
                ws.write_string(r, c, v, f_text_g)
            else:
                ws.write_blank(r, c, None, f_text_g)
        if dzl is not None:
            ws.write_datetime(r, 6, dzl, f_date_g)
        else:
            ws.write_blank(r, 6, None, f_date_g)
        if datarej is not None:
            ws.write_datetime(r, 7, datarej, f_date_g)
        else:
            ws.write_blank(r, 7, None, f_date_g)
        ws.write_formula(
            r, 8, '=IF(OR($G%d="",$H%d="",$H%d<$G%d),"",$H%d-$G%d)'
            % ((r + 1,) * 6), f_int_g)
        r += 1
    zlast = r

    ws.autofilter(HDR_ROW - 1, 0, LAST - 1, 9)
    ws.freeze_panes(HDR_ROW, 0)
    ws.conditional_format(DATA_ROW - 1, 2, LAST - 1, 2,
                          {"type": "duplicate", "format": f_red})

    ws.data_validation(FORM_ROW - 1, 0, FORM_ROW - 1, 0,
                       {"validate": "list", "source": "=Listy!$A$2:$A$%d" % (nl + 1),
                        "show_error": False})
    ws.data_validation(FORM_ROW - 1, 4, FORM_ROW - 1, 4,
                       {"validate": "list",
                        "source": "=Listy!$B$2:$B$%d" % (len(dealerzy) + 1),
                        "show_error": False})
    ws.data_validation(FORM_ROW - 1, 5, FORM_ROW - 1, 5,
                       {"validate": "list",
                        "source": "=Listy!$C$2:$C$%d" % (len(URZEDY_CANON) + 1),
                        "show_error": False})
    for dc in (6, 7):   # daty: złożenia i rejestracji — szybki wybór z listy
        ws.data_validation(FORM_ROW - 1, dc, FORM_ROW - 1, dc,
                           {"validate": "list", "source": "=Listy!$E$2:$E$32",
                            "show_error": False, "show_input": True,
                            "input_title": "Data",
                            "input_message": "Wybierz z listy (31 dni wstecz) "
                                             "albo wpisz RRRR-MM-DD."})

    # ========================================================== PODSUMOWANIE
    ws = wb.add_worksheet("PODSUMOWANIE")
    sheet_order.append(ws)
    ws.set_tab_color("#616161")
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 34)
    ws.set_column("C:C", 16)
    ws.set_column("D:D", 3)
    ws.set_column("E:E", 20)
    ws.set_column("F:G", 16)
    ws.set_column("H:H", 3)
    ws.set_column("I:I", 20)
    ws.set_column("J:K", 16)
    ws.set_column("L:L", 3)
    ws.set_column("M:M", 20)
    ws.set_column("N:O", 16)
    header_band(ws, "  PODSUMOWANIE REJESTRACJI", 15)
    ws.hide_gridlines(2)
    ws.set_row(1, 44)
    if with_vba:
        ws.insert_button(1, 1, {"macro": "GenerujRaport",
                                "caption": "GENERUJ RAPORT (PLIK)",
                                "width": 180, "height": 30,
                                "x_offset": 2, "y_offset": 6})
        ws.insert_button(1, 5, {"macro": "IdzPulpit",
                                "caption": "◀ PULPIT",
                                "width": 96, "height": 30,
                                "x_offset": 8, "y_offset": 6})

    ws.merge_range(2, 1, 2, 2, "  STATYSTYKI", f_sec_band)
    ws.set_row(2, 22)
    stats = [
        ("Pojazdy w rejestracji (oczekujące)",
         "=SUMPRODUCT(--(('W rejestracji'!$A$%d:$A$%d&'W rejestracji'!$C$%d:$C$%d)<>\"\"))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
        ("Pojazdy zarejestrowane (z archiwum)", "=" + cnt_all),
        ("Średni czas rejestracji (dni)",
         "=IF((%s)=0,\"—\",ROUND((%s)/(%s),1))" % (cnt_i, sum_i, cnt_i)),
        ("Maksymalny czas rejestracji (dni)",
         "=IF((%s)=0,\"—\",MAX(%s))" % (cnt_i, ",".join(
             "%s!$I$%d:$I$%d" % (s0, a0, b0) for s0, a0, b0 in REJ_SHEETS))),
        ("Minimalny czas rejestracji (dni)",
         "=IF((%s)=0,\"—\",MIN(%s))" % (cnt_i, ",".join(
             "%s!$I$%d:$I$%d" % (s0, a0, b0) for s0, a0, b0 in REJ_SHEETS))),
        ("Zarejestrowane w bieżącym miesiącu", "=" + cnt_mies),
        ("Najdłużej oczekujący (dni od złożenia)",
         "=IF(COUNT('W rejestracji'!$H$%d:$H$%d)=0,\"—\",MAX('W rejestracji'!$H$%d:$H$%d))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
    ]
    for i, (label, formula) in enumerate(stats):
        ws.write(3 + i, 1, label, f_tbl_text)
        ws.write_formula(3 + i, 2, formula, f_val_box)
    stat_wrej_cell = "$C$4"
    stat_zar_cell = "$C$5"

    # --- FILTR: licz wg urzędu / marki / dealera / miesiąca ----------------
    # pod tabelą WG MARKI (jej ostatni wiersz to RAZEM na 0-idx 5+len(marki))
    FT = max(12, len(marki) + 7)  # 0-indeksowany wiersz pasa FILTR
    ws.merge_range(FT, 8, FT, 10, "  FILTR — policz wg wybranych kryteriów", f_sec_band)
    ws.set_row(FT, 22)
    filt = [("Miesiąc", "F", 9), ("Urząd", "G", len(URZEDY_CANON) + 1),
            ("Marka", "H", len(marki) + 1), ("Dealer", "I", len(dealerzy) + 1),
            ("Współwłaściciel", "D", len(wspolwl) + 1)]
    for i, (label, lcol, n) in enumerate(filt):
        rr = FT + 1 + i
        ws.write(rr, 8, label, f_tbl_text)
        ws.write_string(rr, 9, "(wszystkie)", f_input_free)
        ws.data_validation(rr, 9, rr, 9,
                           {"validate": "list",
                            "source": "=Listy!$%s$2:$%s$%d" % (lcol, lcol, n + 1),
                            "show_error": False})
    FM, FU, FMA, FD, FW = ("$J$%d" % (FT + 2 + i) for i in range(5))

    def sump(sheet, cmarka, cdealer, curzad, cdata, d=DATA_ROW, l=LAST,
             cwsp=None, expr=None):
        g = lambda col: "%s!$%s$%d:$%s$%d" % (sheet, col, d, col, l)
        wsp_cond = ""
        if cwsp:
            wsp_cond = ("*IF(%(fw)s=\"(wszystkie)\",1,--(%(wsp)s=%(fw)s))"
                        % {"fw": FW, "wsp": g(cwsp)})
        tmpl = ("=SUMPRODUCT((%(base)s)"
                "*IF(%(fu)s=\"(wszystkie)\",1,--(%(urz)s=%(fu)s))"
                "*IF(%(fm)s=\"(wszystkie)\",1,--(%(mar)s=%(fm)s))"
                "*IF(%(fd)s=\"(wszystkie)\",1,--(%(dea)s=%(fd)s))"
                + wsp_cond +
                "*IF(%(fmies)s=\"(wszystkie)\",1,"
                "(%(dat)s>=DATEVALUE(%(fmies)s&\"-01\"))"
                "*(%(dat)s<EDATE(DATEVALUE(%(fmies)s&\"-01\"),1))))")
        return tmpl % {"base": expr or ('((%s&%s)<>"")' % (g(cmarka), g("C"))),
                       "urz": g(curzad), "mar": g(cmarka),
                       "dea": g(cdealer), "dat": g(cdata),
                       "fu": FU, "fm": FMA, "fd": FD, "fmies": FM}

    FR = FT + 7  # 0-indeksowany pierwszy wiersz wyników filtra
    ws.merge_range(FR, 8, FR, 9, "W rejestracji (wg daty złożenia)", f_tbl_text)
    ws.write_formula(FR, 10,
                     sump("'W rejestracji'", "A", "D", "F", "G", cwsp="E"),
                     f_val_box)
    ws.merge_range(FR + 1, 8, FR + 1, 9, "Zarejestrowane (wg daty rejestracji)", f_tbl_text)
    zar_f = "+".join(
        sump(s0, "A", "E", "F", "H", a0, b0)[1:] for s0, a0, b0 in REJ_SHEETS)
    ws.write_formula(FR + 1, 10, "=" + zar_f, f_val_box)
    ws.merge_range(FR + 2, 8, FR + 2, 9, "RAZEM (rejestr + zarejestrowane)", f_tbl_text)
    ws.write_formula(FR + 2, 10, "=K%d+K%d" % (FR + 1, FR + 2), f_val_box)
    ws.merge_range(FR + 3, 8, FR + 3, 9, "Średni czas rejestracji (dla filtra)", f_tbl_text)
    # kolumna I bywa tekstem "" (brak dat) — liczymy tylko wartości liczbowe
    czas_sum = "+".join(
        sump(s0, "A", "E", "F", "H", a0, b0,
             expr='IF(ISNUMBER(%(s)s!$I$%(a)d:$I$%(b)d),%(s)s!$I$%(a)d:$I$%(b)d,0)'
             % {"s": s0, "a": a0, "b": b0})[1:]
        for s0, a0, b0 in REJ_SHEETS)
    czas_cnt = "+".join(
        sump(s0, "A", "E", "F", "H", a0, b0,
             expr='(--ISNUMBER(%(s)s!$I$%(a)d:$I$%(b)d))'
             % {"s": s0, "a": a0, "b": b0})[1:]
        for s0, a0, b0 in REJ_SHEETS)
    ws.write_formula(FR + 3, 10,
                     '=IF((%s)=0,"—",ROUND((%s)/(%s),1))'
                     % (czas_cnt, czas_sum, czas_cnt), f_val_box)
    ws.merge_range(FR + 4, 8, FR + 4, 10,
                   "Współwłaściciel filtruje tylko rejestr (katalog nie ma "
                   "tej kolumny). Zarejestrowane i średni czas liczone ze "
                   "wszystkich miesięcy (katalog + Archiwum).", f_note)

    # --- zestawienie miesięczne: 05.2026 – 12.2026 -------------------------
    ws.merge_range(11, 1, 11, 3, "  WG MIESIĄCA (05–12.2026)", f_sec_band)
    ws.set_row(11, 22)
    ws.write(12, 1, "Miesiąc", f_hdr)
    ws.write(12, 2, "Złożone wnioski", f_hdr)
    ws.write(12, 3, "Zarejestrowane", f_hdr)
    f_month = fmt(bg_color="white", border=1, border_color="#D9D9D9",
                  num_format="yyyy-mm", align="center", bold=True)
    for k in range(8):
        rr = 13 + k  # 0-indexed
        mcell = "$B$%d" % (rr + 1)
        ws.write_formula(rr, 1, "=DATE(2026,%d,1)" % (5 + k), f_month)
        ws.write_formula(
            rr, 2,
            "=COUNTIFS('W rejestracji'!$G$%(d)d:$G$%(l)d,\">=\"&%(m)s,"
            "'W rejestracji'!$G$%(d)d:$G$%(l)d,\"<\"&EDATE(%(m)s,1))"
            "+COUNTIFS(Zarejestrowane!$G$%(d)d:$G$%(l)d,\">=\"&%(m)s,"
            "Zarejestrowane!$G$%(d)d:$G$%(l)d,\"<\"&EDATE(%(m)s,1))"
            % {"d": DATA_ROW, "l": LAST, "m": mcell}, f_tbl_int)
        ws.write_formula(
            rr, 3,
            "=" + zsum('COUNTIFS(%%(s)s!$H$%%(a)d:$H$%%(b)d,">="&%(m)s,'
                       '%%(s)s!$H$%%(a)d:$H$%%(b)d,"<"&EDATE(%(m)s,1))'
                       % {"m": mcell}), f_tbl_int)
    ws.set_column("D:D", 16)

    def breakdown(col0, title, items, wcol, zcol):
        """Emit name/count/count table at 0-indexed col0."""
        ws.merge_range(2, col0, 2, col0 + 2, "  " + title, f_sec_band)
        ws.write(3, col0, "Nazwa", f_hdr)
        ws.write(3, col0 + 1, "W rejestracji", f_hdr)
        ws.write(3, col0 + 2, "Zarejestrowane", f_hdr)
        rr = 4
        for it in items:
            ws.write(rr, col0, it, f_tbl_text)
            ws.write_formula(
                rr, col0 + 1,
                "=COUNTIF('W rejestracji'!$%s$%d:$%s$%d,%s%d)"
                % (wcol, DATA_ROW, wcol, LAST,
                   xl_col_to_name(col0), rr + 1), f_tbl_int)
            ws.write_formula(
                rr, col0 + 2,
                "=" + zsum('COUNTIF(%%(s)s!$%s$%%(a)d:$%s$%%(b)d,%s%d)'
                           % (zcol, zcol, xl_col_to_name(col0), rr + 1)),
                f_tbl_int)
            rr += 1
        ws.write(rr, col0, "inne / brak", f_tbl_text)
        c1 = xl_col_to_name(col0 + 1)
        c2 = xl_col_to_name(col0 + 2)
        ws.write_formula(rr, col0 + 1,
                         "=%s-SUM(%s5:%s%d)" % (stat_wrej_cell, c1, c1, rr),
                         f_tbl_int)
        ws.write_formula(rr, col0 + 2,
                         "=%s-SUM(%s5:%s%d)" % (stat_zar_cell, c2, c2, rr),
                         f_tbl_int)
        rr += 1
        ws.write(rr, col0, "RAZEM", f_total_lbl)
        ws.write_formula(rr, col0 + 1,
                         "=SUM(%s5:%s%d)" % (c1, c1, rr), f_total_box)
        ws.write_formula(rr, col0 + 2,
                         "=SUM(%s5:%s%d)" % (c2, c2, rr), f_total_box)

    breakdown(4, "WG URZĘDU",
              by_count(URZEDY_CANON, [r[5] for r in wrej],
                       [z[5] for z in zarch]), "F", "F")
    breakdown(8, "WG MARKI",
              by_count(marki, [r[0] for r in wrej],
                       [z[0] for z in zarch]), "A", "A")

    # WG WSPÓŁWŁAŚCICIELA (FINANSUJĄCEGO) — pod tabelą WG URZĘDU
    col0 = 4
    rw0 = 12
    ws.merge_range(rw0, col0, rw0, col0 + 1,
                   "  WG WSPÓŁWŁAŚCICIELA", f_sec_band)
    ws.set_row(rw0, 22)
    ws.write(rw0 + 1, col0, "Nazwa", f_hdr)
    ws.write(rw0 + 1, col0 + 1, "W rejestracji", f_hdr)
    rr_w = rw0 + 2
    first_w = rr_w + 1  # 1-indeksowany pierwszy wiersz danych
    for it in by_count(wspolwl, [r[4] for r in wrej]):
        ws.write(rr_w, col0, it, f_tbl_text)
        ws.write_formula(
            rr_w, col0 + 1,
            "=COUNTIF('W rejestracji'!$E$%d:$E$%d,%s%d)"
            % (DATA_ROW, LAST, xl_col_to_name(col0), rr_w + 1), f_tbl_int)
        rr_w += 1
    ws.write(rr_w, col0, "inne / brak", f_tbl_text)
    cw = xl_col_to_name(col0 + 1)
    ws.write_formula(rr_w, col0 + 1,
                     "=%s-SUM(%s%d:%s%d)" % (stat_wrej_cell, cw, first_w,
                                             cw, rr_w),
                     f_tbl_int)
    rr_w += 1
    ws.write(rr_w, col0, "RAZEM", f_total_lbl)
    ws.write_formula(rr_w, col0 + 1,
                     "=SUM(%s%d:%s%d)" % (cw, first_w, cw, rr_w), f_total_box)

    # --- TEN TYDZIEŃ (od poniedziałku) -------------------------------------
    tw0 = rr_w + 2
    PON = "(TODAY()-WEEKDAY(TODAY(),2)+1)"
    ws.merge_range(tw0, col0, tw0, col0 + 1, "  TEN TYDZIEŃ", f_sec_band)
    ws.set_row(tw0, 22)
    tyg = [
        ("Złożone wnioski",
         "=COUNTIF('W rejestracji'!$G$%(d)d:$G$%(l)d,\">=\"&%(p)s)"
         "+COUNTIF(Zarejestrowane!$G$%(d)d:$G$%(l)d,\">=\"&%(p)s)"
         % {"d": DATA_ROW, "l": LAST, "p": PON}),
        ("Zarejestrowane",
         "=" + zsum('COUNTIF(%%(s)s!$H$%%(a)d:$H$%%(b)d,">="&%(p)s)'
                    % {"p": PON})),
        ("Planowane odbiory",
         "=COUNTIFS('W rejestracji'!$I$%(d)d:$I$%(l)d,\">=\"&%(p)s,"
         "'W rejestracji'!$I$%(d)d:$I$%(l)d,\"<\"&(%(p)s+7))"
         % {"d": DATA_ROW, "l": LAST, "p": PON}),
    ]
    for i, (lab, f) in enumerate(tyg):
        ws.write(tw0 + 1 + i, col0, lab, f_tbl_text)
        ws.write_formula(tw0 + 1 + i, col0 + 1, f, f_val_box)

    # --- WG OSOBY (prowadzący odbiory) -------------------------------------
    osoby = canonical([r[8] for r in wrej])
    if osoby:
        os0 = tw0 + 5
        ws.merge_range(os0, col0, os0, col0 + 1, "  WG OSOBY (odbiory)",
                       f_sec_band)
        ws.set_row(os0, 22)
        ro = os0 + 1
        for o in by_count(osoby, [r[8] for r in wrej]):
            ws.write(ro, col0, o, f_tbl_text)
            ws.write_formula(ro, col0 + 1,
                             "=COUNTIF('W rejestracji'!$J$%d:$J$%d,%s%d)"
                             % (DATA_ROW, LAST, xl_col_to_name(col0), ro + 1),
                             f_tbl_int)
            ro += 1

    breakdown(12, "WG DEALERA",
              by_count(dealerzy, [r[3] for r in wrej],
                       [z[4] for z in zarch]), "D", "E")

    stopka_r = max(FR + 6,
                   (ro + 2 if osoby else tw0 + 6),
                   4 + len(dealerzy) + 4)
    ws.write(stopka_r, 1, "Stworzone przez: Oskar Figaszewski", f_note)

    # ================================================================ URZĘDY
    ws = wb.add_worksheet("URZĘDY")
    sheet_order.append(ws)
    ws.set_tab_color("#6A1B9A")
    ws.hide_gridlines(2)
    for c, w in enumerate([22, 26, 30, 18, 18, 34]):
        ws.set_column(c, c, w)
    header_band(ws, "  URZĘDY I WSPÓŁWŁAŚCICIELE — ADRESY", 6)
    nav_button(ws, 5)

    f_uhdr = fmt(bold=True, bg_color="#FFF200", border=1)
    f_ucell = fmt(bg_color="white", border=1, border_color="#C8C8C8")
    f_ured = fmt(bold=True, font_color="white", bg_color=RED, border=1,
                 align="center")
    f_uwarn = fmt(bold=True, font_color="white", bg_color=RED, border=1)

    r = 2
    ws.write(r, 0, "Urzędy", f_uhdr)
    ws.write(r, 1, "KOD POCZTOWY", f_uhdr)
    ws.write(r, 2, "ADRES", f_uhdr)
    ws.write(r, 3, "CO POTRZEBA", f_uhdr)
    ws.write(r, 4, "CO POTRZEBA", f_uhdr)
    ws.write(r, 5, "ZAWSZE", f_ured)
    r += 1
    urzedy_info = [
        ("BEMOWO", "01-381 Warszawa", "ul. Powstańców Śląskich 70", "", ""),
        ("WILANÓW", "02-797 Warszawa", "ul. Franciszka Klimczaka 2", "", ""),
        ("ŚRÓDMIEŚCIE", "00-412 Warszawa", "ul. Leona Kruczkowskiego 2",
         "opłaty rozbite", ""),
        ("WAWER", "04-713 Warszawa", "ul. Żegańska 1", "", ""),
        ("OCHOTA", "02-021 Warszawa", "ul. Grójecka 17a", "daty doki",
         "KOPIE DOKI"),
        ("MOKOTÓW", "02-517 Warszawa", "ul. Rakowiecka 25/27", "", ""),
    ]
    for naz, kod, adr, p1, p2 in urzedy_info:
        ws.write(r, 0, naz, f_ucell)
        ws.write(r, 1, kod, f_ucell)
        ws.write(r, 2, adr, f_ucell)
        ws.write(r, 3, p1, f_uwarn if p1 else f_ucell)
        ws.write(r, 4, p2, f_uwarn if p2 else f_ucell)
        ws.write(r, 5, "DATA FV i UMOWA PRZEWŁASZCZENIA", f_ured)
        r += 1

    r += 2
    ws.write(r, 0, "WSPÓŁWŁ (FINANSUJĄCY)", f_uhdr)
    ws.write(r, 1, "ADRES", f_uhdr)
    ws.write(r, 2, "KOD", f_uhdr)
    ws.write(r, 3, "REGON", f_uhdr)
    r += 1
    for naz, adr, kod, regon in [
            ("BMW Financial Services Polska Sp. z o.o.", "ul. Wołoska 22A",
             "02-675 Warszawa", "REGON 143424261"),
            ("MERCEDES-BENZ LEASING POLSKA Sp. z o.o.",
             "ul. Gottlieba Daimlera 1", "02-460 Warszawa", "REGON 012213933"),
            ("PKO LEASING SPÓŁKA AKCYJNA", "ul. Świętokrzyska 36",
             "00-116 Warszawa", "REGON 472191767"),
            ("VELO LEASING S.A.", "Rondo Ignacego Daszyńskiego 2C",
             "00-843 Warszawa", "REGON 367715275"),
            ("M Leasing Sp. z o.o.", "ul. Prosta 18",
             "00-850 Warszawa", "REGON 012527809")]:
        ws.write(r, 0, naz, f_ucell)
        ws.write(r, 1, adr, f_ucell)
        ws.write(r, 2, kod, f_ucell)
        ws.write(r, 3, regon, f_ucell)
        r += 1

    r += 2
    ws.write(r, 0, "Urzędy (kody wyróżników)", f_uhdr)
    ws.write(r, 1, "ORGAN", f_uhdr)
    ws.write(r, 2, "ADRES", f_uhdr)
    ws.write(r, 3, "KOD POCZTOWY", f_uhdr)
    r += 1
    for naz, org, adr, kod in [
            ("BEMOWO (WB)", "Prezydent m.st. Warszawy",
             "ul. Powstańców Śląskich 70", "01-381 Warszawa"),
            ("WILANÓW (WW)", "Prezydent m.st. Warszawy",
             "ul. Franciszka Klimczaka 2", "02-797 Warszawa"),
            ("ŚRÓDMIEŚCIE (WI)", "Prezydent m.st. Warszawy",
             "ul. Leona Kruczkowskiego 2", "00-412 Warszawa"),
            ("WAWER (WT)", "Prezydent m.st. Warszawy",
             "ul. Żegańska 1", "04-713 Warszawa"),
            ("BIAŁOŁĘKA (WA)", "Prezydent m.st. Warszawy",
             "ul. Modlińska 197", "03-122 Warszawa"),
            ("WOLA (WY)", "Prezydent m.st. Warszawy",
             "ul. Solidarności 90", "01-003 Warszawa"),
            ("MOKOTÓW (WE)", "Prezydent m.st. Warszawy",
             "ul. Rakowiecka 25/27", "02-517 Warszawa"),
            ("URSYNÓW (WN)", "Prezydent m.st. Warszawy",
             "al. Komisji Edukacji Narodowej 61", "02-777 Warszawa"),
            ("LEGIONOWO (WL)", "Starosta Legionowski",
             "ul. gen. Władysława Sikorskiego 11", "05-119 Legionowo"),
            ("GDYNIA (GA)", "Prezydent Miasta Gdyni",
             "al. Marszałka Piłsudskiego 52/54", "81-382 Gdynia"),
            ("NOWY DWÓR MAZ (WND)", "Starosta Nowodworski",
             "ul. Paderewskiego 1B", "05-100 Nowy Dwór Mazowiecki")]:
        ws.write(r, 0, naz, f_ucell)
        ws.write(r, 1, org, f_ucell)
        ws.write(r, 2, adr, f_ucell)
        ws.write(r, 3, kod, f_ucell)
        r += 1

    # =============================================================== PRZELEWY
    ws = wb.add_worksheet("PRZELEWY")
    sheet_order.append(ws)
    ws.set_tab_color("#00695C")
    ws.hide_gridlines(2)
    for c, w in enumerate([16, 20, 42, 14, 42]):
        ws.set_column(c, c, w)
    header_band(ws, "  PRZELEWY ZA REJESTRACJĘ — KONTA URZĘDÓW", 5)
    nav_button(ws, 4)

    ws.write(2, 0, "ADRESACI MAILA Z PROŚBĄ O PRZELEW", f_sec)
    ws.write(3, 0, "Do:", f_lbl)
    ws.write(3, 1, "ksiegowosc@99rent.pl", f_ucell)
    ws.write(3, 2, "e.cyc@99rent.pl", f_ucell)
    ws.write(4, 0, "Cc:", f_lbl)
    cc = ["n.jastrzebska@99rent.pl", "d.wroblewski@99rent.pl",
          "a.kokot@99rent.pl", "m.wieczorek@99rent.pl", "k.pietrzyk@99rent.pl",
          "r.piestrzynski@99rent.pl", "j.wojtynska@99rent.pl",
          "flota@99rent.pl", "k.dudek@99rent.pl"]
    for i, adres in enumerate(cc):
        ws.write(4 + i // 3, 1 + i % 3, adres, f_ucell)

    ws.write(9, 0, "Przykład tytułu przelewu:", f_lbl)
    ws.merge_range(9, 1, 9, 4,
                   "Przelew za rejestrację pojazdów WAWER 5x BMW INCHCAPE "
                   "+ VIN każdego auta (VIN-y połącz: =POŁĄCZ.TEKSTY(\"; \";1;zakres))",
                   f_ucell)

    ws.write(11, 0, "KONTA URZĘDÓW", f_sec)
    ws.write(12, 0, "URZĄD", f_hdr)
    ws.write(12, 1, "Przelew 1 — rejestracja", f_hdr)
    ws.write(12, 2, "Konto (rejestracja)", f_hdr)
    ws.write(12, 3, "Przelew 2 — pełnomocn.", f_hdr)
    ws.write(12, 4, "Konto (pełnomocnictwa)", f_hdr)
    ws.set_row(12, 26)
    KONTO_PELN = "21 1030 1508 0000 0005 5000 0070"
    przelewy = [
        ("OCHOTA", "160,00 zł / pojazd", "20 1030 1508 0000 0005 5002 4047"),
        ("ŚRÓDMIEŚCIE", "160,00 zł / pojazd", "07 1030 1508 0000 0005 5001 0119"),
        ("WAWER", "160,00 zł / pojazd", "77 1030 1508 0000 0005 5003 2139"),
        ("BEMOWO", "160,00 zł / pojazd", "50 1030 1508 0000 0005 5000 2167"),
        ("WILANÓW", "160,00 zł / pojazd", "51 1030 1508 0000 0005 5001 6117"),
    ]
    r = 13
    for urzad, kwota, konto in przelewy:
        ws.write(r, 0, urzad, fmt(bold=True, bg_color="white", border=1,
                                  border_color="#C8C8C8"))
        ws.write(r, 1, kwota, f_ucell)
        ws.write(r, 2, konto, f_ucell)
        ws.write(r, 3, "17,00 zł", f_ucell)
        ws.write(r, 4, KONTO_PELN, f_ucell)
        r += 1
    ws.merge_range(r + 1, 0, r + 3, 4,
                   "W tytule przelewu zbiorczego należy wpisać VIN każdego "
                   "samochodu. Zawsze poproś o wysłanie potwierdzenia każdego "
                   "przelewu. Konto przelewu 2 (pełnomocnictwa, 17 zł) jest "
                   "wspólne dla wszystkich urzędów.", f_instr)

    # gotowe teksty maili do kopiuj-wklej (1:1 z dokumentu PRZELEWY.docx)
    f_copy = fmt(bg_color="white", border=1, border_color="#C8C8C8",
                 text_wrap=True, valign="top")
    teksty = [
        ("OCHOTA", "Cześć,", "wpisać VIN każdego samochodu",
         "20 1030 1508 0000 0005 5002 4047", "Ochota"),
        ("ŚRÓDMIEŚCIE", "Hej,", "wpisać VIN każdego samochodu",
         "07 1030 1508 0000 0005 5001 0119", "Śródmieście"),
        ("WAWER", "Hej,", "VIN każdego samochodu",
         "77 1030 1508 0000 0005 5003 2139", "Wawer"),
        ("BEMOWO", "Cześć,", "VIN każdego samochodu",
         "50 1030 1508 0000 0005 5000 2167", "Bemowo"),
        ("WILANÓW", "Hej,", "wpisać VIN każdego samochodu",
         "51 1030 1508 0000 0005 5001 6117", "Wilanów"),
    ]
    r = r + 5
    for urzad, powitanie, vin_fraza, konto1, dzielnica in teksty:
        ws.merge_range(r, 0, r, 4,
                       "  TEKST DO SKOPIOWANIA — " + urzad, f_sec_band)
        ws.set_row(r, 22)
        tekst = (
            powitanie + "\n\n"
            "Bardzo proszę o PILNY przelew za rejestrację pojazdów na konto "
            "Urzędu Dzielnicy " + dzielnica + ". W tytule przelewu zbiorczego "
            "należy " + vin_fraza + ". Proszę również o wysłanie "
            "potwierdzenia każdego przelewu.\n\n"
            "Przelew 1\n"
            "Kwota przelewu za rejestrację pojazdu: 160,00 zł\n"
            "Nr konta bankowego do przelewu: " + konto1 + "\n"
            "W tytule przelewu należy wpisać: \n\n"
            "Przelew 2\n"
            "Kwota przelewu za pełnomocnictwa: 17,00 zł\n"
            "Nr konta bankowego do przelewu: 21 1030 1508 0000 0005 5000 0070\n"
            "W tytule przelewu należy wpisać: ")
        ws.merge_range(r + 1, 0, r + 12, 4, tekst, f_copy)
        r += 14

    ws = wb.add_worksheet("Listy")
    sheet_order.append(ws)
    for c, (title, vals) in enumerate([
            ("Marki", marki), ("Dealerzy", dealerzy),
            ("Urzędy", URZEDY_CANON), ("Współwłaściciele", wspolwl)]):
        ws.write(0, c, title, f_hdr)
        for i, v in enumerate(vals):
            ws.write_string(1 + i, c, v, f_text)
        ws.set_column(c, c, 22)
    # E: podręczny "kalendarz" — ostatnie 31 dni do szybkiego wyboru daty
    ws.write(0, 4, "Daty (31 dni)", f_hdr)
    for i in range(31):
        ws.write_formula(1 + i, 4, "=TODAY()-%d" % i, f_date)
    ws.set_column(4, 4, 14)
    # J: daty w przód (planowany odbiór)
    ws.write(0, 9, "Daty w przód", f_hdr)
    for i in range(15):
        ws.write_formula(1 + i, 9, "=TODAY()+%d" % i, f_date)
    ws.set_column(9, 9, 14)
    # F-I: listy do filtrów PODSUMOWANIA ("(wszystkie)" + wartości)
    MIESIACE = ["2026-%02d" % m for m in range(5, 13)]
    for c, (title, vals) in enumerate([
            ("Filtr miesiąc", MIESIACE), ("Filtr urząd", URZEDY_CANON),
            ("Filtr marka", marki), ("Filtr dealer", dealerzy)], start=5):
        ws.write(0, c, title, f_hdr)
        ws.write_string(1, c, "(wszystkie)", f_text)
        for i, v in enumerate(vals):
            ws.write_string(2 + i, c, v, f_text)
        ws.set_column(c, c, 22)
    ws.hide()

    # ===================================================== ARCHIWA MIESIĘCY
    zhdr_arch = ["Marka", "Model", "VIN", "Nr rejestracyjny", "Dealer",
                 "Urząd", "Data złożenia", "Data rejestracji",
                 "Czas rejestracji (dni)", "Uwagi"]
    for klucz in sorted(stare_by_month):
        ws = wb.add_worksheet("Archiwum %s" % klucz)
        sheet_order.append(ws)
        ws.set_tab_color("#787878")
        for c, w in enumerate([16, 20, 23, 17, 18, 15, 14, 15, 16, 32]):
            ws.set_column(c, c, w)

        # --- statystyki miesiąca na górze (jak PULPIT / PODSUMOWANIE) ------
        header_band(ws, "  ARCHIWUM %s — ZAREJESTROWANE" % klucz, 10)
        arng = "$%%s$%d:$%%s$5000" % ARCH_DATA
        rng = lambda col: (arng % (col, col))
        ws.set_row(1, 34)
        ws.merge_range(1, 0, 1, 1, "ZAREJESTROWANE %s" % klucz, f_cnt_lbl)
        ws.write_formula(
            1, 2, "=SUMPRODUCT(--((%s&%s)<>\"\"))" % (rng("A"), rng("C")),
            f_cnt_num)
        ws.merge_range(1, 4, 1, 5, "ŚREDNI CZAS REJESTRACJI (DNI)", f_cnt_lbl)
        ws.write_formula(
            1, 6, "=IF(COUNT(%s)=0,\"—\",ROUND(AVERAGE(%s),1))"
            % (rng("I"), rng("I")), f_cnt_num)

        ws.merge_range(2, 0, 2, 1, "  WG URZĘDU", f_sec_band)
        ws.merge_range(2, 3, 2, 4, "  WG MARKI", f_sec_band)
        ws.merge_range(2, 6, 2, 7, "  WG DEALERA", f_sec_band)
        ws.set_row(2, 22)

        def mini_tbl(col0, items, datacol):
            rr = 3
            for it in items:
                ws.write(rr, col0, it, f_tbl_text)
                ws.write_formula(
                    rr, col0 + 1, "=COUNTIF(%s,%s%d)"
                    % (rng(datacol), xl_col_to_name(col0), rr + 1), f_tbl_int)
                rr += 1
            ws.write(rr, col0, "inne / brak", f_tbl_text)
            cc = xl_col_to_name(col0 + 1)
            ws.write_formula(rr, col0 + 1,
                             "=$C$2-SUM(%s4:%s%d)" % (cc, cc, rr), f_tbl_int)
            for pad in range(rr + 1, 12):
                ws.write_blank(pad, col0, None)

        cnt_u = collections.Counter(z[5] for z in stare_by_month[klucz] if z[5])
        cnt_m = collections.Counter(z[0] for z in stare_by_month[klucz] if z[0])
        cnt_d = collections.Counter(z[4] for z in stare_by_month[klucz] if z[4])
        mini_tbl(0, [u for u, _ in cnt_u.most_common(8)], "F")
        mini_tbl(3, [m for m, _ in cnt_m.most_common(8)], "A")
        mini_tbl(6, [d for d, _ in cnt_d.most_common(8)], "E")

        # --- tabela danych od wiersza ARCH_DATA ----------------------------
        for c, h in enumerate(zhdr_arch):
            ws.write(ARCH_DATA - 2, c, h, f_hdr)
        ws.set_row(ARCH_DATA - 2, 24)
        for r, z in enumerate(stare_by_month[klucz], start=ARCH_DATA - 1):
            marka, model, vin, nrrej, dealer, urzad, dzl, datarej, cena, uwagi = z
            for c, v in ((0, marka), (1, model), (2, vin), (3, nrrej),
                         (4, dealer), (5, urzad), (9, uwagi)):
                if v:
                    ws.write_string(r, c, v, f_text_g)
                else:
                    ws.write_blank(r, c, None, f_text_g)
            if dzl is not None:
                ws.write_datetime(r, 6, dzl, f_date_g)
            else:
                ws.write_blank(r, 6, None, f_date_g)
            ws.write_datetime(r, 7, datarej, f_date_g)
            czas = None
            if dzl is not None and id(z) not in synth:
                czas = (datarej.date() - dzl.date()).days
                if czas < 0:
                    czas = None
            if czas is not None:
                ws.write_number(r, 8, czas, f_int_g)
            else:
                ws.write_blank(r, 8, None, f_int_g)
        ws.freeze_panes(ARCH_DATA - 1, 0)
        ws.autofilter(ARCH_DATA - 2, 0,
                      ARCH_DATA - 2 + len(stare_by_month[klucz]), 9)

    if with_vba:
        for i, s in enumerate(sheet_order):
            s.set_vba_name("Arkusz%d" % (i + 1))

    wb.close()
    return len(sheet_order)


if __name__ == "__main__":
    import sys
    n = build(sys.argv[1], with_vba=sys.argv[1].endswith(".xlsm"),
              vba_bin="vbaProject.bin")
    print("built", sys.argv[1], "sheets:", n)
