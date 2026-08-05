# -*- coding: utf-8 -*-
"""Build the 99rent vehicle-registration workbook (.xlsm with macros/buttons,
or a formula-identical .xlsx copy for LibreOffice recalc verification)."""
import collections
import datetime
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

    wrej = [(m_marka.get(r[0], r[0]), r[1], r[2], m_dealer.get(r[3], r[3]),
             m_wsp.get(r[4], r[4]), fix_urzad(r[5]), r[6], r[7], r[8])
            for r in wrej]
    zarej = [(m_marka.get(z[0], z[0]), z[1], z[2], z[3],
              m_dealer.get(z[4], z[4]), fix_urzad(z[5]), z[6], z[7], z[8], z[9])
             for z in zarej]

    pilne = []
    for r in wb["Arkusz4"].iter_rows(min_row=4, max_col=2, values_only=True):
        if r[0] or r[1]:
            pilne.append((s(r[0]), s(r[1])))
    return wrej, zarej, pilne


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
    wrej = []
    ws = wb["W rejestracji"]
    for r in range(DATA_ROW, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 10)]
        if not any(s(v) for v in vals[:7]):
            continue
        f = ws.cell(row=r, column=1).fill
        rgb = str(f.fgColor.rgb) if f.patternType == "solid" else ""
        wrej.append((s(vals[0]), s(vals[1]), s(vals[2]), s(vals[3]),
                     s(vals[4]), s(vals[5]), vals[6], s(vals[8]),
                     rgb == "FFFFC7CE"))
    def rows_of(ws, first):
        out = []
        for r in range(first, ws.max_row + 1):
            v = [ws.cell(row=r, column=c).value for c in range(1, 11)]
            if not s(v[2]):
                continue
            out.append((s(v[0]), s(v[1]), s(v[2]), s(v[3]), s(v[4]),
                        s(v[5]), v[6], v[7], None, s(v[9])))
        return out
    zarej = rows_of(wb["Zarejestrowane"], DATA_ROW)
    archiwa = {}
    for name in wb.sheetnames:
        if name.startswith("Archiwum "):
            archiwa[name.split()[1]] = rows_of(wb[name], 2)
    pilne = []
    if "Pilne" in wb.sheetnames:
        for r in wb["Pilne"].iter_rows(min_row=4, max_col=2, values_only=True):
            if r[0] or r[1]:
                pilne.append((s(r[0]), s(r[1])))
    return wrej, zarej, archiwa, pilne


def build(path, with_vba, vba_bin=None, logo="logo99rent.png",
          logo_small="logo99rent_small.png"):
    import os
    if os.path.exists(STATE):
        wrej, zarej, stare_by_month, pilne = load_state(STATE)
    else:
        wrej, zarej, pilne = load_data()
        prog = (datetime.date.today().year, datetime.date.today().month)
        stare_by_month = {}
        for z in zarej:
            if z[7] is not None and (z[7].year, z[7].month) < prog:
                stare_by_month.setdefault("%04d-%02d" % (z[7].year, z[7].month),
                                          []).append(z)
        stare_ids = {id(z) for rows in stare_by_month.values() for z in rows}
        zarej = [z for z in zarej if id(z) not in stare_ids]

    # arkusze, po których liczą się "zarejestrowane": katalog + archiwa
    REJ_SHEETS = [("Zarejestrowane", DATA_ROW, LAST)] + \
        [("'Archiwum %s'" % k, 2, 5000) for k in sorted(stare_by_month)]

    def zsum(tmpl):
        """Suma formuły po katalogu i wszystkich zakładkach Archiwum.
        tmpl używa %(s)s (arkusz), %(a)d (pierwszy wiersz), %(b)d (ostatni)."""
        return "+".join(tmpl % {"s": s0, "a": a0, "b": b0}
                        for s0, a0, b0 in REJ_SHEETS)
    marki = canonical([r[0] for r in wrej] + [z[0] for z in zarej])
    dealerzy = canonical([r[3] for r in wrej] + [z[4] for z in zarej])
    wspolwl = canonical([r[4] for r in wrej])

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
    f_cnt_lbl = fmt(bold=True, font_size=11, align="right", valign="vcenter")
    f_cnt_num = fmt(bold=True, font_color=RED, font_size=14, align="center",
                    valign="vcenter", bg_color="white", border=1,
                    border_color="#D9D9D9")

    def header_band(ws, title, ncols):
        ws.set_row(0, 40)
        ws.merge_range(0, 0, 0, ncols - 2, title, f_band)
        ws.write(0, ncols - 1, "99rent", f_band_sub)
        ws.insert_image(0, 0, logo_small,
                        {"x_scale": 0.14, "y_scale": 0.14,
                         "x_offset": 4, "y_offset": 4,
                         "object_position": 3})

    def nav_button(ws, col):
        if with_vba:
            ws.insert_button(2, col, {"macro": "IdzPulpit", "caption": "◀ PULPIT",
                                      "width": 90, "height": 24})

    sheet_order = []

    # ================================================================ PULPIT
    ws = wb.add_worksheet("PULPIT")
    sheet_order.append(ws)
    ws.set_tab_color(RED)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 2)
    ws.set_column("B:C", 14)
    ws.set_column("D:K", 12)
    header_band(ws, "  RAPORT REJESTRACJI POJAZDÓW", 11)
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
                ("IMPORT HURTOWY", "IdzImport"),
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
        "Kolumna „Dni od złożenia” liczy się sama i podświetla pojazdy czekające zbyt długo (pomarańczowy > 10 dni, czerwony > 21 dni).\n"
        "3.  IMPORT HURTOWY — wklejasz wiele pojazdów naraz i przyciskiem dodajesz je do rejestru (zaznaczenie wierszy = import tylko wybranych); urząd wybierzesz w rejestrze z listy w kolumnie F.\n"
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
    f_leg_y = fmt(bg_color="#FFF9C4", border=1, border_color="#9E9E9E",
                  align="center", font_size=9)
    f_leg_g = fmt(bg_color="#C6EFCE", border=1, border_color="#9E9E9E",
                  align="center", font_size=9)
    f_leg_r = fmt(bg_color="#FFC7CE", border=1, border_color="#9E9E9E",
                  align="center", font_size=9)
    ws.write(19, 1, "LEGENDA KOLORÓW:", f_lbl)
    ws.merge_range(19, 3, 19, 4, "w rejestracji (złożone)", f_leg_y)
    ws.merge_range(19, 5, 19, 6, "zarejestrowany", f_leg_g)
    ws.merge_range(19, 7, 19, 8, "wymaga uwagi / zaległy", f_leg_r)
    ws.write(21, 1, "Wygenerowano na podstawie pliku Raport_rejestracji.xlsx", f_note)

    # ========================================================= W REJESTRACJI
    ws = wb.add_worksheet("W rejestracji")
    sheet_order.append(ws)
    ws.set_tab_color("#F4A100")
    headers = ["Marka", "Model", "VIN", "Dealer", "Współwłaściciel", "Urząd",
               "Data złożenia", "Dni od złożenia", "Uwagi"]
    widths = [14, 20, 23, 17, 17, 15, 14, 14, 32]
    colfmts = [f_text, f_text, f_text, f_text, f_text, f_text, f_date, f_int, f_text]
    for c, (w, cf) in enumerate(zip(widths, colfmts)):
        ws.set_column(c, c, w, cf)
    ws.set_column(9, 9, 2)
    ws.set_column(10, 10, 22)
    header_band(ws, "  POJAZDY W TRAKCIE REJESTRACJI", 9)
    nav_button(ws, 10)
    ws.set_row(1, 24)
    ws.merge_range(1, 0, 1, 1, "Ilość w rejestracji:", f_cnt_lbl)
    ws.write_formula(
        1, 2, "=SUMPRODUCT(--(($A$%d:$A$%d&$C$%d:$C$%d)<>\"\"))"
        % (DATA_ROW, LAST, DATA_ROW, LAST), f_cnt_num)

    ws.merge_range(2, 0, 2, 8,
                   "FORMULARZ — NOWY WNIOSEK:  wypełnij żółte pola i dodaj przez IMPORT HURTOWY lub wpisz bezpośrednio w tabeli",
                   f_form_title)
    for c, h in enumerate(headers):
        ws.write(3, c, h, f_form_label)
    for c in range(9):
        if c == 6:
            ws.write_blank(FORM_ROW - 1, c, None, f_input_date)
        elif c == 7:
            ws.write(FORM_ROW - 1, c, "auto", f_auto)
        else:
            ws.write_blank(FORM_ROW - 1, c, None, f_input)
    ws.set_row(FORM_ROW - 1, 22)
    if with_vba:
        ws.insert_button(6, 10, {"macro": "ZarejestrujZaznaczone",
                                 "caption": "ZAREJESTRUJ ZAZNACZONE ▶",
                                 "width": 170, "height": 34})
        ws.write(9, 10, "Zaznacz wiersze pojazdów i kliknij, aby przenieść "
                 "je do katalogu ZAREJESTROWANE (nr rej. wpiszesz tam w kolumnie D).",
                 f_note)

    for c, h in enumerate(headers):
        ws.write(HDR_ROW - 1, c, h, f_hdr)
    ws.set_row(HDR_ROW - 1, 28)

    r = DATA_ROW - 1  # 0-indexed
    for row in wrej:
        marka, model, vin, dealer, wsp, urzad, dzl, uwagi, is_red = row
        ftxt = f_text_red if is_red else f_text_y
        fdat = f_date_red if is_red else f_date_y
        fint = f_int_red if is_red else f_int_y
        for c, v in ((0, marka), (1, model), (2, vin), (3, dealer),
                     (4, wsp), (5, urzad), (8, uwagi)):
            if v is not None:
                ws.write_string(r, c, v, ftxt)
            else:
                ws.write_blank(r, c, None, ftxt)
        if dzl is not None:
            ws.write_datetime(r, 6, dzl, fdat)
        else:
            ws.write_blank(r, 6, None, fdat)
        ws.write_formula(
            r, 7, '=IF($G%d="","",TODAY()-$G%d)' % (r + 1, r + 1), fint)
        r += 1
    last_data = r  # 0-indexed row after last

    ws.autofilter(HDR_ROW - 1, 0, LAST - 1, 8)
    ws.freeze_panes(HDR_ROW, 0)
    ws.conditional_format(DATA_ROW - 1, 7, LAST - 1, 7,
                          {"type": "cell", "criteria": ">", "value": 21,
                           "format": f_red})
    ws.conditional_format(DATA_ROW - 1, 7, LAST - 1, 7,
                          {"type": "cell", "criteria": "between",
                           "minimum": 11, "maximum": 21, "format": f_amber})
    ws.conditional_format(DATA_ROW - 1, 2, LAST - 1, 2,
                          {"type": "duplicate", "format": f_red})
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

    # ======================================================= IMPORT HURTOWY
    ws = wb.add_worksheet("Import hurtowy")
    sheet_order.append(ws)
    ws.set_tab_color("#F9A825")
    ihdr = ["Marka", "Model", "VIN", "Dealer", "Współwłaściciel", "Urząd",
            "Data złożenia", "Uwagi"]
    iw = [14, 20, 23, 17, 17, 15, 14, 32]
    for c, w in enumerate(iw):
        ws.set_column(c, c, w)
    ws.set_column(8, 8, 2)
    ws.set_column(9, 9, 26)
    header_band(ws, "  IMPORT HURTOWY — WIELE POJAZDÓW NARAZ", 8)
    nav_button(ws, 9)

    ws.merge_range(2, 0, 2, 7,
                   "Wklej 10, 20, 30… pojazdów do tabeli (od wiersza 9) i kliknij "
                   "IMPORTUJ DO REJESTRU (możesz też zaznaczyć tylko wybrane wiersze)", f_form_title)
    ws.merge_range(3, 0, 5, 7,
                   "Wymagany jest VIN (kolumna C), reszta pól opcjonalna. "
                   "Zaznacz wiersze do przeniesienia (bez zaznaczenia przenosi się "
                   "wszystko). Pusta data złożenia = dzisiejsza data. Duplikaty VIN "
                   "są pomijane. Wnioski trafiają do W REJESTRACJI na żółto, a urząd "
                   "możesz wybrać tam z listy w kolumnie F.", f_instr)
    if with_vba:
        ws.insert_button(2, 9, {"macro": "PrzeniesWnioski",
                                "caption": "IMPORTUJ DO REJESTRU",
                                "width": 180, "height": 44})

    for c, h in enumerate(ihdr):
        ws.write(HDR_ROW - 1, c, h, f_hdr)
    ws.set_row(HDR_ROW - 1, 28)
    for rr in range(DATA_ROW - 1, DATA_ROW - 1 + 300):
        for c in range(8):
            ws.write_blank(rr, c, None,
                           f_import_date if c == 6 else f_import)
    ws.data_validation(DATA_ROW - 1, 6, DATA_ROW - 1 + 299, 6,
                       {"validate": "list", "source": "=Listy!$E$2:$E$32",
                        "show_error": False})
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
    ws.set_row(1, 24)
    ws.merge_range(1, 0, 1, 1, "W katalogu (bieżący mies.):", f_cnt_lbl)
    ws.write_formula(
        1, 2, "=SUMPRODUCT(--(($A$%d:$A$%d&$C$%d:$C$%d)<>\"\"))"
        % (DATA_ROW, LAST, DATA_ROW, LAST), f_cnt_num)
    ws.merge_range(1, 4, 1, 5, "Łącznie z archiwum:", f_cnt_lbl)
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
            r, 8, '=IF(OR($G%d="",$H%d=""),"",$H%d-$G%d)'
            % (r + 1, r + 1, r + 1, r + 1), f_int_g)
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
    ws.set_column("C:C", 12)
    ws.set_column("D:D", 3)
    ws.set_column("E:E", 20)
    ws.set_column("F:G", 13)
    ws.set_column("H:H", 3)
    ws.set_column("I:I", 20)
    ws.set_column("J:K", 13)
    ws.set_column("L:L", 3)
    ws.set_column("M:M", 20)
    ws.set_column("N:O", 13)
    header_band(ws, "  PODSUMOWANIE REJESTRACJI", 15)
    ws.hide_gridlines(2)
    ws.set_row(1, 30)
    if with_vba:
        ws.insert_button(1, 1, {"macro": "GenerujRaport",
                                "caption": "GENERUJ RAPORT (PLIK)",
                                "width": 190, "height": 28})
        ws.insert_button(1, 3, {"macro": "IdzPulpit",
                                "caption": "◀ PULPIT",
                                "width": 100, "height": 28})

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
    ws.merge_range(11, 1, 11, 3, "  FILTR — policz wg wybranych kryteriów", f_sec_band)
    ws.set_row(11, 22)
    filt = [("Miesiąc", "F", 9), ("Urząd", "G", len(URZEDY_CANON) + 1),
            ("Marka", "H", len(marki) + 1), ("Dealer", "I", len(dealerzy) + 1)]
    for i, (label, lcol, n) in enumerate(filt):
        rr = 12 + i
        ws.write(rr, 1, label, f_tbl_text)
        ws.write_string(rr, 2, "(wszystkie)", f_input)
        ws.data_validation(rr, 2, rr, 2,
                           {"validate": "list",
                            "source": "=Listy!$%s$2:$%s$%d" % (lcol, lcol, n + 1),
                            "show_error": False})
    FM, FU, FMA, FD = "$C$13", "$C$14", "$C$15", "$C$16"

    def sump(sheet, cmarka, cdealer, curzad, cdata, d=DATA_ROW, l=LAST):
        g = lambda col: "%s!$%s$%d:$%s$%d" % (sheet, col, d, col, l)
        return ("=SUMPRODUCT((%(vin)s<>\"\")"
                "*IF(%(fu)s=\"(wszystkie)\",1,--(%(urz)s=%(fu)s))"
                "*IF(%(fm)s=\"(wszystkie)\",1,--(%(mar)s=%(fm)s))"
                "*IF(%(fd)s=\"(wszystkie)\",1,--(%(dea)s=%(fd)s))"
                "*IF(%(fmies)s=\"(wszystkie)\",1,"
                "(%(dat)s>=DATEVALUE(%(fmies)s&\"-01\"))"
                "*(%(dat)s<EDATE(DATEVALUE(%(fmies)s&\"-01\"),1))))"
                % {"vin": g("C"), "urz": g(curzad), "mar": g(cmarka),
                   "dea": g(cdealer), "dat": g(cdata),
                   "fu": FU, "fm": FMA, "fd": FD, "fmies": FM})

    ws.write(17, 1, "W rejestracji (wg daty złożenia)", f_tbl_text)
    ws.write_formula(17, 2, sump("'W rejestracji'", "A", "D", "F", "G"), f_val_box)
    ws.write(18, 1, "Zarejestrowane (wg daty rejestracji)", f_tbl_text)
    ws.write_formula(18, 2, "=" + "+".join(
        sump(s0, "A", "E", "F", "H", a0, b0)[1:] for s0, a0, b0 in REJ_SHEETS),
        f_val_box)
    ws.merge_range(15, 3, 18, 5,
                   "Wybierz wartości z list (żółte pola) — liczniki obok "
                   "przeliczają się od razu. „(wszystkie)” wyłącza dany filtr. "
                   "Miesiąc: dla rejestru liczy się data złożenia, dla katalogu "
                   "data rejestracji. Miesiące przeniesione do zakładek "
                   "Archiwum liczy tabela WG MIESIĄCA poniżej.", f_note)

    # --- zestawienie miesięczne: 05.2026 – 12.2026 -------------------------
    ws.merge_range(20, 1, 20, 3, "  WG MIESIĄCA (05–12.2026)", f_sec_band)
    ws.set_row(20, 22)
    ws.write(21, 1, "Miesiąc", f_hdr)
    ws.write(21, 2, "Złożone wnioski", f_hdr)
    ws.write(21, 3, "Zarejestrowane", f_hdr)
    f_month = fmt(bg_color="white", border=1, border_color="#D9D9D9",
                  num_format="yyyy-mm", align="center", bold=True)
    for k in range(8):
        rr = 22 + k  # 0-indexed
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
    ws.set_column("D:D", 14)

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

    breakdown(4, "WG URZĘDU", URZEDY_CANON, "F", "F")
    breakdown(8, "WG MARKI", marki, "A", "A")

    # WG WSPÓŁWŁAŚCICIELA (FINANSUJĄCEGO) — pod tabelą WG MARKI, z odstępem
    col0 = 8
    rw0 = 20
    ws.merge_range(rw0, col0, rw0, col0 + 1,
                   "  WG WSPÓŁWŁAŚCICIELA", f_sec_band)
    ws.set_row(rw0, 22)
    ws.write(rw0 + 1, col0, "Nazwa", f_hdr)
    ws.write(rw0 + 1, col0 + 1, "W rejestracji", f_hdr)
    rr_w = rw0 + 2
    first_w = rr_w + 1  # 1-indeksowany pierwszy wiersz danych
    for it in wspolwl:
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
    breakdown(12, "WG DEALERA", dealerzy, "D", "E")

    # ================================================================= PILNE
    ws = wb.add_worksheet("Pilne")
    sheet_order.append(ws)
    ws.set_tab_color("#B71C1C")
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 22)
    ws.set_column("C:D", 10)
    header_band(ws, "  POJAZDY PILNE", 4)
    ws.write(2, 0, "Lista priorytetowa z pierwotnego raportu", f_note)
    ws.write(3, 0, "Marka i model", f_hdr)
    ws.write(3, 1, "VIN", f_hdr)
    r = 4
    for nm, vin in pilne:
        if nm: ws.write_string(r, 0, nm, f_text)
        if vin: ws.write_string(r, 1, vin, f_text)
        r += 1

    # ================================================================= LISTY
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
        for c, w in enumerate([14, 20, 23, 17, 17, 15, 14, 15, 15, 32]):
            ws.set_column(c, c, w)
        for c, h in enumerate(zhdr_arch):
            ws.write(0, c, h, f_hdr)
        ws.set_row(0, 24)
        for r, z in enumerate(stare_by_month[klucz], start=1):
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
            czas = (datarej.date() - dzl.date()).days if dzl is not None else None
            if czas is not None:
                ws.write_number(r, 8, czas, f_int_g)
            else:
                ws.write_blank(r, 8, None, f_int_g)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(stare_by_month[klucz]), 9)

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
