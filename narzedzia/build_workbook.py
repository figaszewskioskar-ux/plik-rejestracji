# -*- coding: utf-8 -*-
"""Build the 99rent vehicle-registration workbook (.xlsm with macros/buttons,
or a formula-identical .xlsx copy for LibreOffice recalc verification)."""
import collections
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

    pilne = []
    for r in wb["Arkusz4"].iter_rows(min_row=4, max_col=2, values_only=True):
        if r[0] or r[1]:
            pilne.append((s(r[0]), s(r[1])))
    return wrej, zarej, pilne


def canonical(values):
    """Group case-insensitively, return most common spelling of each, sorted."""
    groups = collections.defaultdict(collections.Counter)
    for v in values:
        if v:
            groups[v.casefold()][v] += 1
    return sorted((c.most_common(1)[0][0] for c in groups.values()),
                  key=lambda x: x.casefold())


URZEDY_CANON = ["BEMOWO", "BIAŁOŁĘKA", "OCHOTA", "ŚRÓDMIEŚCIE", "WAWER", "WILANÓW"]


def build(path, with_vba, vba_bin=None, logo="logo99rent.png",
          logo_small="logo99rent_small.png"):
    wrej, zarej, pilne = load_data()
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
                 align="left", valign="vcenter", indent=1)
    f_band_sub = fmt(bg_color=RED, font_color="white", font_size=10,
                     align="right", valign="vcenter")
    f_hdr = fmt(bold=True, font_color="white", bg_color=DARK, border=1,
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
    f_text_red = fmt(bg_color="#FFC7CE")
    f_date_red = fmt(bg_color="#FFC7CE", num_format="yyyy-mm-dd")
    f_int_red = fmt(bg_color="#FFC7CE", num_format="0", align="center")
    # żółte wiersze = pojazdy w rejestracji (konwencja z pliku źródłowego)
    f_text_y = fmt(bg_color="#FFF59D")
    f_date_y = fmt(bg_color="#FFF59D", num_format="yyyy-mm-dd")
    f_int_y = fmt(bg_color="#FFF59D", num_format="0", align="center")
    f_import = fmt(bg_color="#FFFDE7", border=1, border_color="#E0D9A0")
    f_import_date = fmt(bg_color="#FFFDE7", border=1, border_color="#E0D9A0",
                        num_format="yyyy-mm-dd")
    f_bandrow = fmt(bg_color=BAND)

    def header_band(ws, title, ncols):
        ws.set_row(0, 40)
        ws.merge_range(0, 0, 0, ncols - 2, title, f_band)
        ws.write(0, ncols - 1, "99rent", f_band_sub)
        ws.insert_image(0, 0, logo_small,
                        {"x_scale": 0.28, "y_scale": 0.28,
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
    ws.insert_image("B3", logo, {"x_scale": 0.105, "y_scale": 0.105,
                                 "object_position": 3})

    kpis = [
        ("=SUMPRODUCT(--(('W rejestracji'!$A$%d:$A$%d&'W rejestracji'!$C$%d:$C$%d)<>\"\"))"
         % (DATA_ROW, LAST, DATA_ROW, LAST), "POJAZDY\nW REJESTRACJI"),
        ("=SUMPRODUCT(--((Zarejestrowane!$A$%d:$A$%d&Zarejestrowane!$C$%d:$C$%d)<>\"\"))"
         % (DATA_ROW, LAST, DATA_ROW, LAST), "POJAZDY\nZAREJESTROWANE"),
        ("=IF(COUNT(Zarejestrowane!$I$%d:$I$%d)=0,\"—\",ROUND(AVERAGE(Zarejestrowane!$I$%d:$I$%d),1))"
         % (DATA_ROW, LAST, DATA_ROW, LAST), "ŚREDNI CZAS\nREJESTRACJI (DNI)"),
        ("=COUNTIFS(Zarejestrowane!$H$%d:$H$%d,\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))"
         % (DATA_ROW, LAST), "ZAREJESTROWANE\nW TYM MIESIĄCU"),
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
                ("NADAJ NR REJ.", "IdzNrRej"),
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
        "3.  IMPORT HURTOWY — wklejasz 10, 20, 30… pojazdów naraz (z VIN-ami) i jednym kliknięciem dodajesz wszystkie do rejestru.\n"
        "4.  Po odebraniu rejestracji: zaznacz pojazdy w W REJESTRACJI i kliknij ZAREJESTRUJ ZAZNACZONE — przechodzą do katalogu "
        "ZAREJESTROWANE z licznikiem dni. Numer rejestracyjny nadasz potem w zakładce NR REJESTRACYJNY (działa też dla pojazdów już w katalogu).\n"
        "5.  ZAREJESTROWANE — pojazd zarejestrowany wcześniej (poza rejestrem) dodasz bezpośrednio przyciskiem DODAJ DO KATALOGU.\n"
        "6.  PODSUMOWANIE — statystyki wg urzędu, dealera i marki oraz czasy rejestracji liczą się automatycznie.\n"
        "Żółte pola = pola do wypełnienia.  Duplikaty VIN są blokowane przez przyciski i podświetlane na czerwono w tabelach."
    )
    ws.merge_range(9, 1, 17, 10, instr, f_instr)
    ws.write(19, 1, "Wygenerowano na podstawie pliku Raport_rejestracji.xlsx", f_note)

    # ========================================================= W REJESTRACJI
    ws = wb.add_worksheet("W rejestracji")
    sheet_order.append(ws)
    ws.set_tab_color("#F4A100")
    headers = ["Marka", "Model", "VIN", "Dealer", "Współwłaściciel", "Urząd",
               "Data złożenia", "Dni od złożenia", "Uwagi"]
    widths = [13, 15, 21, 15, 15, 14, 13, 13, 30]
    colfmts = [f_text, f_text, f_text, f_text, f_text, f_text, f_date, f_int, f_text]
    for c, (w, cf) in enumerate(zip(widths, colfmts)):
        ws.set_column(c, c, w, cf)
    ws.set_column(9, 9, 2)
    ws.set_column(10, 10, 22)
    header_band(ws, "  POJAZDY W TRAKCIE REJESTRACJI", 9)
    nav_button(ws, 10)

    ws.merge_range(2, 0, 2, 8,
                   "FORMULARZ — NOWY WNIOSEK:  wypełnij żółte pola i kliknij DODAJ WNIOSEK",
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
        ws.insert_button(3, 10, {"macro": "DodajWniosek",
                                 "caption": "DODAJ WNIOSEK",
                                 "width": 150, "height": 40})
        ws.insert_button(6, 10, {"macro": "ZarejestrujZaznaczone",
                                 "caption": "ZAREJESTRUJ ZAZNACZONE ▶",
                                 "width": 170, "height": 34})
        ws.write(9, 10, "Zaznacz wiersze pojazdów i kliknij, aby przenieść "
                 "je do katalogu ZAREJESTROWANE (nr rej. nadasz później).",
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
    iw = [13, 18, 21, 15, 15, 14, 13, 30]
    for c, w in enumerate(iw):
        ws.set_column(c, c, w)
    ws.set_column(8, 8, 2)
    ws.set_column(9, 9, 26)
    header_band(ws, "  IMPORT HURTOWY — WIELE POJAZDÓW NARAZ", 8)
    nav_button(ws, 9)

    ws.merge_range(2, 0, 2, 7,
                   "Wklej 10, 20, 30… pojazdów do tabeli poniżej (od wiersza 9) "
                   "i kliknij IMPORTUJ DO REJESTRU", f_form_title)
    ws.merge_range(3, 0, 5, 7,
                   "Kolumny jak w arkuszu W rejestracji: wymagany jest VIN (kolumna C). "
                   "Pusta data złożenia = dzisiejsza data. Pojazdy, których VIN już "
                   "istnieje w rejestrze lub katalogu, zostaną pominięte — po imporcie "
                   "zobaczysz podsumowanie. Zaimportowane wiersze trafiają do arkusza "
                   "W REJESTRACJI z żółtym oznaczeniem, a tabela importu jest czyszczona.",
                   f_instr)
    if with_vba:
        ws.insert_button(2, 9, {"macro": "DodajHurtowo",
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
    zw = [13, 15, 21, 16, 15, 14, 13, 14, 14, 28]
    zfmts = [f_text, f_text, f_text, f_text, f_text, f_text, f_date, f_date,
             f_int, f_text]
    for c, (w, cf) in enumerate(zip(zw, zfmts)):
        ws.set_column(c, c, w, cf)
    ws.set_column(10, 10, 2)
    ws.set_column(11, 11, 24)
    header_band(ws, "  KATALOG POJAZDÓW ZAREJESTROWANYCH", 10)
    nav_button(ws, 11)

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
                ws.write_string(r, c, v, f_text)
        if dzl is not None:
            ws.write_datetime(r, 6, dzl, f_date)
        if datarej is not None:
            ws.write_datetime(r, 7, datarej, f_date)
        ws.write_formula(
            r, 8, '=IF(OR($G%d="",$H%d=""),"",$H%d-$G%d)'
            % (r + 1, r + 1, r + 1, r + 1), f_int)
        r += 1
    zlast = r

    ws.autofilter(HDR_ROW - 1, 0, LAST - 1, 9)
    ws.freeze_panes(HDR_ROW, 0)
    ws.conditional_format(DATA_ROW - 1, 2, LAST - 1, 2,
                          {"type": "duplicate", "format": f_red})
    ws.conditional_format(DATA_ROW - 1, 0, zlast - 1, 9,
                          {"type": "formula",
                           "criteria": "=MOD(ROW(),2)=0", "format": f_bandrow})
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

    # ======================================================= NR REJESTRACYJNY
    ws = wb.add_worksheet("Nr rejestracyjny")
    sheet_order.append(ws)
    ws.set_tab_color("#1565C0")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 24)
    ws.set_column("C:C", 26)
    ws.set_column("D:D", 30)
    ws.set_column("E:E", 4)
    ws.set_column("F:F", 26)
    ws.set_column("G:H", 14)
    header_band(ws, "  NADAWANIE NUMERU REJESTRACYJNEGO", 8)
    nav_button(ws, 6)

    ws.write(3, 1, "Wypełnij pola i kliknij ZAREJESTRUJ POJAZD:", f_form_title)
    ws.write(4, 1, "VIN pojazdu", f_lbl)
    ws.write_blank(4, 2, None, f_input)
    ws.write(5, 1, "Numer rejestracyjny", f_lbl)
    ws.write_blank(5, 2, None, f_input)
    ws.write(6, 1, "Data rejestracji", f_lbl)
    ws.write_blank(6, 2, None, f_input_date)
    ws.write(6, 3, "(puste pole = dzisiejsza data)", f_note)
    for rr in (4, 5, 6):
        ws.set_row(rr, 22)
    ws.data_validation(4, 2, 4, 2,
                       {"validate": "list",
                        "source": "='W rejestracji'!$C$%d:$C$%d" % (DATA_ROW, LAST),
                        "show_error": False})
    ws.data_validation(6, 2, 6, 2,
                       {"validate": "list", "source": "=Listy!$E$2:$E$32",
                        "show_error": False, "show_input": True,
                        "input_title": "Data rejestracji",
                        "input_message": "Wybierz z listy (31 dni wstecz) "
                                         "albo wpisz RRRR-MM-DD."})

    # --- wyszukiwarka VIN (fragment numeru) --------------------------------
    nrows = LAST - DATA_ROW + 1
    ws.write(3, 5, "WYSZUKAJ VIN — wpisz fragment numeru:", f_form_title)
    ws.write_blank(4, 5, None, f_input)
    ws.write(5, 5, "wyniki (W rejestracji, potem katalog):", f_note)
    for k in range(1, 6):
        for cc, listcol, cntcol in ((0, "F", "H"), (1, "G", "I")):
            ws.write_formula(
                5 + k, 5 + cc,
                '=IF($F$5="","",IFERROR(INDEX(Listy!$%(c)s$2:$%(c)s$%(n)d,'
                'MATCH(%(k)d,Listy!$%(h)s$2:$%(h)s$%(n)d,0)),""))'
                % {"c": listcol, "h": cntcol, "n": nrows + 1, "k": k},
                f_text)
    ws.write(5, 6, "(katalog)", f_note)
    if with_vba:
        ws.insert_button(8, 2, {"macro": "ZarejestrujPojazd",
                                "caption": "ZAREJESTRUJ POJAZD",
                                "width": 190, "height": 44})
    ws.set_row(8, 40)

    ws.write(11, 1, "PODGLĄD POJAZDU (dla wpisanego VIN)", f_sec)
    prev = [
        ("Marka", "$A", "$A"), ("Model", "$B", "$B"),
        ("Dealer", "$D", "$E"), ("Urząd", "$F", "$F"),
    ]
    rr = 12
    for label, colw, colz in prev:
        ws.write(rr, 1, label, f_lbl)
        ws.write_formula(
            rr, 2,
            '=IF($C$5="","",IFERROR(INDEX(\'W rejestracji\'!%s$%d:%s$%d,'
            'MATCH($C$5,\'W rejestracji\'!$C$%d:$C$%d,0)),'
            'IFERROR(INDEX(Zarejestrowane!%s$%d:%s$%d,'
            'MATCH($C$5,Zarejestrowane!$C$%d:$C$%d,0)),"nie znaleziono")))'
            % (colw, DATA_ROW, colw, LAST, DATA_ROW, LAST,
               colz, DATA_ROW, colz, LAST, DATA_ROW, LAST), f_text)
        rr += 1
    ws.write(rr, 1, "Data złożenia wniosku", f_lbl)
    ws.write_formula(
        rr, 2,
        '=IF($C$5="","",IFERROR(INDEX(\'W rejestracji\'!$G$%d:$G$%d,'
        'MATCH($C$5,\'W rejestracji\'!$C$%d:$C$%d,0)),'
        'IFERROR(INDEX(Zarejestrowane!$G$%d:$G$%d,'
        'MATCH($C$5,Zarejestrowane!$C$%d:$C$%d,0)),"")))'
        % (DATA_ROW, LAST, DATA_ROW, LAST, DATA_ROW, LAST, DATA_ROW, LAST),
        f_date)
    rr += 1
    ws.write(rr, 1, "Dni od złożenia (dziś)", f_lbl)
    ws.write_formula(
        rr, 2,
        '=IF($C$5="","",IFERROR(TODAY()-INDEX(\'W rejestracji\'!$G$%d:$G$%d,'
        'MATCH($C$5,\'W rejestracji\'!$C$%d:$C$%d,0)),""))'
        % (DATA_ROW, LAST, DATA_ROW, LAST), f_int)
    rr += 1
    ws.write(rr, 1, "Nr rejestracyjny (obecny)", f_lbl)
    ws.write_formula(
        rr, 2,
        '=IF($C$5="","",IFERROR(INDEX(Zarejestrowane!$D$%d:$D$%d,'
        'MATCH($C$5,Zarejestrowane!$C$%d:$C$%d,0)),"—"))'
        % (DATA_ROW, LAST, DATA_ROW, LAST), f_text)
    rr += 2
    ws.merge_range(rr, 1, rr + 4, 3,
                   "Po kliknięciu ZAREJESTRUJ POJAZD pojazd zostaje przeniesiony "
                   "z arkusza W REJESTRACJI do katalogu ZAREJESTROWANE, a licznik "
                   "dni od złożenia wniosku do rejestracji zapisuje się w kolumnie "
                   "„Czas rejestracji (dni)”. Jeśli pojazd jest już w katalogu "
                   "(np. przeniesiony przyciskiem ZAREJESTRUJ ZAZNACZONE), makro "
                   "tylko nadaje mu wpisany numer rejestracyjny.", f_note)

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
    nav_button(ws, 14)
    ws.hide_gridlines(2)

    ws.merge_range(2, 1, 2, 2, "  STATYSTYKI", f_sec_band)
    ws.set_row(2, 22)
    stats = [
        ("Pojazdy w rejestracji (oczekujące)",
         "=SUMPRODUCT(--(('W rejestracji'!$A$%d:$A$%d&'W rejestracji'!$C$%d:$C$%d)<>\"\"))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
        ("Pojazdy zarejestrowane (katalog)",
         "=SUMPRODUCT(--((Zarejestrowane!$A$%d:$A$%d&Zarejestrowane!$C$%d:$C$%d)<>\"\"))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
        ("Średni czas rejestracji (dni)",
         "=IF(COUNT(Zarejestrowane!$I$%d:$I$%d)=0,\"—\",ROUND(AVERAGE(Zarejestrowane!$I$%d:$I$%d),1))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
        ("Maksymalny czas rejestracji (dni)",
         "=IF(COUNT(Zarejestrowane!$I$%d:$I$%d)=0,\"—\",MAX(Zarejestrowane!$I$%d:$I$%d))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
        ("Minimalny czas rejestracji (dni)",
         "=IF(COUNT(Zarejestrowane!$I$%d:$I$%d)=0,\"—\",MIN(Zarejestrowane!$I$%d:$I$%d))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
        ("Zarejestrowane w bieżącym miesiącu",
         "=COUNTIFS(Zarejestrowane!$H$%d:$H$%d,\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))"
         % (DATA_ROW, LAST)),
        ("Najdłużej oczekujący (dni od złożenia)",
         "=IF(COUNT('W rejestracji'!$H$%d:$H$%d)=0,\"—\",MAX('W rejestracji'!$H$%d:$H$%d))"
         % (DATA_ROW, LAST, DATA_ROW, LAST)),
    ]
    for i, (label, formula) in enumerate(stats):
        ws.write(3 + i, 1, label, f_tbl_text)
        ws.write_formula(3 + i, 2, formula, f_val_box)
    stat_wrej_cell = "$C$4"
    stat_zar_cell = "$C$5"

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
                "=COUNTIF(Zarejestrowane!$%s$%d:$%s$%d,%s%d)"
                % (zcol, DATA_ROW, zcol, LAST,
                   xl_col_to_name(col0), rr + 1), f_tbl_int)
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
    # F/G: lustrzane listy VIN z obu arkuszy (do wyszukiwarki VIN)
    ws.write(0, 5, "VIN w rejestracji", f_hdr)
    ws.write(0, 6, "VIN zarejestrowane", f_hdr)
    # H/I: liczniki dopasowań do wyszukiwarki VIN (fragment w 'Nr rejestracyjny'!F5)
    ws.write(0, 7, "trafienia W rej", f_hdr)
    ws.write(0, 8, "trafienia katalog", f_hdr)
    for i in range(DATA_ROW, LAST + 1):
        rr = i - DATA_ROW + 2  # 1-indexed sheet row
        ws.write_formula(rr - 1, 5,
                         "=IF('W rejestracji'!$C$%d=\"\",\"\",'W rejestracji'!$C$%d)"
                         % (i, i), f_text)
        ws.write_formula(rr - 1, 6,
                         "=IF(Zarejestrowane!$C$%d=\"\",\"\",Zarejestrowane!$C$%d)"
                         % (i, i), f_text)
        for src, dst in (("F", "H"), ("G", "I")):
            ws.write_formula(
                rr - 1, ord(dst) - ord("A"),
                '=IF(AND($%(s)s%(r)d<>"",ISNUMBER(SEARCH('
                "'Nr rejestracyjny'!$F$5,$%(s)s%(r)d))),"
                "COUNT($%(d)s$1:%(d)s%(p)d)+1,\"\")"
                % {"s": src, "d": dst, "r": rr, "p": rr - 1}, f_int)
    ws.set_column(5, 8, 22)
    ws.hide()

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
