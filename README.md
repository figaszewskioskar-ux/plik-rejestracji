# Raport rejestracji pojazdów — 99rent

Rozbudowany plik rejestracji pojazdów z makrami (przyciski), licznikiem dni
i logo 99rent.

## Pliki

| Plik | Opis |
|---|---|
| `Raport_rejestracji_99rent.xlsm` | **Główny plik roboczy** — otwórz w Excelu i włącz makra |
| `Raport_rejestracji.xlsx` | Oryginalny plik źródłowy (bez zmian, dla porównania) |
| `narzedzia/` | Skrypty Pythona, którymi wygenerowano plik (build, makra VBA, logo) |

## Arkusze

1. **PULPIT** — logo 99rent, wskaźniki (KPI), przyciski nawigacji i instrukcja.
2. **W rejestracji** — pojazdy oczekujące — wiersze żółte (złożone, 60 szt.) i czerwone (7 szt., zachowane czerwone podświetlenie) z pliku źródłowego.
   Formularz w wierszu 5 + przycisk **DODAJ WNIOSEK**. Kolumna
   *Dni od złożenia* liczy się automatycznie (`=DZIŚ()-data złożenia`)
   i podświetla zaległości (pomarańczowy > 10 dni, czerwony > 21 dni).
3. **Import hurtowy** — wklejasz wiele pojazdów naraz (10, 20, 30…)
   i przycisk **IMPORTUJ DO REJESTRU** dodaje je hurtem do rejestru
   (duplikaty VIN są pomijane).
4. **Zarejestrowane** — nowy katalog pojazdów zarejestrowanych — wiersze zielone
   (odebrane, 867 szt.); 17 z nich ma nr rejestracyjny i datę z arkusza
   źródłowego Arkusz2 (połączone po VIN). Formularz + przycisk
   **DODAJ DO KATALOGU** dla pojazdów zarejestrowanych wcześniej.
   Kolumna *Czas rejestracji (dni)* = data rejestracji − data złożenia wniosku.
   Przycisk **ZAREJESTRUJ ZAZNACZONE** w W rejestracji przenosi wiele
   pojazdów naraz. Daty wybiera się szybko z listy (31 dni wstecz).
4. **Nr rejestracyjny** — zakładka nadawania numeru: wyszukiwarka VIN (fragment numeru), wybierasz VIN z listy,
   wpisujesz numer i datę, klikasz **ZAREJESTRUJ POJAZD** — pojazd przenosi
   się z *W rejestracji* do katalogu *Zarejestrowane* wraz z licznikiem dni.
   Panel podglądu pokazuje dane pojazdu dla wpisanego VIN.
5. **PODSUMOWANIE** — statystyki na żywo: liczby pojazdów, średni / min / maks
   czas rejestracji, zarejestrowane w bieżącym miesiącu oraz zestawienia
   wg urzędu, marki i dealera.
6. **Pilne** — lista priorytetowa z pierwotnego raportu.
7. **Listy** (ukryty) — słowniki do list rozwijanych (marki, dealerzy,
   urzędy, współwłaściciele).

## Zabezpieczenia

- Przyciski blokują duplikaty VIN (w obu arkuszach), formatowanie warunkowe
  podświetla duplikaty na czerwono.
- Żółte pola = pola do wypełnienia; listy rozwijane podpowiadają wartości,
  ale nie blokują wpisania nowych.

## Włączanie makr

Przy pierwszym otwarciu kliknij **Włącz zawartość**. Jeśli Excel blokuje
makra (plik pobrany z internetu): zamknij plik → kliknij prawym przyciskiem →
*Właściwości* → zaznacz *Odblokuj* → OK → otwórz ponownie.

## Regeneracja pliku

```bash
pip install openpyxl xlsxwriter ms-ovba-compression oletools pillow
cd narzedzia
python logo_99rent.py
python make_vba_project.py            # generuje vbaProject.bin z makr VBA
cp ../Raport_rejestracji.xlsx .       # dane źródłowe
python build_workbook.py ../Raport_rejestracji_99rent.xlsm
```
