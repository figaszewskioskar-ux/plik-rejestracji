# -*- coding: utf-8 -*-
"""Generuje vbaProject.bin (makra przycisków) i weryfikuje go oletools/olevba."""
from make_vba_bin import make_vba_project
from vba_source import THISWORKBOOK, MODULE1, SHEET_MODULE_TEMPLATE

N_SHEETS = 12  # PULPIT, W rejestracji, Import hurtowy, Zarejestrowane, PODSUMOWANIE,
               # Do rejestracji, URZEDY, PRZELEWY, Listy + Archiwum 2026-05/06/07

def main(out="vbaProject.bin"):
    modules = {"ThisWorkbook": THISWORKBOOK}
    doc = {"ThisWorkbook"}
    for i in range(1, N_SHEETS + 1):
        name = "Arkusz%d" % i
        modules[name] = SHEET_MODULE_TEMPLATE % name
        doc.add(name)
    modules["Module1"] = MODULE1
    data = make_vba_project(modules, doc)
    with open(out, "wb") as f:
        f.write(data)

    # weryfikacja: kontener OLE + ekstrakcja makr
    import olefile
    ole = olefile.OleFileIO(out)
    streams = ["/".join(p) for p in ole.listdir()]
    ole.close()
    assert "VBA/Module1" in streams and "VBA/dir" in streams, streams
    from oletools.olevba import VBA_Parser
    p = VBA_Parser(out)
    assert p.detect_vba_macros()
    found = {}
    for (_, _, vba_fn, code) in p.extract_macros():
        found[vba_fn] = code
    p.close()
    for sub in ("DodajWniosek", "DodajZarejestrowany", "PrzeniesWnioski",
                "IdzPulpit"):
        assert "Sub " + sub in found["Module1.bas"], sub
    print("vbaProject.bin OK —", len(data), "bajtów,", len(found), "modułów")

if __name__ == "__main__":
    main()
