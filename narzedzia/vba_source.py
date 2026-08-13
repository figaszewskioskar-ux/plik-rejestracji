# -*- coding: utf-8 -*-
"""VBA source for the 99rent registration workbook. ASCII-only (CP1252-safe)."""

THISWORKBOOK = '''Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    ' Wymus automatyczne przeliczanie i pelna kalkulacje przy otwarciu
    ' (gdy sesja Excela byla w trybie recznym, formuly pokazywaly 0 / 1900-01).
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    ' ochrona formul i naglowkow (makra moga pisac, uzytkownik nie)
    On Error Resume Next
    ThisWorkbook.Worksheets("PULPIT").Protect UserInterfaceOnly:=True
    ThisWorkbook.Worksheets("PODSUMOWANIE").Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Module1.StartZegar
    Module1.SprawdzArchiwizacje
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Module1.StopZegar
End Sub
'''

MODULE1 = r'''Attribute VB_Name = "Module1"
Option Explicit

' =====================================================================
'  99rent - Raport rejestracji pojazdow
'  Makra obslugujace przyciski w skoroszycie.
'  Uklad arkuszy:
'    "W rejestracji"   - formularz w wierszu 5, tabela od wiersza 8
'    "Zarejestrowane"  - formularz w wierszu 5, tabela od wiersza 8
' =====================================================================

Private Const ROW_FORM As Long = 5      ' wiersz formularza
Private Const ROW_HDR As Long = 8       ' wiersz naglowka tabel

' zegar na PULPICIE (deklaracje musza byc na gorze modulu)
Private nextTick As Date
Private tickArmed As Boolean

Private Function LastRow(ws As Worksheet) As Long
    ' Ostatni uzyty wiersz w kolumnach A:J (niektore pojazdy nie maja
    ' jeszcze VIN, wiec nie mozna polegac na jednej kolumnie).
    Dim i As Long, r As Long, m As Long
    m = ROW_HDR
    For i = 1 To 10
        r = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
        If r > m Then m = r
    Next i
    LastRow = m
End Function

Private Function FindVinRow(ws As Worksheet, vin As String) As Long
    ' Szybkie wyszukiwanie VIN (Match zamiast petli po komorkach).
    Dim wynik As Variant
    wynik = Application.Match(Trim(vin), _
        ws.Range(ws.Cells(ROW_HDR + 1, 3), ws.Cells(LastRow(ws), 3)), 0)
    If IsError(wynik) Then
        FindVinRow = 0
    Else
        FindVinRow = ROW_HDR + CLng(wynik)
    End If
End Function

Private Function SafeDate(v As Variant, fallback As Date) As Date
    If IsDate(v) Then
        SafeDate = CDate(v)
    Else
        SafeDate = fallback
    End If
End Function

Private Sub PaintRow(ws As Worksheet, r As Long, lastCol As Long, clr As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
    rng.Interior.Color = clr
    rng.Borders.Color = RGB(158, 158, 158)
    rng.Borders.Weight = xlThin
End Sub

Private Sub PaintYellow(ws As Worksheet, r As Long)
    ' Wiersz oczekujacy na rejestracje = zolty (konwencja z pliku zrodlowego)
    PaintRow ws, r, 11, RGB(255, 249, 196)
End Sub

Private Sub PaintGreen(ws As Worksheet, r As Long)
    ' Wiersz zarejestrowany = zielony (konwencja z pliku zrodlowego)
    PaintRow ws, r, 10, RGB(198, 239, 206)
End Sub

' ---------------------------------------------------------------------
' Przycisk: DODAJ WNIOSEK (arkusz "W rejestracji")
' Dodaje nowy pojazd do rejestru pojazdow oczekujacych na rejestracje.
' ---------------------------------------------------------------------
Sub DodajWniosek()
    Dim ws As Worksheet, r As Long, vin As String
    Set ws = ThisWorkbook.Worksheets("W rejestracji")

    vin = Trim(CStr(ws.Cells(ROW_FORM, 3).Value))
    If vin = "" Then
        MsgBox "Podaj numer VIN pojazdu.", vbExclamation, "99rent"
        Exit Sub
    End If
    If FindVinRow(ws, vin) > 0 Then
        MsgBox "Pojazd o VIN " & vin & " jest juz w rejestracji.", vbExclamation, "99rent"
        Exit Sub
    End If
    If FindVinRow(ThisWorkbook.Worksheets("Zarejestrowane"), vin) > 0 Then
        MsgBox "Pojazd o VIN " & vin & " jest juz zarejestrowany.", vbExclamation, "99rent"
        Exit Sub
    End If

    r = LastRow(ws) + 1
    ws.Cells(r, 1).Value = ws.Cells(ROW_FORM, 1).Value          ' Marka
    ws.Cells(r, 2).Value = ws.Cells(ROW_FORM, 2).Value          ' Model
    ws.Cells(r, 3).Value = vin                                   ' VIN
    ws.Cells(r, 4).Value = ws.Cells(ROW_FORM, 4).Value          ' Dealer
    ws.Cells(r, 5).Value = ws.Cells(ROW_FORM, 5).Value          ' Wspolwlasciciel
    ws.Cells(r, 6).Value = ws.Cells(ROW_FORM, 6).Value          ' Urzad
    ws.Cells(r, 7).Value = SafeDate(ws.Cells(ROW_FORM, 7).Value, Date)
    ws.Cells(r, 7).NumberFormat = "yyyy-mm-dd"
    ws.Cells(r, 8).Formula = "=IF($G" & r & "=" & Chr(34) & Chr(34) & _
        "," & Chr(34) & Chr(34) & ",TODAY()-$G" & r & ")"
    If IsDate(ws.Cells(ROW_FORM, 9).Value) Then                  ' Plan. odbior
        ws.Cells(r, 9).Value = CDate(ws.Cells(ROW_FORM, 9).Value)
        ws.Cells(r, 9).NumberFormat = "yyyy-mm-dd"
    End If
    ws.Cells(r, 10).Value = ws.Cells(ROW_FORM, 10).Value        ' Osoba
    ws.Cells(r, 11).Value = ws.Cells(ROW_FORM, 11).Value        ' Uwagi
    PaintYellow ws, r

    ws.Range(ws.Cells(ROW_FORM, 1), ws.Cells(ROW_FORM, 7)).ClearContents
    ws.Range(ws.Cells(ROW_FORM, 9), ws.Cells(ROW_FORM, 11)).ClearContents
    MsgBox "Wniosek dodany do rejestru (wiersz " & r & ").", vbInformation, "99rent"
End Sub

' ---------------------------------------------------------------------
' Przycisk: DODAJ DO KATALOGU (arkusz "Zarejestrowane")
' Dodaje pojazd JUZ ZAREJESTROWANY bezposrednio do katalogu.
' ---------------------------------------------------------------------
Sub DodajZarejestrowany()
    Dim ws As Worksheet, r As Long, vin As String
    Set ws = ThisWorkbook.Worksheets("Zarejestrowane")

    vin = Trim(CStr(ws.Cells(ROW_FORM, 3).Value))
    If vin = "" Then
        MsgBox "Podaj numer VIN pojazdu.", vbExclamation, "99rent"
        Exit Sub
    End If
    If Trim(CStr(ws.Cells(ROW_FORM, 4).Value)) = "" Then
        MsgBox "Podaj numer rejestracyjny pojazdu.", vbExclamation, "99rent"
        Exit Sub
    End If
    If FindVinRow(ws, vin) > 0 Then
        MsgBox "Pojazd o VIN " & vin & " jest juz w katalogu.", vbExclamation, "99rent"
        Exit Sub
    End If

    r = LastRow(ws) + 1
    ws.Cells(r, 1).Value = ws.Cells(ROW_FORM, 1).Value          ' Marka
    ws.Cells(r, 2).Value = ws.Cells(ROW_FORM, 2).Value          ' Model
    ws.Cells(r, 3).Value = vin                                   ' VIN
    ws.Cells(r, 4).Value = ws.Cells(ROW_FORM, 4).Value          ' Nr rej
    ws.Cells(r, 5).Value = ws.Cells(ROW_FORM, 5).Value          ' Dealer
    ws.Cells(r, 6).Value = ws.Cells(ROW_FORM, 6).Value          ' Urzad
    If IsDate(ws.Cells(ROW_FORM, 7).Value) Then
        ws.Cells(r, 7).Value = CDate(ws.Cells(ROW_FORM, 7).Value)
        ws.Cells(r, 7).NumberFormat = "yyyy-mm-dd"
    End If
    ws.Cells(r, 8).Value = SafeDate(ws.Cells(ROW_FORM, 8).Value, Date)
    ws.Cells(r, 8).NumberFormat = "yyyy-mm-dd"
    ws.Cells(r, 9).Formula = "=IF(OR($G" & r & "=" & Chr(34) & Chr(34) & _
        ",$H" & r & "=" & Chr(34) & Chr(34) & _
        ",$H" & r & "<$G" & r & ")," & Chr(34) & Chr(34) & _
        ",$H" & r & "-$G" & r & ")"
    ws.Cells(r, 10).Value = ws.Cells(ROW_FORM, 10).Value        ' Uwagi
    PaintGreen ws, r

    ws.Range(ws.Cells(ROW_FORM, 1), ws.Cells(ROW_FORM, 8)).ClearContents
    ws.Cells(ROW_FORM, 10).ClearContents
    MsgBox "Pojazd " & vin & " dodany do katalogu zarejestrowanych.", _
        vbInformation, "99rent"
End Sub

' ---------------------------------------------------------------------
' Przycisk: PRZYWROC DO REJESTRACJI (arkusz "Zarejestrowane")
' Zaznaczone pojazdy wracaja z katalogu do arkusza W rejestracji
' (np. omylkowo zarejestrowane). Nr rej i data rejestracji sa usuwane.
' ---------------------------------------------------------------------
Sub PrzywrocZaznaczone()
    Dim krok As String
    Dim wsW As Worksheet, wsZ As Worksheet
    Dim obszar As Range, wiersz As Range, rowsToMove As Object, key As Variant
    Dim rw As Long, rz As Long, n As Long
    Dim vin As String, arr() As Long, i As Long, j As Long, tmp As Long

    On Error GoTo Blad
    krok = "start"
    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")

    If ActiveSheet.Name <> wsZ.Name Then
        MsgBox "Przejdz do arkusza Zarejestrowane i zaznacz wiersze pojazdow.", _
            vbExclamation, "99rent"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "Najpierw zaznacz w tabeli wiersze pojazdow do przywrocenia.", _
            vbExclamation, "99rent"
        Exit Sub
    End If

    krok = "zbieranie wierszy"
    Set rowsToMove = CreateObject("Scripting.Dictionary")
    For Each obszar In Selection.Areas
        For Each wiersz In obszar.Rows
            rw = wiersz.Row
            If rw > ROW_HDR And rw <= LastRow(wsZ) Then
                vin = Trim(CStr(wsZ.Cells(rw, 3).Value))
                If vin <> "" And Not rowsToMove.Exists(rw) Then
                    rowsToMove.Add rw, vin
                End If
            End If
        Next wiersz
    Next obszar

    If rowsToMove.Count = 0 Then
        MsgBox "Zaznacz co najmniej jeden wiersz pojazdu (z VIN) w tabeli.", _
            vbExclamation, "99rent"
        Exit Sub
    End If

    StopZegar
    If MsgBox("Przywrocic " & rowsToMove.Count & " pojazd(y) do arkusza " & _
        "W REJESTRACJI?" & vbCrLf & "Nr rejestracyjny i data rejestracji " & _
        "zostana usuniete.", vbYesNo + vbQuestion, "99rent") <> vbYes Then
        StartZegar
        Exit Sub
    End If

    ReDim arr(0 To rowsToMove.Count - 1)
    i = 0
    For Each key In rowsToMove.Keys
        arr(i) = CLng(key)
        i = i + 1
    Next key
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    Application.ScreenUpdating = False
    n = 0
    For i = 0 To UBound(arr)
        rw = arr(i)
        krok = "przywracanie wiersza " & rw
        vin = Trim(CStr(wsZ.Cells(rw, 3).Value))
        If FindVinRow(wsW, vin) = 0 Then
            rz = LastRow(wsW) + 1
            wsW.Cells(rz, 1).Value = wsZ.Cells(rw, 1).Value      ' Marka
            wsW.Cells(rz, 2).Value = wsZ.Cells(rw, 2).Value      ' Model
            wsW.Cells(rz, 3).Value = vin                          ' VIN
            wsW.Cells(rz, 4).Value = wsZ.Cells(rw, 5).Value      ' Dealer
            wsW.Cells(rz, 6).Value = wsZ.Cells(rw, 6).Value      ' Urzad
            If IsDate(wsZ.Cells(rw, 7).Value) Then
                wsW.Cells(rz, 7).Value = CDate(wsZ.Cells(rw, 7).Value)
                wsW.Cells(rz, 7).NumberFormat = "yyyy-mm-dd"
            End If
            wsW.Cells(rz, 8).Formula = "=IF($G" & rz & "=" & Chr(34) & Chr(34) & _
                "," & Chr(34) & Chr(34) & ",TODAY()-$G" & rz & ")"
            wsW.Cells(rz, 11).Value = wsZ.Cells(rw, 10).Value    ' Uwagi
            PaintYellow wsW, rz
            n = n + 1
        End If
        wsZ.Rows(rw).Delete Shift:=xlUp
    Next i
    Application.ScreenUpdating = True

    MsgBox "Przywrocono do rejestracji: " & n & " pojazd(y).", _
        vbInformation, "99rent"
    StartZegar
    Exit Sub
Blad:
    Application.ScreenUpdating = True
    MsgBox "Blad przywracania (etap: " & krok & "):" & vbCrLf & _
        Err.Number & " - " & Err.Description, vbCritical, "99rent"
    StartZegar
End Sub

' ---------------------------------------------------------------------
' Przycisk: IMPORTUJ DO REJESTRU (arkusz "Do rejestracji")
' Wklejasz pojazdy w tabele wzoru, zaznaczasz wiersze (albo nic -
' wtedy bierze wszystkie) i klikasz: pojazdy przechodza do W rejestracji.
' Wymagane kolumny A-G oraz komplet dokumentow = TAK (kolumna H);
' braki zostaja podswietlone na czerwono i wiersz nie jest przenoszony.
' ---------------------------------------------------------------------
Sub PrzeniesWnioski()
    Dim krok As String
    Dim wsI As Worksheet, wsW As Worksheet, wsZ As Worksheet
    Dim obszar As Range, wiersz As Range, doPrzen As Object, key As Variant
    Dim i As Long, r As Long, rw As Long, n As Long, skipped As Long
    Dim lastI As Long, vin As String
    Dim arr() As Long, j As Long, tmp As Long

    On Error GoTo Blad
    krok = "start"
    Set wsI = ThisWorkbook.Worksheets("Do rejestracji")
    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")
    Set doPrzen = CreateObject("Scripting.Dictionary")
    lastI = LastRow(wsI)

    krok = "zbieranie wierszy"
    If ActiveSheet.Name = wsI.Name And TypeName(Selection) = "Range" Then
        For Each obszar In Selection.Areas
            For Each wiersz In obszar.Rows
                rw = wiersz.Row
                If rw > ROW_HDR And rw <= lastI Then
                    vin = Trim(CStr(wsI.Cells(rw, 3).Value))
                    If vin <> "" And Not doPrzen.Exists(rw) Then
                        doPrzen.Add rw, vin
                    End If
                End If
            Next wiersz
        Next obszar
    End If
    If doPrzen.Count = 0 Then
        ' brak zaznaczenia w tabeli -> wszystkie wnioski z VIN
        For i = ROW_HDR + 1 To lastI
            If Trim(CStr(wsI.Cells(i, 3).Value)) <> "" Then
                doPrzen.Add i, Trim(CStr(wsI.Cells(i, 3).Value))
            End If
        Next i
    End If

    If doPrzen.Count = 0 Then
        MsgBox "Wklej pojazdy w tabele od wiersza " & (ROW_HDR + 1) & _
            " (kolumna VIN jest wymagana), zaznacz wiersze i kliknij ponownie.", _
            vbExclamation, "99rent"
        Exit Sub
    End If

    StopZegar
    If MsgBox("Przeniesc " & doPrzen.Count & " wniosek/wnioski do arkusza " & _
        "W REJESTRACJI?", vbYesNo + vbQuestion, "99rent") <> vbYes Then
        StartZegar
        Exit Sub
    End If

    ReDim arr(0 To doPrzen.Count - 1)
    i = 0
    For Each key In doPrzen.Keys
        arr(i) = CLng(key)
        i = i + 1
    Next key
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    Application.ScreenUpdating = False
    Dim braki As Long, c2 As Long, pelny As Boolean
    braki = 0
    n = 0: skipped = 0
    For i = 0 To UBound(arr)
        rw = arr(i)
        krok = "sprawdzanie wiersza " & rw
        ' wymagane kolumny A-G oraz komplet dokumentow = TAK (kolumna H)
        pelny = True
        For c2 = 1 To 7
            If Trim(CStr(wsI.Cells(rw, c2).Value)) = "" Then
                pelny = False
                wsI.Cells(rw, c2).Interior.Color = RGB(255, 199, 206)
            End If
        Next c2
        If UCase(Trim(CStr(wsI.Cells(rw, 8).Value))) <> "TAK" Then
            pelny = False
            wsI.Cells(rw, 8).Interior.Color = RGB(255, 199, 206)
        End If
        If Not pelny Then
            braki = braki + 1
            GoTo NastepnyWiersz
        End If
        krok = "przenoszenie wiersza " & rw
        vin = Trim(CStr(wsI.Cells(rw, 3).Value))
        If FindVinRow(wsW, vin) > 0 Or FindVinRow(wsZ, vin) > 0 Then
            skipped = skipped + 1
        Else
            r = LastRow(wsW) + 1
            wsW.Cells(r, 1).Value = wsI.Cells(rw, 1).Value    ' Marka
            wsW.Cells(r, 2).Value = wsI.Cells(rw, 2).Value    ' Model
            wsW.Cells(r, 3).Value = vin                        ' VIN
            wsW.Cells(r, 4).Value = wsI.Cells(rw, 4).Value    ' Dealer
            wsW.Cells(r, 5).Value = wsI.Cells(rw, 5).Value    ' Wspolwl.
            wsW.Cells(r, 6).Value = wsI.Cells(rw, 6).Value    ' Urzad
            wsW.Cells(r, 7).Value = SafeDate(wsI.Cells(rw, 7).Value, Date)
            wsW.Cells(r, 7).NumberFormat = "yyyy-mm-dd"
            wsW.Cells(r, 8).Formula = "=IF($G" & r & "=" & Chr(34) & Chr(34) & _
                "," & Chr(34) & Chr(34) & ",TODAY()-$G" & r & ")"
            wsW.Cells(r, 11).Value = wsI.Cells(rw, 10).Value  ' Uwagi
            PaintYellow wsW, r
            n = n + 1
        End If
        wsI.Range(wsI.Cells(rw, 1), wsI.Cells(rw, 10)).Delete Shift:=xlUp
NastepnyWiersz:
    Next i
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Przeniesiono do rejestracji: " & n & " wniosek/wnioski."
    If skipped > 0 Then msg = msg & vbCrLf & _
        "Pominieto (VIN juz istnieje): " & skipped & "."
    If braki > 0 Then msg = msg & vbCrLf & _
        "NIE przeniesiono " & braki & " wiersza/y (podswietlone na " & _
        "czerwono) - wymagane kolumny A-G oraz komplet dokumentow = TAK."
    MsgBox msg, vbInformation, "99rent"
    StartZegar
    Exit Sub
Blad:
    Application.ScreenUpdating = True
    MsgBox "Blad przenoszenia wnioskow (etap: " & krok & "):" & vbCrLf & _
        Err.Number & " - " & Err.Description, vbCritical, "99rent"
    StartZegar
End Sub

' ---------------------------------------------------------------------
' Przycisk: ZAREJESTRUJ ZAZNACZONE (arkusz "W rejestracji")
' Zaznacz dowolne komorki wierszy pojazdow (mozna wiele naraz) i kliknij:
' pojazdy przechodza do katalogu "Zarejestrowane" z dzisiejsza data
' rejestracji. Numer rejestracyjny nadasz potem w zakladce Nr rejestracyjny.
' ---------------------------------------------------------------------
Sub ZarejestrujZaznaczone()
    Dim krok As String
    Dim wsW As Worksheet, wsZ As Worksheet
    Dim obszar As Range, wiersz As Range, rowsToMove As Object, key As Variant
    Dim rw As Long, rz As Long, n As Long
    Dim vin As String, dataZl As Variant, arr() As Long, i As Long, j As Long, tmp As Long

    On Error GoTo Blad
    krok = "start"
    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")

    krok = "sprawdzanie zaznaczenia"
    If ActiveSheet.Name <> wsW.Name Then
        MsgBox "Przejdz do arkusza W rejestracji i zaznacz wiersze pojazdow.", _
            vbExclamation, "99rent"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "Najpierw zaznacz w tabeli wiersze pojazdow do zarejestrowania.", _
            vbExclamation, "99rent"
        Exit Sub
    End If

    krok = "zbieranie wierszy"
    Set rowsToMove = CreateObject("Scripting.Dictionary")
    For Each obszar In Selection.Areas
        For Each wiersz In obszar.Rows
            rw = wiersz.Row
            If rw > ROW_HDR And rw <= LastRow(wsW) Then
                vin = Trim(CStr(wsW.Cells(rw, 3).Value))
                If vin <> "" And Not rowsToMove.Exists(rw) Then
                    rowsToMove.Add rw, vin
                End If
            End If
        Next wiersz
    Next obszar

    If rowsToMove.Count = 0 Then
        MsgBox "Zaznacz co najmniej jeden wiersz pojazdu (z VIN) w tabeli " & _
            "(wiersze od " & (ROW_HDR + 1) & " w dol).", vbExclamation, "99rent"
        Exit Sub
    End If

    StopZegar
    If MsgBox("Przeniesc " & rowsToMove.Count & " pojazd(y) do katalogu " & _
        "ZAREJESTROWANE z dzisiejsza data rejestracji?" & vbCrLf & _
        "Numer rejestracyjny wpiszesz w kolumnie D katalogu.", _
        vbYesNo + vbQuestion, "99rent") <> vbYes Then Exit Sub

    ' posortuj wiersze malejaco, zeby usuwanie nie przesuwalo numeracji
    ReDim arr(0 To rowsToMove.Count - 1)
    i = 0
    For Each key In rowsToMove.Keys
        arr(i) = CLng(key)
        i = i + 1
    Next key
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    Application.ScreenUpdating = False
    n = 0
    For i = 0 To UBound(arr)
        rw = arr(i)
        krok = "przenoszenie wiersza " & rw
        vin = Trim(CStr(wsW.Cells(rw, 3).Value))
        If FindVinRow(wsZ, vin) = 0 Then
            dataZl = wsW.Cells(rw, 7).Value
            rz = LastRow(wsZ) + 1
            wsZ.Cells(rz, 1).Value = wsW.Cells(rw, 1).Value      ' Marka
            wsZ.Cells(rz, 2).Value = wsW.Cells(rw, 2).Value      ' Model
            wsZ.Cells(rz, 3).Value = vin                          ' VIN
            wsZ.Cells(rz, 5).Value = wsW.Cells(rw, 4).Value      ' Dealer
            wsZ.Cells(rz, 6).Value = wsW.Cells(rw, 6).Value      ' Urzad
            If IsDate(dataZl) Then
                wsZ.Cells(rz, 7).Value = CDate(dataZl)
                wsZ.Cells(rz, 7).NumberFormat = "yyyy-mm-dd"
            End If
            wsZ.Cells(rz, 8).Value = Date                         ' Data rejestracji
            wsZ.Cells(rz, 8).NumberFormat = "yyyy-mm-dd"
            wsZ.Cells(rz, 9).Formula = "=IF(OR($G" & rz & "=" & Chr(34) & Chr(34) & _
                ",$H" & rz & "=" & Chr(34) & Chr(34) & _
                ",$H" & rz & "<$G" & rz & ")," & Chr(34) & Chr(34) & _
                ",$H" & rz & "-$G" & rz & ")"
            wsZ.Cells(rz, 10).Value = wsW.Cells(rw, 11).Value    ' Uwagi
            PaintGreen wsZ, rz
            n = n + 1
        End If
        wsW.Rows(rw).Delete Shift:=xlUp
    Next i
    Application.ScreenUpdating = True

    MsgBox "Przeniesiono do katalogu: " & n & " pojazd(y)." & vbCrLf & _
        "Numer rejestracyjny wpisz w kolumnie D katalogu.", _
        vbInformation, "99rent"
    StartZegar
    Exit Sub
Blad:
    Application.ScreenUpdating = True
    MsgBox "Blad podczas przenoszenia (etap: " & krok & "):" & vbCrLf & _
        Err.Number & " - " & Err.Description, vbCritical, "99rent"
End Sub


' ---------------------------------------------------------------------
' Zegar na PULPICIE: data i godzina z sekundami, odswiezany co 30 s
' (czestszy zapis czyscilby schowek i przerywal prace w arkuszach).
' ---------------------------------------------------------------------
Sub StartZegar()
    On Error Resume Next
    ThisWorkbook.Worksheets("PULPIT").Range("J2").Value = Now
    nextTick = Now + TimeSerial(0, 0, 30)
    Application.OnTime nextTick, "TykZegara"
    tickArmed = True
End Sub

Sub TykZegara()
    On Error Resume Next
    ThisWorkbook.Worksheets("PULPIT").Range("J2").Value = Now
    nextTick = Now + TimeSerial(0, 0, 30)
    Application.OnTime nextTick, "TykZegara"
    tickArmed = True
End Sub

Sub StopZegar()
    On Error Resume Next
    If tickArmed Then
        Application.OnTime nextTick, "TykZegara", , False
        tickArmed = False
    End If
End Sub

' ---------------------------------------------------------------------
' Archiwizacja: pojazdy zarejestrowane w POPRZEDNICH miesiacach
' przenosza sie do nowego pliku Archiwum_zarejestrowane_RRRR-MM.xlsx
' (w folderze tego skoroszytu), zeby katalog nie zasmiecal sie starymi
' wpisami. Uruchamia sie przyciskiem lub automatycznie przy pierwszym
' otwarciu w nowym miesiacu (z pytaniem).
' ---------------------------------------------------------------------
Sub SprawdzArchiwizacje()
    On Error Resume Next
    Dim wsL As Worksheet, znacznik As String
    Set wsL = ThisWorkbook.Worksheets("Listy")
    znacznik = Format(Date, "yyyy-mm")
    If CStr(wsL.Range("K1").Value) = znacznik Then Exit Sub
    wsL.Range("K1").Value = znacznik
    If LiczStareMiesiace() = 0 Then Exit Sub
    If MsgBox("Nowy miesiac. W katalogu ZAREJESTROWANE sa pojazdy " & _
        "zarejestrowane w poprzednich miesiacach (" & LiczStareMiesiace() & _
        " szt.)." & vbCrLf & "Przeniesc je teraz do pliku archiwum?", _
        vbYesNo + vbQuestion, "99rent") = vbYes Then
        ArchiwizujStareMiesiace
    End If
End Sub

Private Function LiczStareMiesiace() As Long
    Dim wsZ As Worksheet, i As Long, n As Long, d As Variant
    Dim progu As Date
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")
    progu = DateSerial(Year(Date), Month(Date), 1)
    For i = ROW_HDR + 1 To LastRow(wsZ)
        d = wsZ.Cells(i, 8).Value
        If IsDate(d) Then
            If CDate(d) < progu Then n = n + 1
        End If
    Next i
    LiczStareMiesiace = n
End Function

Sub ArchiwizujStareMiesiace()
    On Error GoTo Blad
    Dim wsZ As Worksheet, i As Long, d As Variant
    Dim progu As Date, klucz As Variant, razem As Long, opis As String
    Dim mies As Object
    Set mies = CreateObject("Scripting.Dictionary")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")
    progu = DateSerial(Year(Date), Month(Date), 1)

    For i = ROW_HDR + 1 To LastRow(wsZ)
        d = wsZ.Cells(i, 8).Value
        If IsDate(d) Then
            If CDate(d) < progu Then mies(Format(CDate(d), "yyyy-mm")) = 1
        End If
    Next i

    If mies.Count = 0 Then
        MsgBox "Brak pojazdow zarejestrowanych w poprzednich miesiacach.", _
            vbInformation, "99rent"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    razem = 0: opis = ""
    For Each klucz In mies.Keys
        razem = razem + ArchiwizujMiesiac(CStr(klucz))
        opis = opis & vbCrLf & "  Archiwum " & klucz
    Next klucz
    Application.ScreenUpdating = True

    MsgBox "Zarchiwizowano " & razem & " pojazd(y) do zakladek na dole pliku:" & _
        opis, vbInformation, "99rent"
    Exit Sub
Blad:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Blad archiwizacji: " & Err.Description, vbCritical, "99rent"
End Sub

Private Function SciezkaArchiwum() As String
    SciezkaArchiwum = ThisWorkbook.Path
    If SciezkaArchiwum = "" Then SciezkaArchiwum = Application.DefaultFilePath
End Function

Private Function ArchiwizujMiesiac(klucz As String) As Long
    ' Przenosi pojazdy zarejestrowane w miesiacu 'klucz' (RRRR-MM) do
    ' zakladki "Archiwum RRRR-MM" na koncu tego skoroszytu.
    ' Gdy zakladka istnieje - dopisuje na koncu.
    Dim wsZ As Worksheet, wsA As Worksheet
    Dim i As Long, r As Long, c As Long, n As Long, d As Variant
    Dim nazwa As String

    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")
    nazwa = "Archiwum " & klucz

    On Error Resume Next
    Set wsA = ThisWorkbook.Worksheets(nazwa)
    On Error GoTo 0

    If wsA Is Nothing Then
        Set wsA = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsA.Name = nazwa
        wsA.Tab.Color = RGB(120, 120, 120)
        For c = 1 To 10
            With wsA.Cells(1, c)
                .Value = wsZ.Cells(ROW_HDR, c).Value
                .Font.Bold = True
                .Font.Color = vbWhite
                .Interior.Color = RGB(63, 63, 63)
                .Borders.Color = RGB(120, 120, 120)
                .Borders.Weight = xlThin
            End With
        Next c
        wsA.Rows(1).RowHeight = 24
        wsA.Columns(3).ColumnWidth = 23
        wsA.Columns(10).ColumnWidth = 32
        r = 2
    Else
        r = wsA.Cells(wsA.Rows.Count, 3).End(xlUp).Row + 1
        If r < 2 Then r = 2
    End If

    n = 0
    For i = LastRow(wsZ) To ROW_HDR + 1 Step -1
        d = wsZ.Cells(i, 8).Value
        If IsDate(d) Then
            If Format(CDate(d), "yyyy-mm") = klucz Then
                For c = 1 To 10
                    wsA.Cells(r, c).Value = wsZ.Cells(i, c).Value
                Next c
                wsA.Cells(r, 7).NumberFormat = "yyyy-mm-dd"
                wsA.Cells(r, 8).NumberFormat = "yyyy-mm-dd"
                With wsA.Range(wsA.Cells(r, 1), wsA.Cells(r, 10))
                    .Interior.Color = RGB(198, 239, 206)
                    .Borders.Color = RGB(158, 158, 158)
                    .Borders.Weight = xlThin
                End With
                r = r + 1
                wsZ.Rows(i).Delete Shift:=xlUp
                n = n + 1
            End If
        End If
    Next i
    wsA.Columns("A:J").AutoFit
    ArchiwizujMiesiac = n
End Function

' ---------------------------------------------------------------------
' Przycisk: GENERUJ RAPORT (arkusz "PODSUMOWANIE")
' Tworzy ladny, gotowy do wyslania plik z najwazniejszymi informacjami:
' pojazdy w rejestracji, zarejestrowane w biezacym miesiacu, statystyki
' wg urzedu oraz lista aut zarejestrowanych w tym miesiacu.
' ---------------------------------------------------------------------
Sub GenerujRaport()
    On Error GoTo Blad
    Dim wsW As Worksheet, wsZ As Worksheet, wsP As Worksheet
    Dim wbR As Workbook, ws As Worksheet
    Dim i As Long, j As Long, r As Long, c As Long, n As Long
    Dim wRej As Long, wKat As Long, wMies As Long
    Dim sumaDni As Double, ileDni As Long, maxCzek As Long
    Dim d As Variant, dz As Variant, progu As Date, plik As String
    Dim grupy As Object, k As Variant, czesci() As String
    Dim ksztalt As Shape, najw As Shape

    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")
    Set wsP = ThisWorkbook.Worksheets("PULPIT")
    progu = DateSerial(Year(Date), Month(Date), 1)
    Set grupy = CreateObject("Scripting.Dictionary")
    Dim odCzerwca As Date, wRok As Long
    odCzerwca = DateSerial(Year(Date), 6, 1)

    For i = ROW_HDR + 1 To LastRow(wsW)
        If Trim(CStr(wsW.Cells(i, 3).Value)) <> "" Then
            wRej = wRej + 1
            k = Trim(CStr(wsW.Cells(i, 1).Value)) & "|" & _
                Trim(CStr(wsW.Cells(i, 4).Value)) & "|" & _
                Trim(CStr(wsW.Cells(i, 6).Value))
            grupy(k) = grupy(k) + 1
            dz = wsW.Cells(i, 7).Value
            If IsDate(dz) Then
                If Date - CDate(dz) > maxCzek Then maxCzek = Date - CDate(dz)
            End If
        End If
    Next i
    For i = ROW_HDR + 1 To LastRow(wsZ)
        If Trim(CStr(wsZ.Cells(i, 3).Value)) <> "" Then
            wKat = wKat + 1
            d = wsZ.Cells(i, 8).Value
            If IsDate(d) Then
                If CDate(d) >= progu Then wMies = wMies + 1
            End If
            If IsNumeric(wsZ.Cells(i, 9).Value) And _
               Trim(CStr(wsZ.Cells(i, 9).Value)) <> "" Then
                If CDbl(wsZ.Cells(i, 9).Value) >= 0 Then
                    sumaDni = sumaDni + CDbl(wsZ.Cells(i, 9).Value)
                    ileDni = ileDni + 1
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = False
    Set wbR = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbR.Worksheets(1)
    ws.Name = "Raport"
    ActiveWindow.DisplayGridlines = False
    ws.Columns(1).ColumnWidth = 3
    For c = 2 To 7
        ws.Columns(c).ColumnWidth = 18
    Next c

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 8))
        .Merge
        .Value = "   RAPORT REJESTRACJI POJAZDOW  -  99rent  -  stan na " & _
            Format(Now, "yyyy-mm-dd hh:mm")
        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = vbWhite
        .Interior.Color = RGB(227, 6, 19)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 36

    ' logo z PULPITU (najwieksze zdjecie)
    For Each ksztalt In wsP.Shapes
        If ksztalt.Type = 13 Then
            If najw Is Nothing Then
                Set najw = ksztalt
            ElseIf ksztalt.Width > najw.Width Then
                Set najw = ksztalt
            End If
        End If
    Next ksztalt
    If Not najw Is Nothing Then
        On Error Resume Next   ' logo jest ozdoba - raport ma powstac zawsze
        najw.Copy
        ws.Paste ws.Range("B3")
        With ws.Shapes(ws.Shapes.Count)
            .LockAspectRatio = True
            .Width = 90
        End With
        On Error GoTo Blad
    End If

    ' zarejestrowane od czerwca tego roku (katalog + archiwa)
    Dim wsY As Worksheet, fY As Long, lY As Long
    For Each wsY In ThisWorkbook.Worksheets
        If wsY.Name = "Zarejestrowane" Or Left(wsY.Name, 8) = "Archiwum" Then
            If wsY.Name = "Zarejestrowane" Then
                fY = ROW_HDR + 1
            Else
                fY = 2
            End If
            lY = wsY.Cells(wsY.Rows.Count, 3).End(xlUp).Row
            For i = fY To lY
                If Trim(CStr(wsY.Cells(i, 3).Value)) <> "" Then
                    If IsDate(wsY.Cells(i, 8).Value) Then
                        If CDate(wsY.Cells(i, 8).Value) >= odCzerwca Then
                            wRok = wRok + 1
                        End If
                    End If
                End If
            Next i
        End If
    Next wsY

    Dim kpiL(4) As String, kpiW(4) As Variant, c0 As Long
    kpiL(0) = "POJAZDY" & vbLf & "W REJESTRACJI": kpiW(0) = wRej
    kpiL(1) = "POJAZDY" & vbLf & "ZAREJESTROWANE": kpiW(1) = wKat
    kpiL(2) = "ZAREJESTROWANE" & vbLf & "W TYM MIESIACU": kpiW(2) = wMies
    If ileDni > 0 Then
        kpiL(3) = "SREDNI CZAS" & vbLf & "REJESTRACJI (DNI)"
        kpiW(3) = Round(sumaDni / ileDni, 1)
    Else
        kpiL(3) = "SREDNI CZAS" & vbLf & "REJESTRACJI (DNI)": kpiW(3) = "-"
    End If
    kpiL(4) = "ZAREJESTROWANE" & vbLf & "OD CZERWCA " & Year(Date): kpiW(4) = wRok
    For i = 0 To 4
        c0 = 4 + i
        With ws.Cells(3, c0)
            .Value = kpiW(i)
            .Font.Bold = True
            .Font.Size = 22
            .Font.Color = RGB(227, 6, 19)
            .Interior.Color = vbWhite
            .HorizontalAlignment = xlCenter
            .Borders.Color = RGB(217, 217, 217)
            .Borders.Weight = xlThin
        End With
        With ws.Cells(4, c0)
            .Value = kpiL(i)
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(89, 89, 89)
            .Interior.Color = vbWhite
            .HorizontalAlignment = xlCenter
            .WrapText = True
            .Borders.Color = RGB(217, 217, 217)
            .Borders.Weight = xlThin
        End With
    Next i
    ws.Rows(3).RowHeight = 34
    ws.Rows(4).RowHeight = 26

    r = 7
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Merge
        .Value = "  W REJESTRACJI - MARKA / DEALER / URZAD / SZTUK"
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(227, 6, 19)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 22
    r = r + 1
    Dim naglowkiG As Variant
    naglowkiG = Array("Marka", "Dealer", "Urzad", "Sztuk")
    For i = 0 To 3
        With ws.Cells(r, 2 + i)
            .Value = naglowkiG(i)
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(63, 63, 63)
            .Borders.Weight = xlThin
        End With
    Next i
    r = r + 1
    For Each k In grupy.Keys
        czesci = Split(CStr(k), "|")
        ws.Cells(r, 2).Value = czesci(0)
        ws.Cells(r, 3).Value = czesci(1)
        ws.Cells(r, 4).Value = czesci(2)
        ws.Cells(r, 5).Value = grupy(k)
        ws.Cells(r, 5).HorizontalAlignment = xlCenter
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
            .Interior.Color = RGB(255, 249, 196)
            .Borders.Color = RGB(158, 158, 158)
            .Borders.Weight = xlThin
        End With
        r = r + 1
    Next k
    ws.Cells(r, 2).Value = "RAZEM"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 5).Value = wRej
    ws.Cells(r, 5).Font.Bold = True
    ws.Cells(r, 5).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Interior.Color = RGB(242, 242, 242)
        .Borders.Weight = xlThin
    End With
    r = r + 2

    ' ranking urzedow: kto obsluguje najszybciej (katalog + archiwa)
    Dim urzS As Object, urzN As Object, urzA As Object
    Dim wsX As Worksheet, firstR As Long, lastR As Long, u As Variant
    Set urzS = CreateObject("Scripting.Dictionary")
    Set urzN = CreateObject("Scripting.Dictionary")
    Set urzA = CreateObject("Scripting.Dictionary")
    For Each wsX In ThisWorkbook.Worksheets
        If wsX.Name = "Zarejestrowane" Or Left(wsX.Name, 8) = "Archiwum" Then
            If wsX.Name = "Zarejestrowane" Then
                firstR = ROW_HDR + 1
            Else
                firstR = 2
            End If
            lastR = wsX.Cells(wsX.Rows.Count, 3).End(xlUp).Row
            For i = firstR To lastR
                If Trim(CStr(wsX.Cells(i, 3).Value)) <> "" Then
                    u = UCase(Trim(CStr(wsX.Cells(i, 6).Value)))
                    If u = "" Then u = "(BRAK URZEDU)"
                    urzA(u) = urzA(u) + 1
                    If IsNumeric(wsX.Cells(i, 9).Value) And _
                       Trim(CStr(wsX.Cells(i, 9).Value)) <> "" Then
                        If CDbl(wsX.Cells(i, 9).Value) >= 0 Then
                            urzS(u) = urzS(u) + CDbl(wsX.Cells(i, 9).Value)
                            urzN(u) = urzN(u) + 1
                        End If
                    End If
                End If
            Next i
        End If
    Next wsX

    Dim kl() As String, av() As Double, cnt2() As Long, nk As Long
    Dim tmpS As String, tmpD As Double, tmpL As Long, maxCnt As Long
    nk = urzA.Count
    If nk > 0 Then
        ReDim kl(0 To nk - 1)
        ReDim av(0 To nk - 1)
        ReDim cnt2(0 To nk - 1)
        i = 0
        For Each u In urzA.Keys
            kl(i) = CStr(u)
            cnt2(i) = urzA(u)
            av(i) = 9999
            If urzN.Exists(u) Then
                If urzN(u) > 0 Then av(i) = urzS(u) / urzN(u)
            End If
            i = i + 1
        Next u
        For i = 0 To nk - 2
            For j = i + 1 To nk - 1
                If av(j) < av(i) Then
                    tmpD = av(i): av(i) = av(j): av(j) = tmpD
                    tmpS = kl(i): kl(i) = kl(j): kl(j) = tmpS
                    tmpL = cnt2(i): cnt2(i) = cnt2(j): cnt2(j) = tmpL
                End If
            Next j
        Next i

        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
            .Merge
            .Value = "  RANKING URZEDOW - KTO OBSLUGUJE NAJSZYBCIEJ  " & _
                "(liczba urzedow: " & nk & ")"
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(227, 6, 19)
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(r).RowHeight = 22
        r = r + 1
        Dim nagR As Variant
        nagR = Array("Miejsce", "Urzad", "Sr. czas rej. (dni)", _
            "Zarejestrowane (szt.)")
        For i = 0 To 3
            With ws.Cells(r, 2 + i)
                .Value = nagR(i)
                .Font.Bold = True
                .Font.Color = vbWhite
                .Interior.Color = RGB(63, 63, 63)
                .Borders.Weight = xlThin
            End With
        Next i
        r = r + 1
        maxCnt = 0
        For i = 0 To nk - 1
            If cnt2(i) > maxCnt Then maxCnt = cnt2(i)
        Next i
        For i = 0 To nk - 1
            ws.Cells(r, 2).Value = i + 1
            ws.Cells(r, 2).HorizontalAlignment = xlCenter
            ws.Cells(r, 3).Value = kl(i)
            If av(i) < 9999 Then
                ws.Cells(r, 4).Value = Round(av(i), 1)
            Else
                ws.Cells(r, 4).Value = "-"
            End If
            ws.Cells(r, 4).HorizontalAlignment = xlCenter
            ws.Cells(r, 5).Value = cnt2(i)
            ws.Cells(r, 5).HorizontalAlignment = xlCenter
            With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
                .Interior.Color = IIf(i = 0, RGB(198, 239, 206), vbWhite)
                .Borders.Color = RGB(200, 200, 200)
                .Borders.Weight = xlThin
            End With
            If cnt2(i) = maxCnt Then
                ws.Cells(r, 5).Font.Bold = True
                ws.Cells(r, 5).Font.Color = RGB(227, 6, 19)
            End If
            r = r + 1
        Next i
        ws.Cells(r, 2).Value = "1. miejsce = najkrotszy sredni czas; " & _
            "czerwona liczba = najwiecej rejestracji. Ranking liczony ze " & _
            "WSZYSTKICH miesiecy (katalog + zakladki Archiwum)."
        ws.Cells(r, 2).Font.Italic = True
        ws.Cells(r, 2).Font.Size = 9
        r = r + 2
    End If

    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
        .Merge
        .Value = "  ZAREJESTROWANE W TYM MIESIACU (" & Format(Date, "yyyy-mm") & ")"
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(227, 6, 19)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 22
    r = r + 1
    Dim naglowki As Variant
    naglowki = Array("Marka", "Model", "VIN", "Nr rejestracyjny", _
        "Data rejestracji", "Czas (dni)")
    For i = 0 To 5
        With ws.Cells(r, 2 + i)
            .Value = naglowki(i)
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(63, 63, 63)
            .Borders.Weight = xlThin
        End With
    Next i
    r = r + 1
    n = 0
    For i = ROW_HDR + 1 To LastRow(wsZ)
        d = wsZ.Cells(i, 8).Value
        If IsDate(d) Then
            If CDate(d) >= progu Then
                ws.Cells(r, 2).Value = wsZ.Cells(i, 1).Value
                ws.Cells(r, 3).Value = wsZ.Cells(i, 2).Value
                ws.Cells(r, 4).Value = wsZ.Cells(i, 3).Value
                ws.Cells(r, 5).Value = wsZ.Cells(i, 4).Value
                ws.Cells(r, 6).Value = CDate(d)
                ws.Cells(r, 6).NumberFormat = "yyyy-mm-dd"
                ws.Cells(r, 7).Value = wsZ.Cells(i, 9).Value
                With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
                    .Interior.Color = RGB(198, 239, 206)
                    .Borders.Color = RGB(158, 158, 158)
                    .Borders.Weight = xlThin
                End With
                r = r + 1
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then
        ws.Cells(r, 2).Value = "(brak rejestracji w biezacym miesiacu)"
        ws.Cells(r, 2).Font.Italic = True
    End If

    ws.Columns("B:G").AutoFit
    For c = 2 To 7
        If ws.Columns(c).ColumnWidth < 14 Then ws.Columns(c).ColumnWidth = 14
    Next c

    ' stopka
    If n > 0 Then r = r + 1
    ws.Cells(r + 1, 2).Value = "Stworzone przez: Oskar Figaszewski"
    ws.Cells(r + 1, 2).Font.Italic = True
    ws.Cells(r + 1, 2).Font.Size = 9
    ws.Cells(r + 1, 2).Font.Color = RGB(127, 127, 127)

    plik = SciezkaArchiwum() & Application.PathSeparator & _
        "Raport_99rent_" & Format(Now, "yyyy-mm-dd_hhmm") & ".xlsx"
    Application.DisplayAlerts = False
    wbR.SaveAs Filename:=plik, FileFormat:=51
    Application.DisplayAlerts = True
    wbR.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "Raport zapisany:" & vbCrLf & plik, vbInformation, "99rent"
    Exit Sub
Blad:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Blad generowania raportu: " & Err.Description, vbCritical, "99rent"
End Sub

' --------------------------- nawigacja (przyciski na pulpicie) --------
Sub IdzWRejestracji()
    Application.Goto ThisWorkbook.Worksheets("W rejestracji").Range("A1"), True
End Sub

Sub IdzZarejestrowane()
    Application.Goto ThisWorkbook.Worksheets("Zarejestrowane").Range("A1"), True
End Sub

Sub IdzPodsumowanie()
    Application.Goto ThisWorkbook.Worksheets("PODSUMOWANIE").Range("A1"), True
End Sub

Sub IdzImport()
    Application.Goto ThisWorkbook.Worksheets("Do rejestracji").Range("A1"), True
End Sub

Sub IdzPulpit()
    Application.Goto ThisWorkbook.Worksheets("PULPIT").Range("A1"), True
End Sub
'''


SHEET_MODULE_TEMPLATE = '''Attribute VB_Name = "%s"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'''
