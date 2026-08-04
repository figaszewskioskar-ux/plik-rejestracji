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
    Dim r As Range
    Set r = ws.Range(ws.Cells(ROW_HDR + 1, 3), ws.Cells(LastRow(ws), 3))
    Dim c As Range
    For Each c In r.Cells
        If UCase(Trim(CStr(c.Value))) = UCase(Trim(vin)) Then
            FindVinRow = c.Row
            Exit Function
        End If
    Next c
    FindVinRow = 0
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
    PaintRow ws, r, 9, RGB(255, 249, 196)
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
    ws.Cells(r, 9).Value = ws.Cells(ROW_FORM, 9).Value          ' Uwagi
    PaintYellow ws, r

    ws.Range(ws.Cells(ROW_FORM, 1), ws.Cells(ROW_FORM, 7)).ClearContents
    ws.Cells(ROW_FORM, 9).ClearContents
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
        ",$H" & r & "=" & Chr(34) & Chr(34) & ")," & Chr(34) & Chr(34) & _
        ",$H" & r & "-$G" & r & ")"
    ws.Cells(r, 10).Value = ws.Cells(ROW_FORM, 10).Value        ' Uwagi
    PaintGreen ws, r

    ws.Range(ws.Cells(ROW_FORM, 1), ws.Cells(ROW_FORM, 8)).ClearContents
    ws.Cells(ROW_FORM, 10).ClearContents
    MsgBox "Pojazd " & vin & " dodany do katalogu zarejestrowanych.", _
        vbInformation, "99rent"
End Sub

' ---------------------------------------------------------------------
' Przycisk: IMPORTUJ DO REJESTRU (arkusz "Import hurtowy")
' Wkleja sie wiele pojazdow naraz (10, 20, 30...) w tabele importu,
' a makro przenosi je wszystkie do arkusza "W rejestracji".
' ---------------------------------------------------------------------
Sub DodajHurtowo()
    Dim wsI As Worksheet, wsW As Worksheet, wsZ As Worksheet
    Dim i As Long, r As Long, n As Long, skipped As Long, lastI As Long
    Dim vin As String, msg As String

    Set wsI = ThisWorkbook.Worksheets("Import hurtowy")
    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")

    lastI = LastRow(wsI)
    If lastI <= ROW_HDR Then
        MsgBox "Wklej pojazdy w tabele od wiersza " & (ROW_HDR + 1) & _
            " (kolumna VIN jest wymagana).", vbExclamation, "99rent"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    n = 0: skipped = 0
    For i = ROW_HDR + 1 To lastI
        vin = Trim(CStr(wsI.Cells(i, 3).Value))
        If vin <> "" Then
            If FindVinRow(wsW, vin) > 0 Or FindVinRow(wsZ, vin) > 0 Then
                skipped = skipped + 1
            Else
                r = LastRow(wsW) + 1
                wsW.Cells(r, 1).Value = wsI.Cells(i, 1).Value    ' Marka
                wsW.Cells(r, 2).Value = wsI.Cells(i, 2).Value    ' Model
                wsW.Cells(r, 3).Value = vin                       ' VIN
                wsW.Cells(r, 4).Value = wsI.Cells(i, 4).Value    ' Dealer
                wsW.Cells(r, 5).Value = wsI.Cells(i, 5).Value    ' Wspolwl.
                wsW.Cells(r, 6).Value = wsI.Cells(i, 6).Value    ' Urzad
                wsW.Cells(r, 7).Value = SafeDate(wsI.Cells(i, 7).Value, Date)
                wsW.Cells(r, 7).NumberFormat = "yyyy-mm-dd"
                wsW.Cells(r, 8).Formula = "=IF($G" & r & "=" & Chr(34) & Chr(34) & _
                    "," & Chr(34) & Chr(34) & ",TODAY()-$G" & r & ")"
                wsW.Cells(r, 9).Value = wsI.Cells(i, 8).Value    ' Uwagi
                PaintYellow wsW, r
                n = n + 1
            End If
        End If
    Next i
    If n > 0 Then
        wsI.Range(wsI.Cells(ROW_HDR + 1, 1), wsI.Cells(lastI, 8)).ClearContents
    End If
    Application.ScreenUpdating = True

    msg = "Zaimportowano pojazdow: " & n & "."
    If skipped > 0 Then msg = msg & vbCrLf & _
        "Pominieto (VIN juz istnieje): " & skipped & "."
    MsgBox msg, vbInformation, "99rent"
End Sub

' ---------------------------------------------------------------------
' Przycisk: ZAREJESTRUJ ZAZNACZONE (arkusz "W rejestracji")
' Zaznacz dowolne komorki wierszy pojazdow (mozna wiele naraz) i kliknij:
' pojazdy przechodza do katalogu "Zarejestrowane" z dzisiejsza data
' rejestracji. Numer rejestracyjny nadasz potem w zakladce Nr rejestracyjny.
' ---------------------------------------------------------------------
Sub ZarejestrujZaznaczone()
    Dim wsW As Worksheet, wsZ As Worksheet
    Dim cell As Range, rowsToMove As Object, key As Variant
    Dim rw As Long, rz As Long, n As Long
    Dim vin As String, dataZl As Variant, arr() As Long, i As Long, j As Long, tmp As Long

    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")

    If ActiveSheet.Name <> wsW.Name Then
        MsgBox "Przejdz do arkusza W rejestracji i zaznacz wiersze pojazdow.", _
            vbExclamation, "99rent"
        Exit Sub
    End If

    Set rowsToMove = CreateObject("Scripting.Dictionary")
    For Each cell In Selection.Cells
        If cell.Row > ROW_HDR And cell.Row <= LastRow(wsW) Then
            vin = Trim(CStr(wsW.Cells(cell.Row, 3).Value))
            If vin <> "" And Not rowsToMove.Exists(cell.Row) Then
                rowsToMove.Add cell.Row, vin
            End If
        End If
    Next cell

    If rowsToMove.Count = 0 Then
        MsgBox "Zaznacz co najmniej jeden wiersz pojazdu (z VIN) w tabeli.", _
            vbExclamation, "99rent"
        Exit Sub
    End If

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
                ",$H" & rz & "=" & Chr(34) & Chr(34) & ")," & Chr(34) & Chr(34) & _
                ",$H" & rz & "-$G" & rz & ")"
            wsZ.Cells(rz, 10).Value = wsW.Cells(rw, 9).Value     ' Uwagi
            PaintGreen wsZ, rz
            n = n + 1
        End If
        wsW.Rows(rw).Delete Shift:=xlUp
    Next i
    Application.ScreenUpdating = True

    MsgBox "Przeniesiono do katalogu: " & n & " pojazd(y)." & vbCrLf & _
        "Numer rejestracyjny wpisz w kolumnie D katalogu.", _
        vbInformation, "99rent"
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
    Application.Goto ThisWorkbook.Worksheets("Import hurtowy").Range("A1"), True
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
