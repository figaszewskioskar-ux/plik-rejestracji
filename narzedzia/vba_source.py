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
'''

MODULE1 = r'''Attribute VB_Name = "Module1"
Option Explicit

' =====================================================================
'  99rent - Raport rejestracji pojazdow
'  Makra obslugujace przyciski w skoroszycie.
'  Uklad arkuszy:
'    "W rejestracji"   - formularz w wierszu 5, tabela od wiersza 8
'    "Zarejestrowane"  - formularz w wierszu 5, tabela od wiersza 8
'    "Nr rejestracyjny" - VIN w C5, nr rej w C6, data rejestracji w C7
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
    If IsNumeric(ws.Cells(ROW_FORM, 10).Value) And _
       Trim(CStr(ws.Cells(ROW_FORM, 10).Value)) <> "" Then
        ws.Cells(r, 10).Value = CDbl(ws.Cells(ROW_FORM, 10).Value)
    End If

    ws.Range(ws.Cells(ROW_FORM, 1), ws.Cells(ROW_FORM, 8)).ClearContents
    ws.Cells(ROW_FORM, 10).ClearContents
    MsgBox "Pojazd " & vin & " dodany do katalogu zarejestrowanych.", _
        vbInformation, "99rent"
End Sub

' ---------------------------------------------------------------------
' Przycisk: ZAREJESTRUJ POJAZD (arkusz "Nr rejestracyjny")
' Nadaje numer rejestracyjny pojazdowi oczekujacemu: przenosi go
' z "W rejestracji" do katalogu "Zarejestrowane" i liczy dni.
' ---------------------------------------------------------------------
Sub ZarejestrujPojazd()
    Dim wsN As Worksheet, wsW As Worksheet, wsZ As Worksheet
    Dim vin As String, nrRej As String
    Dim dataRej As Date, dataZl As Variant
    Dim rw As Long, rz As Long, dni As String

    Set wsN = ThisWorkbook.Worksheets("Nr rejestracyjny")
    Set wsW = ThisWorkbook.Worksheets("W rejestracji")
    Set wsZ = ThisWorkbook.Worksheets("Zarejestrowane")

    vin = Trim(CStr(wsN.Range("C5").Value))
    nrRej = Trim(CStr(wsN.Range("C6").Value))
    If vin = "" Then
        MsgBox "Podaj numer VIN pojazdu.", vbExclamation, "99rent"
        Exit Sub
    End If
    If nrRej = "" Then
        MsgBox "Podaj numer rejestracyjny.", vbExclamation, "99rent"
        Exit Sub
    End If
    dataRej = SafeDate(wsN.Range("C7").Value, Date)

    rw = FindVinRow(wsW, vin)
    If rw = 0 Then
        MsgBox "Nie znaleziono VIN " & vin & " w arkuszu W rejestracji." _
            & vbCrLf & "Uzyj formularza w arkuszu Zarejestrowane, jesli " _
            & "pojazd nie przechodzil przez rejestr.", vbExclamation, "99rent"
        Exit Sub
    End If
    If FindVinRow(wsZ, vin) > 0 Then
        MsgBox "Pojazd o VIN " & vin & " jest juz w katalogu.", vbExclamation, "99rent"
        Exit Sub
    End If

    dataZl = wsW.Cells(rw, 7).Value

    rz = LastRow(wsZ) + 1
    wsZ.Cells(rz, 1).Value = wsW.Cells(rw, 1).Value             ' Marka
    wsZ.Cells(rz, 2).Value = wsW.Cells(rw, 2).Value             ' Model
    wsZ.Cells(rz, 3).Value = vin                                 ' VIN
    wsZ.Cells(rz, 4).Value = nrRej                               ' Nr rej
    wsZ.Cells(rz, 5).Value = wsW.Cells(rw, 4).Value             ' Dealer
    wsZ.Cells(rz, 6).Value = wsW.Cells(rw, 6).Value             ' Urzad
    If IsDate(dataZl) Then
        wsZ.Cells(rz, 7).Value = CDate(dataZl)
        wsZ.Cells(rz, 7).NumberFormat = "yyyy-mm-dd"
    End If
    wsZ.Cells(rz, 8).Value = dataRej
    wsZ.Cells(rz, 8).NumberFormat = "yyyy-mm-dd"
    wsZ.Cells(rz, 9).Formula = "=IF(OR($G" & rz & "=" & Chr(34) & Chr(34) & _
        ",$H" & rz & "=" & Chr(34) & Chr(34) & ")," & Chr(34) & Chr(34) & _
        ",$H" & rz & "-$G" & rz & ")"

    wsW.Rows(rw).Delete Shift:=xlUp

    If IsDate(dataZl) Then
        dni = CStr(CLng(dataRej - CDate(dataZl)))
        MsgBox "Pojazd " & vin & " zarejestrowany jako " & nrRej & "." & _
            vbCrLf & "Czas rejestracji: " & dni & " dni.", vbInformation, "99rent"
    Else
        MsgBox "Pojazd " & vin & " zarejestrowany jako " & nrRej & ".", _
            vbInformation, "99rent"
    End If

    wsN.Range("C5:C7").ClearContents
End Sub

' --------------------------- nawigacja (przyciski na pulpicie) --------
Sub IdzWRejestracji()
    Application.Goto ThisWorkbook.Worksheets("W rejestracji").Range("A1"), True
End Sub

Sub IdzZarejestrowane()
    Application.Goto ThisWorkbook.Worksheets("Zarejestrowane").Range("A1"), True
End Sub

Sub IdzNrRej()
    Application.Goto ThisWorkbook.Worksheets("Nr rejestracyjny").Range("A1"), True
End Sub

Sub IdzPodsumowanie()
    Application.Goto ThisWorkbook.Worksheets("PODSUMOWANIE").Range("A1"), True
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
