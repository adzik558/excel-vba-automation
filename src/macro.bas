Attribute VB_Name = "Module2"
Option Explicit

Sub CreateYearlyReport()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Dim yrWS As Worksheet
    Dim srcRange As Range
    Dim pasteRow As Long
    Dim firstPaste As Boolean
    
    ' SprawdŸ czy istnieje arkusz YEARLY REPORT, jeœli nie - utwórz
    On Error Resume Next
    Set yrWS = ThisWorkbook.Worksheets("YEARLY REPORT")
    On Error GoTo ErrHandler
    If yrWS Is Nothing Then
        Set yrWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        yrWS.Name = "YEARLY REPORT"
    End If
    
    ' Wstaw nag³ówki tylko raz (jeœli pierwszy wiersz nie zawiera oczekiwanych nag³ówków)
    If Application.WorksheetFunction.CountA(yrWS.Rows(1)) = 0 Then
        InsertHeadersToSheet yrWS
        FormatHeadersOnSheet yrWS
    End If
    
    firstPaste = True
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> yrWS.Name Then
            ' Zak³adamy, ¿e dane zaczynaj¹ siê w A2 i maj¹ co najmniej jedn¹ niepust¹ komórkê w kolumnie A
            If Application.WorksheetFunction.CountA(ws.Range("A:A")) > 1 Then
                ' ZnajdŸ zakres danych: od A2 do ostatniego wiersza i ostatniej kolumny
                Dim lastRow As Long, lastCol As Long
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column ' zak³ada, ¿e wiersz 2 ma pe³ne dane
                
                If lastRow >= 2 And lastCol >= 1 Then
                    Set srcRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
                    
                    ' ZnajdŸ wiersz do wklejenia na YEARLY REPORT (pierwsza pusta po istniej¹cych danych)
                    pasteRow = yrWS.Cells(yrWS.Rows.Count, "A").End(xlUp).Row + 1
                    ' Jeœli w Yearly Report jest tylko nag³ówek i nic poza tym to upewnij siê, ¿e pasteRow = 2
                    If pasteRow < 2 Then pasteRow = 2
                    
                    ' Kopiuj wartoœci i formaty (bez schodzenia do schowka)
                    srcRange.Copy
                    yrWS.Cells(pasteRow, "A").PasteSpecial xlPasteValuesAndNumberFormats
                    yrWS.Cells(pasteRow, "A").PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                    
                    firstPaste = False
                End If
            End If
        End If
    Next ws
    
    ' Po wklejeniu wszystkiego, ustaw formu³ê sumy w kolumnie F (jeœli istnieje kolumna F)
    AutomateTotalSUM_OnSheet yrWS
    
    ' Formatowanie nag³ówków finalne
    FormatHeadersOnSheet yrWS
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Roczny raport zaktualizowany.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Wyst¹pi³ b³¹d: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' Wstawia nag³ówki na wskazanym arkuszu (nie nadpisuje danych)
Private Sub InsertHeadersToSheet(targetWS As Worksheet)
    With targetWS
        .Rows(1).Insert Shift:=xlDown
        .Range("A1").Value = "Division"
        .Range("B1").Value = "Category"
        .Range("C1").Value = "Jan"
        .Range("D1").Value = "Feb"
        .Range("E1").Value = "Mar"
        .Range("F1").Value = "Total"
    End With
End Sub

Private Sub FormatHeadersOnSheet(targetWS As Worksheet)
    With targetWS.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = targetWS.Parent.Theme.ThemeColorScheme(1) ' fallback - mo¿na ustawiæ konkretny kolor
        .Font.Size = 12
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    ' Formatowanie kolumn z wartoœciami (C:F) jako waluta je¿eli istniej¹
    On Error Resume Next
    With targetWS
        .Range("C2", .Cells(.Rows.Count, "F").End(xlUp)).Style = "Currency"
        .Columns("B:F").EntireColumn.AutoFit
    End With
    On Error GoTo 0
End Sub

Private Sub AutomateTotalSUM_OnSheet(targetWS As Worksheet)
    Dim lastRow As Long
    With targetWS
        ' ZnajdŸ ostatni u¿yty wiersz w kolumnie F (jeœli kolumna F nie istnieje, wyjdŸ)
        lastRow = .Cells(.Rows.Count, "F").End(xlUp).Row
        If lastRow < 2 Then Exit Sub ' brak danych
        ' Wstaw sumê w pierwszej wolnej komórce pod ostatni¹ wartoœci¹
        .Cells(lastRow + 1, "F").Formula = "=SUM(F2:F" & lastRow & ")"
    End With
End Sub

