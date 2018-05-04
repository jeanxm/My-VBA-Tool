Attribute VB_Name = "Module1"
Public pintMaxL As Integer
Public pintMaxR As Integer
Public pstrRst1 As String
Public pstrRst2 As String
Public pstrRst3 As String
Public prngRst1 As Range
Public prngRst2 As Range
Public prngRst3 As Range
Public pshtRst1 As Worksheet
Public pshtRst2 As Worksheet
Public pshtRst3 As Worksheet

Sub LimitCells(ByVal prngCells As Range)
    Dim cell As Range, pstrCell As String, arr As Variant, brr As Variant, crr As Variant, m As Long, n As Integer, str1 As String, rng1 As Range
    ReDim brr(1 To prngCells.Cells.Count)
    m = 1
    For Each cell In prngCells
        pstrCell = cell.Text
        arr = Split(pstrCell, Chr(10)) 'LBound is 0
        If UBound(arr) > pintMaxR - 1 Then
            'cell.ClearContents
            cell = AdjustCells(cell)
            cell.Select
            Call FormatCell
            brr(m) = cell.Address
            m = m + 1
        Else
            For n = 0 To UBound(arr)
                If Len(arr(n)) > pintMaxL Then
                    'cell.ClearContents
                    cell = AdjustCells(cell)
                    cell.Select
                    Call FormatCell
                    brr(m) = cell.Address
                    m = m + 1
                    Exit For
                End If
            Next
        End If
    Next
    'when any cell is out of range
    If m > 1 Then
        ReDim crr(1 To m - 1)
        For n = 1 To m - 1
            crr(n) = brr(n)
        Next
        str1 = Join(crr, ",")
        MsgBox "Please refill " & str1 & ". Each cell should not have more than " & pintMaxR & " lines and each line should not have more than " & pintMaxL & " characters.", vbExclamation
        Erase crr
    End If
    Erase brr
    Erase arr
End Sub

Sub CheckCells()
    pintMaxL = ThisWorkbook.Worksheets(1).Range("B8")
    pintMaxR = ThisWorkbook.Worksheets(1).Range("B7")
    pstrRst1 = ThisWorkbook.Worksheets(1).Range("B2")
    pstrRst2 = ThisWorkbook.Worksheets(1).Range("B3")
    pstrRst3 = ThisWorkbook.Worksheets(1).Range("B4")
    If FindSht(ThisWorkbook.Worksheets(1).Range("A2").Text) = True Then Set pshtRst1 = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets(1).Range("A2").Text)
    If FindSht(ThisWorkbook.Worksheets(1).Range("A3").Text) = True Then Set pshtRst2 = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets(1).Range("A3").Text)
    If FindSht(ThisWorkbook.Worksheets(1).Range("A4").Text) = True Then Set pshtRst3 = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets(1).Range("A4").Text)
    Call GetRst
    
    If FindSht(ThisWorkbook.Worksheets(1).Range("A2").Text) = True And pstrRst1 <> "" Then
        pshtRst1.Activate
        Call LimitCells(prngRst1)
    End If
    
    If FindSht(ThisWorkbook.Worksheets(1).Range("A3").Text) = True And pstrRst2 <> "" Then
        pshtRst2.Activate
        Call LimitCells(prngRst2)
    End If
    
    If FindSht(ThisWorkbook.Worksheets(1).Range("A4").Text) = True And pstrRst3 <> "" Then
        pshtRst3.Activate
        Call LimitCells(prngRst3)
    End If
End Sub
 
Function AdjustCells(ByVal pstrCell As String)
    Dim str As String, trr As Variant, n As Integer
    str = Left(Application.WorksheetFunction.Substitute(pstrCell, Chr(10), ""), pintMaxL * pintMaxR)
    If Len(str) > pintMaxL Then
        ReDim trr(1 To pintMaxR)
        n = 0
        Do Until Mid(str, pintMaxL * n + 1, pintMaxL) = ""
            trr(n + 1) = Mid(str, pintMaxL * n + 1, pintMaxL)
            n = n + 1
        Loop
        AdjustCells = Join(trr, Chr(10))
        Erase trr
    Else
        AdjustCells = str
    End If
End Function

Sub GetRst()
    If FindSht(ThisWorkbook.Worksheets(1).Range("A2").Text) = True And pstrRst1 <> "" Then
         Set prngRst1 = pshtRst1.Range(pstrRst1)
    End If
    If FindSht(ThisWorkbook.Worksheets(1).Range("A3").Text) = True And pstrRst2 <> "" Then
         Set prngRst2 = pshtRst2.Range(pstrRst2)
    End If
    If FindSht(ThisWorkbook.Worksheets(1).Range("A4").Text) = True And pstrRst3 <> "" Then
         Set prngRst3 = pshtRst3.Range(pstrRst3)
    End If
End Sub

Function FindSht(ByVal pstrSht As String) As Boolean
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name = pstrSht Then
            FindSht = True
            Exit Function
        End If
    Next
    FindSht = False
End Function

Sub FormatCell()
    With ActiveWindow.RangeSelection
        .ColumnWidth = 20
        .RowHeight = 30
    End With
End Sub
