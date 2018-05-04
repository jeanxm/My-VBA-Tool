VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    pintMaxL = ThisWorkbook.Worksheets(1).Range("B8")
    pintMaxR = ThisWorkbook.Worksheets(1).Range("B7")
    pstrRst1 = ThisWorkbook.Worksheets(1).Range("B2")
    pstrRst2 = ThisWorkbook.Worksheets(1).Range("B3")
    pstrRst3 = ThisWorkbook.Worksheets(1).Range("B4")
    If FindSht(ThisWorkbook.Worksheets(1).Range("A2").Text) = True Then Set pshtRst1 = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets(1).Range("A2").Text)
    If FindSht(ThisWorkbook.Worksheets(1).Range("A3").Text) = True Then Set pshtRst2 = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets(1).Range("A3").Text)
    If FindSht(ThisWorkbook.Worksheets(1).Range("A4").Text) = True Then Set pshtRst3 = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets(1).Range("A4").Text)
    Call GetRst
    
    If FindSht(ThisWorkbook.Worksheets(1).Range("A2").Text) = True And Sh.Name = pshtRst1.Name And pstrRst1 <> "" Then
        If Not Application.Intersect(Target, prngRst1) Is Nothing Then
            Sh.Activate
            Call LimitCells(Application.Intersect(Target, prngRst1))
        End If
    End If
    
    If FindSht(ThisWorkbook.Worksheets(1).Range("A3").Text) = True And Sh.Name = pshtRst2.Name And pstrRst2 <> "" Then
        If Not Application.Intersect(Target, prngRst2) Is Nothing Then
            Sh.Activate
            Call LimitCells(Application.Intersect(Target, prngRst2))
        End If
    End If
    
    If FindSht(ThisWorkbook.Worksheets(1).Range("A4").Text) = True And Sh.Name = pshtRst3.Name And pstrRst3 <> "" Then
        If Not Application.Intersect(Target, prngRst3) Is Nothing Then
            Sh.Activate
            Call LimitCells(Application.Intersect(Target, prngRst3))
        End If
    End If
End Sub
