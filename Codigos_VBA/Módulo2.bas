Attribute VB_Name = "Módulo2"
Sub ConsolidacaoDeDados_1()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long
    Dim Plan1 As String
    Dim PlanConsolidada As String
    Dim FPath As String
    
    

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    
    'Define the name as "PlanConsolidada1" and the costom path to export in the end of this macro'
    
    FPath = "C:\Users\vinic\Documents\GitHub\ANP\assets\PlanConsolidada1"
    
  

    'Delete the sheet "PlanConsolidada1" if it exist
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("PlanConsolidada1").Delete
    On Error GoTo 0
    Application.DisplayAlerts = False
    
    

    'Add a worksheet with the name "PlanConsolidada1"
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "PlanConsolidada1"
     
    'Add a title in the first row"
    
    With Sheets("PlanConsolidada1")
    .Range("a1").Value = "COMBUSTÍVEL"
    .Range("b1").Value = "ANO"
    .Range("c1").Value = "REGIÃO"
    .Range("d1").Value = "ESTADO"
    .Range("e1").Value = "UNIDADE"
    .Range("f1").Value = "JAN"
    .Range("g1").Value = "FEV"
    .Range("h1").Value = "MAR"
    .Range("i1").Value = "ABR"
    .Range("j1").Value = "MAI"
    .Range("k1").Value = "JUN"
    .Range("l1").Value = "JUL"
    .Range("m1").Value = "AGO"
    .Range("n1").Value = "SET"
    .Range("o1").Value = "OUT"
    .Range("p1").Value = "NOV"
    .Range("q1").Value = "DEZ"
    .Range("r1").Value = "TOTAL"

End With
    
    
    'Fill in the start row (without header)
    StartRow = 2

    'loop through all worksheets and copy the data to the DestSh
    For Each sh In ActiveWorkbook.Worksheets

        'Loop through all worksheets except the PlanConsolidada worksheet and the
        'Plan1 worksheet, you can ad more sheets to the array if you want.
        If IsError(Application.Match(sh.Name, _
                                     Array(DestSh.Name, "Plan1"), 0)) Then

            'Find the last row with data on the DestSh and sh
            
            Last = DestSh.Cells.Find(What:="*", After:=DestSh.Range("A1"), Lookat:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
            shLast = sh.Cells.Find(What:="*", After:=sh.Range("A1"), Lookat:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
            

            'If sh is not empty and if the last row >= StartRow copy the CopyRng
            If shLast > 0 And shLast >= StartRow Then

                'Set the range that you want to copy
                Set CopyRng = sh.Range(sh.Rows(StartRow), sh.Rows(shLast))

                'Test if there enough rows in the DestSh to copy all the data
                If Last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                    MsgBox "There are not enough rows in the Destsh"
                    GoTo ExitTheSub
                End If

                'This example copies values/formats, if you only want to copy the
                'values or want to copy everything look below example 1 on this page
                CopyRng.Copy
                With DestSh.Cells(Last + 1, "A")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With

            End If

        End If
    Next



'Export the new worksheet to a new archive'


ThisWorkbook.Sheets("PlanConsolidada1").Copy
ActiveWorkbook.SaveAs FPath, FileFormat:=xlCSVUTF8





'Delete worksheets after merge data keeping just the Plan1'

For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> "Plan1" Then
       sh.Delete
    End If
Application.DisplayAlerts = False
Next sh







'FINAL'

ExitTheSub:

    'Application.GoTo DestSh.Cells(1)

    'AutoFit the column width in the DestSh sheet
    'DestSh.Columns.AutoFit

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
End Sub

Sub ConsolidacaoDeDados_2()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim shLast As Long
    Dim CopyRng As Range
    Dim StartRow As Long
    Dim Plan1 As String
    Dim PlanConsolidada As String
    Dim FPath As String
    
    

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    
    'Define the name as "PlanConsolidada1" and the costom path to export in the end of this macro'
    
    FPath = "C:\Users\vinic\Documents\GitHub\ANP\assets\PlanConsolidada2"
    
  

    'Delete the sheet "PlanConsolidada1" if it exist
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("PlanConsolidada2").Delete
    On Error GoTo 0
    Application.DisplayAlerts = False
    
    

    'Add a worksheet with the name "PlanConsolidada1"
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.Name = "PlanConsolidada2"
     
    'Add a title in the first row"
    
    With Sheets("PlanConsolidada2")
    .Range("a1").Value = "COMBUSTÍVEL"
    .Range("b1").Value = "ANO"
    .Range("c1").Value = "REGIÃO"
    .Range("d1").Value = "ESTADO"
    .Range("e1").Value = "UNIDADE"
    .Range("f1").Value = "JAN"
    .Range("g1").Value = "FEV"
    .Range("h1").Value = "MAR"
    .Range("i1").Value = "ABR"
    .Range("j1").Value = "MAI"
    .Range("k1").Value = "JUN"
    .Range("l1").Value = "JUL"
    .Range("m1").Value = "AGO"
    .Range("n1").Value = "SET"
    .Range("o1").Value = "OUT"
    .Range("p1").Value = "NOV"
    .Range("q1").Value = "DEZ"
    .Range("r1").Value = "TOTAL"

End With
    
    
    'Fill in the start row (without header)
    StartRow = 2

    'loop through all worksheets and copy the data to the DestSh
    For Each sh In ActiveWorkbook.Worksheets

        'Loop through all worksheets except the PlanConsolidada worksheet and the
        'Plan1 worksheet, you can ad more sheets to the array if you want.
        If IsError(Application.Match(sh.Name, _
                                     Array(DestSh.Name, "Plan1"), 0)) Then

            'Find the last row with data on the DestSh and sh
            
            Last = DestSh.Cells.Find(What:="*", After:=DestSh.Range("A1"), Lookat:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
            shLast = sh.Cells.Find(What:="*", After:=sh.Range("A1"), Lookat:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
            

            'If sh is not empty and if the last row >= StartRow copy the CopyRng
            If shLast > 0 And shLast >= StartRow Then

                'Set the range that you want to copy
                Set CopyRng = sh.Range(sh.Rows(StartRow), sh.Rows(shLast))

                'Test if there enough rows in the DestSh to copy all the data
                If Last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                    MsgBox "There are not enough rows in the Destsh"
                    GoTo ExitTheSub
                End If

                'This example copies values/formats, if you only want to copy the
                'values or want to copy everything look below example 1 on this page
                CopyRng.Copy
                With DestSh.Cells(Last + 1, "A")
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End With

            End If

        End If
    Next



'Export the new worksheet to a new archive'


ThisWorkbook.Sheets("PlanConsolidada2").Copy
ActiveWorkbook.SaveAs FPath, FileFormat:=xlCSVUTF8





'Delete worksheets after merge data keeping just the Plan1'

For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> "Plan1" Then
       sh.Delete
    End If
Application.DisplayAlerts = False
Next sh







'FINAL'

ExitTheSub:

    'Application.GoTo DestSh.Cells(1)

    'AutoFit the column width in the DestSh sheet
    'DestSh.Columns.AutoFit

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
End Sub



