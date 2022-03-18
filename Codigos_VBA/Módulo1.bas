Attribute VB_Name = "Módulo1"
Sub Desagrupar_TabDinFixa_1()
'
' Macro1 Macro
'
'
For Each c In Worksheets("Plan1").Range("C54:W54")
    c.Select
    Selection.ShowDetail = True
    Sheets("Plan1").Select
Next c

    
    
    
End Sub

Sub Desagrupar_TabDinFixa_2()
'
' Macro1 Macro
'
'
For Each c In Worksheets("Plan1").Range("C134:J134")
    c.Select
    Selection.ShowDetail = True
    Sheets("Plan1").Select
Next c

    
    
    
End Sub
