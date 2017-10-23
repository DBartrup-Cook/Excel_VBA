Attribute VB_Name = "mdl_WorkbookSharing"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : MakeExclusive & MakeShared
' Author    : Andy Pope
' Date      : 24/09/2003
' Purpose   : Adds and removes workbook sharing.
'             http://www.ozgrid.com/forum/showthread.php?t=16086
'---------------------------------------------------------------------------------------
Sub MakeExclusive()
     
    If ActiveWorkbook.MultiUserEditing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.ExclusiveAccess
        Application.DisplayAlerts = True
    End If
     
End Sub
 
Sub MakeShared()
     
    If Not ActiveWorkbook.MultiUserEditing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs ActiveWorkbook.Name, accessmode:=xlShared
        Application.DisplayAlerts = True
    End If
     
End Sub
