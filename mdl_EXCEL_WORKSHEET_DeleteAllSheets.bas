Attribute VB_Name = "mdl_DeleteAllSheets"
Option Explicit

Public Sub DeleteAllSheets(Optional TargetBook As Workbook)

    Dim wrkSht As Worksheet

    If TargetBook Is Nothing Then
        TargetBook = ThisWorkbook
    End If
    
    Application.DisplayAlerts = False
    For Each wrkSht In TargetBook.Worksheets
        wrkSht.Delete
    Next wrkSht
    Application.DisplayAlerts = True

End Sub
