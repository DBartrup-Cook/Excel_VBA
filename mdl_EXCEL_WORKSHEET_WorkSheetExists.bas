Attribute VB_Name = "mdl_WorkSheetExists"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : WorkSheetExists
' Author    : Darren Bartrup-Cook
' Date      : 21/01/2014
' Purpose   : Attempts to set a reference to the worksheet, returns False if it fails.
'---------------------------------------------------------------------------------------
Public Function WorkSheetExists(SheetName As String, Optional WrkBk As Workbook) As Boolean
    Dim wrkSht As Worksheet
    
    If WrkBk Is Nothing Then
        Set WrkBk = ThisWorkbook
    End If
    
    On Error Resume Next
        Set wrkSht = WrkBk.Worksheets(SheetName)
        WorkSheetExists = (Err.Number = 0)
        Set wrkSht = Nothing
    On Error GoTo 0
End Function
