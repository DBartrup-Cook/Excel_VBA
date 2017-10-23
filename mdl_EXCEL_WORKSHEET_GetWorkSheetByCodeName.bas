Attribute VB_Name = "mdl_GetWorkSheetByCodeName"
Option Explicit

'----------------------------------------------------------------------------------
' Procedure : GetWorkSheet
' Author    : Darren Bartrup-Cook
' Date      : 03/03/2016
' Purpose   : Returns a reference to a worksheet given the codename.
'             Not useful if the worksheet is in ThisWorkbook, but allows referencing
'             to worksheets in other workbooks by codename.
'-----------------------------------------------------------------------------------
Public Function GetWorkSheet(sCodeName As String, Optional wrkBook As Workbook) As Worksheet

    Dim wrkSht As Worksheet

    If wrkBook Is Nothing Then
        Set wrkBook = ThisWorkbook
    End If
    
    For Each wrkSht In wrkBook.Worksheets
        If wrkSht.CodeName = sCodeName Then
            Set GetWorkSheet = wrkSht
            Exit For
        End If
    Next wrkSht

End Function
