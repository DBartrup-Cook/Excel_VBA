Attribute VB_Name = "mdl_PrepRawData"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : DeleteColumns
' Author    : Darren Bartrup-Cook
' Date      : 17/12/2013
' Purpose   : Deletes the named columns on either the active worksheet or the specified worksheet.
' To Use    : DeleteColumns "GD:GL,FY:GB,FU:FW,FS", "Lettings_AppSurvey4WKS"
'             DeleteColumns "B"
'---------------------------------------------------------------------------------------
Public Sub DeleteColumns(ColumnLetters As String, Optional SheetName As String = "")

    Dim wrkSht As Worksheet
    Dim cols As Variant
    Dim rRange As Range
    Dim x As Long

    On Error GoTo ERROR_HANDLER

    '//Resolve the sheet name.
    If SheetName = "" Then
        Set wrkSht = ActiveSheet
    Else
        Set wrkSht = ThisWorkbook.Worksheets(SheetName)
    End If
    
    If InStr(ColumnLetters, ",") > 0 Then
    '//If the passed argument contains commas then it's multiple ranges.
        cols = Split(ColumnLetters, ",")
        For x = LBound(cols) To UBound(cols)
            If rRange Is Nothing Then
                Set rRange = wrkSht.Columns(cols(x))
            Else
                Set rRange = Union(rRange, wrkSht.Columns(cols(x)))
            End If
        Next x
    Else
    '//Otherwise it's a single range.
        Set rRange = wrkSht.Columns(ColumnLetters)
    End If
    
    '//Delete the appropriate columns.
    If Not rRange Is Nothing Then rRange.Delete
    
    On Error GoTo 0
    Exit Sub

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure DeleteColumns."
            Err.Clear
            Application.EnableEvents = True
    End Select
    
End Sub


Public Sub Test()

    DeleteColumns "GD:GL,FY:GB,FU:FW,FS,FN:FQ,FA:FL,BW:EY,AS:BP,B:AQ", "Lettings_AppSurvey4WKS"
    DeleteColumns "GI:GK,FU:FY,FS,FN:FQ,FA:FL,BW:EZ,AS:BO,B:AQ", "Lettings_LLSurvey4WKS"
    

End Sub

