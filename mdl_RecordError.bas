Attribute VB_Name = "mdl_RecordError"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : RecordError
' Purpose   : Save a record of any error in the 'Error.log' stored in the same
'             directory as the backend.
'---------------------------------------------------------------------------------------
Public Sub RecordError(ErrMsg As String)
    Dim lFile As Long
    lFile = FreeFile
    
    Open ThisWorkbook.Path & "\Error.log" For Append As #lFile
    Print #lFile, ErrMsg & vbNewLine & Now() & " | " & GetSystemNames(ComputerName) & "| " & _
            GetSystemNames(ComputerUser) & vbNewLine & String(52, 175)
    Close #lFile
    
End Sub

