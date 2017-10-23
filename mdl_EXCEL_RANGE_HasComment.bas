Attribute VB_Name = "mdl_HasComment"
Option Explicit

'----------------------------------------------------------------------------------
' Procedure : HasComment
' Author    : Darren Bartrup-Cook
' Date      : 25/02/2016
' Purpose   : Returns TRUE/FALSE if based on cell containing comment.
'             This will return TRUE if a comment has been created, but no text added.
'-----------------------------------------------------------------------------------
Public Function HasComment(Target As Range) As Boolean

    On Error GoTo ERROR_HANDLER

    If Target.Cells.Count = 1 Then
        With Target
            HasComment = Not .Comment Is Nothing
        End With
    Else
        Err.Raise 513, "HasComment()", "Argument must reference single cell."
    End If

    On Error GoTo 0
    Exit Function

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure Module1.HasComment."
            Err.Clear
            Application.EnableEvents = True
    End Select
    

End Function
