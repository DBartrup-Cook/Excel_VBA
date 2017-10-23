Attribute VB_Name = "mdl_HasValidation"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : HasValidation
' Author    : Darren Bartrup-Cook
' Date      : 13/12/2013
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function HasValidation(Target As Range) As Variant

    On Error Resume Next
    If Target.Cells.Count = 1 Then
        HasValidation = Not (IsEmpty(Target.Validation.Formula1))
    Else
        HasValidation = CVErr(xlErrValue)
    End If
    On Error GoTo 0

End Function


