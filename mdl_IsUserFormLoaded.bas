Attribute VB_Name = "mdl_IsUserFormLoaded"
Option Explicit

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
    For Each UForm In VBA.UserForms
        IsUserFormLoaded = UForm.Name = UFName
        If IsUserFormLoaded Then
            Exit For
        End If
    Next
End Function
