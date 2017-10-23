Attribute VB_Name = "mdl_WorkBookIsOpen"
Option Explicit

'----------------------------------------------------------------------------------
' Procedure : WorkBookIsOpen
' Author    : Darren Bartrup-Cook
' Date      : 22/10/2014
' Purpose   : Returns TRUE if the named file is open.
'-----------------------------------------------------------------------------------
Public Function WorkBookIsOpen(FullFilePath As String) As Boolean
    
    Dim ff As Long

    On Error Resume Next
    
    ff = FreeFile()
    Open FullFilePath For Input Lock Read As #ff
    Close ff
    WorkBookIsOpen = (Err.Number <> 0)
    
    On Error GoTo 0

End Function
