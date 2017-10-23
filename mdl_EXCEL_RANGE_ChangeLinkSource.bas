Attribute VB_Name = "mdl_ChangeLinkSource"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : ChangeLink
' Author    : Darren Bartrup-Cook
' Date      : 09/01/2014
' Purpose   : Updates external links.
'---------------------------------------------------------------------------------------
Public Sub ChangeLink(FromLink As Workbook, ToLink As Workbook)
    Dim arrLinks As Variant
    
    On Error GoTo ERROR_HANDLER

    arrLinks = ToLink.LinkSources(xlExcelLinks)
    If IsEmpty(arrLinks) Then
        MsgBox ToLink.Name & " does not contain any links.", vbOKOnly, "No External Links"
        Exit Sub
    End If
    
    ToLink.ChangeLink Name:=FromLink.Name, _
                   NewName:=ToLink.FullName, Type:=xlLinkTypeExcelLinks

    On Error GoTo 0
    Exit Sub

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure ChangeLink." & vbCr & vbCr & _
                    "Please contact the spreadsheet designer."
            Err.Clear
            Application.EnableEvents = True
    End Select
                   
End Sub
