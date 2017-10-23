Attribute VB_Name = "mdl_FindPrecedents"
Option Explicit

Sub FindPrecedents()
    ' written by Bill Manville
    ' With edits from PaulS
    ' this procedure finds the cells which are the direct precedents of the active cell
    Dim rLast As Range, iLinkNum As Integer, iArrowNum As Integer
    Dim stMsg As String
    Dim bNewArrow As Boolean
    Application.ScreenUpdating = False
    ActiveCell.ShowPrecedents
    Set rLast = ActiveCell
    iArrowNum = 1
    iLinkNum = 1
    bNewArrow = True
    Do
        Do
            Application.Goto rLast
            On Error Resume Next
            ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=iArrowNum, LinkNumber:=iLinkNum
            If Err.Number > 0 Then Exit Do
            On Error GoTo 0
            If rLast.Address(external:=True) = ActiveCell.Address(external:=True) Then Exit Do
            bNewArrow = False
            If rLast.Worksheet.Parent.Name = ActiveCell.Worksheet.Parent.Name Then
                If rLast.Worksheet.Name = ActiveCell.Parent.Name Then
                    ' local
                    stMsg = stMsg & vbNewLine & Selection.Address
                Else
                    stMsg = stMsg & vbNewLine & "'" & Selection.Parent.Name & "'!" & Selection.Address
                End If
            Else
                ' external
                stMsg = stMsg & vbNewLine & Selection.Address(external:=True)
            End If
            iLinkNum = iLinkNum + 1  ' try another link
        Loop
        If bNewArrow Then Exit Do
        iLinkNum = 1
        bNewArrow = True
        iArrowNum = iArrowNum + 1  'try another arrow
    Loop
    rLast.Parent.ClearArrows
    Application.Goto rLast
    MsgBox "Precedents are" & stMsg
    Exit Sub
End Sub

