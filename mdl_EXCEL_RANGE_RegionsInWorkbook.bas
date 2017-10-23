Attribute VB_Name = "mdl_RegionsInWorkbook"
Option Explicit

Sub Test()
    Dim aLists  As Variant
    Dim aLists1 As Variant
    '//Find lists in a different workbook.
    aLists = FindRegionsInWorkbook(Workbooks("Bearwood Data Cleanse.xls"))
    '//Find lists in the this workbook.
'    aLists1 = FindRegionsInWorkbook(ThisWorkbook)
    Debug.Assert False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindRegionsInWorkbook
' Author    : Zack Barresse (MVP), Oregon, USA. (http://www.mrexcel.com/forum/showthread.php?t=309052)
' Date      : 20/03/2008
' Purpose   : Returns each region in each worksheet within the workbook in the 'sRegion' variable.
'---------------------------------------------------------------------------------------
Public Function FindRegionsInWorkbook(wrkBk As Workbook) As Variant
    Dim ws As Worksheet, rRegion As Range, sRegion As String, sCheck As String
    Dim sAddys As String, arrAddys() As String, aRegions() As Variant
    Dim iCnt As Long, i As Long, j As Long
    '//Cycle through each worksheet in workbook.
    j = 0
    For Each ws In wrkBk.Worksheets
        sAddys = vbNullString
        sRegion = vbNullString
        On Error Resume Next
        '//Find all ranges of constant & formula valies in worksheet.
        sAddys = ws.Cells.SpecialCells(xlCellTypeConstants, 23).Address(0, 0) & ","
        sAddys = sAddys & ws.Cells.SpecialCells(xlCellTypeFormulas, 23).Address(0, 0)
        If Right(sAddys, 1) = "," Then sAddys = Left(sAddys, Len(sAddys) - 1)
        On Error GoTo 0
        If sAddys = vbNullString Then GoTo SkipWs
        '//Put each seperate range into an array.
        If InStr(1, sAddys, ",") = 0 Then
            ReDim arrAddys(0 To 0)
            arrAddys(0) = "'" & ws.Name & "'!" & sAddys
        Else
            arrAddys = Split(sAddys, ",")
            For i = LBound(arrAddys) To UBound(arrAddys)
                arrAddys(i) = "'" & ws.Name & "'!" & arrAddys(i)
            Next i
        End If
        '//Place region that range sits in into sRegion (if not already in there).
        For i = LBound(arrAddys) To UBound(arrAddys)
            If InStr(1, sRegion, ws.Range(arrAddys(i)).CurrentRegion.Address(0, 0)) = 0 Then
                sRegion = sRegion & ws.Range(arrAddys(i)).CurrentRegion.Address(0, 0) & "," '*** no sheet
                sCheck = Right(arrAddys(i), Len(arrAddys(i)) - InStr(1, arrAddys(i), "!"))
                ReDim Preserve aRegions(0 To j)
                aRegions(j) = Left(arrAddys(i), InStr(1, arrAddys(i), "!") - 1) & "!" & ws.Range(sCheck).CurrentRegion.Address(0, 0)
                j = j + 1
            End If
        Next i
SkipWs:
    Next ws
    On Error GoTo ErrHandle
    FindRegionsInWorkbook = aRegions
    Exit Function
ErrHandle:
    'things you might want done if no lists were found...
End Function
