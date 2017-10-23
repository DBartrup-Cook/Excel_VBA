Attribute VB_Name = "mdl_Worksheet_FilterOnDblClick"
Option Explicit

Private rCellClicked As Range

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    With Sheet1
        If .FilterMode Then
            .Range("$P$2:$P$672").SpecialCells(xlCellTypeVisible) = "DONE"
            .ShowAllData
            rCellClicked.Select
        Else
            .Range("$A$1:$T$672").AutoFilter Field:=Target.Column, Criteria1:=Target.Value
            Set rCellClicked = Target
        End If
    End With

End Sub
