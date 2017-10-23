Attribute VB_Name = "mdl_Filtering"
Option Explicit

'--------------------------------------------------------------------------------------
' Procedure : SetAutoFilter
' Author    : Darren Bartrup-Cook
' Date      : 15/01/2014
' Purpose   : Applies an auto-filter to range specified by rData.
'---------------------------------------------------------------------------------------
Public Sub SetAutoFilter(Optional rData As Range, Optional bState As Boolean = True)

    Dim wrkSht As Worksheet

    On Error GoTo ERROR_HANDLER

    Set wrkSht = rData.Parent 'Get reference to worksheet.

    If bState Then
        'Turn auto-filter on, show all columns and show all data.
        If wrkSht.AutoFilterMode Then wrkSht.AutoFilterMode = False 'Turn off auto-filter.
        wrkSht.Cells.EntireColumn.Hidden = False 'Unhide all columns.
        If Not wrkSht.AutoFilterMode Then rData.AutoFilter 'Turn on the auto-filter.
        If wrkSht.FilterMode Then wrkSht.ShowAllData 'Clear any filters applied.
    Else
        'Turn auto-filter off and show all all columns.
        If wrkSht.AutoFilterMode Then wrkSht.AutoFilterMode = False 'Turn off auto-filter.
        wrkSht.Cells.EntireColumn.Hidden = False 'Unhide all columns.
    End If

    On Error GoTo 0
    Exit Sub

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure SetAutoFilter." & vbCr & vbCr & _
                    "Please contact the spreadsheet designer."
            Err.Clear
            Application.EnableEvents = True
    End Select
    

End Sub


'--------------------------------------------------------------------------------------
' Procedure : SetFilter
' Author    : Darren Bartrup-Cook
' Date      : 15/01/2014
' Purpose   : Applies a filter to a range.
'             ParamArray is an array of filter arguments containing either 2 or 4 elements.
'        e.g. SetFilter rMainData, Array(7, "<>"), Array(13, "On Hold", "=", xlOr)
'---------------------------------------------------------------------------------------
Public Sub SetFilter(rDataRange As Range, ParamArray sFilters())
    Dim wrkSht As Worksheet
    Dim x As Long
    
    Set wrkSht = rDataRange.Parent 'Get reference to worksheet.
    
    With wrkSht
        If Not .AutoFilterMode Then rDataRange.AutoFilter 'Turn on the auto-filter
        If .FilterMode Then .ShowAllData 'Clear any filters applied.
    End With
    With rDataRange
        For x = LBound(sFilters) To UBound(sFilters)
            Select Case UBound(sFilters(x))
                Case 1 '2 elements to array.
                    .AutoFilter Field:=sFilters(x)(0), Criteria1:=sFilters(x)(1)
                Case 3 '4 elements to array.
                    .AutoFilter Field:=sFilters(x)(0), Criteria1:=sFilters(x)(1), _
                        Operator:=sFilters(x)(3), Criteria2:=sFilters(x)(2)
            End Select
        Next x
    End With
    
End Sub
