Attribute VB_Name = "mdl_SelectTop10FilteredRows"
Option Explicit

Sub Test()

    Dim wrkSht As Worksheet
    Dim rContentsTot As Range
    Dim rChToKey As Range
    Dim rCTD As Range
    Dim tbl As ListObject
    Dim lTblFirstCol As Long
    Dim rVisible As Range
    Dim rRow As Range
    Dim rSelection As Range
    Dim lCounter As Long
    
    Set wrkSht = ThisWorkbook.Worksheets("Data")
    Set tbl = wrkSht.ListObjects("tb_DATA")
    
    With tbl
    
        lTblFirstCol = .HeaderRowRange.Column
    
        Set rContentsTot = .HeaderRowRange.Find("Contents Total")
        Set rChToKey = .HeaderRowRange.Find("Ch To key")
        Set rCTD = .HeaderRowRange.Find("Logistics/CTD")
        
        'Only continue if all columns have been found.
        If Not rContentsTot Is Nothing And Not rChToKey Is Nothing And Not rCTD Is Nothing Then
    
            'Turn on table autofilter if it's not on, or show all data if it is.
            If .AutoFilter Is Nothing Then
                .Range.AutoFilter
            Else
                .AutoFilter.ShowAllData
            End If
            
            'Filter as required.  Field:=1 is first column in table, Field:=5 is the fifth.
            'NB - Is there a better way to return the correct column within the table?
            '     This works, but feel I should be able to equate the found column number to
            '     a column number within the table.
            With .Range
                .AutoFilter Field:=rCTD.Column - lTblFirstCol + 1, Criteria1:="3"
                .AutoFilter Field:=rChToKey.Column - lTblFirstCol + 1, Criteria1:="6"
                .AutoFilter Field:=rContentsTot.Column - lTblFirstCol + 1, Criteria1:="Rebill"
            End With
            
            '********************************
            'Not sure why I added this next bit of code - it gets the first 10 rows
            'of filtered data.  Think I must've combined two posts into one in my
            'mind overnight (started answering this question yesterday).
            'Anyway.... it's there, it does stuff, so I'm leaving it in
            '********************************
            
            'Now to grab the top 10 visible rows in the table.
            Set rVisible = .DataBodyRange.SpecialCells(xlCellTypeVisible)
            For Each rRow In rVisible.Rows
                If lCounter < 10 Then
                    If lCounter = 0 Then
                        Set rSelection = rRow
                    Else
                        Set rSelection = Application.Union(rSelection, rRow)
                    End If
                    lCounter = lCounter + 1
                Else
                    Exit For
                End If
            Next rRow
            
            'Remove filter and select the top 10 rows that appeared in the filter.
            .AutoFilter.ShowAllData
            rSelection.Select
            
            '******************
            'End of extra code that you may or may not need
            '******************
            
        Else
            'Raise error as not all columns have been found.
        End If
        
    End With
    
End Sub

