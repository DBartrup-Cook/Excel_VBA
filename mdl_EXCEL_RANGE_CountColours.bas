Attribute VB_Name = "mdl_UDF_CountColours"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : CountColours
' Author    : Darren Bartrup-Cook
' Date      : 13/11/2013
' Purpose   : Counts the background colour of a selected range.
'             The first argument is the range to be counted, the second is either a range
'             or a number identifying the colour to count.
' To Use    : =COUNTCOLOURS(A1:A5,A6) or =COUNTCOLOURS(A1:A5,34)
'---------------------------------------------------------------------------------------
Public Function CountColours(Target As Range, Colour As Variant) As Long

    Dim rCell As Range
    Dim lColour As Long
    Dim x As Long
    
    Application.Volatile False
    
    Select Case TypeName(Colour)
        Case "Range"
            If Colour.Cells.Count = 1 Then
                lColour = Colour.Interior.ColorIndex
            Else
                CountColours = CVErr(xlErrValue)
            End If
        Case "Double"
            lColour = Colour
        Case Else
            CountColours = CVErr(xlErrValue)
    End Select
    
    For Each rCell In Target
        If rCell.Interior.ColorIndex = lColour Then
            x = x + 1
        End If
    Next rCell
    
    CountColours = x

End Function
