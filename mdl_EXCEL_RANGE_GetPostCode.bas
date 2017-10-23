Attribute VB_Name = "mdl_GetPostCode"
Option Explicit

Public Function GetPostCode(AddressRange As Range) As Variant

    Dim rCell As Range
    Dim sAddressString As String

    For Each rCell In AddressRange
        sAddressString = sAddressString & " " & rCell.Value
    Next rCell
    sAddressString = Trim(sAddressString)

    GetPostCode = ValidatePostCode(sAddressString)

End Function

Public Function ValidatePostCode(strData As String) As Variant

    Dim RE As Object, REMatches As Object

    Dim UKPostCode As String

    'Pattern could probably be improved.
    UKPostCode = "(?:(?:A[BL]|B[ABDHLNRST]?|C[ABFHMORTVW]|D[ADEGHLNTY]|E[CHNX]?|F[KY]|G[LUY]?|" _
                & "H[ADGPRSUX]|I[GMPV]|JE|K[ATWY]|L[ADELNSU]?|M[EKL]?|N[EGNPRW]?|O[LX]|P[AEHLOR]|R[GHM]|S[AEGKLMNOPRSTWY]?|" _
                & "T[ADFNQRSW]|UB|W[ACDFNRSV]?|YO|ZE)\d(?:\d|[A-Z])? \d[A-Z]{2})"

    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .MultiLine = False
        .Global = False
        .IgnoreCase = True
        .Pattern = UKPostCode
    End With

    Set REMatches = RE.Execute(strData)
    If REMatches.Count = 0 Then
        ValidatePostCode = CVErr(xlErrValue)
    Else
        ValidatePostCode = REMatches(0)
    End If

End Function

