Attribute VB_Name = "RegularExpressions"
' GetRegexGroups:
'   Returns an array containing all the matches. If there are none, the array
'   will be empty.
Public Function GetRegexGroups(sInput As String, sPattern As String, _
    Optional lGlobal As Boolean = True, Optional lIgnoreCase As Boolean = True, _
    Optional lMultiLine As Boolean = True)

    Dim oRegex As New RegExp
    Dim oRegexResults As Object
    Dim asMatches() As String

    With oRegex
        .Global = lGlobal
        .IgnoreCase = lIgnoreCase
        .MultiLine = lMultiLine
        .Pattern = sPattern

        Set oRegexResults = .Execute(sInput)
    End With

    If oRegexResults.Count <> 0 Then
        ReDim asMatches(oRegexResults.Count - 1)
    End If

    For i = 0 To oRegexResults.Count - 1
        asMatches(i) = oRegexResults(i).Value
    Next

    GetRegexGroups = asMatches
End Function




' Regex
'   Wrapper for GetRegexGroups. Returns matches as "match_1|match_2|...|match_n",
'   or an empty string if there are none. Can be used as an Excel formula.
Public Function Regex(sInput As String, sPattern As String, _
    Optional lIgnoreCase As Boolean = True, Optional lGlobal = True, _
    Optional lMultiLine = True) As String

    Dim asMatches As Variant

    asMatches = GetRegexGroups(sInput, sPattern, (lGlobal), (lIgnoreCase), _
        (lMultiLine))

    Regex = ""
    For Each sMatch In asMatches
        Regex = Regex + "|" + sMatch
    Next

    Regex = Replace(Regex, "|", "", 1, 1)
End Function



