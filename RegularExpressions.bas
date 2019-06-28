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
    Else
        ReDim asMatches(0)
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

    asMatches = ExecuteRegex(sInput, sPattern, (lGlobal), (lIgnoreCase), _
        (lMultiLine))

    Regex = ""
    For Each sMatch In asMatches
        Regex = Regex + "|" + sMatch
    Next

    Regex = Replace(Regex, "|", "", 1, 1)
End Function




Public Function ExecuteRegex(sInput As String, sPattern As String, _
    Optional lGlobal As Boolean = True, Optional lIgnoreCase As Boolean = True, _
    Optional lMultiLine As Boolean = True) As Variant

    Dim oRegex As New RegExp
    Dim oRegexMatches As Object
    Dim nItems As Integer
    Dim asMatches() As String

    With oRegex
        .Global = lGlobal
        .IgnoreCase = lIgnoreCase
        .MultiLine = lMultiLine
        .Pattern = sPattern

        Set oRegexMatches = .Execute(sInput)
    End With

    nItems = 0
    nId = 0
    For Each oMatch In oRegexMatches
        If oMatch.SubMatches.Count = 0 Then
            nItems = nItems + 1
        Else
            nItems = nItems + oMatch.SubMatches.Count
        End If
    Next

    If nItems <> 0 Then
        ReDim asMatches(nItems - 1)
        For Each oMatch In oRegexMatches
            If oMatch.SubMatches.Count = 0 Then
                asMatches(nId) = oMatch.Value
            Else
                For Each sSubMatch In oMatch.SubMatches
                    asMatches(nId) = sSubMatch
                    nId = nId + 1
                Next
                nId = nId - 1
            End If

            nId = nId + 1
        Next
    End If

    ExecuteRegex = asMatches
End Function
