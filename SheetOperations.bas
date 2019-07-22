Attribute VB_Name = "SheetOperations"
' GetColumnNumberByName
'       Pretty self-explanatory, really...
'       Returns 0 if it can't find the column.
'       *It's case-sensitive!*
'
'       Parameters:
'               sName: Column name.
'               [nRowNo]: Row number where the column names are located.
'               [nColNo]: Number of the first column to check.
Public Function GetColumnNumberByName(sName As String, _
    Optional nRowNo As Integer = 1, Optional nColNo As Integer = 1) As Integer

        Dim j As Integer
        Dim sVal As String

        GetColumnNumberByName = 0

        j = nColNo
        Do While True
                sVal = Cells(nRowNo, j)

                If sVal = "" Then
                        Exit Do
                ElseIf sVal = sName Then
                        GetColumnNumberByName = j
                        Exit Do
                End If

                j = j + 1
        Loop
End Function




' GetLastRow
'       Returns the number of the last row with content for a given column in the
'       current worksheet.
Public Function GetLastRow(Optional nColNo As Integer = 1) As Integer
    Dim sVal As String

    GetLastRow = 1

    Do While True
        sVal = Cells(GetLastRow, nColNo)

        If sVal = "" Then
            Exit Do
        End If
        GetLastRow = GetLastRow + 1
    Loop
End Function




' BlankColumns
'       Sets the values of the specified columns to "", or to a specified value.
'
'       Parameters:
'               sColumns: String containing the columns whose values are to be changed,
'                       separated by commas. Column names may contain special characters
'                       or spaces.
'               nStartRow: Row from which to start changing values.
'               nLastRow: Row at which to stop.
'               [sValue]: Value to use. Defaults to "".
Public Sub BlankColumns(sColumns As String, nStartRow As Integer, _
    nLastRow As Integer, Optional sValue As String = "")

    Dim asColumns As Variant, sColumn As Variant

    asColumns = Split(sColumns, ",")

    For Each sColumn In asColumns
        nCol = GetColumnNumberByName(CStr(sColumn))
        If nCol <> 0 Then
            Range(Cells(nStartRow, nCol), Cells(nLastRow, nCol)).Value = sValue
        End If
    Next
End Sub




' ApplyFilter
'       Applies a filter to a specified range in the current worksheet.
'
'       Parameters:
'               sFields: String containing the names of the columns to filter,
'                       separated by commas.
'               sCriteria: String containing the criteria for each field. Only supports
'                       one per field.
'
'               *Must use at least one of these*
'               [sRng]: Range, in A1-style notation.
'               [oRng]: Range, as Range.
Public Sub ApplyFilter(sFields As String, sCriteria As String, _
    Optional sRng As String = "", Optional oRng As Range)

    Dim asFields As Variant, asCriteria As Variant
    Dim i As Integer, nColNo As Integer

    asFields = Split(sFields, ",")
    asCriteria = Split(sCriteria, ",")

    If UBound(asFields) <> UBound(asCriteria) Then
        Debug.Print ("WARNING: ApplyFilter: Incorrect number of criteria.")
        Exit Sub
    End If

    If sRng <> "" Then
        oRng = Range(sRng)
    End If

    With oRng
        For i = 0 To UBound(asFields)
            nColNo = GetColumnNumberByName(CStr(asFields(i)))
            If nColNo <> 0 Then
                .AutoFilter Field:=nColNo, Criteria1:=asCriteria(i)
            End If
        Next
    End With
End Sub
