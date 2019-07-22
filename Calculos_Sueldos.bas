Attribute VB_Name = "Calculos_Sueldos"
' CalcVacaciones
'    Devuelve la cantidad de días de vacaciones (al año) que le corresponden
'    al empleado según su antigüedad
Public Function CalcVacaciones(nAntiguedad As Integer) As Integer
    If nAntiguedad < 0 Then
        CalcVacaciones = 0
    ElseIf nAntiguedad < 6 * 30 Then
        CalcVacaciones = Int(nAntiguedad / 20)
    ElseIf nAntiguedad < 5 * 12 * 30 Then
        CalcVacaciones = 14
    ElseIf nAntiguedad < 10 * 12 * 30 Then
        CalcVacaciones = 21
    ElseIf nAntiguedad < 20 * 12 * 30 Then
        CalcVacaciones = 28
    Else
        CalcVacaciones = 35
    End If
End Function




' CalcDiasSAC
'   Según las fechas de ingreso y egreso de un empleado, devuelve la cantidad
'   de días de aguinaldo que le corresponden en este último período, útil
'   para cálculo de liquidaciones finales.
Public Function CalcDiasSAC(dIng As Date, dEgr As Date) As Integer
    ' Si dEgr > MedioAño:
    '   InicioSAC = Max(dIng, MedioAño)
    ' Si no:
    '   InicioSAC = Max(PrimerDia, dIng)

    ' Devuelvo int: dEgr - InicioSAC

    Dim dInicioSAC As Date, dMidYear As Date, dFDOfYear As Date

    dFDOfYear = DateSerial(Year(dEgr), 1, 1)
    dMidYear = DateSerial(Year(dEgr), 6, 30)

    If dEgr > dMidYear Then
        dInicioSAC = WorksheetFunction.Max(dIng, dMidYear)
    Else
        dInicioSAC = WorksheetFunction.Max(dFDOfYear, dIng)
    End If

    CalcDiasSAC = dEgr - dInicioSAC

End Function
