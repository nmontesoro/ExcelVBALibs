Attribute VB_Name = "Calculos_Sueldos"
' CalcVacaciones
'    Devuelve la cantidad de d�as de vacaciones (al a�o) que le corresponden
'    al empleado seg�n su antig�edad
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
'   Seg�n las fechas de ingreso y egreso de un empleado, devuelve la cantidad
'   de d�as de aguinaldo que le corresponden en este �ltimo per�odo, �til
'   para c�lculo de liquidaciones finales.
Public Function CalcDiasSAC(dIng As Date, dEgr As Date) As Integer
    ' Si dEgr > MedioA�o:
    '   InicioSAC = Max(dIng, MedioA�o)
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
