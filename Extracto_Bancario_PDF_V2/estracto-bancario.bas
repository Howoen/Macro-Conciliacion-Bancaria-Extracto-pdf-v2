Attribute VB_Name = "M�dulo2"
Sub SepararYCompararValores()

Dim ws As Worksheet, hojaD�bito As Worksheet, hojaCr�dito As Worksheet
Dim celda As Range, valor As Variant
Dim filaPos As Long, filaNeg As Long
Dim i As Long
Dim ultimaFilaE As Long, ultimaFilaA As Long, ultimaFilaB As Long
Dim listaPositivos As Range, listaNegativos As Range
Dim debitos As Range, creditos As Range

Set ws = ThisWorkbook.Sheets("Hoja1")
filaPos = 2
filaNeg = 2

' Limpiar columnas H e I antes de copiar
ws.Range("H2:H10000").ClearContents
ws.Range("I2:I10000").ClearContents

' Detectar �ltima fila con datos en columnas E, A y B
ultimaFilaE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
ultimaFilaA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
ultimaFilaB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

' Separar positivos y negativos desde E2 hasta �ltima fila de E
For Each celda In ws.Range("E2:E" & ultimaFilaE)
    valor = celda.Value
    If IsNumeric(valor) Then
        If valor > 0 Then
            ws.Cells(filaPos, "H").Value = valor
            filaPos = filaPos + 1
        ElseIf valor < 0 Then
            ws.Cells(filaNeg, "I").Value = Abs(valor) ' Convertir a positivo con Abs
            filaNeg = filaNeg + 1
        End If
    End If
Next celda

' Definir rangos din�micos
Set listaPositivos = ws.Range("H2:H" & filaPos - 1)
Set listaNegativos = ws.Range("I2:I" & filaNeg - 1)
Set debitos = ws.Range("A2:A" & ultimaFilaA)
Set creditos = ws.Range("B2:B" & ultimaFilaB)

' Eliminar hojas anteriores si existen
On Error Resume Next
Application.DisplayAlerts = False
ThisWorkbook.Sheets("NoEnD�bitos").Delete
ThisWorkbook.Sheets("NoEnCr�ditos").Delete
Application.DisplayAlerts = True
On Error GoTo 0

' Crear hojas nuevas
Set hojaD�bito = ThisWorkbook.Sheets.Add(After:=ws)
hojaD�bito.Name = "NoEnD�bitos"
Set hojaCr�dito = ThisWorkbook.Sheets.Add(After:=hojaD�bito)
hojaCr�dito.Name = "NoEnCr�ditos"

' Comparar positivos con d�bitos
i = 1
For Each celda In listaPositivos
    If WorksheetFunction.CountIf(debitos, celda.Value) = 0 Then
        hojaD�bito.Cells(i, 1).Value = celda.Value
        i = i + 1
    End If
Next celda

' Comparar negativos convertidos con cr�ditos
i = 1
For Each celda In listaNegativos
    If WorksheetFunction.CountIf(creditos, celda.Value) = 0 Then
        hojaCr�dito.Cells(i, 1).Value = celda.Value
        i = i + 1
    End If
Next celda

MsgBox "Proceso finalizado con �xito.", vbInformation


End Sub
