Sub NuevaAccion()
'
' NuevaAccion Macro
'
' Keyboard Shortcut: Ctrl+q
'
    If Selection.count = 1 Then
        If ActiveSheet.Name = "SEGUIMIENTO" Then
            If Range("A" & Selection.Row).Value <> "" And Range("A" & Selection.Row).Value <> "ID_CONDICION" Then
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = Range("A" & Selection.Row).Value
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = Range("B" & Selection.Row).Value
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FILA_ACCION") = Selection.Row
                newAccion.Show (0)
            End If
        End If
    End If
End Sub

Sub ModificarAccion()
'
' ModificarAccion Macro
'
' Keyboard Shortcut: Ctrl+e
'
    If Selection.count = 1 Then
        If ActiveSheet.Name = "SEGUIMIENTO" Then
            If Range("A" & Selection.Row).Value <> "" And Range("A" & Selection.Row).Value <> "ID_CONDICION" Then
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = Range("A" & Selection.Row).Value
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = Range("B" & Selection.Row).Value
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = Range("C" & Selection.Row).Value
                If NUEVA_ACCION Then
                    busqSeguimiento.Show (0)
                Else
                    busqCondicion.Show (0)
                End If
            End If
        End If
    End If
End Sub

Sub LimpiarTablas()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("dataSet"), ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("dataSet").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("dataSetTemp"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("dataSetTemp").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP2).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP2).Range("dataSetTemp2"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP2).Range("dataSetTemp2").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP3).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP3).Range("dataSetTemp3"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP3).Range("dataSetTemp3").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range("dataSetTemp4"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range("dataSetTemp4").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP6).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP6).Range("dataSetTemp6"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP6).Range("dataSetTemp6").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range("dataSetTemp9"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range("dataSetTemp9").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("dataSetTemp11"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("dataSetTemp11").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range("dataSetTemp13"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range("dataSetTemp13").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP14).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP14).Range("dataSetTemp14"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP14).Range("dataSetTemp14").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range("dataSetTemp16"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range("dataSetTemp16").End(xlDown)).ClearContents
End Sub
