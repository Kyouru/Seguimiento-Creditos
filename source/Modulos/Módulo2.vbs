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


