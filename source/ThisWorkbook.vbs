Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    
            'ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
            
            'ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("dataSet"), ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("dataSet").End(xlDown)).ClearContents
            
            'ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
End Sub

Private Sub Workbook_Open()

    'Desbloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD

    If MATRIZ Then
        Hoja1.Range("TITULO") = "SEGUIMIENTO MATRIZ"
        Hoja1.btActualizarDesembolso.Visible = True
        Hoja1.btMantenimientoLista.Visible = True
        Hoja1.btNuevaAccion.Caption = "Nueva Accion"
        Hoja1.btMantenimiento.Caption = "Mantenimiento DB"
    Else
        If NUEVA_ACCION Then
            Hoja1.Range("TITULO") = "SEGUIMIENTO"
            Hoja1.btNuevaAccion.Caption = "Nueva Accion"
            Hoja1.btMantenimiento.Caption = "Mantenimiento DB"
        Else
            Hoja1.Range("TITULO") = "INGRESO CONDICIONES"
            Hoja1.btNuevaAccion.Caption = "Ver Accion"
            Hoja1.btMantenimiento.Caption = "Nueva Condicion"
        End If
        Hoja1.btActualizarDesembolso.Visible = False
        Hoja1.btMantenimientoLista.Visible = False
    End If
    
    If POSTERGAR Then
        Hoja1.btPostergar.Visible = True
        Hoja1.btReporteP.Visible = True
    Else
        Hoja1.btPostergar.Visible = False
        Hoja1.btReporteP.Visible = False
    End If
    
    Hoja1.Columns("A:C").NumberFormat = "General"
    Hoja1.Range("D1:M1").NumberFormat = "General"
    
    'Bloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

End Sub
