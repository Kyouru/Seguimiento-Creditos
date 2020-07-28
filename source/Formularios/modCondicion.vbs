Private Sub btCancelar_Click()
    Unload Me
    busqCondicion.Show (0)
End Sub

Private Sub btGuardar_Click()
    'Comprobar que la Condicion este seleccionada correctamente
    If cmbTipo.ListIndex <> -1 Then
        'Comprobar que el Detalle de la Condicion no sea en blanco
        If tbDetalle.Text <> "" Then
            'Comprobar que alguien apruebe la Condicion
            If cmbAprobado.Text <> "" Then
                strSQL = "UPDATE DB_CONDICION SET ID_TIPO_CONDICION_FK = " & cmbTipo.List(cmbTipo.ListIndex, 1) & _
                ", DETALLE = '" & Replace(tbDetalle.Text, "'", "''") & "', APROBADO_POR = '" & cmbAprobado.Text & _
                "' WHERE ID_CONDICION = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION")
                
                OpenDB
                On Error GoTo Handle:
                cnn.Execute (strSQL)
                closeRS
                
                'Actualiza Lista de Condiciones
                busqCondicion.ActualizarHoja
                busqCondicion.ActualizarLista
                
                'Desbloquea Hoja
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
            
                With ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO)
                    Set total = .Range("A:A")
                    Set r = total.Cells.Find(What:=.Range("ID_CONDICION"), LookAt:=xlWhole)
                    If Not r Is Nothing Then
                        rowNumber = r.Row
                    Else
                        rowNumber = .Range("FILA_ACCION")
                    End If
                    .Cells(rowNumber, .Range("FECHA_ACCION").Column - 1) = tbDetalle.Text
                    .Cells(rowNumber, .Range("FECHA_ACCION").Column - 2) = cmbTipo.Text
                End With
            
                'Bloquea Hoja
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

                Unload Me
                busqCondicion.Show (0)
            Else
                MsgBox "Aprobado por Vacio"
            End If
        Else
            MsgBox "Detalle Vacio"
        End If
    Else
        MsgBox "Error en Tipo"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    
    strSQL = "SELECT * FROM DB_TIPO_CONDICION"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbTipo.Clear
        Dim cont As Integer
        cont = 0
        Do While Not rs.EOF
            cmbTipo.AddItem rs.Fields("NOMBRE_TIPO")
            cmbTipo.List(cont, 1) = rs.Fields("ID_TIPO_CONDICION")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    closeRS
    cmbTipo.ListIndex = 0
    
    strSQL = "SELECT * FROM ((((DB_TIPO_CONDICION " & _
    "LEFT JOIN DB_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK) " & _
    "LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK) " & _
    "LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
    "LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK) " & _
    "LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK " & _
    "WHERE ID_CONDICION = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & _
    " AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDoi.Caption = lbDoi.Caption & rs.Fields("DOI")
        lbSolicitud.Caption = lbSolicitud.Caption & rs.Fields("SOLICITUD")
        lbProducto.Caption = lbProducto.Caption & rs.Fields("NOMBRE_PRODUCTO")
        lbMoneda.Caption = lbMoneda.Caption & rs.Fields("NOMBRE_MONEDA")
        lbMonto.Caption = lbMonto.Caption & Format(rs.Fields("MONTO"), "#,##0.00")
        lbDesembolso.Caption = lbDesembolso.Caption & rs.Fields("FECHA_DESEMBOLSO")
        cmbTipo.Text = rs.Fields("NOMBRE_TIPO")
        tbDetalle.Text = rs.Fields("DETALLE")
        cmbAprobado.Text = rs.Fields("APROBADO_POR")
        
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = rs.Fields("ID_SOCIO")
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = rs.Fields("ID_PRESTAMO")
    End If
    closeRS
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
