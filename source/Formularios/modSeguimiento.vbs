Dim detalleAnterior As String
Dim usuarioAnterior As String
Dim fechaAnterior As Date
Dim estadoAnterior As Integer

Private Sub btCalendario_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA") = ""
    frmCalendario.Show
    tbFecha.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA")
End Sub

Private Sub btCalendarioInicio_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA") = ""
    frmCalendario.Show
    tbFechaInicio.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA")
End Sub

Private Sub btCancelar_Click()
    Unload Me
    busqSeguimiento.Show (0)
End Sub

Private Sub btGuardar_Click()
    If tbDetalleAccion.Text <> "" Then
        strSQL = "UPDATE DB_SEGUIMIENTO SET FECHA_PROXIMA_ACCION = "
        If tbFecha.Visible Then
            If IsDate(tbFecha.Text) Then
                strSQL = strSQL & "#" & Format(tbFecha.Text, "yyyy/mm/dd") & "#, "
            Else
                strSQL = strSQL & "NULL, "
            End If
        Else
            strSQL = strSQL & " #" & Format(Now(), "yyyy/mm/dd") & "#, "
        End If
        
        strSQL = strSQL & "DETALLE_ACCION = '" & Replace(tbDetalleAccion.Text, "'", "''") & "', " & _
        "ID_ESTADO_SEGUIMIENTO_FK = " & cmbEstado.List(cmbEstado.ListIndex, 1) & ", " & _
        "FECHA_ACCION = #" & Format(tbFechaInicio.Text, "yyyy-mm-dd hh:mm:ss") & "#, " & _
        "USUARIO = '" & cmbUsuario.Text & "'" & _
        " WHERE ID_SEGUIMIENTO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO")
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        closeRS
        
        'Agrega entrada en el Historico de las Ediciones de Seguimiento
        strSQL = " INSERT INTO DB_SEGUIMIENTO_EDICION ( " & _
                        "ID_SEGUIMIENTO_FK, " & _
                        "DETALLE_ANTERIOR, " & _
                        "ID_ESTADO_SEGUIMIENTO_ANTERIOR, " & _
                        "USUARIO_EDITOR, " & _
                        "FECHA_PROXIMA_ANTERIOR, " & _
                        "FECHA_EDICION, " & _
                        "USUARIO_ANTERIOR, " & _
                        "ID_TIPO_EDICION_FK) " & _
                    "VALUES (" & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") & ", " & _
                            "'" & detalleAnterior & "', " & _
                            estadoAnterior & ", " & _
                            "'" & Application.UserName & "', " & _
                            "#" & Format(fechaAnterior, "YYYY-MM-DD") & "#, " & _
                            "#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "#, " & _
                            "'" & usuarioAnterior & "', " & _
                            "1)"

        OpenDB
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        closeRS
        
        busqSeguimiento.ActualizarHoja
        busqSeguimiento.ActualizarLista
        
        'Desbloquea Hoja
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
        
        With ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO)
            Set total = .Range("B:B")
            Set r = total.Cells.Find(What:=.Range("ID_SEGUIMIENTO"), LookAt:=xlWhole)
            If Not r Is Nothing Then
                rowNumber = r.Row
            Else
                rowNumber = .Range("FILA_ACCION")
            End If
            .Cells(rowNumber, .Range("FECHA_ACCION").Column) = Format(Now(), "yyyy/mm/dd")
            .Cells(rowNumber, .Range("FECHA_ACCION").Column + 1) = tbDetalleAccion.Text
            .Cells(rowNumber, .Range("FECHA_ACCION").Column + 2) = cmbEstado.List(cmbEstado.ListIndex, 0)
            If tbFecha.Visible Then
                .Cells(rowNumber, .Range("FECHA_ACCION").Column + 3) = Format(tbFecha.Text, "yyyy/mm/dd")
            End If
        End With
        
        'Bloquea Hoja
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

        Unload Me
        
        busqSeguimiento.Show (0)
    Else
        MsgBox "Falta Detalle de la Accion"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub cmbEstado_Change()
    If cmbEstado.ListIndex <> -1 Then
        If cmbEstado.List(cmbEstado.ListIndex, 2) = True Then
            tbFecha.Visible = False
            tbFecha.Enabled = False
            btCalendario.Visible = False
        Else
            tbFecha.Visible = True
            tbFecha.Enabled = True
            btCalendario.Visible = True
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    
    strSQL = "SELECT * FROM DB_ESTADO_SEGUIMIENTO WHERE DB_ESTADO_SEGUIMIENTO.ANULADO = FALSE"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbEstado.Clear
        Dim cont As Integer
        cont = 0
        Do While Not rs.EOF
            cmbEstado.AddItem rs.Fields("NOMBRE_ESTADO_SEGUIMIENTO")
            cmbEstado.List(cont, 1) = rs.Fields("ID_ESTADO_SEGUIMIENTO")
            cmbEstado.List(cont, 2) = rs.Fields("FIN")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    closeRS
    
    strSQL = "SELECT * FROM ((((((DB_TIPO_CONDICION " & _
    "LEFT JOIN DB_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK) " & _
    "LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK) " & _
    "LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
    "LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK) " & _
    "LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN DB_SEGUIMIENTO ON DB_SEGUIMIENTO.ID_CONDICION_FK = DB_CONDICION.ID_CONDICION) " & _
    "LEFT JOIN DB_ESTADO_SEGUIMIENTO ON DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK " & _
    "WHERE ID_CONDICION = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & _
    " AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE " & _
    " AND DB_SEGUIMIENTO.ID_SEGUIMIENTO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO")
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        lbSolicitud.Caption = lbSolicitud.Caption & rs.Fields("SOLICITUD")
        lbProducto.Caption = lbProducto.Caption & rs.Fields("NOMBRE_PRODUCTO")
        lbMoneda.Caption = lbMoneda.Caption & rs.Fields("NOMBRE_MONEDA")
        lbMonto.Caption = lbMonto.Caption & Format(rs.Fields("MONTO"), "#,##0.00")
        lbDesembolso.Caption = lbDesembolso.Caption & rs.Fields("FECHA_DESEMBOLSO")
        lbTipo.Caption = lbTipo.Caption & rs.Fields("NOMBRE_TIPO")
        lbDetalle.Caption = lbDetalle.Caption & rs.Fields("DETALLE")
        lbAprobado.Caption = lbAprobado.Caption & rs.Fields("APROBADO_POR")
        tbFecha.Text = tbFecha.Text & rs.Fields("FECHA_PROXIMA_ACCION")
        tbFechaInicio.Text = tbFechaInicio.Text & rs.Fields("FECHA_ACCION")
        tbDetalleAccion.Text = tbDetalleAccion.Text & rs.Fields("DETALLE_ACCION")
        lbSeguimiento.Caption = lbSeguimiento.Caption & rs.Fields("USUARIO")
        cmbEstado.Text = rs.Fields("NOMBRE_ESTADO_SEGUIMIENTO")
        
        detalleAnterior = "" & rs.Fields("DETALLE")
        usuarioAnterior = "" & rs.Fields("USUARIO")
        fechaAnterior = rs.Fields("FECHA_PROXIMA_ACCION")
        estadoAnterior = rs.Fields("ID_ESTADO_SEGUIMIENTO_FK")
        
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = rs.Fields("ID_SOCIO")
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = rs.Fields("ID_PRESTAMO")
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = rs.Fields("ID_CONDICION")
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = rs.Fields("ID_SEGUIMIENTO")
    End If
    
    cmbUsuario.Text = Application.UserName
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
