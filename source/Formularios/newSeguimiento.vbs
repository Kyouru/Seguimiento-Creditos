Private Sub btCalendario_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA") = ""
    frmCalendario.Show
    tbFecha.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA")
End Sub

Private Sub btCancelar_Click()
    Unload Me
    busqSeguimiento.Show (0)
End Sub

Private Sub btGuardar_Click()
    If cmbEstado.ListIndex <> -1 Then
    If tbFecha.Text <> "" Then
    If IsDate(tbFecha.Text) Then
    If cmbUsuario.Text <> "" Then
    strSQL = "INSERT INTO DB_SEGUIMIENTO (FECHA_ACCION, ID_ESTADO_SEGUIMIENTO_FK, ID_CONDICION_FK, " & _
    "USUARIO, DETALLE_ACCION, FECHA_PROXIMA_ACCION) VALUES (#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & _
    "#, " & cmbEstado.List(cmbEstado.ListIndex, 1) & ", " & _
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & ", '" & cmbUsuario.Text & "'"
    If tbDetalleAccion.Text <> "" Then
        strSQL = strSQL & ", '" & Replace(tbDetalleAccion.Text, "'", "''") & "'"
    Else
        strSQL = strSQL & ", NULL"
    End If
    If tbFecha.Visible Then
        strSQL = strSQL & ", #" & Format(tbFecha.Text, "YYYY/MM/DD") & "#)"
    Else
        strSQL = strSQL & ", #" & Format(Now(), "YYYY/MM/DD") & "#)"
    End If
    
    OpenDB
    On Error GoTo Handle:
    cnn.Execute strSQL
    closeRS
    
    busqSeguimiento.ActualizarHoja
    busqSeguimiento.ActualizarLista
    
    Unload Me
    busqSeguimiento.Show (0)
    
    Else
        MsgBox "Usuario Vacio"
    End If
    Else
        MsgBox "Fecha no Valida"
    End If
    Else
        MsgBox "Fecha Vacia"
    End If
    Else
        MsgBox "Estado no Valido"
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
        If cmbEstado.List(cmbEstado.ListIndex, 2) Then
            tbFecha.Value = ""
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
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
        lbSolicitud.Caption = lbSolicitud.Caption & rs.Fields("SOLICITUD")
        lbProducto.Caption = lbProducto.Caption & rs.Fields("NOMBRE_PRODUCTO")
        lbMoneda.Caption = lbMoneda.Caption & rs.Fields("NOMBRE_MONEDA")
        lbMonto.Caption = lbMonto.Caption & Format(rs.Fields("MONTO"), "#,##0.00")
        lbDesembolso.Caption = lbDesembolso.Caption & rs.Fields("FECHA_DESEMBOLSO")
        lbTipo.Caption = lbTipo.Caption & rs.Fields("NOMBRE_TIPO")
        lbDetalle.Caption = lbDetalle.Caption & rs.Fields("DETALLE")
        lbAprobado.Caption = lbAprobado.Caption & rs.Fields("APROBADO_POR")
    End If
    
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
    cmbEstado.ListIndex = 0
    cmbUsuario.Text = Application.UserName
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
