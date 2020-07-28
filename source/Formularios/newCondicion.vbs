
Private Sub btCancelar_Click()
    Unload Me
    busqCondicion.Show (0)
End Sub

Private Sub btGuardar_Click()
    If cmbTipo.ListIndex <> -1 Then
    If tbDetalle.Text <> "" Then
    If cmbAprobado.Text <> "" Then
    strSQL = "INSERT INTO DB_CONDICION (ID_TIPO_CONDICION_FK, DETALLE, APROBADO_POR, ID_PRESTAMO_FK) VALUES (" & _
            "'" & cmbTipo.List(cmbTipo.ListIndex, 1) & "','" & Replace(tbDetalle.Text, "'", "''") & "','" & _
            cmbAprobado.Text & "'," & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") & ")"
    OpenDB
    On Error GoTo Handle:
    cnn.Execute strSQL
    
    If cmbTipo.Text = "SIN CONDICIONES" And cmbTipo.Text = "ANULADO" And cmbTipo.Text = "DENEGADO" Then
        strSQL = "INSERT INTO DB_SEGUIMIENTO (FECHA_ACCION, ID_ESTADO_SEGUIMIENTO_FK, ID_CONDICION_FK, USUARIO) VALUES (#" & _
            Format(Now, "yyyy-mm-dd hh:mm:ss") & "#, 3, @@IDENTITY, '" & cmbAprobado.Text & "')"
    Else
        strSQL = "INSERT INTO DB_SEGUIMIENTO (FECHA_ACCION, ID_ESTADO_SEGUIMIENTO_FK, ID_CONDICION_FK, USUARIO) VALUES (#" & _
            Format(Now, "yyyy-mm-dd hh:mm:ss") & "#, 1, @@IDENTITY, '" & cmbAprobado.Text & "')"
    End If
    
    On Error GoTo Handle:
    cnn.Execute strSQL
    closeRS
    Else
        MsgBox "Usuario Vacio"
    End If
    Else
        MsgBox "Detalle Vacio"
    End If
    Else
        MsgBox "Tipo Errado"
    End If
    
    busqCondicion.ActualizarHoja
    busqCondicion.ActualizarLista
    
    Unload Me
    busqCondicion.Show (0)
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
    cmbTipo.ListIndex = 0
    closeRS
    
    strSQL = "SELECT * FROM (((DB_PRESTAMO LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
    "LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK) " & _
    "LEFT JOIN DB_ESTADO_PRESTAMO ON DB_ESTADO_PRESTAMO.ID_ESTADO_PRESTAMO = DB_PRESTAMO.ID_ESTADO_PRESTAMO_FK " & _
    "WHERE ID_PRESTAMO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO")
    
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
        lbMonto.Caption = lbMonto.Caption & rs.Fields("MONTO")
        lbDesembolso.Caption = lbDesembolso.Caption & rs.Fields("FECHA_DESEMBOLSO")
    End If
    
    cmbAprobado.Text = Application.UserName
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

