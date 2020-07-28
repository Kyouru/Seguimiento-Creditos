Private Sub btCalendario_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA") = ""
    frmCalendario.Show
    tbDesembolso.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA")
End Sub

Private Sub btCancelar_Click()
    Unload Me
    busqPrestamo.Show (0)
End Sub

Private Sub btGuardar_Click()
    If tbDesembolso.Text = "" Or IsDate(tbDesembolso.Text) Then
    If tbMonto.Text = "" Or IsNumeric(tbMonto.Text) Then
    If cmbProducto.ListIndex <> -1 Then
    If cmbMoneda.ListIndex <> -1 Then
    
        strSQL = "UPDATE DB_PRESTAMO SET SOLICITUD = '" & tbSolicitud.Text & "', ID_PRODUCTO_FK = " & _
        cmbProducto.List(cmbProducto.ListIndex, 1) & ", ID_MONEDA_FK = " & _
        cmbMoneda.List(cmbMoneda.ListIndex, 1) & ", MONTO = "
        If tbMonto.Text <> "" Then
            strSQL = strSQL & tbMonto.Text
        Else
            strSQL = strSQL & "NULL"
        End If
        strSQL = strSQL & ", FECHA_DESEMBOLSO = "
        If tbDesembolso.Text <> "" Then
            If CDate(tbDesembolso.Text) > Now Then
                MsgBox "Desembolso Futuro"
                Exit Sub
            End If
            strSQL = strSQL & "#" & Format(CDate(tbDesembolso.Text), "yyyy/mm/dd") & "#,"
        Else
            strSQL = strSQL & "NULL,"
        End If
        strSQL = strSQL & " ID_ESTADO_PRESTAMO_FK = " & cmbEstado.List(cmbEstado.ListIndex, 1) & _
        " WHERE ID_PRESTAMO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO")
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute strSQL
        closeRS
        
        busqPrestamo.ActualizarHoja
        busqPrestamo.ActualizarLista
        
        Unload Me
        busqPrestamo.Show (0)
    Else
        MsgBox "Error en Moneda"
    End If
    Else
        MsgBox "Error en Producto"
    End If
    Else
        MsgBox "Monto Incorrecto"
    End If
    Else
        MsgBox "Desembolso Incorrecto"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    
    Dim cont As Integer
    strSQL = "SELECT * FROM DB_PRODUCTO WHERE ANULADO = FALSE"
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbProducto.Clear
        cont = 0
        Do While Not rs.EOF
            cmbProducto.AddItem rs.Fields("NOMBRE_PRODUCTO")
            cmbProducto.List(cont, 1) = rs.Fields("ID_PRODUCTO")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    closeRS
    
    strSQL = "SELECT * FROM DB_MONEDA WHERE ANULADO = FALSE"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbMoneda.Clear
        cont = 0
        Do While Not rs.EOF
            cmbMoneda.AddItem rs.Fields("NOMBRE_MONEDA")
            cmbMoneda.List(cont, 1) = rs.Fields("ID_MONEDA")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    closeRS
    
    strSQL = "SELECT * FROM DB_ESTADO_PRESTAMO WHERE ANULADO = FALSE"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbEstado.Clear
        cont = 0
        Do While Not rs.EOF
            cmbEstado.AddItem rs.Fields("NOMBRE_ESTADO_PRESTAMO")
            cmbEstado.List(cont, 1) = rs.Fields("ID_ESTADO_PRESTAMO")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
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
        tbSolicitud.Text = rs.Fields("SOLICITUD")
        cmbProducto.Text = rs.Fields("NOMBRE_PRODUCTO")
        cmbMoneda.Text = rs.Fields("NOMBRE_MONEDA")
        tbMonto.Text = rs.Fields("MONTO")
        If Not IsNull(rs.Fields("FECHA_DESEMBOLSO")) Then
            tbDesembolso.Text = Format(rs.Fields("FECHA_DESEMBOLSO"), "DD/MM/YYYY")
        End If
        cmbEstado.Text = rs.Fields("NOMBRE_ESTADO_PRESTAMO")
    End If
    closeRS
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
