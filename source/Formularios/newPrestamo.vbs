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
    If IsNumeric(tbMonto.Text) Then
    If tbSolicitud.Text <> "" Then
    If cmbProducto.ListIndex <> -1 Then
    If cmbMoneda.ListIndex <> -1 Then
    
    strSQL = "INSERT INTO DB_PRESTAMO (SOLICITUD, ID_PRODUCTO_FK, ID_MONEDA_FK, MONTO, FECHA_DESEMBOLSO, " & _
    "ID_ESTADO_PRESTAMO_FK, ID_SOCIO_FK) VALUES ('" & tbSolicitud.Text & "'," & _
    cmbProducto.List(cmbProducto.ListIndex, 1) & "," & cmbMoneda.List(cmbMoneda.ListIndex, 1) & "," & _
    tbMonto.Text & ","
    If tbDesembolso.Text <> "" Then
        If CDate(tbDesembolso.Text) > Now Then
            MsgBox "Desembolso Futuro"
            Exit Sub
        End If
        strSQL = strSQL & "#" & Format(CDate(tbDesembolso.Text), "yyyy/mm/dd") & "#,"
    Else
        strSQL = strSQL & "NULL,"
    End If
    
    strSQL = strSQL & "1," & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") & ")"
    
    OpenDB
    On Error GoTo Handle:
    cnn.Execute strSQL
    strSQL = "SELECT @@IDENTITY"
    rs.Open strSQL, cnn
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = rs.Fields(0)
    closeRS
    
    busqPrestamo.ActualizarHoja
    busqPrestamo.ActualizarLista
    
    Unload Me
    
    newCondicion.Show (0)
    Else
        MsgBox "Moneda Incorrecto"
    End If
    Else
        MsgBox "Producto Incorrecto"
    End If
    Else
        MsgBox "Solicitud Vacia"
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

    strSQL = "SELECT * FROM DB_SOCIO" & _
    " WHERE ID_SOCIO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO")
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        lbNombre.Caption = lbNombre.Caption & rs.Fields("NOMBRE_SOCIO")
        lbCodigo.Caption = lbCodigo.Caption & rs.Fields("CODIGO_SOCIO")
        lbDOI.Caption = lbDOI.Caption & rs.Fields("DOI")
    End If
    closeRS
    
    Dim cont As Integer
    strSQL = "SELECT * FROM DB_PRODUCTO"
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
    cmbProducto.ListIndex = 0
    
    strSQL = "SELECT * FROM DB_MONEDA"
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
    cmbMoneda.ListIndex = 0
    
    If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("SOLICITUD") <> "" Then
        tbSolicitud.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("SOLICITUD")
        tbDesembolso.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("DESEMBOLSO")
        tbMonto.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("MONTO")
        cmbProducto.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("PRODUCTO")
        cmbMoneda.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("MONEDA")
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
