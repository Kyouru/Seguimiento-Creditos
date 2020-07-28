Private Sub btAtras_Click()
    Unload Me
    busqCondicion.Show (0)
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            'Confirmacion anter de anular el Seguimiento
            Dim resp As Integer
            resp = MsgBox("Esta seguro que desea eliminar este seguimiento?", vbYesNo + vbQuestion, "Borrar Seguimiento")
            If resp = vbYes Then
                strSQL = "UPDATE DB_SEGUIMIENTO SET ANULADO = TRUE WHERE ID_SEGUIMIENTO = " & ListBox1.List(ListBox1.ListIndex)
                
                OpenDB
                On Error GoTo Handle:
                cnn.Execute (strSQL)
                closeRS
                
                ActualizarHoja
                ActualizarLista
            End If
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btEliminar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Modifica el Seguimiento Seleccionado
Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = ListBox1.List(ListBox1.ListIndex)
            Unload Me
            modSeguimiento.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

'Ingresa un Nuevo Seguimiento
Private Sub btNuevo_Click()
    Unload Me
    newSeguimiento.Show (0)
End Sub

'Consulta el estado al que paso la condicion al realizar ese Seguimiento
Private Sub ListBox1_Change()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            lbEstado.Caption = "Estado Seleccionado: " & ListBox1.List(ListBox1.ListIndex, 4)
        End If
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btModificar_Click
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
        
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = rs.Fields("ID_SOCIO")
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = rs.Fields("ID_PRESTAMO")
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = rs.Fields("ID_CONDICION")
    End If
    closeRS
    
    'Actualizar la Lista
    ActualizarHoja
    ActualizarLista
    
    'En caso halla Seguimientos, se selecciona la primera por defecto
    If ListBox1.ListCount > 0 Then
        ListBox1.ListIndex = 0
    End If
    
    'En caso se halla abierto este formulario a traves del macro ModificarAccion del Modulo2, se busca el seguimiento especifico a editar
    If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") <> "" Then
        Dim i As Integer
        For i = 0 To ListBox1.ListCount - 1
            If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = ListBox1.List(i) Then
                ListBox1.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
    
'Se Solicita todos los Seguimientos de la Condicion Seleccionada previamente y se Copian a una hoja Temporal para luego poder agregarlos a la ListBox
Public Sub ActualizarHoja()
    strSQL = "SELECT ID_SEGUIMIENTO, FECHA_ACCION, FECHA_PROXIMA_ACCION, DETALLE_ACCION, NOMBRE_ESTADO_SEGUIMIENTO, ID_CONDICION_FK, USUARIO FROM (((DB_ESTADO_SEGUIMIENTO LEFT JOIN DB_SEGUIMIENTO ON DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK) LEFT JOIN DB_CONDICION ON DB_SEGUIMIENTO.ID_CONDICION_FK = DB_CONDICION.ID_CONDICION) LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK) LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK WHERE ID_CONDICION = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & " AND DB_SEGUIMIENTO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_SOCIO.ANULADO = FALSE ORDER BY FECHA_ACCION ASC"
    
    'Limpia Hoja Temporal
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range("dataSetTemp4"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range("dataSetTemp4").End(xlDown)).ClearContents
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Application.Calculation = xlCalculationManual
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4).Range("dataSetTemp4").CopyFromRecordset rs
        Application.Calculation = xlCalculationAutomatic
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP4)
        ListBox1.ColumnWidths = "0;50;50;250;60;0;20"
        ListBox1.ColumnCount = 7
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp4").Address, Len(.Range("dataSetTemp4").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp4").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub
