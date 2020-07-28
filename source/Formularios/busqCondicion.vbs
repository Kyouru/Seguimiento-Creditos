Private Sub btAtras_Click()
    Unload Me
    busqPrestamo.Show (0)
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            'Confirmacion antes de Anular la Condicion
            Dim resp As Integer
            resp = MsgBox("Esta seguro que desea eliminar esta condición?", vbYesNo + vbQuestion, "Borrar Condición")
            If resp = vbYes Then
                OpenDB
                strSQL = "UPDATE DB_CONDICION SET ANULADO = TRUE WHERE ID_CONDICION = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute (strSQL)
                strSQL = "UPDATE DB_SEGUIMIENTO SET ANULADO = TRUE WHERE ID_CONDICION_FK = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute (strSQL)
                closeRS
                
                'Actualizar la Lista de Condiciones
                ActualizarHoja
                ActualizarLista
            End If
        Else
            MsgBox "Seleccione una entrada con datos"
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btEliminar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Modifica la Condicion Seleccionada
Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = ListBox1.List(ListBox1.ListIndex)
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = ""
            Unload Me
            modCondicion.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub btNuevo_Click()
    Unload Me
    newCondicion.Show (0)
End Sub

'Lista los seguimientos de la condicion seleccionada
Private Sub btSeleccionar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = ListBox1.List(ListBox1.ListIndex)
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = ""
            Unload Me
            busqSeguimiento.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

'Consulta el ultimo estado de la Condicion (Seguimiento con Fecha Accion mas reciente)
Public Sub ListBox1_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            strSQL = "SELECT S.ID_SEGUIMIENTO, S.NOMBRE_ESTADO_SEGUIMIENTO, R.MAXFECHA FROM (SELECT ID_CONDICION_FK, MAX(FECHA_ACCION) AS MAXFECHA FROM DB_SEGUIMIENTO WHERE DB_SEGUIMIENTO.ANULADO = FALSE GROUP BY ID_CONDICION_FK) AS R INNER JOIN (SELECT * FROM DB_SEGUIMIENTO LEFT JOIN DB_ESTADO_SEGUIMIENTO ON DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK) AS S ON S.ID_CONDICION_FK = R.ID_CONDICION_FK AND S.FECHA_ACCION = R.MAXFECHA WHERE S.ID_CONDICION_FK = " & ListBox1.List(ListBox1.ListIndex) & ""
            OpenDB
            On Error GoTo Handle:
            rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                lbEstado.Caption = "Estado: " & rs.Fields("NOMBRE_ESTADO_SEGUIMIENTO")
            Else
                lbEstado.Caption = "Estado: SIN INICIAR"
            End If
            closeRS
        End If
    End If
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ListBox1_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'Lista los seguimientos de la condicion seleccionada
    If NUEVA_ACCION Then
        btSeleccionar_Click
    End If
End Sub

Private Sub UserForm_Initialize()
    
    If NUEVA_ACCION Then
        btSeleccionar.Visible = True
    Else
        btSeleccionar.Visible = False
    End If
    
    strSQL = "SELECT DOI, CODIGO_SOCIO, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_SOCIO, NOMBRE_MONEDA, MONTO, " & _
    "FECHA_DESEMBOLSO FROM ((DB_PRESTAMO LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK)" & _
    "LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK " & _
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
        lbMonto.Caption = lbMonto.Caption & Format(rs.Fields("MONTO"), "#,##0.00")
        If IsDate(rs.Fields("FECHA_DESEMBOLSO")) Then
            lbDesembolso.Caption = lbDesembolso.Caption & rs.Fields("FECHA_DESEMBOLSO")
        Else
            lbDesembolso.Caption = lbDesembolso.Caption & "SIN DESEMBOLSAR"
        End If
    End If
    closeRS
    
    'Actualiza la Lista de Condiciones
    ActualizarHoja
    ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Se Solicita todas las Condiciones del Prestamo Seleccionado previamente y se Copian a una hoja Temporal para luego poder agregarlos a la ListBox
Public Sub ActualizarHoja()
    strSQL = "SELECT ID_CONDICION, NOMBRE_TIPO, DETALLE, APROBADO_POR, ID_PRESTAMO_FK FROM ((" & _
    "DB_CONDICION LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK)" & _
    " LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK)" & _
    " LEFT JOIN DB_TIPO_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK" & _
    " WHERE DB_PRESTAMO.ID_PRESTAMO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") & _
    " AND DB_CONDICION.ANULADO = FALSE AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE "
    
    Dim primerTipo As Boolean: primerTipo = True
    
    If SEGUIMIENTO_GENERAL Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'SEGUIMIENTO' OR DB_TIPO_CONDICION.NOMBRE_TIPO = 'ANULADO' OR DB_TIPO_CONDICION.NOMBRE_TIPO = 'SIN CONDICIONES' OR DB_TIPO_CONDICION.NOMBRE_TIPO = 'DENEGADO'"
    End If
    
    If SEGUIMIENTO_GARANTIA Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'GARANTIA'"
    End If
    
    If SEGUIMIENTO_SEGURO Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'SEGURO'"
    End If
    
    If SEGUIMIENTO_COVENANT Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'COVENANT'"
    End If
    
    If Not primerTipo Then
        strSQL = strSQL & ")"
    End If
    
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP3)
        'Limpiar Hoja Temporal
        .Range(.Range("dataSetTemp3"), .Range("dataSetTemp3").End(xlDown)).ClearContents
        
        OpenDB
        On Error GoTo Handle:
        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            .Range("dataSetTemp3").CopyFromRecordset rs
        End If
        closeRS
    End With
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP3)
        ListBox1.ColumnWidths = "0;60;270;80;0"
        ListBox1.ColumnCount = 5
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp3").Address, Len(.Range("dataSetTemp3").Address) - 1) & .Range("A3").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp3").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
        
        'En caso de que se provenga de un nivel superior (busqSeguimiento -> Atras) se selecciona la Condicion a la que pertenecia el Seguimiento
        'Case contrario se selecciona la primera condicion si la hubiese
        If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") <> "" Then
            For i = 0 To (ListBox1.ListCount - 1)
                If ListBox1.List(i, 0) = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") Then
                    ListBox1.ListIndex = i
                    Exit For
                End If
            Next
        Else
            If ListBox1.ListCount > 0 Then
                ListBox1.ListIndex = 0
            End If
        End If
        
    End With
End Sub
