
'Actualiza la Lista
Private Sub btBuscar_Click()
    ActualizarHoja
    ActualizarLista
End Sub

Private Sub btEliminar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            'Confirmacion antes de anular el Prestamo
            Dim resp As Integer
            resp = MsgBox("Esta seguro que desea eliminar este socio?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 3))
            If resp = vbYes Then
            
                OpenDB
                On Error GoTo Handle:
                strSQL = "UPDATE DB_SOCIO SET ANULADO = TRUE WHERE ID_SOCIO = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute (strSQL)
                strSQL = "UPDATE DB_PRESTAMO SET ANULADO = TRUE WHERE ID_SOCIO_FK = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute (strSQL)
                strSQL = "UPDATE DB_CONDICION SET ANULADO = TRUE WHERE ID_PRESTAMO_FK IN (SELECT ID_PRESTAMO FROM DB_PRESTAMO WHERE ID_SOCIO_FK = " & ListBox1.List(ListBox1.ListIndex) & ")"
                cnn.Execute (strSQL)
                strSQL = "UPDATE DB_SOCIO SET ANULADO = TRUE WHERE ID_SOCIO = " & ListBox1.List(ListBox1.ListIndex)
                cnn.Execute ("UPDATE DB_SEGUIMIENTO SET ANULADO = TRUE WHERE ID_CONDICION_FK IN (SELECT ID_CONDICION FROM DB_CONDICION WHERE ID_PRESTAMO_FK IN (SELECT ID_PRESTAMO FROM DB_PRESTAMO WHERE ID_SOCIO_FK = " & ListBox1.List(ListBox1.ListIndex)) & "))"
                closeRS
            
                'Actualizar la ListBox
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

'Limpia todos los campos de busqueda
Private Sub btLimpiar_Click()
    tbCodSocio.Text = ""
    tbDOI.Text = ""
    cmbGrupo.Text = ""
    tbNomSocio.Text = ""
End Sub

'Modifica el Socio Seleccionado
Private Sub btModificar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = ListBox1.List(ListBox1.ListIndex)
            Unload Me
            modSocio.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

'Nuevo Socio
Private Sub btNuevo_Click()
    Unload Me
    newSocio.Show (0)
End Sub

Private Sub btSalir_Click()
    Unload Me
End Sub

'Busca todos los Prestamos del Socio seleccionado
Private Sub btSeleccionar_Click()
    'Solo si se selecciono algun item de la lista
    If ListBox1.ListIndex <> -1 Then
        'Solo si se selecciono algun item de la lista y no es vacio
        If ListBox1.List(ListBox1.ListIndex) <> "" Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = ListBox1.List(ListBox1.ListIndex)
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = ""
            Unload Me
            busqPrestamo.Show (0)
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub

Private Sub cmbGrupo_Change()
    btBuscar_Click
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btSeleccionar_Click
End Sub

Private Sub tbCodSocio_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub tbDOI_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub tbNomSocio_Change()
    tbNomSocio.Text = UCase(tbNomSocio.Text)
End Sub

Private Sub tbNomSocio_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btBuscar_Click
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim cont As Integer
    strSQL = "SELECT * FROM DB_GRUPO WHERE DB_GRUPO.ANULADO = FALSE ORDER BY NOMBRE_GRUPO"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbGrupo.Clear
        cont = 0
        Do While Not rs.EOF
            cmbGrupo.AddItem rs.Fields("NOMBRE_GRUPO")
            cmbGrupo.List(cont, 1) = rs.Fields("ID_GRUPO")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    
    'Actualizar la Lista de Socios
    ActualizarHoja
    ActualizarLista
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

'Se Solicita todos los Socios que cumplan los filtros y se Copian a una hoja Temporal para luego poder agregarlos a la ListBox
Public Sub ActualizarHoja()
    strSQL = "SELECT ID_SOCIO, NOMBRE_GRUPO, DOI, CODIGO_SOCIO, NOMBRE_SOCIO FROM DB_SOCIO LEFT JOIN DB_GRUPO ON DB_GRUPO.ID_GRUPO = DB_SOCIO.ID_GRUPO_FK WHERE 1=1"
    If tbCodSocio.Text <> "" Then
        strSQL = strSQL & " AND CODIGO_SOCIO LIKE '%" & tbCodSocio.Text & "%'"
    End If
    If tbNomSocio.Text <> "" Then
        strSQL = strSQL & " AND NOMBRE_SOCIO LIKE '%" & tbNomSocio.Text & "%'"
    End If
    If tbDOI.Text <> "" Then
        strSQL = strSQL & " AND DOI LIKE '%" & tbDOI.Text & "%'"
    End If
    If cmbGrupo.Text <> "" Then
        strSQL = strSQL & " AND ID_GRUPO_FK = " & cmbGrupo.List(cmbGrupo.ListIndex, 1)
    End If
    strSQL = strSQL & " AND DB_SOCIO.ANULADO = FALSE ORDER BY NOMBRE_GRUPO, NOMBRE_SOCIO"
    
    'Limpiar Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("dataSetTemp"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("dataSetTemp").End(xlDown)).ClearContents
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP).Range("dataSetTemp").CopyFromRecordset rs
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
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP)
        ListBox1.ColumnWidths = "0;100;80;65;200"
        ListBox1.ColumnCount = 5
        ListBox1.ColumnHeads = True
        
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp").Address, Len(.Range("dataSetTemp").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
        
        'En caso de que se provenga de un nivel superior (busqPrestamo -> Atras) se selecciona el socio al que se le otorgo el prestamo
        If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") <> "" Then
            For i = 0 To (ListBox1.ListCount - 1)
                If ListBox1.List(i, 0) = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") Then
                    ListBox1.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End With
End Sub
