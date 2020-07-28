Private Sub btCancelar_Click()
    Unload Me
    busqSocio.Show (0)
End Sub

Private Sub btGuardar_Click()
    If cmbGrupo.ListIndex <> -1 Then
    If tbCodigo.Text <> "" Then
    If tbDOI.Text <> "" Then
    If tbNombre.Text <> "" Then
        strSQL = "UPDATE DB_SOCIO SET ID_GRUPO_FK = " & cmbGrupo.List(cmbGrupo.ListIndex, 1) & _
        ", CODIGO_SOCIO = '" & tbCodigo.Text & "', DOI = '" & tbDOI.Text & "', NOMBRE_SOCIO = '" & _
        tbNombre.Text & "' WHERE ID_SOCIO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO")
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute (strSQL)
        closeRS
        
        busqSocio.ActualizarHoja
        busqSocio.ActualizarLista
        
        Unload Me
        busqSocio.Show (0)
    Else
        MsgBox "Nombre Vacio"
    End If
    Else
        MsgBox "DOI Vacio"
    End If
    Else
        MsgBox "Codigo de Socio Vacio"
    End If
    Else
        MsgBox "Grupo no Valido"
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
    strSQL = "SELECT * FROM DB_GRUPO WHERE DB_GRUPO.ANULADO = FALSE"
    
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
    closeRS
    
    strSQL = "SELECT * FROM DB_SOCIO" & _
    " WHERE ID_SOCIO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO")
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        tbNombre.Text = rs.Fields("NOMBRE_SOCIO")
        tbCodigo.Text = rs.Fields("CODIGO_SOCIO")
        tbDOI.Text = rs.Fields("DOI")
        cmbGrupo.ListIndex = rs.Fields("ID_GRUPO_FK") - 1
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
