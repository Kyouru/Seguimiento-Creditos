Private Sub btCancelar_Click()
    Unload Me
    busqSocio.Show (0)
End Sub

Private Sub btGuardar_Click()
    If tbCodigo.Text <> "" And tbDOI.Text <> "" And tbNombre.Text <> "" Then
        strSQL = "INSERT INTO DB_SOCIO (ID_GRUPO_FK, DOI, CODIGO_SOCIO, NOMBRE_SOCIO) " & _
        "VALUES (" & cmbGrupo.List(cmbGrupo.ListIndex, 1) & ", '" & tbDOI.Text & "', '" & _
        tbCodigo.Text & "', '" & Replace(tbNombre.Text, "'", "''") & "')"
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute strSQL
        strSQL = "SELECT @@IDENTITY"
        rs.Open strSQL, cnn
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = rs.Fields(0)
        closeRS
        
        Unload Me
        prestamoDesembolsado.Show (0)
    Else
        MsgBox "Informacion Incompleta"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub tbNombre_Change()
    tbNombre.Text = UCase(tbNombre.Text)
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
    cmbGrupo.ListIndex = 0
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub
