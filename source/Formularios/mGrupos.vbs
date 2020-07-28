Private Sub btAgregar_Click()
    Dim myValue As Variant
    myValue = InputBox("Nombre del Nuevo Grupo Economico:", "Nuevo Grupo Economico")
    If myValue <> "" Then
        strSQL = "INSERT INTO DB_GRUPO (NOMBRE_GRUPO) VALUES ('" & UCase(myValue) & "');"
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute strSQL
        closeRS
        
        ActualizarHoja
        ActualizarLista
    End If
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btAgregar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btCerrar_Click()
    Unload Me
End Sub

Public Sub ActualizarHoja()
    strSQL = "SELECT ID_GRUPO, NOMBRE_GRUPO FROM DB_GRUPO WHERE DB_GRUPO.ANULADO = FALSE"
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range("dataSetTemp9"), _
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range("dataSetTemp9").End(xlDown)).ClearContents
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9).Range("dataSetTemp9").CopyFromRecordset rs
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP9)
        ListBox1.ColumnWidths = "40;80;"
        ListBox1.ColumnCount = 2
        ListBox1.ColumnHeads = True
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp9").Address, Len(.Range("dataSetTemp9").Address) - 1) & .Range("A3").End(xlDown).Row
        Else
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp9").Address
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

Private Sub btEliminar_Click()
    If ListBox1.ListIndex <> -1 Then
        'Confirmar la Anulacion del Grupo
        Dim resp As Integer
        resp = MsgBox("Esta seguro que desea eliminar este grupo?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 1))
        If resp = vbYes Then
        
            OpenDB
            On Error GoTo Handle:
            cnn.Execute ("UPDATE DB_GRUPO SET DB_GRUPO.ANULADO = TRUE WHERE ID_GRUPO = " & ListBox1.List(ListBox1.ListIndex, 0))
            closeRS
            
            ActualizarHoja
            ActualizarLista
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

Private Sub btModificar_Click()
    If ListBox1.ListIndex <> -1 Then
        Dim myValue As Variant
        myValue = InputBox("Nuevo Nombre del Grupo Económico:", "Modificar Grupo Económico", ListBox1.List(ListBox1.ListIndex, 1))
        If myValue <> "" Then
            strSQL = "UPDATE DB_GRUPO SET NOMBRE_GRUPO = '" & UCase(myValue) & "' WHERE ID_GRUPO = " & ListBox1.List(ListBox1.ListIndex, 0)
            
            OpenDB
            On Error GoTo Handle:
            cnn.Execute strSQL
            closeRS
            
            ActualizarHoja
            ActualizarLista
        End If
    Else
        MsgBox "Seleccione una entrada"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btModificar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    ActualizarHoja
    ActualizarLista
End Sub
