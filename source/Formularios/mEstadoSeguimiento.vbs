Private Sub btAgregar_Click()
    Dim myValue, myValue2 As Variant
    myValue = InputBox("Nombre del Nuevo Estado de Seguimiento:", "Nuevo Estado de Seguimiento")
    If myValue <> "" Then
        strSQL = "INSERT INTO DB_ESTADO_SEGUIMIENTO (NOMBRE_ESTADO_SEGUIMIENTO, FIN) VALUES ('" & UCase(myValue) & "', FALSE);"
        
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

    strSQL = "SELECT ID_ESTADO_SEGUIMIENTO, NOMBRE_ESTADO_SEGUIMIENTO, FIN FROM DB_ESTADO_SEGUIMIENTO WHERE DB_ESTADO_SEGUIMIENTO.ANULADO = FALSE"
    
    'Limpiar Hoja Temporal
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range("dataSetTemp13"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range("dataSetTemp13").End(xlDown)).ClearContents
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13).Range("dataSetTemp13").CopyFromRecordset rs
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP13)
        ListBox1.ColumnWidths = "40;80;40"
        ListBox1.ColumnCount = 3
        ListBox1.ColumnHeads = True
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp13").Address, Len(.Range("dataSetTemp13").Address) - 1) & .Range("A3").End(xlDown).Row
        Else
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp13").Address
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
    End With
End Sub

Private Sub btEliminar_Click()
    If ListBox1.ListIndex <> -1 Then
    
        'Confirmar la Anulacion del Estado de Seguimiento
        Dim resp As Integer
        resp = MsgBox("Esta seguro que desea eliminar este estado de seguimiento?", vbYesNo + vbQuestion, ListBox1.List(ListBox1.ListIndex, 1))
        If resp = vbYes Then
            strSQL = "UPDATE DB_ESTADO_SEGUIMIENTO SET DB_ESTADO_SEGUIMIENTO.ANULADO = TRUE WHERE ID_ESTADO_SEGUIMIENTO = " & ListBox1.List(ListBox1.ListIndex, 0)
            
            OpenDB
            On Error GoTo Handle:
            cnn.Execute (strSQL)
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

Private Sub btFin_Click()
    If ListBox1.ListIndex <> -1 Then
        If ListBox1.List(ListBox1.ListIndex, 2) Then
            strSQL = "UPDATE DB_ESTADO_SEGUIMIENTO SET FIN = FALSE WHERE ID_ESTADO_SEGUIMIENTO = " & ListBox1.List(ListBox1.ListIndex, 0)
        Else
            strSQL = "UPDATE DB_ESTADO_SEGUIMIENTO SET FIN = TRUE WHERE ID_ESTADO_SEGUIMIENTO = " & ListBox1.List(ListBox1.ListIndex, 0)
        End If
        
        OpenDB
        On Error GoTo Handle:
        cnn.Execute strSQL
        closeRS
        
        ActualizarHoja
        ActualizarLista
    Else
        MsgBox "Seleccione una entrada"
    End If
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btFin_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btModificar_Click()
    If ListBox1.ListIndex <> -1 Then
        Dim myValue As Variant
        myValue = InputBox("Nombre del Estado de Seguimiento:", "Modificar Estado de Seguimiento", ListBox1.List(ListBox1.ListIndex, 1))
        If myValue <> "" Then
            strSQL = "UPDATE DB_ESTADO_SEGUIMIENTO SET NOMBRE_ESTADO_SEGUIMIENTO = '" & UCase(myValue) & "' WHERE ID_ESTADO_SEGUIMIENTO = " & ListBox1.List(ListBox1.ListIndex, 0)
            
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

