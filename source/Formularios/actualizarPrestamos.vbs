Private Sub ActualizarHoja()

     strSQL = "SELECT ID_PRESTAMO, CODIGO_SOCIO, NOMBRE_SOCIO, MONTO, NOMBRE_PRODUCTO, NOMBRE_MONEDA, FECHA_DESEMBOLSO, SOLICITUD FROM (((DB_PRESTAMO" & _
            " LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
            " LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK) " & _
            " LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK)"

    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15").End(xlDown)).ClearContents
    OpenDB
    
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15").Cells(1, 1).CopyFromRecordset rs
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("B:B").NumberFormat = "0.00"
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("B:B").Style = "Comma"
    End If
    closeRS
    
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    Call fso.CopyFile(ThisWorkbook.Sheets(NOMBRE_HOJA_L).Range("DB_PATH_REGISTRO_SISGO"), ActiveWorkbook.Path & "\DATABASE\REGISTRO\Registro SISGO.xlsx", 1)

    OpenDB3
    strSQL = "SELECT * FROM [" & 2017 & "$]"

    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7").End(xlDown)).ClearContents
    
    On Error GoTo Handle3:
    rs3.Open strSQL, cnn3, adOpenKeyset, adLockOptimistic
    
    If rs3.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7").CopyFromRecordset rs3
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("J:J").NumberFormat = "0.00"
        cerrar = False
        
        strSQL = "SELECT B.[ID_PRESTAMO], B.[SOCIO], B.[NOMBRE], B.[SOLICITUD], B.[MONTO], B.[FECHA_DESEMBOLSO], A.[F_AUTORIZACION] FROM [" & NOMBRE_HOJA_TEMP15 & "$] AS B LEFT JOIN [" & NOMBRE_HOJA_TEMP7 & "$] AS A ON A.[SOLICITUD2] = B.[SOLICITUD] WHERE A.[SOLICITUD] IS NOT NULL AND (B.[FECHA_DESEMBOLSO] IS NULL OR A.[F_AUTORIZACION] <> B.[FECHA_DESEMBOLSO]) ORDER BY A.[F_AUTORIZACION] DESC"
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range("dataSetTemp16"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range("dataSetTemp16").End(xlDown)).ClearContents
        
        OpenDB2
        On Error GoTo Handle2:
        rs2.Open strSQL, cnn2, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16).Range("dataSetTemp16").CopyFromRecordset rs2
        End If
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
Handle2:
    If cnn2.Errors.count > 0 Then
        Call Error_Handle(cnn2.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn2.Errors.Item(0).Number, cnn2.Errors.Item(0).Description)
    End If
    cnn2.Errors.Clear
    closeRS2
Handle3:
    If cnn3.Errors.count > 0 Then
        Call Error_Handle(cnn3.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn3.Errors.Item(0).Number, cnn3.Errors.Item(0).Description)
    End If
    cnn3.Errors.Clear
    closeRS3
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7").End(xlDown)).ClearContents
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP15).Range("dataSetTemp15").End(xlDown)).ClearContents

End Sub

'Agrega la Hoja Temporal a la ListBox
Public Sub ActualizarLista()
    With ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP16)
        ListBox1.ColumnWidths = "0;45;180;80;80;60;80;"
        ListBox1.ColumnCount = 7
        ListBox1.ColumnHeads = True
        'En caso halla mas de una fila
        If .Range("A3") <> "" Then
            ListBox1.RowSource = .Name & "!" & Left(.Range("dataSetTemp16").Address, Len(.Range("dataSetTemp16").Address) - 1) & .Range("A2").End(xlDown).Row
        Else
            'En caso halla solamente una fila
            If .Range("A2") <> "" Then
                ListBox1.RowSource = .Name & "!" & .Range("dataSetTemp16").Address
            'En caso no hallan datos
            Else
                ListBox1.RowSource = ""
                ListBox1.ColumnHeads = False
            End If
        End If
        
        If ListBox1.ListCount > 0 Then
            ListBox1.ListIndex = 0
        End If
    End With
    
End Sub

Private Sub btActualizar_Click()
    
    OpenDB
    On Error GoTo Handle:
    For i = 0 To (ListBox1.ListCount - 1)
        If ListBox1.Selected(i) Then
            strSQL = "UPDATE DB_PRESTAMO SET FECHA_DESEMBOLSO = #" & Format(ListBox1.List(i, 6), "YYYY-MM-DD") & "# WHERE ID_PRESTAMO = " & ListBox1.List(i, 0)
            cnn.Execute strSQL
        End If
    Next
    MsgBox "Actualizado Exitosamente"
    Unload Me
    'Sheets(NOMBRE_HOJA_TEMP16).Visible = True
    'Sheets(NOMBRE_HOJA_TEMP16).Activate
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - ActualizarHoja", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub btSeleccionarTodo_Click()
    For i = 0 To (ListBox1.ListCount - 1)
        ListBox1.Selected(i) = True
    Next
End Sub

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ActualizarHoja
    ActualizarLista
End Sub
