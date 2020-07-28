
Private cerrar As Boolean

Private Sub btAceptar_Click()
    If ListBox1.ListIndex <> -1 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("MONEDA") = Left(ListBox1.List(ListBox1.ListIndex, 0), 1)
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("DESEMBOLSO") = ListBox1.List(ListBox1.ListIndex, 1)
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("PRODUCTO") = ListBox1.List(ListBox1.ListIndex, 4)
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("SOLICITUD") = ListBox1.List(ListBox1.ListIndex, 5)
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("MONTO") = ListBox1.List(ListBox1.ListIndex, 6)
        Unload Me
        newPrestamo.Show (0)
    Else
        MsgBox "Seleccione una entrada"
    End If
End Sub


Private Sub btCancelar_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("SOLICITUD") = ""
    Unload Me
    newPrestamo.Show
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btAceptar_Click
End Sub

Private Sub UserForm_Activate()
    
    If cerrar Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("SOLICITUD") = ""
        Unload Me
        newPrestamo.Show
    End If
End Sub

Private Sub UserForm_Initialize()
    
    ActualizarHoja
    ActualizarLista
End Sub

Public Sub ActualizarHoja()
    Dim codsocio As String
    
    strSQL = "SELECT CODIGO_SOCIO FROM DB_SOCIO WHERE ID_SOCIO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO")
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        codsocio = rs.Fields(0)
    End If
    closeRS
    
    strSQL = "SELECT ID_SOCIO, ID_PRESTAMO, CODIGO_SOCIO, SOLICITUD, MONTO, FECHA_DESEMBOLSO, NOMBRE_PRODUCTO, NOMBRE_MONEDA FROM " & _
    "((((DB_PRESTAMO LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK) " & _
    "LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK) " & _
    "LEFT JOIN DB_ESTADO_PRESTAMO ON DB_ESTADO_PRESTAMO.ID_ESTADO_PRESTAMO = DB_PRESTAMO.ID_ESTADO_PRESTAMO_FK) " & _
    "LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
    "WHERE CODIGO_SOCIO='" & codsocio & "' AND DB_PRESTAMO.ANULADO = FALSE "
    
    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("dataSetTemp11"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("dataSetTemp11").End(xlDown)).ClearContents
    
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("dataSetTemp11").CopyFromRecordset rs
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("B:B").NumberFormat = "@"
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("E:E").NumberFormat = "0.00"
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("E:E").Style = "Comma"
    End If
    closeRS
    
    OpenDB3
    strSQL = "SELECT * FROM [" & 2017 & "$] WHERE [SOCIO] = '" & codsocio & "'"

    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7").End(xlDown)).ClearContents
    
    On Error GoTo Handle3:
    rs3.Open strSQL, cnn3, adOpenKeyset, adLockOptimistic
    
    If rs3.RecordCount > 0 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("dataSetTemp7").CopyFromRecordset rs3
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("J:J").NumberFormat = "0.00"
        cerrar = False
        
        strSQL = "SELECT A.[MONEDA], A.[F_AUTORIZACION], A.[SOCIO], A.[NOMBRE], A.[PRODUCTO], A.[SOLICITUD2], A.[MONTO] FROM [" & NOMBRE_HOJA_TEMP7 & "$] AS A LEFT JOIN [" & NOMBRE_HOJA_TEMP11 & "$] AS B ON A.[SOLICITUD2] = B.[SOLICITUD] WHERE B.[SOLICITUD] IS NULL"
        
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12").End(xlDown)).ClearContents
        
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP7).Range("L:L").NumberFormat = "@"
        ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP11).Range("D:D").NumberFormat = "@"
        OpenDB2
        On Error GoTo Handle2:
        rs2.Open strSQL, cnn2, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12").CopyFromRecordset rs2
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("B:B").NumberFormat = "DD/MM/YYYY"
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("G:G").NumberFormat = "0.00"
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("G:G").Style = "Comma"
        Else
            MsgBox "No Hay Creditos Desembolsados de este Socio desde el 2016 y que no se encuentren ingresados al la Base de Datos"
            cerrar = True
        End If
    Else
        MsgBox "No Hay Creditos Desembolsados de este Socio desde el 2016"
        cerrar = True
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
End Sub


Public Sub ActualizarLista()
    With ListBox1
        .ColumnWidths = "60;60;50;100;40;80;50;"
        .ColumnCount = 7
        .ColumnHeads = True
        If ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("A3") <> "" Then
            .RowSource = ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Name & "!" & Left(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12").Address, Len(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12").Address) - 1) & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("A2").End(xlDown).Row
        Else
            If ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("A2") <> "" Then
                .RowSource = ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Name & "!" & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP12).Range("dataSetTemp12").Address
            Else
                .RowSource = ""
                .ColumnHeads = False
            End If
        End If
    End With
End Sub
