Private Sub btGenerarRango_Click()
    With ActiveSheet
    If IsDate(.tbFechaInicio.Text) Then
        If IsDate(.tbFechaFin.Text) Then
            strSQL = "SELECT ID_CONDICION, ID_SEGUIMIENTO, NOMBRE_GRUPO, DOI, CODIGO_SOCIO, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_SOCIO, NOMBRE_MONEDA, MONTO, FECHA_DESEMBOLSO, NOMBRE_TIPO, DETALLE, FECHA_ACCION, DETALLE_ACCION, NOMBRE_ESTADO_SEGUIMIENTO, FECHA_PROXIMA_ACCION FROM ((((((( " & _
" DB_SEGUIMIENTO LEFT JOIN DB_CONDICION ON DB_SEGUIMIENTO.ID_CONDICION_FK = DB_CONDICION.ID_CONDICION)" & _
                " LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK)" & _
                " LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK)" & _
                " LEFT JOIN DB_ESTADO_SEGUIMIENTO ON DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK)" & _
                " LEFT JOIN DB_TIPO_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK)" & _
                " LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK)" & _
                " LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK)" & _
                " LEFT JOIN DB_GRUPO ON DB_GRUPO.ID_GRUPO = DB_SOCIO.ID_GRUPO_FK" & _
            " WHERE DB_SEGUIMIENTO.ANULADO = FALSE AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE"
            If .obAccion.Value Then
                strSQL = strSQL & " AND FECHA_ACCION >= #" & fechaStrStr(.tbFechaInicio.Text) & "# AND FECHA_ACCION <= #" & fechaStrStr(.tbFechaFin.Text) & "#"
            Else
                strSQL = strSQL & " AND FECHA_PROXIMA_ACCION >= #" & fechaStrStr(.tbFechaInicio.Text) & "# AND FECHA_PROXIMA_ACCION <= #" & fechaStrStr(.tbFechaFin.Text) & "#"
            End If
            
            If .cbSoloMica Then
                strSQL = strSQL & " AND (DB_CONDICION.ID_TIPO_CONDICION_FK = 1 OR DB_CONDICION.ID_TIPO_CONDICION_FK = 4)"
            End If
            
            strSQL = strSQL & " ORDER BY NOMBRE_SOCIO, SOLICITUD, FECHA_ACCION ASC"
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5").End(xlDown)).ClearContents
            OpenDB
            If cnn.State = adStateOpen Then
                rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount > 0 Then
                    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5").CopyFromRecordset rs
                End If
            End If
            closeRS
            Dim vArr
            vArr = Split(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("FECHA_DE_DESEMBOLSO").Address(True, False), "$")
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(vArr(0) & "1:" & vArr(0) & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA")).FillDown
            vArr = Split(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("ESTADO").Address(True, False), "$")
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(vArr(0) & "1:" & vArr(0) & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA")).FillDown
            .PivotTables("PorDesembolso").ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Name & "!R1C1:R" & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA").Value & "C19")
            .PivotTables("PorEstado").ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Name & "!R1C1:R" & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA").Value & "C19")
            
            ThisWorkbook.RefreshAll
        Else
            MsgBox "Error en Fecha Fin"
        End If
    Else
        MsgBox "Error en Fecha Inicio"
    End If
    End With
End Sub

Private Sub btTodo_Click()
    With ActiveSheet
    If IsDate(.tbFechaInicio.Text) Then
        If IsDate(.tbFechaFin.Text) Then
            strSQL = "SELECT ID_CONDICION, ID_SEGUIMIENTO, NOMBRE_GRUPO, DOI, CODIGO_SOCIO, SOLICITUD, NOMBRE_PRODUCTO, NOMBRE_SOCIO, NOMBRE_MONEDA, MONTO, FECHA_DESEMBOLSO, NOMBRE_TIPO, DETALLE, FECHA_ACCION, DETALLE_ACCION, NOMBRE_ESTADO_SEGUIMIENTO, FECHA_PROXIMA_ACCION FROM ((((((( " & _
" DB_SEGUIMIENTO LEFT JOIN DB_CONDICION ON DB_SEGUIMIENTO.ID_CONDICION_FK = DB_CONDICION.ID_CONDICION)" & _
                " LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK)" & _
                " LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK)" & _
                " LEFT JOIN DB_ESTADO_SEGUIMIENTO ON DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK)" & _
                " LEFT JOIN DB_TIPO_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK)" & _
                " LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK)" & _
                " LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK)" & _
                " LEFT JOIN DB_GRUPO ON DB_GRUPO.ID_GRUPO = DB_SOCIO.ID_GRUPO_FK" & _
            " WHERE DB_SEGUIMIENTO.ANULADO = FALSE AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE"
            
            If .cbSoloMica Then
                strSQL = strSQL & " AND (DB_CONDICION.ID_TIPO_CONDICION_FK = 1 OR DB_CONDICION.ID_TIPO_CONDICION_FK = 4)"
            End If
            
            strSQL = strSQL & " ORDER BY NOMBRE_SOCIO, SOLICITUD, FECHA_ACCION ASC"
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5").End(xlDown)).ClearContents
            OpenDB
            If cnn.State = adStateOpen Then
                On Error GoTo Handle:
                rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                If rs.RecordCount > 0 Then
                    ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5").CopyFromRecordset rs
                End If
                
                Dim vArr
                vArr = Split(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("FECHA_DE_DESEMBOLSO").Address(True, False), "$")
                ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(vArr(0) & "1:" & vArr(0) & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA")).FillDown
                vArr = Split(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("ESTADO").Address(True, False), "$")
                ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(vArr(0) & "1:" & vArr(0) & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA")).FillDown
            
                .PivotTables("PorDesembolso").ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Name & "!R1C1:R" & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA").Value & "C19")
                .PivotTables("PorEstado").ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Name & "!R1C1:R" & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA").Value & "C19")
            
                ThisWorkbook.RefreshAll
            End If
            closeRS
        Else
            MsgBox "Error en Fecha Fin"
        End If
    Else
        MsgBox "Error en Fecha Inicio"
    End If
    End With
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.CodeName & " - btTodo_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub


