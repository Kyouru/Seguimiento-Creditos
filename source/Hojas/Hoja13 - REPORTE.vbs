
Private Sub btGenerar_Click()
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
            " WHERE DB_SEGUIMIENTO.ANULADO = FALSE AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE "
            strSQL = strSQL & " AND FECHA_PROXIMA_ACCION >= #" & fechaStrStr(.tbFechaInicio.Text) & "# AND FECHA_PROXIMA_ACCION <= #" & fechaStrStr(.tbFechaFin.Text) & "# AND ("
            
            If cbGarantia.Value Then
                strSQL = strSQL & "NOMBRE_TIPO = 'GARANTIA'"
            End If
            If cbSeguro.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_TIPO = 'SEGURO'"
            End If
            If cbSeguimiento.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_TIPO = 'SEGUIMIENTO'"
            End If
            If cbCovenant.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_TIPO = 'COVENANT'"
            End If
            If cbAnulado.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_TIPO = 'ANULADO'"
            End If
            If cbSinCondiciones.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_TIPO = 'SIN CONDICIONES'"
            End If
            If cbDenegado.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_TIPO = 'DENEGADO'"
            End If
            
            
            strSQL = strSQL & ") AND ("
            
            
            If cbEnProceso.Value Then
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'EN PROCESO'"
            End If
            If cbFinalizado.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'FINALIZADO'"
            End If
            If cbEAnulado.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'ANULADO'"
            End If
            If tbESinCondiciones.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'SIN CONDICIONES'"
            End If
            If cbExonerado.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'EXONERADO'"
            End If
            If cbEDenegado.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'DENEGADO'"
            End If
            If cbStandBy.Value Then
                If Right(strSQL, 1) <> "(" Then
                    strSQL = strSQL & " OR "
                End If
                strSQL = strSQL & "NOMBRE_ESTADO_SEGUIMIENTO = 'STAND BY'"
            End If
            strSQL = strSQL & ")"
            
            strSQL = strSQL & " ORDER BY NOMBRE_SOCIO, SOLICITUD, FECHA_ACCION ASC"
            ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range(ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5"), ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("dataSetTemp5").End(xlDown)).ClearContents
            Debug.Print strSQL
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
            .PivotTables("TablaReporte").ChangePivotCache ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Name & "!R1C1:R" & ThisWorkbook.Sheets(NOMBRE_HOJA_TEMP5).Range("CUENTA").Value & "C19")
            
            ThisWorkbook.RefreshAll
        Else
            MsgBox "Error en Fecha Fin"
        End If
    Else
        MsgBox "Error en Fecha Inicio"
    End If
    End With
End Sub


