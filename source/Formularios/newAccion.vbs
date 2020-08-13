
Private Sub btCalendario_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA") = ""
    frmCalendario.Show
    tbFecha.Text = ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FECHA")
End Sub

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub cmbEstado_Change()
    If cmbEstado.ListIndex <> -1 Then
        If cmbEstado.List(cmbEstado.ListIndex, 2) Or Not NUEVA_ACCION Then
            tbFecha.Visible = False
            btCalendario.Visible = False
            lbFecha.Visible = False
        Else
            tbFecha.Visible = True
            btCalendario.Visible = True
            lbFecha.Visible = True
        End If
    End If
End Sub

Private Sub btGuardar_Click()
    If IsDate(tbFecha.Text) Or Not tbFecha.Visible Then
        strSQL = "INSERT INTO DB_SEGUIMIENTO (FECHA_ACCION, DETALLE_ACCION, ID_ESTADO_SEGUIMIENTO_FK," & _
        "ID_CONDICION_FK, USUARIO, FECHA_PROXIMA_ACCION) VALUES (#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "#, "
        If tbDetalleAccion.Text <> "" Then
            strSQL = strSQL & "'" & tbDetalleAccion.Text & "',"
        Else
            strSQL = strSQL & " NULL, "
        End If
        strSQL = strSQL & cmbEstado.List(cmbEstado.ListIndex, 1) & ", " & _
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & ",'" & Me.cmbUsuario.Text & "',"
        If Not tbFecha.Visible Then
            strSQL = strSQL & " NULL);"
        Else
            If CDate(tbFecha.Text) >= Format(Now(), "yyyy/mm/dd") Then
                strSQL = strSQL & " #" & Format(tbFecha.Text, "yyyy/mm/dd") & "#)"
            Else
                MsgBox "Fecha Proxima no puede ser tiempo pasado"
                Exit Sub
            End If
        End If
        
            OpenDB
            On Error GoTo Handle:
            cnn.Execute (strSQL)
            closeRS
            
            Dim total As Range
            Dim r As Range
            Dim rowNumber As Long
            
            'Desbloquea Hoja
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
            
            With ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO)
                Set total = .Range("B:B")
                Set r = total.Cells.Find(What:=.Range("ID_SEGUIMIENTO"), LookAt:=xlWhole)
                If Not r Is Nothing Then
                    rowNumber = r.Row
                Else
                    rowNumber = .Range("FILA_ACCION")
                End If
                .Cells(rowNumber, .Range("FECHA_ACCION").Column) = Format(Now(), "yyyy/mm/dd")
                .Cells(rowNumber, .Range("FECHA_ACCION").Column + 1) = tbDetalleAccion.Text
                .Cells(rowNumber, .Range("FECHA_ACCION").Column + 2) = cmbEstado.List(cmbEstado.ListIndex, 0)
                If tbFecha.Visible Then
                    .Cells(rowNumber, .Range("FECHA_ACCION").Column + 3) = Format(tbFecha.Text, "yyyy/mm/dd")
                End If
            End With
            
            'Bloquea Hoja
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

            Unload Me
    Else
        MsgBox "Fecha no Ingresada o Errada"
    End If
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - btGuardar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Private Sub UserForm_Initialize()
    
    If NUEVA_ACCION Then
        tbDetalleAccion.Visible = True
        cmbEstado.Visible = True
        tbFecha.Visible = True
        btCalendario.Visible = True
        cmbUsuario.Visible = True
        btGuardar.Visible = True
        Label1.Visible = True
        Label3.Visible = True
        Label5.Visible = True
        
        tbAnterior.Height = 245
    Else
        tbDetalleAccion.Visible = False
        cmbEstado.Visible = False
        tbFecha.Visible = False
        btCalendario.Visible = False
        cmbUsuario.Visible = False
        btGuardar.Visible = False
        Label1.Visible = False
        Label3.Visible = False
        Label5.Visible = False
        
        tbAnterior.Height = 330
    End If

    strSQL = "SELECT * FROM DB_ESTADO_SEGUIMIENTO WHERE DB_ESTADO_SEGUIMIENTO.ANULADO = FALSE"
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        cmbEstado.Clear
        Dim cont As Integer
        cont = 0
        Do While Not rs.EOF
            cmbEstado.AddItem rs.Fields("NOMBRE_ESTADO_SEGUIMIENTO")
            cmbEstado.List(cont, 1) = rs.Fields("ID_ESTADO_SEGUIMIENTO")
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    cmbEstado.ListIndex = 0
    
    strSQL = "SELECT * FROM ((((DB_SEGUIMIENTO LEFT JOIN DB_ESTADO_SEGUIMIENTO ON " & _
    "DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK) " & _
    "RIGHT JOIN DB_CONDICION ON DB_CONDICION.ID_CONDICION = DB_SEGUIMIENTO.ID_CONDICION_FK) " & _
    "LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK) " & _
    "LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
    "WHERE DB_SEGUIMIENTO.ANULADO = FALSE AND ID_CONDICION_FK = " & _
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & " ORDER BY FECHA_ACCION ASC"
    'OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Me.Caption = Me.Caption & " | " & rs.Fields("CODIGO_SOCIO") & " | " & rs.Fields("NOMBRE_SOCIO") & _
        " | " & rs.Fields("SOLICITUD")
        Do While Not rs.EOF
            tbAnterior.Text = tbAnterior.Text & " " & ChrW(&H25A0) & " "
            If rs.Fields("FECHA_ACCION") = CDate("1999/01/01") Then
                tbAnterior.Text = tbAnterior.Text & "MIGRADO: "
            Else
                tbAnterior.Text = tbAnterior.Text & rs.Fields("FECHA_ACCION") & ": "
            End If
            tbAnterior.Text = tbAnterior.Text & rs.Fields("DETALLE_ACCION") & vbCrLf & "  " & ChrW(8594) & _
            " ESTADO: " & rs.Fields("NOMBRE_ESTADO_SEGUIMIENTO") & vbCrLf & "  " & ChrW(8594) & _
            " PROXIMA ACCION: " & rs.Fields("FECHA_PROXIMA_ACCION") & vbCrLf & "  " & ChrW(8594) & _
            " USUARIO: " & rs.Fields("USUARIO")
            
            tbAnterior.Text = tbAnterior.Text & vbCrLf & vbCrLf
            lbCondicion.Caption = lbCondicion.Caption & rs.Fields("DETALLE")
            cmbEstado.ListIndex = rs.Fields("ID_ESTADO_SEGUIMIENTO") - 1
            rs.MoveNext
        Loop
    End If
    cmbUsuario.Text = Application.UserName
    
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.Name & " - UserForm_Initialize", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

