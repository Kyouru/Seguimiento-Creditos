
Public Sub verificarVersion()
    Dim fso As FileSystemObject
    Dim txtStream As TextStream
    
    Dim ultimaVersion As String
    
    Set fso = New FileSystemObject
    Set txtStream = fso.OpenTextFile(ThisWorkbook.Sheets(NOMBRE_HOJA_L).Range("PATH_SEG") & "VERSIONES\version", ForReading, False)
    
    ultimaVersion = txtStream.ReadLine
    
    If [VERSION_SEGUIMIENTO] <> ultimaVersion Then
        Me.Range(Me.Range("VERSION_SEGUIMIENTO"), Me.Range("VERSION_SEGUIMIENTO").Offset(0, 1)).Interior.Pattern = xlSolid
        Me.Range(Me.Range("VERSION_SEGUIMIENTO"), Me.Range("VERSION_SEGUIMIENTO").Offset(0, 1)).Interior.Color = 49407
        If Left(Right(ultimaVersion, 3), 1) < Left(Right([VERSION_SEGUIMIENTO], 3), 1) Then
            Range("VERSION_SEGUIMIENTO").Offset(0, 1).Value = "Version Desarrollo"
        Else
            If Right(ultimaVersion, 1) < Right([VERSION_SEGUIMIENTO], 1) And Left(Right(ultimaVersion, 3), 1) = Left(Right([VERSION_SEGUIMIENTO], 3), 1) Then
                Range("VERSION_SEGUIMIENTO").Offset(0, 1).Value = "Version Desarrollo"
            Else
                Range("VERSION_SEGUIMIENTO").Offset(0, 1).Value = "No es la ultima Version"
            End If
        End If
    Else
        Me.Range(Me.Range("VERSION_SEGUIMIENTO"), Me.Range("VERSION_SEGUIMIENTO").Offset(0, 1)).Interior.Pattern = xlNone
        Range("VERSION_SEGUIMIENTO").Offset(0, 1).Value = ""
    End If
    
    txtStream.Close
    
    Set fso = Nothing
    Set txtStream = Nothing
    
End Sub
Private Sub btActualizarDesembolso_Click()
    actualizarPrestamos.Show (0)
End Sub

Private Sub btEditar_Click()
    If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FILA_ACCION") <> "NO" Then
        'Macro en Mulo2, abre formulario busqSeguimiento listando todos los Seguimientos de la Condicion
        ModificarAccion
    End If
End Sub

Private Sub btCalendario_Click()
    
    With ActiveSheet
        .Range("FECHA") = ""
        frmCalendario.Show
        'Verifica que se marco una fecha en el formulario frmCalendario
        If .Range("FECHA") <> "" Then
            'Desbloquea Hoja
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
            
            QuitarFiltros
            'Se filtra por la fecha seleccionada en el calendario
            ActiveSheet.Range("ENCABEZADO").AutoFilter Field:=18, Criteria1:="=" & Format(.Range("FECHA"), "YYYY-MM-DD")
            
            'Bloquea Hoja
            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
        End If
    End With
End Sub

Private Sub btHoy2_Click()
    'Desbloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
    QuitarFiltros
    'Se filtra por la fecha de Hoy
    
    ActiveWorkbook.Worksheets("SEGUIMIENTO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SEGUIMIENTO").Sort.SortFields.Add2 Key:=Range( _
        "CABECERA_FECHA_PROXIMA"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SEGUIMIENTO").Sort
        .SetRange Range("ENCABEZADO").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.Range("ENCABEZADO").AutoFilter Field:=13, Criteria1:= _
        "=COVENANT", Operator:=xlOr, Criteria2:="=SEGUIMIENTO"
    ActiveSheet.Range("ENCABEZADO").AutoFilter Field:=17, Criteria1:= _
        "=(EN BLANCO)", Operator:=xlOr, Criteria2:="=EN PROCESO"
    ActiveSheet.Range("ENCABEZADO").AutoFilter Field:=18, Criteria1:= _
        "<=" & Format(Date, "YYYY-MM-DD"), Operator:=xlAnd
        
    'Bloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
End Sub

Private Sub btMantenimiento_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SOCIO") = ""
    busqSocio.Show (0)
End Sub

Private Sub btMantenimientoLista_Click()
    mLista.Show
End Sub

Private Sub btNuevaAccion_Click()
    If ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FILA_ACCION") <> "NO" Then
        'Macro en Mulo2, abre formulario newAccion listando todos los Seguimientos anteriores de la Condicion y permite ingresar un nuevo seguimiento
        NuevaAccion
    End If
End Sub

Private Sub btPostergar_Click()
    
    If Selection.count = 1 Then
        If ActiveSheet.Name = "SEGUIMIENTO" Then
            If Range("A" & Selection.Row).Value <> "" And Range("A" & Selection.Row).Value <> "ID_CONDICION" Then
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") = Range("A" & Selection.Row).Value
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") = Range("B" & Selection.Row).Value
                ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_PRESTAMO") = Range("C" & Selection.Row).Value
                
                With ActiveSheet
                    .Range("FECHA") = ""
                    frmCalendario.Show
                    
                    'Verifica que se marco una fecha en el formulario frmCalendario
                    If .Range("FECHA") <> "" Then
                        strSQL = "SELECT * FROM ((((((DB_TIPO_CONDICION " & _
                        "LEFT JOIN DB_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK) " & _
                        "LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK) " & _
                        "LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK) " & _
                        "LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK) " & _
                        "LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK) " & _
                        "LEFT JOIN DB_SEGUIMIENTO ON DB_SEGUIMIENTO.ID_CONDICION_FK = DB_CONDICION.ID_CONDICION) " & _
                        "LEFT JOIN DB_ESTADO_SEGUIMIENTO ON DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK " & _
                        "WHERE ID_CONDICION = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_CONDICION") & _
                        " AND DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE " & _
                        " AND DB_SEGUIMIENTO.ID_SEGUIMIENTO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO")
                        
                        Dim detalle As String
                        Dim estado As Integer
                        Dim fecha As Date
                        Dim strfecha As String
                        Dim usuario As String
                        
                        OpenDB
                        On Error GoTo Handle:
                        rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                        If rs.RecordCount > 0 Then
                            detalle = rs.Fields("DETALLE_ACCION")
                            estado = rs.Fields("ID_ESTADO_SEGUIMIENTO_FK")
                            fecha = rs.Fields("FECHA_PROXIMA_ACCION")
                            
                            If IsNull(rs.Fields("FECHA_PROXIMA_ACCION")) Then
                                strfecha = "NULL"
                            Else
                                strfecha = "#" & Format(rs.Fields("FECHA_PROXIMA_ACCION"), "YYYY-MM-DD") & "#"
                            End If
                            
                            If IsNull(rs.Fields("USUARIO")) Then
                                usuario = "NULL"
                            Else
                                usuario = rs.Fields("USUARIO")
                            End If
                        End If
                        
                        closeRS
                        
                        If .Range("FECHA") > fecha Then
                        
                            strSQL = "UPDATE DB_SEGUIMIENTO SET FECHA_PROXIMA_ACCION = #" & Format(.Range("FECHA"), "YYYY-MM-DD") & "#" & _
                            " WHERE ID_SEGUIMIENTO = " & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO")
                            
                            OpenDB
                            On Error GoTo Handle:
                            cnn.Execute (strSQL)
                            closeRS
                            
                            strSQL = " INSERT INTO DB_SEGUIMIENTO_EDICION ( " & _
                                        "ID_SEGUIMIENTO_FK, " & _
                                        "DETALLE_ANTERIOR, " & _
                                        "ID_ESTADO_SEGUIMIENTO_ANTERIOR, " & _
                                        "USUARIO_EDITOR, " & _
                                        "FECHA_PROXIMA_ANTERIOR, " & _
                                        "FECHA_EDICION, " & _
                                        "USUARIO_ANTERIOR, " & _
                                        "ID_TIPO_EDICION_FK) " & _
                                    "VALUES (" & ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("ID_SEGUIMIENTO") & ", " & _
                                            "'" & detalle & "', " & _
                                            estado & ", " & _
                                            "'" & Application.UserName & "', " & _
                                            strfecha & ", " & _
                                            "#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "#, " & _
                                            "'" & usuario & "', " & _
                                            "2)"
                            OpenDB
                            On Error GoTo Handle:
                            cnn.Execute (strSQL)
                            closeRS
                        
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
                                .Cells(rowNumber, .Range("FECHA_ACCION").Column + 3) = Format(.Range("FECHA"), "yyyy/mm/dd")
                                
                            End With
                            
                            'Bloquea Hoja
                            ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
                        Else
                            MsgBox "Fecha Inferior a la accion actual"
                        End If
                    End If
                End With
                
            End If
        End If
    End If
    
    
Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.CodeName & " - btActualizar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    
    'Se Desconecta de la Base de Datos
    closeRS
    
    'Bloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

    
End Sub

Private Sub btReporteP_Click()
    frmCalendario2.Show
End Sub

Private Sub btSinFiltro_Click()
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
    QuitarFiltros
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
End Sub

Private Sub btActualizar_Click()

    strSQL = "SELECT ID_CONDICION, ID_SEGUIMIENTO, ID_PRESTAMO, NOMBRE_GRUPO, DOI, CODIGO_SOCIO, SOLICITUD," & _
    " NOMBRE_PRODUCTO, NOMBRE_SOCIO, NOMBRE_MONEDA, MONTO, FECHA_DESEMBOLSO, NOMBRE_TIPO, DETALLE," & _
    " FECHA_ACCION, DETALLE_ACCION, NOMBRE_ESTADO_SEGUIMIENTO, FECHA_PROXIMA_ACCION FROM (((((((((" & _
    " SELECT ID_CONDICION_FK, MAX(FECHA_ACCION) AS MAXFECHA FROM DB_SEGUIMIENTO" & _
    " WHERE DB_SEGUIMIENTO.ANULADO = FALSE GROUP BY ID_CONDICION_FK) AS R" & _
    " LEFT JOIN (SELECT * FROM DB_SEGUIMIENTO LEFT JOIN DB_ESTADO_SEGUIMIENTO ON" & _
    " DB_ESTADO_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO = DB_SEGUIMIENTO.ID_ESTADO_SEGUIMIENTO_FK) AS S" & _
    " ON S.ID_CONDICION_FK = R.ID_CONDICION_FK AND S.FECHA_ACCION = R.MAXFECHA)" & _
    " LEFT JOIN DB_CONDICION ON DB_CONDICION.ID_CONDICION = R.ID_CONDICION_FK)" & _
    " LEFT JOIN DB_TIPO_CONDICION ON DB_TIPO_CONDICION.ID_TIPO_CONDICION = DB_CONDICION.ID_TIPO_CONDICION_FK)" & _
    " LEFT JOIN DB_PRESTAMO ON DB_PRESTAMO.ID_PRESTAMO = DB_CONDICION.ID_PRESTAMO_FK)" & _
    " LEFT JOIN DB_SOCIO ON DB_SOCIO.ID_SOCIO = DB_PRESTAMO.ID_SOCIO_FK)" & _
    " LEFT JOIN DB_MONEDA ON DB_MONEDA.ID_MONEDA = DB_PRESTAMO.ID_MONEDA_FK)" & _
    " LEFT JOIN DB_GRUPO ON DB_GRUPO.ID_GRUPO = DB_SOCIO.ID_GRUPO_FK)" & _
    " LEFT JOIN DB_PRODUCTO ON DB_PRODUCTO.ID_PRODUCTO = DB_PRESTAMO.ID_PRODUCTO_FK)" & _
    " WHERE DB_SOCIO.ANULADO = FALSE AND DB_PRESTAMO.ANULADO = FALSE AND DB_CONDICION.ANULADO = FALSE"
    
    Dim primerTipo As Boolean: primerTipo = True
    
    If SEGUIMIENTO_GENERAL Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'SEGUIMIENTO' OR DB_TIPO_CONDICION.NOMBRE_TIPO = 'ANULADO' OR DB_TIPO_CONDICION.NOMBRE_TIPO = 'SIN CONDICIONES' OR DB_TIPO_CONDICION.NOMBRE_TIPO = 'DENEGADO'"
    End If
    
    If SEGUIMIENTO_GARANTIA Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'GARANTIA'"
    End If
    
    If SEGUIMIENTO_SEGURO Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'SEGURO'"
    End If
    
    If SEGUIMIENTO_COVENANT Then
        If primerTipo Then
            strSQL = strSQL & " AND ("
            primerTipo = False
        Else
            strSQL = strSQL & " OR "
        End If
        strSQL = strSQL & "DB_TIPO_CONDICION.NOMBRE_TIPO = 'COVENANT'"
    End If
    
    If Not primerTipo Then
        strSQL = strSQL & ")"
    End If
    
    If MATRIZ Then
        strSQL = strSQL & " ORDER BY NOMBRE_GRUPO, NOMBRE_SOCIO ASC"
    Else
        strSQL = strSQL & " ORDER BY NOMBRE_SOCIO, SOLICITUD ASC"
    End If
    
    'Desbloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Unprotect SHEET_PASSWORD
    
    verificarVersion
    
    QuitarFiltros
    
    'Limpiar Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range(Me.Range("dataSet"), Me.Range("dataSet").End(xlDown)).ClearContents
    
    'Se Conecta a la Base de Datos
    OpenDB
    On Error GoTo Handle:
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        Me.Range("dataSet").CopyFromRecordset rs
    End If

Handle:
    If cnn.Errors.count > 0 Then
        'Log del Error
        Call Error_Handle(cnn.Errors.Item(0).Source, Me.CodeName & " - btActualizar_Click", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    
    'Se Desconecta de la Base de Datos
    closeRS
    
    'Bloquea Hoja
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

End Sub

Private Sub QuitarFiltros()
    'Se agrega un criterio cualquiera al filtro para asegurarnos que el filtro exista
    ActiveSheet.Range("ENCABEZADO").AutoFilter Field:=1, Criteria1:="<>Cualquier Cosa"
    'Remueve el filtro
    Selection.AutoFilter
    'Agrega un filtro sin criterio de filtrado
    ActiveSheet.Range("ENCABEZADO").AutoFilter
End Sub

'Para Recolectar la fila activa, usado para Editar Condiciones y Nueva Accion (Seguimiento), Macros en Modulo2
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'En caso de no una unica celda activa
    ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FILA_ACCION") = "NO"
    
    'Evitar el error por desbordamiento (Muchas Celdas)
    On Error Resume Next
    If Target.count = 1 Then
        ThisWorkbook.Sheets(NOMBRE_HOJA_SEGUIMIENTO).Range("FILA_ACCION") = Target.Row
    End If
End Sub
