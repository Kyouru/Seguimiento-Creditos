Option Explicit
Public cnn As New ADODB.Connection
Public cnn2 As New ADODB.Connection
Public cnn3 As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public strSQL As String
Public Const NOMBRE_HOJA_SEGUIMIENTO As String = "SEGUIMIENTO"
Public Const NOMBRE_HOJA_CARGA As String = "NUEVA CARGA"

Public Const NOMBRE_HOJA_TEMP As String = "TEMP"
Public Const NOMBRE_HOJA_TEMP2 As String = "TEMP2"
Public Const NOMBRE_HOJA_TEMP3 As String = "TEMP3"
Public Const NOMBRE_HOJA_TEMP4 As String = "TEMP4"
Public Const NOMBRE_HOJA_TEMP5 As String = "TEMP5"
Public Const NOMBRE_HOJA_TEMP6 As String = "TEMP6"
Public Const NOMBRE_HOJA_TEMP7 As String = "TEMP7"
Public Const NOMBRE_HOJA_TEMP8 As String = "TEMP8"
Public Const NOMBRE_HOJA_TEMP9 As String = "TEMP9"
Public Const NOMBRE_HOJA_TEMP10 As String = "TEMP10"
Public Const NOMBRE_HOJA_TEMP11 As String = "TEMP11"
Public Const NOMBRE_HOJA_TEMP12 As String = "TEMP12"
Public Const NOMBRE_HOJA_TEMP13 As String = "TEMP13"
Public Const NOMBRE_HOJA_TEMP14 As String = "TEMP14"
Public Const NOMBRE_HOJA_TEMP15 As String = "TEMP15"
Public Const NOMBRE_HOJA_TEMP16 As String = "TEMP16"

Public Const NOMBRE_HOJA_REPORTE_POSTERGADOS As String = "REPORTE REVISION"
Public Const NOMBRE_HOJA_L As String = "L"
Public Const SHEET_PASSWORD As String = "KyouruKenji"

Public Const MATRIZ As Boolean = True
Public Const NUEVA_ACCION As Boolean = True
Public Const POSTERGAR As Boolean = True

Public Const SEGUIMIENTO_GENERAL As Boolean = True
Public Const SEGUIMIENTO_COVENANT As Boolean = True
Public Const SEGUIMIENTO_GARANTIA As Boolean = True
Public Const SEGUIMIENTO_SEGURO As Boolean = True

Public Const NOMBRE_HOJA_REPORTE_SEG As String = "REPORTE SEGUIMIENTO"

Public g_objFSO As Scripting.FileSystemObject
Public g_scrText As Scripting.TextStream

Public colLabelEvent As Collection 'Collection of labels for event handling
Public colLabels As Collection     'Collection of the date labels
Public bSecondDate As Boolean      'True if finding second date
Public sActiveDay As String        'Last day selected
Public lDays As Long               'Number of days in month
Public lFirstDay As Long           'Day selected, e.g. 19th
Public lStartPos As Long
Public lSelMonth As Long           'The selected month
Public lSelYear As Long            'The selected year
Public lSelMonth1 As Long          'Used to check if same date is selected twice
Public lSelYear1 As Long           'Used to check if same date is selected twice


Sub Start()
UserForm1.Show
End Sub


Public Sub OpenDB()
    If cnn.State = adStateOpen Then cnn.Close
    On Error GoTo Handle
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & _
        ThisWorkbook.Sheets(NOMBRE_HOJA_L).Range("PATH_SEG") & "\" & ThisWorkbook.Sheets(NOMBRE_HOJA_L).Range("NAME_DB")
        cnn.Open
    Exit Sub
Handle:
    If cnn.Errors.count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, "Mulo1", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub closeRS()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If cnn.State = adStateOpen Then cnn.Close
    Set cnn = Nothing
End Sub

Public Sub OpenDB2()
    If cnn2.State = adStateOpen Then cnn2.Close
    cnn2.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & Application.ActiveWorkbook.Name
    cnn2.Open
End Sub

Public Sub closeRS2()
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If cnn2.State = adStateOpen Then cnn2.Close
    Set cnn2 = Nothing
End Sub

Public Sub OpenDB3()
    If cnn3.State = adStateOpen Then cnn3.Close
    cnn3.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ThisWorkbook.Sheets(NOMBRE_HOJA_L).Range("PATH_SEG") & "DATABASE\REGISTRO\Registro SISGO.xlsx"
    cnn3.Open
End Sub

Public Sub closeRS3()
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    If cnn3.State = adStateOpen Then cnn3.Close
    Set cnn3 = Nothing
End Sub

Public Function LogFile_WriteError(ByVal sRoutineName As String, _
                             ByVal sMessage As String)
Dim sText As String
Dim logPath As String
    logPath = ThisWorkbook.Path & "\log.txt"
   On Error GoTo ErrorHandler
   If (g_objFSO Is Nothing) Then
      Set g_objFSO = New FileSystemObject
   End If
   If (g_scrText Is Nothing) Then
      If (g_objFSO.FileExists(logPath) = False) Then
         Set g_scrText = g_objFSO.OpenTextFile(logPath, IOMode.ForWriting, True)
      Else
         Set g_scrText = g_objFSO.OpenTextFile(logPath, IOMode.ForAppending)
      End If
   End If
   sText = sText & Format(Date, "DD/MM/YYYY") & " " & Time() & ";"
   sText = sText & sRoutineName & ";"
   sText = sText & sMessage & ";"
   g_scrText.WriteLine sText
   g_scrText.Close
   Set g_scrText = Nothing
   Exit Function
ErrorHandler:
   Set g_scrText = Nothing
   Call MsgBox("No se pudo escribir en el fichero log", vbCritical, "LogFile_WriteError")
End Function

Public Sub Error_Handle(ByVal sRoutineName As String, _
                         ByVal sObject As String, _
                         ByVal currentStrSQL As String, _
                         ByVal sErrorNo As String, _
                         ByVal sErrorDescription As String)
Dim sMessage As String
   sMessage = sObject & ";" & currentStrSQL & ";" & sErrorNo & ";" & sErrorDescription & ";" & Application.UserName
   Call MsgBox(sErrorNo & vbCrLf & sErrorDescription, vbCritical, sRoutineName & " - " & sObject & " - Error")
   Call LogFile_WriteError(sRoutineName, sMessage)
End Sub

Public Function fechaStrStr(fechaDDMMYYYY As String)
    Dim splitfecha As Variant
    splitfecha = Split(fechaDDMMYYYY, "/")
    fechaStrStr = splitfecha(2) & "-" & splitfecha(1) & "-" & splitfecha(0)
End Function
