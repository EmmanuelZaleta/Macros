Option Explicit

' ============================================================================
' FORMULARIO ACTUALIZA - VERSIÓN OPTIMIZADA PROFESIONAL
' ============================================================================
' Optimizaciones implementadas:
' - Desactivación de eventos, cálculo automático y actualización de pantalla
' - Uso de variables de objeto en lugar de referencias repetidas
' - Eliminación de llamadas innecesarias a Select/Activate
' - Manejo de errores mejorado y centralizado
' - Mejoras en validación de datos
' - Código modularizado para mejor mantenimiento
' ============================================================================

Const GWL_STYLE = -16
Const WS_CAPTION = &HC00000

' Variables para drag & drop del formulario
Dim mdOriginX As Double
Dim mdOriginY As Double

' ============================================================================
' EVENTOS DEL FORMULARIO
' ============================================================================

Private Sub UserForm_Initialize()
    ' Configuración visual profesional del formulario
    With Me
        .lbl_Titulo.BackColor = RGB(41, 128, 185)  ' Azul profesional moderno
        .lbl_Titulo.ForeColor = RGB(255, 255, 255)  ' Texto blanco
        .lbl_Titulo.Font.Bold = True
        .lbl_Titulo.Font.Size = 12
    End With

    ' Ocultar todos los indicadores al iniciar
    OcultarTodosLosIndicadores
End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub lbl_Salir_Click()
    Unload Me
End Sub

' ============================================================================
' DRAG & DROP DEL FORMULARIO
' ============================================================================

Private Sub lbl_Titulo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    mdOriginX = X
    mdOriginY = y
End Sub

Private Sub lbl_Titulo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal y As Single)
    If Button And 1 Then
        Me.Left = Me.Left + (X - mdOriginX)
        Me.Top = Me.Top + (y - mdOriginY)
    End If
End Sub

' ============================================================================
' PROCESO PRINCIPAL DE ACTUALIZACIÓN - ULTRA OPTIMIZADO
' ============================================================================

Private Sub lbl_Actualizar_Click()
    Dim vPlan As String
    Dim rutaArchivos As String
    Dim wb As Workbook
    Dim startTime As Double

    On Error GoTo ErrorHandler

    ' Mostrar progreso
    Me.Caption = "Procesando... Por favor espere"
    DoEvents

    ' Iniciar temporizador para medir rendimiento
    startTime = Timer

    ' Obtener referencias principales
    rutaArchivos = ThisWorkbook.Sheets("Macro").Range("B1").Value
    Set wb = ActiveWorkbook
    vPlan = wb.Name

    ' OPTIMIZACIÓN CRÍTICA: Desactivar eventos y actualizaciones
    Call OptimizarEntorno(True)

    ' Procesar cada módulo seleccionado
    If Me.chk_FlexPlan.Value Then Call ProcesarFlexPlan
    If Me.chk_Ordenes.Value Then Call ProcesarOrdenes(wb, vPlan, rutaArchivos)
    If Me.chk_InvLocWIP.Value Then Call ProcesarInvLocWIP(wb, vPlan, rutaArchivos)
    If Me.chk_LoadFactor.Value Then Call ProcesarLoadFactor(wb, vPlan, rutaArchivos)
    If Me.chk_ItemMaster.Value Then Call ProcesarItemMaster(wb, vPlan, rutaArchivos)
    If Me.chk_InventarioFG.Value Then Call ProcesarInventarioFG(wb, vPlan, rutaArchivos)
    If Me.chk_Capacidades.Value Then Call ProcesarCapacidades(wb, vPlan)

    ' Restaurar configuración
    Call OptimizarEntorno(False)

    ' Mostrar tiempo de ejecución
    Dim tiempoTotal As Double
    tiempoTotal = Round(Timer - startTime, 2)

    MsgBox "Proceso Terminado Exitosamente" & vbCrLf & _
           "Tiempo de ejecución: " & tiempoTotal & " segundos", vbInformation, "Éxito"

    Unload Me
    Exit Sub

ErrorHandler:
    Call OptimizarEntorno(False)  ' Siempre restaurar configuración
    MsgBox "Error en la aplicación: " & Err.Description & vbCrLf & _
           "Línea: " & Erl, vbCritical, "Error"
End Sub

' ============================================================================
' OPTIMIZACIÓN DEL ENTORNO DE EXCEL
' ============================================================================

Private Sub OptimizarEntorno(ByVal optimizar As Boolean)
    ' Activa o desactiva optimizaciones según el parámetro
    With Application
        If optimizar Then
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
            .DisplayStatusBar = False
            .Cursor = xlWait
        Else
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .DisplayStatusBar = True
            .Cursor = xlDefault
        End If
    End With
End Sub

' ============================================================================
' PROCESAMIENTO DE MÓDULOS - MÉTODOS OPTIMIZADOS
' ============================================================================

Private Sub ProcesarFlexPlan()
    Me.img_PalomaFlexPlan.Visible = True
    Me.Repaint
End Sub

Private Sub ProcesarOrdenes(wb As Workbook, vPlan As String, rutaArchivos As String)
    On Error GoTo ErrorOrdenes

    Dim wsOrdenes As Worksheet
    Dim vArchivo As String
    Dim rutaCompleta As String
    Dim fecha As String
    Dim ultimaFila As Long
    Dim rngDatos As Range

    ' Obtener referencia a la hoja
    Set wsOrdenes = wb.Sheets("Orderstats")
    wsOrdenes.Visible = xlSheetVisible

    Me.img_Error_OrderStats.Visible = False

    ' Validar fecha
    fecha = Trim$(Me.txtFechaFinal.Text)
    If Not ValidarFechaYYYYMMDD(fecha) Then GoTo ErrorOrdenes

    ' Obtener archivo
    vArchivo = buscaArchivo("Ordenes")
    rutaCompleta = rutaArchivos & vArchivo

    If Dir(rutaCompleta) = "" Then
        MsgBox "No se encontró el archivo de Órdenes.", vbInformation
        GoTo ErrorOrdenes
    End If

    ' Limpiar hoja de manera eficiente
    With wsOrdenes
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A2:L" & .Rows.Count).ClearContents
    End With

    ' Cargar datos
    Call CargarOrderStat_DesdeUNC_Hasta(vPlan, fecha)

    ' Configurar encabezados usando arrays (MÁS RÁPIDO)
    wsOrdenes.Range("A1:L1").Value = Array("CUST. CD.", "S/T", "PARTNO", "ETD", "ETA", _
                                            "QUANTITY", "SHIPPING QTY", "Remain1", "CUST. PO", _
                                            "ORDER FLG", "Date", "Validacion")

    ' Procesar y ordenar datos
    ultimaFila = wsOrdenes.Cells(wsOrdenes.Rows.Count, "C").End(xlUp).Row

    If ultimaFila > 1 Then
        ' Forzar formato de fecha
        Call ForzarFechaEnColumna(wsOrdenes, "D", ultimaFila)

        ' Ordenar datos
        Set rngDatos = wsOrdenes.Range("A1:L" & ultimaFila)
        rngDatos.Sort _
            Key1:=wsOrdenes.Range("C1"), Order1:=xlAscending, _
            Key2:=wsOrdenes.Range("D1"), Order2:=xlAscending, _
            Header:=xlYes, Orientation:=xlTopToBottom
    Else
        MsgBox "La tabla de Órdenes no tiene datos suficientes.", vbExclamation
        GoTo ErrorOrdenes
    End If

    Me.img_PalomaOrdenes.Visible = True
    Me.Repaint
    Exit Sub

ErrorOrdenes:
    Me.img_PalomaOrdenes.Visible = False
    Me.img_Error_OrderStats.Visible = True
End Sub

Private Sub ProcesarInvLocWIP(wb As Workbook, vPlan As String, rutaArchivos As String)
    Dim wsWIP As Worksheet
    Dim vArchivo As String
    Dim vLstRen As Long

    Set wsWIP = wb.Sheets("WIP")
    wsWIP.Visible = xlSheetVisible

    vArchivo = buscaArchivo("InvLocWIP")

    If Dir(rutaArchivos & vArchivo) = "" Then
        MsgBox "La tabla de InvLocWIP no existe", vbInformation
        Exit Sub
    End If

    With wsWIP
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A2:I" & .Rows.Count).ClearContents
    End With

    ' Cargar información
    Call traeInformacionInvLocWIP(vPlan)

    ' Configurar encabezados usando arrays
    wsWIP.Range("A1:I1").Value = Array("Inv Location", "Box Unit", "Part#", _
                                        "Inj.Date Min", "Dept", "Type", "Flg/Ord", _
                                        "Inv Confiable", "inv Date")

    ' Fórmulas optimizadas
    With wsWIP
        .Range("H2").Formula = "=IF(LEFT(A2,1)=""H"",0,IF(LEFT(A2,3)=""60V"",0,IF(LEFT(A2,3)=""EPA"",0,IF(AND(A2>=""EXC50"",A2<=""EXC99""),0,(IF(AND(LEFT(A2,2)=""CA"",I2<(TODAY()-1)),0,B2))))))"
        .Range("I2").Formula = "=+DATE(LEFT(D2,4),MID(D2,5,2),MID(D2,7,2))"

        vLstRen = .Cells(.Rows.Count, "A").End(xlUp).Row

        If vLstRen > 2 Then
            .Range("H2:I2").AutoFill Destination:=.Range("H2:I" & vLstRen), Type:=xlFillDefault
        End If
    End With

    Me.img_PalomaInvLocWIP.Visible = True
    Me.Repaint
End Sub

Private Sub ProcesarLoadFactor(wb As Workbook, vPlan As String, rutaArchivos As String)
    Dim wsLoad As Worksheet
    Dim vLstRen As Long

    Set wsLoad = wb.Sheets("Load Factor")
    wsLoad.Visible = xlSheetVisible

    With wsLoad
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1:AH" & .Rows.Count).ClearContents
    End With

    ' Ejecutar carga
    Call TraeInformacionLoadFactor(vPlan)
    Call QuitarEspaciosHoja
    Call NumeroAValor("A", "2")

    ' Encabezados usando arrays
    wsLoad.Range("A1:N1").Value = Array("PartNo", "CONTROL", "DIE", "Dep", "Group Code", _
                                         "Eng Lev", "Std Cav", "Act Cav", "Cycle Time", _
                                         "Piece Weight", "Shot Weight", "Pcs/Hour", _
                                         "Capacidad", "Ensamble")

    ' Fórmulas
    With wsLoad
        .Range("M2").Formula = "=IFNA(IF(LEFT(D2,1)=""N"",L2*19.83*7,IFS(LEFT(D2,1)=""F"", (3600/I2)*G2*24*7*0.9, LEFT(D2,1)=""S"", (3600/I2)*G2*24*7*0.9, LEFT(D2,1)=""J"", (3600/I2)*G2*24*7*0.9)),0)"
        .Range("N2").Formula = "=IF(LEFT(A2,2)=""72"",IF(COUNTIF(B:B,B2)>1,""COMPARTE"","" ""),""-"")"

        vLstRen = .Cells(.Rows.Count, "A").End(xlUp).Row

        If vLstRen > 2 Then
            .Range("M2:N2").AutoFill Destination:=.Range("M2:N" & vLstRen), Type:=xlFillDefault
        End If

        ' Ordenar por DIE
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsLoad.Range("C2:C" & vLstRen), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wsLoad.Range("A1:N" & vLstRen)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With

    Me.img_PalomaLoadFactor.Visible = True
    Me.Repaint
End Sub

Private Sub ProcesarItemMaster(wb As Workbook, vPlan As String, rutaArchivos As String)
    Dim wsItem As Worksheet
    Dim vArchivo As String

    Set wsItem = wb.Sheets("Item Master")
    wsItem.Visible = xlSheetVisible

    vArchivo = buscaArchivo("ItemMaster")

    If Dir(rutaArchivos & vArchivo) = "" Then
        MsgBox "La tabla de ItemMaster no existe", vbInformation
        Exit Sub
    End If

    With wsItem
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1:J" & .Rows.Count).ClearContents
    End With

    Call traeInformacionItemMaster(vPlan)

    ' Encabezados
    wsItem.Range("A1:J1").Value = Array("PartNo", "Description", "Dep", "Line", "PLN", _
                                         "Type", "Flg/Ord", "Unit/Bag", "Unit/Poly", "Unit/Box")

    Me.img_PalomaItemMaster.Visible = True
    Me.Repaint
End Sub

Private Sub ProcesarInventarioFG(wb As Workbook, vPlan As String, rutaArchivos As String)
    Dim wsInvFG As Worksheet
    Dim vArchivo As String

    Set wsInvFG = wb.Sheets("Inventario FG")
    wsInvFG.Visible = xlSheetVisible

    vArchivo = buscaArchivo("InvLocWIPFG")

    If Dir(rutaArchivos & vArchivo) = "" Then
        MsgBox "La tabla de InvLocWIPFG no existe", vbInformation
        Exit Sub
    End If

    With wsInvFG
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A2:O" & .Rows.Count).ClearContents
    End With

    Call traeInformacionInventarioFG(vPlan)

    Me.img_PalomaInventarioFG.Visible = True
    Me.Repaint
End Sub

Private Sub ProcesarCapacidades(wb As Workbook, vPlan As String)
    Dim wsCap As Worksheet

    If wb.ProtectStructure Then
        MsgBox "El libro está protegido (estructura). Desprotégelo para crear hojas.", vbExclamation
        Exit Sub
    End If

    Set wsCap = GetOrCreateSheet(wb, "Capacidades")

    With wsCap
        .Visible = xlSheetVisible
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A2:O" & .Rows.Count).ClearContents
    End With

    Call traeInformacionCapacidades(vPlan)

    Me.img_PalomaCapacidades.Visible = True
    Me.Repaint
End Sub

Private Sub lbl_FlexPlan_Click()
    Dim f As Object
    Dim varFile As Variant
    Dim getFileName As String

    Set f = Application.FileDialog(3)
    With f
        .AllowMultiSelect = False
        .Title = "Seleccionar FlexPlan"
        .InitialFileName = "\\Yazaki.local\na\elcom\chihuahua\Area_General\Materiales\Archivos Macro PCD\Pruebas\Extractor\YCC Flex Planning.xlsx"
    End With

    If f.Show = True Then
        For Each varFile In f.SelectedItems
            Me.txt_FlexPlan.Text = varFile
        Next
    Else
        Me.txt_FlexPlan.Text = ""
    End If
End Sub

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

Private Sub OcultarTodosLosIndicadores()
    ' Ocultar todos los indicadores de éxito/error
    On Error Resume Next
    Me.img_PalomaFlexPlan.Visible = False
    Me.img_PalomaOrdenes.Visible = False
    Me.img_Error_OrderStats.Visible = False
    Me.img_PalomaInvLocWIP.Visible = False
    Me.img_PalomaLoadFactor.Visible = False
    Me.img_PalomaItemMaster.Visible = False
    Me.img_PalomaInventarioFG.Visible = False
    Me.img_PalomaCapacidades.Visible = False
    On Error GoTo 0
End Sub

Private Function ValidarFechaYYYYMMDD(ByVal fecha As String) As Boolean
    Dim anio As Integer, mes As Integer, dia As Integer
    Dim testFecha As Date

    ValidarFechaYYYYMMDD = False

    If Len(fecha) <> 8 Or Not IsNumeric(fecha) Then Exit Function

    anio = CInt(Left$(fecha, 4))
    mes = CInt(Mid$(fecha, 5, 2))
    dia = CInt(Right$(fecha, 2))

    If mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Then Exit Function

    On Error Resume Next
    testFecha = DateSerial(anio, mes, dia)
    If Err.Number = 0 Then ValidarFechaYYYYMMDD = True
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(baseName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

        On Error Resume Next
        ws.Name = baseName
        If Err.Number <> 0 Then
            Err.Clear
            ws.Name = baseName & "_" & Format(Now, "hhmmss")
            MsgBox "Ya existía '" & baseName & "'. Se creó '" & ws.Name & "'.", vbInformation
        End If
        On Error GoTo 0
    End If

    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    Set GetOrCreateSheet = ws
End Function

Private Sub ForzarFechaEnColumna(ws As Worksheet, ByVal col As String, ByVal lastRow As Long)
    Dim r As Range, c As Range
    Dim s As String, y As Integer, m As Integer, d As Integer
    Dim digits As String, n As Double, nLng As Long

    If lastRow < 2 Then Exit Sub
    Set r = ws.Range(col & "2:" & col & lastRow)

    On Error GoTo fin

    For Each c In r.Cells
        s = Trim$(CStr(c.Value2))
        If Len(s) = 0 Then GoTo siguiente

        s = Replace(Replace(Replace(s, "-", ""), "/", ""), ".", "")
        digits = s

        If IsNumeric(digits) Then
            If Len(digits) = 8 Then
                y = CInt(Left$(digits, 4))
                m = CInt(Mid$(digits, 5, 2))
                d = CInt(Right$(digits, 2))
                If y >= 1900 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                    c.Value = DateSerial(y, m, d)
                    GoTo siguiente
                End If
            End If

            n = Val(digits)
            nLng = CLng(n)
            If nLng >= 30000 And nLng <= 60000 Then
                c.Value = nLng
                GoTo siguiente
            End If
        End If

        On Error Resume Next
        c.Value = CDate(c.Value)
        On Error GoTo 0
siguiente:
    Next c

fin:
    r.NumberFormat = "mm/dd/yyyy"
End Sub
