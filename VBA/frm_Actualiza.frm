VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Actualiza 
   Caption         =   "Centro de actualización"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9180
   OleObjectBlob   =   "frm_Actualiza.frx":0000
   StartUpPosition =   1  'CenterOwner
   BackColor       =   &H00F9F5F2&
   Begin MSForms.Label lblTitulo 
      Caption         =   "Actualizar registros"
      Height          =   360
      Left            =   240
      Top             =   180
      Width           =   5400
      FontBold        =   -1  'True
      FontSize        =   16
      ForeColor       =   &H00202020&
      BackStyle       =   0  'fmBackStyleTransparent
   End
   Begin MSForms.TextBox txtBuscar 
      Height          =   360
      Left            =   240
      Top             =   780
      Width           =   3000
      FontName        =   "Segoe UI"
      FontSize        =   10
      ForeColor       =   &H00202020&
      SpecialEffect   =   2  'fmSpecialEffectSunken
      BorderStyle     =   1  'fmBorderStyleSingle
      BackColor       =   &H00FFFFFF&
      TabIndex        =   0
   End
   Begin MSForms.ComboBox cboCampo 
      Height          =   360
      Left            =   3360
      Top             =   780
      Width           =   2100
      FontName        =   "Segoe UI"
      FontSize        =   10
      ForeColor       =   &H00202020&
      Style           =   2  'fmStyleDropDownList
      BackColor       =   &H00FFFFFF&
      TabIndex        =   1
   End
   Begin MSForms.CommandButton cmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   360
      Left            =   5580
      Top             =   780
      Width           =   1200
      BackColor       =   &H00FFFFFF&
      FontName        =   "Segoe UI"
      FontSize        =   10
      TabIndex        =   2
   End
   Begin MSForms.CommandButton cmdExportar 
      Caption         =   "Exportar..."
      Height          =   360
      Left            =   6840
      Top             =   780
      Width           =   1080
      BackColor       =   &H00FFFFFF&
      FontName        =   "Segoe UI"
      FontSize        =   10
      TabIndex        =   3
   End
   Begin MSForms.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   7980
      Top             =   780
      Width           =   900
      BackColor       =   &H00FFFFFF&
      FontName        =   "Segoe UI"
      FontSize        =   10
      TabIndex        =   4
   End
   Begin MSForms.ListBox lstRegistros 
      Height          =   4200
      Left            =   240
      Top             =   1320
      Width           =   4200
      ColumnHeads     =   0   'False
      ColumnCount     =   4
      FontName        =   "Segoe UI"
      FontSize        =   10
      ForeColor       =   &H00202020&
      IntegralHeight  =   0   'False
      ListStyle       =   0  'fmListStylePlain
      MultiSelect     =   0  'fmMultiSelectSingle
      TabIndex        =   5
   End
   Begin MSForms.Frame fraEdicion 
      Height          =   4320
      Left            =   4560
      Top             =   1320
      Width           =   4380
      Caption         =   "Detalle del registro"
      BackColor       =   &H00FFFFFF&
      FontName        =   "Segoe UI"
      FontSize        =   10
      TabIndex        =   6
      ScrollBars      =   2  'fmScrollBarsVertical
      ScrollHeight    =   4320
   End
   Begin MSForms.CommandButton cmdGuardar 
      Caption         =   "Guardar cambios"
      Height          =   420
      Left            =   4560
      Top             =   5760
      Width           =   1860
      BackColor       =   &H0073C2F9&
      ForeColor       =   &H00FFFFFF&
      FontName        =   "Segoe UI"
      FontSize        =   10
      TabIndex        =   7
   End
   Begin MSForms.CommandButton cmdRestablecer 
      Caption         =   "Restablecer"
      Height          =   420
      Left            =   6480
      Top             =   5760
      Width           =   1260
      BackColor       =   &H00FFFFFF&
      FontName        =   "Segoe UI"
      FontSize        =   10
      TabIndex        =   8
   End
   Begin MSForms.Label lblEstado 
      Height          =   300
      Left            =   240
      Top             =   5880
      Width           =   4200
      Caption         =   "Listo"
      BackStyle       =   0  'fmBackStyleTransparent
      FontName        =   "Segoe UI"
      FontSize        =   9
      ForeColor       =   &H00666666&
   End
End
Attribute VB_Name = "frm_Actualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NOMBRE_TABLA As String = "tblActualizacion"
Private Const COLUMNA_CLAVE As Long = 1

Private mTabla As TablaDatos
Private mIndices() As Long
Private mCampos As Collection
Private mFilaActiva As Long
Private mEstaCargando As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo manejador
    mEstaCargando = True
    ConfigurarTema
    CargarTablaDatos
    mEstaCargando = False
    Exit Sub
manejador:
    MsgBox "No fue posible inicializar el formulario: " & Err.Description, vbCritical
    Unload Me
End Sub

Private Sub ConfigurarTema()
    Me.Caption = "Centro de actualización"
    lblTitulo.Font.Name = "Segoe UI"
    lblTitulo.Font.Bold = True
    lblTitulo.ForeColor = &H00202020
    lblEstado.Caption = "Listo"
    cboCampo.Clear
    cboCampo.BackColor = &H00FFFFFF
    txtBuscar.BackColor = &H00FFFFFF
End Sub

Private Sub CargarTablaDatos()
    Dim i As Long

    mTabla = CargarTabla(NOMBRE_TABLA)

    cboCampo.Clear
    If IsArray(mTabla.Encabezados) Then
        For i = LBound(mTabla.Encabezados) To UBound(mTabla.Encabezados)
            cboCampo.AddItem mTabla.Encabezados(i)
        Next i
        If cboCampo.ListCount > 0 Then cboCampo.ListIndex = 0
    End If

    ConstruirControlesDeEdicion
    RefrescarListado
End Sub

Private Sub ConstruirControlesDeEdicion()
    Dim etiqueta As MSForms.Label
    Dim campo As MSForms.TextBox
    Dim topActual As Double
    Dim i As Long

    Set mCampos = New Collection

    Do While fraEdicion.Controls.Count > 0
        fraEdicion.Controls.Remove fraEdicion.Controls(0).Name
    Loop

    topActual = 480

    If Not IsArray(mTabla.Encabezados) Then Exit Sub

    For i = LBound(mTabla.Encabezados) To UBound(mTabla.Encabezados)
        Set etiqueta = fraEdicion.Controls.Add("Forms.Label.1", "lblCampo" & i, True)
        etiqueta.Caption = mTabla.Encabezados(i)
        etiqueta.Font.Bold = True
        etiqueta.Left = 180
        etiqueta.Top = topActual - 240
        etiqueta.Width = fraEdicion.InsideWidth - 360
        etiqueta.ForeColor = &H00303030

        Set campo = fraEdicion.Controls.Add("Forms.TextBox.1", "txtCampo" & i, True)
        campo.Left = 180
        campo.Top = topActual
        campo.Width = fraEdicion.InsideWidth - 360
        campo.Height = 300
        campo.Tag = CStr(i)
        campo.Font.Name = "Segoe UI"
        campo.Font.Size = 10
        campo.BackColor = &H00F4F8FD
        campo.BorderStyle = fmBorderStyleSingle
        campo.SpecialEffect = fmSpecialEffectSunken
        campo.EnterKeyBehavior = False
        campo.TabIndex = i - 1
        campo.Locked = (i = COLUMNA_CLAVE)
        mCampos.Add campo, CStr(i)

        topActual = topActual + 480
    Next i

    fraEdicion.ScrollHeight = topActual + 240
End Sub

Private Sub RefrescarListado()
    Dim datos As Variant
    Dim columnas As Long
    Dim filas As Long
    Dim filtro As String
    Dim campoIndice As Long

    lstRegistros.Clear
    lblEstado.Caption = "Cargando..."

    If Not EsMatrizValida(mTabla.Valores) Then
        lblEstado.Caption = "Sin registros disponibles."
        Exit Sub
    End If

    datos = mTabla.Valores
    filas = UBound(datos, 1)
    columnas = UBound(datos, 2)

    filtro = Trim$(txtBuscar.Value)
    If cboCampo.ListIndex < 0 Then
        campoIndice = COLUMNA_CLAVE
    Else
        campoIndice = cboCampo.ListIndex + 1
    End If

    Dim buffer() As Variant
    Dim indice() As Long
    Dim fila As Long
    Dim col As Long
    Dim contador As Long
    ReDim buffer(0 To filas - 1, 0 To columnas - 1)
    ReDim indice(0 To filas - 1)

    For fila = 1 To filas
        If filtro = "" Or InStr(1, CStr(datos(fila, campoIndice)), filtro, vbTextCompare) > 0 Then
            For col = 1 To columnas
                buffer(contador, col - 1) = CStr(datos(fila, col))
            Next col
            indice(contador) = fila
            contador = contador + 1
        End If
    Next fila

    If contador = 0 Then
        lblEstado.Caption = "No se encontraron coincidencias."
        Erase mIndices
        Exit Sub
    End If

    ReDim Preserve buffer(0 To contador - 1, 0 To columnas - 1)
    ReDim Preserve indice(0 To contador - 1)

    mIndices = indice
    lstRegistros.ColumnCount = columnas
    lstRegistros.List = buffer
    lblEstado.Caption = contador & " registro(s) mostrados."
End Sub

Private Sub lstRegistros_Click()
    MostrarDetalleSeleccionado
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MostrarDetalleSeleccionado
End Sub

Private Sub MostrarDetalleSeleccionado()
    Dim idx As Long
    Dim filaOrigen As Long
    Dim col As Long
    Dim campo As MSForms.TextBox

    If lstRegistros.ListIndex < 0 Then Exit Sub
    If Not EsMatrizValida(mTabla.Valores) Then Exit Sub
    If Not HayIndices() Then Exit Sub

    idx = lstRegistros.ListIndex
    If idx < LBound(mIndices) Or idx > UBound(mIndices) Then Exit Sub

    filaOrigen = mIndices(idx)
    mFilaActiva = filaOrigen

    For col = LBound(mTabla.Encabezados) To UBound(mTabla.Encabezados)
        Set campo = mCampos(CStr(col))
        campo.Value = CStr(mTabla.Valores(filaOrigen, col))
    Next col

    lblEstado.Caption = "Editando registro " & mTabla.Valores(filaOrigen, COLUMNA_CLAVE)
End Sub

Private Function HayIndices() As Boolean
    On Error GoTo sinIndices
    If UBound(mIndices) >= LBound(mIndices) Then HayIndices = True
    Exit Function
sinIndices:
    HayIndices = False
End Function

Private Sub txtBuscar_Change()
    If mEstaCargando Then Exit Sub
    RefrescarListado
End Sub

Private Sub cboCampo_Change()
    If mEstaCargando Then Exit Sub
    RefrescarListado
End Sub

Private Sub cmdRefrescar_Click()
    On Error GoTo manejador
    CargarTablaDatos
    lblEstado.Caption = "Datos actualizados desde la hoja."
    Exit Sub
manejador:
    lblEstado.Caption = Err.Description
End Sub

Private Sub cmdExportar_Click()
    Dim destino As Variant

    If Not EsMatrizValida(mTabla.Valores) Then
        MsgBox "No hay información para exportar.", vbInformation
        Exit Sub
    End If

    destino = Application.GetSaveAsFilename(InitialFileName:="Registros.xlsx", _
                                            FileFilter:="Libro de Excel (*.xlsx), *.xlsx")
    If VarType(destino) = vbBoolean And destino = False Then Exit Sub

    ExportarRango CStr(destino), mTabla.Valores, mTabla.Encabezados
    MsgBox "Se exportaron " & CStr(UBound(mTabla.Valores, 1)) & " registros.", vbInformation
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdRestablecer_Click()
    If Not EsMatrizValida(mTabla.Valores) Then Exit Sub
    If mFilaActiva = 0 Then Exit Sub
    Dim col As Long
    Dim campo As MSForms.TextBox

    For col = LBound(mTabla.Encabezados) To UBound(mTabla.Encabezados)
        Set campo = mCampos(CStr(col))
        campo.Value = CStr(mTabla.Valores(mFilaActiva, col))
    Next col
    lblEstado.Caption = "Valores originales restaurados."
End Sub

Private Sub cmdGuardar_Click()
    Dim valores As Variant
    Dim col As Long
    Dim campo As MSForms.TextBox
    Dim clave As Variant

    If mFilaActiva = 0 Then
        MsgBox "Seleccione un registro para actualizar.", vbExclamation
        Exit Sub
    End If

    valores = CrearVectorFila(UBound(mTabla.Encabezados))

    For col = LBound(mTabla.Encabezados) To UBound(mTabla.Encabezados)
        Set campo = mCampos(CStr(col))
        valores(1, col) = campo.Value
    Next col

    clave = valores(1, COLUMNA_CLAVE)
    If EsValorVacio(clave) Then
        MsgBox "La clave del registro no puede quedar vacía.", vbExclamation
        Exit Sub
    End If

    On Error GoTo manejador
    If ActualizarFilaTabla(NOMBRE_TABLA, COLUMNA_CLAVE, clave, valores) Then
        For col = LBound(mTabla.Encabezados) To UBound(mTabla.Encabezados)
            mTabla.Valores(mFilaActiva, col) = valores(1, col)
        Next col
        RefrescarListado
        SeleccionarPorClave CStr(clave)
        lblEstado.Caption = "Registro actualizado correctamente."
    Else
        MsgBox "No fue posible ubicar el registro en la hoja.", vbExclamation
    End If
    Exit Sub
manejador:
    MsgBox "Se produjo un error al guardar: " & Err.Description, vbCritical
End Sub

Private Sub SeleccionarPorClave(ByVal clave As String)
    Dim i As Long
    If Not HayIndices() Then Exit Sub
    For i = LBound(mIndices) To UBound(mIndices)
        If CStr(mTabla.Valores(mIndices(i), COLUMNA_CLAVE)) = clave Then
            lstRegistros.ListIndex = i
            MostrarDetalleSeleccionado
            Exit For
        End If
    Next i
End Sub
