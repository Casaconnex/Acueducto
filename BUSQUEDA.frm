VERSION 5.00
Begin VB.Form BUSQUEDA 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BUSQUEDA"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "OCR A Extended"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Resultados"
      ForeColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   120
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.ListBox Lista 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   2205
      ItemData        =   "BUSQUEDA.frx":0000
      Left            =   120
      List            =   "BUSQUEDA.frx":000D
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton buscar 
      Caption         =   "Buscar..."
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Parametro 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elemento a buscar:"
      ForeColor       =   &H00868686&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2160
   End
End
Attribute VB_Name = "BUSQUEDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buscar_Click()
If SALVADO = False And BUSQ = True Then
    SALVADO = True
    Guardar_BaseDatos BusquedaE
End If
If Lista.ListIndex = 0 Then
    Data1.DatabaseName = App.Path + "\encuesta.mdb"
    Data1.RecordSource = "select * from tabla1 where cedula =" & Val(Parametro.Text)
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "Cédula no encontrada!", vbInformation, "CEDULA"
        Exit Sub
    End If

ElseIf Lista.ListIndex = 1 Then
    Data1.DatabaseName = App.Path + "\encuesta.mdb"
    Data1.RecordSource = "select * from tabla1 where codigo =" & Val(Parametro.Text)
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "Código no encontrado!", vbInformation, "CÓDIGO"
        Exit Sub
    End If

ElseIf Lista.ListIndex = 2 Then
    Data1.DatabaseName = App.Path + "\encuesta.mdb"
    Data1.RecordSource = "select * from tabla1 where ruta =" & Val(Parametro.Text)
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "Ruta no encontrada!", vbInformation, "RUTA"
        Exit Sub
    End If
    
End If
    
    Data1.Recordset.MoveLast
    Label2.Caption = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst
    CargarE
    SALVADO = False
    BUSQ = True
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
POSICIONLISTA = 0

End Sub

Private Sub Lista_Click()
    Parametro.SelStart = 0
    Parametro.SelLength = Len(Parametro.Text)
    Parametro.SetFocus
    POSICIONLISTA = Lista.ListIndex
End Sub

Private Sub Parametro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    buscar_Click
End If
End Sub
Private Sub CargarE()
If Data1.Recordset!nombre_suscriptor <> Null Or Data1.Recordset!nombre_suscriptor <> "" Then
BusquedaE.Nombre = Data1.Recordset!nombre_suscriptor
End If
BusquedaE.Cedula = Data1.Recordset!Cedula
BusquedaE.codigo = Data1.Recordset!codigo
BusquedaE.ruta = Data1.Recordset!ruta
BusquedaE.Ubicacion = Data1.Recordset!ubicacion_casa
BusquedaE.NoPisos = Data1.Recordset!no_pisos
If Data1.Recordset!direccion_predio <> Null Or Data1.Recordset!direccion_predio <> "" Then BusquedaE.direccion = Data1.Recordset!direccion_predio
If Data1.Recordset!numero_catastral <> Null Or Data1.Recordset!numero_catastral <> "" Then BusquedaE.NoCatastro = Data1.Recordset!numero_catastral
BusquedaE.EstadoPredio = Data1.Recordset!estado_predio
BusquedaE.NumeroPersonas = Data1.Recordset!numero_personas_casa
BusquedaE.NumeroFamilias = Data1.Recordset!numero_familias_casa
BusquedaE.NumeroNinos = Data1.Recordset!numero_menores_5
BusquedaE.Abastecimiento = Data1.Recordset!conectado_sistema
BusquedaE.OtraFuente = Data1.Recordset!otra_fuente
If Data1.Recordset!Cual <> Null Or Data1.Recordset!Cual <> "" Then
    BusquedaE.cual1 = Data1.Recordset!Cual
Else
    BusquedaE.cual1 = ""
End If
BusquedaE.OpinionCalidad = Data1.Recordset!calidad_agua
BusquedaE.OpinionCantidad = Data1.Recordset!cantidad_agua_suficiente
BusquedaE.UsoPredio = Data1.Recordset!uso_predio
BusquedaE.DiametroConexion = Data1.Recordset!diametro_conexion
BusquedaE.MaterialConexion = Data1.Recordset!tipo_materiales
BusquedaE.EstadoMedidor = Data1.Recordset!estado_medidor
If Data1.Recordset!numero_medidor <> Null Or Data1.Recordset!numero_medidor <> "" Then
    BusquedaE.NumeroMedidor = Data1.Recordset!numero_medidor
Else
    BusquedaE.NumeroMedidor = ""
End If
If Data1.Recordset!marca_medidor <> Null Or Data1.Recordset!marca_medidor <> "" Then
    BusquedaE.MarcaMedidor = Data1.Recordset!marca_medidor
Else
    BusquedaE.MarcaMedidor = ""
End If
BusquedaE.lectura = Data1.Recordset!lectura
BusquedaE.EstadoCajilla = Data1.Recordset!estado_cajilla
BusquedaE.TipoConexion = Data1.Recordset!tipo_conexion_usuario
BusquedaE.TanqueAlmacena = Data1.Recordset!tanque_almacenamiento
BusquedaE.AlmacenaAgua = Data1.Recordset!almacena_agua
BusquedaE.HierveAgua = Data1.Recordset!hierve_agua
If Data1.Recordset!reparacion_instalacion <> Null Or Data1.Recordset!reparacion_instalacion <> "" Then BusquedaE.ReparacionesInstalacion = Data1.Recordset!reparacion_instalacion
BusquedaE.QuienRepara = Data1.Recordset!quien_realiza
BusquedaE.Goteras = Data1.Recordset!gotea_llaves_grifos
BusquedaE.TipoServicioSanitario = Data1.Recordset!tipo_servicio_sanitario
BusquedaE.ProblemasInstalacionSanitarias = Data1.Recordset!problemas_instalacion
If Data1.Recordset!cuales <> Null Or Data1.Recordset!cuales <> "" Then BusquedaE.cual2 = Data1.Recordset!cuales
BusquedaE.InodoroLimpio = Data1.Recordset!inodoro_limpio
BusquedaE.EstadoCaseta = Data1.Recordset!estado_caseta
BusquedaE.TaponadaConexion = Data1.Recordset!taponada_conexion
BusquedaE.SolucionConexion = Data1.Recordset!solucion_problema
If Data1.Recordset!cual_solucion <> Null Or Data1.Recordset!cual_solucion <> "" Then BusquedaE.Cuales3 = Data1.Recordset!cual_solucion
BusquedaE.QueHaceBasuras = Data1.Recordset!destino_basuras
BusquedaE.BasurasCasa = Data1.Recordset!existencia_basuras_casa
BusquedaE.BarridoPorSemana = Data1.Recordset!veces_barrido_semana
BusquedaE.RecoleccionPorSemana = Data1.Recordset!veces_recoleCcion_semana
BusquedaE.OpinionEntidad = Data1.Recordset!opinion_administracion
BusquedaE.RespaldoEntidad = Data1.Recordset!respaldo_entidad
If Data1.Recordset!Observaciones <> Null Or Data1.Recordset!Observaciones <> "" Then
    BusquedaE.Observaciones = Data1.Recordset!Observaciones
Else
    BusquedaE.Observaciones = ""
End If
BusquedaE.DESHABITADA = Data1.Recordset!DESABITADA
Data1.Recordset.Delete
Cargar_Form
End Sub

Private Sub Timer1_Timer()
Lista.ListIndex = POSICIONLISTA
Lista.Enabled = True
Timer1.Enabled = False
Parametro.SelStart = 0
Parametro.SelLength = Len(Parametro.Text)
Parametro.SetFocus
End Sub
