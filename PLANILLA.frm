VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PLANILLA 
   BackColor       =   &H00000000&
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   5760
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   5280
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   4800
   End
   Begin VB.CommandButton resumen 
      Caption         =   "Imprimir Reporte"
      Height          =   855
      Index           =   1
      Left            =   10080
      Picture         =   "PLANILLA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton resumen 
      Caption         =   "Imprimir Resumen"
      Height          =   855
      Index           =   0
      Left            =   10080
      Picture         =   "PLANILLA.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Timer Timer11 
      Left            =   1320
      Top             =   6120
   End
   Begin VB.Timer Timer10 
      Left            =   1320
      Top             =   5520
   End
   Begin VB.Timer Timer9 
      Left            =   5160
      Top             =   4680
   End
   Begin VB.Timer Timer8 
      Left            =   4560
      Top             =   4680
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   4680
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6480
      Top             =   7320
   End
   Begin VB.Frame RESU 
      BackColor       =   &H00000000&
      Caption         =   "RESULTADOS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   2880
      TabIndex        =   1
      Top             =   5400
      Width           =   2775
      Begin MSFlexGridLib.MSFlexGrid PORCENTAJES 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483641
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   6120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   5520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   5520
   End
   Begin MSFlexGridLib.MSFlexGrid MALLA 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   6000
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   6480
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   8640
      Picture         =   "PLANILLA.frx":0204
      Stretch         =   -1  'True
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   8640
      Picture         =   "PLANILLA.frx":2386
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   1695
   End
End
Attribute VB_Name = "PLANILLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Data1.DatabaseName = App.Path + "\Encuesta.mdb"
Data2.DatabaseName = App.Path + "\Encuesta.mdb"
PORCENTAJES.Cols = 4
PORCENTAJES.Rows = 75
CARGAR_MALLA M
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For Y = 0 To 1
  resumen(Y).BackColor = &H8000000F
Next Y
End Sub

Private Sub Form_Resize()
MALLA.Width = PLANILLA.Width * 0.95
MALLA.Height = PLANILLA.Height * 1 / 2 - 100
MALLA.Left = 100
MALLA.Top = 0
RESU.Left = 100
RESU.Top = MALLA.Height + 100
RESU.Width = PLANILLA.Width * 3 / 4
RESU.Height = PLANILLA.Height / 2 - 100
PORCENTAJES.Left = 100
PORCENTAJES.Top = 300
PORCENTAJES.Width = RESU.Width - 200
PORCENTAJES.Height = RESU.Height - 350
PORCENTAJES.ColWidth(0) = 300
PORCENTAJES.ColWidth(1) = PORCENTAJES.Width / 2 + 200
PORCENTAJES.ColWidth(2) = PORCENTAJES.Width * 0.2
PORCENTAJES.ColWidth(3) = PORCENTAJES.Width * 0.2
Image1(0).Left = PLANILLA.Width - Image1(0).Width - 200
Image1(1).Left = PLANILLA.Width - Image1(1).Width - 200
Image1(0).Top = PLANILLA.Height - Image1(0).Height - 200
Image1(1).Top = PLANILLA.Height - Image1(1).Height - 200
resumen(0).Left = Image1(1).Left
resumen(1).Left = Image1(1).Left
PORCENTAJES.Clear
CARGAR_REPORTES Data2, REPOR
cargar_resultados PORCENTAJES, REPOR
INICIALIZAR
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image1(0).Visible = False
Image1(1).Visible = True
Timer6.Enabled = True
End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

Image1(1).Visible = False
Image1(0).Visible = True
End If
End Sub



Private Sub Label3_Click()

End Sub

Private Sub resul_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub MALLA_DblClick()
    SELCAMPOS.Show
End Sub

Private Sub PORCENTAJES_Click()
    'MsgBox "la fila es la " & PORCENTAJES.RowSel
    If PORCENTAJES.RowSel <> 1 Then
        CARGAR_CONSULTA REPOR(PORCENTAJES.RowSel - 1).SQL
    Else
        Consulta = "SELECT Tabla1.Nombre_suscriptor, " & _
               "Tabla1.Cedula, Tabla1.Codigo, Tabla1.Ruta," & _
               "Tabla1.Ubicacion_casa, Tabla1.no_pisos, " & _
               "Tabla1.Direccion_predio, Tabla1.Numero_catastral, " & _
               "Tabla1.Estado_predio, Tabla1.Numero_personas_casa, " & _
               "Tabla1.Numero_familias_casa, Tabla1.Numero_menores_5, " & _
               "Tabla1.Conectado_sistema, Tabla1.Otra_fuente, Tabla1.Cual," & _
               "Tabla1.Calidad_agua, Tabla1.Cantidad_agua_suficiente, " & _
               "Tabla1.Uso_predio, Tabla1.Diametro_conexion, Tabla1.Tipo_materiales, " & _
               "Tabla1.Estado_medidor, Tabla1.Numero_medidor, Tabla1.Marca_medidor, " & _
               "Tabla1.Lectura, Tabla1.Estado_cajilla, Tabla1.Tipo_conexion_usuario," & _
               "Tabla1.Tanque_almacenamiento, Tabla1.Almacena_agua, Tabla1.Hierve_agua, " & _
               "Tabla1.Reparacion_instalacion, Tabla1.Quien_realiza, Tabla1.Gotea_llaves_grifos," & _
               "Tabla1.Tipo_servicio_sanitario, Tabla1.Problemas_instalacion, Tabla1.Cuales," & _
               "Tabla1.Inodoro_limpio, Tabla1.Estado_caseta, Tabla1.Taponada_conexion, " & _
               "Tabla1.Solucion_problema, Tabla1.Cual_solucion, Tabla1.Destino_basuras," & _
               "Tabla1.Existencia_basuras_casa, Tabla1.Veces_barrido_semana,Tabla1.Veces_recoleccion_semana ," & _
               "Tabla1.Opinion_administracion, Tabla1.Respaldo_entidad, Tabla1.Observaciones, Tabla1.Desabitada" & _
               " From Tabla1   ORDER BY RUTA , DIRECCION_PREDIO"
    End If
    Timer1.Enabled = True
End Sub

Private Sub PORCENTAJES_EnterCell()
    If PORCENTAJES.RowSel <> 1 Then
        CARGAR_CONSULTA REPOR(PORCENTAJES.RowSel - 1).SQL
    Else
        Consulta = "SELECT Tabla1.Nombre_suscriptor, " & _
               "Tabla1.Cedula, Tabla1.Codigo, Tabla1.Ruta," & _
               "Tabla1.Ubicacion_casa, Tabla1.no_pisos, " & _
               "Tabla1.Direccion_predio, Tabla1.Numero_catastral, " & _
               "Tabla1.Estado_predio, Tabla1.Numero_personas_casa, " & _
               "Tabla1.Numero_familias_casa, Tabla1.Numero_menores_5, " & _
               "Tabla1.Conectado_sistema, Tabla1.Otra_fuente, Tabla1.Cual," & _
               "Tabla1.Calidad_agua, Tabla1.Cantidad_agua_suficiente, " & _
               "Tabla1.Uso_predio, Tabla1.Diametro_conexion, Tabla1.Tipo_materiales, " & _
               "Tabla1.Estado_medidor, Tabla1.Numero_medidor, Tabla1.Marca_medidor, " & _
               "Tabla1.Lectura, Tabla1.Estado_cajilla, Tabla1.Tipo_conexion_usuario," & _
               "Tabla1.Tanque_almacenamiento, Tabla1.Almacena_agua, Tabla1.Hierve_agua, " & _
               "Tabla1.Reparacion_instalacion, Tabla1.Quien_realiza, Tabla1.Gotea_llaves_grifos," & _
               "Tabla1.Tipo_servicio_sanitario, Tabla1.Problemas_instalacion, Tabla1.Cuales," & _
               "Tabla1.Inodoro_limpio, Tabla1.Estado_caseta, Tabla1.Taponada_conexion, " & _
               "Tabla1.Solucion_problema, Tabla1.Cual_solucion, Tabla1.Destino_basuras," & _
               "Tabla1.Existencia_basuras_casa, Tabla1.Veces_barrido_semana,Tabla1.Veces_recoleccion_semana ," & _
               "Tabla1.Opinion_administracion, Tabla1.Respaldo_entidad, Tabla1.Observaciones, Tabla1.Desabitada" & _
               " From Tabla1  ORDER BY RUTA, DIRECCION_PREDIO"
    End If
    Timer1.Enabled = True
End Sub

Private Sub resumen_Click(Index As Integer)
Select Case Index
    Case 0: Imprimir_Resumen PORCENTAJES, "PORCENTAJES ESTADISTICOS"
    Case 1: calculo_maximo_columnas MALLA
            IMPRIMIR_REPORTE MALLA, PORCENTAJES.TextMatrix(PORCENTAJES.RowSel, 1)
End Select
End Sub

Private Sub resumen_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For Y = 0 To 1
    If Y <> Index Then
        resumen(Y).BackColor = &H8000000F
    Else
        resumen(Y).BackColor = &HEED14A
    End If
    
Next Y
End Sub

Private Sub Timer1_Timer()
MALLA.Clear

Data1.RecordSource = Consulta
Data1.Refresh
If Data1.Recordset.EOF = False Then
Data1.Recordset.MoveLast
End If
tam = Data1.Recordset.RecordCount
MALLA.Rows = tam + 1
MALLA.Cols = COLUMN + 1

cargar_rejilla MALLA, M


MALLA.ColAlignment(9) = 1
MALLA.ColAlignment(18) = 1
If Data1.Recordset.EOF = False Then
Data1.Recordset.MoveFirst
i = 1
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Timer9.Enabled = True
Timer10.Enabled = True
Timer1.Enabled = False
Else
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer1.Enabled = False
End If
Timer1.Enabled = False
End Sub

Private Sub Timer10_Timer()
    cargar
End Sub

Private Sub Timer11_Timer()
    cargar
End Sub

Private Sub Timer12_Timer()

If Data3.Recordset.EOF = False Then
   Data4.Recordset.MoveFirst
   While Data4.Recordset.EOF = False
        If Data3.Recordset!ruta = Data4.Recordset!ruta Then
            ENCONTRADO = True
            
        Else
            Data4.Recordset.MoveNext
            ENCONTRADO = False
        End If
   Wend
   If ENCONTRADO = True Then
        Data3.Recordset.MoveNext
        REPOR(83).cantidad = REPOR(83).cantidad + 1
        PORCENTAJES.TextMatrix(84, 2) = REPOR(83).cantidad
        ENCONTRADO = False
    End If
   
Else
    Timer12.Enabled = False
    Timer13.Enabled = False
    Timer14.Enabled = False
End If


End Sub
Private Sub Timer13_Timer()

If Data3.Recordset.EOF = False Then
   Data4.Recordset.MoveFirst
   While Data4.Recordset.EOF = False
        If Data3.Recordset!ruta = Data4.Recordset!ruta Then
            ENCONTRADO = True
            Data4.Recordset.MoveLast
        Else
            Data4.Recordset.MoveNext
            ENCONTRADO = False
        End If
   Wend
   If ENCONTRADO = True Then
        Data3.Recordset.MoveNext
        REPOR(83).cantidad = REPOR(83).cantidad + 1
        PORCENTAJES.TextMatrix(84, 2) = REPOR(83).cantidad
        ENCONTRADO = False
    End If
   
Else
    Timer12.Enabled = False
    Timer13.Enabled = False
    Timer14.Enabled = False
End If

End Sub
Private Sub Timer14_Timer()

If Data3.Recordset.EOF = False Then
   Data4.Recordset.MoveFirst
   While Data4.Recordset.EOF = False
        If Data3.Recordset!ruta = Data4.Recordset!ruta Then
            ENCONTRADO = True
            
        Else
            Data4.Recordset.MoveNext
            ENCONTRADO = False
        End If
   Wend
   If ENCONTRADO = True Then
        Data3.Recordset.MoveNext
        REPOR(83).cantidad = REPOR(83).cantidad + 1
        PORCENTAJES.TextMatrix(84, 2) = REPOR(83).cantidad
        ENCONTRADO = False
    End If
   
Else
    Timer12.Enabled = False
    Timer13.Enabled = False
    Timer14.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
    cargar
End Sub
Private Sub Timer3_Timer()
    cargar
End Sub
Private Sub Timer4_Timer()
    cargar
End Sub

Private Sub Timer5_Timer()
    cargar
End Sub

Public Sub cargar()
    
    MALLA.TextMatrix(i, 0) = i
    If IsNull(Data1.Recordset!nombre_suscriptor) = False Then
    MALLA.TextMatrix(i, 1) = Data1.Recordset!nombre_suscriptor
    End If
    

    If Data1.Recordset!Cedula <> 0 Then
        MALLA.TextMatrix(i, 2) = Data1.Recordset!Cedula
    Else
        MALLA.TextMatrix(i, 2) = ""
    End If
    
    If Data1.Recordset!codigo <> 0 Then
        MALLA.TextMatrix(i, 3) = Data1.Recordset!codigo
    End If
    
    If Data1.Recordset!ruta <> 0 Then
        MALLA.TextMatrix(i, 4) = Data1.Recordset!ruta
    End If
    
    If Data1.Recordset!ubicacion_casa = True Then
        MALLA.TextMatrix(i, 5) = "Zona Urbana"
    Else
        MALLA.TextMatrix(i, 5) = "Zona Rural"
    End If
    
    If IsNull(Data1.Recordset!no_pisos) = False Then
    MALLA.TextMatrix(i, 6) = Data1.Recordset!no_pisos
    End If
    If IsNull(Data1.Recordset!direccion_predio) = False Then
        MALLA.TextMatrix(i, 7) = Data1.Recordset!direccion_predio
    End If
    
    If IsNull(Data1.Recordset!numero_catastral) = False Then
        MALLA.TextMatrix(i, 8) = Data1.Recordset!numero_catastral
    End If
    
    Select Case Data1.Recordset!estado_predio
        Case 0: MALLA.TextMatrix(i, 9) = "Sin deter"
        Case 1: MALLA.TextMatrix(i, 9) = "Lote"
        Case 2: MALLA.TextMatrix(i, 9) = "En construcción"
        Case 3: MALLA.TextMatrix(i, 9) = "Construido"
    End Select
    If IsNull(Data1.Recordset!numero_personas_casa) = False Then
    MALLA.TextMatrix(i, 10) = Data1.Recordset!numero_personas_casa
    End If
    If IsNull(Data1.Recordset!numero_familias_casa) = False Then
    MALLA.TextMatrix(i, 11) = Data1.Recordset!numero_familias_casa
    End If
    If IsNull(Data1.Recordset!numero_menores_5) = False Then
    MALLA.TextMatrix(i, 12) = Data1.Recordset!numero_menores_5
    End If
    
    If Data1.Recordset!conectado_sistema = True Then
        MALLA.TextMatrix(i, 13) = "Sí"
    Else
        MALLA.TextMatrix(i, 13) = "No"
    End If
    
    If Data1.Recordset!otra_fuente = True Then
        MALLA.TextMatrix(i, 14) = "Sí"
    Else
        MALLA.TextMatrix(i, 14) = "No"
    End If
    
    If IsNull(Data1.Recordset!Cual) = False Then
        MALLA.TextMatrix(i, 15) = Data1.Recordset!Cual
    End If
    
    If Data1.Recordset!calidad_agua = -1 Then
        MALLA.TextMatrix(i, 16) = "Buena"
    ElseIf Data1.Recordset!calidad_agua = 2 Then
        MALLA.TextMatrix(i, 16) = "Mala"
    ElseIf Data1.Recordset!calidad_agua = 0 Then
        MALLA.TextMatrix(i, 16) = ""
    End If
    
    If Data1.Recordset!cantidad_agua_suficiente = True Then
        MALLA.TextMatrix(i, 17) = "Sí"
    Else
        MALLA.TextMatrix(i, 17) = "No"
    End If
    
    
    Select Case Data1.Recordset!uso_predio
        Case 0: MALLA.TextMatrix(i, 18) = ""
        Case 1: MALLA.TextMatrix(i, 18) = "Residencial"
        Case 2: MALLA.TextMatrix(i, 18) = "Comercial"
        Case 3: MALLA.TextMatrix(i, 18) = "Industrial"
        Case 4: MALLA.TextMatrix(i, 18) = "Oficial"
        Case 5: MALLA.TextMatrix(i, 18) = "Mixto"
    End Select
    
    Select Case Data1.Recordset!diametro_conexion
        Case 0: MALLA.TextMatrix(i, 19) = ""
        Case 1: MALLA.TextMatrix(i, 19) = Chr(189) & "'"
        Case 2: MALLA.TextMatrix(i, 19) = Chr(190) & "'"
        Case 3: MALLA.TextMatrix(i, 19) = "1'"
        Case 4: MALLA.TextMatrix(i, 19) = ">1'"
    End Select
    
    Select Case Data1.Recordset!tipo_materiales
        Case 0: MALLA.TextMatrix(i, 20) = ""
        Case 1: MALLA.TextMatrix(i, 20) = "P.V.C."
        Case 2: MALLA.TextMatrix(i, 20) = "Galvanizado"
        Case 3: MALLA.TextMatrix(i, 20) = "Manguera"
        Case 4: MALLA.TextMatrix(i, 20) = "Otro"
    End Select
    
    Select Case Data1.Recordset!estado_medidor
        Case 0: MALLA.TextMatrix(i, 21) = ""
        Case 1: MALLA.TextMatrix(i, 21) = "Registrando"
        Case 2: MALLA.TextMatrix(i, 21) = "Detenido"
        Case 3: MALLA.TextMatrix(i, 21) = "Nublado"
        Case 4: MALLA.TextMatrix(i, 21) = "Dañado"
        Case 5: MALLA.TextMatrix(i, 21) = "Sin medidor"
    End Select
    
    If IsNull(Data1.Recordset!numero_medidor) = False Then
        MALLA.TextMatrix(i, 22) = Data1.Recordset!numero_medidor
    End If
    
    If IsNull(Data1.Recordset!marca_medidor) = False Then
        MALLA.TextMatrix(i, 23) = Data1.Recordset!marca_medidor
    End If
    If IsNull(Data1.Recordset!lectura) = False Then
    MALLA.TextMatrix(i, 24) = Data1.Recordset!lectura
    End If
    Select Case Data1.Recordset!estado_cajilla
        Case 0: MALLA.TextMatrix(i, 25) = ""
        Case 1: MALLA.TextMatrix(i, 25) = "Bueno"
        Case 2: MALLA.TextMatrix(i, 25) = "Malo"
        Case 3: MALLA.TextMatrix(i, 25) = "No existe"
    End Select
    
    Select Case Data1.Recordset!tipo_conexion_usuario
        Case 0: MALLA.TextMatrix(i, 26) = ""
        Case 1: MALLA.TextMatrix(i, 26) = "Legal"
        Case 2: MALLA.TextMatrix(i, 26) = "No Incluida S"
        Case 3: MALLA.TextMatrix(i, 26) = "Multiusuario"
        Case 4: MALLA.TextMatrix(i, 26) = "Clandestina"
                
        Case 5: MALLA.TextMatrix(i, 26) = "Provisional"
        Case 6: MALLA.TextMatrix(i, 26) = "No existe"
    End Select
    
    If Data1.Recordset!tanque_almacenamiento = True Then
        MALLA.TextMatrix(i, 27) = "Sí"
    Else
        MALLA.TextMatrix(i, 27) = "No"
    End If
    
    If Data1.Recordset!almacena_agua = True Then
        MALLA.TextMatrix(i, 28) = "Si"
    Else
        MALLA.TextMatrix(i, 28) = "No"
    End If
    
    Select Case Data1.Recordset!hierve_agua
        Case 0: MALLA.TextMatrix(i, 29) = ""
        Case 1: MALLA.TextMatrix(i, 29) = "Siempre"
        Case 2: MALLA.TextMatrix(i, 29) = "Alguna veces"
        Case 3: MALLA.TextMatrix(i, 29) = "Nunca"
        Case 4: MALLA.TextMatrix(i, 29) = "Sólo para niños"
    End Select
    
    If IsNull(Data1.Recordset!reparacion_instalacion) = False Then
        MALLA.TextMatrix(i, 30) = Data1.Recordset!reparacion_instalacion
    End If
    
    Select Case Data1.Recordset!quien_realiza
        Case 0: MALLA.TextMatrix(i, 31) = ""
        Case 1: MALLA.TextMatrix(i, 31) = "Fontanero"
        Case 2: MALLA.TextMatrix(i, 31) = "Familiar"
        Case 3: MALLA.TextMatrix(i, 31) = "Particular"
    End Select
    
    If Data1.Recordset!gotea_llaves_grifos = True Then
        MALLA.TextMatrix(i, 32) = "Sí"
    Else
        MALLA.TextMatrix(i, 32) = "No"
    End If
    
    Select Case Data1.Recordset!tipo_servicio_sanitario
        Case 0: MALLA.TextMatrix(i, 33) = ""
        Case 1: MALLA.TextMatrix(i, 33) = "Inodoro con conexión al alcantarillado"
        Case 2: MALLA.TextMatrix(i, 33) = "Inodoro o taza con tanque séptico"
        Case 3: MALLA.TextMatrix(i, 33) = "Inodoro sin conexión al alcantarillado"
        Case 4: MALLA.TextMatrix(i, 33) = "Ninguno"
    End Select
    
    If Data1.Recordset!problemas_instalacion = True Then
        MALLA.TextMatrix(i, 34) = "Sí"
    Else
        MALLA.TextMatrix(i, 34) = "No"
    End If
    
    If IsNull(Data1.Recordset!cuales) = False Then
        MALLA.TextMatrix(i, 35) = Data1.Recordset!cuales
    End If
    
    If Data1.Recordset!inodoro_limpio = True Then
        MALLA.TextMatrix(i, 36) = "Sí"
    Else
        MALLA.TextMatrix(i, 36) = "No"
    End If
    
    Select Case Data1.Recordset!estado_caseta
        Case 0: MALLA.TextMatrix(i, 37) = ""
        Case 1: MALLA.TextMatrix(i, 37) = "Bueno"
        Case 2: MALLA.TextMatrix(i, 37) = "Malo"
        Case 3: MALLA.TextMatrix(i, 37) = "No Existe"
    End Select
    
    If Data1.Recordset!taponada_conexion = True Then
        MALLA.TextMatrix(i, 38) = "Sí"
    Else
        MALLA.TextMatrix(i, 38) = "No"
    End If
    
    Select Case Data1.Recordset!solucion_problema
        Case 0: MALLA.TextMatrix(i, 39) = ""
        Case 1: MALLA.TextMatrix(i, 39) = "Usted Mismo lo solucionó"
        Case 2: MALLA.TextMatrix(i, 39) = "Llamó al operador"
        Case 3: MALLA.TextMatrix(i, 39) = "Otro"
    End Select
    
    If IsNull(Data1.Recordset!cual_solucion) = False Then
        MALLA.TextMatrix(i, 40) = Data1.Recordset!cual_solucion
    End If
    
    Select Case Data1.Recordset!destino_basuras
        Case 0: MALLA.TextMatrix(i, 41) = ""
        Case 1: MALLA.TextMatrix(i, 41) = "La Quema"
        Case 2: MALLA.TextMatrix(i, 41) = "La Arroja"
        Case 3: MALLA.TextMatrix(i, 41) = "Carro Recolector"
        Case 4: MALLA.TextMatrix(i, 41) = "La Entierra"
    End Select
    
    If Data1.Recordset!existencia_basuras_casa = True Then
        MALLA.TextMatrix(i, 42) = "Sí"
    Else
        MALLA.TextMatrix(i, 42) = "No"
    End If
    If IsNull(Data1.Recordset!veces_barrido_semana) = False Then
    MALLA.TextMatrix(i, 43) = Data1.Recordset!veces_barrido_semana
    End If
    If IsNull(Data1.Recordset!veces_recoleCcion_semana) = False Then
    MALLA.TextMatrix(i, 44) = Data1.Recordset!veces_recoleCcion_semana
    End If
    
    Select Case Data1.Recordset!opinion_administracion
        Case 0: MALLA.TextMatrix(i, 45) = "No opinó"
        Case 1: MALLA.TextMatrix(i, 45) = "Buena"
        Case 2: MALLA.TextMatrix(i, 45) = "Regular"
        Case 3: MALLA.TextMatrix(i, 45) = "Mala"
    End Select
    
    If Data1.Recordset!respaldo_entidad = True Then
        MALLA.TextMatrix(i, 46) = "Sí"
    Else
        MALLA.TextMatrix(i, 46) = "No"
    End If
    
    If IsNull(Data1.Recordset!Observaciones) = False Then
        MALLA.TextMatrix(i, 47) = ""
        For j = 1 To Len(Data1.Recordset!Observaciones)
            If Mid(Data1.Recordset!Observaciones, j, 1) <> Chr(13) Then
                MALLA.TextMatrix(i, 47) = MALLA.TextMatrix(i, 47) & Mid(Data1.Recordset!Observaciones, j, 1)
            Else
                MALLA.TextMatrix(i, 47) = MALLA.TextMatrix(i, 47) & " "
                j = j + 1
            End If
        Next j
    End If
    
    If Data1.Recordset!DESABITADA = True Then
        MALLA.TextMatrix(i, 48) = "Sí"
    Else
        MALLA.TextMatrix(i, 48) = "No"
    End If
    
    
    
    Data1.Recordset.MoveNext
    i = i + 1
    If i = tam + 1 Then
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Timer7.Enabled = False
        Timer8.Enabled = False
        Timer9.Enabled = False
        Timer10.Enabled = False
        MALLA.Visible = True
    End If
End Sub

Private Sub Timer6_Timer()
    PLANILLA.Hide
    Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
    cargar
End Sub

Private Sub Timer8_Timer()
    cargar
End Sub

Private Sub Timer9_Timer()
    cargar
End Sub

Private Sub INICIALIZAR()

Data3.DatabaseName = App.Path + "\encuesta.mdb"
Data4.DatabaseName = App.Path + "\encuesta.mdb"
Data3.RecordSource = "Select * from tabla1"
Data4.RecordSource = "Select * from tabla2"
Data3.Refresh
Data4.Refresh
ENCONTRADO = False
REPOR(83).cantidad = 0
'Timer12.Enabled = True
'Timer13.Enabled = True
'Timer14.Enabled = True
End Sub
