VERSION 5.00
Begin VB.Form Rangos 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generación e imporesión personalizada"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   1830
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox fin 
      Height          =   345
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox ini 
      Height          =   345
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el número de ruta final:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el número de ruta incial:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3225
   End
End
Attribute VB_Name = "Rangos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aceptar_Click()
RegIni = Val(ini.Text)
RegFin = Val(fin.Text)

If RegIni > RegFin Then
    MsgBox "Error. No se puede generar reporte!" & vbCrLf & _
    "Valor incial mayor que valor final", vbInformation, "Reportes personalizados"
    Exit Sub
End If

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
               " From Tabla1   where ruta>=" & RegIni & " AND ruta <= " & RegFin & " ORDER BY RUTA, DIRECCION_PREDIO"
COLUMN = 48
PLANILLA.Timer1.Enabled = True
PLANILLA.Show

End Sub

Private Sub fin_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    aceptar.SetFocus
    aceptar_Click
End If
End Sub

Private Sub ini_Click()
ini.SelStart = 0
ini.SelLength = Len(ini.Text)
End Sub

Private Sub ini_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    fin.SelStart = 0
    fin.SelLength = Len(ini.Text)
    fin.SetFocus
End If
End Sub
