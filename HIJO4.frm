VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form HIJO4 
   BackColor       =   &H00000000&
   ClientHeight    =   8880
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
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Deshabitada"
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   9960
      TabIndex        =   25
      Top             =   6840
      Width           =   1815
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "NO"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   960
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "SI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "OBSERVACIONES"
      ForeColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   8895
      Begin VB.TextBox Observaciones 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Timer validar 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   6960
   End
   Begin MSComctlLib.ProgressBar guardando 
      Height          =   135
      Left            =   6840
      TabIndex        =   22
      ToolTipText     =   "Guardando....."
      Top             =   7635
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7320
      Top             =   7635
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   840
      Top             =   8520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8160
      Top             =   7635
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9120
      Top             =   7635
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   8520
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "RELACIONES CON LA OFICINA DE SERVICIOS PUBLICOS"
      ForeColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   11655
      Begin VB.Label entidad 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   20
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label entidad 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         Height          =   195
         Index           =   1
         Left            =   9720
         TabIndex        =   19
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Se siente respaldado por la entidad administradora del servicio?"
         ForeColor       =   &H00868686&
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   11535
      End
      Begin VB.Label opina 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Mala ( )"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   17
         Top             =   720
         Width           =   960
      End
      Begin VB.Label opina 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Regular ( )"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   16
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label opina 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Buena ( )"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Qué opina de la entidad administradora respecto a los servicios de acueducto, alcantarillado     y aseo?"
         ForeColor       =   &H00868686&
         Height          =   435
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "MANEJO DE BASURAS"
      ForeColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.TextBox recoleccion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10200
         TabIndex        =   12
         Text            =   "0"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox barrido 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10200
         TabIndex        =   10
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Cuántas veces por semana la empresa presta el servicio de recolección de basuras:"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   10080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Cuántas veces por semana la empresa presta el servicio de barrido de las calles:"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   9960
      End
      Begin VB.Label interior 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         Height          =   195
         Index           =   1
         Left            =   9840
         TabIndex        =   8
         Top             =   600
         Width           =   720
      End
      Begin VB.Label interior 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         Height          =   195
         Index           =   0
         Left            =   8640
         TabIndex        =   7
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Observe si existen basuras en el patio o en el interior de la casa"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   8280
      End
      Begin VB.Label basura 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "La entierra ( )"
         Height          =   195
         Index           =   3
         Left            =   9720
         TabIndex        =   5
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label basura 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Carro recolector ( )"
         Height          =   195
         Index           =   2
         Left            =   7080
         TabIndex        =   4
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label basura 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "La arroja ( )"
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label basura 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "La quema ( )"
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Qué hace con las basuras?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Image nuevo 
      Height          =   615
      Index           =   1
      Left            =   9960
      Picture         =   "HIJO4.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo....."
      Top             =   7635
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image nuevo 
      Height          =   615
      Index           =   0
      Left            =   9960
      Picture         =   "HIJO4.frx":02A0
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo....."
      Top             =   7635
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   10920
      Picture         =   "HIJO4.frx":0570
      Stretch         =   -1  'True
      Top             =   7635
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   10920
      Picture         =   "HIJO4.frx":0810
      Stretch         =   -1  'True
      Top             =   7635
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   8760
      Picture         =   "HIJO4.frx":0AE0
      Stretch         =   -1  'True
      ToolTipText     =   "Regresar..."
      Top             =   7635
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   8760
      Picture         =   "HIJO4.frx":3954
      Stretch         =   -1  'True
      ToolTipText     =   "Regresar..."
      Top             =   7635
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label mensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EED14A&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   7635
      Width           =   150
   End
End
Attribute VB_Name = "HIJO4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub barrido_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    
    recoleccion.SelStart = 0
    recoleccion.SelLength = Len(recoleccion.Text)
    recoleccion.SetFocus
End If
End Sub

Private Sub basura_Click(Index As Integer)
Select Case Index
    Case 0: basura(0).Caption = "La quema (X)"
            basura(1).Caption = "La arroja ( )"
            basura(2).Caption = "Carro recolector ( )"
            basura(3).Caption = "La entierra ( )"
            Basuras = 1
    Case 1: basura(0).Caption = "La quema ( )"
            basura(1).Caption = "La arroja (X)"
            basura(2).Caption = "Carro recolector ( )"
            basura(3).Caption = "La entierra ( )"
            Basuras = 2
    Case 2: basura(0).Caption = "La quema ( )"
            basura(1).Caption = "La arroja ( )"
            basura(2).Caption = "Carro recolector (X)"
            basura(3).Caption = "La entierra ( )"
            Basuras = 3
    Case 3: basura(0).Caption = "La quema ( )"
            basura(1).Caption = "La arroja ( )"
            basura(2).Caption = "Carro recolector ( )"
            basura(3).Caption = "La entierra (X)"
            Basuras = 4
End Select
End Sub


Private Sub entidad_Click(Index As Integer)
Select Case Index
    Case 0: entidad(0).Caption = "SI (X)"
            entidad(1).Caption = "NO ( )"
            Respaldo = 1
    Case 1: entidad(0).Caption = "SI ( )"
            entidad(1).Caption = "NO (X)"
            Respaldo = 2
End Select
End Sub

Private Sub Form_Load()
CONT = 0
mensaje = ""
Basuras = 0
BasuraCasa = 0
EntidadA = 0
Respaldo = 0
BotellonA = 0
Timer4.Enabled = True
If EstadoP = 1 Then
    Frame1.Visible = False
End If
PADRE.siguiente.Enabled = False
PADRE.guardar.Enabled = True
Formu = 4
End Sub

Private Sub Form_Resize()
If PADRE.WindowState <> 1 And EstadoP <> 1 Then
    Frame1.Left = 100
    Frame1.Width = HIJO4.Width - 400
    Frame2.Left = 100
    Frame2.Width = HIJO4.Width - 400
    
ElseIf EstadoP = 1 Then
    Frame2.Top = Frame1.Top
    Frame2.Left = 100
    Frame2.Width = HIJO4.Width - 400
End If
Frame3.Left = 100
Frame3.Width = HIJO4.Width - 400
Observaciones.Width = Frame3.Width - 250
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image1(0).Visible = False
Image1(1).Visible = True
Timer2.Enabled = True
End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

Image1(1).Visible = False
Image1(0).Visible = True
End If
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

Image2(0).Visible = False
Image2(1).Visible = True
Timer3.Enabled = True
End If
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

Image2(1).Visible = False
Image2(0).Visible = True
End If
End Sub

Private Sub interior_Click(Index As Integer)
Select Case Index
    Case 0: interior(0).Caption = "SI (X)"
            interior(1).Caption = "NO ( )"
            BasuraCasa = 1
    Case 1: interior(0).Caption = "SI ( )"
            interior(1).Caption = "NO (X)"
            BasuraCasa = 2
End Select
End Sub

Private Sub nuevo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

nuevo(0).Visible = False
nuevo(1).Visible = True
Timer5.Enabled = True
End If
End Sub

Private Sub nuevo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

nuevo(1).Visible = False
nuevo(0).Visible = True
End If
End Sub

Private Sub Observaciones_Change()
Observaciones.Text = UCase(Observaciones.Text)
Observaciones.SelStart = Len(Observaciones.Text)
End Sub

Private Sub opina_Click(Index As Integer)
Select Case Index
    Case 0: opina(0).Caption = "Buena (X)"
            opina(1).Caption = "Regular ( )"
            opina(2).Caption = "Mala ( )"
            EntidadA = 1
    Case 1: opina(0).Caption = "Buena ( )"
            opina(1).Caption = "Regular (X)"
            opina(2).Caption = "Mala ( )"
            EntidadA = 2
    Case 2: opina(0).Caption = "Buena ( )"
            opina(1).Caption = "Regular ( )"
            opina(2).Caption = "Mala (X)"
            EntidadA = 3
End Select
End Sub

Private Sub recoleccion_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    Observaciones.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
Select Case CONT
    Case 0: mensaje = "Para pasar a la siguiente caja de texto" & vbCrLf & " solo basta dar Enter"
    Case 1: mensaje = "Haga click en los recuadros verdes para" & vbCrLf & " seleccionar una opción."
    Case 2: mensaje = "Presione Ctrl + G para guardar."
    Case 3: mensaje = "Presione Ctrl + T para ir a la pantalla anterior."
End Select

If CONT < 3 Then
    CONT = CONT + 1
Else
    CONT = 0
End If
End Sub

Private Sub Timer2_Timer()
i = 0
'If Validar_HIJO4 = False Then
 '   Timer2.Enabled = False
  '  Exit Sub
'End If
Timer6.Enabled = True
Guardar_HIJO4
If SALVADO = False And BUSQ = True Then
    SALVADO = True
End If

Guardar_BaseDatos guardar

guardando.Visible = False
PADRE.nuevo.Enabled = True
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
HIJO4.Hide
HIJO3.WindowState = 2
HIJO3.Show
Formu = 3
validar.Enabled = False
HIJO3.validar.Enabled = False
Timer3.Enabled = False
Timer1.Enabled = False

End Sub

Private Sub Timer4_Timer()
If EstadoP <> 1 Then
barrido.SelStart = 0
barrido.SelLength = 1
barrido.SetFocus
Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
 If SALVADO = False And BUSQ = True Then
 If MsgBox("LOS DATOS INGRESADOS HASTA EL MOMENTO NO SERAN ALMACENADOS POR EL SISTEMA" & vbCrLf & "¿DESEA GUARDAR AHORA ?", vbQuestion + vbYesNo, "ADVERTENCIA") = vbYes Then
    Guardar_BaseDatos guardar
 ElseIf SALVADO = False And BUSQ = True Then
    Guardar_BaseDatos BusquedaE
 End If
End If
Unload HIJO1
Unload HIJO2
Unload HIJO3
Load HIJO1
HIJO1.WindowState = 2
HIJO1.Show
Timer1.Enabled = False
Timer5.Enabled = False
Unload HIJO4
End Sub

Private Sub Timer6_Timer()
HIJO4.MousePointer = 11
guardando.Visible = True
If i < 5 Then
guardando.Value = i
i = i + 1
Else
guardando.Visible = False
HIJO4.MousePointer = 0
If MsgBox("Desea ingresar otra encuesta?", vbQuestion + vbYesNo, "NUEVA ENCUESTA") = vbYes Then
    Timer5.Enabled = True
End If
Timer6.Enabled = False

End If
End Sub

Private Sub validar_Timer()
If Validar_FORMU4 = True Then
    PADRE.guardar.Enabled = True
Else
    PADRE.guardar.Enabled = True 'AQUI CAMBIE A TRUE*******
End If
End Sub
