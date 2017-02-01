VERSION 5.00
Begin VB.Form HIJO3 
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer validar 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6240
      Top             =   8520
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   840
      Top             =   8520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   8520
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "ALCANTARILLADO"
      ForeColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   6000
      Width           =   11655
      Begin VB.TextBox cual3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   38
         Top             =   1320
         Width           =   10575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cual?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label solucion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Otro ( )"
         Height          =   195
         Index           =   2
         Left            =   6840
         TabIndex        =   36
         Top             =   960
         Width           =   960
      End
      Begin VB.Label solucion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Llamó al operador ( )"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   35
         Top             =   960
         Width           =   2520
      End
      Begin VB.Label solucion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Usted mismo lo solucionó ( )"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   34
         Top             =   960
         Width           =   3360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Qué hizo para solucionar el problema?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   4800
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Se ha taponado su conexión de alcantarillado?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   5760
      End
      Begin VB.Label taponado 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         Height          =   195
         Index           =   0
         Left            =   6360
         TabIndex        =   31
         Top             =   240
         Width           =   720
      End
      Begin VB.Label taponado 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   30
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "ESTADO DE LAS INSTALACIONES SANITARIAS"
      ForeColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   11655
      Begin VB.Label caseta 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "No existe ( )"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   28
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label caseta 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Malo ( )"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   27
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label caseta 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Bueno ( )"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Observe en qué estado se encuentra la caseta de la instalación sanitaria:"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   9120
      End
      Begin VB.Label limpio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         Height          =   195
         Index           =   1
         Left            =   7680
         TabIndex        =   24
         Top             =   360
         Width           =   720
      End
      Begin VB.Label limpio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         Height          =   195
         Index           =   0
         Left            =   6600
         TabIndex        =   23
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Observe si el inodoro / taza / letrina está limpio"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   6360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "SANEAMIENTO EN LA VIVIENDA"
      ForeColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   11535
      Begin VB.TextBox cual2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   1800
         Width           =   10335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuales?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label instal 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         Height          =   195
         Index           =   0
         Left            =   7200
         TabIndex        =   18
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label instal 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         Height          =   195
         Index           =   1
         Left            =   8280
         TabIndex        =   17
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. A tenido usted problemas con su instalación sanitaria?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   6840
      End
      Begin VB.Label servicio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Ninguno                                ( )"
         Height          =   195
         Index           =   3
         Left            =   6360
         TabIndex        =   15
         Top             =   1080
         Width           =   5040
      End
      Begin VB.Label servicio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Inodoro sin conexión al alcantarillado ( )"
         Height          =   195
         Index           =   2
         Left            =   6360
         TabIndex        =   14
         Top             =   720
         Width           =   5040
      End
      Begin VB.Label servicio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Inodoro o taza con tanque séptico           ( )"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   13
         Top             =   1080
         Width           =   5640
      End
      Begin VB.Label servicio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Inodoro con conexión al alcantarillado      ( )"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   5640
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Observe que tipo de servicio sanitario existe en la vivienda:"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   7680
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "INSTALACIONES INTRADOMICILIARIAS"
      ForeColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.TextBox Reparaciones 
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   10935
      End
      Begin VB.Label goteando 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         Height          =   195
         Index           =   1
         Left            =   8880
         TabIndex        =   9
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label goteando 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   8
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Observe si hay llaves, grifos, tuberias o inodoros goteando"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   7440
      End
      Begin VB.Label operaciones 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Particular ( )"
         Height          =   195
         Index           =   2
         Left            =   8040
         TabIndex        =   6
         Top             =   1200
         Width           =   1680
      End
      Begin VB.Label operaciones 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Familiar ( )"
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   5
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label operaciones 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Fontanero ( )"
         Height          =   195
         Index           =   0
         Left            =   4440
         TabIndex        =   4
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Quién realiza las operaciones?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Qué reparaciones se realizan en su vivienda en las instalaciones de agua?"
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9120
      End
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   9720
      Picture         =   "HIJO3.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Regresar..."
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   9720
      Picture         =   "HIJO3.frx":2E74
      Stretch         =   -1  'True
      ToolTipText     =   "Regresar..."
      Top             =   7800
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
      TabIndex        =   39
      Top             =   7800
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   500
      Index           =   1
      Left            =   10800
      Picture         =   "HIJO3.frx":5C74
      Stretch         =   -1  'True
      ToolTipText     =   "Continuar..."
      Top             =   7800
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Image Image1 
      Height          =   500
      Index           =   0
      Left            =   10800
      Picture         =   "HIJO3.frx":8A74
      Stretch         =   -1  'True
      ToolTipText     =   "Continuar..."
      Top             =   7800
      Width           =   1000
   End
End
Attribute VB_Name = "HIJO3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub caseta_Click(Index As Integer)
Select Case Index
    Case 0: caseta(0).Caption = "Bueno (X)"
            caseta(1).Caption = "Malo ( )"
            caseta(2).Caption = "No existe ( )"
            EstadoCaseta = 1
    Case 1: caseta(0).Caption = "Bueno ( )"
            caseta(1).Caption = "Malo (X)"
            caseta(2).Caption = "No existe ( )"
            EstadoCaseta = 2
    Case 2: caseta(0).Caption = "Bueno ( )"
            caseta(1).Caption = "Malo ( )"
            caseta(2).Caption = "No existe (X)"
            EstadoCaseta = 3
End Select
End Sub

Private Sub cual2_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
End Sub

Private Sub cual3_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
End Sub

Private Sub Form_Load()
Frame4.Caption = "MANTENIMIENTO DE LA SOLUCIONES PARA DISPOSICIÓN DE EXCRETAS - ALCANTARILLADO"
CONT = 0
mensaje = ""
Operacion = 0
Goteo = 0
ServicioSanitario = 0
ProblemaInstalacion = 0
EstadoCaseta = 0
Alcantarillado = 0
SolucionP = 0
If EstadoP = 1 Then
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
End If
PADRE.siguiente.Enabled = False
Formu = 3
End Sub

Private Sub Form_Resize()
If PADRE.WindowState <> 1 And EstadoP <> 1 Then
    Frame1.Left = 100
    Frame1.Width = HIJO3.Width - 400
    Frame2.Left = 100
    Frame2.Width = HIJO3.Width - 400
    Frame3.Left = 100
    Frame3.Width = HIJO3.Width - 400
    Frame4.Left = 100
    Frame4.Width = HIJO3.Width - 400
Else
    Frame4.Top = Frame1.Top
    Frame4.Left = 100
    Frame4.Width = HIJO3.Width - 400
End If
End Sub

Private Sub goteando_Click(Index As Integer)
Select Case Index
    Case 0: goteando(0).Caption = "SI (X)"
            goteando(1).Caption = "NO ( )"
            Goteo = 1
    Case 1: goteando(0).Caption = "SI ( )"
            goteando(1).Caption = "NO (X)"
            Goteo = 2
End Select
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

Private Sub instal_Click(Index As Integer)

Select Case Index
    Case 0: instal(0).Caption = "SI (X)"
            instal(1).Caption = "NO ( )"
            ProblemaInstalacion = 1
            cual2.Enabled = True
            cual2.Text = Cual
            cual2.SetFocus
    Case 1: instal(0).Caption = "SI ( )"
            instal(1).Caption = "NO (X)"
            ProblemaInstalacion = 2
            Cual = cual2.Text
            cual2.Text = ""
            cual2.Enabled = False
           
            
End Select
End Sub

Private Sub limpio_Click(Index As Integer)
Select Case Index
    Case 0: limpio(0).Caption = "SI (X)"
            limpio(1).Caption = "NO ( )"
            InodoroL = 1
    Case 1: limpio(0).Caption = "SI ( )"
            limpio(1).Caption = "NO (X)"
            InodoroL = 2
End Select
End Sub

Private Sub operaciones_Click(Index As Integer)
Select Case Index
    Case 0: operaciones(0).Caption = "Fontanero (X)"
            operaciones(1).Caption = "Familiar ( )"
            operaciones(2).Caption = "Particular ( )"
            Operacion = 1
    Case 1: operaciones(0).Caption = "Fontanero ( )"
            operaciones(1).Caption = "Familiar (X)"
            operaciones(2).Caption = "Particular ( )"
            Operacion = 2
    Case 2: operaciones(0).Caption = "Fontanero ( )"
            operaciones(1).Caption = "Familiar ( )"
            operaciones(2).Caption = "Particular (X)"
            Operacion = 3
End Select
End Sub

Private Sub Reparaciones_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)

End Sub

Private Sub servicio_Click(Index As Integer)
Select Case Index
    Case 0: servicio(0).Caption = "Inodoro con conexión al alcantarillado      (X)"
            servicio(1).Caption = "Inodoro o taza con tanque séptico           ( )"
            servicio(2).Caption = "Inodoro sin conexión al alcantarillado ( )"
            servicio(3).Caption = "Ninguno                                ( )"
            ServicioSanitario = 1
    Case 1: servicio(0).Caption = "Inodoro con conexión al alcantarillado      ( )"
            servicio(1).Caption = "Inodoro o taza con tanque séptico           (X)"
            servicio(2).Caption = "Inodoro sin conexión al alcantarillado ( )"
            servicio(3).Caption = "Ninguno                                ( )"
            ServicioSanitario = 2
    Case 2: servicio(0).Caption = "Inodoro con conexión al alcantarillado      ( )"
            servicio(1).Caption = "Inodoro o taza con tanque séptico           ( )"
            servicio(2).Caption = "Inodoro sin conexión al alcantarillado (X)"
            servicio(3).Caption = "Ninguno                                ( )"
            ServicioSanitario = 3
    Case 3: servicio(0).Caption = "Inodoro con conexión al alcantarillado      ( )"
            servicio(1).Caption = "Inodoro o taza con tanque séptico           ( )"
            servicio(2).Caption = "Inodoro sin conexión al alcantarillado ( )"
            servicio(3).Caption = "Ninguno                                (X)"
            ServicioSanitario = 4
End Select
End Sub

Private Sub tanque_Click(Index As Integer)

End Sub

Private Sub solucion_Click(Index As Integer)
Select Case Index
    Case 0: solucion(0).Caption = "Usted mismo lo solucionó (X)"
            solucion(1).Caption = "Llamó al operador ( )"
            solucion(2).Caption = "Otro ( )"
            If cual3.Text <> "" Then
                CualX = cual3.Text
                cual3.Text = ""
            End If
            SolucionP = 1
    Case 1: solucion(0).Caption = "Usted mismo lo solucionó ( )"
            solucion(1).Caption = "Llamó al operador (X)"
            solucion(2).Caption = "Otro ( )"
            If cual3.Text <> "" Then
                CualX = cual3.Text
                cual3.Text = ""
            End If
            SolucionP = 2
    Case 2: solucion(0).Caption = "Usted mismo lo solucionó ( )"
            solucion(1).Caption = "Llamó al operador ( )"
            solucion(2).Caption = "Otro (X)"
            cual3.Text = CualX
            cual3.Enabled = True
            cual3.SetFocus
            SolucionP = 3
        
End Select
End Sub

Private Sub taponado_Click(Index As Integer)
Select Case Index
    Case 0: taponado(0).Caption = "SI (X)"
            taponado(1).Caption = "NO ( )"
            Alcantarillado = 1
    Case 1: taponado(0).Caption = "SI ( )"
            taponado(1).Caption = "NO (X)"
            Alcantarillado = 2
            cual3.Text = ""
End Select
PADRE.siguiente.Enabled = True 'CAMBIO TEMPORAL
End Sub

Private Sub Timer1_Timer()
Select Case CONT
    Case 0: mensaje = "Para pasar a la siguiente caja de texto" & vbCrLf & " solo basta dar Enter"
    Case 1: mensaje = "Haga click en los recuadros verdes para" & vbCrLf & " seleccionar una opción."
    Case 2: mensaje = "Presione Ctrl + I para ir a la pantalla siguiente."
    Case 3: mensaje = "Presione Ctrl + T para ir a la pantalla anterior."
End Select

If CONT < 3 Then
    CONT = CONT + 1
Else
    CONT = 0
End If
End Sub

Private Sub Timer2_Timer()
'If Validar_HIJO3 = False Then
 '   Timer2.Enabled = False
  '  Exit Sub
'End If
Guardar_HIJO3
PADRE.siguiente.Enabled = False
validar.Enabled = False
Formu = 4
HIJO4.Show
Timer2.Enabled = False
Timer1.Enabled = False
End Sub

Private Sub Timer3_Timer()
HIJO3.Hide
HIJO2.WindowState = 2
HIJO2.Show
Formu = 2
validar.Enabled = False
HIJO2.validar.Enabled = True
Timer3.Enabled = False
Timer1.Enabled = False

End Sub

Private Sub validar_Timer()
If Validar_FORMU3 = True Then
    PADRE.siguiente = True
Else
    PADRE.siguiente.Enabled = False
End If
End Sub
