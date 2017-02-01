VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MENU 
   BackColor       =   &H00000000&
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "MENU PRINCIPAL"
      ForeColor       =   &H00C0C0C0&
      Height          =   4215
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   4380
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3201
         ButtonWidth     =   2408
         ButtonHeight    =   3043
         AllowCustomize  =   0   'False
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   10
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Reportes..."
               Key             =   "reportes"
               Object.ToolTipText     =   "Generar Reportes"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Encuestas..."
               Key             =   "encuestas"
               Object.ToolTipText     =   "Realizar Nueva Encuesta"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2400
         Top             =   2760
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   2040
         Top             =   2760
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1680
         Top             =   2760
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   3360
      End
      Begin ComctlLib.ProgressBar progreso 
         Height          =   150
         Left            =   1600
         TabIndex        =   3
         Top             =   4020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   265
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4080
         Top             =   3000
      End
      Begin ComctlLib.StatusBar estado 
         Height          =   255
         Left            =   20
         TabIndex        =   2
         Top             =   3960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   3
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   6
               AutoSize        =   2
               TextSave        =   "05/08/02"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Bevel           =   2
               Enabled         =   0   'False
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   2
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Salir..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1702
         TabIndex        =   4
         Top             =   3200
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   5
         Top             =   3240
         Width           =   975
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   4080
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   84
         ImageHeight     =   95
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENU.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MENU.frx":10E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   4650
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "MENU.frx":21CC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim POS As Integer

Private Sub Form_Load()
i = 0
MENU.WindowState = 2
estado.Panels(3).Text = "Hora..."
Label3.Caption = "ACTUALIZACION DE CATASTRO DE USUARIOS DE ACUEDUCTO" & vbCrLf & _
                         " ALCANTARILLADO Y ASEO"
End Sub

Private Sub Form_Resize()
Image1.Left = (MENU.Width - Image1.Width) / 2
Label3.Left = (MENU.Width - Label3.Width) / 2
Frame1.Left = (MENU.Width - Frame1.Width) / 2

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H0&
End Sub

Private Sub Label1_Click()
Timer3.Enabled = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF00&
End Sub

Private Sub Timer1_Timer()
estado.Panels(3).Text = Time
End Sub

Private Sub Timer2_Timer()
If i > 100 Then
        Timer2.Enabled = False
        i = 0
        MENU.MousePointer = 0
        If POS = 1 Then
            'MENU.Hide
            
            Toolbar1.Buttons(10).Value = tbrUnpressed
            Toolbar1.Buttons(1).Value = tbrUnpressed
            progreso.Value = 0
            PLANILLA.WindowState = 2
            Consulta = "SELECT * from tabla1"
            COLUMN = 48
            PLANILLA.Timer1.Enabled = True
            PLANILLA.Show

        End If
        If POS = 2 Then
             Toolbar1.Buttons(10).Value = tbrUnpressed
            Toolbar1.Buttons(1).Value = tbrUnpressed
             'MENU.Hide
            progreso.Value = 0
            HIJO1.WindowState = 2
            HIJO1.Show
        End If
        
        Exit Sub
End If
progreso.Value = i
i = i + 9



End Sub

Private Sub Timer3_Timer()
Label1.Visible = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Label1.Visible = True
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "reportes":
                        MENU.MousePointer = 11
                        Toolbar1.Buttons(1).Value = tbrPressed
                        Toolbar1.Buttons(10).Value = tbrUnpressed
                        POS = 1
                        Timer2.Enabled = True
                        
    Case "encuestas":
                        MENU.MousePointer = 11
                        Toolbar1.Buttons(10).Value = tbrPressed
                        Toolbar1.Buttons(1).Value = tbrUnpressed
                        POS = 2
                        Timer2.Enabled = True
                        
                  
End Select
End Sub
