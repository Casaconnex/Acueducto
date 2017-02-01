VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de"
   ClientHeight    =   3540
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5760
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2443.371
   ScaleMode       =   0  'User
   ScaleWidth      =   5408.938
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4560
      Top             =   3120
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   1
      Left            =   1680
      Picture         =   "frmAbout.frx":074C
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   0
      Left            =   1680
      Picture         =   "frmAbout.frx":6300
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1521.93
      Y2              =   1521.93
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizaciòn de catastro de usuarios de acueducto, alcantarillado y aseo"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E18648&
      Height          =   1170
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   4125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título de la aplicación"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E18648&
      Height          =   300
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1532.283
      Y2              =   1532.283
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 1.0"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E18648&
      Height          =   300
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then

Image1(0).Visible = False
Image1(1).Visible = True
Timer2.Enabled = True
End If
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then

Image1(1).Visible = False
Image1(0).Visible = True
End If
End Sub

Private Sub Timer2_Timer()
Unload Me
End Sub
