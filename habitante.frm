VERSION 5.00
Begin VB.Form habitante 
   BackColor       =   &H80000008&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Habitantes"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6780
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
   ScaleHeight     =   5865
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Caption         =   "Niños menores de cinco años"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   6375
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pertenecen a familias conectadas al sistema"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   4410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pertenecen a familias No conectadas al sistema"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   4740
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3240
         TabIndex        =   13
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Caption         =   "Número de habitantes"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3240
         TabIndex        =   4
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Habitantes No conectados al sistema"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   3690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Habitantes conectados al sistema"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   3360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de familias"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "habitante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label10_Click()

End Sub

Private Sub cerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = App.Path + "\encuesta.mdb"
    
    Data1.RecordSource = "SELECT SUM (Numero_familias_casa) AS CANTI FROM TABLA1 "
    Data1.Refresh
    Label5.Caption = Data1.Recordset!CANTI
    
    Data1.RecordSource = "SELECT SUM (Numero_personas_casa) AS CANTI FROM TABLA1 WHERE Conectado_sistema = TRUE"
    Data1.Refresh
    Label6.Caption = Data1.Recordset!CANTI
    
    Data1.RecordSource = "SELECT SUM (Numero_personas_casa) AS CANTI FROM TABLA1 WHERE Conectado_sistema = FALSE"
    Data1.Refresh
    Label7.Caption = Data1.Recordset!CANTI
    
    Label8 = Val(Label6) + Val(Label7)
    
    Data1.RecordSource = "SELECT SUM (Numero_menores_5) AS CANTI FROM TABLA1 WHERE Conectado_sistema = TRUE"
    Data1.Refresh
    Label12.Caption = Data1.Recordset!CANTI
    
    Data1.RecordSource = "SELECT SUM (Numero_menores_5) AS CANTI FROM TABLA1 WHERE Conectado_sistema = FALSE"
    Data1.Refresh
    Label11.Caption = Data1.Recordset!CANTI
        
    Label9 = Val(Label11) + Val(Label12)
End Sub

