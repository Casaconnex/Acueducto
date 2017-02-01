VERSION 5.00
Begin VB.Form HIJO1 
   BackColor       =   &H00000000&
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer validar 
      Interval        =   1
      Left            =   3360
      Top             =   8520
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5160
      Top             =   8640
   End
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   2280
      Top             =   8640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1680
      Top             =   8640
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "COBERTURA"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   240
      TabIndex        =   28
      Top             =   5400
      Width           =   11655
      Begin VB.TextBox cual1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   37
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label cantidad 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   10800
         TabIndex        =   43
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label cantidad 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   9960
         TabIndex        =   42
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. La cantidad de agua suministrada por el sistema le alcanza para todas sus necesidades?"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   10680
      End
      Begin VB.Label calidad_agua 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Mala ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7920
         TabIndex        =   40
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label calidad_agua 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Buena ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6120
         TabIndex        =   39
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Qué opina de la calidad de agua del sistema?"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   5640
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cual?"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   6480
         TabIndex        =   36
         Top             =   720
         Width           =   600
      End
      Begin VB.Label fuente 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5520
         TabIndex        =   35
         Top             =   720
         Width           =   720
      End
      Begin VB.Label fuente 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4320
         TabIndex        =   34
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Recolecta agua de otra fuente?"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   3960
      End
      Begin VB.Label Abastece 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "NO ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   8160
         TabIndex        =   32
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Abastece 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "SI ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6960
         TabIndex        =   31
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Esta conectado al sistema de abastecimiento de agua?"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   6600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11655
      Begin VB.TextBox pisos 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   44
         Text            =   "1"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox no_ninos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "0"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox no_familias 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11040
         MaxLength       =   1
         TabIndex        =   24
         Text            =   "0"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox no_personas 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "0"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox no_catastro 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox direccion 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox ruta 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   9
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox codigo 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   7
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox cc 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Nombre 
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pisos:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   7320
         TabIndex        =   45
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de niños menores de cinco (5) años que viven en la casa:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   7560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de familias que viven en la casa: "
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   6000
         TabIndex        =   25
         Top             =   2520
         Width           =   4920
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de personas que viven en la casa:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   4800
      End
      Begin VB.Label estado 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Construido (X)"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7320
         TabIndex        =   21
         Top             =   2040
         Width           =   1680
      End
      Begin VB.Label estado 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "En construccion ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4560
         TabIndex        =   20
         Top             =   2040
         Width           =   2280
      End
      Begin VB.Label estado 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Lote ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   19
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del predio:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2160
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Catastral:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   7800
         TabIndex        =   17
         Top             =   1560
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección del predio: "
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2640
      End
      Begin VB.Label ubica 
         BackColor       =   &H00808000&
         Caption         =   "Zona Rural (X)"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   13
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label ubica 
         BackColor       =   &H00808000&
         Caption         =   "Zona Urbana ( )"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   12
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación de la casa:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2520
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   5880
         TabIndex        =   10
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.C"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   8760
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre suscriptor:"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00868686&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2160
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   8640
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
      TabIndex        =   46
      Top             =   7635
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   10800
      Picture         =   "HIJO1.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Continuar..."
      Top             =   7635
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   10800
      Picture         =   "HIJO1.frx":2E00
      Stretch         =   -1  'True
      ToolTipText     =   "Continuar..."
      Top             =   7635
      Width           =   1005
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASPECTOS RELACIONADOS CON EL AGUA"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00868686&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   5160
      Width           =   4950
   End
   Begin VB.Label NO_FORM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "AvantGarde Bk BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   525
   End
   Begin VB.Label TITULO 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TITULO"
      BeginProperty Font 
         Name            =   "AvantGarde Bk BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E18648&
      Height          =   330
      Left            =   5175
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "HIJO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Abastece_Click(Index As Integer)
Select Case Index
    Case 0: Abastece(0).Caption = "SI (X)"
            Abastece(1).Caption = "NO ( )"
            Abaste = 1
    Case 1: Abastece(0).Caption = "SI ( )"
            Abastece(1).Caption = "NO (X)"
            Abaste = 2
End Select
End Sub

Private Sub calidad_agua_Click(Index As Integer)
Select Case Index
    Case 0: calidad_agua(0).Caption = "Buena (X)"
            calidad_agua(1).Caption = "Mala ( )"
            Calidad = -1
    Case 1: calidad_agua(0).Caption = "Buena ( )"
            calidad_agua(1).Caption = "Mala (X)"
            Calidad = 2
End Select

End Sub

Private Sub cantidad_Click(Index As Integer)
Select Case Index
    Case 0: cantidad(0).Caption = "SI (X)"
            cantidad(1).Caption = "NO ( )"
            CantidadA = 1
    Case 1: cantidad(0).Caption = "SI ( )"
            cantidad(1).Caption = "NO (X)"
            CantidadA = 2
End Select
End Sub

Private Sub cc_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    If Mid(cc.Text, 1, 1) = "0" Then
        MsgBox "La cédula no puede empezar por cero (0)!", vbCritical, "CEDULA"
        cc.SelStart = 0
        cc.SelLength = 1
        cc.SetFocus
        Exit Sub
    End If
    codigo.Enabled = True
    codigo.SetFocus
    VarPos = 2
End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    If Mid(codigo.Text, 1, 1) = "0" Then
        MsgBox "El código no puede empezar por cero (0)!", vbCritical, "CODIGO"
        codigo.SelStart = 0
        codigo.SelLength = 1
        codigo.SetFocus
        Exit Sub
    End If
    If codigo.Text <> "" Then
        Data2.DatabaseName = App.Path + "\encuesta.mdb"
        Data2.RecordSource = "select * from tabla1"
        Data2.Refresh
        Data2.Recordset.FindFirst "codigo = " & Val(codigo.Text)
        If Data2.Recordset.NoMatch Then
            ruta.Enabled = True
            ruta.SetFocus
            VarPos = 3
        Else
            MsgBox "El código " & Val(codigo.Text) & " ya Existe ", vbInformation, "CÓDIGO"
            codigo.Text = ""
            codigo.SetFocus
            Exit Sub
        End If
    End If
    ruta.Enabled = True
    ruta.SetFocus
    VarPos = 3
End If

End Sub

Private Sub cual1_Change()
If EstadoP = 0 Then
    MsgBox "Debe seleccionar primero una de las opciones de" & vbCrLf & _
    "estado del predio.", vbInformation, "ESTADO DEL PREDIO"
    Exit Sub
End If

If Abaste = 0 Then
    MsgBox "Debe seleccionar primero si el usuario está o no" & vbrclf & _
    "conectado al sistema de abastecimiento de agua.", vbInformation, "ABASTECIMIENTO"
End If
End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And direccion.Text <> "" Then
    no_catastro.Enabled = True
    no_catastro.SetFocus
    VarPos = 6
End If
End Sub

Private Sub estado_Click(Index As Integer)
If pisos.Text <> "" And pisos.Text <> "0" Then
    NP = Val(pisos.Text)
End If
Select Case Index
    Case 0: estado(0).Caption = "Lote (X)"
            estado(1).Caption = "En construccion ( )"
            estado(2).Caption = "Construido ( )"
            EstadoP = 1
            no_personas.Text = 0
            no_familias.Text = 0
            no_ninos.Text = 0
            NP = Val(pisos.Text)
            pisos.Text = 0
            Activar_caja
            
    Case 1: estado(0).Caption = "Lote ( )"
            estado(1).Caption = "En construccion (X)"
            estado(2).Caption = "Construido ( )"
            EstadoP = 2
            no_personas.Text = 0
            no_familias.Text = 0
            no_ninos.Text = 0
            pisos.Text = NP
            Activar_caja
            
    Case 2: estado(0).Caption = "Lote ( )"
            estado(1).Caption = "En construccion ( )"
            estado(2).Caption = "Construido (X)"
            EstadoP = 3
            no_personas.Text = 0
            no_familias.Text = 0
            no_ninos.Text = 0
            pisos.Text = NP
            Activar_caja
            
End Select
End Sub

Private Sub Form_Load()
TITULO = ""
TITULO = TITULO + "SERVICIOS PUBLICOS" & vbCrLf _
       & "ACTUALIZACION DE CATASTRO DE USUARIOS DE ACUEDUCTO, ALCANTARILLADO Y ASEO" & vbCrLf _
       & "MUNICIPIO DE UBATE"
Data1.DatabaseName = App.Path + "\ENCUESTA.MDB"
Timer1.Enabled = True
CONT = 0
mensaje = ""
'variables de control
Ubicac = 0
'ASIGNACION TEMPORAL
Ubicac = 2
'------
EstadoP = 3
VarPos = 0
Abaste = 0
Recolecta = 0
Calidad = 0
CantidadA = 0
Formu = 1
PADRE.siguiente.Enabled = False

End Sub

Private Sub Form_Resize()
TITULO.Left = (HIJO1.Width - TITULO.Width) / 2
If PADRE.WindowState <> 1 Then
    Frame1.Left = 100
    Frame1.Width = HIJO1.Width - 400
    Frame2.Left = 100
    Frame2.Width = HIJO1.Width - 400
End If
End Sub

Private Sub fuente_Click(Index As Integer)
Select Case Index
    Case 0: fuente(0).Caption = "SI (X)"
            fuente(1).Caption = "NO ( )"
            cual1.Enabled = True
            cual1.SetFocus
            Recolecta = 1
    Case 1: fuente(0).Caption = "SI ( )"
            fuente(1).Caption = "NO (X)"
            Recolecta = 2
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

Private Sub no_catastro_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
        If no_catastro.Text <> "" Then
            Data2.DatabaseName = App.Path + "\encuesta.mdb"
            Data2.RecordSource = "SELECT * FROM TABLA1"
            Data2.Refresh
            Data2.Recordset.FindFirst " NUMERO_CATASTRAL =" & comillas & no_catastro.Text & comillas
            If Data2.Recordset.NoMatch Then
                If EstadoP > 1 Then
                    no_personas.Enabled = True
                    no_personas.SelStart = 0
                    no_personas.SelLength = Len(no_personas.Text)
                    no_personas.SetFocus
                End If
            Else
                MsgBox "El Número Catastral " & no_catastro.Text & " Ya existe dentro del sistema", vbInformation, "NÚMERO DE CATASTRO"
                no_catastro.Text = ""
                no_catastro.SetFocus
                Exit Sub
            
            End If
        End If
    If EstadoP > 1 Then
        no_personas.Enabled = True
        no_personas.SelStart = 0
        no_personas.SelLength = Len(no_personas.Text)
        no_personas.SetFocus
    End If
End If
End Sub

Private Sub no_familias_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 And no_familias.Text <> "" Then
    no_ninos.Enabled = True
    no_ninos.SelStart = 0
    no_ninos.SelLength = Len(no_ninos.Text)
    no_ninos.SetFocus
    VarPos = 8
End If
End Sub

Private Sub no_ninos_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 And no_ninos.Text <> "" Then
    If Val(no_ninos.Text) > Val(no_personas.Text) Then
        MsgBox "La cantidad de niños no puede exceder a la cantidad " & vbCrLf & _
        "de personas que viven en la casa.", vbCritical, "NIÑOS MENORES DE 5 AÑOS"
        no_ninos.SelStart = 0
        no_ninos.SelLength = Len(no_ninos.Text)
        no_ninos.SetFocus
        Exit Sub
    End If
    VarPos = 9
End If
End Sub

Private Sub no_personas_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 And no_personas.Text <> "" Then
    no_familias.Enabled = True
    no_familias.SelStart = 0
    no_familias.SelLength = Len(no_familias.Text)
    no_familias.SetFocus
    VarPos = 7
End If
End Sub

Private Sub Nombre_Change()
Nombre.Text = UCase(Nombre.Text)
Nombre.SelStart = Len(Nombre.Text)
End Sub

Private Sub Nombre_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 And Nombre.Text <> "" Then

    cc.Enabled = True
    cc.SetFocus
    VarPos = 1
End If
End Sub

Private Sub pisos_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    If Ubicac = 0 Then
        MsgBox "Para poder continuar debe seleccionar" & vbCrLf & _
        "primero la ubicación de la casa. ", vbInformation, "UBICACION"
        Timer4.Enabled = True
    Else
        direccion.Enabled = True
        direccion.SetFocus
        VarPos = 5
    End If

End If

End Sub

Private Sub ruta_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 And ruta.Text <> "" Then
        If ruta.Text <> "" Then
        Data2.DatabaseName = App.Path + "\encuesta.mdb"
        Data2.RecordSource = "select * from tabla1"
        Data2.Refresh
        Data2.Recordset.FindFirst "RUTA = " & Val(ruta.Text)
        If Data2.Recordset.NoMatch Then
            pisos.Enabled = True
            pisos.SetFocus
            VarPos = 3
        Else
            MsgBox "La ruta " & Val(ruta.Text) & " ya Existe ", vbInformation, "RUTA"
            ruta.Text = ""
            ruta.SetFocus
            Exit Sub
        End If
    End If

    pisos.Enabled = True
    pisos.SelStart = 0
    pisos.SelLength = Len(pisos.Text)
    pisos.SetFocus
    VarPos = 4
End If

End Sub

Private Sub Timer1_Timer()
'extrae numero de formulario
Data1.RecordSource = "SELECT COUNT(CODIGO)AS NUMERO FROM TABLA1"
Data1.Refresh
If Data1.Recordset!numero = 0 Then
    NO_FORM = "Formulario 1"
Else
    NO_FORM = "Formulario " & Data1.Recordset!numero + 1
End If
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
'If Validar_HIJO1 = False Then
 '   Timer2.Enabled = False
  '  Exit Sub
'End If
Guardar_HIJO1
Formu = 2
PADRE.anterior.Enabled = True
PADRE.siguiente.Enabled = False
validar.Enabled = False
HIJO2.Show
Timer2.Enabled = False
Timer3.Enabled = False

End Sub

Private Sub Timer3_Timer()
Select Case CONT
    Case 0: mensaje = "Para pasar a la siguiente caja de texto" & vbCrLf & " solo basta dar Enter"
    Case 1: mensaje = "Haga click en los recuadros verdes para" & vbCrLf & " seleccionar una opción."
    Case 2: mensaje = "Al ingresar el número de cédula hagalo sin incluir punto."
    Case 3: mensaje = "Presione Ctrl + I para ir a la pantalla siguiente."
    Case 4: mensaje = "Presione Ctrl + T para ir a la pantalla anterior."
    
    
End Select

If CONT < 4 Then
    CONT = CONT + 1
Else
    CONT = 0
End If
End Sub

Private Sub Timer4_Timer()
'SendKeys (Chr(13))
Timer4.Enabled = False
End Sub

Private Sub ubica_Click(Index As Integer)
Select Case Index
    Case 0: ubica(0).Caption = "Zona Urbana (X)"
            ubica(1).Caption = "Zona Rural ( )"
            Ubicac = 1
            Activar_caja
    Case 1: ubica(0).Caption = "Zona Urbana ( )"
            ubica(1).Caption = "Zona Rural (X)"
            Ubicac = 2
            Activar_caja
End Select
End Sub

Private Sub validar_Timer()
If Validar_FORMU1 = True Then
    PADRE.siguiente = True
Else
    PADRE.siguiente.Enabled = False
End If
End Sub
