VERSION 5.00
Begin VB.Form HIJO2 
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
   Begin VB.Timer validar 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1440
      Top             =   8400
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   840
      Top             =   8400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   8400
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "RIESGO"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   40
      Top             =   5400
      Width           =   11775
      Begin VB.Label hierve 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Solo para los niños ( )"
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
         Index           =   3
         Left            =   8520
         TabIndex        =   51
         Top             =   1080
         Width           =   2760
      End
      Begin VB.Label hierve 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Nunca ( )"
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
         TabIndex        =   50
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label hierve 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Algunas veces ( )"
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
         Left            =   5160
         TabIndex        =   49
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Label hierve 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Siempre ( )"
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
         Left            =   3720
         TabIndex        =   48
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Hierve el agua para beber?"
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
         TabIndex        =   47
         Top             =   1080
         Width           =   3480
      End
      Begin VB.Label consumo 
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
         Left            =   5400
         TabIndex        =   46
         Top             =   720
         Width           =   720
      End
      Begin VB.Label consumo 
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
         TabIndex        =   45
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Almacena agua para el consumo?"
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
         TabIndex        =   44
         Top             =   720
         Width           =   3960
      End
      Begin VB.Label tanque 
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
         Left            =   5400
         TabIndex        =   43
         Top             =   360
         Width           =   720
      End
      Begin VB.Label tanque 
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
         TabIndex        =   42
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Tiene tanque de almacenamiento?"
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
         Top             =   360
         Width           =   4080
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "INFORMACIÓN DE LA INSTALACIÓN"
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
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.TextBox lectura 
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
         Left            =   10200
         MaxLength       =   5
         TabIndex        =   28
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox marca_medidor 
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
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   26
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox no_medidor 
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label tipo_conex 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "No existe                   ( )"
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
         Index           =   5
         Left            =   4920
         TabIndex        =   39
         Top             =   4800
         Width           =   3720
      End
      Begin VB.Label tipo_conex 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Provisional                 ( )"
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
         Index           =   4
         Left            =   4920
         TabIndex        =   38
         Top             =   4440
         Width           =   3720
      End
      Begin VB.Label tipo_conex 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Clandestina                 ( )"
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
         Index           =   3
         Left            =   4920
         TabIndex        =   37
         Top             =   4080
         Width           =   3720
      End
      Begin VB.Label tipo_conex 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Multiusuario                 ( )"
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
         Left            =   480
         TabIndex        =   36
         Top             =   4800
         Width           =   3840
      End
      Begin VB.Label tipo_conex 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "No incluido en el sistema    ( )"
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
         Left            =   480
         TabIndex        =   35
         Top             =   4440
         Width           =   3840
      End
      Begin VB.Label tipo_conex 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Legal                        ( )"
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
         Left            =   480
         TabIndex        =   34
         Top             =   4080
         Width           =   3840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7. Determine que tipo de conexiòn tiene el usuario:"
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
         Top             =   3720
         Width           =   6120
      End
      Begin VB.Label estado_cajilla 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "No existe ( )"
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
         Left            =   5880
         TabIndex        =   32
         Top             =   3360
         Width           =   1560
      End
      Begin VB.Label estado_cajilla 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Malo ( )"
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
         TabIndex        =   31
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label estado_cajilla 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Bueno ( )"
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
         Left            =   3120
         TabIndex        =   30
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6. Estado de la cajilla:"
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
         TabIndex        =   29
         Top             =   3360
         Width           =   2880
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lectura:"
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
         Left            =   9120
         TabIndex        =   27
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca del medidor:"
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
         Left            =   4800
         TabIndex        =   25
         Top             =   2880
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. No. del medidor:"
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
         Top             =   2880
         Width           =   2280
      End
      Begin VB.Label estado_medidor 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Sin medidor ( )"
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
         Index           =   4
         Left            =   7800
         TabIndex        =   22
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label estado_medidor 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Dañado ( )"
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
         Index           =   3
         Left            =   6240
         TabIndex        =   21
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label estado_medidor 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Nublado ( )"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   2520
         Width           =   1320
      End
      Begin VB.Label estado_medidor 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Detenido ( )"
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
         Left            =   2640
         TabIndex        =   19
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label estado_medidor 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Registrando ( )"
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
         Left            =   480
         TabIndex        =   18
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Observe en que estado se encuentra el medidor:"
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
         TabIndex        =   17
         Top             =   2160
         Width           =   5880
      End
      Begin VB.Label tipo_conexion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Otro ( )"
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
         Index           =   3
         Left            =   6480
         TabIndex        =   16
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label tipo_conexion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Manguera ( )"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label tipo_conexion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Galvanizado ( )"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Label tipo_conexion 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "PVC ( )"
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
         Left            =   480
         TabIndex        =   13
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Observe en qué tipo de materiales se encuentra la conexión de acueducto:"
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
         TabIndex        =   12
         Top             =   1440
         Width           =   9000
      End
      Begin VB.Label diametro 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   ">1'( )"
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
         Index           =   3
         Left            =   10800
         TabIndex        =   11
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label diametro 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "1'( )"
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
         Left            =   9840
         TabIndex        =   10
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label diametro 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "3/4'( )"
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
         Left            =   8520
         TabIndex        =   9
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label diametro 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "1/2'(X)"
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
         Left            =   7080
         TabIndex        =   8
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label uso_predio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Mixto ( )"
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
         Index           =   4
         Left            =   9480
         TabIndex        =   7
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label uso_predio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Oficial ( )"
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
         Index           =   3
         Left            =   7320
         TabIndex        =   6
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label uso_predio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Industrial ( )"
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
         Left            =   5160
         TabIndex        =   5
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label uso_predio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Comercial ( )"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Observe que uso le están dando en la actualidad al predio"
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
         Top             =   360
         Width           =   7200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Observe en qué diámetro está la conexión del servicio:"
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
         TabIndex        =   2
         Top             =   1080
         Width           =   6840
      End
      Begin VB.Label uso_predio 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Residencial ( )"
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
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1800
      End
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   9720
      Picture         =   "HIJO2.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Regresar..."
      Top             =   7635
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   9720
      Picture         =   "HIJO2.frx":2E00
      Stretch         =   -1  'True
      ToolTipText     =   "Regresar..."
      Top             =   7635
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
      TabIndex        =   52
      Top             =   7635
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   500
      Index           =   0
      Left            =   10800
      Picture         =   "HIJO2.frx":5C74
      Stretch         =   -1  'True
      ToolTipText     =   "Continuar..."
      Top             =   7635
      Width           =   1000
   End
   Begin VB.Image Image1 
      Height          =   500
      Index           =   1
      Left            =   10800
      Picture         =   "HIJO2.frx":8AFE
      Stretch         =   -1  'True
      ToolTipText     =   "Continuar..."
      Top             =   7635
      Visible         =   0   'False
      Width           =   1000
   End
End
Attribute VB_Name = "HIJO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub consumo_Click(Index As Integer)
Select Case Index
    Case 0: consumo(0).Caption = "SI (X)"
            consumo(1).Caption = "NO ( )"
            AlmacenaA = 1
    Case 1: consumo(0).Caption = "SI ( )"
            consumo(1).Caption = "NO (X)"
            AlmacenaA = 2
End Select
End Sub

Private Sub diametro_Click(Index As Integer)
Select Case Index
    Case 0: diametro(0).Caption = "1/2'(X)"
            diametro(1).Caption = "3/4'( )"
            diametro(2).Caption = "1'( )"
            diametro(3).Caption = ">1'( )"
            DiametroC = 1
    Case 1: diametro(0).Caption = "1/2'( )"
            diametro(1).Caption = "3/4'(X)"
            diametro(2).Caption = "1'( )"
            diametro(3).Caption = ">1'( )"
            DiametroC = 2
    Case 2: diametro(0).Caption = "1/2'( )"
            diametro(1).Caption = "3/4'( )"
            diametro(2).Caption = "1'(X)"
            diametro(3).Caption = ">1'( )"
            DiametroC = 3
    Case 3: diametro(0).Caption = "1/2'( )"
            diametro(1).Caption = "3/4'( )"
            diametro(2).Caption = "1'( )"
            diametro(3).Caption = ">1'(X)"
            DiametroC = 4
End Select
End Sub

Private Sub estado_cajilla_Click(Index As Integer)
Select Case Index
    Case 0: estado_cajilla(0).Caption = "Bueno (X)"
            estado_cajilla(1).Caption = "Malo ( )"
            estado_cajilla(2).Caption = "No existe ( )"
            EstadoC = 1
    Case 1: estado_cajilla(0).Caption = "Bueno ( )"
            estado_cajilla(1).Caption = "Malo (X)"
            estado_cajilla(2).Caption = "No existe ( )"
            EstadoC = 2
    Case 2: estado_cajilla(0).Caption = "Bueno ( )"
            estado_cajilla(1).Caption = "Malo ( )"
            estado_cajilla(2).Caption = "No existe (X)"
            EstadoC = 3
End Select
End Sub

Private Sub estado_medidor_Click(Index As Integer)
Select Case Index
    Case 0: estado_medidor(0).Caption = "Registrando (X)"
            estado_medidor(1).Caption = "Detenido ( )"
            estado_medidor(2).Caption = "Nublado ( )"
            estado_medidor(3).Caption = "Dañado ( )"
            estado_medidor(4).Caption = "Sin medidor ( )"
            EstadoM = 1
            no_medidor.Enabled = True
    Case 1: estado_medidor(0).Caption = "Registrando ( )"
            estado_medidor(1).Caption = "Detenido (X)"
            estado_medidor(2).Caption = "Nublado ( )"
            estado_medidor(3).Caption = "Dañado ( )"
            estado_medidor(4).Caption = "Sin medidor ( )"
            EstadoM = 2
            no_medidor.Enabled = True
    Case 2: estado_medidor(0).Caption = "Registrando ( )"
            estado_medidor(1).Caption = "Detenido ( )"
            estado_medidor(2).Caption = "Nublado (X)"
            estado_medidor(3).Caption = "Dañado ( )"
            estado_medidor(4).Caption = "Sin medidor ( )"
            EstadoM = 3
            no_medidor.Enabled = True
    Case 3: estado_medidor(0).Caption = "Registrando ( )"
            estado_medidor(1).Caption = "Detenido ( )"
            estado_medidor(2).Caption = "Nublado ( )"
            estado_medidor(3).Caption = "Dañado (X)"
            estado_medidor(4).Caption = "Sin medidor ( )"
            EstadoM = 4
            no_medidor.Enabled = True
            
    Case 4: estado_medidor(0).Caption = "Registrando ( )"
            estado_medidor(1).Caption = "Detenido ( )"
            estado_medidor(2).Caption = "Nublado ( )"
            estado_medidor(3).Caption = "Dañado ( )"
            estado_medidor(4).Caption = "Sin medidor (X)"
            no_medidor.Enabled = False
            marca_medidor.Enabled = False
            lectura.Enabled = False
            EstadoM = 5
End Select
End Sub

Private Sub Form_Load()
CONT = 0
mensaje = ""
UsoP = 0
DiametroC = 1
MaterialC = 0
EstadoM = 0
EstadoC = 0
TipoC = 0
TanqueA = 0
AlmacenaA = 0
HierveA = 0
PADRE.siguiente.Enabled = False
PADRE.anterior.Enabled = True
Formu = 2
End Sub

Private Sub Form_Resize()
If PADRE.WindowState <> 1 Then
    Frame1.Left = 100
    Frame1.Width = HIJO2.Width - 400
    Frame2.Left = 100
    Frame2.Width = HIJO2.Width - 400
End If
End Sub

Private Sub hierve_Click(Index As Integer)
Select Case Index
    Case 0: hierve(0).Caption = "Siempre (X)"
            hierve(1).Caption = "Algunas veces ( )"
            hierve(2).Caption = "Nunca ( )"
            hierve(3).Caption = "Solo para los niños ( )"
            HierveA = 1
    Case 1: hierve(0).Caption = "Siempre ( )"
            hierve(1).Caption = "Algunas veces (X)"
            hierve(2).Caption = "Nunca ( )"
            hierve(3).Caption = "Solo para los niños ( )"
            HierveA = 2
    Case 2: hierve(0).Caption = "Siempre ( )"
            hierve(1).Caption = "Algunas veces ( )"
            hierve(2).Caption = "Nunca (X)"
            hierve(3).Caption = "Solo para los niños ( )"
            HierveA = 3
    Case 3: hierve(0).Caption = "Siempre ( )"
            hierve(1).Caption = "Algunas veces ( )"
            hierve(2).Caption = "Nunca ( )"
            hierve(3).Caption = "Solo para los niños (X)"
            HierveA = 4
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

Private Sub lectura_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_numero(KeyAscii)
End Sub


Private Sub marca_medidor_KeyPress(KeyAscii As Integer)
KeyAscii = Validar_letra(KeyAscii)
If KeyAscii = 13 Then
    lectura.Enabled = True
    lectura.SetFocus
End If
End Sub


Private Sub no_medidor_Change()
If UsoP = 0 Then
    MsgBox "Marque el uso actual del predio.", vbInformation, "USO DEL PREDIO"
    Exit Sub
ElseIf DiametroC = 0 Then
    MsgBox "Marque el diametro de la conexión.", vbInformation, "DIAMETRO DE LA CONEXION"
    Exit Sub
ElseIf MaterialC = 0 Then
    MsgBox "Marque el tipo de material de la conexión.", vbInformation, "MATERIAL DE LA CONEXION"
    Exit Sub
ElseIf EstadoM = 0 Then
    MsgBox "Marque el estado en que se encuentra el medidor.", vbInformation, "ESTADO DEL MEDIDOR"
    Exit Sub
End If
End Sub

Private Sub no_medidor_KeyPress(KeyAscii As Integer)
'KeyAscii = Validar_numero(KeyAscii)
If KeyAscii = 13 Then
    marca_medidor.Enabled = True
    marca_medidor.SetFocus
End If
End Sub

Private Sub tanque_Click(Index As Integer)
Select Case Index
    Case 0: tanque(0).Caption = "SI (X)"
            tanque(1).Caption = "NO ( )"
            TanqueA = 1
    Case 1: tanque(0).Caption = "SI ( )"
            tanque(1).Caption = "NO (X)"
            TanqueA = 2
End Select
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
'If Validar_HIJO2 = False Then
 '   Timer2.Enabled = False
  '  Exit Sub
'End If
Guardar_HIJO2
PADRE.siguiente.Enabled = False
validar.Enabled = False
Formu = 3
HIJO3.Show
Timer2.Enabled = False
Timer1.Enabled = False

End Sub

Private Sub Timer3_Timer()
HIJO2.Hide
HIJO1.WindowState = 2
HIJO1.validar.Enabled = True
validar.Enabled = False
Formu = 1
PADRE.anterior.Enabled = False
HIJO1.Show
Timer3.Enabled = False
Timer1.Enabled = False

End Sub

Private Sub tipo_conex_Click(Index As Integer)
Select Case Index
    Case 0: tipo_conex(0).Caption = "Legal                        (X)"
            tipo_conex(1).Caption = "No incluido en el sistema    ( )"
            tipo_conex(2).Caption = "Multiusuario                 ( )"
            tipo_conex(3).Caption = "Clandestina                 ( )"
            tipo_conex(4).Caption = "Provisional                 ( )"
            tipo_conex(5).Caption = "No existe                   ( )"
            TipoC = 1
    Case 1: tipo_conex(0).Caption = "Legal                        ( )"
            tipo_conex(1).Caption = "No incluido en el sistema    (X)"
            tipo_conex(2).Caption = "Multiusuario                 ( )"
            tipo_conex(3).Caption = "Clandestina                 ( )"
            tipo_conex(4).Caption = "Provisional                 ( )"
            tipo_conex(5).Caption = "No existe                   ( )"
            TipoC = 2
    Case 2: tipo_conex(0).Caption = "Legal                        ( )"
            tipo_conex(1).Caption = "No incluido en el sistema    ( )"
            tipo_conex(2).Caption = "Multiusuario                 (X)"
            tipo_conex(3).Caption = "Clandestina                 ( )"
            tipo_conex(4).Caption = "Provisional                 ( )"
            tipo_conex(5).Caption = "No existe                   ( )"
            TipoC = 3
    Case 3: tipo_conex(0).Caption = "Legal                        ( )"
            tipo_conex(1).Caption = "No incluido en el sistema    ( )"
            tipo_conex(2).Caption = "Multiusuario                 ( )"
            tipo_conex(3).Caption = "Clandestina                 (X)"
            tipo_conex(4).Caption = "Provisional                 ( )"
            tipo_conex(5).Caption = "No existe                   ( )"
            TipoC = 4
    Case 4: tipo_conex(0).Caption = "Legal                        ( )"
            tipo_conex(1).Caption = "No incluido en el sistema    ( )"
            tipo_conex(2).Caption = "Multiusuario                 ( )"
            tipo_conex(3).Caption = "Clandestina                 ( )"
            tipo_conex(4).Caption = "Provisional                 (X)"
            tipo_conex(5).Caption = "No existe                   ( )"
            TipoC = 5
    Case 5: tipo_conex(0).Caption = "Legal                        ( )"
            tipo_conex(1).Caption = "No incluido en el sistema    ( )"
            tipo_conex(2).Caption = "Multiusuario                 ( )"
            tipo_conex(3).Caption = "Clandestina                 ( )"
            tipo_conex(4).Caption = "Provisional                 ( )"
            tipo_conex(5).Caption = "No existe                   (X)"
            TipoC = 6
End Select
End Sub

Private Sub tipo_conexion_Click(Index As Integer)
Select Case Index
    Case 0: tipo_conexion(0).Caption = "PVC (X)"
            tipo_conexion(1).Caption = "Galvanizado ( )"
            tipo_conexion(2).Caption = "Manguera ( )"
            tipo_conexion(3).Caption = "Otro ( )"
            MaterialC = 1
    Case 1: tipo_conexion(0).Caption = "PVC ( )"
            tipo_conexion(1).Caption = "Galvanizado (X)"
            tipo_conexion(2).Caption = "Manguera ( )"
            tipo_conexion(3).Caption = "Otro ( )"
            MaterialC = 2
    Case 2: tipo_conexion(0).Caption = "PVC ( )"
            tipo_conexion(1).Caption = "Galvanizado ( )"
            tipo_conexion(2).Caption = "Manguera (X)"
            tipo_conexion(3).Caption = "Otro ( )"
            MaterialC = 3
    Case 3: tipo_conexion(0).Caption = "PVC ( )"
            tipo_conexion(1).Caption = "Galvanizado ( )"
            tipo_conexion(2).Caption = "Manguera ( )"
            tipo_conexion(3).Caption = "Otro (X)"
            MaterialC = 4
End Select
End Sub

Private Sub uso_predio_Click(Index As Integer)
Select Case Index
    Case 0: uso_predio(0).Caption = "Residencial (X)"
            uso_predio(1).Caption = "Comercial ( )"
            uso_predio(2).Caption = "Industrial ( )"
            uso_predio(3).Caption = "Oficial ( )"
            uso_predio(4).Caption = "Mixto ( )"
            UsoP = 1
    Case 1: uso_predio(0).Caption = "Residencial ( )"
            uso_predio(1).Caption = "Comercial (X)"
            uso_predio(2).Caption = "Industrial ( )"
            uso_predio(3).Caption = "Oficial ( )"
            uso_predio(4).Caption = "Mixto ( )"
            UsoP = 2
    Case 2: uso_predio(0).Caption = "Residencial ( )"
            uso_predio(1).Caption = "Comercial ( )"
            uso_predio(2).Caption = "Industrial (X)"
            uso_predio(3).Caption = "Oficial ( )"
            uso_predio(4).Caption = "Mixto ( )"
            UsoP = 3
    Case 3: uso_predio(0).Caption = "Residencial ( )"
            uso_predio(1).Caption = "Comercial ( )"
            uso_predio(2).Caption = "Industrial ( )"
            uso_predio(3).Caption = "Oficial (X)"
            uso_predio(4).Caption = "Mixto ( )"
            UsoP = 4
    Case 4: uso_predio(0).Caption = "Residencial ( )"
            uso_predio(1).Caption = "Comercial ( )"
            uso_predio(2).Caption = "Industrial ( )"
            uso_predio(3).Caption = "Oficial ( )"
            uso_predio(4).Caption = "Mixto (X)"
            UsoP = 5
End Select
End Sub

Private Sub validar_Timer()
If Validar_FORMU2 = True Then
    PADRE.siguiente = True
Else
    PADRE.siguiente.Enabled = False
End If
End Sub
