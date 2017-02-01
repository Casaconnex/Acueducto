VERSION 5.00
Begin VB.Form SELCAMPOS 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección campos a mostrar en Reporte"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7110
   Icon            =   "campos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton boton 
      Caption         =   "A&plicar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   2640
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aplicar sin cerrar la ventana para poder modificar la selecciòn"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton boton 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   5040
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cierra sin realizar cambios en el reporte"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton boton 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Aceptar, cierra y aplica en el reporte"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton boton 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "devolver todos los elementos"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton boton 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "pasar todos los elementos"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton boton 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pasar un  elemento"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton boton 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Devolver un elemento"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ListBox listadestino 
      Height          =   3375
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox listaoriginal 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Campos a mostrar"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista Original de Campos"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "SELCAMPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub boton_Click(Index As Integer)
    Select Case Index

    Case 0:
            If listaoriginal.ListIndex = -1 Then
                If listaoriginal.ListCount > 0 Then
                    listaoriginal.ListIndex = 0
                End If
            End If
            If listaoriginal.ListCount > 0 Then
                listadestino.AddItem listaoriginal.Text
                listaoriginal.RemoveItem listaoriginal.ListIndex
                If listaoriginal.ListCount > 0 Then
                    listaoriginal.ListIndex = 0
                    listaoriginal.ListIndex = 0
                Else
                    listadestino.ListIndex = 0
                End If
            End If
    Case 1:
            If listadestino.ListIndex = -1 Then
                If listadestino.ListCount > 0 Then listadestino.ListIndex = 0
            End If
            If listadestino.ListCount > 0 Then
                listaoriginal.AddItem listadestino.Text
                listadestino.RemoveItem listadestino.ListIndex
                If listadestino.ListCount > 0 Then
                    listaoriginal.ListIndex = 0
                    listadestino.ListIndex = 0
                Else
                    listaoriginal.ListIndex = 0
                End If
            End If
    Case 2:
            If listaoriginal.ListCount > 0 Then
                listadestino.Clear
                listaoriginal.Clear
                For X = 0 To COLUMN
                    listadestino.AddItem M.TITULO(X)
                Next X
                listadestino.ListIndex = 0
            End If
    Case 3:
            If listadestino.ListCount > 0 Then
                listadestino.Clear
                listaoriginal.Clear
                For X = 0 To COLUMN
                    listaoriginal.AddItem M.TITULO(X)
                Next X
                listaoriginal.ListIndex = 0
            End If
            
    Case 4:
            If listadestino.ListCount >= 1 Then
                SELECCION_DE_CAMPOS
                cargar_rejilla PLANILLA.MALLA, M
                SELCAMPOS.Hide
            Else
                SELCAMPOS.Hide
            End If
            
    Case 5: SELCAMPOS.Hide
    Case 6:
        If listadestino.ListCount >= 1 Then
                SELECCION_DE_CAMPOS
                cargar_rejilla PLANILLA.MALLA, M
                
            Else
                SELCAMPOS.Hide
            End If
    
    End Select
End Sub

Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For X = 0 To 6
        If X <> Index Then
            boton(X).FontSize = 12
            boton(X).BackColor = &H8000000F
        Else
            boton(X).FontSize = 22
            boton(X).BackColor = &HFFFF80
            boton(X).SetFocus
        End If
    Next X
End Sub

Private Sub Form_Load()
listaoriginal.Clear
    For X = 0 To COLUMN
        listaoriginal.AddItem M.TITULO(X)
    Next X
    listaoriginal.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For X = 0 To 5
        boton(X).FontSize = 12
        boton(X).BackColor = &H8000000F
    Next X
End Sub


Private Sub SELECCION_DE_CAMPOS()
    Dim COMPARA() As String
    For X = 0 To COLUMN
        M.ACTIVO(X) = False
    Next X
    ReDim COMPARA(0 To listadestino.ListCount - 1)
    For Z = 0 To listadestino.ListCount - 1
        listadestino.ListIndex = Z
        COMPARA(Z) = listadestino.Text
    Next Z
    For X = 0 To COLUMN
    
        For Z = 0 To listadestino.ListCount - 1
            
            If M.TITULO(X) = COMPARA(Z) Then
                M.ACTIVO(X) = True
            End If
        Next Z
    Next X
End Sub

Private Sub listadestino_DblClick()
    boton_Click 1
End Sub

Private Sub listaoriginal_DblClick()
    boton_Click 0
End Sub
