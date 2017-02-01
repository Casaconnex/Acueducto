VERSION 5.00
Begin VB.MDIForm PADRE 
   BackColor       =   &H00000000&
   Caption         =   "Servicios Públicos"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   8880
   Icon            =   "PADRE.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu opciones 
      Caption         =   "&Opciones"
      Begin VB.Menu guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu nuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu u 
         Caption         =   "-"
      End
      Begin VB.Menu siguiente 
         Caption         =   "Sigu&iente"
         Shortcut        =   ^I
      End
      Begin VB.Menu anterior 
         Caption         =   "An&terior"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu ll 
         Caption         =   "-"
      End
      Begin VB.Menu menup 
         Caption         =   "Menu Principal"
         Shortcut        =   ^M
      End
      Begin VB.Menu ol 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu busque 
         Caption         =   "&Busqueda"
         Shortcut        =   ^B
      End
      Begin VB.Menu personali 
         Caption         =   "Personalizado"
      End
      Begin VB.Menu habitantes 
         Caption         =   "&Habitantes"
         Shortcut        =   ^H
      End
      Begin VB.Menu general 
         Caption         =   "&General"
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu acerca 
         Caption         =   "A&cerca de"
      End
   End
End
Attribute VB_Name = "PADRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Win32 Then
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long
#End If

'constantes a utilizar para inhabilitar los comandos
Const SC_SIZE = &HF000
Const SC_MOVE = &HF010
Const SC_MINIMIZE = &HF020
Const SC_MAXIMIZE = &HF030
Const SC_CLOSE = &HF060
Const SC_RESTORE = &HF120
Const MF_SEPARATOR = &H800
Const MF_BYPOSITION = &H400
Const MF_BYCOMMAND = &H0

Private Sub acerca_Click()
    frmAbout.Show
End Sub

Private Sub anterior_Click()
Select Case Formu
    Case 1: PADRE.anterior.Enabled = False
    Case 2: PADRE.anterior.Enabled = True
            HIJO2.Timer3.Enabled = True
            HIJO1.validar.Enabled = True
            HIJO2.validar.Enabled = False
            Formu = 1
    Case 3: PADRE.anterior.Enabled = True
            HIJO3.Timer3.Enabled = True
            HIJO2.validar.Enabled = True
            HIJO3.validar.Enabled = False
            Formu = 2
    Case 4: PADRE.anterior.Enabled = True
            HIJO4.Timer3.Enabled = True
            HIJO3.validar.Enabled = True
            HIJO4.validar.Enabled = False
            Formu = 3
End Select
End Sub

Private Sub busque_Click()

BUSQUEDA.Show
BUSQUEDA.Timer1.Enabled = True
End Sub

Private Sub general_Click()
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
COLUMN = 48
PLANILLA.Timer1.Enabled = True
PLANILLA.Show
End Sub

Private Sub guardar_Click()
HIJO4.Timer2.Enabled = True
End Sub

Private Sub habitantes_Click()
    habitante.Show
End Sub

Private Sub MDIForm_Load()
'inhabilitar
#If Win32 Then
    Dim hWnd&, hMenu&, Success&
#End If
Dim i%

hWnd = Me.hWnd
hMenu = GetSystemMenu(hWnd, 0)

'quita los menus
Success = DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
Success = DeleteMenu(hMenu, SC_MOVE, MF_BYCOMMAND)
Success = DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)
Success = DeleteMenu(hMenu, SC_RESTORE, MF_BYCOMMAND)
Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
comillas = Chr(34)
'HIJO1.Show
MENU.Show
End Sub


Private Sub MDIForm_Resize()
If PADRE.WindowState <> 1 Then
    PADRE.WindowState = 2
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If SALVADO = False And BUSQ = True Then
    Guardar_BaseDatos BusquedaE
End If
End
End Sub

Private Sub menup_Click()

MENU.Show

End Sub

Private Sub nuevo_Click()
HIJO4.Timer5.Enabled = True
End Sub

Private Sub personali_Click()
Rangos.Show
End Sub

Private Sub salir_Click()
    Unload Me
    
End Sub

Private Sub siguiente_Click()
Select Case Formu
    Case 1:
            HIJO1.Timer2.Enabled = True
            HIJO1.validar.Enabled = False
            HIJO2.validar.Enabled = True
            PADRE.anterior.Enabled = True
            PADRE.siguiente.Enabled = False
            Formu = 2
    Case 2: PADRE.siguiente.Enabled = False
            HIJO2.Timer2.Enabled = True
            HIJO2.validar.Enabled = False
            HIJO3.validar.Enabled = True
            PADRE.anterior.Enabled = True
            Formu = 3
    Case 3: PADRE.siguiente.Enabled = False
            HIJO3.Timer2.Enabled = True
            HIJO3.validar.Enabled = False
            HIJO4.validar.Enabled = True
            PADRE.anterior.Enabled = True
            Formu = 4
    Case 4: PADRE.siguiente.Enabled = False
            
End Select
End Sub
