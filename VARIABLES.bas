Attribute VB_Name = "VARIABLES"
Public Formu As Integer
Public Consulta As String
'-----
Public CONT As Integer
Public Ubicac As Integer
Public EstadoP As Integer
Public VarPos As Integer
Public Abaste As Integer
Public Recolecta As Integer
Public Calidad As Integer
Public CantidadA As Integer
Public NP As Integer
'hijo2
Public UsoP As Integer
Public DiametroC As Integer
Public MaterialC As Integer
Public EstadoM As Integer
Public EstadoC As Integer
Public TipoC As Integer
Public TanqueA As Integer
Public AlmacenaA As Integer
Public HierveA As Integer
'hijo3
Public Operacion As Integer
Public Goteo As Integer
Public ServicioSanitario As Integer
Public ProblemaInstalacion As Integer
Public InodoroL As Integer
Public EstadoCaseta As Integer
Public Alcantarillado As Integer
Public SolucionP As Integer
Public Cual As String
Public CualX As String
'hijo4
Public Basuras As Integer
Public BasuraCasa As Integer
Public EntidadA As Integer
Public Respaldo As Integer

Type Encuesta1
    Nombre As String
    Cedula As Long
    codigo As Long
    ruta As Long
    Ubicacion As Boolean
    NoPisos As Integer
    direccion As String
    NoCatastro As String
    EstadoPredio As Integer
    NumeroPersonas As Integer
    NumeroFamilias As Integer
    NumeroNinos As Integer
    Abastecimiento As Boolean
    OtraFuente As Boolean
    cual1 As String
    OpinionCalidad As Boolean
    OpinionCantidad As Boolean
    UsoPredio As Integer
    DiametroConexion As Integer
    MaterialConexion As Integer
    EstadoMedidor As Integer
    NumeroMedidor As String
    MarcaMedidor As String
    lectura As Long
    EstadoCajilla As Integer
    TipoConexion As Integer
    TanqueAlmacena As Boolean
    AlmacenaAgua As Boolean
    HierveAgua As Integer
    ReparacionesInstalacion As String
    QuienRepara As Integer
    Goteras As Boolean
    TipoServicioSanitario As Integer
    ProblemasInstalacionSanitarias As Boolean
    cual2 As String
    InodoroLimpio As Boolean
    EstadoCaseta As Integer
    TaponadaConexion As Boolean
    SolucionConexion As Integer
    Cuales3 As String
    QueHaceBasuras As Integer
    BasurasCasa As Boolean
    BarridoPorSemana As Integer
    RecoleccionPorSemana As Integer
    OpinionEntidad As Integer
    RespaldoEntidad As Boolean
    Observaciones As String
    DESHABITADA As Boolean
End Type

Public guardar As Encuesta1
Public BusquedaE As Encuesta1
Public comillas As String
Public i As Integer
Public X As Integer
Public Z As Integer
Public SALVADO As Boolean
Public BUSQ As Boolean
Public POSICIONLISTA As Integer
Public tam As Integer
Public COLUMN As Integer

Type COLUMNAS
    TITULO(48) As String
    TAMAÑO(48) As Long
    ACTIVO(48) As Boolean
    tamcarac(48) As Integer
End Type

Public M As COLUMNAS
Public filas As Integer




Type reportes
    ROTULO As String
    cantidad As Long
    PORCEN As Double
    SQL As String
    Numcomparantes As Integer
    COMPARANTES() As Integer
End Type

Public REPOR(0 To 90) As reportes

Public Y As Integer

Public ini As Integer
Public ENCONTRADO As Boolean

'-----generacion reporte personalizado
Public RegIni
Public RegFin

