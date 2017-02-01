Attribute VB_Name = "FUNCIONES"
Public Function Validar_letra(Caracter As Integer) As Integer
If Caracter <> 13 Then
    If Caracter <> 8 Then
        If Caracter <> 32 Then
            If (Caracter = 241 Or Caracter = 209) Then
            Validar_letra = Caracter
            Exit Function
            End If
            If Caracter >= 65 And Caracter <= 90 Then
            
                Validar_letra = Caracter
            ElseIf Caracter >= 97 And Caracter <= 122 Then
                Validar_letra = Caracter
            Else
                Validar_letra = 0
            End If
        Else
            Validar_letra = Caracter
        End If
    Else
        Validar_letra = Caracter
    End If
Else
    Validar_letra = Caracter
End If
End Function

Public Function Validar_numero(numero As Integer) As Integer
If numero <> 13 Then
    If numero <> 8 Then
        If numero >= 48 And numero <= 57 Then
            Validar_numero = numero
        Else
            Validar_numero = 0
        End If
    Else
        Validar_numero = numero
    End If
Else
    Validar_numero = numero
End If
End Function

Public Sub Activar_caja()
Select Case VarPos
    Case 0: HIJO1.Nombre.Enabled = True
            HIJO1.Nombre.SetFocus
    Case 1: HIJO1.cc.Enabled = True
            HIJO1.cc.SetFocus
    Case 2: HIJO1.codigo.Enabled = True
            HIJO1.codigo.SetFocus
    Case 3: HIJO1.ruta.Enabled = True
            HIJO1.ruta.SetFocus
    Case 4: HIJO1.pisos.Enabled = True
            HIJO1.pisos.SetFocus
    Case 5: HIJO1.direccion.Enabled = True
            HIJO1.direccion.SetFocus
    Case 6: HIJO1.no_catastro.Enabled = True
            HIJO1.no_catastro.SetFocus
      
    
End Select
End Sub

Public Function Validar_HIJO1() As Boolean
If HIJO1.Nombre.Text = "" Then
    MsgBox "Falta ingresar el nombre del suscriptor!", vbCritical, "NOMBRE SUSCRIPTOR"
    HIJO1.Nombre.Enabled = True
    HIJO1.Nombre.SetFocus
    Validar_HIJO1 = False
    Exit Function
ElseIf HIJO1.ruta.Text = "" Then
    MsgBox "Falta ingresar la ruta!", vbCritical, "RUTA"
    HIJO1.ruta.Enabled = True
    HIJO1.ruta.SetFocus
    Validar_HIJO1 = False
    Exit Function
ElseIf HIJO1.pisos.Text = "" Then
    MsgBox "Falta ingresar el número de pisos!", vbCritical, "PISOS"
    HIJO1.pisos.Enabled = True
    HIJO1.pisos.SetFocus
    Validar_HIJO1 = False
    Exit Function
ElseIf HIJO1.direccion.Text = "" Then
    MsgBox "Falta ingresar la dirección!", vbCritical, "DIRECCION"
    HIJO1.direccion.Enabled = True
    HIJO1.direccion.SetFocus
    Validar_HIJO1 = False
    Exit Function
ElseIf Val(HIJO1.no_ninos.Text) > Val(HIJO1.no_personas.Text) Then
    MsgBox "La cantidad de niños no puede exceder a la cantidad " & vbCrLf & _
        "de personas que viven en la casa.", vbCritical, "NIÑOS MENORES DE 5 AÑOS"
        HIJO1.no_ninos.SelStart = 0
        HIJO1.no_ninos.SelLength = Len(HIJO1.no_ninos.Text)
        HIJO1.no_ninos.SetFocus
        Validar_HIJO1 = False
        Exit Function
ElseIf Recolecta = 1 And HIJO1.cual1.Text = "" Then
        MsgBox "Especifique cual es la Fuente de la cual  se recolecta Agua ", vbInformation, "OTRA FUENTE (SI)"
        HIJO1.cual1.SetFocus
        Validar_HIJO1 = False
        Exit Function
End If
'-------------
If Ubicac = 0 Then
    MsgBox "Para poder continuar debe seleccionar" & vbCrLf & _
        "primero la ubicación de la casa. ", vbInformation, "UBICACION"
        Validar_HIJO1 = False
    Exit Function
ElseIf EstadoP = 0 Then
    MsgBox "Debe seleccionar primero una de las opciones de" & vbCrLf & _
    "estado del predio.", vbInformation, "ESTADO DEL PREDIO"
    Validar_HIJO1 = False
    Exit Function
ElseIf Abaste = 0 Then
    MsgBox "Debe seleccionar primero si el usuario está o no " & vbrclf & _
    "conectado al sistema de abastecimiento de agua.", vbInformation, "ABASTECIMIENTO"
    Validar_HIJO1 = False
    Exit Function
ElseIf Recolecta = 0 Then
    MsgBox "Debe seleccionar primero si se recolecta agua de " & vbrclf & _
    "otra fuente", vbCritical, "FUENTE"
    Validar_HIJO1 = False
    Exit Function
ElseIf Calidad = 0 Then
    MsgBox "Debe seleccionar cual es la calidad del agua!", vbCritical, "CALIDAD DE AGUA"
    Validar_HIJO1 = False
    Exit Function
 ElseIf CantidadA = 0 Then
    MsgBox "Seleccione una de las opciones respecto a la cantidad de agua!", vbCritical, "CANTIDAD DE AGUA"
    Validar_HIJO1 = False
    Exit Function
End If
Validar_HIJO1 = True
PADRE.siguiente.Enabled = True
End Function

Public Function Validar_FORMU1() As Boolean
If HIJO1.Nombre.Text = "" Then
    Validar_FORMU1 = False
    Exit Function
ElseIf HIJO1.ruta.Text = "" Then
    Validar_FORMU1 = False
    Exit Function
ElseIf HIJO1.pisos.Text = "" Then
    Validar_FORMU1 = False
    Exit Function
ElseIf HIJO1.direccion.Text = "" Then
    Validar_FORMU1 = False
    Exit Function
ElseIf Val(HIJO1.no_ninos.Text) > Val(HIJO1.no_personas.Text) Then
    Validar_FORMU1 = False
        Exit Function
ElseIf Recolecta = 1 And HIJO1.cual1.Text = "" Then
        Validar_FORMU1 = False
        Exit Function
End If
'-------------
If Ubicac = 0 Then
    Validar_FORMU1 = False
    Exit Function
ElseIf EstadoP = 0 Then
    Validar_FORMU1 = False
    Exit Function
ElseIf Abaste = 0 Then
    Validar_FORMU1 = False
    Exit Function
ElseIf Recolecta = 0 Then
    Validar_FORMU1 = False
    Exit Function
ElseIf Calidad = 0 Then
    Validar_FORMU1 = False
    Exit Function
ElseIf CantidadA = 0 Then
    Validar_FORMU1 = False
    Exit Function
End If
Validar_FORMU1 = True

End Function
Public Sub Guardar_HIJO1()
guardar.Nombre = HIJO1.Nombre.Text
guardar.Cedula = Val(HIJO1.cc.Text)
guardar.codigo = Val(HIJO1.codigo.Text)
guardar.ruta = Val(HIJO1.ruta.Text)
If Ubicac = 1 Then
    guardar.Ubicacion = True
Else
    guardar.Ubicacion = False
End If
guardar.NoPisos = Val(HIJO1.pisos.Text)
guardar.direccion = HIJO1.direccion.Text
guardar.NoCatastro = HIJO1.no_catastro.Text
guardar.EstadoPredio = EstadoP
guardar.NumeroPersonas = Val(HIJO1.no_personas.Text)
guardar.NumeroFamilias = Val(HIJO1.no_familias.Text)
guardar.NumeroNinos = Val(HIJO1.no_ninos.Text)
If Abaste = 1 Then
    guardar.Abastecimiento = True
Else
    guardar.Abastecimiento = False
End If
If Recolecta = 1 Then
    guardar.OtraFuente = True
Else
    guardar.OtraFuente = False
End If
guardar.cual1 = HIJO1.cual1.Text
If Calidad = 1 Then
    guardar.OpinionCalidad = True
ElseIf Calidad = 2 Then
    guardar.OpinionCalidad = False
End If
If CantidadA = 1 Then
    guardar.OpinionCantidad = True
Else
    guardar.OpinionCantidad = False
End If
'termina hijo1
End Sub

Public Sub Guardar_HIJO2()
guardar.UsoPredio = UsoP
guardar.DiametroConexion = DiametroC
guardar.MaterialConexion = MaterialC
guardar.EstadoMedidor = EstadoM
guardar.NumeroMedidor = HIJO2.no_medidor.Text
guardar.MarcaMedidor = HIJO2.marca_medidor.Text
guardar.lectura = Val(HIJO2.lectura.Text)
guardar.EstadoCajilla = EstadoC
guardar.TipoConexion = TipoC
If TanqueA = 1 Then
    guardar.TanqueAlmacena = True
Else
    guardar.TanqueAlmacena = False
End If
If AlmacenaA = 1 Then
    guardar.AlmacenaAgua = True
Else
    guardar.AlmacenaAgua = False
End If
guardar.HierveAgua = HierveA

End Sub

Public Function Validar_HIJO2() As Boolean
If UsoP = 0 Then
    MsgBox "Marque el uso actual del predio.", vbInformation, "USO DEL PREDIO"
    Validar_HIJO2 = False
    Exit Function
ElseIf DiametroC = 0 Then
    MsgBox "Marque el diametro de la conexión.", vbInformation, "DIAMETRO DE LA CONEXION"
    Validar_HIJO2 = False
    Exit Function
ElseIf MaterialC = 0 Then
    MsgBox "Marque el tipo de material de la conexión.", vbInformation, "MATERIAL DE LA CONEXION"
    Validar_HIJO2 = False
    Exit Function
ElseIf EstadoM = 0 Then
    MsgBox "Marque el estado en que se encuentra el medidor.", vbInformation, "ESTADO DEL MEDIDOR"
    Validar_HIJO2 = False
    Exit Function
ElseIf EstadoC = 0 Then
    MsgBox "Marque el estado en que se encuentra la cajilla.", vbInformation, "ESTADO DE LA CAJILLA"
    Validar_HIJO2 = False
    Exit Function
ElseIf TipoC = 0 Then
    MsgBox "Marque el tipo de conexión.", vbInformation, "TIPO DE CONEXION"
    Validar_HIJO2 = False
    Exit Function
ElseIf TanqueA = 0 Then
    MsgBox "Indique si tiene tanque de almacenamiento.", vbInformation, "TANQUE DE ALMACENAMIENTO"
    Validar_HIJO2 = False
    Exit Function
ElseIf AlmacenaA = 0 Then
    MsgBox "Indique si almacena agua.", vbInformation, "ALMACENA AGUA"
    Validar_HIJO2 = False
    Exit Function
ElseIf HierveA = 0 Then
    MsgBox "Indique si hierve el agua.", vbInformation, "HIERVE AGUA"
    Validar_HIJO2 = False
    Exit Function
End If

Validar_HIJO2 = True
PADRE.siguiente.Enabled = True
PADRE.anterior.Enabled = True
End Function
Public Function Validar_FORMU2() As Boolean
If UsoP = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf DiametroC = 0 Then
   Validar_FORMU2 = False
    Exit Function
ElseIf MaterialC = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf EstadoM = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf EstadoC = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf TipoC = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf TanqueA = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf AlmacenaA = 0 Then
    Validar_FORMU2 = False
    Exit Function
ElseIf HierveA = 0 Then
    Validar_FORMU2 = False
    Exit Function
End If

Validar_FORMU2 = True

End Function

Public Function Validar_FORMU3() As Boolean
If EstadoP <> 1 Then
If HIJO3.Reparaciones.Text = "" Then
    Validar_FORMU3 = False
    Exit Function
ElseIf Operacion = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf Goteo = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf ServicioSanitario = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf ProblemaInstalacion = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf InodoroL = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf EstadoCaseta = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf Alcantarillado = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf SolucionP = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf ProblemaInstalacion = 1 And HIJO3.cual2.Text = "" Then
    Validar_FORMU3 = False
    Exit Function
ElseIf SolucionP = 3 And HIJO3.cual3.Text = "" Then
    Validar_FORMU3 = False
    Exit Function
End If
Else
If Alcantarillado = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf SolucionP = 0 Then
    Validar_FORMU3 = False
    Exit Function
ElseIf ProblemaInstalacion = 1 And HIJO3.cual2.Text = "" Then
    Validar_FORMU3 = False
    Exit Function
ElseIf SolucionP = 3 And HIJO3.cual3.Text = "" Then
    Validar_FORMU3 = False
    Exit Function
End If
End If

Validar_FORMU3 = True
End Function
Public Function Validar_HIJO3() As Boolean
If EstadoP <> 1 Then
If HIJO3.Reparaciones.Text = "" Then
    MsgBox "Especifique las reparaciones  de la instalación.", vbInformation, "REPARACIONES"
    Validar_HIJO3 = False
    Exit Function
ElseIf Operacion = 0 Then
    MsgBox "Especifique quien realiza las repaciones.", vbInformation, "OPERACIONES"
    Validar_HIJO3 = False
    Exit Function
ElseIf Goteo = 0 Then
    MsgBox "Observe si hay llaves, grifos, tuberias o inodoros goteando.", vbInformation, "GOTERAS"
    Validar_HIJO3 = False
    Exit Function
ElseIf ServicioSanitario = 0 Then
    MsgBox "Elija el tipo de servicio sanitario en la vivienda.", vbInformation, "SERVICIO SANITARIO"
    Validar_HIJO3 = False
    Exit Function
ElseIf ProblemaInstalacion = 0 Then
    MsgBox "Especifique si hay problemas con la instalación.", vbInformation, "SERVICIO SANITARIO"
    Validar_HIJO3 = False
    Exit Function
ElseIf InodoroL = 0 Then
    MsgBox "Indique el estado del inodoro.", vbInformation, "INODORO"
    Validar_HIJO3 = False
    Exit Function
ElseIf EstadoCaseta = 0 Then
    MsgBox "Indique en que estado se encuentra la caseta.", vbInformation, "ESTADO DE LA CASETA"
    Validar_HIJO3 = False
    Exit Function
ElseIf Alcantarillado = 0 Then
    MsgBox "Indique si hay problemas de taponamiento en el alcantarillado.", vbInformation, "ALCANTARILLADO"
    Validar_HIJO3 = False
    Exit Function
ElseIf SolucionP = 0 Then
    MsgBox "Elija la solución del problema.", vbInformation, "SOLUCION DEL PROBLEMA"
    Validar_HIJO3 = False
    Exit Function
ElseIf ProblemaInstalacion = 1 And HIJO3.cual2.Text = "" Then
    MsgBox "Indique Cual fue el problema con la instalación sanitaria ", vbInformation, "PROBLEMA INSTALACIÓN SANITARIA"
    Validar_HIJO3 = False
    Exit Function
ElseIf SolucionP = 3 And HIJO3.cual3.Text = "" Then
    MsgBox "Indique que se hizo para solucionar el problema ", vbInformation, "PROBLEMA INSTALACIÓN SANITARIA"
    Validar_HIJO3 = False
    Exit Function
End If
Else
If Alcantarillado = 0 Then
    MsgBox "Indique si hay problemas de taponamiento en el alcantarillado.", vbInformation, "ALCANTARILLADO"
    Validar_HIJO3 = False
    Exit Function
ElseIf SolucionP = 0 Then
    MsgBox "Elija la solución del problema.", vbInformation, "SOLUCION DEL PROBLEMA"
    Validar_HIJO3 = False
    Exit Function
ElseIf ProblemaInstalacion = 1 And HIJO3.cual2.Text = "" Then
    MsgBox "Indique Cual fue el problema con la instalación sanitaria ", vbInformation, "PROBLEMA INSTALACIÓN SANITARIA"
    Validar_HIJO3 = False
    Exit Function
ElseIf SolucionP = 3 And HIJO3.cual3.Text = "" Then
    MsgBox "Indique que se hizo para solucionar el problema ", vbInformation, "PROBLEMA INSTALACIÓN SANITARIA"
    Validar_HIJO3 = False
    Exit Function
End If
End If

Validar_HIJO3 = True
End Function

Public Sub Guardar_HIJO3()
guardar.ReparacionesInstalacion = HIJO3.Reparaciones.Text
guardar.QuienRepara = Operacion
If Goteo = 1 Then
    guardar.Goteras = True
Else
     guardar.Goteras = False
End If
guardar.TipoServicioSanitario = ServicioSanitario
If ProblemaInstalacion = 1 Then
    guardar.ProblemasInstalacionSanitarias = True
Else
     guardar.ProblemasInstalacionSanitarias = False
End If
guardar.cual2 = HIJO3.cual2.Text
If InodoroL = 1 Then
    guardar.InodoroLimpio = True
Else
     guardar.InodoroLimpio = False
End If
guardar.EstadoCaseta = EstadoCaseta
If Alcantarillado = 1 Then
    guardar.TaponadaConexion = True
Else
     guardar.TaponadaConexion = False
End If
guardar.SolucionConexion = SolucionP
If HIJO3.cual3.Enabled = True Then
    guardar.Cuales3 = HIJO3.cual3.Text
ElseIf HIJO3.cual3.Enabled = False Then
    guardar.Cuales3 = ""
End If

End Sub

Public Function Validar_HIJO4() As Boolean
If EstadoP <> 1 Then
If Basuras = 0 Then
    MsgBox "Indique que se hace con las basuras.", vbInformation, "BASURAS"
    Validar_HIJO4 = False
    Exit Function
ElseIf BasuraCasa = 0 Then
    MsgBox "Indique si hay basura en el interior de la casa.", vbInformation, "BASURAS"
    Validar_HIJO4 = False
    Exit Function
ElseIf HIJO4.barrido.Text = "" Then
    MsgBox "Especifique la cantidad de veces que se barre por semana.", vbInformation, "BARRIDOS"
    Validar_HIJO4 = False
    Exit Function
ElseIf HIJO4.recoleccion.Text = "" Then
    MsgBox "Especifique la cantidad de veces que se hace la recolección en la semana.", vbInformation, "BARRIDOS"
    Validar_HIJO4 = False
    Exit Function
ElseIf EntidadA = 0 Then
    MsgBox "Indique la opinión sobre la Administración.", vbInformation, "ENTIDAD"
    Validar_HIJO4 = False
    Exit Function
ElseIf Respaldo = 0 Then
    MsgBox "Indique la opinión respecto al respaldo que ofrece la." & vbCrLf & _
    " la entidad administradora.", vbInformation, "RESPALDO"
    Validar_HIJO4 = False
    Exit Function
End If
Else
If EntidadA = 0 Then
    MsgBox "Indique la opinión sobre la Administración.", vbInformation, "ENTIDAD"
    Validar_HIJO4 = False
    Exit Function
ElseIf Respaldo = 0 Then
    MsgBox "Indique la opinión respecto al respaldo que ofrece la." & vbCrLf & _
    " la entidad administradora.", vbInformation, "RESPALDO"
    Validar_HIJO4 = False
    Exit Function
End If
End If
Validar_HIJO4 = True
End Function
Public Function Validar_FORMU4() As Boolean
If EstadoP <> 1 Then
If Basuras = 0 Then
    Validar_FORMU4 = False
    Exit Function
ElseIf BasuraCasa = 0 Then
    Validar_FORMU4 = False
    Exit Function
ElseIf HIJO4.barrido.Text = "" Then
    Validar_FORMU4 = False
    Exit Function
ElseIf HIJO4.recoleccion.Text = "" Then
    Validar_FORMU4 = False
    Exit Function
ElseIf EntidadA = 0 Then
    Validar_FORMU4 = False
    Exit Function
ElseIf Respaldo = 0 Then
    Validar_FORMU4 = False
    Exit Function
End If
Else
If EntidadA = 0 Then
    Validar_FORMU4 = False
    Exit Function
ElseIf Respaldo = 0 Then
    Validar_FORMU4 = False
    Exit Function
End If
End If
Validar_FORMU4 = True
End Function
Public Sub Guardar_HIJO4()
guardar.QueHaceBasuras = Basuras
If BasuraCasa = 1 Then
    guardar.BasurasCasa = True
Else
    guardar.BasurasCasa = False
End If
guardar.BarridoPorSemana = Val(HIJO4.barrido.Text)
guardar.RecoleccionPorSemana = Val(HIJO4.recoleccion.Text)
guardar.OpinionEntidad = EntidadA
If Respaldo = 1 Then
    guardar.RespaldoEntidad = True
Else
    guardar.RespaldoEntidad = False
End If
guardar.Observaciones = HIJO4.Observaciones.Text

End Sub

Public Sub Guardar_BaseDatos(guardar As Encuesta1)
'conectar bd
HIJO4.Data1.DatabaseName = App.Path + "\Encuesta.mdb"
HIJO4.Data1.RecordSource = "select * from tabla1"
HIJO4.Data1.Refresh
'guardar datos
HIJO4.Data1.Recordset.MoveLast
HIJO4.Data1.Recordset.AddNew
HIJO4.Data1.Recordset!nombre_suscriptor = guardar.Nombre
HIJO4.Data1.Recordset!Cedula = guardar.Cedula
HIJO4.Data1.Recordset!codigo = guardar.codigo
HIJO4.Data1.Recordset!ruta = guardar.ruta
HIJO4.Data1.Recordset!ubicacion_casa = guardar.Ubicacion
HIJO4.Data1.Recordset!no_pisos = guardar.NoPisos
HIJO4.Data1.Recordset!direccion_predio = guardar.direccion
HIJO4.Data1.Recordset!numero_catastral = guardar.NoCatastro
HIJO4.Data1.Recordset!estado_predio = guardar.EstadoPredio
HIJO4.Data1.Recordset!numero_personas_casa = guardar.NumeroPersonas
HIJO4.Data1.Recordset!numero_familias_casa = guardar.NumeroFamilias
HIJO4.Data1.Recordset!numero_menores_5 = guardar.NumeroNinos
HIJO4.Data1.Recordset!conectado_sistema = guardar.Abastecimiento
HIJO4.Data1.Recordset!otra_fuente = guardar.OtraFuente
HIJO4.Data1.Recordset!Cual = guardar.cual1
HIJO4.Data1.Recordset!calidad_agua = guardar.OpinionCalidad
HIJO4.Data1.Recordset!cantidad_agua_suficiente = guardar.OpinionCantidad
HIJO4.Data1.Recordset!uso_predio = guardar.UsoPredio
HIJO4.Data1.Recordset!diametro_conexion = guardar.DiametroConexion
HIJO4.Data1.Recordset!tipo_materiales = guardar.MaterialConexion
HIJO4.Data1.Recordset!estado_medidor = guardar.EstadoMedidor
HIJO4.Data1.Recordset!numero_medidor = guardar.NumeroMedidor
HIJO4.Data1.Recordset!marca_medidor = guardar.MarcaMedidor
HIJO4.Data1.Recordset!lectura = guardar.lectura
HIJO4.Data1.Recordset!estado_cajilla = guardar.EstadoCajilla
HIJO4.Data1.Recordset!tipo_conexion_usuario = guardar.TipoConexion
HIJO4.Data1.Recordset!tanque_almacenamiento = guardar.TanqueAlmacena
HIJO4.Data1.Recordset!almacena_agua = guardar.AlmacenaAgua
HIJO4.Data1.Recordset!hierve_agua = guardar.HierveAgua
HIJO4.Data1.Recordset!reparacion_instalacion = guardar.ReparacionesInstalacion
HIJO4.Data1.Recordset!quien_realiza = guardar.QuienRepara
HIJO4.Data1.Recordset!gotea_llaves_grifos = guardar.Goteras
HIJO4.Data1.Recordset!tipo_servicio_sanitario = guardar.TipoServicioSanitario
HIJO4.Data1.Recordset!problemas_instalacion = guardar.ProblemasInstalacionSanitarias
HIJO4.Data1.Recordset!cuales = guardar.cual2
HIJO4.Data1.Recordset!inodoro_limpio = guardar.InodoroLimpio
HIJO4.Data1.Recordset!estado_caseta = guardar.EstadoCaseta
HIJO4.Data1.Recordset!taponada_conexion = guardar.TaponadaConexion
HIJO4.Data1.Recordset!solucion_problema = guardar.SolucionConexion
HIJO4.Data1.Recordset!cual_solucion = guardar.Cuales3
HIJO4.Data1.Recordset!destino_basuras = guardar.QueHaceBasuras
HIJO4.Data1.Recordset!existencia_basuras_casa = guardar.BasurasCasa
HIJO4.Data1.Recordset!veces_barrido_semana = guardar.BarridoPorSemana
HIJO4.Data1.Recordset!veces_recoleCcion_semana = guardar.RecoleccionPorSemana
HIJO4.Data1.Recordset!opinion_administracion = guardar.OpinionEntidad
HIJO4.Data1.Recordset!respaldo_entidad = guardar.RespaldoEntidad
HIJO4.Data1.Recordset!Observaciones = guardar.Observaciones
If HIJO4.Option1.Value = True Then
    HIJO4.Data1.Recordset!DESABITADA = True
Else
    HIJO4.Data1.Recordset!DESABITADA = False
End If
HIJO4.Data1.Recordset.Update
End Sub

Public Sub Cargar_Form()
Dim i As Integer

'---------------HIJO 1-------------------
'----------------------------------------

HIJO1.Nombre.Text = BusquedaE.Nombre
HIJO1.cc.Text = BusquedaE.Cedula
HIJO1.codigo.Text = BusquedaE.codigo
HIJO1.ruta.Text = BusquedaE.ruta
If BusquedaE.Ubicacion = True Then
    HIJO1.ubica(0).Caption = Mid(HIJO1.ubica(0).Caption, 1, Len(HIJO1.ubica(0).Caption) - 2) & "X" & ")"
    HIJO1.ubica(1).Caption = Mid(HIJO1.ubica(1).Caption, 1, Len(HIJO1.ubica(1).Caption) - 2) & " " & ")"
Else
    HIJO1.ubica(1).Caption = Mid(HIJO1.ubica(1).Caption, 1, Len(HIJO1.ubica(1).Caption) - 2) & "X" & ")"
    HIJO1.ubica(0).Caption = Mid(HIJO1.ubica(0).Caption, 1, Len(HIJO1.ubica(0).Caption) - 2) & " " & ")"
End If
HIJO1.pisos.Text = BusquedaE.NoPisos
HIJO1.direccion.Text = BusquedaE.direccion
HIJO1.no_catastro.Text = BusquedaE.NoCatastro
HIJO1.no_personas.Text = BusquedaE.NumeroPersonas
HIJO1.no_familias.Text = BusquedaE.NumeroFamilias
HIJO1.no_ninos.Text = BusquedaE.NumeroNinos
For i = 0 To 3
    If i + 1 <> BusquedaE.EstadoPredio And i <> 3 Then
        HIJO1.estado(i).Caption = Mid(HIJO1.estado(i).Caption, 1, Len(HIJO1.estado(i).Caption) - 2) & " " & ")"
    ElseIf i <> 3 Then
        HIJO1.estado(i).Caption = Mid(HIJO1.estado(i).Caption, 1, Len(HIJO1.estado(i).Caption) - 2) & "X" & ")"
    End If
Next i
If BusquedaE.Abastecimiento = True Then
    HIJO1.Abastece(0).Caption = Mid(HIJO1.Abastece(0).Caption, 1, Len(HIJO1.Abastece(0).Caption) - 2) & "X" & ")"
    HIJO1.Abastece(1).Caption = Mid(HIJO1.Abastece(1).Caption, 1, Len(HIJO1.Abastece(1).Caption) - 2) & " " & ")"
    Abaste = 1
Else
    HIJO1.Abastece(0).Caption = Mid(HIJO1.Abastece(0).Caption, 1, Len(HIJO1.Abastece(0).Caption) - 2) & " " & ")"
    HIJO1.Abastece(1).Caption = Mid(HIJO1.Abastece(1).Caption, 1, Len(HIJO1.Abastece(1).Caption) - 2) & "X" & ")"
    Abaste = 2
End If


If BusquedaE.OtraFuente = True Then
    HIJO1.fuente(0).Caption = Mid(HIJO1.fuente(0).Caption, 1, Len(HIJO1.fuente(0).Caption) - 2) & "X" & ")"
    HIJO1.fuente(1).Caption = Mid(HIJO1.fuente(1).Caption, 1, Len(HIJO1.fuente(1).Caption) - 2) & " " & ")"
    Recolecta = 1
Else
    HIJO1.fuente(0).Caption = Mid(HIJO1.fuente(0).Caption, 1, Len(HIJO1.fuente(0).Caption) - 2) & " " & ")"
    HIJO1.fuente(1).Caption = Mid(HIJO1.fuente(1).Caption, 1, Len(HIJO1.fuente(1).Caption) - 2) & "X" & ")"
    Recolecta = 2
End If
HIJO1.cual1.Text = BusquedaE.cual1
If BusquedaE.OpinionCalidad = True Then
    HIJO1.calidad_agua(0).Caption = Mid(HIJO1.calidad_agua(0).Caption, 1, Len(HIJO1.calidad_agua(0).Caption) - 2) & "X" & ")"
    HIJO1.calidad_agua(1).Caption = Mid(HIJO1.calidad_agua(1).Caption, 1, Len(HIJO1.calidad_agua(1).Caption) - 2) & " " & ")"
    Calidad = 1
Else
    HIJO1.calidad_agua(0).Caption = Mid(HIJO1.calidad_agua(0).Caption, 1, Len(HIJO1.calidad_agua(0).Caption) - 2) & " " & ")"
    HIJO1.calidad_agua(1).Caption = Mid(HIJO1.calidad_agua(1).Caption, 1, Len(HIJO1.calidad_agua(1).Caption) - 2) & "X" & ")"
    Calidad = 2
End If

If BusquedaE.OpinionCantidad = True Then
    HIJO1.cantidad(0).Caption = Mid(HIJO1.cantidad(0).Caption, 1, Len(HIJO1.cantidad(0).Caption) - 2) & "X" & ")"
    HIJO1.cantidad(1).Caption = Mid(HIJO1.cantidad(1).Caption, 1, Len(HIJO1.cantidad(1).Caption) - 2) & " " & ")"
    CantidadA = 1
Else
    HIJO1.cantidad(0).Caption = Mid(HIJO1.cantidad(0).Caption, 1, Len(HIJO1.cantidad(0).Caption) - 2) & " " & ")"
    HIJO1.cantidad(1).Caption = Mid(HIJO1.cantidad(1).Caption, 1, Len(HIJO1.cantidad(1).Caption) - 2) & "X" & ")"
    CantidadA = 2
End If

'---------------HIJO 2-------------------
'----------------------------------------

For i = 0 To 5
    If i + 1 <> BusquedaE.UsoPredio And i <> 5 Then
        HIJO2.uso_predio(i).Caption = Mid(HIJO2.uso_predio(i).Caption, 1, Len(HIJO2.uso_predio(i).Caption) - 2) & " " & ")"
    ElseIf i <> 5 Then
        HIJO2.uso_predio(i).Caption = Mid(HIJO2.uso_predio(i).Caption, 1, Len(HIJO2.uso_predio(i).Caption) - 2) & "X" & ")"
    End If
Next i
UsoP = BusquedaE.UsoPredio

For i = 0 To 4
    If i + 1 <> BusquedaE.DiametroConexion And i <> 4 Then
        HIJO2.diametro(i).Caption = Mid(HIJO2.diametro(i).Caption, 1, Len(HIJO2.diametro(i).Caption) - 2) & " " & ")"
    ElseIf i <> 4 Then
        HIJO2.diametro(i).Caption = Mid(HIJO2.diametro(i).Caption, 1, Len(HIJO2.diametro(i).Caption) - 2) & "X" & ")"
    End If
Next i
DiametroC = BusquedaE.DiametroConexion

For i = 0 To 4
    If i + 1 <> BusquedaE.MaterialConexion And i <> 4 Then
        HIJO2.tipo_conexion(i).Caption = Mid(HIJO2.tipo_conexion(i).Caption, 1, Len(HIJO2.tipo_conexion(i).Caption) - 2) & " " & ")"
    ElseIf i <> 4 Then
        HIJO2.tipo_conexion(i).Caption = Mid(HIJO2.tipo_conexion(i).Caption, 1, Len(HIJO2.tipo_conexion(i).Caption) - 2) & "X" & ")"
    End If
Next i
MaterialC = BusquedaE.MaterialConexion

For i = 0 To 5
    If i + 1 <> BusquedaE.EstadoMedidor And i <> 5 Then
        HIJO2.estado_medidor(i).Caption = Mid(HIJO2.estado_medidor(i).Caption, 1, Len(HIJO2.estado_medidor(i).Caption) - 2) & " " & ")"
    ElseIf i <> 5 Then
        HIJO2.estado_medidor(i).Caption = Mid(HIJO2.estado_medidor(i).Caption, 1, Len(HIJO2.estado_medidor(i).Caption) - 2) & "X" & ")"
    End If
Next i
EstadoM = BusquedaE.EstadoMedidor


HIJO2.no_medidor.Text = ""
HIJO2.no_medidor.Text = BusquedaE.NumeroMedidor
HIJO2.marca_medidor.Text = ""
HIJO2.marca_medidor.Text = BusquedaE.MarcaMedidor
HIJO2.lectura.Text = ""
HIJO2.lectura.Text = BusquedaE.lectura

For i = 0 To 3
    If i + 1 <> BusquedaE.EstadoCajilla And i <> 3 Then
        HIJO2.estado_cajilla(i).Caption = Mid(HIJO2.estado_cajilla(i).Caption, 1, Len(HIJO2.estado_cajilla(i).Caption) - 2) & " " & ")"
    ElseIf i <> 3 Then
        HIJO2.estado_cajilla(i).Caption = Mid(HIJO2.estado_cajilla(i).Caption, 1, Len(HIJO2.estado_cajilla(i).Caption) - 2) & "X" & ")"
    End If
Next i
EstadoC = BusquedaE.EstadoCajilla

For i = 0 To 6
    If i + 1 <> BusquedaE.TipoConexion And i <> 6 Then
        HIJO2.tipo_conex(i).Caption = Mid(HIJO2.tipo_conex(i).Caption, 1, Len(HIJO2.tipo_conex(i).Caption) - 2) & " " & ")"
    ElseIf i <> 6 Then
        HIJO2.tipo_conex(i).Caption = Mid(HIJO2.tipo_conex(i).Caption, 1, Len(HIJO2.tipo_conex(i).Caption) - 2) & "X" & ")"
    End If
Next i
TipoC = BusquedaE.TipoConexion

If BusquedaE.TanqueAlmacena = True Then
     HIJO2.tanque(0).Caption = Mid(HIJO2.tanque(0).Caption, 1, Len(HIJO2.tanque(0).Caption) - 2) & "X" & ")"
     HIJO2.tanque(1).Caption = Mid(HIJO2.tanque(1).Caption, 1, Len(HIJO2.tanque(1).Caption) - 2) & " " & ")"
     TanqueA = 1
Else
     HIJO2.tanque(0).Caption = Mid(HIJO2.tanque(0).Caption, 1, Len(HIJO2.tanque(0).Caption) - 2) & " " & ")"
     HIJO2.tanque(1).Caption = Mid(HIJO2.tanque(1).Caption, 1, Len(HIJO2.tanque(1).Caption) - 2) & "X" & ")"
     TanqueA = 2
End If

If BusquedaE.AlmacenaAgua = True Then
     HIJO2.consumo(0).Caption = Mid(HIJO2.consumo(0).Caption, 1, Len(HIJO2.consumo(0).Caption) - 2) & "X" & ")"
     HIJO2.consumo(1).Caption = Mid(HIJO2.consumo(1).Caption, 1, Len(HIJO2.consumo(1).Caption) - 2) & " " & ")"
     AlmacenaA = 1
Else
     HIJO2.consumo(0).Caption = Mid(HIJO2.consumo(0).Caption, 1, Len(HIJO2.consumo(0).Caption) - 2) & " " & ")"
     HIJO2.consumo(1).Caption = Mid(HIJO2.consumo(1).Caption, 1, Len(HIJO2.consumo(1).Caption) - 2) & "X" & ")"
     AlmacenaA = 2
End If

For i = 0 To 4
    If i + 1 <> BusquedaE.HierveAgua And i <> 4 Then
        HIJO2.hierve(i).Caption = Mid(HIJO2.hierve(i).Caption, 1, Len(HIJO2.hierve(i).Caption) - 2) & " " & ")"
    ElseIf i <> 4 Then
        HIJO2.hierve(i).Caption = Mid(HIJO2.hierve(i).Caption, 1, Len(HIJO2.hierve(i).Caption) - 2) & "X" & ")"
    End If
Next i
HierveA = BusquedaE.HierveAgua
HIJO2.Hide

'---------------HIJO 3-------------------
'----------------------------------------

HIJO3.Reparaciones.Text = BusquedaE.ReparacionesInstalacion

For i = 0 To 3
    If i + 1 <> BusquedaE.QuienRepara And i <> 3 Then
        HIJO3.operaciones(i).Caption = Mid(HIJO3.operaciones(i).Caption, 1, Len(HIJO3.operaciones(i).Caption) - 2) & " " & ")"
    ElseIf i <> 3 Then
        HIJO3.operaciones(i).Caption = Mid(HIJO3.operaciones(i).Caption, 1, Len(HIJO3.operaciones(i).Caption) - 2) & "X" & ")"
    End If
Next i
Operacion = BusquedaE.QuienRepara

If BusquedaE.Goteras = True Then
    HIJO3.goteando(0).Caption = Mid(HIJO3.goteando(0).Caption, 1, Len(HIJO3.goteando(0).Caption) - 2) & "X" & ")"
    HIJO3.goteando(1).Caption = Mid(HIJO3.goteando(1).Caption, 1, Len(HIJO3.goteando(1).Caption) - 2) & " " & ")"
    Goteo = 1
Else
    HIJO3.goteando(0).Caption = Mid(HIJO3.goteando(0).Caption, 1, Len(HIJO3.goteando(0).Caption) - 2) & " " & ")"
    HIJO3.goteando(1).Caption = Mid(HIJO3.goteando(1).Caption, 1, Len(HIJO3.goteando(1).Caption) - 2) & "X" & ")"
    Goteo = 2
End If

For i = 0 To 4
    If i + 1 <> BusquedaE.TipoServicioSanitario And i <> 4 Then
        HIJO3.servicio(i).Caption = Mid(HIJO3.servicio(i).Caption, 1, Len(HIJO3.servicio(i).Caption) - 2) & " " & ")"
    ElseIf i <> 4 Then
        HIJO3.servicio(i).Caption = Mid(HIJO3.servicio(i).Caption, 1, Len(HIJO3.servicio(i).Caption) - 2) & "X" & ")"
    End If
Next i
ServicioSanitario = BusquedaE.TipoServicioSanitario

If BusquedaE.ProblemasInstalacionSanitarias = True Then
    HIJO3.instal(0).Caption = Mid(HIJO3.instal(0).Caption, 1, Len(HIJO3.instal(0).Caption) - 2) & "X" & ")"
    HIJO3.instal(1).Caption = Mid(HIJO3.instal(1).Caption, 1, Len(HIJO3.instal(1).Caption) - 2) & " " & ")"
    ProblemaInstalacion = 1
Else
    HIJO3.instal(0).Caption = Mid(HIJO3.instal(0).Caption, 1, Len(HIJO3.instal(0).Caption) - 2) & " " & ")"
    HIJO3.instal(1).Caption = Mid(HIJO3.instal(1).Caption, 1, Len(HIJO3.instal(1).Caption) - 2) & "X" & ")"
    ProblemaInstalacion = 2
End If

HIJO3.cual2.Text = BusquedaE.cual2

If BusquedaE.InodoroLimpio = True Then
    HIJO3.limpio(0).Caption = Mid(HIJO3.limpio(0).Caption, 1, Len(HIJO3.limpio(0).Caption) - 2) & "X" & ")"
    HIJO3.limpio(1).Caption = Mid(HIJO3.limpio(1).Caption, 1, Len(HIJO3.limpio(1).Caption) - 2) & " " & ")"
    InodoroL = 1
Else
    HIJO3.limpio(0).Caption = Mid(HIJO3.limpio(0).Caption, 1, Len(HIJO3.limpio(0).Caption) - 2) & " " & ")"
    HIJO3.limpio(1).Caption = Mid(HIJO3.limpio(1).Caption, 1, Len(HIJO3.limpio(1).Caption) - 2) & "X" & ")"
    InodoroL = 2
End If

For i = 0 To 3
    If i + 1 <> BusquedaE.EstadoCaseta And i <> 3 Then
        HIJO3.caseta(i).Caption = Mid(HIJO3.caseta(i).Caption, 1, Len(HIJO3.caseta(i).Caption) - 2) & " " & ")"
    ElseIf i <> 3 Then
        HIJO3.caseta(i).Caption = Mid(HIJO3.caseta(i).Caption, 1, Len(HIJO3.caseta(i).Caption) - 2) & "X" & ")"
    End If
Next i
EstadoCaseta = BusquedaE.EstadoCaseta

If BusquedaE.TaponadaConexion = True Then
    HIJO3.taponado(0).Caption = Mid(HIJO3.taponado(0).Caption, 1, Len(HIJO3.taponado(0).Caption) - 2) & "X" & ")"
    HIJO3.taponado(1).Caption = Mid(HIJO3.taponado(1).Caption, 1, Len(HIJO3.taponado(1).Caption) - 2) & " " & ")"
    Alcantarillado = 1
Else
    HIJO3.taponado(0).Caption = Mid(HIJO3.taponado(0).Caption, 1, Len(HIJO3.taponado(0).Caption) - 2) & " " & ")"
    HIJO3.taponado(1).Caption = Mid(HIJO3.taponado(1).Caption, 1, Len(HIJO3.taponado(1).Caption) - 2) & "X" & ")"
    Alcantarillado = 2
End If

For i = 0 To 3
    If i + 1 <> BusquedaE.SolucionConexion And i <> 3 Then
        HIJO3.solucion(i).Caption = Mid(HIJO3.solucion(i).Caption, 1, Len(HIJO3.solucion(i).Caption) - 2) & " " & ")"
    ElseIf i <> 3 Then
        HIJO3.solucion(i).Caption = Mid(HIJO3.solucion(i).Caption, 1, Len(HIJO3.solucion(i).Caption) - 2) & "X" & ")"
    End If
Next i
SolucionP = BusquedaE.SolucionConexion
HIJO3.cual3.Text = BusquedaE.Cuales3

HIJO3.Hide

'---------------HIJO 4-------------------
'----------------------------------------

For i = 0 To 4
    If i + 1 <> BusquedaE.QueHaceBasuras And i <> 4 Then
        HIJO4.basura(i).Caption = Mid(HIJO4.basura(i).Caption, 1, Len(HIJO4.basura(i).Caption) - 2) & " " & ")"
    ElseIf i <> 4 Then
        HIJO4.basura(i).Caption = Mid(HIJO4.basura(i).Caption, 1, Len(HIJO4.basura(i).Caption) - 2) & "X" & ")"
    End If
Next i
Basuras = BusquedaE.QueHaceBasuras

If BusquedaE.BasurasCasa = True Then
    HIJO4.interior(0).Caption = Mid(HIJO4.interior(0).Caption, 1, Len(HIJO4.interior(0).Caption) - 2) & "X" & ")"
    HIJO4.interior(1).Caption = Mid(HIJO4.interior(1).Caption, 1, Len(HIJO4.interior(1).Caption) - 2) & " " & ")"
    BasuraCasa = 1
Else
    HIJO4.interior(0).Caption = Mid(HIJO4.interior(0).Caption, 1, Len(HIJO4.interior(0).Caption) - 2) & " " & ")"
    HIJO4.interior(1).Caption = Mid(HIJO4.interior(1).Caption, 1, Len(HIJO4.interior(1).Caption) - 2) & "X" & ")"
    BasuraCasa = 2
End If

HIJO4.barrido.Text = BusquedaE.BarridoPorSemana
HIJO4.recoleccion.Text = BusquedaE.RecoleccionPorSemana

For i = 0 To 3
    If i + 1 <> BusquedaE.OpinionEntidad And i <> 3 Then
        HIJO4.opina(i).Caption = Mid(HIJO4.opina(i).Caption, 1, Len(HIJO4.opina(i).Caption) - 2) & " " & ")"
    ElseIf i <> 3 Then
        HIJO4.opina(i).Caption = Mid(HIJO4.opina(i).Caption, 1, Len(HIJO4.opina(i).Caption) - 2) & "X" & ")"
    End If
Next i
EntidadA = BusquedaE.OpinionEntidad

If BusquedaE.RespaldoEntidad = True Then
    HIJO4.entidad(0).Caption = Mid(HIJO4.entidad(0).Caption, 1, Len(HIJO4.entidad(0).Caption) - 2) & "X" & ")"
    HIJO4.entidad(1).Caption = Mid(HIJO4.entidad(1).Caption, 1, Len(HIJO4.entidad(1).Caption) - 2) & " " & ")"
    Respaldo = 1
Else
    HIJO4.entidad(0).Caption = Mid(HIJO4.entidad(0).Caption, 1, Len(HIJO4.entidad(0).Caption) - 2) & " " & ")"
    HIJO4.entidad(1).Caption = Mid(HIJO4.entidad(1).Caption, 1, Len(HIJO4.entidad(1).Caption) - 2) & "X" & ")"
    Respaldo = 2
End If

HIJO4.Observaciones.Text = BusquedaE.Observaciones

If BusquedaE.DESHABITADA = True Then
    HIJO4.Option1.Value = True
Else
    HIJO4.Option2.Value = True
End If
HIJO4.Timer4.Enabled = False
HIJO4.Hide
HIJO1.WindowState = 2
HIJO1.Show
End Sub

Public Sub CARGAR_MALLA(MALLA As COLUMNAS)
Dim max As Integer
MALLA.TITULO(0) = "No."
MALLA.TITULO(1) = "Nombre"
MALLA.TITULO(2) = "Cédula"
MALLA.TITULO(3) = "Código"
MALLA.TITULO(4) = "Ruta"
MALLA.TITULO(5) = "Ubicación de la Casa"
MALLA.TITULO(6) = "No Pisos"
MALLA.TITULO(7) = "Dirección Predio"
MALLA.TITULO(8) = "No. Catastro"
MALLA.TITULO(9) = "Estado Predio"
MALLA.TITULO(10) = "No. Personas"
MALLA.TITULO(11) = "No. Familias"
MALLA.TITULO(12) = "No Niños < 5"
MALLA.TITULO(13) = "Conec. Sistema"
MALLA.TITULO(14) = "Otra Fuente"
MALLA.TITULO(15) = "Cual"
MALLA.TITULO(16) = "Calidad Agua"
MALLA.TITULO(17) = "Cantidad Suficiente"
MALLA.TITULO(18) = "Uso Predio"
MALLA.TITULO(19) = "Diametro Conex"
MALLA.TITULO(20) = "Material Conexión"
MALLA.TITULO(21) = "Estado Medidor"
MALLA.TITULO(22) = "Número Medidor"
MALLA.TITULO(23) = "Marca Medidor"
MALLA.TITULO(24) = "Lectura"
MALLA.TITULO(25) = "Estado Cajilla"
MALLA.TITULO(26) = "Tipo Conexión"
MALLA.TITULO(27) = "Tanque Almac"
MALLA.TITULO(28) = "Almac a. Cons"
MALLA.TITULO(29) = "Hierve Agua"
MALLA.TITULO(30) = "Reparaciones Instalaciones Agua"
MALLA.TITULO(31) = "Quien Repara"
MALLA.TITULO(32) = "G-t-i Goteando"
MALLA.TITULO(33) = "Tipo Servicio Sanitario"
MALLA.TITULO(34) = "Problemas Instala"
MALLA.TITULO(35) = "Cuales"
MALLA.TITULO(36) = "Inodoro Limpio"
MALLA.TITULO(37) = "Estado Caseta"
MALLA.TITULO(38) = "Taponamiento Conex"
MALLA.TITULO(39) = "Solución"
MALLA.TITULO(40) = "Cual"
MALLA.TITULO(41) = "Destino Basuras"
MALLA.TITULO(42) = "Basura Interior Casa"
MALLA.TITULO(43) = "No. Veces barrido/semana"
MALLA.TITULO(44) = "No. Veces Recoleción/semana"
MALLA.TITULO(45) = "Opinión Respecto Servicios"
MALLA.TITULO(46) = "Respaldo Entidad"
MALLA.TITULO(47) = "Observaciones"
MALLA.TITULO(48) = "Deshabitada"

For j = 0 To 48
    MALLA.TAMAÑO(j) = 1300
Next j

MALLA.TAMAÑO(0) = 500
MALLA.TAMAÑO(1) = 2500
MALLA.TAMAÑO(5) = 2000
MALLA.TAMAÑO(6) = 900
MALLA.TAMAÑO(7) = 2500
MALLA.TAMAÑO(9) = 1500
MALLA.TAMAÑO(15) = 1500
MALLA.TAMAÑO(16) = 800
MALLA.TAMAÑO(17) = 800
MALLA.TAMAÑO(17) = 1500
MALLA.TAMAÑO(18) = 1500
MALLA.TAMAÑO(21) = 1500
MALLA.TAMAÑO(26) = 1500
MALLA.TAMAÑO(29) = 1500
MALLA.TAMAÑO(30) = 2000
MALLA.TAMAÑO(33) = 3800
MALLA.TAMAÑO(34) = 1500
MALLA.TAMAÑO(35) = 2000
MALLA.TAMAÑO(36) = 1500
MALLA.TAMAÑO(37) = 1500
MALLA.TAMAÑO(38) = 1500
MALLA.TAMAÑO(39) = 3000
MALLA.TAMAÑO(40) = 2000
MALLA.TAMAÑO(41) = 2000
MALLA.TAMAÑO(42) = 2000
MALLA.TAMAÑO(43) = 2500
MALLA.TAMAÑO(44) = 3000
MALLA.TAMAÑO(45) = 3000
MALLA.TAMAÑO(46) = 3000
MALLA.TAMAÑO(47) = 4500


'malla.tamcarac(1)=
For j = 0 To 48
    MALLA.ACTIVO(j) = True
Next j
End Sub
Public Sub cargar_rejilla(rejillas As MSFlexGrid, mallas As COLUMNAS)
    For j = 0 To 48
        rejillas.TextMatrix(0, j) = mallas.TITULO(j)
        If mallas.ACTIVO(j) = True Then
            rejillas.ColWidth(j) = mallas.TAMAÑO(j)
        Else
            rejillas.ColWidth(j) = 0
        End If
    Next j
End Sub
Public Sub CARGAR_REPORTES(base As Data, REPOR() As reportes)
    
    '-*-*-*-*-*-* cargando base de datos*-*-*-*-*-
    base.DatabaseName = App.Path + "\encuesta.mdb"
        
    '-*-*-*-*-*-* cargando datos sobre los reportes*-*-*-*-
    REPOR(0).ROTULO = "TOTAL ENCUESTAS"
    REPOR(0).SQL = "SELECT COUNT(RUTA) AS CANTI FROM TABLA1"
    REPOR(0).Numcomparantes = 0
    
    REPOR(1).ROTULO = "Encuestas sin Código"
    REPOR(1).SQL = "SELECT COUNT(CODIGO) AS CANTI FROM TABLA1 WHERE CODIGO=0"
    REPOR(1).Numcomparantes = 1
    ReDim REPOR(1).COMPARANTES(1 To 1)
    REPOR(1).COMPARANTES(1) = 0
    
    REPOR(2).ROTULO = "Encuestas sin Ruta"
    REPOR(2).SQL = "SELECT COUNT(ruta) AS CANTI FROM TABLA1 WHERE RUTA=0 "
    REPOR(2).Numcomparantes = 1
    ReDim REPOR(2).COMPARANTES(1 To 1)
    REPOR(2).COMPARANTES(1) = 0
    
    REPOR(3).ROTULO = "Predios Urbanos"
    REPOR(3).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Ubicacion_casa= True"
    REPOR(3).Numcomparantes = 2
    ReDim REPOR(3).COMPARANTES(1 To 2)
    REPOR(3).COMPARANTES(1) = 0
    REPOR(3).COMPARANTES(2) = 4
    
    REPOR(4).ROTULO = "Predios Rurales"
    REPOR(4).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Ubicacion_casa= FALSE"
    REPOR(4).Numcomparantes = 2
    ReDim REPOR(4).COMPARANTES(1 To 2)
    REPOR(4).COMPARANTES(1) = 0
    REPOR(4).COMPARANTES(2) = 3
    
    REPOR(5).ROTULO = "Estado Predio: Lote"
    REPOR(5).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_predio= 1"
    REPOR(5).Numcomparantes = 2
    ReDim REPOR(5).COMPARANTES(1 To 2)
    REPOR(5).COMPARANTES(1) = 5
    REPOR(5).COMPARANTES(2) = 6
    
    REPOR(6).ROTULO = "Estado Predio: En Construcción"
    REPOR(6).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_predio= 2"
    REPOR(6).Numcomparantes = 2
    ReDim REPOR(6).COMPARANTES(1 To 2)
    REPOR(6).COMPARANTES(1) = 4
    REPOR(6).COMPARANTES(2) = 6
    
    REPOR(7).ROTULO = "Estado Predio: Construido"
    REPOR(7).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_predio= 3"
    REPOR(7).Numcomparantes = 2
    ReDim REPOR(7).COMPARANTES(1 To 2)
    REPOR(7).COMPARANTES(1) = 4
    REPOR(7).COMPARANTES(2) = 5
    
    REPOR(8).ROTULO = "Conexión Sistema: Sí"
    REPOR(8).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Conectado_sistema= true"
    REPOR(8).Numcomparantes = 1
    ReDim REPOR(8).COMPARANTES(1 To 1)
    REPOR(8).COMPARANTES(1) = 9
    
    REPOR(9).ROTULO = "Conexión Sistema: No"
    REPOR(9).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Conectado_sistema= False"
    REPOR(9).Numcomparantes = 1
    ReDim REPOR(9).COMPARANTES(1 To 1)
    REPOR(9).COMPARANTES(1) = 8
    
    REPOR(10).ROTULO = "Otra Fuente: Sí"
    REPOR(10).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Otra_fuente= True AND DESABITADA=FALSE"
    REPOR(10).Numcomparantes = 1
    ReDim REPOR(10).COMPARANTES(1 To 1)
    REPOR(10).COMPARANTES(1) = 11
        
    REPOR(11).ROTULO = "Otra Fuente: No"
    REPOR(11).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Otra_fuente= False AND DESABITADA=FALSE"
    REPOR(11).Numcomparantes = 1
    ReDim REPOR(11).COMPARANTES(1 To 1)
    REPOR(11).COMPARANTES(1) = 10
    
    REPOR(12).ROTULO = "Opinión Calidad: Buena"
    REPOR(12).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Calidad_agua= True AND DESABITADA=FALSE"
    REPOR(12).Numcomparantes = 1
    ReDim REPOR(12).COMPARANTES(1 To 1)
    REPOR(12).COMPARANTES(1) = 13
        
    REPOR(13).ROTULO = "Opinión Calidad: Mala"
    REPOR(13).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Calidad_agua= False AND DESABITADA=FALSE"
    REPOR(13).Numcomparantes = 1
    ReDim REPOR(13).COMPARANTES(1 To 1)
    REPOR(13).COMPARANTES(1) = 12
    
    REPOR(14).ROTULO = "Cantidad de Agua Suficiente necesidades: Sí "
    REPOR(14).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Cantidad_agua_suficiente= True AND DESABITADA=FALSE"
    REPOR(14).Numcomparantes = 1
    ReDim REPOR(14).COMPARANTES(1 To 1)
    REPOR(14).COMPARANTES(1) = 15
    
    REPOR(15).ROTULO = "Cantidad de Agua Suficiente necesidades: No "
    REPOR(15).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Cantidad_agua_suficiente= False AND DESABITADA=FALSE"
    REPOR(15).Numcomparantes = 1
    ReDim REPOR(15).COMPARANTES(1 To 1)
    REPOR(15).COMPARANTES(1) = 14
    
    REPOR(16).ROTULO = "Uso Predio: Residencial"
    REPOR(16).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Uso_predio= 1"
    REPOR(16).Numcomparantes = 4
    ReDim REPOR(16).COMPARANTES(1 To 4)
    REPOR(16).COMPARANTES(1) = 17
    REPOR(16).COMPARANTES(2) = 18
    REPOR(16).COMPARANTES(3) = 19
    REPOR(16).COMPARANTES(4) = 20
    
    REPOR(17).ROTULO = "Uso Predio: Comercial"
    REPOR(17).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Uso_predio= 2"
    REPOR(17).Numcomparantes = 4
    ReDim REPOR(17).COMPARANTES(1 To 4)
    REPOR(17).COMPARANTES(1) = 16
    REPOR(17).COMPARANTES(2) = 18
    REPOR(17).COMPARANTES(3) = 19
    REPOR(17).COMPARANTES(4) = 20
    
    REPOR(18).ROTULO = "Uso Predio: Industrial"
    REPOR(18).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Uso_predio= 3"
    REPOR(18).Numcomparantes = 4
    ReDim REPOR(18).COMPARANTES(1 To 4)
    REPOR(18).COMPARANTES(1) = 16
    REPOR(18).COMPARANTES(2) = 17
    REPOR(18).COMPARANTES(3) = 19
    REPOR(18).COMPARANTES(4) = 20
    
    REPOR(19).ROTULO = "Uso Predio: Oficial"
    REPOR(19).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Uso_predio= 4"
    REPOR(19).Numcomparantes = 4
    ReDim REPOR(19).COMPARANTES(1 To 4)
    REPOR(19).COMPARANTES(1) = 16
    REPOR(19).COMPARANTES(2) = 17
    REPOR(19).COMPARANTES(3) = 18
    REPOR(19).COMPARANTES(4) = 20
    
    REPOR(20).ROTULO = "Uso Predio: Mixto"
    REPOR(20).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Uso_predio= 5"
    REPOR(20).Numcomparantes = 4
    ReDim REPOR(20).COMPARANTES(1 To 4)
    REPOR(20).COMPARANTES(1) = 16
    REPOR(20).COMPARANTES(2) = 17
    REPOR(20).COMPARANTES(3) = 18
    REPOR(20).COMPARANTES(4) = 19
    
    REPOR(21).ROTULO = "Uso Predio: No determinado"
    REPOR(21).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Uso_predio= 0"
    REPOR(21).Numcomparantes = 0
    
    REPOR(22).ROTULO = "Diametro conexión: 1/2'"
    REPOR(22).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Diametro_conexion= 1"
    REPOR(22).Numcomparantes = 3
    ReDim REPOR(22).COMPARANTES(1 To 3)
    REPOR(22).COMPARANTES(1) = 23
    REPOR(22).COMPARANTES(2) = 24
    REPOR(22).COMPARANTES(3) = 25
    
    REPOR(23).ROTULO = "Diametro conexión: 3/4'"
    REPOR(23).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Diametro_conexion= 2"
    REPOR(23).Numcomparantes = 3
    ReDim REPOR(23).COMPARANTES(1 To 3)
    REPOR(23).COMPARANTES(1) = 22
    REPOR(23).COMPARANTES(2) = 24
    REPOR(23).COMPARANTES(3) = 25
    
    REPOR(24).ROTULO = "Diametro conexión: 1'"
    REPOR(24).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Diametro_conexion= 3"
    REPOR(24).Numcomparantes = 3
    ReDim REPOR(24).COMPARANTES(1 To 3)
    REPOR(24).COMPARANTES(1) = 22
    REPOR(24).COMPARANTES(2) = 23
    REPOR(24).COMPARANTES(3) = 25
    
    REPOR(25).ROTULO = "Diametro conexión: >1'"
    REPOR(25).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Diametro_conexion= 4"
    REPOR(25).Numcomparantes = 3
    ReDim REPOR(25).COMPARANTES(1 To 3)
    REPOR(25).COMPARANTES(1) = 22
    REPOR(25).COMPARANTES(2) = 23
    REPOR(25).COMPARANTES(3) = 24
    
    REPOR(26).ROTULO = "Material conexión: P.V.C."
    REPOR(26).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_materiales= 1"
    REPOR(26).Numcomparantes = 3
    ReDim REPOR(26).COMPARANTES(1 To 3)
    REPOR(26).COMPARANTES(1) = 27
    REPOR(26).COMPARANTES(2) = 28
    REPOR(26).COMPARANTES(3) = 29
    
    REPOR(27).ROTULO = "Material conexión: Galvanizado"
    REPOR(27).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_materiales= 2"
    REPOR(27).Numcomparantes = 3
    ReDim REPOR(27).COMPARANTES(1 To 3)
    REPOR(27).COMPARANTES(1) = 26
    REPOR(27).COMPARANTES(2) = 28
    REPOR(27).COMPARANTES(3) = 29
    
    REPOR(28).ROTULO = "Material conexión: Manguera"
    REPOR(28).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_materiales= 3"
    REPOR(28).Numcomparantes = 3
    ReDim REPOR(28).COMPARANTES(1 To 3)
    REPOR(28).COMPARANTES(1) = 26
    REPOR(28).COMPARANTES(2) = 27
    REPOR(28).COMPARANTES(3) = 29
    
    REPOR(29).ROTULO = "Material conexión: Otro"
    REPOR(29).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_materiales= 4"
    REPOR(29).Numcomparantes = 3
    ReDim REPOR(29).COMPARANTES(1 To 3)
    REPOR(29).COMPARANTES(1) = 26
    REPOR(29).COMPARANTES(2) = 27
    REPOR(29).COMPARANTES(3) = 28
    
    REPOR(30).ROTULO = "Estado Medidor: Registrando"
    REPOR(30).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_medidor= 1"
    REPOR(30).Numcomparantes = 4
    ReDim REPOR(30).COMPARANTES(1 To 4)
    REPOR(30).COMPARANTES(1) = 31
    REPOR(30).COMPARANTES(2) = 32
    REPOR(30).COMPARANTES(3) = 33
    REPOR(30).COMPARANTES(4) = 34
    
    REPOR(31).ROTULO = "Estado Medidor: Detenido"
    REPOR(31).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_medidor= 2"
    REPOR(31).Numcomparantes = 4
    ReDim REPOR(31).COMPARANTES(1 To 4)
    REPOR(31).COMPARANTES(1) = 30
    REPOR(31).COMPARANTES(2) = 32
    REPOR(31).COMPARANTES(3) = 33
    REPOR(31).COMPARANTES(4) = 34
    
    REPOR(32).ROTULO = "Estado Medidor: Nublado"
    REPOR(32).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_medidor= 3"
    REPOR(32).Numcomparantes = 4
    ReDim REPOR(32).COMPARANTES(1 To 4)
    REPOR(32).COMPARANTES(1) = 30
    REPOR(32).COMPARANTES(2) = 31
    REPOR(32).COMPARANTES(3) = 33
    REPOR(32).COMPARANTES(4) = 34
    
    REPOR(33).ROTULO = "Estado Medidor: Dañado"
    REPOR(33).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_medidor= 4"
    REPOR(33).Numcomparantes = 4
    ReDim REPOR(33).COMPARANTES(1 To 4)
    REPOR(33).COMPARANTES(1) = 30
    REPOR(33).COMPARANTES(2) = 31
    REPOR(33).COMPARANTES(3) = 32
    REPOR(33).COMPARANTES(4) = 34
    
    REPOR(34).ROTULO = "Estado Medidor: Sin medidor"
    REPOR(34).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_medidor= 4"
    REPOR(34).Numcomparantes = 4
    ReDim REPOR(34).COMPARANTES(1 To 4)
    REPOR(34).COMPARANTES(1) = 30
    REPOR(34).COMPARANTES(2) = 31
    REPOR(34).COMPARANTES(3) = 32
    REPOR(34).COMPARANTES(4) = 33
    
    REPOR(35).ROTULO = "Estado Cajilla: Bueno"
    REPOR(35).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_cajilla= 1"
    REPOR(35).Numcomparantes = 2
    ReDim REPOR(35).COMPARANTES(1 To 2)
    REPOR(35).COMPARANTES(1) = 36
    REPOR(35).COMPARANTES(2) = 37
    
    REPOR(36).ROTULO = "Estado Cajilla: Malo"
    REPOR(36).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_cajilla= 2"
    REPOR(36).Numcomparantes = 2
    ReDim REPOR(36).COMPARANTES(1 To 2)
    REPOR(36).COMPARANTES(1) = 35
    REPOR(36).COMPARANTES(2) = 37
    
    REPOR(37).ROTULO = "Estado Cajilla: No Existe"
    REPOR(37).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_cajilla= 3"
    REPOR(37).Numcomparantes = 2
    ReDim REPOR(37).COMPARANTES(1 To 2)
    REPOR(37).COMPARANTES(1) = 35
    REPOR(37).COMPARANTES(2) = 36
    
    REPOR(38).ROTULO = "Tipo de Conexión: Legal"
    REPOR(38).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_conexion_usuario= 1"
    REPOR(38).Numcomparantes = 5
    ReDim REPOR(38).COMPARANTES(1 To 5)
    REPOR(38).COMPARANTES(1) = 39
    REPOR(38).COMPARANTES(2) = 40
    REPOR(38).COMPARANTES(3) = 41
    REPOR(38).COMPARANTES(4) = 42
    REPOR(38).COMPARANTES(5) = 43
    
    REPOR(39).ROTULO = "Tipo de Conexión: No Incluida Sistema"
    REPOR(39).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_conexion_usuario= 2"
    REPOR(39).Numcomparantes = 5
    ReDim REPOR(39).COMPARANTES(1 To 5)
    REPOR(39).COMPARANTES(1) = 38
    REPOR(39).COMPARANTES(2) = 40
    REPOR(39).COMPARANTES(3) = 41
    REPOR(39).COMPARANTES(4) = 42
    REPOR(39).COMPARANTES(5) = 43
    
    REPOR(40).ROTULO = "Tipo de Conexión: Multiusuario"
    REPOR(40).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_conexion_usuario= 3"
    REPOR(40).Numcomparantes = 5
    ReDim REPOR(40).COMPARANTES(1 To 5)
    REPOR(40).COMPARANTES(1) = 38
    REPOR(40).COMPARANTES(2) = 39
    REPOR(40).COMPARANTES(3) = 41
    REPOR(40).COMPARANTES(4) = 42
    REPOR(40).COMPARANTES(5) = 43
    
    REPOR(41).ROTULO = "Tipo de Conexión: Clandestina"
    REPOR(41).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_conexion_usuario= 4"
    REPOR(41).Numcomparantes = 5
    ReDim REPOR(41).COMPARANTES(1 To 5)
    REPOR(41).COMPARANTES(1) = 38
    REPOR(41).COMPARANTES(2) = 39
    REPOR(41).COMPARANTES(3) = 40
    REPOR(41).COMPARANTES(4) = 42
    REPOR(41).COMPARANTES(5) = 43
    
    REPOR(42).ROTULO = "Tipo de Conexión: Provisional"
    REPOR(42).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_conexion_usuario= 5"
    REPOR(42).Numcomparantes = 5
    ReDim REPOR(42).COMPARANTES(1 To 5)
    REPOR(42).COMPARANTES(1) = 38
    REPOR(42).COMPARANTES(2) = 39
    REPOR(42).COMPARANTES(3) = 40
    REPOR(42).COMPARANTES(4) = 41
    REPOR(42).COMPARANTES(5) = 43
    
    REPOR(43).ROTULO = "Tipo de Conexión: No existe"
    REPOR(43).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_conexion_usuario= 6"
    REPOR(43).Numcomparantes = 5
    ReDim REPOR(43).COMPARANTES(1 To 5)
    REPOR(43).COMPARANTES(1) = 38
    REPOR(43).COMPARANTES(2) = 39
    REPOR(43).COMPARANTES(3) = 40
    REPOR(43).COMPARANTES(4) = 41
    REPOR(43).COMPARANTES(5) = 42
    
    REPOR(44).ROTULO = "Tanque de almacenamiento: Sí"
    REPOR(44).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tanque_almacenamiento= True AND Desabitada=FALSE"
    REPOR(44).Numcomparantes = 1
    ReDim REPOR(44).COMPARANTES(1 To 1)
    REPOR(44).COMPARANTES(1) = 45
    
    REPOR(45).ROTULO = "Tanque de almacenamiento: No"
    REPOR(45).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tanque_almacenamiento= False AND Desabitada=FALSE"
    REPOR(45).Numcomparantes = 1
    ReDim REPOR(45).COMPARANTES(1 To 1)
    REPOR(45).COMPARANTES(1) = 44
    
    REPOR(46).ROTULO = "Almacena Agua Consumo: Sí"
    REPOR(46).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Almacena_agua= True AND Desabitada=FALSE"
    REPOR(46).Numcomparantes = 1
    ReDim REPOR(46).COMPARANTES(1 To 1)
    REPOR(46).COMPARANTES(1) = 47
    
    REPOR(47).ROTULO = "Almacena Agua Consumo: No"
    REPOR(47).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Almacena_agua= False AND Desabitada=FALSE"
    REPOR(47).Numcomparantes = 1
    ReDim REPOR(47).COMPARANTES(1 To 1)
    REPOR(47).COMPARANTES(1) = 46
    
    REPOR(48).ROTULO = "Hierve Agua: Siempre"
    REPOR(48).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Hierve_agua= 1 AND Desabitada=FALSE"
    REPOR(48).Numcomparantes = 3
    ReDim REPOR(48).COMPARANTES(1 To 3)
    REPOR(48).COMPARANTES(1) = 49
    REPOR(48).COMPARANTES(2) = 50
    REPOR(48).COMPARANTES(3) = 51
    
    REPOR(49).ROTULO = "Hierve Agua: Alguna veces"
    REPOR(49).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Hierve_agua= 2 AND Desabitada=FALSE"
    REPOR(49).Numcomparantes = 3
    ReDim REPOR(49).COMPARANTES(1 To 3)
    REPOR(49).COMPARANTES(1) = 48
    REPOR(49).COMPARANTES(2) = 50
    REPOR(49).COMPARANTES(3) = 51
    
    REPOR(50).ROTULO = "Hierve Agua: Nunca"
    REPOR(50).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Hierve_agua= 3 AND Desabitada=FALSE"
    REPOR(50).Numcomparantes = 3
    ReDim REPOR(50).COMPARANTES(1 To 3)
    REPOR(50).COMPARANTES(1) = 48
    REPOR(50).COMPARANTES(2) = 49
    REPOR(50).COMPARANTES(3) = 51
    
    REPOR(51).ROTULO = "Hierve Agua: Solo par los niños"
    REPOR(51).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Hierve_agua= 4 AND Desabitada=FALSE"
    REPOR(51).Numcomparantes = 3
    ReDim REPOR(51).COMPARANTES(1 To 3)
    REPOR(51).COMPARANTES(1) = 48
    REPOR(51).COMPARANTES(2) = 49
    REPOR(51).COMPARANTES(3) = 50
    
    REPOR(52).ROTULO = "Grifos, tuberias o inodoros goteando: Sí"
    REPOR(52).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Gotea_llaves_grifos= true AND Desabitada=FALSE"
    REPOR(52).Numcomparantes = 1
    ReDim REPOR(52).COMPARANTES(1 To 1)
    REPOR(52).COMPARANTES(1) = 53
    
    REPOR(53).ROTULO = "Grifos, tuberias o inodoros goteando: No"
    REPOR(53).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Gotea_llaves_grifos= False AND Desabitada=FALSE"
    REPOR(53).Numcomparantes = 1
    ReDim REPOR(53).COMPARANTES(1 To 1)
    REPOR(53).COMPARANTES(1) = 52
    
    REPOR(54).ROTULO = "Inodoro con conexión al alcantarillado"
    REPOR(54).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_servicio_sanitario= 1 AND Desabitada=FALSE"
    REPOR(54).Numcomparantes = 3
    ReDim REPOR(54).COMPARANTES(1 To 3)
    REPOR(54).COMPARANTES(1) = 55
    REPOR(54).COMPARANTES(2) = 56
    REPOR(54).COMPARANTES(3) = 57
    
    REPOR(55).ROTULO = "Inodoro o taza con tanque séptico"
    REPOR(55).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_servicio_sanitario= 2 AND Desabitada=FALSE"
    REPOR(55).Numcomparantes = 3
    ReDim REPOR(55).COMPARANTES(1 To 3)
    REPOR(55).COMPARANTES(1) = 54
    REPOR(55).COMPARANTES(2) = 56
    REPOR(55).COMPARANTES(3) = 57
    
    REPOR(56).ROTULO = "Inodoro sin conexiòn al alcantarillado"
    REPOR(56).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_servicio_sanitario= 3 AND Desabitada=FALSE"
    REPOR(56).Numcomparantes = 3
    ReDim REPOR(56).COMPARANTES(1 To 3)
    REPOR(56).COMPARANTES(1) = 54
    REPOR(56).COMPARANTES(2) = 55
    REPOR(56).COMPARANTES(3) = 57
    
    REPOR(57).ROTULO = "Ninguno"
    REPOR(57).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Tipo_servicio_sanitario= 4 AND Desabitada=FALSE"
    REPOR(57).Numcomparantes = 3
    ReDim REPOR(57).COMPARANTES(1 To 3)
    REPOR(57).COMPARANTES(1) = 54
    REPOR(57).COMPARANTES(2) = 55
    REPOR(57).COMPARANTES(3) = 56
    
    REPOR(58).ROTULO = "Problemas instalación sanitaria: Sí"
    REPOR(58).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Problemas_instalacion= True AND Desabitada=FALSE"
    REPOR(58).Numcomparantes = 1
    ReDim REPOR(58).COMPARANTES(1 To 1)
    REPOR(58).COMPARANTES(1) = 59
    
    REPOR(59).ROTULO = "Problemas instalación sanitaria: No"
    REPOR(59).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Problemas_instalacion= False AND Desabitada=FALSE"
    REPOR(59).Numcomparantes = 1
    ReDim REPOR(59).COMPARANTES(1 To 1)
    REPOR(59).COMPARANTES(1) = 58
    
    REPOR(60).ROTULO = "Inodoro/taza/letrina limpio: Sí"
    REPOR(60).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Inodoro_limpio= True AND Desabitada=FALSE"
    REPOR(60).Numcomparantes = 1
    ReDim REPOR(60).COMPARANTES(1 To 1)
    REPOR(60).COMPARANTES(1) = 61
    
    REPOR(61).ROTULO = "Inodoro/taza/letrina limpio: No"
    REPOR(61).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Inodoro_limpio= False AND Desabitada=FALSE"
    REPOR(61).Numcomparantes = 1
    ReDim REPOR(61).COMPARANTES(1 To 1)
    REPOR(61).COMPARANTES(1) = 60
    
    REPOR(62).ROTULO = "Estado caseta Instalación sanitaria: Bueno"
    REPOR(62).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_caseta= 1 AND Desabitada=FALSE"
    REPOR(62).Numcomparantes = 2
    ReDim REPOR(62).COMPARANTES(1 To 2)
    REPOR(62).COMPARANTES(1) = 63
    REPOR(62).COMPARANTES(2) = 64
       
    REPOR(63).ROTULO = "Estado caseta Instalación sanitaria: Malo"
    REPOR(63).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_caseta= 2 AND Desabitada=FALSE"
    REPOR(63).Numcomparantes = 2
    ReDim REPOR(63).COMPARANTES(1 To 2)
    REPOR(63).COMPARANTES(1) = 62
    REPOR(63).COMPARANTES(2) = 64
       
    REPOR(64).ROTULO = "Estado caseta Instalación sanitaria: No existe"
    REPOR(64).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Estado_caseta= 3 AND Desabitada=FALSE"
    REPOR(64).Numcomparantes = 2
    ReDim REPOR(64).COMPARANTES(1 To 2)
    REPOR(64).COMPARANTES(1) = 62
    REPOR(64).COMPARANTES(2) = 63
    
    REPOR(65).ROTULO = "Taponamiento conexión alcantarillado: Sí"
    REPOR(65).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Taponada_conexion= true AND Desabitada=FALSE"
    REPOR(65).Numcomparantes = 1
    ReDim REPOR(65).COMPARANTES(1 To 1)
    REPOR(65).COMPARANTES(1) = 66
       
    REPOR(66).ROTULO = "Taponamiento conexión alcantarillado: No"
    REPOR(66).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Taponada_conexion= False AND Desabitada=FALSE"
    REPOR(66).Numcomparantes = 1
    ReDim REPOR(66).COMPARANTES(1 To 1)
    REPOR(66).COMPARANTES(1) = 65
       
    REPOR(67).ROTULO = "Solución Preblema Taponamiento: Usuario Mismo"
    REPOR(67).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Solucion_problema= 1 AND Desabitada=FALSE"
    REPOR(67).Numcomparantes = 2
    ReDim REPOR(67).COMPARANTES(1 To 2)
    REPOR(67).COMPARANTES(1) = 68
    REPOR(67).COMPARANTES(2) = 69
       
    REPOR(68).ROTULO = "Solución Preblema Taponamiento: Operador"
    REPOR(68).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Solucion_problema= 2 AND Desabitada=FALSE"
    REPOR(68).Numcomparantes = 2
    ReDim REPOR(68).COMPARANTES(1 To 2)
    REPOR(68).COMPARANTES(1) = 67
    REPOR(68).COMPARANTES(2) = 69
       
    REPOR(69).ROTULO = "Solución Preblema Taponamiento: Otro"
    REPOR(69).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Solucion_problema= 3 AND Desabitada=FALSE"
    REPOR(69).Numcomparantes = 2
    ReDim REPOR(69).COMPARANTES(1 To 2)
    REPOR(69).COMPARANTES(1) = 67
    REPOR(69).COMPARANTES(2) = 68
    
    REPOR(70).ROTULO = "Destino basuras: Las quema"
    REPOR(70).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Destino_basuras= 1 AND Desabitada=FALSE"
    REPOR(70).Numcomparantes = 3
    ReDim REPOR(70).COMPARANTES(1 To 3)
    REPOR(70).COMPARANTES(1) = 71
    REPOR(70).COMPARANTES(2) = 72
    REPOR(70).COMPARANTES(3) = 73
    
    REPOR(71).ROTULO = "Destino basuras: La arroja"
    REPOR(71).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Destino_basuras= 2 AND Desabitada=FALSE"
    REPOR(71).Numcomparantes = 3
    ReDim REPOR(71).COMPARANTES(1 To 3)
    REPOR(71).COMPARANTES(1) = 70
    REPOR(71).COMPARANTES(2) = 72
    REPOR(71).COMPARANTES(3) = 73
    
    REPOR(72).ROTULO = "Destino basuras: Carro recolector"
    REPOR(72).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Destino_basuras= 3 AND Desabitada=FALSE"
    REPOR(72).Numcomparantes = 3
    ReDim REPOR(72).COMPARANTES(1 To 3)
    REPOR(72).COMPARANTES(1) = 70
    REPOR(72).COMPARANTES(2) = 71
    REPOR(72).COMPARANTES(3) = 73
    
    REPOR(73).ROTULO = "Destino basuras: La entierra"
    REPOR(73).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Destino_basuras= 4 AND Desabitada=FALSE"
    REPOR(73).Numcomparantes = 3
    ReDim REPOR(73).COMPARANTES(1 To 3)
    REPOR(73).COMPARANTES(1) = 70
    REPOR(73).COMPARANTES(2) = 71
    REPOR(73).COMPARANTES(3) = 72
    
    REPOR(74).ROTULO = "Basuras interior casa: Sí"
    REPOR(74).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Existencia_basuras_casa= True AND Desabitada=FALSE"
    REPOR(74).Numcomparantes = 1
    ReDim REPOR(74).COMPARANTES(1 To 1)
    REPOR(74).COMPARANTES(1) = 75
    
    REPOR(75).ROTULO = "Basuras interior casa: No"
    REPOR(75).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Existencia_basuras_casa= False AND Desabitada=FALSE"
    REPOR(75).Numcomparantes = 1
    ReDim REPOR(75).COMPARANTES(1 To 1)
    REPOR(75).COMPARANTES(1) = 74
    
    REPOR(76).ROTULO = "Oponion entidad/servicios: Buena"
    REPOR(76).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Opinion_administracion= 1 AND Desabitada=FALSE"
    REPOR(76).Numcomparantes = 2
    ReDim REPOR(76).COMPARANTES(1 To 3)
    REPOR(76).COMPARANTES(1) = 77
    REPOR(76).COMPARANTES(2) = 78
    REPOR(76).COMPARANTES(3) = 79
    
    REPOR(77).ROTULO = "Oponion entidad/servicios: Regular"
    REPOR(77).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Opinion_administracion= 2 AND Desabitada=FALSE"
    REPOR(77).Numcomparantes = 2
    ReDim REPOR(77).COMPARANTES(1 To 3)
    REPOR(77).COMPARANTES(1) = 76
    REPOR(77).COMPARANTES(2) = 78
    REPOR(77).COMPARANTES(3) = 79
    
    REPOR(78).ROTULO = "Oponion entidad/servicios: Mala"
    REPOR(78).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Opinion_administracion= 3 AND Desabitada=FALSE"
    REPOR(78).Numcomparantes = 2
    ReDim REPOR(78).COMPARANTES(1 To 3)
    REPOR(78).COMPARANTES(1) = 76
    REPOR(78).COMPARANTES(2) = 77
    REPOR(78).COMPARANTES(3) = 79
    
    REPOR(79).ROTULO = "Oponion entidad/servicios: No opina"
    REPOR(79).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Opinion_administracion= 0 AND Desabitada=FALSE"
    REPOR(79).Numcomparantes = 3
    ReDim REPOR(79).COMPARANTES(1 To 3)
    REPOR(79).COMPARANTES(1) = 76
    REPOR(79).COMPARANTES(2) = 77
    REPOR(79).COMPARANTES(3) = 78
    
    REPOR(80).ROTULO = "Respaldo Entidad: Sí"
    REPOR(80).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Respaldo_entidad= True AND Desabitada=FALSE"
    REPOR(80).Numcomparantes = 1
    ReDim REPOR(80).COMPARANTES(1 To 1)
    REPOR(80).COMPARANTES(1) = 81
    
    REPOR(81).ROTULO = "Respaldo Entidad: No"
    REPOR(81).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Respaldo_entidad= False AND Desabitada=FALSE"
    REPOR(81).Numcomparantes = 1
    ReDim REPOR(81).COMPARANTES(1 To 1)
    REPOR(81).COMPARANTES(1) = 80
    
    REPOR(82).ROTULO = "Deshabitadas: Sí"
    REPOR(82).SQL = "SELECT COUNT(Ruta) AS CANTI FROM TABLA1 WHERE Desabitada= True"
    REPOR(82).Numcomparantes = 1
    ReDim REPOR(82).COMPARANTES(1 To 1)
    REPOR(82).COMPARANTES(1) = 0
    
    REPOR(83).ROTULO = "Incluidas en la Encuesta y no pertenencientes al Sistema"
    REPOR(83).SQL = "SELECT COUNT (RUTA) as canti FROM TABLA1 WHERE RUTA NOT IN( SELECT RUTA FROM TABLA2)"
    
    REPOR(84).ROTULO = "Incluidas en el sistema pero no en la encuesta"
    REPOR(84).SQL = "SELECT COUNT (RUTA)  as canti FROM TABLA2 WHERE RUTA NOT IN( SELECT RUTA FROM TABLA1)"
    
    filas = 82
       
    For j = 0 To filas
        base.RecordSource = REPOR(j).SQL
        base.Refresh
        REPOR(j).cantidad = base.Recordset!CANTI
        REPOR(j).PORCEN = REPOR(j).cantidad * 100 / REPOR(0).cantidad
    Next j
    
    
End Sub


Public Sub cargar_resultados(REJILLA As MSFlexGrid, REPOR() As reportes)
    REJILLA.Rows = filas + 2
    REJILLA.ColWidth(0) = 400
    For Y = 0 To filas
        REJILLA.TextMatrix(Y + 1, 0) = Y + 1
        REJILLA.TextMatrix(Y + 1, 1) = REPOR(Y).ROTULO
        REJILLA.TextMatrix(Y + 1, 2) = REPOR(Y).cantidad
        
        For X = 1 To 4
            REJILLA.TextMatrix(Y + 1, 3) = REJILLA.TextMatrix(Y + 1, 3) + Mid(REPOR(Y).PORCEN, X, 1)
        Next X
        If REPOR(Y).PORCEN > 0.1 And REPOR(Y).PORCEN <> 0 Then
            REJILLA.TextMatrix(Y + 1, 3) = REJILLA.TextMatrix(Y + 1, 3) + " %"
        ElseIf REPOR(Y).PORCEN > 0.005 Then
            REJILLA.TextMatrix(Y + 1, 3) = REJILLA.TextMatrix(Y + 1, 3) + " E -02" + " %"
        Else
            REJILLA.TextMatrix(Y + 1, 3) = "0 %"
        End If
    Next Y
    REJILLA.TextMatrix(0, 1) = "ITEM"
    REJILLA.TextMatrix(0, 2) = "CANTIDAD"
    REJILLA.TextMatrix(0, 3) = "Porcentaje"
'    REJILLA.SetFocus
End Sub
Public Sub CARGAR_CONSULTA(CONSUL As String)
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
               " From Tabla1 "
    X = 1
        While Mid(CONSUL, X, 1) <> "W"
            X = X + 1
        Wend
        Consulta = Consulta + Mid(CONSUL, X, Len(CONSUL) - X + 1)
        Consulta = Consulta + " ORDER BY  RUTA, DIRECCION_PREDIO"
End Sub

Public Sub Imprimir_Resumen(PLANILLA As MSFlexGrid, TITULO As String)
Dim ESPACIO As Integer
Printer.Font = "Draft 10cpi"
'Printer.FontSize = 8
Printer.FontBold = True
Printer.PaperSize = vbPRPSLegal
ESPACIO = (82 - Len(TITULO)) / 2
Printer.Print String(ESPACIO, " ");
Printer.Print TITULO
Printer.Print
Printer.FontBold = False
For Y = 0 To PLANILLA.Rows - 1
    If Y = 0 Then
        Printer.FontBold = True
    Else
        Printer.FontBold = False
    End If
    For X = 1 To PLANILLA.Cols - 1
        Select Case X
            Case 1: Printer.Print String(5, " ");
            Case 2:
                    If 45 - Len(PLANILLA.TextMatrix(Y, X - 1)) > 0 Then
                        Printer.Print String(45 - Len(PLANILLA.TextMatrix(Y, X - 1)), " ");
                    End If
            Case 3: Printer.Print String(15 - Len(PLANILLA.TextMatrix(Y, X - 1)), " ");
        End Select
        Printer.Print PLANILLA.TextMatrix(Y, X);
    Next X
    If Y = 0 Then Printer.Print vbCrLf
    Printer.Print
Next Y

Printer.EndDoc
End Sub
Public Sub calculo_maximo_columnas(REJILLA As MSFlexGrid)
    Dim MAXIMO As Integer
    For Y = 0 To REJILLA.Cols - 1
        If M.ACTIVO(Y) = True Then
            MAXIMO = Len(REJILLA.TextMatrix(0, Y))
            For X = 0 To REJILLA.Rows - 1
                If Len(REJILLA.TextMatrix(X, Y)) > MAXIMO Then
                    MAXIMO = Len(REJILLA.TextMatrix(X, Y))
                End If
            Next X
            M.tamcarac(Y) = MAXIMO + 1
        End If
    Next Y
End Sub
Public Sub IMPRIMIR_REPORTE(REJILLA As MSFlexGrid, TITULO As String)

Dim ESPACIO As Integer
Dim POSIM As Integer
Dim POS As Integer
Printer.Font = "Draft 10cpi"
'Printer.FontSize = 8
Printer.PaperSize = vbPRPSLegal
IMPRIMIR_ENCABEZADO TITULO, REJILLA
Printer.FontBold = False

For X = 1 To REJILLA.Rows - 1
    If X Mod 41 <> 0 Then
        Printer.Print String(5, " ")
        POSIM = 1
        For Y = 0 To REJILLA.Cols - 1
            If M.ACTIVO(Y) = True Then
                If POSIM = 1 Then
                    Printer.Print REJILLA.TextMatrix(X, Y);
                    POSIM = 2
                Else
                    POS = Y - 1
                    While M.ACTIVO(POS) = False And POS > 0
                        POS = POS - 1
                    Wend
                    If M.tamcarac(POS) > Len(REJILLA.TextMatrix(X, POS)) Then
                        Printer.Print String(M.tamcarac(POS) - Len(REJILLA.TextMatrix(X, POS)), " ");
                    End If
                    Printer.Print REJILLA.TextMatrix(X, Y);
                End If
            End If
        Next Y
        Printer.Print
    Else
        Printer.NewPage
        IMPRIMIR_ENCABEZADO TITULO, REJILLA
        Printer.FontBold = False
        Printer.Print String(5, " ")
        POSIM = 1
        For Y = 0 To REJILLA.Cols - 1
            If M.ACTIVO(Y) = True Then
                If POSIM = 1 Then
                    Printer.Print REJILLA.TextMatrix(X, Y);
                    POSIM = 2
                Else
                    POS = Y - 1
                    While M.ACTIVO(POS) = False And POS > 0
                        POS = POS - 1
                    Wend
                    If M.tamcarac(POS) > Len(REJILLA.TextMatrix(X, POS)) Then
                        Printer.Print String(M.tamcarac(POS) - Len(REJILLA.TextMatrix(X, POS)), " ");
                    End If
                    Printer.Print REJILLA.TextMatrix(X, Y);
                End If
            End If
        Next Y
        Printer.Print
    End If
Next X

Printer.EndDoc
End Sub
Public Sub IMPRIMIR_ENCABEZADO(TITULO As String, REJILLA As MSFlexGrid)
Dim T As Integer
Dim R As Integer
Dim POS As Integer
Printer.Font = "Draft 10cpi"
Printer.FontSize = 10
Printer.FontBold = True
ESPACIO = (82 - Len(TITULO)) / 2
Printer.Print String(ESPACIO, " ");
Printer.Print TITULO
Printer.Print
For T = 0 To 0
    
    Printer.Print String(5, " ")
    POSIM = 1
    For R = 0 To REJILLA.Cols - 1
        If M.ACTIVO(R) = True Then
            If POSIM = 1 Then
                Printer.Print REJILLA.TextMatrix(T, R);
                POSIM = 2
            Else
                POS = R - 1
                While M.ACTIVO(POS) = False And POS > 0
                    POS = POS - 1
                Wend
                If M.tamcarac(POS) > Len(REJILLA.TextMatrix(T, POS)) Then
                    Printer.Print String(M.tamcarac(POS) - Len(REJILLA.TextMatrix(T, POS)), " ");
                End If
                Printer.Print REJILLA.TextMatrix(T, R);
            End If
        End If
    Next R
    Printer.Print
Next T
End Sub
