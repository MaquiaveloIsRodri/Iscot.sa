' Diaria - Carga EP de tabla Curiculum
Sub Main
    Stop
	Set self = omensajeps.value
    Set xfso = CreateObject("Scripting.FileSystemObject")
    archivoLog = "c:\util\"&nombreusuario()&"-ImportarCVGoogle.txt"
	Carteles = (Self.ClassId = "{20773538-B07F-11D2-9229-0000214166F2}")
	
'    If xfso.FileExists(archivoLog) Then
'        xfso.DeleteFile(archivoLog)
'    End If


    ' Set request = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	' request.open "GET", "https://docs.google.com/spreadsheets/d/e/2PACX-1vSItIwVkeLeRM1nB_EkbkObQOpMrFVYp6ci4njqqrvvFnn3KJMoFqEc5HeJ_ts7u3pRBm7kPOwzNzbU/pub?output=tsv", False 
	' 'request.setRequestHeader "Content-type", "application/x-www-form-urlencoded; charset=UTF-8"
	' request.send 
	' texto = request.responseText

    ' 'texto = URLGet("https://docs.google.com/spreadsheets/d/e/2PACX-1vSItIwVkeLeRM1nB_EkbkObQOpMrFVYp6ci4njqqrvvFnn3KJMoFqEc5HeJ_ts7u3pRBm7kPOwzNzbU/pub","output=tsv","","")
	' cvs = Split(texto,chr(13))   

    Set xCone2 = CreateObject("ADODB.CONNECTION")
    xCone2.connectionString = StringConexion("CALIPSO", Self.WorkSpace)
    xCone2.connectiontimeOut = 150
    xCone2.Open
    set query = RecordSet(xCone2, "select top 1 * from producto")
    query.close
    query.activeconnection.commandtimeout=0
    query.source="select top 1 * from producto" ' en q2 esta la consulta
    query.open
    query.close
    query.activeconnection.commandtimeout=0
		 
    query.source= "SELECT * FROM VP_CURRICULUM"
    query.open
    If query.eof Then
        If Carteles Then Call MsgBox("No existen nuevos empleados para importar.",64,"Información")
        Exit Sub
    End If   
    
    Set oSistema		= ExisteBo(Self,"appcalipso"      ,"id","BE582E8B-30BE-4F7B-9032-ABF58AE718BF",nil,True,False,"=")
    Set xcompania		= ExisteBo(Self,"compania"        ,"id","63E432B5-726E-4195-8F9D-2025C20D7467",nil,True,False,"=")
	Set oUORRHH 		= InstanciarBO("{DD7EE41C-3171-4D8B-9463-76F0F84A2B26}", "UORECURSOSHUMANOS", Self.Workspace)
    Set oTiposNomina	= ExisteBo(Self,"TipoClasificador","NOMBRE","Tipo de Pago",nil,True,False,"=")
    Set oProvincias		= ExisteBo(Self,"TipoClasificador","NOMBRE","EP - Provincias",nil,True,False,"=")
    Set oCiudades		= ExisteBo(Self,"TipoClasificador","NOMBRE","EP - Ciudades",nil,True,False,"=")
    Set oTiposJornada	= ExisteBo(Self,"TipoClasificador","NOMBRE","Tipo de Jornada",nil,True,False,"=")
    Set oNivelesPri		= ExisteBo(Self,"TipoClasificador","NOMBRE","EP - Educacion Primaria",nil,True,False,"=")
    Set oNivelesSec		= ExisteBo(Self,"TipoClasificador","NOMBRE","EP - Educacion Secundaria",nil,True,False,"=")
    Set oEstadosOP		= ExisteBo(Self,"TipoClasificador","NOMBRE","EP - Estado",nil,True,False,"=")
    Set oConceptosAfip  = ExisteBo(Self, "TIPOCLASIFICADOR", "ID", "939F99DC-E5D3-4545-89CF-CFF9822D467B", nil, True, false, "=")
    Set oEstado         = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "ID", "{D554A05B-CE21-418F-B2A1-EDEE3A2E7D9E}", nil, True, false, "=") ' IMPORTACION WEB
    Set oCondicion      = ExisteBo(Self, "CONDICION", "ID", "{817AC20C-0BA7-4D59-9B1E-E9D3929A013C}", nil, True, false, "=") ' ACTIVO ? 147D8D14-617B-44DE-A099-A658AF25603E
    contador = 0    : contA = 0 :    contN = 0 
	'If Weekday(xFechaHasta) = 2 Then xFechaDesde = DateAdd("d", -3,xFechaHasta) Else xFechaDesde = DateAdd("d", -1,xFechaHasta) End If
	'xFechaHasta = DateAdd("d", -1,xFechaHasta)
	If Carteles Then Call ProgressControl(Self.Workspace, "Importando CVs de Google", 0, query.properties.count)

    Do While Not query.eof
		If Carteles Then Call ProgressControlAvance(Self.Workspace, fechaCarga)
        mensaje = ""
        xError                  = False : a = 1 
        Set oEmpleadoPotencial  = Nothing
        Set oEmpleado           = Nothing
        set oSexo               = Nothing
        set oProvincia          = Nothing
        set oLocalidad          = Nothing
        set oBarrio	            = Nothing
        existeEP                = False
        Set oSecundario         = Nothing
        xFechaIngreso			= #30/12/1899#	
        Set oDisponibilidad     = Nothing
        Set oMovilidadPropia     = Nothing

        If Not IsNull(query("ID").value) Then                   xID				        = Replace(Replace(query("ID").value,"'",""),chr(34),"") Else xID = query("ID").value  End If
        If Not IsNull(query("APELLIDO").value) Then             xApellido				= Replace(Replace(query("APELLIDO").value,"'",""),chr(34),"") Else xApellido = query("APELLIDO").value  End If
        If Not IsNull(query("NOMBRE").value) Then               xNombre					= Replace(Replace(query("NOMBRE").value,"'",""),chr(34),"") Else xNombre = query("NOMBRE").value  End If
        If Not IsNull(query("CUIT").value) Then                 xCuil					= Replace(Replace(query("CUIT").value,"'",""),chr(34),"") Else xCuil = query("CUIT").value  End If
        If Not IsNull(query("NUMERODOCUMENTO").value) Then      xNroDocumento			= Replace(Replace(query("NUMERODOCUMENTO").value,"'",""),chr(34),"") Else xNroDocumento = query("NUMERODOCUMENTO").value  End If
        If Not IsNull(query("FECHACARGA").value) Then           xFechaCarga             = Replace(Replace(query("FECHACARGA").value ,"'",""),chr(34),"") Else xFechaCarga = query("FECHACARGA").value  End If
        If Not IsNull(query("SEXO").value) Then                 xSexo					= Replace(Replace(query("SEXO").value,"'",""),chr(34),"") Else xSexo = query("SEXO").value  End If
        If Not IsNull(query("PROVINCIA").value) Then            xProvincia				= Replace(Replace(query("PROVINCIA").value,"'",""),chr(34),"") Else xProvincia = query("PROVINCIA").value  End If
        If Not IsNull(query("LOCALIDAD").value) Then            xLocalidad				= Replace(Replace(query("LOCALIDAD").value,"'",""),chr(34),"") Else xLocalidad = query("LOCALIDAD").value  End If
        If Not IsNull(query("FECHANACIMIENTO").value) Then      xFechaNacimiento		= Replace(Replace(query("FECHANACIMIENTO").value,"'",""),chr(34),"") Else xFechaNacimiento = query("FECHANACIMIENTO").value  End If
        If Not IsNull(query("CALLE").value) Then                xCalle					= Replace(Replace(query("CALLE").value,"'",""),chr(34),"") Else xCalle = query("CALLE").value  End If
        If Not IsNull(query("NUMERO").value) Then               xNumero					= Replace(Replace(query("NUMERO").value,"'",""),chr(34),"") Else xNumero = query("NUMERO").value  End If
        If Not IsNull(query("BARRIO").value) Then               xBarrio					= Replace(Replace(query("BARRIO").value,"'",""),chr(34),"") Else xBarrio = query("BARRIO").value  End If
        If Not IsNull(query("TELEFONO").value) Then             xTelefonoParticular		= Replace(Replace(query("TELEFONO").value,"'",""),chr(34),"") Else xTelefonoParticular = query("TELEFONO").value  End If
        If Not IsNull(query("TELEFONOALTERNATIVO").value) Then  xTelefonoAlternativo    = Replace(Replace(query("TELEFONOALTERNATIVO").value,"'",""),chr(34),"") Else xTelefonoAlternativo = query("TELEFONOALTERNATIVO").value  End If
        If Not IsNull(query("SECUNDARIO").value) Then           xSecundario             = Replace(Replace(query("SECUNDARIO").value ,"'",""),chr(34),"") Else xSecundario = query("SECUNDARIO").value  End If
        If Not IsNull(query("LICENCIACONDUCIR").value) Then     xLicenciaConducir       = Replace(Replace(query("LICENCIACONDUCIR").value ,"'",""),chr(34),"") Else xLicenciaConducir = query("LICENCIACONDUCIR").value  End If
        If Not IsNull(query("MOVILIDADPROPIA").value) Then      xMovilidadPropia        = Replace(Replace(query("MOVILIDADPROPIA").value ,"'",""),chr(34),"") Else xMovilidadPropia = query("MOVILIDADPROPIA").value  End If
        If Not IsNull(query("TIPOMOVILIDAD").value) Then        xTipoMovilidad          = Replace(Replace(query("TIPOMOVILIDAD").value ,"'",""),chr(34),"") Else xTipoMovilidad = query("TIPOMOVILIDAD").value  End If
        If Not IsNull(query("ORIGEN").value) Then               xReferencia             = Replace(Replace(query("ORIGEN").value ,"'",""),chr(34),"") Else xReferencia = query("ORIGEN").value  End If
        If Not IsNull(query("EMAIL").value) Then                xEmail                  = Replace(Replace(query("EMAIL").value ,"'",""),chr(34),"") Else xEmail = query("EMAIL").value  End If
        If Not IsNull(query("CV").value) Then                   linkCv                  = Replace(Replace(query("CV").value ,"'",""),chr(34),"") Else linkCv = query("CV").value  End If
        If Not IsNull(query("EXPERIENCIA").value) Then          xExperiencia            = Replace(Replace(query("EXPERIENCIA").value ,"'",""),chr(34),"") Else xExperiencia = query("EXPERIENCIA").value  End If
        If Not IsNull(query("EMPLEADOPOTENCIAL_ID").value) Then xEP_ID                  = Replace(Replace(query("EMPLEADOPOTENCIAL_ID").value ,"'",""),chr(34),"") Else xEP_ID = query("EMPLEADOPOTENCIAL_ID").value  End If
        If Not IsNull(query("DETALLE").value) Then              xDetalle                = Replace(Replace(query("DETALLE").value ,"'",""),chr(34),"") Else xDetalle = query("DETALLE").value  End If
        If Not IsNull(query("DISPONIBILIDAD").value) Then       xDisponibilidad         = Replace(Replace(query("DISPONIBILIDAD").value ,"'",""),chr(34),"") Else xDisponibilidad = query("DISPONIBILIDAD").value  End If
        If Not IsNull(query("LIMPIEZAGENERAL").value) Then      xExperienciaLimp        = Replace(Replace(query("LIMPIEZAGENERAL").value,"'",""),chr(34),"") Else xExperienciaLimp = query("LIMPIEZAGENERAL").value  End If
        If Not IsNull(query("LIMPIEZAHOSPITALES").value) Then   xExperienciaHosp        = Replace(Replace(query("LIMPIEZAHOSPITALES").value,"'",""),chr(34),"") Else xExperienciaHosp = query("LIMPIEZAHOSPITALES").value  End If
        If Not IsNull(query("ESPACIOSVERDES").value) Then       xExperienciaEV          = Replace(Replace(query("ESPACIOSVERDES").value,"'",""),chr(34),"") Else xExperienciaEV = query("ESPACIOSVERDES").value  End If
        If Not IsNull(query("ELEVADORES").value) Then           xExperienciaMula        = Replace(Replace(query("ELEVADORES").value,"'",""),chr(34),"") Else xExperienciaMula = query("ELEVADORES").value  End If
        If Not IsNull(query("MANIANA").value) Then              xManiana                = Replace(Replace(query("MANIANA").value ,"'",""),chr(34),"") Else xManiana = query("MANIANA").value  End If
        If Not IsNull(query("TARDE").value) Then                xTarde                  = Replace(Replace(query("TARDE").value ,"'",""),chr(34),"") Else xTarde = query("TARDE").value  End If
        If Not IsNull(query("NOCHE").value) Then                xNoche                  = Replace(Replace(query("NOCHE").value ,"'",""),chr(34),"") Else xNoche = query("NOCHE").value  End If
        If Not IsNull(query("POSTULACION").value) Then          xPostulacion            = Replace(Replace(query("POSTULACION").value,"'",""),chr(34),"") Else xPostulacion = query("POSTULACION").value  End If
        If Not IsNull(query("LAT").value) Then                  xLat                    = Replace(Replace(query("LAT").value ,"'",""),chr(34),"") Else xLat = query("LAT").value  End If
        If Not IsNull(query("LON").value) Then                  xLon                    = Replace(Replace(query("LON").value ,"'",""),chr(34),"") Else xLon = query("LON").value  End If

        contador = contador +1                 
        
        Set oEmpleado = ExisteBo(Self, "EMPLEADO", "CUIT", xCuil, nil, True, False, "=")
        If oEmpleado Is Nothing Then Set oEmpleado = ExisteBo(Self, "EMPLEADO", "NUMERODOCUMENTO", xNroDocumento, nil, True, False, "=")
        If Not oEmpleado Is Nothing Then 
           mensaje = mensaje + " " + "Existe Empleado; " : xError = True
        End If
        Set oEmpleadoPotencial = ExisteBo(Self, "EMPLEADOPOTENCIAL", "CUIT", xCuil, nil, True, False, "=")
        If oEmpleadoPotencial Is Nothing Then Set oEmpleadoPotencial = ExisteBo(Self, "EMPLEADOPOTENCIAL", "NUMERODOCUMENTO", xNroDocumento, nil, True, False, "=")
        If Not oEmpleadoPotencial Is Nothing Then
           existeEP = True
           If Not oEmpleadoPotencial.BOextension.Estado.ID = "{D554A05B-CE21-418F-B2A1-EDEE3A2E7D9E}" Then  'IMPORTACION WEB
               mensaje = mensaje + " " + "Estado EP no permite cambios; "
               xError = True
           End If
        End If


        If verificaCuit(xCuil) <> "CORRECTO" Then 
            mensaje = mensaje + " " + "Cuit; "
            xError = True
        End If

        If Not xSexo = "" And Not IsNull(xSexo) Then
            Set oSexo = ExisteBo(Self, "SEXO", "ID", xSexo, nil, True, False, "=")
        End If
        If oSexo Is Nothing Then 
            mensaje = mensaje + " " + "Sexo; " : xError = True
        End If
        If Not xProvincia = "" And Not IsNull(xProvincia) Then
            Set oProvincia = ExisteBo(Self, "PROVINCIA", "ID", xProvincia, nil, True, False, "=")
        End If
        If oProvincia Is Nothing Then
            mensaje = mensaje + " " + "Provincia; " : xError = True
        End If

        xLocalidad  = UCase(xLocalidad)
        xLocalidad  = Replace(xLocalidad, "Á", "A")
        xLocalidad 	= Replace(xLocalidad, "É", "E")
        xLocalidad  = Replace(xLocalidad, "Í", "I")
        xLocalidad  = Replace(xLocalidad, "Ó", "O")
        xLocalidad  = Replace(xLocalidad, "Ú", "U")
        xLocalidad  = Trim(xLocalidad)
        Set oLocalidad = ExisteBo(Self, "CIUDAD", "NOMBRE", xLocalidad, nil, True, False, "=")
        If oLocalidad Is Nothing Then
            mensaje = mensaje + " " + "Localidad; " 
            If Not oProvincia Is Nothing Then
                set oLocalidad=crearbo("CIUDAD",Self)
                oLocalidad.Nombre = xLocalidad
                oProvincia.Ciudades.add(oLocalidad)
            End If
        End If
        'If xLocalidad = "CORDOBA CAPITAL" Then Set oLocalidad = InstanciarBO("{ABDD5B39-D4BC-4949-B631-7AA3B4B2A49A}", "CIUDAD", self.Workspace )
        
        If IsDate(xFechaNacimiento) Then
            xFechaNacimiento = cDate(xFechaNacimiento)
        Else
            mensaje = mensaje + " " + "Fecha de Nacimiento; " : xError = True
        End If

        Select Case uCase(xSecundario)
            Case "COMPLETO"           : xSecundario = "COMPLETA"
            Case "INCOMPLETO"         : xSecundario = "INCOMPLETA"
        End Select
        Set oSecundario = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xSecundario, onivelesSec.Valores, True, False, "=")
        If oSecundario Is Nothing Then
            mensaje = mensaje + " " + "Secundario; "
        End If

        Select Case uCase(xDisponibilidad)
            Case true   : Set oDisponibilidad = InstanciarBO("{EDEDCA77-1543-4E05-B17F-263BCF91D1F3}", "ITEMTIPOCLASIFICADOR", self.Workspace )
            Case false  : Set oDisponibilidad = InstanciarBO("{EDEDCA77-1543-4E05-B17F-263BCF91D1F3}", "ITEMTIPOCLASIFICADOR", self.Workspace )
        End Select
        
        Select Case uCase(xMovilidadPropia)
            Case false  :   Set oMovilidadPropia = InstanciarBO("{A5F41E8C-1F8C-4835-95F3-75FC333A74F7}", "ITEMTIPOCLASIFICADOR", self.Workspace )
            Case true   :   Set oMovilidadPropia = InstanciarBO("{D0889701-A025-451E-9937-117082FC7FAB}", "ITEMTIPOCLASIFICADOR", self.Workspace )
        End Select
        
        If Not xError Then
        'Comentar lineas de código con bandera para actualizar
        'Bandera = True
            ' VALIDAR SI EXISTE PREVIAMENTE EL EMPLEADO POTENCIAL, ENTONCES, ES UNA ACTUALIZACION DE CV	
            If Not existeEP Then
                Set oEmpleadoPotencial = crearbo("EMPLEADOPOTENCIAL", oUORRHH)
                xCompania.EmpleadosPotenciales.Add(oEmpleadoPotencial)
                oEmpleadoPotencial.UnidadOperativa = oUORRHH
                    
                Set oUDEmpleadoPotencial = crearbo("UD_EMPLEADOPOTENCIAL", oEmpleadoPotencial)
                oEmpleadoPotencial.boextension                                      = oUDEmpleadoPotencial
                oUDEmpleadoPotencial.bo_owner                                       = oEmpleadoPotencial
                oEmpleadoPotencial.BoExtension.Nombre	                            = xNombre
                oEmpleadoPotencial.BoExtension.Apellido	                            = xApellido
                oEmpleadoPotencial.BoExtension.Cuil		                            = xCuil
                If Not oSexo Is Nothing Then oEmpleadoPotencial.BoExtension.Sexo    = oSexo
                contN = contN + 1
				Call registroSeleccion(oEmpleadoPotencial,"Creacion EP","Creacion EP",Date)
            Else
                contA = contA + 1
                'Bandera = False
            End If
            'If Bandera Then
            If Not oEstado Is Nothing Then oEmpleadoPotencial.BoExtension.Estado                        = oEstado
            If Not oCondicion Is Nothing Then oEmpleadoPotencial.BoExtension.Condicion                  = oCondicion
            If Not oProvincia Is Nothing Then oEmpleadoPotencial.BoExtension.Provincia                  = oProvincia
            If Not oLocalidad Is Nothing Then oEmpleadoPotencial.BoExtension.Localidad	                = oLocalidad
            oEmpleadoPotencial.BoExtension.FechaNacimiento	                                            = xFechaNacimiento
            oEmpleadoPotencial.BoExtension.Calle                                                        = xCalle
            oEmpleadoPotencial.BoExtension.Nro                                                          = xNumero
            oEmpleadoPotencial.BoExtension.Barrio                                                       = xBarrio
            oEmpleadoPotencial.BoExtension.TelefonoParticular                                           = xTelefonoParticular
            If Not oSecundario Is Nothing Then oEmpleadoPotencial.BoExtension.Secundario                = oSecundario
            If Not IsNull(xLicenciaConducir) Then oEmpleadoPotencial.BoExtension.CarnetConducir         = xLicenciaConducir
            oEmpleadoPotencial.BoExtension.Email                                                        = xEmail
            If Not oDisponibilidad Is Nothing Then oEmpleadoPotencial.BoExtension.DISPONIBILIDADHORARIA = oDisponibilidad
            If Not oMovilidadPropia Is Nothing Then oEmpleadoPotencial.BoExtension.MOVILIDADPROPIA      = oMovilidadPropia
            oEmpleadoPotencial.BoExtension.Movilidad                                                    = xTipoMovilidad
            oEmpleadoPotencial.BoExtension.TelefonoAlternativo                                          = xTelefonoAlternativo
            oEmpleadoPotencial.BoExtension.DetalleExperiencias                                          = xExperiencia
            If Not IsNull(xExperienciaLimp) Then oEmpleadoPotencial.BoExtension.EXPERIENCIALIMPIEZA     = xExperienciaLimp
            ' oEmpleadoPotencial.BoExtension.EXPERIENCIAMANTENIMIENTO                                     = experienciaMantenimiento
            If Not IsNull(xExperienciaEV) Then oEmpleadoPotencial.BoExtension.EXPERIENCIAESPACIOSVERDES = xExperienciaEV
            If Not IsNull(xExperienciaHosp) Then oEmpleadoPotencial.BoExtension.EXPERIENCIAHOSPITAL     = xExperienciaHosp
            If Not IsNull(xExperienciaMula) Then oEmpleadoPotencial.BoExtension.EXPERIENCIAMULERO       = xExperienciaMula 
            ' oEmpleadoPotencial.BoExtension.EXPERIENCIAOBRAS                                             = experienciaObras 
            If Not IsNull(xManiana) Then oEmpleadoPotencial.BoExtension.DISPONIBILIDADMANIANA           = xManiana 
            If Not IsNull(xTarde) Then oEmpleadoPotencial.BoExtension.DISPONIBILIDADTARDE               = xTarde 
            If Not IsNull(xNoche) Then oEmpleadoPotencial.BoExtension.DISPONIBILIDADNOCHE               = xNoche 
            oEmpleadoPotencial.BoExtension.OBSERVACIONES                                                = xPostulacion 
            oEmpleadoPotencial.BoExtension.REFERENCIABUSQUEDA                                           = xReferencia 
            oEmpleadoPotencial.BoExtension.LINKCV                                                       = linkCv
            oEmpleadoPotencial.BoExtension.AntiguedadVacaciones                                         = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHAANTIGUEDADRECONOCIDA                                    = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHAESTUDIO                                                 = #30/12/1899#
            oEmpleadoPotencial.BoExtension.INDUCCION                                                    = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHADISPONIBILIDAD                                          = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHAINGRESO                                                 = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHANACIMIENTOESPOSA                                        = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHANACIMIENTOHIJO1                                         = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHANACIMIENTOHIJO2                                         = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHANACIMIENTOHIJO3                                         = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHANACIMIENTOHIJO4                                         = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHAENTRAVISTA1                                             = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHAENTREVISTA2                                             = #30/12/1899#
            oEmpleadoPotencial.BoExtension.FECHAESTADO                                             	    = cDate(xFechaCarga)

            'call WorkSpaceCheck(oEmpleadoPotencial.Workspace)
            'End If
            a = workspacecheck(Self.workspace)
            If a = 0 Then
                'If mensaje <> "" Then
					Call registroSeleccion(oEmpleadoPotencial,"Postulacion","Postulacion",cDate(xFechaCarga))
				
                    set xVector = NewVector()
                    set xBucket = NewBucket()
                    'xBucket.Value = "UPDATE CURRICULUM SET EMPLEADOPOTENCIAL_ID = '"& oEmpleadoPotencial.id &"' WHERE ID = '" & xID & "'"
                    xBucket.Value = "UPDATE CURRICULUM SET EMPLEADOPOTENCIAL_ID = '"& oEmpleadoPotencial.id &"', DETALLE = '"& mensaje &"' WHERE ID = '" & xID & "'"
                    xVector.Add(xBucket)
                    ExecutarSQL xVector, "DistrObj", "", Self.Workspace, -1
                    call WorkSpaceCheck(Self.Workspace)
                ' Else
                '     set xVector = NewVector()
                '     set xBucket = NewBucket()
                '     xBucket.Value = "UPDATE CURRICULUM SET DETALLE = '"& mensaje &"' WHERE ID = '" & xID & "'"
                '     xVector.Add(xBucket)
                '     ExecutarSQL xVector, "DistrObj", "", Self.Workspace, -1
                '     call WorkSpaceCheck(Self.Workspace) 
                ' End If
            End If
        Else
            If existeEP Then 
                q = "UPDATE CURRICULUM SET EMPLEADOPOTENCIAL_ID = '"& oEmpleadoPotencial.id &"', DETALLE = '"& mensaje &"' WHERE ID = '" & xID & "'"
            Else
                q = "UPDATE CURRICULUM SET DETALLE = '"& mensaje &"' WHERE ID = '" & xID & "'"
            End If 
            set xVector = NewVector()
            set xBucket = NewBucket()
            xBucket.Value = q
            xVector.Add(xBucket)
            ExecutarSQL xVector, "DistrObj", "", Self.Workspace, -1
            call WorkSpaceCheck(Self.Workspace) 
        End If
        query.MoveNext
    loop
	If Carteles Then Call ProgressControlFinish(Self.Workspace)
	'Enviar correo

   ' set xFSO = CreateObject("Scripting.FileSystemObject")
    set xArchivo = xfso.OpenTextFile("C:\util\html\email_cvs.html")
    
    htmlBody = xArchivo.readAll()
    htmlBody = replace(htmlBody,"xCargados" , contador)
    htmlBody = replace(htmlBody,"xDesechados" , contador -contA - contN)
    htmlBody = replace(htmlBody,"xActualizados" , contA) 
    htmlBody = replace(htmlBody,"xCreados" , contN) 
    xSubject = "[CVs actualizados]" 
    xcorreos = "rodrigocasomolina@iscot.com.ar;pablosans@iscot.com.ar;santiagogatsch@iscot.com.ar"
    xBody    = ""            
    xadjunto = ""            
    call enviar_aviso_sinmsg(Self, xcorreos, xSubject, xBody, xadjunto,htmlBody)  
end sub
