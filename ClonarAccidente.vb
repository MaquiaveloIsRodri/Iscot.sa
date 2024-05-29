'2024-05-23 Clonar Accidente - Rodrigo Fierro
sub main
    stop
    'Creo un nuevo elemento en la UD
    set xAccidente = CrearBO("UD_ACCIDENTES", self.WorkSpace)

    set xEmpleado = Nothing
    xEmpleadoID = ""

    'Validacion para empleado inactivo
    if self.empleado.activestatus = 1 then
        'Obtenemos el id del empleado inactivo
		xEmpleadoID = self.empleado.id
        'Armamos la conexion a la base
		set xCon = CreateObject("adodb.connection")
		xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
		xCon.ConnectionTimeout 	= 150
		xCon.Open
        'Activa al empleado
		xCon.Execute("UPDATE EMPLEADO SET ACTIVESTATUS = 0 WHERE ID =  '" & xEmpleadoID & "'" )
		If self.workspace.intransaction then self.workspace.commit
		Set xempleadoActivado =  InstanciarBO(  xEmpleadoID, "EMPLEADO", self.workspace )
		'Lo asigna
        xAccidente.empleado = xempleadoActivado
        xAccidente.ESTADO = "inactivo"
		'Lo pone como inactivo de nuevo
        xCon.Execute("UPDATE EMPLEADO SET ACTIVESTATUS = 1 WHERE ID =  '" & xEmpleadoID & "'" )
    Else
        xAccidente.empleado = self.empleado
    end if

    'Coloco primero este campo, ya que el mismo se encarga de asignar los activos
    xAccidente.CATEGORIA                       = self.CATEGORIA

    'Empresa
    xAccidente.FECHAINGRESOEMPRESA = self.FECHAINGRESOEMPRESA
    xAccidente.COSTOTOTAL          = self.COSTOTOTAL
    xAccidente.COSTOEMPRESA        = self.COSTOEMPRESA
    xAccidente.DIASDEBAJATOTALES   = self.DIASDEBAJATOTALES
    xAccidente.UNIDAD              = self.UNIDAD
    xAccidente.DIASBAJAEMPRESA     = self.DIASBAJAEMPRESA
    xAccidente.CONTROLEXTENSION    = self.CONTROLEXTENSION
    xAccidente.ESTADOACCIDENTE     = self.ESTADOACCIDENTE
    xAccidente.DIASBAJALABORAL     = self.DIASBAJALABORAL
    xAccidente.ACCION              = self.ACCION
    xAccidente.COSTORECUPERO       = self.COSTORECUPERO
    'Legales
    xAccidente.FECHACONCLUSION                 = self.FECHACONCLUSION
    xAccidente.NROJUICIO                       = self.NROJUICIO
    xAccidente.FECHAACCIDENTE                  = self.FECHAACCIDENTE
    xAccidente.DESGLOSE                        = self.DESGLOSE
    xAccidente.AUTODENUNCIA                    = self.AUTODENUNCIA
    xAccidente.FECHAALTA                       = self.FECHAALTA
    xAccidente.ESTADODEJUCIO                   = self.ESTADODEJUCIO
    xAccidente.ANALISTAART                     = self.ANALISTAART
    xAccidente.PROVINCIA                       = self.PROVINCIA
    xAccidente.DIASDEBAJATOTALES               = self.DIRECTASDEBAJATOTALES
    xAccidente.FECHAINGRESO                    = self.FECHAINGRESO
    xAccidente.LETRADODESIGNADO                = self.LETRADODESIGNADO
    xAccidente.SUBCAUSA                        = self.SUBCAUSA
    xAccidente.JURIDICCIONJUICIO               = self.JURIDICCIONJUICIO
    xAccidente.STATUS                          = self.STATUS
    'lIQUIDACIONES
    xAccidente.COSTORECUPERO                   = self.COSTORECUPERO
    xAccidente.COSTOTOTAL                      = self.COSTOTOTAL
    xAccidente.COSTOEMPRESA                    = self.COSTOEMPRESA
    'Siniestro
    xAccidente.ZONACUERPO                      = self.ZONACUERPO
    xAccidente.GESTIONRECHAZO                  = self.GESTIONRECHAZO
    xAccidente.DIASEMANA                       = self.DIASEMANA
    xAccidente.CAUZARAIZ                       = self.CAUZARAIZ
    xAccidente.INTERVALO                       = self.INTERVALO
    xAccidente.RECHAZADOART                    = self.RECHAZADOART
    xAccidente.TIPOACCIDENTE                   = self.TIPOACCIDENTE
    xAccidente.LESION                          = self.LESION
    xAccidente.DIASBAJALABORAL                 = self.DIRECTAJALABORAL
    xAccidente.SINIESTROSASOCIADOS             = self.SINIESTROSASOCIADOS
    xAccidente.ART                             = self.ART
    xAccidente.FORMAACCIDENTE                  = self.FORMAACCIDENTE
    xAccidente.FECHAPREVISTA                   = self.FECHAPREVISTA
    xAccidente.EVENTO                          = self.EVENTO
    xAccidente.SECTORACC                       = self.SECTORACC
    xAccidente.UNIDAD                          = self.UNIDAD
    xAccidente.FECHAACCIDENTE                  = self.FECHAACCIDENTE
    xAccidente.TURNO                           = self.TURNO
    xAccidente.GESTIONRECHAZO                  = self.GESTIONRECHAZ
    xAccidente.GESTIONEFECTIVADERECHAZO        = self.GESTIONEFECTIVADERECHAZO
    xAccidente.NROSINIESTRO                    = self.NROSINIESTRO
    xAccidente.FECHAALTA                       = self.FECHAALTA
    xAccidente.DIASDEBAJATOTALES               = self.DIASDEBAJATOTALES
    xAccidente.CENTRODECOSTOS                  = self.CENTRODECOSTOS

end sub