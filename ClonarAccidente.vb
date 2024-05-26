sub main
    stop
    'Creo un nuevo elemento en la UD
    set xAccidente = CrearBO("UD_ACCIDENTES", self.WorkSpace)
    if self.empleado.activestatus = then  
    xAccidente.empleado = self.empleado

    'Creo un nuevo elemento en la UD
    'Empresa
    xAccidente.FECHAINGRESOEMPRESA = FECHAINGRESOEMPRESA
    xAccidente.COSTOTOTAL          = COSTOTOTAL
    xAccidente.COSTOEMPRESA        = COSTOEMPRESA
    xAccidente.DIASDEBAJATOTALES   = DIASDEBAJATOTALES
    xAccidente.UNIDAD              = UNIDAD
    xAccidente.DIASBAJAEMPRESA     = DIASBAJAEMPRESA
    xAccidente.CONTROLEXTENSION    = CONTROLEXTENSION
    xAccidente.ESTADOACCIDENTE     = ESTADOACCIDENTE
    xAccidente.DIASBAJALABORAL     = DIASBAJALABORAL
    xAccidente.ACCION              = ACCION
    xAccidente.COSTORECUPERO       = COSTORECUPERO
    'Legales
    FECHACONCLUSION                 = FECHACONCLUSION
    NROJUICIO                       = NROJUICIO
    FECHAACCIDENTE                  = FECHAACCIDENTE
    DESGLOSE                        = DESGLOSE
    AUTODENUNCIA                    = AUTODENUNCIA
    FECHAALTA                       = FECHAALTA
    ESTADODEJUCIO                   = ESTADODEJUCIO
    ANALISTAART                     = ANALISTAART
    PROVINCIA                       = PROVINCIA
    DIASDEBAJATOTALES               = DIRECTASDEBAJATOTALES
    FECHAINGRESO                    = FECHAINGRESO
    LETRADODESIGNADO                = LETRADODESIGNADO
    SUBCAUSA                        = SUBCAUSA
    JURIDICCIONJUICIO               = JURIDICCIONJUICIO
    STATUS                          = STATUS
    'lIQUIDACIONES
    COSTORECUPERO                   = COSTORECUPERO
    COSTOTOTAL                      = COSTOTOTAL
    COSTOEMPRESA                    = COSTOEMPRESA
    'Siniestro
    ZONACUERPO                      = ZONACUERPO
    GESTIONRECHAZO                  = GESTIONRECHAZO
    DIASEMANA                       = DIASEMANA
    CAUZARAIZ                       = CAUZARAIZ
    INTERVALO                       = INTERVALO
    RECHAZADOART                    = RECHAZADOART
    TIPOACCIDENTE                   = TIPOACCIDENTE

end sub