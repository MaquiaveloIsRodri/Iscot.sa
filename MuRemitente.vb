sub main
    stop

    if newvalue.value.id = "{6FD8E331-C1E4-4C5E-8379-B9F0FEBBF579}" then
        'motivo     D36C3DB2-869F-4B5C-B2E1-99E819748769
        Set oMotivos  = ExisteBo(owner, "TIPOCLASIFICADOR", "ID", "{D36C3DB2-869F-4B5C-B2E1-99E819748769}", nil, True, false, "=")
        Set oMotivo   = ExisteBo(owner, "ITEMTIPOCLASIFICADOR", "ID", "{2FD9E38F-EED2-4860-966B-5242C8AD0F68}" , oMotivos.Valores, True, False, "=") 'PENDIENTE A RESPONDER

        owner.MOTIVO = oMotivo

        'Novedades   724B7CD5-E6A1-40E7-8492-0B280DAC9131
        Set oNovedades  = ExisteBo(owner, "TIPOCLASIFICADOR", "ID", "{724B7CD5-E6A1-40E7-8492-0B280DAC9131}", nil, True, false, "=")
        Set oNovedad    = ExisteBo(owner, "ITEMTIPOCLASIFICADOR", "ID", "{0CA43204-7E51-48A8-8693-0A6B53E9A893}" , oNovedades.Valores, True, False, "=") 'PENDIENTE A RESPONDER

        owner.NOVEDADESINTERCAMBIO = oNovedad
        
    end if


end sub