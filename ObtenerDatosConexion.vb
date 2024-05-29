' Creacion:     2024-05-28
' Creador:      Rodrigo Fierro
' Descripcion:  ME se encarga de cargar el telefono del supervisor
sub main
    stop

    if not newValue.value is nothing then
        set empleado = newValue.value
        set xtel		= ExisteBo(Self,"UD_CELULARES","RESPONSABLE_ID",empleado.id,nil,True,False,"=")
        if not xtel is nothing then
            owner.boextension.telefono = xtel.numerolinea
        end if
    end if
end sub