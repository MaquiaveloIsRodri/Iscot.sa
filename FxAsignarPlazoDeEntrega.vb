sub main
    stop

    set xView = NewCompoundView(Self, "CENTROCOSTOS", Self.Workspace, nil, true)
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xView.AddBOCol("NOMBRE")
	xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NOMBRE"), false))


    For Each xCc in xView.viewitems
        plazoEntrega = 0
        if not xCc.bo.boextension.zona.nombre is nothing
            Select Case xCc.bo.boextension.zona.nombre
                Case "San Juan"
                    plazoEntrega = 7
                Case "Buenos Aires - San Isidro"
                    plazoEntrega = 5
                Case "Santa Fe - Resto de la Provincia"
                    plazoEntrega = 4
                Case "San Luis - ciudad de San Luis"
                    MsgBox "El centro de costos "& xCc.bo.nombre &" tiene a san luis como zona.", 64, "Informacion"
                Case "Buenos Aires - Pilar"
                    plazoEntrega = 5
                Case "Capital Federal"
                    plazoEntrega = 7
                Case "Buenos Aires - Escobar"
                    plazoEntrega = 5
                Case "Buenos Aires - Tres de Febrero"
                    plazoEntrega = 5
                Case "Cordoba - Gran Cordoba"
                    plazoEntrega = 3
                Case "Buenos Aires - Gran Buenos Aires"
                    plazoEntrega = 5
                Case Else
                    MsgBox "El centro de costos "& xCc.bo.nombre &" no tiene una zona asignada.", 64, "Informacion"
            End Select


            for each xProveedor in xCc.bo.boextension.proveedores
                set xPlazoEntrega           = crearbo("UD_PLAZOENTREGA",self)
                xPlazoEntrega.centrocostos  = xCc.bo
                xPlazoEntrega.PROVEEDOR     = xProveedor
                xPlazoEntrega.plazoEntrega  = plazoEntrega
                xCc.bo.boextension.PLAZOENTREGA.add(xPlazoEntrega)
            next
            plazoEntrega = 0
        end if
    next

    MsgBox "La correccion se realizo correctamente.", 64, "Informacion"

end sub
