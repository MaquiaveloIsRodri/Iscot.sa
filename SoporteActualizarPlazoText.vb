sub main
    stop

    set xView = NewCompoundView(Self, "CENTROCOSTOS", Self.Workspace, nil, true)
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xView.AddBOCol("NOMBRE")
	xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NOMBRE"), false))


    For Each xCc in xView.viewitems


            for each xProveedor in xCc.bo.boextension.PLAZOSENTREGA
                xPlazoEntrega.PROVEEDORTEXT     = xProveedor.nombre
            next
		If self.workspace.intransaction Then Self.workspace.commit
            plazoEntrega = 0
        end if
    next

    MsgBox "La correccion se realizo correctamente.", 64, "Informacion"

end sub
