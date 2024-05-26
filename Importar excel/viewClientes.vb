' 20/09/2023 - Candidatos Ítem PIC Servicios - @josefantasia.
'ServiciosPIC
sub main
	stop
	set xTr = Owner.PlaceOwner
	
	if xTr.CentroCostos is nothing then
		MsgBox "Indique: Centro de Costos.", 48, "Aviso"
		aView.AddFilter(NewFilterSpec(aView.ColumnFromPath("ACTIVESTATUS"), " = ", 9))
		exit sub
	end if
	
    'cambio 25-04-2024 Rodrigo fierro



	aView.AddJoin(NewJoinSpec(aView.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), false))
    aView.AddJoin(NewJoinSpec(NewColumnSpec("CENTROCOSTOS", "BOEXTENSION", ""), NewColumnSpec("UD_CENTROCOSTOS", "ID", ""), false))
    aView.AddJoin(NewJoinSpec(NewColumnSpec("UD_CENTROCOSTOS", "PLAZOSENTREGA", ""), NewColumnSpec("UD_PLAZOENTREGA", "BO_PLACE_ID", ""), false))
    aView.AddJoin(NewJoinSpec(NewColumnSpec("UD_PLAZOENTREGA", "PROVEEDOR", ""), NewColumnSpec("PROVEEDOR", "ID", ""), false))
    aView.AddJoin(NewJoinSpec(NewColumnSpec("PROVEEDOR", "ENTEASOCIADO", ""), NewColumnSpec("PERSONA", "ID", ""), false))
    aView.AddJoin(NewJoinSpec(NewColumnSpec("PROVEEDOR", "LISTAPRECIO", ""), NewColumnSpec("LISTAPRECIO", "ID", ""), false))
    aView.AddJoin(NewJoinSpec(NewColumnSpec("LISTAPRECIO", "PRECIOS", ""), NewColumnSpec("PRECIO", "BO_PLACE", ""), false))
    aView.AddFilter(NewFilterSpec(NewColumnSpec("PROVEEDOR", "ACTIVESTATUS", ""), " = ", 0))
    aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "REFERENCIA", ""), " = ", xTr.ID))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "DESDEFECHA", ""), " <= ", Now))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "HASTAFECHA", ""), " >= ", Now))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("CENTROCOSTOS", "ID", ""), " = ", xTr.CentroCostos.ID))


    aView.AddJoin(NewJoinSpec(NewColumnSpec("PRECIO", "BO_PLACE", ""), NewColumnSpec("LISTAPRECIO", "PRECIOS", ""), false))

	aView.AddJoin(NewJoinSpec(NewColumnSpec("PRECIO", "BO_PLACE", ""), NewColumnSpec("LISTAPRECIO", "PRECIOS", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("LISTAPRECIO", "ID", ""), NewColumnSpec("PROVEEDOR", "LISTAPRECIO", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("PROVEEDOR", "ENTEASOCIADO", ""), NewColumnSpec("PERSONA", "ID", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("PROVEEDOR", "ID", ""), NewColumnSpec("PERSLIST", "ITEM", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("PERSLIST", "ID", ""), NewColumnSpec("BOLIST", "BO_ITEMS", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("BOLIST", "ID", ""), NewColumnSpec("UD_CENTROCOSTOS", "PROVEEDORSERVICIO", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("UD_CENTROCOSTOS", "ID", ""), NewColumnSpec("CENTROCOSTOS", "BOEXTENSION", ""), false))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PROVEEDOR", "ACTIVESTATUS", ""), " = ", 0))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "DESDEFECHA", ""), " <= ", Now))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "HASTAFECHA", ""), " >= ", Now))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("CENTROCOSTOS", "ID", ""), " = ", xTr.CentroCostos.ID))


	aView.AddJoin(NewJoinSpec(aView.ColumnFromPath("ID"), NewColumnSpec("PRECIO", "REFERENCIA", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("PRECIO", "BO_PLACE", ""), NewColumnSpec("LISTAPRECIO", "PRECIOS", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("LISTAPRECIO", "ID", ""), NewColumnSpec("PROVEEDOR", "LISTAPRECIO", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("PROVEEDOR", "ENTEASOCIADO", ""), NewColumnSpec("PERSONA", "ID", ""), false))
	' aView.AddJoin(NewJoinSpec(NewColumnSpec("PROVEEDOR", "ID", ""), NewColumnSpec("PERSLIST", "ITEM", ""), false))
	' aView.AddJoin(NewJoinSpec(NewColumnSpec("PERSLIST", "ID", ""), NewColumnSpec("BOLIST", "BO_ITEMS", ""), false))
    aView.NewInListFilterSpec(NewColumnSpec("PROVEEDOR", "ENTEASOCIADO", ""), NewColumnSpec("PERSONA", "ID", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("BOLIST", "ID", ""), NewColumnSpec("UD_CENTROCOSTOS", "PROVEEDORSERVICIO", ""), false))
	aView.AddJoin(NewJoinSpec(NewColumnSpec("UD_CENTROCOSTOS", "ID", ""), NewColumnSpec("CENTROCOSTOS", "BOEXTENSION", ""), false))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PROVEEDOR", "ACTIVESTATUS", ""), " = ", 0))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "DESDEFECHA", ""), " <= ", Now))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("PRECIO", "HASTAFECHA", ""), " >= ", Now))
	aView.AddFilter(NewFilterSpec(NewColumnSpec("CENTROCOSTOS", "ID", ""), " = ", xTr.CentroCostos.ID))
	
	Set xCol		= aView.AddColumn(NewColumnSpec("PRECIO", "VALOR2_IMPORTE", ""))
	xCol.Caption	= "($) Precio"
	xCol.Browse		= True
	Set xCol		= aView.AddColumn(NewColumnSpec("PERSONA", "NOMBRE", ""))
	xCol.Caption	= "Razón Social"
	xCol.Browse		= True
	Set xCol		= aView.AddColumn(NewColumnSpec("PROVEEDOR", "DENOMINACION", ""))
	xCol.Caption	= "Nombre de Fantasia"
	xCol.Browse		= True
end sub