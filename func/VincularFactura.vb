View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "FLAG", ""), NewColumnSpec("FLAG", "ID", ""), True))
View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "CENTROCOSTOS", ""), NewColumnSpec("CENTROCOSTOS", "ID", ""), True))
View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "BOEXTENSION", ""), NewColumnSpec("UD_ORDENCOMPRA", "ID", ""), True))
	'View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "ID", ""), NewColumnSpec("PENDIENTE", "TRANSACCION", ""), False))
	'View.AddFilter(NewFilterSpec(NewColumnSpec("PENDIENTE", "RELACIONTRORIG", ""), " = ", "{967DD785-8E31-4864-9160-C6EA0D73D355}"))
	View.AddJoin(NewJoinSpec(NewColumnSpec("UD_ORDENCOMPRA", "EMPLEADO", ""), NewColumnSpec("EMPLEADO", "ID", ""), True))
	View.AddJoin(NewJoinSpec(NewColumnSpec("EMPLEADO", "ENTEASOCIADO", ""), NewColumnSpec("PERSONAFISICA", "ID", ""), True))
    
	Set xCol        = View.AddColumn(NewColumnSpec("TRORDENCOMPRA", "NOMBREDESTINATARIO", ""))
    xCol.Caption    = "Proveedor"
    xCol.Browse     = True