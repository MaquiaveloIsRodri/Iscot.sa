' 23/07/2015 - José Luis Fantasia.
' Vista de Reclamo Compra.
sub main
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRSOLICITUD", "FLAG", ""), NewColumnSpec("FLAG", "ID", ""), true))
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRSOLICITUD", "CENTROCOSTOS", ""), NewColumnSpec("CENTROCOSTOS", "ID", ""), true))
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRSOLICITUD", "BOEXTENSION", ""), NewColumnSpec("UD_RECLAMOCOMPRA", "ID", ""), true))
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRSOLICITUD", "RESPONSABLE", ""), NewColumnSpec("EMPLEADO", "ID", ""), true))
	View.AddJoin(NewJoinSpec(NewColumnSpec("EMPLEADO", "ENTEASOCIADO", ""), NewColumnSpec("PERSONAFISICA", "ID", ""), true))
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRSOLICITUD", "VINCULOTR", ""), NewColumnSpec("TRORDENCOMPRA", "ID", ""), true))
	View.AddJoin(NewJoinSpec(NewColumnSpec("UD_RECLAMOCOMPRA", "TIPORECLAMO", ""), NewColumnSpec("ITEMTIPOCLASIFICADOR", "ID", ""), true))

	'Actualizacion de la view 21/11/2023 RFIERRO
	Set xCol        = View.AddBOCol("NOMBREDESTINATARIO")
	xCol.Caption    = "Proveedor"
	xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("CENTROCOSTOS", "NOMBRE", ""))
	xCol.Caption    = "Centro de Costos"
	xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("ITEMTIPOCLASIFICADOR", "NOMBRE", ""))
	xCol.Caption    = "Tipo de Reclamo"
	xCol.Browse     = True
	Set xCol        = View.AddBOCol("ESTADO")
	xCol.Caption    = "Estado"
	xCol.Browse     = True
	Set xCol        = View.AddBOCol("DETALLE")
	xCol.Caption    = "Descripción"
	xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_RECLAMOCOMPRA", "SOLUCION", ""))
	xCol.Caption    = "Solución"
	xCol.Browse     = True

	View.AddOrderColumn(NewOrderSpec(NewColumnSpec("TRSOLICITUD", "FECHAACTUAL", ""), True))
end sub