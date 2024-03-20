' 24/10/2023 - MU Vincular OC RFIERRO
sub main
	stop
	if Self.Estado <> "A" then
		MsgBox "La Orden debe estar Abierta.", 48, "Aviso"
		exit sub
	end if
	if Self.Destinatario is nothing then
		MsgBox "No indicó: Proveedor.", 48, "Aviso"
		exit sub
	end if
	'OC con filtro sobre el proveedor
	set xViewOC = NewCompoundView(Self, "TRORDENCOMPRA", Self.Workspace, nil, true)
	xViewOC.AddFilter(NewFilterSpec(NewColumnSpec("TRORDENCOMPRA", "NOMBREDESTINATARIO", ""), " = ", " " & Self.Destinatario.name &  ""))
	xViewOC.AddBOCol("NOMBRE")
	xViewOC.AddBOCol("NOMBREDESTINATARIO")

	'VALIDAR
	if xViewOC.ViewItems.Count = 0 then
		MsgBox "No se encontró con el proveedor: " & Self.Destinatario.name &   " .", 48, "Aviso"
		exit sub
	end if

    Set xContainerOrden = NewContainer()
    xContainerOrden.Add(xViewOC)

    set xVisualVar = VisualVarEditor("ORDEN DE COMPRA")
    call AddVarObj( xVisualVar, "00Orde" ,"Orden de compra", "Indique", nothing , xContainerOrden, Self.WorkSpace )
    aceptar = ShowVisualVar(xVisualVar)
    if not aceptar then exit sub
	'Trae OC
    set xOCseleccionada  	= GetValueVisualVar(xVisualVar, "00Orde", "Indique")

	if xOCseleccionada is nothing then 
		MsgBox "No selecciono una OC", 48, "Aviso"
		exit sub
	end if

	res = MsgBox("¿Vincular Factura: " & xOCseleccionada.name & " ?.", 292, "Pregunta")
	if res <> 6 then exit sub

	'Borra los items
	for each xItem in Self.ItemsTransaccion
		xItem.Delete
	next

	'Asignamos la OC
	set Self.VinculoTr		= xOCseleccionada
	'Creamos los items en la orden de compra original
	for each xItemF in xOCseleccionada.ItemsTransaccion
		'Creamos los items en OD
		set xItemNuevo = CrearItemTransaccion(Self)
		xItemNuevo.Referencia				= xItemF.Referencia
		xItemNuevo.Cantidad.Cantidad		= xItemF.Cantidad.Cantidad
		xItemNuevo.Valor.Importe			= xItemF.Valor.Importe
		xItemNuevo.PorcentajeBonificacion	= xItemF.PorcentajeBonificacion
	next
end sub
