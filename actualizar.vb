sub main
	stop
	set xVisualVar = VisualVarEditor("Cambiar Fecha de Entrega Real")
	call AddVarDate(xVisualVar, "00FECHAENTREGAREAL", "Nueva Fecha", "Parámetros", Date)
	aceptar = ShowVisualVar(xVisualVar)

    if aceptar then
        for each xOc in container
			Self.BOEXTENSION.FechaEntregaReal = nuevaFecha
			call WorkSpaceCheck(self.workspace)
        next
    end if

	MsgBox "Cambio de fecha de entrega real finalizado!!.", 64, "Informacion"
end sub






' creado: 11/06/2013 - Jose Fantasia.
' MU Cambiar Fecha de Entrega Real.
sub main
	stop
	set xVisualVar = VisualVarEditor("Cambiar Fecha de Entrega Real")
	call AddVarDate(xVisualVar, "00FECHAENTREGAREAL", "Nueva Fecha", "Parámetros", Date)
	aceptar = ShowVisualVar(xVisualVar)
	if aceptar then
		nuevaFecha = CDate(Int(GetValueVisualVar(xVisualVar, "00FECHAENTREGAREAL", "Parámetros")))
		' AGREGADO: 17/01/2014 - Jose Fantasia.
		res = MsgBox("¿Desea hacer un cambio masivo de fechas?", 292, "Fecha Entrega REAL")
		if res = 6 then
			set xVisualVarOC = VisualVarEditor("ORDENES ENTRE FECHAS")
			call AddVarDate(xVisualVarOC, "00FECHADESDE", "Fecha Desde", "Parámetros", Date)
			call AddVarDate(xVisualVarOC, "05FECHAHASTA", "Fecha Hasta", "Parámetros", Date)
			aceptar = ShowVisualVar(xVisualVarOC)
			if aceptar then
				fechaDesde = CDate(Int(GetValueVisualVar(xVisualVarOC, "00FECHADESDE", "Parámetros")))
				fechaHasta = CDate(Int(GetValueVisualVar(xVisualVarOC, "05FECHAHASTA", "Parámetros")))
				if fechaDesde > fechaHasta then
					MsgBox "La Fecha Desde NO puede ser Mayor que la Fecha Hasta.", 48, "Aviso"
					exit sub
				end if
				
				set xContainer = NewContainer()
				xContainer.Add(ObtenerView(fechaDesde, fechaHasta))
				set xSeleccionados = SelectViewItems(Self.Workspace, "TRORDENCOMPRA", xContainer)
				if xSeleccionados.Size <= 0 then
					MsgBox "No seleccionó ningún ítem.", 48, "Aviso"
					exit sub
				end if
				
				for each xItemSel in xSeleccionados
					xItemSel.BO.BOEXTENSION.FechaEntregaReal = nuevaFecha
					call WorkSpaceCheck(self.workspace)
				next
			else
				exit sub
			end if
		else
			Self.BOEXTENSION.FechaEntregaReal = nuevaFecha
			call WorkSpaceCheck(self.workspace)
		end if
		MsgBox "Fecha(s) Actualizadas!!.", 64, "Información"
	end if
end sub

function ObtenerView(pFechaDesde, pFechaHasta)
	set xView = NewCompoundView(Self, "TRORDENCOMPRA", Self.Workspace, nil, true)
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), false))
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_ORDENCOMPRA", "ID", ""), false))
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("FLAG"), NewColumnSpec("FLAG", "ID", ""), false))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("FLAG"), " <> ", "{C10833F4-5E1E-47DA-803E-5FBF135BEA51}"))		' Anulado.
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ESTADO"), " <> ", "N"))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("FECHAACTUAL"), " >= ", pFechaDesde))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("FECHAACTUAL"), " <= ", pFechaHasta))
	xView.AddBOCol("NOMBRE").Caption													= " "
	xView.AddBOCol("FLAG").Caption														= " "
	xView.AddBOCol("NOMBRE").Caption													= "Transacción"
	xView.AddBOCol("NOMBREDESTINATARIO").Caption 										= "Proveedor"
	xView.AddColumn(NewColumnSpec("CENTROCOSTOS", "NOMBRE", "")).Caption 				= "Centro Costos"
	xView.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "FECHAENTREGAREAL", "")).Caption 	= "Fecha Ent. Real"
	xView.AddBOCol("NOTA").Caption 					   	 								= "Observaciones"
	xView.AddBOCol("ESTADO").Caption 					   	 							= "Estado"
	xView.AddColumn(NewColumnSpec("FLAG", "DESCRIPCION", "")).Caption					= "Flag"
	xView.AddBOCol("USUARIO").Caption					 								= "Usuario"
	xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NUMERODOCUMENTO"), true))
	set ObtenerView = xView
end function
