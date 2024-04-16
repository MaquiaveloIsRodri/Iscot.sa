' 23/02/2021 - PIC - MU Masivo Autorizar PICs.
sub main
	stop

	' PICs Pendientes.
	set xContainer = NewContainer()
	xContainer.Add(ViewPedidos)
	set xSel = SelectViewItems(Self.Workspace, "Autorizar PICs Pendientes", xContainer)
	if xSel.Size <= 0 then
		MsgBox "No seleccionó ningún PIC.", 48, "Aviso"
		exit sub
	end if
	
	call ProgressControl(Self.Workspace, "Masivo Autorizar PICs", 0, 10)
	
	for each xTr in xSel
		SendDebug " > Autorizando PIC: " & xTr.BO.FECHAACTUAL & " - " & xTr.BO.CENTROCOSTOS.NOMBRE & " - " & xTr.BO.Detalle
		call ProgressControlAvance(Self.Workspace, "Autorizando PIC..." & vbNewLine _
			& " - Fecha: " & xTr.BO.FECHAACTUAL & vbNewLine _
			& " - CC: " & xTr.BO.CENTROCOSTOS.NOMBRE & vbNewLine _
			& " - Detalle: " & xTr.BO.Detalle)
		call EjecutarTransicion(xTr.BO, "Autorizar Pedido")
	next
	
	call ProgressControlFinish(Self.Workspace)
end sub

function ViewPedidos()
	set xViewTr = NewCompoundView(Self, "TRORDENCOMPRA", Self.Workspace, nil, true)
	xViewTr.AddJoin(NewJoinSpec(xViewTr.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), false))
	xViewTr.AddJoin(NewJoinSpec(xViewTr.ColumnFromPath("FLAG"), NewColumnSpec("FLAG", "ID", ""), false))
	xViewTr.AddJoin(NewJoinSpec(xViewTr.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_PEDIDOINTERNOCOMPRA", "ID", ""), false))
	xViewTr.AddFilter(NewFilterSpec(xViewTr.ColumnFromPath("FLAG"), " = ", "{FE439A4C-8FC8-49DF-B730-607BC332D173}"))		' Pendiente de Autorizar.
	xViewTr.AddBOCol("NOMBRE").Caption												= " "
	xViewTr.AddBOCol("FLAG").Caption												= " "
	xViewTr.AddBOCol("NOMBRE").Caption												= "Transacción"
	xViewTr.AddColumn(NewColumnSpec("CENTROCOSTOS", "NOMBRE", "")).Caption			= "Centro Costos"
	xViewTr.AddBOCol("VALORTOTAL").Caption											= "Total"
	xViewTr.AddBOCol("DETALLE").Caption												= "Detalle"
	xViewTr.AddColumn(NewColumnSpec("FLAG", "DESCRIPCION", "")).Caption				= "Flag"
	xViewTr.AddBOCol("USUARIO").Caption					 							= "Usuario"
	xViewTr.AddOrderColumn(NewOrderSpec(xViewTr.ColumnFromPath("FECHAACTUAL"), true))
	set ViewPedidos = xViewTr
end function
