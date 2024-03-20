' CREADO: 17/10/2014 - Jose Fantasia.
' OTE: Orden de Compra.
' MU: Acciones Actualizar Datos.
sub main
	stop
	if Self.Estado <> "A" then
		MsgBox "La Orden de Compra debe estar Abierta.", 64, "Información"
		exit sub
	end if

	' Tipos de Pago Activos.
	set xView = NewCompoundView(Self, "TIPOPAGO", Self.Workspace, nil, true)
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xView.AddBOCol("NOMBRE")
	xView.AddBOCol("OBSERVACION")

	set xVisualVar = VisualVarEditor("CAMBIAR DATOS OC:" & Self.NumeroDocumento)
	call AddVarView(xVisualVar, "00TIPOPAGO", "Tipo de Pago", "Parametros:", xView, "NOMBRE%OBSERVACION")
	call AddVarDate(xVisualVar, "05FECHAESTIMADA", "Fecha Ent. Estimada", "Parametros:", Self.FechaEntrega)
	call AddVarDate(xVisualVar, "10FECHAREAL", "Fecha Ent. Real", "Parametros:", Self.BOEXTENSION.FechaEntregaReal)
	call AddVarBoolean(xVisualVar, "15VERIFICADA", "Verificada", "Parametros:", Self.BOEXTENSION.Verificada)
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then exit sub

	tipoPago_id		= GetValueVisualVar(xVisualVar, "00TIPOPAGO", "Parametros:")
	fechaEstimada 	= CDate(Int(GetValueVisualVar(xVisualVar, "05FECHAESTIMADA", "Parametros:")))
	fechaReal 		= CDate(Int(GetValueVisualVar(xVisualVar, "10FECHAREAL", "Parametros:")))
	verificada		= GetValueVisualVar(xVisualVar, "15VERIFICADA", "Parametros:")
	set xTipoPago 	= nothing

	if tipoPago_id = "" and Self.TipoPago is nothing then
		MsgBox "No indicó el Tipo de Pago.", 42, "Aviso"
		exit sub
	end if
	
	if tipoPago_id <> "" then
	    set xTipoPago = InstanciarBO(tipoPago_id, "TIPOPAGO", Self.Workspace)
		if not xTipoPago is nothing then 
	   	    Self.TipoPago = xTipoPago
		end if
	end if
	Self.FechaEntrega					    = fechaEstimada
	Self.BOEXTENSION.FechaEntregaReal		= fechaReal
	Self.BOEXTENSION.Verificada			    = verificada
	call WorkSpaceCheck(Self.Workspace)
	
	call EjecutarTransicion( Self, "Cerrar Orden" )
end sub