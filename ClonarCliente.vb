' Rodrigo Fierro
sub main
	stop
	' Cabecera.
	set xTrNueva = CrearTransaccion("CoCli", Self.UnidadOperativa)
	xTrNueva.FECHAACTUAL 					= Self.FECHAACTUAL
	xTrNueva.DETALLE						= Self.DETALLE
	xTrNueva.NOTA				            = Self.NOTA
	xTrNueva.ORIGINANTE		                = Self.ORIGINANTE
    xTrNueva.COTIZACION				        = Self.COTIZACION
    'xTrNueva.DESTINO				        = Self.DESTINO
    'xTrNueva.IMPORTE		                = Self.IMPORTE'REVISAR
	xTrNueva.CENTROCOSTOS		            = Self.CENTROCOSTOS
    xTrNueva.TOTALCANCELAR		            = Self.TOTALCANCELAR
	xTrNueva.BOExtension.MEDIOPAGO			= Self.BOExtension.MEDIOPAGO
	xTrNueva.BOEXTENSION.FECHAREAL			= Self.BOEXTENSION.FECHAREAL


	' Ítems.
	for each xItem in Self.ItemsTransaccion
		set xItemNuevo = CrearItemTransaccion(xTrNueva)
		xItemNuevo.DESTINO					        = xItem.DESTINO
		xItemNuevo.REFERENCIATIPO			        = xItem.REFERENCIATIPO
		xItemNuevo.VALORORI_IMPORTE	                = xItem.VALORORI_IMPORTE
		xItemNuevo.DATOSCOMPLEMENTARIOS			    = xItem.DATOSCOMPLEMENTARIOS
		xItemNuevo.DETALLE			                = xItem.DETALLE
		xItemNuevo.BOExtension.CONCILIADA	        = xItem.BOExtension.CONCILIADA
		xItemNuevo.BOExtension.DestinoConciliada	= xItem.BOExtension.DestinoConciliada
	next

	ShowBO(xTrNueva)
	res = MsgBox("¿Confirma que desea aplicar los cambios?", 36, "Confirmar")
	if res = 6 then
		call WorkSpaceCheck(xTrNueva.Workspace)
	else
		xTrNueva.Workspace.Rollback
	end if
end sub
