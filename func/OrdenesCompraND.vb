
sub main
	stop
	if not (PerteneceAGrupo( "CALIPSO FACTURACION" ) or UCase(NombreUsuario()) = "SCALIPSO") then
		MsgBox "Usuario No Autorizado.", 48, "Aviso"
		exit sub
	end if

	' Creamos la ND.
	'set xFlag			 	= InstanciarBO("D6B3920D-C3FA-4B1F-921E-D653E3B9E170", "FLAG", Self.Workspace)	  ' Facturado.
	set xUO 				= InstanciarBO("CEA52A93-43D9-429A-9708-AD75A1183343", "UOADMINISTRACIONCOMPRAS", Self.Workspace)
    set xTrNueva 			= CrearTransaccion("NoDeCo", xUO)
	xTrNueva.Destinatario 	= Self.Destinatario
	set xTrNueva.TIPOPAGO	= self.TIPOPAGO


    'set self.VinculoTr	= xTrNueva
    'self.VinculosTransaccionales.Add(xTrNueva)

	nro_item = 0
	for each xItem in self.ItemsTransaccion
		nro_item = nro_item + 1
		' ND.
		set xItemNuevo 							= CrearItemTransaccion(xTrNueva)

        xItemNuevo.REFERENCIA               = xItem.REFERENCIA
        xItemNuevo.DESCRIPCION              = xItem.DESCRIPCION
        xItemNuevo.CANTIDAD.CANTIDAD        = xItem.CANTIDAD.CANTIDAD
        xItemNuevo.VALOR.IMPORTE            = xItem.VALOR.IMPORTE
        xItemNuevo.PORCENTAJEBONIFICACION   = xItem.PORCENTAJEBONIFICACION
        xItemNuevo.IMPORTEBONIFICADO        = xItem.IMPORTEBONIFICADO
        xItemNuevo.TOTALSINDESCUENTOS       = xItem.TOTALSINDESCUENTOS
        xItemNuevo.CENTROCOSTOS             = xItem.CENTROCOSTOS

    next



	MsgBox "Recuerde Verificar antes de Guardar.", 48, "Aviso"
	ShowBO(xTrNueva)
	res = MsgBox("¿Guardar Nota de Débito?", 36, "Confirmar")
	if res = 6 then
		a = WorkSpaceCheck(xTrNueva.Workspace)
		if a = 0 then
			MsgBox "Nota de Débito Creada", 64, "Información"
		else
			MsgBox "Por algun campo necesario, no se genero la ND", 64, "Información"
		end if
	else
		xTrNueva.Workspace.Rollback
	end if
end sub


