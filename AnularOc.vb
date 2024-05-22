' 25/10/2018 - José Luis Fantasia.
' 04/07/2019 - Actualizamos Talles.
' Transición: PIP RH > Cerrar > Ejecuta.
sub main
	stop
	
	' Creamos 1 PIP por CC.
	set xCon = CreateObject("adodb.connection")
	xCon.ConnectionString 	= StringConexion("calipso", Transaccion.Workspace)
	xCon.ConnectionTimeout 	= 150
	
	call ProgressControl(Transaccion.Workspace, "Cerrando PIP RH #" & Transaccion.NumeroDocumento, 0, 4)
	
	set xRs = RecordSet(xCon, "select top 1 * from producto")
	xRs.Close
	xRs.ActiveConnection.CommandTimeout = 0
	xRs.Source = "SELECT DISTINCT cc.ID, cc.NOMBRE " & _
		"FROM TRSOLICITUD AS pip WITH(NOLOCK) " & _
		"INNER JOIN UD_PIPRH AS upip WITH(NOLOCK) ON pip.BOEXTENSION_ID = upip.ID " & _
		"INNER JOIN UD_EMPLEADOPIPRH AS uitem WITH(NOLOCK) ON upip.ITEMS_ID = uitem.BO_PLACE_ID " & _
		"INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON uitem.CENTROCOSTOS_ID = cc.ID " & _
		"WHERE pip.TIPOTRANSACCION_ID = '{B97FC3BF-902D-4B78-B388-3712988F1DB5}' " & _
		"AND pip.ID = '" & Transaccion.ID & "' " & _
		"ORDER BY cc.NOMBRE"
	
	xRs.Open
	do while not xRs.EOF
		set xCC = InstanciarBO(xRs("ID").Value, "CENTROCOSTOS", Transaccion.Workspace)
		call ProgressControlAvance(Transaccion.Workspace, "Creando PIP para Compras..." & vbNewLine & xCC.NOMBRE)
		
		' Responsable.
		set xView = NewCompoundView(Transaccion, "EMPLEADO", Transaccion.Workspace, nil, true)
		xView.NoFlushBuffers = true		' WITH(NOLOCK)
		xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_EMPLEADO", "ID", ""), false))
		xView.AddJoin(NewJoinSpec(NewColumnSpec("UD_EMPLEADO", "USUARIO", ""), NewColumnSpec("USUARIO", "ID", ""), false))
		xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
		xView.AddFilter(NewFilterSpec(NewColumnSpec("USUARIO", "NOMBRE", ""), " = ", Transaccion.Usuario))
		set xResponsable 		= nothing
		for each xItem in xView.ViewItems
			set xResponsable 	= xItem.BO
		next
		if xResponsable is nothing then
			set xResponsable 	= InstanciarBO("{0AF11CD0-D9FB-4C60-BAB4-CF184EC09D4B}", "EMPLEADO", Transaccion.Workspace)		' GENERICO RH.
		end if
		
		' PIP Compras.
		set xTrNueva = CrearTransaccion("PIP", Transaccion.UnidadOperativa)
		call NoMensaje(xTrNueva, true)
		set xTrNueva.BOEXTENSION.PIPRH		= Transaccion
		xTrNueva.FechaActual				= Now()
		xTrNueva.Destinatario 				= xResponsable
		xTrNueva.CentroCostos				= xCC
		xTrNueva.ImputacionContable			= Transaccion.ImputacionContable
		xTrNueva.Detalle					= Transaccion.Detalle
		xTrNueva.BOEXTENSION.Urgente		= Transaccion.BOEXTENSION.Urgente
		xTrNueva.BOEXTENSION.FechaEntrega	= Transaccion.BOEXTENSION.FechaEntrega
		xTrNueva.BOEXTENSION.DETALLEREPOSICION	= Transaccion.BOEXTENSION.DETALLEREPOSICION



		
		set xDeposito 		= Nothing
		set xUbicacion		= Nothing
		'if not xCC.BOExtension.Deposito is nothing then
		'    set xDeposito 	= xCC.BOExtension.Deposito
		'	
		'	set xView = NewCompoundView(xTrNueva, "UBICACION", xTrNueva.Workspace, nil, true)
		'	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", xDeposito.Ubicaciones.ID))
		'	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("NOMBRE"), " = ", "Nueva"))
		'	if xView.ViewItems.Count > 0 then
		'		set xUbicacion = xView.ViewItems.First.Current.BO
		'	end if
		'end if
		
		' Ítems.
		query	= "SELECT uitem.PRODUCTO_ID " & _
			",uitem.CANTIDAD " & _
			",uitem.EMPLEADO_ID " & _
			",CONVERT(date, uitem.FECHAENTREGA, 103) AS FECHAENTREGA " & _
			",uitem.ADJUNTO,uitem.OBSERVACIONES" & _
			"FROM TRSOLICITUD AS pip WITH(NOLOCK) " & _
			"INNER JOIN UD_PIPRH AS upip WITH(NOLOCK) ON pip.BOEXTENSION_ID = upip.ID " & _
			"INNER JOIN UD_EMPLEADOPIPRH AS uitem WITH(NOLOCK) ON upip.ITEMS_ID = uitem.BO_PLACE_ID " & _
			"INNER JOIN EMPLEADO AS empl WITH(NOLOCK) ON uitem.EMPLEADO_ID = empl.ID " & _
			"INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = per.ID " & _
			"WHERE pip.TIPOTRANSACCION_ID = '{B97FC3BF-902D-4B78-B388-3712988F1DB5}' " & _
			"AND pip.ID = '" & Transaccion.ID & "' " & _
			"AND uitem.CENTROCOSTOS_ID = '" & xCC.ID & "' " & _
			"ORDER BY per.NOMBRE"
		
		if Transaccion.ImputacionContable.ID = "{8E9C59B0-57B2-4C59-84E4-4FA6337BC8CE}" then		' Normal.
			' No interesa el Empleado, agrupamos por Cantidad
			' y ordenamos según le conviene a Compras.
			query	= "SELECT uitem.PRODUCTO_ID " & _
				",tipo.NOMBRE " & _
				",prod.DESCRIPCION " & _
				",SUM(uitem.CANTIDAD) AS CANTIDAD, uitem.OBSERVACIONES " & _
				"FROM TRSOLICITUD AS pip WITH(NOLOCK) " & _
				"INNER JOIN UD_PIPRH AS upip WITH(NOLOCK) ON pip.BOEXTENSION_ID = upip.ID " & _
				"INNER JOIN UD_EMPLEADOPIPRH AS uitem WITH(NOLOCK) ON upip.ITEMS_ID = uitem.BO_PLACE_ID " & _
				"INNER JOIN PRODUCTO AS prod WITH(NOLOCK) ON uitem.PRODUCTO_ID = prod.ID " & _
				"INNER JOIN UD_PRODUCTO AS uprod WITH(NOLOCK) ON prod.BOEXTENSION_ID = uprod.ID " & _
				"INNER JOIN ITEMTIPOCLASIFICADOR AS tipo WITH(NOLOCK) ON uprod.TIPOINDUMENTARIA2_ID = tipo.ID " & _
				"WHERE pip.TIPOTRANSACCION_ID = '{B97FC3BF-902D-4B78-B388-3712988F1DB5}' " & _
				"AND pip.ID = '" & Transaccion.ID & "' " & _
				"AND uitem.CENTROCOSTOS_ID = '" & xCC.ID & "' " & _
				"GROUP BY uitem.PRODUCTO_ID, tipo.NOMBRE, prod.DESCRIPCION, uitem.OBSERVACIONES " & _
				"ORDER BY tipo.NOMBRE, prod.DESCRIPCION"
		end if
		
		set xRsI = RecordSet(xCon, "select top 1 * from producto")
		xRsI.Close
		xRsI.ActiveConnection.CommandTimeout = 0
		xRsI.Source = query
		xRsI.Open
		do while not xRsI.EOF
			set xItemNuevo = CrearItemTransaccion(xTrNueva)
			senddebug(xRsI("PRODUCTO_ID").Value)
			set xItemNuevo.Referencia				= InstanciarBO(xRsI("PRODUCTO_ID").Value, "PRODUCTO", Transaccion.Workspace)
			if xItemNuevo.Referencia Is Nothing Then
			   senddebug(xRsI("PRODUCTO_ID").Value)
			   stop
			End If
			xItemNuevo.Descripcion					= xItemNuevo.Referencia.Descripcion
			xItemNuevo.OBSERVACION.MEMO				= xRsI("OBSERVACIONES").Value
			xItemNuevo.Cantidad.Cantidad			= CDbl(xRsI("CANTIDAD").Value)
			if Transaccion.ImputacionContable.ID <> "{8E9C59B0-57B2-4C59-84E4-4FA6337BC8CE}" then		' Normal.
			    Set xEmpleado = InstanciarBO(xRsI("EMPLEADO_ID").Value, "EMPLEADO", Transaccion.Workspace)
				If xEmpleado.ActiveStatus <> 0 Then set xEmpleado = InstanciarBO("{0AF11CD0-D9FB-4C60-BAB4-CF184EC09D4B}", "EMPLEADO", Transaccion.Workspace)		' GENERICO RH.
				set xItemNuevo.BOExtension.Empleado	= xEmpleado
				xItemNuevo.BOExtension.FechaEntrega	= CDate(xRsI("FECHAENTREGA").Value)
				If Not IsNull(xRsI("ADJUNTO").Value) Then xItemNuevo.BOExtension.Adjunto = xRsI("ADJUNTO").Value
			end if
			'set xItemNuevo.DepositoOri 				= xDeposito
			'set xItemNuevo.UbicacionOri 			= xUbicacion
			
			xRsI.MoveNext
		loop
		
		if xTrNueva.Workspace.InTransaction then xTrNueva.Workspace.Commit
		call ProgressControlAvance(Transaccion.Workspace, "PIP Compras #" & xTrNueva.NumeroDocumento & " Creado!!" & vbNewLine & xCC.NOMBRE)
		Transaccion.VinculosTransaccionales.Add(xTrNueva)
		'call EjecutarTransicion(xTrNueva, "Cerrar RH")
		
		xRs.MoveNext
	loop
	
	' 04/07/2019 - Actualizamos Talles.
	for each xItemRH in Transaccion.BOExtension.Items
		if (not xItemRH.Empleado is nothing) and (not xItemRH.Producto is nothing) then
			if not xItemRH.Producto.BOExtension.TIPOINDUMENTARIA2 is nothing then
				for each xItemTalle in xItemRH.Empleado.BOExtension.IndumentariaTalles
					If Not xItemTalle.Producto Is Nothing Then
					   if not xItemTalle.Producto.BOExtension.TIPOINDUMENTARIA2 is nothing then
						  if xItemRH.Producto.BOExtension.TIPOINDUMENTARIA2.ID = xItemTalle.Producto.BOExtension.TIPOINDUMENTARIA2.ID then
							 xItemTalle.Delete
							 exit for
						   end if
						end if
					End If
				next
			end if
			
			call ProgressControlAvance(Transaccion.Workspace, "Actualizando Talles: " & xItemRH.Empleado.EnteAsociado.Nombre & vbNewLine & "Producto: " & xItemRH.Producto.Descripcion)
			
			set xTalle = CrearBO("UD_EMPLEADOTALLES", xItemRH.Empleado)
			set xTalle.Producto		= xItemRH.Producto
			xTalle.Descripcion		= xItemRH.Producto.Descripcion
			xTalle.Cantidad			= xItemRH.Cantidad
			xItemRH.Empleado.BOExtension.IndumentariaTalles.Add(xTalle)
		end if
	next
	set xFSO 		= CreateObject("Scripting.FileSystemObject")
	set xArchivo = xFSO.OpenTextFile("C:\util\html\email-NuevoPicCompra.html")
	htmlBody = xArchivo.readAll()

	items = items & "<tr>"
	items = items & "<td>"& Transaccion.name & "</td>"
	items = items & "<td>"& Transaccion.FechaActual & "</td>"
	items = items & "<td>"& xResponsable.name & "</td>"
	items = items & "<td>"& XCC.name & "</td>"
	items = items & "<td>"& Transaccion.Detalle & "</td>"
	items = items & "<td>"& Transaccion.BOEXTENSION.Urgente & "</td>"
	items = items & "<td>"& Transaccion.BOEXTENSION.FechaEntrega & "</td>"
	items = items & "</tr>"

	htmlBody = replace(htmlBody,"xItems" , items)


	xSubject = "[Nuevo Pip] " & Date()
	xBody    = ""
	xcorreos = "rodrigofierrro@gmail.com;"
	xadjunto = ""

 '   set xRs2 = RecordSet(xCon, "select top 1 * from producto")
'	xRs2.Close
'	xRs2.ActiveConnection.CommandTimeout = 0
'	xRs2.Source = "SELECT itc.NOMBRE Correo " & _
'		"FROM ITEMTIPOCLASIFICADOR itc " & _
'		"WHERE itc.BO_PLACE_ID = '{10678499-05FA-47B0-9007-FD2DEB0958D7}'   " & _
'		"AND  itc.ACTIVESTATUS <> 2  "
'	
'	xRs2.Open
'	do while not xRs2.EOF
'	    xcorreos = xcorreos & ";"& xRs2("Correo").Value 
'	    
'		xRs2.MoveNext
'   loop
	call enviar_aviso_sinmsg(Transaccion,"paniolcentral@iscot.com.ar", xSubject, xBody, xadjunto,htmlBody)
	call ProgressControlFinish(Transaccion.Workspace)
end sub


















