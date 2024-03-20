'rodrigo fierro 04/10/2023
Sub Main
	stop
	set xFactura = Transaccion.VinculoTr

	if Not xFactura is nothing Then
		xstatus = Transaccion.VinculoTr.boextension.Empleado.activestatus
		if xstatus <> 0 Then
        	xEmpleadoID = Transaccion.VinculoTr.boextension.Empleado.id
        	set xCon = CreateObject("adodb.connection")
			xCon.ConnectionString 	= StringConexion("calipso", transaccion.Workspace)
			xCon.ConnectionTimeout 	= 150
			xCon.Open
			xCon.Execute("UPDATE EMPLEADO SET ACTIVESTATUS = 0 WHERE ID =  '" & xEmpleadoID & "'" )
			transaccion.nota = "inactivo"
			If transaccion.workspace.intransaction then transaccion.workspace.commit
			Set xempleadoActivado =  InstanciarBO(  xEmpleadoID, "EMPLEADO", transaccion.workspace )
			'Completamos los datos
			transaccion.BOEXTENSION.EMPLEADO = xempleadoActivado
			transaccion.CENTROCOSTOS = Transaccion.VinculoTr.boextension.CENTROCOSTOS
			transaccion.boextension.CUIL = Transaccion.VinculoTr.boextension.CUIL
			transaccion.boextension.MOTIVOEGRESO  = Transaccion.VinculoTr.boextension.MOTIVOBAJA
			transaccion.boextension.NUMEROEXPEDIENTE = Transaccion.VinculoTr.boextension.NUMEROEXPEDIENTE
			'Validacion de las fechas
			if isDate(Transaccion.VinculoTr.boextension.FECHAINGRESO) then 
				transaccion.boextension.FECHAINGRESO = Transaccion.VinculoTr.boextension.FECHAINGRESO 
			else
				transaccion.boextension.FECHABAJA = #00:00:00# 
			end if
    		if isDate(Transaccion.VinculoTr.boextension.FECHABAJA) then 
				transaccion.boextension.FECHABAJA = Transaccion.VinculoTr.boextension.FECHABAJA 
			else
				transaccion.boextension.FECHABAJA = #00:00:00# 
			end if
			xCon.Execute("UPDATE EMPLEADO SET ACTIVESTATUS = 1 WHERE ID =  '" & xEmpleadoID & "'" )
		else
			'Completamos los datos
			transaccion.BOEXTENSION.LEGAJO = Transaccion.VinculoTr.boextension.codigo
			transaccion.BOEXTENSION.EMPLEADO = Transaccion.VinculoTr.boextension.Empleado
			transaccion.CENTROCOSTOS = Transaccion.VinculoTr.boextension.CENTROCOSTOS
			transaccion.boextension.CUIL = Transaccion.VinculoTr.boextension.CUIL
			transaccion.boextension.MOTIVOEGRESO  = Transaccion.VinculoTr.boextension.MOTIVOBAJA
			transaccion.boextension.NUMEROEXPEDIENTE = Transaccion.VinculoTr.boextension.NUMEROEXPEDIENTE
			'Validacion de las fechas
			if isDate(Transaccion.VinculoTr.boextension.FECHAINGRESO) then 
				transaccion.boextension.FECHAINGRESO = Transaccion.VinculoTr.boextension.FECHAINGRESO 
			else
				transaccion.boextension.FECHABAJA = #00:00:00# 
			end if
        	if isDate(Transaccion.VinculoTr.boextension.FECHABAJA) then 
				transaccion.boextension.FECHABAJA = Transaccion.VinculoTr.boextension.FECHABAJA 
			else
				transaccion.boextension.FECHABAJA = #00:00:00# 
			end if
		end if
	Else
        transaccion.boextension.LEGAJO = ""
		transaccion.boextension.Empleado = Nothing 
		transaccion.boextension.CENTROCOSTOS = Nothing
		transaccion.boextension.CUIL = ""
		transaccion.boextension.FECHAINGRESO = #00:00:00#
        transaccion.boextension.FECHABAJA = #00:00:00#
		transaccion.boextension.MOTIVOEGRESO  = nothing
		transaccion.boextension.NUMEROEXPEDIENTE = ""
	End If
End Sub

Sub Main
	stop
 	set xFactura = Transaccion.VinculoTr

	if Not xFactura is nothing Then
		xEmpleado = Transaccion.VinculoTr.destinatario
		if xEmpleado.activestatus <> 0 Then
        	xEmpleadoID = Transaccion.VinculoTr.destinatario.id
        	set xCon = CreateObject("adodb.connection")
			xCon.ConnectionString 	= StringConexion("calipso", transaccion.Workspace)
			xCon.ConnectionTimeout 	= 150
			xCon.Open
			xCon.Execute("UPDATE EMPLEADO SET ACTIVESTATUS = 0 WHERE ID =  '" & xEmpleadoID & "'" )
			transaccion.nota = "inactivo"
			If transaccion.workspace.intransaction then transaccion.workspace.commit
			Set xempleadoActivado =  InstanciarBO(  xEmpleadoID, "EMPLEADO", transaccion.workspace )
			transaccion.BOEXTENSION.EMPLEADO = xempleadoActivado
			transaccion.BOEXTENSION.LEGAJO = xempleadoActivado.codigo
        	transaccion.boextension.fechabaja = Transaccion.VinculoTr.boextension.fechabaja
			transaccion.boextension.motivobaja = Transaccion.VinculoTr.boextension.motivoegreso
 			transaccion.boextension.MOTIVOBAJAAFIP = Transaccion.VinculoTr.boextension.MOTIVOAFIP
			xCon.Execute("UPDATE EMPLEADO SET ACTIVESTATUS = 1 WHERE ID =  '" & xEmpleadoID & "'" )
		else
			transaccion.BOEXTENSION.EMPLEADO = xEmpleado
			transaccion.BOEXTENSION.LEGAJO = xEmpleado.codigo
        	transaccion.boextension.fechabaja = Transaccion.VinculoTr.boextension.fechabaja
			transaccion.boextension.motivobaja = Transaccion.VinculoTr.boextension.motivoegreso
 			transaccion.boextension.MOTIVOBAJAAFIP = Transaccion.VinculoTr.boextension.MOTIVOAFIP
	Else
  	   	transaccion.boextension.legajo = ""
		transaccion.boextension.empleado = nothing
		transaccion.boextension.cuil = ""
		transaccion.boextension.centrocostos = nothing
		transaccion.boextension.FECHAINGRESO = #00:00:00#
		transaccion.boextension.fechabaja = #00:00:00#
		transaccion.boextension.MOTIVOBAJA = ""
		transaccion.boextension.MOTIVOBAJAAFIP = nothing
	End If
End Sub

