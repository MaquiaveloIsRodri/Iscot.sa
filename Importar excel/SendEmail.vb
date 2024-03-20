Sub Main
	stop
	'Valida que el campo traiga datos
	If newvalue.value.name = "" Then 
		Call MsgBox("Tipo de cuenta no existente.",64, "Información")
		exit sub
	End If


	Set oTipoCuenta = ExisteBo(transaccion,"TipoClasificador","ID","3930B69B-B768-400E-96B8-F99AF5A8FC05",nil,True,False,"=")
	set xTipoCuenta = ExisteBo(transaccion, "ITEMTIPOCLASIFICADOR", "id", newvalue.value.id, oTipoCuenta.Valores , True, False, "=")	{...}	IBO

	'Valida objeto
	If  xTipoCuenta.name <> "S - Sueldo" Then 
		exit sub
	end if

	'Carga datos
	set oempleado  = transaccion.boextension.empleado
    For Each xItem In transaccion.ItemsTransaccion
		i = i + 1
        xitem.boextension.cbu = oempleado.cbu
	Next
End Sub




