' CREADO: 29/11/2013 - Jose Fantasia.
' Neuronal Orden de Compra - Abierta.
sub main
	If aObject.flag.descripcion = "Pendiente" Then Exit Sub
	if aObject.FechaEntrega = CDate("01/01/2000") then
	    AddFailedConstraint "No indicó la Fecha de Entrega Estimada.", 0 
	end if
	if aObject.centrocostos is nothing then
	    AddFailedConstraint "No indico Centro de Costo.", 0 
	end if
	if aObject.BOEXTENSION.FechaEntregaReal = CDate("01/01/2000") then
	    'AddFailedConstraint "No indicó la Fecha de Entrega Real.", 0 
	end if
	if aObject.BOEXTENSION.SolicitaMiniCC then
	    for each xItem in aObject.ItemsTransaccion
			if xItem.BOEXTENSION.MiniCC is nothing then
	    	    AddFailedConstraint "No indicó el Registro Repuesto para el ítem: " & xItem.NumeroItem, 2
			end if
		next 
	end if

	'26/09/23 validación de Ret. Gan. Bienes de Cambio y Uso para productos
    for each xItem in aObject.ItemsTransaccion
		If UCASE(classname(xItem.Referencia)) = "PRODUCTO" Then
		   set xImp = Nothing
		   Set xImp = GetPosicionImpuesto(xItem.referencia, "Ret. Gan. Bienes de Cambio y Uso")
		   If Not xImp Is Nothing Then
	          pos = xImp.PosicionImpuesto.name
		      if pos = "No Gravado"  then
	             AddFailedConstraint "Producto " & xItem.referencia.descripcion & " tiene Ret. Gan. Bienes de Cambio y Uso como No Gravado" & pos, 0
              end if
		   Else
		      AddFailedConstraint "Producto " & xItem.referencia.descripcion & " no tiene cargada posicion de impuesto Ret. Gan. Bienes de Cambio y Uso" & pos, 0
		   End If
        End If
		
	next
end sub