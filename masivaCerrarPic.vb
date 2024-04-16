sub main

    stop

    for each xPic in container
    	if xPic.flag is nothing then
            Call Msgbox("Transacción sin flag, contacte a sistemas",64, "Información")
		   Exit Sub		
        Else
            If xPic.flag.id = "{77A2D825-05AF-498A-B91E-814149A20AF0}" then ' cerrada
                Call Msgbox("La Baja " & xPic.nombre & " ya fué cerrada.",64, "Información")
                Exit sub
            End If

		end if
        call EjecutarTransicion( xPic , "Cerrar PIP Compra" )
    next
end sub
