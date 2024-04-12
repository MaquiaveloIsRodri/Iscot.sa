sub main
	stop
	' Para escoger varios ítems del View.
	set xRs	= ConsultarSQL("SELECT * FROM V_INTERCAMBIOSFILTER", Self.Workspace)


	do while not xRs.EOF
		Set oNovedades  = ExisteBo(self, "TIPOCLASIFICADOR", "ID", "{724B7CD5-E6A1-40E7-8492-0B280DAC9131}", nil, True, false, "=")
        Set oNovedad    = ExisteBo(self, "ITEMTIPOCLASIFICADOR", "ID", "{E28EBB36-0497-43D6-9C4F-8504DEA9D7DC}" , oNovedades.Valores, True, False, "=") 'Entregada
		set intercambio = ExisteBo(self, "UD_INTERCAMBIOEPISTOLAR", "ID", xRs("ID").Value , nil, True, False, "=")
		intercambio.NOVEDADESINTERCAMBIO = oNovedad
		If intercambio.BO.Workspace.InTransaction Then
			intercambio.BO.Workspace.Commit
		End If
		Call Esperar (3)

		xRs.MoveNext
	loop

end sub

Sub Esperar(Tiempo)
	ComienzoTiempo = Timer 
	FinTiempo = ComienzoTiempo + Tiempo
	do while FinTiempo > Timer
		if ComienzoTiempo > Timer then 
			FinTiempo = FinTiempo - 24 * 60 * 60 
		end if 
	loop 
end Sub
