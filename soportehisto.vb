' SelectView.
sub main
	stop
	' Para escoger varios ítems del View.
	set xView = NewCompoundView(Self, "UD_CELULARES", Self.Workspace, nil, true)
	xView.NoFlushBuffers = true		' WITH(NOLOCK)
	'xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_TABLA", "ID", ""), false))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("HISTORICOCELULARES"), " IS ", NOTHING))
	'xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NOMBRE"), false))
	
	contador = 0
	For Each xItem in xView.viewitems
		linea = xItem.BO.NUMEROLINEA
		xItem.BO.NUMEROLINEA = "test"
		xItem.BO.NUMEROLINEA = linea
		If xItem.BO.Workspace.InTransaction Then
		   xItem.BO.Workspace.Commit
		End If
		
		contador = contador + 1
		Call Esperar (3)
		SENDDEBUG xItem.BO.HISTORICOCELULARES.Id
		
	Next

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
