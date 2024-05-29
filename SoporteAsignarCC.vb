' Fecha Creación: 2024/05/24 
' Descripcion: Se agrego el campo CC a la UD, por ende completo todos los 
' User: Rodrigo Fierro
sub main
	stop
	' Para escoger varios ítems del View.
	set xView = NewCompoundView(Self, "UD_INTERCAMBIOEPISTOLAR", Self.Workspace, nil, true)
	xView.NoFlushBuffers = true		' WITH(NOLOCK)

	contador = 0
	For Each xItem in xView.viewitems
        if not xItem.BO.empleado is nothing then
            xItem.BO.centrocostos = xItem.BO.empleado.centrocostos
            If xItem.BO.Workspace.InTransaction Then
                xItem.BO.Workspace.Commit
            End If

            contador = contador + 1
            Call Esperar (3)
        end if

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