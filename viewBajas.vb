SUB MAIN
	View.AddJoin(NewJoinSpec(View.ColumnFromPath("DESTINATARIO"), NewColumnSpec("EMPLEADO", "ID", ""), False))
	View.AddJoin(NewJoinSpec(View.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_BAJAPERSONAL", "ID", ""), True))
	View.AddJoin(NewJoinSpec(View.ColumnFromPath("FLAG"), NewColumnSpec("FLAG", "ID", ""), True))
	View.AddJoin(NewJoinSpec(View.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), True))
	View.AddFilter(NewFilterSpec(View.ColumnFromPath("ESTADO"), " <> ", "N"))

	


	Set xCol        = View.AddColumn(NewColumnSpec("EMPLEADO", "CODIGO", ""))
	xCol.Caption    = "Legajo"
    	xCol.Browse     = True
	
	Set datadef     = View.AddColumn(NewColumnSpec("EMPLEADO", "DESCRIPCION", ""))
 	datadef.caption = "Nombre"
 	datadef.browse  = True

	Set datadef     = View.AddColumn(NewColumnSpec("CENTROCOSTOS", "Nombre", ""))
 	datadef.caption = "Centro Costos"
 	datadef.browse  = True

	Set datadef     = View.AddColumn(NewColumnSpec("EMPLEADO", "Cuit", ""))
 	datadef.caption = "Cuit"
 	datadef.browse  = True

   	Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "FECHARECEPCION", "")) 
 	datadef.caption = "Fecha Recepcion"
 	datadef.browse  = True	

	Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "FECHABAJA", ""))
 	datadef.caption = "Fecha de Baja"
 	datadef.browse  = True	

	Set datadef     = View.AddColumn(NewColumnSpec("FLAG", "DESCRIPCION", ""))
 	datadef.caption = "Estado"
 	datadef.browse  = True	

   	Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "FECHAENVIO", "")) 
 	datadef.caption = "Fecha Envio"
 	datadef.browse  = True	

    
    	Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "DEPOSITALIQUIDACIONFINAL", ""))
 	datadef.caption = "Deposita LF"
 	datadef.browse  = True
    
    	Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "DESCUENTAINDUMENTARIA", "")) 
 	datadef.caption = "Desc. Indum."
 	datadef.browse  = True
    
    	Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "DESCUENTATARJETA", ""))  
 	datadef.caption = "Desc. Tarj"
 	datadef.browse  = True


	'View.AddOrderColumn(NewOrderSpec(NewColumnSpec ("UD_BAJAPERSONAL", "MOMENTO", ""), True))
	View.AddOrderColumn(NewOrderSpec(View.ColumnFromPath("MOMENTO"), True))
	

END SUB


Public Sub FillDetailViews( aBO, aResultContainer, aOriginalDetailViews )
  DIM xObj
  On Error Resume Next
  For Each xObj in aOriginalDetailViews
    if CheckDetailView( xObj ) then
      aResultContainer.Add( xObj )
    End if
  Next
    aResultContainer.Add(GetViewBajas(aBO))
End Sub



Private Function GetViewBajas(aBO)
	Set xView = Nothing
	stop
	If Not aBO.empleado Is Nothing Then		
		Set xView = NewCompoundView(aBO, "TRSOLICITUD", aBO.Workspace, Nil, True)
        xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_BAJAPERSONAL", "ID", ""), false))
        xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("FLAG"), NewColumnSpec("FLAG", "ID", ""), True))
        xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), True))
        xView.AddFilter(NewFilterSpec(NewColumnSpec("TRSOLICITUD", "DESTINATARIO", ""), " = ", aBO.empleado))
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ESTADO"), " <> ", "N"))

        Set xCol        = xView.AddColumn(NewColumnSpec("EMPLEADO", "CODIGO", ""))
        xCol.Caption    = "Legajo"
        xCol.Browse     = True
        
        Set datadef     = xView.AddColumn(NewColumnSpec("EMPLEADO", "DESCRIPCION", ""))
        datadef.caption = "Nombre"
        datadef.browse  = True

        Set datadef     = xView.AddColumn(NewColumnSpec("CENTROCOSTOS", "Nombre", ""))
        datadef.caption = "Centro Costos"
        datadef.browse  = True

        Set datadef     = View.AddColumn(NewColumnSpec("EMPLEADO", "Cuit", ""))
        datadef.caption = "Cuit"
        datadef.browse  = True

        Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "FECHARECEPCION", "")) 
        datadef.caption = "Fecha Recepcion"
        datadef.browse  = True	

        Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "FECHABAJA", ""))
        datadef.caption = "Fecha de Baja"
        datadef.browse  = True	

        Set datadef     = View.AddColumn(NewColumnSpec("FLAG", "DESCRIPCION", ""))
        datadef.caption = "Estado"
        datadef.browse  = True	

        Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "FECHAENVIO", "")) 
        datadef.caption = "Fecha Envio"
        datadef.browse  = True	

        
            Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "DEPOSITALIQUIDACIONFINAL", ""))
        datadef.caption = "Deposita LF"
        datadef.browse  = True
        
            Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "DESCUENTAINDUMENTARIA", "")) 
        datadef.caption = "Desc. Indum."
        datadef.browse  = True
        
            Set datadef     = View.AddColumn(NewColumnSpec("UD_BAJAPERSONAL", "DESCUENTATARJETA", ""))  
        datadef.caption = "Desc. Tarj"
        datadef.browse  = True


        'View.AddOrderColumn(NewOrderSpec(NewColumnSpec ("UD_BAJAPERSONAL", "MOMENTO", ""), True))
        View.AddOrderColumn(NewOrderSpec(View.ColumnFromPath("MOMENTO"), True))

	End If
	
	Set GetViewBajas = xView
End Function

