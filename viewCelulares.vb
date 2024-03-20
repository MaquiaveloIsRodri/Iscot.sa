Sub Main
  stop
  view.addjoin(NewJoinSpec( view.ColumnFromPath( "ORIGINANTE"), NewColumnSpec( "EMPLEADO", "id", "EMPLEADO" ), True ))
  View.AddJoin(NewJoinSpec(NewColumnSpec("UD_HISTORICOCELULAR", "CENTROCOSTOSANTERIOR", ""), NewColumnSpec("CENTROCOSTOS", "ID", ""), true))
  View.AddJoin(NewJoinSpec(NewColumnSpec("UD_HISTORICOCELULAR", "CENTROCOSTOSNUEVO", ""), NewColumnSpec("CENTROCOSTOS", "ID", ""), true)) 
  View.AddJoin(NewJoinSpec(NewColumnSpec("UD_HISTORICOCELULAR", "SERVICIOS", ""), NewColumnSpec("SERVICIO", "ID", ""), true))
  View.AddJoin(NewJoinSpec(NewColumnSpec("UD_HISTORICOCELULAR", "SERVICIOS", ""), NewColumnSpec("SERVICIO", "ID", ""), true))


  '-------------------------------------------------------------------------------------
  Set datadef = view.AddBOCol("FECHAHASTA")
  datadef.caption = "Fecha Hasta"
  datadef.browse = True
  'responsable--------------------------------------------------------------------------
  Set datadef = view.addcolumn(NewColumnSpec( "ORIGINANTE", "descripcion", "" ))
  datadef.caption = "Originante"
  datadef.browse = true

  'centro de costos anterior------------------------------------------------------------
  Set datadef = view.AddColumn (NewColumnSpec("CENTROCOSTOSANTERIOR", "NOMBRE", ""))
  datadef.caption = "Centro Costos Anterior"
  datadef.browse = True

  'centro de costos nuevo---------------------------------------------------------------
  Set datadef = view.AddColumn (NewColumnSpec("CENTROCOSTOSNUEVO", "NOMBRE", ""))
  datadef.caption = "Centro Costos Nuevo"
  datadef.browse = True

  'servicios Anterior-------------------------------------------------------------------
  Set datadef = view.AddColumn (NewColumnSpec("SERVICIOANTERIOR", "DESCRIPCION", ""))
  datadef.caption = "Plan Anterior"
  datadef.browse = True

  'servicios Anterior-------------------------------------------------------------------
  Set datadef = view.AddColumn (NewColumnSpec("SERVICIONUEVO", "DESCRIPCION", ""))
  datadef.caption = "Plan Nuevo"
  datadef.browse = True

  '-------------------------------------------------------------------------------------
  Set datadef = view.AddBOCol("MOTIVO_N")
  datadef.caption = "Motivo"
  datadef.browse = True


  view.searchdatadefname = "NOMBRE"
End Sub
