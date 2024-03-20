Sub Main
  stop
  view.addjoin(NewJoinSpec( view.ColumnFromPath( "Empleado" ), NewColumnSpec( "Responsable", "id", "" ), True ))
  view.addjoin(NewJoinSpec( view.ColumnFromPath( "centrocostos" ), NewColumnSpec( "Empleado", "centrocostos_id", "" ), True ))
  'Codigo--------------------------------------------------------------------------
  Set datadef = view.AddBOCol("CODIGO")
  datadef.caption = "Código"
  datadef.browse = True
  'Modelo------------------------------------------------------------------------- 
  Set datadef = view.addbocol("MARCAMODELO")
  datadef.caption = "Marca/modelo"
  datadef.browse = True
  'responsable--------------------------------------------------------------------------
  Set datadef = view.addbocol("RESPONSABLE_N")
  datadef.caption = "Responsable"
  datadef.browse = true
  'numero sim--------------------------------------------------------------------------
  Set datadef = view.AddBOCol("NUMEROSIM")
  datadef.caption = "Numero Sim"
  datadef.browse = True
  'Imei----------------------------------------------------------------------------
  Set datadef = view.AddBOCol ("IMEI")
  datadef.caption = "IMEI"
  datadef.browse = True

  view.searchdatadefname = "NOMBRE"
End Sub

