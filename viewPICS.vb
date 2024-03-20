' Creado: 13/03/2013 - José Fantasia.
' Vista de Orden de Compra.
sub main
	stop
	' Empleado del Usuario.


	set xEmpleado = nothing
	set xViewEmpl = NewCompoundView(Self, "EMPLEADO", Self.Workspace, nil, true)
	xViewEmpl.AddJoin(NewJoinSpec(xViewEmpl.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_EMPLEADO", "ID", ""), false))
	xViewEmpl.AddJoin(NewJoinSpec(NewColumnSpec("UD_EMPLEADO", "Usuario", ""), NewColumnSpec("USUARIO", "ID", ""), false))
	xViewEmpl.AddFilter(NewFilterSpec(xViewEmpl.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xViewEmpl.AddFilter(NewFilterSpec(NewColumnSpec("USUARIO", "NOMBRE", ""), " = ", NombreUsuario()))
	xViewEmpl.AddOrderColumn(NewOrderSpec(xViewEmpl.ColumnFromPath("CODIGO"), false))
	if xViewEmpl.ViewItems.Size > 0 then
		set xEmpleado = xViewEmpl.ViewItems.First.Current.BO
	end if
	
    View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "FLAG", ""), NewColumnSpec("FLAG", "ID", ""), True))
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "CENTROCOSTOS", ""), NewColumnSpec("CENTROCOSTOS", "ID", ""), True))
	View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "BOEXTENSION", ""), NewColumnSpec("UD_ORDENCOMPRA", "ID", ""), True))
	'View.AddJoin(NewJoinSpec(NewColumnSpec("TRORDENCOMPRA", "ID", ""), NewColumnSpec("PENDIENTE", "TRANSACCION", ""), False))
	'View.AddFilter(NewFilterSpec(NewColumnSpec("PENDIENTE", "RELACIONTRORIG", ""), " = ", "{967DD785-8E31-4864-9160-C6EA0D73D355}"))
	View.AddJoin(NewJoinSpec(NewColumnSpec("UD_ORDENCOMPRA", "EMPLEADO", ""), NewColumnSpec("EMPLEADO", "ID", ""), True))
	View.AddJoin(NewJoinSpec(NewColumnSpec("EMPLEADO", "ENTEASOCIADO", ""), NewColumnSpec("PERSONAFISICA", "ID", ""), True))
    
	Set xCol        = View.AddColumn(NewColumnSpec("TRORDENCOMPRA", "NOMBREDESTINATARIO", ""))
    xCol.Caption    = "Proveedor"
    xCol.Browse     = True
    Set xCol        = View.AddColumn(NewColumnSpec("CENTROCOSTOS", "NOMBRE", ""))
    xCol.Caption    = "Centro de Costos"
    xCol.Browse     = True
    Set xCol        = View.AddBOCol("VALORTOTAL")
    xCol.Caption    = "Monto Total"
    xCol.Browse     = True
    Set xCol        = View.AddBOCol("ESTADO")
    xCol.Caption    = "Estado"
    xCol.Browse     = True
    Set xCol        = View.AddColumn(NewColumnSpec("FLAG", "DESCRIPCION", ""))
    xCol.Caption    = "Flag"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "VERIFICADA", ""))
    xCol.Caption    = "Verificada (F)"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "PAGOSEMANAL", ""))
    xCol.Caption    = "Pago Semanal"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "PENDIENTEENTREGA", ""))
    xCol.Caption    = "Con Pendiente de Entrega"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "SOLICITANTEPIC_N", ""))
    xCol.Caption    = "Responsable OC"
    xCol.Browse     = True
	Set xCol        = View.AddBOCol("USUARIO")
    xCol.Caption    = "Usuario"
    xCol.Browse     = True
    Set xCol        = View.AddBOCol("Nota")
    xCol.Caption    = "Observaciones"
    xCol.Browse     = True
    Set xCol        = View.AddBOCol("NUMERODOCUMENTO")
    xCol.Caption    = "Número"
    xCol.Browse     = True
    Set xCol        = View.AddBOCol("FECHAACTUAL")
    xCol.Caption    = "Fecha"
    xCol.Browse     = True
	Set xCol        = View.AddBOCol("FECHAENTREGA")
    xCol.Caption    = "Fecha Ent. Estimada"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "FECHAENTREGAREAL", ""))
    xCol.Caption    = "Fecha Ent. Real"
    xCol.Browse     = True
	'Set xCol        = View.AddColumn(NewColumnSpec("PENDIENTE", "SALDADA", ""))
    'xCol.Caption    = "Facturada"
    'xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "OBSERVAINTERNA", ""))
    xCol.Caption    = "Observación Interna"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "FECHREGIS", ""))
    xCol.Caption    = "Fecha Registración"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "FACTURA", ""))
    xCol.Caption    = "¿Se Factura?"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "COMPRADOR", ""))
    xCol.Caption    = "Comprador"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("PERSONAFISICA", "NOMBRE", ""))
    xCol.Caption    = "Supervisor"
    xCol.Browse     = True
	Set xCol        = View.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "ENVIADA", ""))
    xCol.Caption    = "PDF Enviado"
    xCol.Browse     = True
	
	if UCase(NombreUsuario()) <> "DESARROLLO3" AND UCase(NombreUsuario()) <> "RFIERRO" AND UCase(NombreUsuario()) <> "DESARROLLO4" and UCase(NombreUsuario()) <> "rfierro" AND ucase(nombreusuario()) <> "CDELCAMPILLO" AND ucase(nombreusuario()) <> "SCALIPSO"_
	    AND ucase(nombreusuario()) <> "PCORTEZ" AND ucase(nombreusuario()) <> "HYORIO" AND ucase(nombreusuario()) <> "CLANGE" then
	if xEmpleado is nothing then	' Filtro para que no vea nada.
		View.AddFilter(NewFilterSpec(View.ColumnFromPath("ESTADO"), " = ", "X"))
	else
		if xEmpleado.BOEXTENSION.CENTROSVERTODOS = false then
			' Sólo ve los Centros que tiene permitido.
			contador = 1
			for each iCC in xEmpleado.BOEXTENSION.CENTROSVERLISTA
				set xFiltro = NewFilterSpec(View.ColumnFromPath("CENTROCOSTOS"), " = ", iCC.ID)
				if contador = 1 then
					xFiltro.BeginBlock 	= "("
				else
					xFiltro.Conector 	= "OR"
				end if
				View.AddFilter(xFiltro)
				contador = contador + 1
			next
			xFiltro.EndBlock = ")"
			View.AddFilter(xFiltro)
		end if
	end if
	end if
	
	View.AddOrderColumn(NewOrderSpec(NewColumnSpec("TRORDENCOMPRA", "NUMERODOCUMENTO", ""), True))
end sub
