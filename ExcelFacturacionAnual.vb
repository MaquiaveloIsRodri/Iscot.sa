' Creado: 26/09/2013 - Jose Fantasia.
' MU Excel. Exportar Comparacion de Facturacion.
sub main
	stop
	' VisualVar.
	set xVisualVar = VisualVarEditor("Comparación de Facturación")
	call AddVarInteger(xVisualVar, "00MES", "Mes", "Ingrese los siguientes datos:", Month(Date))
	call AddVarInteger(xVisualVar, "05ANIO", "Año", "Ingrese los siguientes datos:", Year(Date))
	aceptar = ShowVisualVar(xVisualVar)
	if aceptar then
		mes 	= CInt(GetValueVisualVar(xVisualVar, "00MES", "Ingrese los siguientes datos:"))
		anio 	= CInt(GetValueVisualVar(xVisualVar, "05ANIO", "Ingrese los siguientes datos:"))
		if mes < 1 or mes > 12 then
			MsgBox "El Mes ingresado es incorrecto.", 48, "Advertencia"
			exit sub
		end if
		if anio < 2000 then
			MsgBox "El Año ingresado es incorrecto.", 48, "Advertencia"
			exit sub
		end if
		
		' Conexion.
		set xCone2 = CreateObject("adodb.connection")
		xCone2.ConnectionString 	= StringConexion("calipso", Self.Workspace)
		xCone2.ConnectionTimeout 	= 150
		
		set xRstF = RecordSet(xCone2, "select top 1 * from producto")
		xRstF.Close
		xRstF.ActiveConnection.CommandTimeout = 0

		query = "exec SP_Ventas_ComparacionFacturacion " & anio & ", " & mes
		
		' --- EXCEL ---.
		call ProgressControl(Self.Workspace, "COMPARACION DE FACTURACION" , 0, 50)
		
		Dim HojaExcel
		set HojaExcel = CreateObject("excel.application")
		HojaExcel.Workbooks.Add

		' Cabecera.
		HojaExcel.ActiveSheet.Cells(1, 1).Value = "ISCOT SERVICES S.A."
		HojaExcel.ActiveSheet.Cells(2, 1).Value = "COMPARACION DE FACTURACION"
		HojaExcel.ActiveSheet.Cells(3, 1).Value = "MES: " & mes
		for ind = 1 to 3	' Hago un fondo gris y negrita a las primeras 4 filas y 5 columnas.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(ind, 1), HojaExcel.ActiveSheet.Cells(ind, 6)).Font.Bold 			= true
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(ind, 1), HojaExcel.ActiveSheet.Cells(ind, 6)).Interior.ColorIndex 	= 43
		next
		
		' Format.
		HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "$ #,##0.00"
		HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "$ #,##0.00"
		HojaExcel.ActiveSheet.Columns(4).NumberFormat	= "$ #,##0.00"
		
		txt_mes 	= ""
		txt_mesAnt 	= ""
		select case mes
		case 1
			txt_mes		= "ENERO"
			txt_mesAnt 	= "DICIEMBRE"
		case 2
			txt_mes		= "FEBRERO"
			txt_mesAnt 	= "ENERO"
		case 3
			txt_mes		= "MARZO"
			txt_mesAnt 	= "FEBRERO"
		case 4
			txt_mes		= "ABRIL"
			txt_mesAnt 	= "MARZO"
		case 5
			txt_mes		= "MAYO"
			txt_mesAnt 	= "ABRIL"
		case 6
			txt_mes		= "JUNIO"
			txt_mesAnt 	= "MAYO"
		case 7
			txt_mes		= "JULIO"
			txt_mesAnt 	= "JUNIO"
		case 8
			txt_mes		= "AGOSTO"
			txt_mesAnt 	= "JULIO"
		case 9
			txt_mes		= "SEPTIEMBRE"
			txt_mesAnt 	= "AGOSTO"
		case 10
			txt_mes		= "OCTUBRE"
			txt_mesAnt 	= "SEPTIEMBRE"
		case 11
			txt_mes		= "NOVIEMBRE"
			txt_mesAnt 	= "OCTUBRE"
		case else
			txt_mes		= "DICIEMBRE"
			txt_mesAnt 	= "NOVIEMBRE"
		end select
		
		HojaExcel.ActiveSheet.Cells(5, 1).Value = "CLIENTE (CC)"
		HojaExcel.ActiveSheet.Cells(5, 2).Value = "Facturación " & txt_mesAnt
		HojaExcel.ActiveSheet.Cells(5, 3).Value = "Facturación " & txt_mes
		HojaExcel.ActiveSheet.Cells(5, 4).Value = "Diferencia"
		HojaExcel.ActiveSheet.Cells(5, 5).Value = "Variación"
		HojaExcel.ActiveSheet.Cells(5, 6).Value = "Observación en la facturación respecto al mes anterior"
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 6)).Font.Bold 			= true
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 6)).Interior.ColorIndex 	= 43
		
		' Detalle.
		R = 7
		total 		= 0
		totalAnt 	= 0
		
		xRstF.Source = query
		xRstF.Open
		do while not xRstF.EOF
			call ProgressControlAvance(Self.Workspace, "Procesando. Por favor espere!. CC:" & CStr(xRstF("NombreCC").Value))
			
			HojaExcel.ActiveSheet.Cells(R, 1).Value = CStr(xRstF("NombreCC").Value)
			HojaExcel.ActiveSheet.Cells(R, 2).Value = CDbl(xRstF("TotalCCAnterior").Value)
			HojaExcel.ActiveSheet.Cells(R, 3).Value = CDbl(xRstF("TotalCC").Value)
			totalAnt	= totalAnt + CDbl(xRstF("TotalCCAnterior").Value)
			total		= total + CDbl(xRstF("TotalCC").Value)
			diferencia 	= CDbl(xRstF("TotalCCAnterior").Value) - CDbl(xRstF("TotalCC").Value)
			HojaExcel.ActiveSheet.Cells(R, 4).Value = Abs(CDbl(diferencia))
			if diferencia <> 0 then
				if diferencia > 0 then
					HojaExcel.ActiveSheet.Cells(R, 5).Interior.ColorIndex = 3
				else
					HojaExcel.ActiveSheet.Cells(R, 5).Interior.ColorIndex = 4
				end if
			end if
			
			R = R + 1
			xRstF.MoveNext
		loop
		
		' Totales.
		R = R + 1
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 6)).Font.Bold 			= true
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 6)).Interior.ColorIndex 	= 43
		
		HojaExcel.ActiveSheet.Cells(R, 1).Value = "TOTAL"
		HojaExcel.ActiveSheet.Cells(R, 2).Value = CDbl(totalAnt)
		HojaExcel.ActiveSheet.Cells(R, 3).Value = CDbl(total)
		diferencia = totalAnt - total
		HojaExcel.ActiveSheet.Cells(R, 4).Value = Abs(CDbl(diferencia))
		if diferencia <> 0 then
			if diferencia > 0 then
				HojaExcel.ActiveSheet.Cells(R, 5).Interior.ColorIndex = 3
			else
				HojaExcel.ActiveSheet.Cells(R, 5).Interior.ColorIndex = 4
			end if
		end if
		
		' Cabecera Maestros.
		R = R + 4
		HojaExcel.ActiveSheet.Cells(R, 1).Value = "RESUMEN MAESTROS"
		HojaExcel.ActiveSheet.Cells(R, 2).Value = "Facturación " & txt_mesAnt
		HojaExcel.ActiveSheet.Cells(R, 3).Value = "Facturación " & txt_mes
		HojaExcel.ActiveSheet.Cells(R, 4).Value = "Diferencia"
		HojaExcel.ActiveSheet.Cells(R, 5).Value = "Variación"
		HojaExcel.ActiveSheet.Cells(R, 6).Value = "Observación en la facturación respecto al mes anterior"
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 6)).Font.Bold 			= true
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 6)).Interior.ColorIndex 	= 43
		R = R + 2
		
		' Detalle Maestros.
		set xView = NewCompoundView(Self, "CENTROCOSTOS", Self.Workspace, nil, true)
		xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_CENTROCOSTOS", "ID", ""), false))
		xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
		xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("CENTROCOSTOS"), " = ", "{FEC5CA68-DA85-4722-88DC-8C9CC461167D}"))
		xView.AddFilter(NewFilterSpec(NewColumnSpec("UD_CENTROCOSTOS", "PRORRATEABLE", ""), " = ", true))
		xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NOMBRE"), false))
		
		for each xCC in xView.ViewItems
			call ProgressControlAvance(Self.Workspace, "Procesando. Por favor espere!. MAESTRO:" & xCC.BO.Nombre)
			aux_totalMaestro 		= 0
			aux_totalMaestroAnt 	= 0
			
			set xRstIng = RecordSet(xCone2, "select top 1 * from producto")
			xRstIng.Close
			xRstIng.ActiveConnection.commandTimeout = 0
			xRstIng.Source = QueryFacturadoCC(anio, mes, xCC.BO)
			xRstIng.Open
			do while not xRstIng.EOF
				aux_totalMaestro 	= CDbl(xRstIng("TotalMaestro").Value)
				aux_totalMaestroAnt = CDbl(xRstIng("TotalMaestroAnterior").Value)
				xRstIng.MoveNext
			loop
			
			HojaExcel.ActiveSheet.Cells(R, 1).Value = xCC.BO.Nombre
			HojaExcel.ActiveSheet.Cells(R, 2).Value = CDbl(aux_totalMaestroAnt)
			HojaExcel.ActiveSheet.Cells(R, 3).Value = CDbl(aux_totalMaestro)
			diferencia = CDbl(aux_totalMaestroAnt) - CDbl(aux_totalMaestro)
			HojaExcel.ActiveSheet.Cells(R, 4).Value = Abs(CDbl(diferencia))
			if diferencia <> 0 then
				if diferencia > 0 then
					HojaExcel.ActiveSheet.Cells(R, 5).Interior.ColorIndex = 3
				else
					HojaExcel.ActiveSheet.Cells(R, 5).Interior.ColorIndex = 4
				end if
			end if
			
			R = R + 1
		next
		
		' --- FIN EXCEL ---.
		HojaExcel.ActiveSheet.Columns("A:F").AutoFit
		call ProgressControlFinish(Self.Workspace)
		HojaExcel.Visible = true
		set HojaExcel = nothing
	end if
end sub

function QueryFacturadoCC(pAnio, pMes, pCentroCostos)
	filtro = "and cc.ID in ("
	for each xItemCC in pCentroCostos.BOEXTENSION.CCPRORRATEABLES
		filtro = filtro & "'" & xItemCC.CENTROCOSTOS.ID & "', "
	next
	filtro = filtro & "'" & pCentroCostos.ID & "', "
	filtro = Left(filtro, Len(filtro) - 2)
	filtro = filtro & ") "
		
	query = "select ISNULL(SUM(TotalMaestro), 0) as TotalMaestro, ISNULL(SUM(TotalMaestroAnterior), 0) as TotalMaestroAnterior " _
		& "from ( " _
		& "select ISNULL(SUM(item.HABER2_IMPORTE - item.DEBE2_IMPORTE), 0) as TotalMaestro, " _
		& "( " _
		& "	select ISNULL(SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE), 0) " _
		& "	from V_ITEMCONTABLE item2 with(nolock) " _
		& "	where item2.CENTROCOSTOS_ID = cc.ID " _
		& "	and item2.ESTADOTR = 'C' " _
		& "	and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}' " _
		& "	and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = (case when " & pMes & " = 1 then 12 else " & pMes & " - 1 end) " _
		& "	and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = (case when " & pMes & " = 1 then " & pAnio & " - 1 else " & pAnio & " end) " _
		& "	and item2.REFERENCIA_ID IN ( " _
		& "		select ID " _
		& "		from CUENTA with(nolock) " _
		& "		where ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}' " _
		& "		and ACTIVESTATUS = 0 " _
		& "	) " _
		& ") as TotalMaestroAnterior " _
		& "from V_ITEMCONTABLE item with(nolock) " _
		& "inner join V_CUENTA cta with(nolock) on cta.ID = item.REFERENCIA_ID " _
		& "inner join V_CENTROCOSTOS cc with(nolock) on cc.ID = item.CENTROCOSTOS_ID " _
		& "where item.ESTADOTR = 'C' " _
		& "and item.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}' " _
		& "and CAST(SUBSTRING(item.FECHAVENCIMIENTO, 5, 2) as Int) = " & pMes & " " _
		& "and CAST(LEFT(item.FECHAVENCIMIENTO, 4) as Int) = " & pAnio & " " _
		& "and cta.ID in ( " _
		& "		select ID " _
		& "		from V_CUENTA with(nolock) " _
		& "		where ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}' " _
		& "		and ACTIVESTATUS = 0 " _
		& ") " _
		& filtro _
		& "group by cc.ID, cta.CODIGO, cta.DESCRIPCION " _
		& ") Q"
	QueryFacturadoCC = query	
end function
