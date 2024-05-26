' 01/09/2017 - MU OrSer Reporte Pendientes - @josefantasia.
' 22/11/2023 - Agregamos el Estado y las Referencias.
sub main
	stop
	query 	= "exec SP_OrdenServicio_Pendientes"
	set xRs = ConsultarSQL(query, Self.Workspace)
	
	if xRs.EOF then
		MsgBox "No hay órdenes pendientes para mostrar.", 48, "Aviso"
		exit sub
	end if
	
	call ProgressControl(Self.Workspace, "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES", 0, 400)
	
	' Excel.
	set HojaExcel = CreateObject("Excel.Application")
	HojaExcel.Workbooks.Add
	
	numero_hoja	= 0
	centro		= ""
	periodo		= ""
	total_c		= 0.0
	total_p 	= 0.0
	
	do while not xRs.EOF
		call ProgressControlAvance(Self.Workspace, "CC: " & xRs("CENTROCOSTOS").Value & vbNewLine & "OS: " & xRs("OS").Value)
		
		if centro <> xRs("CENTROCOSTOS").Value then
			if centro <> "" then
				HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "Subtotal"
				HojaExcel.ActiveSheet.Cells(R, 5).Value 		= total_p
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
				R = R + 1
				HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "TOTAL"
				HojaExcel.ActiveSheet.Cells(R, 5).Value 		= total_c
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
			end if
			
			centro 		= xRs("CENTROCOSTOS").Value
			periodo		= ""
			total_p 	= 0.0
			total_c		= 0.0
			
			HojaExcel.ActiveSheet.Columns("B:Z").AutoFit
			
			' Mombre de Hoja.
			numero_hoja = numero_hoja + 1
			if numero_hoja > 3 then
				HojaExcel.Sheets.Add
			end if
			nombre_hoja = "Hoja" & CStr(numero_hoja)
			HojaExcel.Sheets(nombre_hoja).Select
			nueva_hoja 	= xRs("CENTROCOSTOS").Value
			nueva_hoja 	= Replace(nueva_hoja, ":", "")
			nueva_hoja 	= Replace(nueva_hoja, "\", "")
			nueva_hoja 	= Replace(nueva_hoja, "/", "")
			nueva_hoja 	= Replace(nueva_hoja, "?", "")
			nueva_hoja 	= Replace(nueva_hoja, "*", "")
			nueva_hoja 	= Replace(nueva_hoja, "[", "")
			nueva_hoja 	= Replace(nueva_hoja, "]", "")
			nueva_hoja 	= Left(Trim(nueva_hoja), 31)
			if Len(nueva_hoja) = 0 then
				nueva_hoja = "Hoja" & CStr(numero_hoja)
			end if
			HojaExcel.Sheets(nombre_hoja).Name = nueva_hoja
			
			' Format.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Name 		= "Calibri"
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Size 		= 10
			HojaExcel.ActiveSheet.Columns("A").ColumnWidth 	= 15
			HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "@"
			HojaExcel.ActiveSheet.Columns(5).NumberFormat	= "$ #,##0.00"
			
			' Cabecera.
			HojaExcel.ActiveSheet.Cells(1, 1).Value 		= "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES - " & FormatDateTime(Now, 2) & "  " & FormatDateTime(Now, 4)
			HojaExcel.ActiveSheet.Cells(2, 1).Value 		= "SERVICIO: " & xRs("CENTROCOSTOS").Value
			HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "RESP. DE FACTURACIÓN: " & xRs("RESPONSABLE").Value
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 13)).Font.Bold 				= true 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 13)).Interior.ColorIndex 		= 10 	' Fondo.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 13)).Font.Color				= vbWhite
			
			' Columnas.
			HojaExcel.ActiveSheet.Cells(5, 1).Value 		= "Periodo"
			HojaExcel.ActiveSheet.Cells(5, 2).Value 		= "Orden"
			HojaExcel.ActiveSheet.Cells(5, 3).Value 		= "Cliente"
			HojaExcel.ActiveSheet.Cells(5, 4).Value 		= "Sector"
			HojaExcel.ActiveSheet.Cells(5, 5).Value 		= "Total"
			HojaExcel.ActiveSheet.Cells(5, 6).Value 		= "Observaciones"
			HojaExcel.ActiveSheet.Cells(5, 7).Value 		= "Descripción"
			HojaExcel.ActiveSheet.Cells(5, 8).Value 		= "Tipo"
			HojaExcel.ActiveSheet.Cells(5, 9).Value 		= "Año"
			HojaExcel.ActiveSheet.Cells(5, 10).Value 		= "Solicitante"
			HojaExcel.ActiveSheet.Cells(5, 11).Value 		= "Autorizante"
			HojaExcel.ActiveSheet.Cells(5, 12).Value 		= "Fe. Estado"
			HojaExcel.ActiveSheet.Cells(5, 13).Value 		= "Observa. Estado"
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 13)).Font.Bold 				= true 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 13)).Interior.ColorIndex 		= 43 	' Fondo.
			
			R = 6
		end if
		if periodo <> "" and periodo <> xRs("PERIODO").Value then
			HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "Subtotal"
			HojaExcel.ActiveSheet.Cells(R, 5).Value 		= total_p
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
			total_p	= 0.0
			R = R + 1
		end if
		
		periodo = xRs("PERIODO").Value
		
		redim datos(12)
		datos(0)		= NombreMes(CDate(xRs("FECHASERVICIO").Value)) & "-" & Year(CDate(xRs("FECHASERVICIO").Value))
		datos(1)		= xRs("OS").Value
		datos(2)		= xRs("CLIENTE").Value
		datos(3)		= xRs("SECTOR").Value
		datos(4)		= CDbl(xRs("TOTAL").Value)
		datos(5)		= xRs("OBSERVACIONES").Value
		datos(6)		= xRs("DESCRIPCION").Value
		datos(7)		= xRs("TIPOSERVICIO").Value
		datos(8)		= xRs("ANIO").Value
		datos(9)		= xRs("SOLICITANTE").Value
		datos(10)		= xRs("RECLAMARA").Value
		datos(11)		= CDate(xRs("FECHAESTADO").Value)
		datos(12)		= xRs("OBSERVAESTADO").Value
		set rango 				= HojaExcel.ActiveSheet.Range("A" & R)
		rango.Resize(1, 13) 	= datos
		
		if CDate(xRs("FECHAESTADO").Value) = CDate("01/01/2000") then
			HojaExcel.ActiveSheet.Cells(R, 12).Value	= ""
		end if
		
		' -- COLORES --
		' 22/11/2023 - Estado de Orden.
		if xRs("ESTADOORDEN") = "01" then		' Interna.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 6
		end if
		' Fecha Servicio.
		' 21/03/2024 - Cambiamos a que se calcule con el último día del mes.
		' Antes: dias = DateDiff("d", CDate(xRs("FECHASERVICIO").Value), Now)
		ultimoDia 	= DateSerial(Year(CDate(xRs("FECHASERVICIO").Value)), Month(CDate(xRs("FECHASERVICIO").Value)) + 1, 1) - 1
		dias 		= DateDiff("d", ultimoDia, Now)
		if dias > 30 then
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 1)).Interior.ColorIndex 		= 3
		end if
		' Fecha Estado.
		if CDate(xRs("FECHAESTADO").Value) <> CDate("01/01/2000") then
			dias = DateDiff("d", CDate(xRs("FECHAESTADO").Value), Now)
			if dias > 15 then
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 12), HojaExcel.ActiveSheet.Cells(R, 12)).Interior.ColorIndex 	= 3
			end if
		end if
		
		total_p = total_p + CDbl(xRs("TOTAL").Value)
		total_c	= total_c + CDbl(xRs("TOTAL").Value)
		R = R + 1
		xRs.MoveNext
	loop

	' Totales de la última hoja.
	HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "Subtotal"
	HojaExcel.ActiveSheet.Cells(R, 5).Value 		= total_p
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
	R = R + 1
	HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "TOTAL"
	HojaExcel.ActiveSheet.Cells(R, 5).Value 		= total_c
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
	
	HojaExcel.ActiveSheet.Columns("B:Z").AutoFit
	
	' -------------
	' -- RESUMEN --
	' -------------
	
	call ProgressControlAvance(Self.Workspace, "Armando Resumen...")
	
	numero_hoja = numero_hoja + 1
	HojaExcel.Sheets.Add
	HojaExcel.Sheets("Hoja" & CStr(numero_hoja)).Name = "00 RESUMEN"
	HojaExcel.Sheets("00 RESUMEN").Select
	
	' Format.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Size 		= 10
	HojaExcel.ActiveSheet.Columns("A").ColumnWidth 	= 35
	HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(4).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(5).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(6).NumberFormat	= "$ #,##0.00"
	
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value 		= "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES"
	HojaExcel.ActiveSheet.Cells(2, 1).Value 		= "RESUMEN"
	HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "PERIODO ACTUAL: " & UCase(NombreMes(Now)) & "-" & Year(Now)
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Interior.ColorIndex 		= 10 	' Fondo.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Font.Color				= vbWhite
	
	' Columnas.
	HojaExcel.ActiveSheet.Cells(5, 1).Value 		= "Centro de Costos"
	HojaExcel.ActiveSheet.Cells(5, 2).Value 		= "Total"
	HojaExcel.ActiveSheet.Cells(5, 3).Value 		= NombreMes(Now) & "-" & Year(Now)
	HojaExcel.ActiveSheet.Cells(5, 4).Value 		= NombreMes(DateAdd("m", -1, Now)) & "-" & Year(DateAdd("m", -1, Now))
	HojaExcel.ActiveSheet.Cells(5, 5).Value 		= NombreMes(DateAdd("m", -2, Now)) & "-" & Year(DateAdd("m", -2, Now))
	HojaExcel.ActiveSheet.Cells(5, 6).Value 		= "Resto de meses"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 6)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 6)).Interior.ColorIndex 		= 43 	' Fondo.
	
	call CargaExcelManual(Self.Workspace, "exec SP_OrdenServicio_Pendientes_Totales2", HojaExcel, 6, "SinCabe")
	
	I = 6
	do while Trim(HojaExcel.ActiveSheet.Cells(I, 1).Value) <> ""
		I = I + 1
	loop
	
	HojaExcel.ActiveSheet.Cells(I, 1).Value 	= "Total"
	HojaExcel.ActiveSheet.Cells(I, 2).Formula	= "=SUM(B6:B" & I - 1 & ")"
    HojaExcel.ActiveSheet.Cells(I, 3).Formula	= "=SUM(C6:C" & I - 1 & ")"
    HojaExcel.ActiveSheet.Cells(I, 4).Formula	= "=SUM(D6:D" & I - 1 & ")"
    HojaExcel.ActiveSheet.Cells(I, 5).Formula	= "=SUM(E6:E" & I - 1 & ")"
    HojaExcel.ActiveSheet.Cells(I, 6).Formula	= "=SUM(F6:F" & I - 1 & ")"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(I, 1), HojaExcel.ActiveSheet.Cells(I, 6)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(I, 1), HojaExcel.ActiveSheet.Cells(I, 6)).Interior.ColorIndex 		= 43 	' Fondo.
	
	HojaExcel.ActiveSheet.Columns("B:Z").AutoFit

    ' ------------------------------------
	' -- Comparación de Facturación -----
	' ------------------------------------

	call ProgressControlAvance(Self.Workspace, "Armando Comparación de Facturación...")



	' Conexion.
	set xCone2 = CreateObject("adodb.connection")
	xCone2.ConnectionString 	= StringConexion("calipso", Self.Workspace)
	xCone2.ConnectionTimeout 	= 150

	set xRstF = RecordSet(xCone2, "select top 1 * from producto")
	xRstF.Close
	xRstF.ActiveConnection.CommandTimeout = 0
    mes = Month(Date)
    anio = Year(Date)

	query = "exec SP_Ventas_ComparacionFacturacion " & anio & ", " & mes


	numero_hoja = numero_hoja + 1
	HojaExcel.Sheets.Add
	HojaExcel.Sheets("Hoja" & CStr(numero_hoja)).Name = "00 Comparación Facturación"
	HojaExcel.Sheets("00 Comparación Facturación").Select
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "ISCOT SERVICES S.A."
	HojaExcel.ActiveSheet.Cells(2, 1).Value = "FACTURACION ANUAL"
	HojaExcel.ActiveSheet.Cells(3, 1).Value = "AÑO: " & anio
	for ind = 1 to 3	' Hago un fondo gris y negrita a las primeras 4 filas y 5 columnas.
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(ind, 1), HojaExcel.ActiveSheet.Cells(ind, 6)).Font.Bold 			= true
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(ind, 1), HojaExcel.ActiveSheet.Cells(ind, 6)).Interior.ColorIndex 	= 43
	next

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

	' -----------------
	' -- REFERENCIAS --
	' -----------------
	
	call ProgressControlAvance(Self.Workspace, "Armando Referencias...")
	
	numero_hoja = numero_hoja + 1
	HojaExcel.Sheets.Add
	HojaExcel.Sheets("Hoja" & CStr(numero_hoja)).Name = "ZZ REFERENCIAS"
	HojaExcel.Sheets("ZZ REFERENCIAS").Select
	
	' Format.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Size 		= 10
	HojaExcel.ActiveSheet.Columns("A").ColumnWidth 	= 15
	
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value 		= "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES"
	HojaExcel.ActiveSheet.Cells(2, 1).Value 		= "REFERENCIAS"
	HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "COLORES EN HOJAS DE DETALLES"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 2)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 2)).Interior.ColorIndex 		= 10 	' Fondo.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(3, 2)).Font.Color				= vbWhite
	
	' Columnas.
	HojaExcel.ActiveSheet.Cells(5, 1).Value 		= "Color"
	HojaExcel.ActiveSheet.Cells(5, 2).Value 		= "Descripción"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 2)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 2)).Interior.ColorIndex 		= 43 	' Fondo.
	
	' Referencias.
	HojaExcel.ActiveSheet.Cells(6, 1).Interior.ColorIndex 	= 6
	HojaExcel.ActiveSheet.Cells(6, 2).Value					= "Orden Interna."
	HojaExcel.ActiveSheet.Cells(7, 1).Interior.ColorIndex 	= 3
	HojaExcel.ActiveSheet.Cells(7, 2).Value					= "Fecha de Servicio mayor que 30 días."
	HojaExcel.ActiveSheet.Cells(8, 1).Interior.ColorIndex 	= 3
	HojaExcel.ActiveSheet.Cells(8, 2).Value					= "Fecha de Estado mayor que 15 días."
	
	HojaExcel.ActiveSheet.Columns("B").AutoFit
	
	' -- FIN --
	' Ordena la Hojas Alfabéticamente.
	For a = 1 To HojaExcel.Sheets.Count
		For s = a + 1 To HojaExcel.Sheets.Count
			If UCase(HojaExcel.Sheets(a).Name) > UCase(HojaExcel.Sheets(s).Name) Then
				HojaExcel.Sheets(s).Move HojaExcel.Sheets(a)
			End If
		Next
	Next
	
	HojaExcel.Sheets("00 RESUMEN").Name 		= "RESUMEN"
    HojaExcel.Sheets("00 Comparación Facturación").Name 		= "Comparación Facturación"
	HojaExcel.Sheets("ZZ REFERENCIAS").Name 	= "REFERENCIAS"
	call ProgressControlFinish(Self.Workspace)
	HojaExcel.Visible 	= true
	set HojaExcel 		= nothing
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