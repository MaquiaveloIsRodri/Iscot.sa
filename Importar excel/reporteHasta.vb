' 02/03/2023 - MU Reporte Pendientes Hasta.
sub main
	stop
	' VisualVar.
	set xVisualVar = VisualVarEditor("REPORTE PENDIENTES HASTA")
	call AddVarDate(xVisualVar, "05FECHAHASTA", "Fecha Hasta", "Parametros", Now)
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then	exit sub
	
	fechaHasta 		= CDate(Int(GetValueVisualVar(xVisualVar, "05FECHAHASTA", "Parametros")))
	fechaHastaStr	= Year(fechaHasta) & Right("00" & Month(fechaHasta), 2) & Right("00" & Day(fechaHasta), 2)
	
	query = "exec SP_OrdenServicio_PendientesHasta '" & fechaHastaStr & "'"
	set xRs	= ConsultarSQL(query, Self.Workspace)
	
	call ProgressControl(Self.Workspace, "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES HASTA", 0, 650)
	SendDebug "Inicio Pendientes Hasta"
	
	' Excel.
	set HojaExcel = CreateObject("Excel.Application")
	HojaExcel.Workbooks.Add
	
	nroHoja = 0
	cc		= ""
	mes		= ""
	tot_mes = 0.0
	tot_cc	= 0.0
	
	do while not xRs.EOF
		call ProgressControlAvance(Self.Workspace, "CC: " & xRs("CENTROCOSTOS").Value & vbNewLine & "OS: " & xRs("OS").Value)
		
		if cc <> xRs("CENTROCOSTOS").Value then
			if cc <> "" then
				HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "Subtotal"
				HojaExcel.ActiveSheet.Cells(R, 5).Value 		= tot_mes
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
				R = R + 1
				HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "TOTAL"
				HojaExcel.ActiveSheet.Cells(R, 5).Value 		= tot_cc
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
			end if
			
			cc 		= xRs("CENTROCOSTOS").Value
			mes		= ""
			tot_mes = 0.0
			tot_cc	= 0.0
			HojaExcel.ActiveSheet.Columns("B:Z").AutoFit
			
			nroHoja = nroHoja + 1
			if nroHoja > 3 then
				HojaExcel.Sheets.Add
			end if
			nomHoja = "Hoja" & CStr(nroHoja)
			HojaExcel.Sheets(nomHoja).Select
			nuevoNomHoja = xRs("CENTROCOSTOS").Value
			nuevoNomHoja = Replace(nuevoNomHoja, ":", "")
			nuevoNomHoja = Replace(nuevoNomHoja, "\", "")
			nuevoNomHoja = Replace(nuevoNomHoja, "/", "")
			nuevoNomHoja = Replace(nuevoNomHoja, "?", "")
			nuevoNomHoja = Replace(nuevoNomHoja, "*", "")
			nuevoNomHoja = Replace(nuevoNomHoja, "[", "")
			nuevoNomHoja = Replace(nuevoNomHoja, "]", "")
			nuevoNomHoja = Left(Trim(nuevoNomHoja), 31)
			if Len(nuevoNomHoja) = 0 then
				nuevoNomHoja = "Hoja" & CStr(nroHoja)
			end if
			HojaExcel.Sheets(nomHoja).Name = nuevoNomHoja
			
			' Format.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Name 		= "Calibri"
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Size 		= 10
			HojaExcel.ActiveSheet.Columns("A").ColumnWidth 	= 15
			HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "@"
			HojaExcel.ActiveSheet.Columns(5).NumberFormat	= "$ #,##0.00"
			
			' Cabecera.
			HojaExcel.ActiveSheet.Cells(1, 1).Value 		= "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES - HASTA: " & FormatDateTime(fechaHasta, 2)
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
		if mes <> "" and mes <> xRs("PERIODO").Value then
			HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "Subtotal"
			HojaExcel.ActiveSheet.Cells(R, 5).Value 		= tot_mes
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
			tot_mes	= 0.0
			R = R + 1
		end if
		
		mes = xRs("PERIODO").Value
		
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
		' 29/11/2023 - Estado de Orden.
		if xRs("ESTADOORDEN") = "01" then		' Interna.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 6
		end if
		' Fecha Servicio.
		dias = DateDiff("d", CDate(xRs("FECHASERVICIO").Value), Now)
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
		
		tot_mes 		= tot_mes + CDbl(xRs("TOTAL").Value)
		tot_cc			= tot_cc + CDbl(xRs("TOTAL").Value)
		R = R + 1
		xRs.MoveNext
	loop
	
	' Totales de la última hoja.
	HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "Subtotal"
	HojaExcel.ActiveSheet.Cells(R, 5).Value 		= tot_mes
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
	R = R + 1
	HojaExcel.ActiveSheet.Cells(R, 1).Value 		= "TOTAL"
	HojaExcel.ActiveSheet.Cells(R, 5).Value 		= tot_cc
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 		= 15 	' Fondo.
	
	HojaExcel.ActiveSheet.Columns("B:Z").AutoFit
	
	' -----------------
	' -- REFERENCIAS --
	' -----------------
	
	call ProgressControlAvance(Self.Workspace, "Armando Referencias...")
	
	nroHoja = nroHoja + 1
	HojaExcel.Sheets.Add
	HojaExcel.Sheets("Hoja" & CStr(nroHoja)).Name = "ZZ REFERENCIAS"
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
	
	HojaExcel.Sheets("ZZ REFERENCIAS").Name 	= "REFERENCIAS"
	SendDebug "FIN Pendientes Hasta!!!!"
	call ProgressControlFinish(Self.Workspace)
	HojaExcel.Visible 	= true
	set HojaExcel 		= nothing
end sub
