' 01/09/2017 - José Luis Fantasia.
' MU Facturacion Reporte Pendientes.
sub main
	stop
	
	set xCon = CreateObject("adodb.connection")
	xCon.ConnectionString 	= StringConexion("CALIPSO", Self.Workspace)
	xCon.ConnectionTimeout 	= 150
	
	set xRs = RecordSet(xCon, "select top 1 * from producto")
	xRs.Close
	xRs.ActiveConnection.CommandTimeout = 0
	xRs.Source = "exec SP_OrdenServicio_Pendientes"
	xRs.Open
	
	call ProgressControl(Self.Workspace, "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES", 0, 780)
	SendDebug "Inicio Pendientes"
	
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
			HojaExcel.ActiveSheet.Cells(1, 1).Value 		= "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES"
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

		' Colores.
		' Fecha Servicio.
		dias			= DateDiff("d", CDate(xRs("FECHASERVICIO").Value), Now)
		if dias > 30 then
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 1)).Interior.ColorIndex 		= 3
		end if
		' Fecha Estado.
		if CDate(xRs("FECHAESTADO").Value) <> CDate("01/01/2000") then
			dias			= DateDiff("d", CDate(xRs("FECHAESTADO").Value), Now)
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
	
	' Ordena la Hojas Alfabéticamente.
	For a = 1 To HojaExcel.Sheets.Count
		For s = a + 1 To HojaExcel.Sheets.Count
			If UCase(HojaExcel.Sheets(a).Name) > UCase(HojaExcel.Sheets(s).Name) Then
				HojaExcel.Sheets(s).Move HojaExcel.Sheets(a)
			End If
		Next
	Next



	'Resumen
	call ProgressControlAvance(Self.Workspace, "Resumen")
	nroHoja = nroHoja + 1
	HojaExcel.Sheets.Add
	HojaExcel.Sheets("Hoja" & CStr(nroHoja)).Name = "Resumen"
	HojaExcel.Sheets("Resumen").Move HojaExcel.Sheets(1)
	HojaExcel.Sheets("Resumen").Select
	' Format.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1000, 30)).Font.Size 		= 10

	HojaExcel.ActiveSheet.Columns("A").ColumnWidth 	= 15
	HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "$ #,##0.00"

	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value 		= "ORDENES DE SERVICIO Y CRÉDITO PENDIENTES"
	HojaExcel.ActiveSheet.Cells(2, 1).Value 		= "RESUMEN"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 2)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 2)).Interior.ColorIndex 		= 10 	' Fondo.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 2)).Font.Color				= vbWhite

	' Columnas.
	HojaExcel.ActiveSheet.Cells(5, 1).Value 		= "Centro de Costos"
	HojaExcel.ActiveSheet.Cells(5, 2).Value 		= "Total"
	HojaExcel.ActiveSheet.Cells(5, 3).Value 		= "Mes Actual"
	HojaExcel.ActiveSheet.Cells(5, 4).Value 		= "2do Mes"
	HojaExcel.ActiveSheet.Cells(5, 5).Value 		= "3r Mes"
	HojaExcel.ActiveSheet.Cells(5, 6).Value 		= "Resto de meses"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5,6)).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5,6)).Interior.ColorIndex 		= 43 	' Fondo.
	HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(4).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(5).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(6).NumberFormat	= "$ #,##0.00"

	call CargaExcelManual( Self.Workspace, "exec SP_OrdenServicio_Pendientes_Totales", HojaExcel, 6, "SinCabe" )
	HojaExcel.ActiveSheet.Columns("A:B").Autofit

	xRs.Close
	xRs.ActiveConnection.CommandTimeout = 0
	xRs.Source = "exec SP_OrdenServicio_Pendientes"
	xRs.Open
	MontoActual = 0.0
	Monto2da = 0.0
	Monto3ra = 0.0
	resto = 0.0
	total = 0.0
	Cc	= ""
	mes2Da = month(Date)-1
	mes3r = month(Date)-2
	if mes2Da < 10 then mes2Da = 0 & mes2Da
	if mes3r < 10 then mes3r = 0 & mes3r
	I = 6
	relative    = 0.0
ban         = true 
do while not xRs.EOF
    if cc <> xRs("CENTROCOSTOS").Value then
        if cc <> "" then
            ban = false
        end if
        cc          = xRs("CENTROCOSTOS").Value
	end if

	Select Case xRs("PERIODO").Value
		case year(date) & month(Date)
            if ban then
				MontoActual = MontoActual + CDbl(xRs("TOTAL").Value) + relative
				total       = total + MontoActual
                xRs.MoveNext
            end if
                relative = CDbl(xRs("TOTAL").Value)
		case year(date) & Mes2Da
            if ban then
            	Monto2da    = Monto2da + CDbl(xRs("TOTAL").Value) + relative
				total       = total + Monto2da
                xRs.MoveNext
            end if
			relative = CDbl(xRs("TOTAL").Value)

		case year(date) & Mes3r
            if ban then
				Monto3ra = Monto3ra + CDbl(xRs("TOTAL").Value) + relative
				total    = total + Monto3ra
                xRs.MoveNext
            end if
            relative = CDbl(xRs("TOTAL").Value)
		Case Else
            if ban then
				resto = resto + CDbl(xRs("TOTAL").Value) + relative
				total = total + resto
                xRs.MoveNext
            end if
			relative = CDbl(xRs("TOTAL").Value)
	End Select

    relative = 0.0
    ban      = true
	HojaExcel.ActiveSheet.Cells(I, 3).Value 		= MontoActual
	HojaExcel.ActiveSheet.Cells(I, 4).Value 		= Monto2da
	HojaExcel.ActiveSheet.Cells(I, 5).Value 		= Monto3ra
	HojaExcel.ActiveSheet.Cells(I, 6).Value 		= resto
    MontoActual = 0.0
	Monto2da 	= 0.0
	Monto3ra 	= 0.0
	resto 		= 0.0
	I = I + 1
    xRs.MoveNext
loop
	I = I + 1
	HojaExcel.ActiveSheet.Cells(I, 1).Value 		= "Total"
	HojaExcel.ActiveSheet.Cells(I, 2).Value 		= total
	HojaExcel.ActiveSheet.Cells(I, 1).Font.Bold 				= true 	' Negrita.
	HojaExcel.ActiveSheet.Cells(I, 1).Interior.ColorIndex 		= 43 	' Fondo.
	SendDebug "FIN Pendientes!!!!"
	call ProgressControlFinish(Self.Workspace)
	HojaExcel.Visible 	= true
	set HojaExcel 		= nothing
end sub
