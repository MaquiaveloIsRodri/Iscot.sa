' 25/09/2015 - José Luis Fantasia.
' MU Saldo a Cobrar Clientes.
sub main
	stop
	
	set xVisualVar = VisualVarEditor("SALDO A COBRAR")
	call AddVarDate(xVisualVar, "00FECHADESDE", "Fecha Desde", "Parametros", Date)
	call AddVarDate(xVisualVar, "05FECHAHASTA", "Fecha Hasta", "Parametros", Date)
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then exit sub
	
	fechaDesde 		= CDate(Int(GetValueVisualVar(xVisualVar, "00FECHADESDE", "Parametros")))
	fechaHasta 		= CDate(Int(GetValueVisualVar(xVisualVar, "05FECHAHASTA", "Parametros")))
	fechaDesdeStr	= Year(fechaDesde) & Right("00" & Month(fechaDesde), 2) & Right("00" & Day(fechaDesde), 2)
	fechaHastaStr	= Year(fechaHasta) & Right("00" & Month(fechaHasta), 2) & Right("00" & Day(fechaHasta), 2)
	
	if fechaDesde > fechaHasta then
		MsgBox "La Fecha Desde NO puede ser Mayor que la Fecha Hasta.", 48, "Aviso"
		exit sub
	end if
	
	set xCon = CreateObject("adodb.connection")
	xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
	xCon.ConnectionTimeout 	= 150
	
	set xRs = RecordSet(xCon, "select top 1 * from producto")
	xRs.Close
	xRs.ActiveConnection.CommandTimeout = 0
	
	call ProgressControl(Self.Workspace, "Saldo a Cobrar Clientes" , 0, 210)
	SendDebug "Inicio A Cobrar"
	
	' --EXCEL
	dim HojaExcel
	set HojaExcel = createObject("Excel.Application")
	HojaExcel.Workbooks.Add
	
	' -------------------------
	' -- Hoja1: COMPROBANTES --
	' -------------------------
	
	HojaExcel.Sheets("Hoja1").Name	= "COMPROBANTES"

	' Fuente.
	HojaExcel.ActiveSheet.Cells.Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Cells.Font.Size 		= 11
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "SALDO A COBRAR CLIENTES"
	HojaExcel.ActiveSheet.Cells(2, 1).Value = "Periodo: " & FormatDateTime(fechaDesde, 2) & " - " & FormatDateTime(fechaHasta, 2)
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 13)).Font.Bold 			= true 	   	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 13)).Font.Color			= vbWhite	' Blanco.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 13)).Interior.ColorIndex 	= 10 		' Fondo: Verde oscuro.
	' Columnas.
	HojaExcel.ActiveSheet.Cells(4, 1).Value 		= "Cliente"
	HojaExcel.ActiveSheet.Cells(4, 2).Value 		= "Comprobante"
	HojaExcel.ActiveSheet.Cells(4, 3).Value 		= "Saldo"
	HojaExcel.ActiveSheet.Cells(4, 4).Value		    = "Vencida"
	HojaExcel.ActiveSheet.Cells(4, 5).Value 		= "Tipo"
	HojaExcel.ActiveSheet.Cells(4, 6).Value 		= "Servicio"
	HojaExcel.ActiveSheet.Cells(4, 7).Value 		= "Días a Hoy"
	HojaExcel.ActiveSheet.Cells(4, 8).Value 		= "Centro de Costos"
	HojaExcel.ActiveSheet.Cells(4, 9).Value 		= "Seguimiento"
	HojaExcel.ActiveSheet.Cells(4, 10).Value 		= "Fe. Estado"
	HojaExcel.ActiveSheet.Cells(4, 11).Value 		= "Observa. Estado"
    HojaExcel.ActiveSheet.Cells(4, 12).Value 		= "No vencido"
    HojaExcel.ActiveSheet.Cells(4, 13).Value 		= "Vencido"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(4, 1), HojaExcel.ActiveSheet.Cells(4, 13)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(4, 1), HojaExcel.ActiveSheet.Cells(4, 13)).Interior.ColorIndex 	= 43 	' Fondo.
	' Formato.
	HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "#,##0.00"
	HojaExcel.ActiveSheet.Columns(7).NumberFormat	= "#,##0"
	HojaExcel.ActiveSheet.Columns(12).NumberFormat	= "#,##0.00"
    HojaExcel.ActiveSheet.Columns(13).NumberFormat	= "#,##0.00"

	xRs.Source = "exec SP_Cliente_ACobrar '" & fechaDesdeStr & "', '" & fechaHastaStr & "'"
	xRs.Open
	R = 5
	aux_cliente 	= ""
	aux_total		= 0.0
	do while not xRs.EOF
		call ProgressControlAvance(Self.Workspace, xRs("Cliente").Value & " - " & xRs("Comprobante").Value)
		
		if aux_cliente <> xRs("Cliente").Value then
			if aux_cliente <> "" then
			   	HojaExcel.ActiveSheet.Cells(R, 1).Value 	= aux_cliente
				HojaExcel.ActiveSheet.Cells(R, 2).Value 	= "TOTAL"
				HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CDbl(aux_total)
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 			= true 	' Negrita.
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 	= 15 	' Fondo.
				aux_total = 0.0
				R = R + 1
			end if
			
			aux_cliente	= xRs("Cliente").Value
			HojaExcel.ActiveSheet.Cells(R, 1).Value = xRs("Cliente").Value
		end if
		
		aux_total 	= aux_total + CDbl(xRs("Saldo").Value)
		
		' 24/06/2021 - Para mejorar la velocidad.
		redim datos(11)
		datos(0)		= xRs("Comprobante").Value
		datos(1)		= CDbl(xRs("Saldo").Value)
		datos(2)		= xRs("Vencida").Value
		datos(3)		= xRs("Tipo").Value
		datos(4)		= xRs("Servicio").Value
		datos(5)		= xRs("DiasHoy").Value
		datos(6)		= xRs("CentroCostos").Value
		datos(7)		= xRs("Seguimiento").Value
		datos(8)		= CDate(xRs("FECHAESTADO").Value)
		datos(9)		= xRs("OBSERVAESTADO").Value
        datos(10)		= xRs("NoVencido").Value
        datos(11)		= xRs("Vencido").Value

		set rango 				= HojaExcel.ActiveSheet.Range("B" & R)
		rango.Resize(1, 12) 	= datos

		if UCase(xRs("Vencida").Value) = "VENCIDA" then
			HojaExcel.ActiveSheet.Cells(R, 4).Interior.ColorIndex 	= 3 	' Fondo: Rojo.
			HojaExcel.ActiveSheet.Cells(R, 4).Font.Color			= vbWhite
		end if
		
		select case xRs("EstadoPago").Value
		case "01"
			HojaExcel.ActiveSheet.Cells(R, 2).Interior.ColorIndex 	= 35 	' Fondo: Verde claro.
		case "02"
			HojaExcel.ActiveSheet.Cells(R, 2).Interior.ColorIndex 	= 4 	' Fondo: Verde intenso.
		case "03"
			HojaExcel.ActiveSheet.Cells(R, 2).Interior.ColorIndex 	= 38 	' Fondo: Rosa oscuro.
        case "04"
			HojaExcel.ActiveSheet.Cells(R, 2).Interior.ColorIndex 	= 10 	' Fondo: Rosa oscuro.
		end select
		
		' if CLng(xRs("DiasHoy").Value) > 1 and CLng(xRs("DiasHoy").Value) <= 5 then
			' HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 6), HojaExcel.ActiveSheet.Cells(R, 6)).Interior.ColorIndex 	= 44 	' Fondo.
		' end if
		' if CLng(xRs("DiasHoy").Value) > 5 and CLng(xRs("DiasHoy").Value) <= 10 then
			' HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 6), HojaExcel.ActiveSheet.Cells(R, 6)).Interior.ColorIndex 	= 46 	' Fondo.
		' end if
		if CLng(xRs("DiasHoy").Value) >= 1 and CLng(xRs("DiasHoy").Value) <= 10 then
			HojaExcel.ActiveSheet.Cells(R, 6).Interior.ColorIndex 	= 6 	' Fondo: Amarillo.
		end if
		if CLng(xRs("DiasHoy").Value) > 10 then
			HojaExcel.ActiveSheet.Cells(R, 6).Interior.ColorIndex 	= 3 	' Fondo: Rojo.
		end if
		
		' 24/06/2021 - Nuevos colores y columnas.
		if CDate(xRs("FECHAESTADO").Value) = CDate("01/01/2000") then
			HojaExcel.ActiveSheet.Cells(R, 9).Value	= ""
		end if
		' Fecha Servicio.
		'dias			= DateDiff("d", CDate(xRs("FECHASERVICIO").Value), Now)
		'if dias > 30 then
		'	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 6), HojaExcel.ActiveSheet.Cells(R, 6)).Interior.ColorIndex 		= 3		' Fondo: Rojo.
		'end if
		' Fecha Estado.
		if CDate(xRs("FECHAESTADO").Value) <> CDate("01/01/2000") then
			dias		= DateDiff("d", CDate(xRs("FECHAESTADO").Value), Now)
			if dias > 15 then
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 9), HojaExcel.ActiveSheet.Cells(R, 9)).Interior.ColorIndex = 3		' Fondo: Rojo.
			end if
		end if
		
		R = R + 1
		xRs.MoveNext
	loop
	
	HojaExcel.ActiveSheet.Cells(R, 2).Value 	= "TOTAL"
	HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CDbl(aux_total)
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 13)).Interior.ColorIndex 	= 15 	' Fondo.
	HojaExcel.ActiveSheet.Columns("A:Z").AutoFit
	'HojaExcel.ActiveSheet.Columns("I").ColumnWidth 	= 15
	'HojaExcel.ActiveSheet.Columns(8).EntireColumn.Delete  	' 14/07/2021 - quitamos esta columna a pedido de Estefi.
	' Calcula las sumas acumuladas

     xRs.MoveFirst ' Volver al primer registro para calcular las sumas desde el inicio
     Do While Not xRs.EOF
    ' if xRs("EstadoPago").Value <> "04" then
        sumaSaldo = sumaSaldo + CDbl(xRs("Saldo").Value)
        sumaVencido = sumaVencido + CDbl(xRs("Vencido").Value)
        sumaNoVencido = sumaNoVencido + CDbl(xRs("NoVencido").Value)
    ' end if
     xRs.MoveNext
     Loop

     ' Agrega la última fila con las sumas acumuladas
      HojaExcel.ActiveSheet.Cells(R, 1).Value = "Total acumulado"
      HojaExcel.ActiveSheet.Cells(R, 3).Value = sumaSaldo
      HojaExcel.ActiveSheet.Cells(R, 11).Value = sumaNoVencido
      HojaExcel.ActiveSheet.Cells(R, 12).Value = sumaVencido

      ' Ajusta automáticamente el ancho de las columnas.
      HojaExcel.ActiveSheet.Columns("A:Z").AutoFit

	' ------------------------
	' -- Hoja2: REFERENCIAS --
	' ------------------------
	
	HojaExcel.Sheets("Hoja2").Select
	HojaExcel.Sheets("Hoja2").Name	= "REFERENCIAS"
	
	' Fuente.
	HojaExcel.ActiveSheet.Cells.Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Cells.Font.Size 		= 11
	
	' Estado del Pago.
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "Estado del Pago"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1, 1)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(1, 2)).Interior.ColorIndex 	= 15 	' Fondo: Gris.
	
	HojaExcel.ActiveSheet.Cells(2, 1).Interior.ColorIndex 	= 35 	' 01.
	HojaExcel.ActiveSheet.Cells(2, 2).Value = "Cargada para el Pago"
	HojaExcel.ActiveSheet.Cells(3, 1).Interior.ColorIndex 	= 4 	' 02.
	HojaExcel.ActiveSheet.Cells(3, 2).Value = "Pago Agendado/ Confirmado"
	HojaExcel.ActiveSheet.Cells(4, 1).Interior.ColorIndex 	= 38 	' 03.
	HojaExcel.ActiveSheet.Cells(4, 2).Value = "Con Problemas"
      HojaExcel.ActiveSheet.Cells(5, 1).Interior.ColorIndex 	= 10 	' 04.
	HojaExcel.ActiveSheet.Cells(5, 2).Value = "Pago Generado"

	' Vencida.
	HojaExcel.ActiveSheet.Cells(7, 1).Value = "Vencida Días"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(7, 1), HojaExcel.ActiveSheet.Cells(7, 1)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(7, 1), HojaExcel.ActiveSheet.Cells(7, 2)).Interior.ColorIndex 	= 15 	' Fondo: Gris.
	
	' HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(8, 1), HojaExcel.ActiveSheet.Cells(8, 1)).Interior.ColorIndex 	= 44 	' 5.
	' HojaExcel.ActiveSheet.Cells(8, 2).Value = "Entre 1 y 5"
	' HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(9, 1), HojaExcel.ActiveSheet.Cells(9, 1)).Interior.ColorIndex 	= 46 	' 10.
	' HojaExcel.ActiveSheet.Cells(9, 2).Value = "Entre 6 y 10"
	
	HojaExcel.ActiveSheet.Cells(8, 1).Interior.ColorIndex 	= 6 	' Fondo: Amarillo.
	HojaExcel.ActiveSheet.Cells(8, 2).Value = "Entre 1 y 10"
	HojaExcel.ActiveSheet.Cells(9, 1).Interior.ColorIndex 	= 3 	' Fondo: Rojo.
	HojaExcel.ActiveSheet.Cells(9, 2).Value = "Más de 10"
	
	' Fechas.
	HojaExcel.ActiveSheet.Cells(12, 1).Value = "Fechas"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(12, 1), HojaExcel.ActiveSheet.Cells(12, 1)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(12, 1), HojaExcel.ActiveSheet.Cells(12, 2)).Interior.ColorIndex	= 15 	' Fondo: Gris.
	
	HojaExcel.ActiveSheet.Cells(13, 2).Value = "Servicio: más de 30 días"
	HojaExcel.ActiveSheet.Cells(14, 1).Interior.ColorIndex	= 3 	' Fondo: Rojo.
	HojaExcel.ActiveSheet.Cells(14, 2).Value = "Fe. Estado: más de 15 días"
	
	HojaExcel.ActiveSheet.Columns("A:Z").AutoFit
	
	' ---------------------
	' -- Hoja3: A COBRAR --
	' ---------------------
   	HojaExcel.Sheets.Add
    HojaExcel.Sheets("Hoja4").Select
    HojaExcel.Sheets("Hoja4").Name	= "A COBRAR"

	
	' Fuente.
	HojaExcel.ActiveSheet.Cells.Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Cells.Font.Size 		= 11
	
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "FECHAS A COBRAR"
	HojaExcel.ActiveSheet.Cells(2, 1).Value = "Periodo: " & FormatDateTime(fechaDesde, 2) & " - " & FormatDateTime(fechaHasta, 2)
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 4)).Font.Bold 			= true 	   	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 4)).Font.Color			= vbWhite	' Blanco.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 4)).Interior.ColorIndex 	= 10 		' Fondo: Verde oscuro.
	' Columnas.
	HojaExcel.ActiveSheet.Cells(4, 1).Value 		= "Cliente"
	HojaExcel.ActiveSheet.Cells(4, 2).Value 		= "Saldo"
	HojaExcel.ActiveSheet.Cells(4, 3).Value 		= "Fe. a Cobrar"
	HojaExcel.ActiveSheet.Cells(4, 4).Value			= "Estado"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(4, 1), HojaExcel.ActiveSheet.Cells(4, 4)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(4, 1), HojaExcel.ActiveSheet.Cells(4, 4)).Interior.ColorIndex 	= 43 	' Fondo.
	' Formatos.
	HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "dd/MM/yyyy"

	query			= "exec SP_Cliente_ACobrarEstado '" & fechaDesdeStr & "', '" & fechaHastaStr & "'"
	call CargaExcelManual(Self.WorkSpace, query, HojaExcel, 5, "SinCabe")
	HojaExcel.ActiveSheet.Columns("A:Z").AutoFit


	' ---------------------
	' -- Hoja4:Vencida+15 dias  --
	' ---------------------
	'RFIERRO 20/11/2023
	HojaExcel.Sheets("Hoja3").Select
	HojaExcel.Sheets("Hoja3").Name	= "Vencidas +15dias"

	' Fuente.
	HojaExcel.ActiveSheet.Cells.Font.Name 		= "Calibri"
	HojaExcel.ActiveSheet.Cells.Font.Size 		= 11
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "Saldos con mas de 15 dias"
	HojaExcel.ActiveSheet.Cells(2, 1).Value = "Periodo: " & FormatDateTime(fechaDesde, 2) & " - " & FormatDateTime(fechaHasta, 2)
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 10)).Font.Bold 			= true 	   	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 10)).Font.Color			= vbWhite	' Blanco.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(1, 1), HojaExcel.ActiveSheet.Cells(2, 10)).Interior.ColorIndex 	= 10 		' Fondo: Verde oscuro.
	' Columnas.
	HojaExcel.ActiveSheet.Cells(4, 1).Value 		= "Cliente"
	HojaExcel.ActiveSheet.Cells(4, 2).Value 		= "Comprobante"
	HojaExcel.ActiveSheet.Cells(4, 3).Value 		= "Saldo"
	HojaExcel.ActiveSheet.Cells(4, 4).Value 		= "Tipo"
	HojaExcel.ActiveSheet.Cells(4, 5).Value 		= "Servicio"
	HojaExcel.ActiveSheet.Cells(4, 6).Value 		= "Días a Hoy"
	HojaExcel.ActiveSheet.Cells(4, 7).Value 		= "Centro de Costos"
	HojaExcel.ActiveSheet.Cells(4, 8).Value 		= "Seguimiento"
	HojaExcel.ActiveSheet.Cells(4, 9).Value 		= "Fe. Estado"
	HojaExcel.ActiveSheet.Cells(4, 10).Value 		= "Observa. Estado"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(4, 1), HojaExcel.ActiveSheet.Cells(4, 10)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(4, 1), HojaExcel.ActiveSheet.Cells(4, 10)).Interior.ColorIndex 	= 43 	' Fondo.
	'Formato.
	HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "#,##0.00"
	HojaExcel.ActiveSheet.Columns(7).NumberFormat	= "#,##0"
	HojaExcel.ActiveSheet.Columns(12).NumberFormat	= "#,##0.00"
    HojaExcel.ActiveSheet.Columns(13).NumberFormat	= "#,##0.00"

    	xRs.Close
	xRs.ActiveConnection.CommandTimeout = 0
	xRs.Source = "exec SP_Cliente_ACobrar15dias '" & fechaDesdeStr & "', '" & fechaHastaStr & "'"
	xRs.Open
	R = 5
	aux_cliente 	= ""
	aux_total		= 0.0
	do while not xRs.EOF
		call ProgressControlAvance(Self.Workspace, xRs("Cliente").Value & " - " & xRs("Comprobante").Value)

		if aux_cliente <> xRs("Cliente").Value then
			if aux_cliente <> "" then
			   	HojaExcel.ActiveSheet.Cells(R, 1).Value 	= aux_cliente
				HojaExcel.ActiveSheet.Cells(R, 2).Value 	= "TOTAL"
				HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CDbl(aux_total)
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 10)).Font.Bold 			= true 	' Negrita.
				HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 10)).Interior.ColorIndex 		= 15 	' Fondo.
				aux_total = 0.0
				R = R + 1
			end if

			aux_cliente	= xRs("Cliente").Value
			HojaExcel.ActiveSheet.Cells(R, 1).Value = xRs("Cliente").Value
		end if

		aux_total 	= aux_total + CDbl(xRs("Saldo").Value)

		redim datos(8)
		datos(0)		= xRs("Comprobante").Value
		datos(1)		= CDbl(xRs("Saldo").Value)
		datos(2)		= xRs("Tipo").Value
		datos(3)		= xRs("Servicio").Value
		datos(4)		= xRs("DiasHoy").Value
		datos(5)		= xRs("CentroCostos").Value
		datos(6)		= xRs("Seguimiento").Value
		datos(7)		= CDate(xRs("FECHAESTADO").Value)
		datos(8)		= xRs("OBSERVAESTADO").Value


		set rango 				= HojaExcel.ActiveSheet.Range("B" & R)
		rango.Resize(1, 9) 	= datos

		R = R + 1
		xRs.MoveNext
	loop

	HojaExcel.ActiveSheet.Cells(R, 2).Value 	= "TOTAL"
	HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CDbl(aux_total)
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 10)).Font.Bold 			= true 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 10)).Interior.ColorIndex 		= 15 	' Fondo.
	HojaExcel.ActiveSheet.Columns("A:Z").AutoFit

    xRs.MoveFirst ' Volver al primer registro para calcular las sumas desde el inicio
    sumaSaldo15 = 0
    Do While Not xRs.EOF
        sumaSaldo15 = sumaSaldo15 + CDbl(xRs("Saldo").Value)
        xRs.MoveNext
    Loop

    ' Agrega la última fila con las suma Saldo
    HojaExcel.ActiveSheet.Cells(R, 1).Value = "Total acumulado"
    HojaExcel.ActiveSheet.Cells(R, 3).Value = sumaSaldo15

    ' Ajusta automáticamente el ancho de las columnas.
    HojaExcel.ActiveSheet.Columns("A:Z").AutoFit



	call ProgressControlFinish(Self.Workspace)
	SendDebug "FIN A Cobrar!!!!"
	HojaExcel.Visible 	= true
	set HojaExcel 		= nothing
end sub
