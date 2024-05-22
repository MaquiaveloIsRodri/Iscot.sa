' 19/05/2023 - Indicador Composición Mensual.
sub main
	stop
	
	if not (PerteneceAGrupo( "CALIPSO FACTURACION" ) or UCase(NombreUsuario()) = "SCALIPSO" or UCase(NombreUsuario()) = "NNIESS") then
		MsgBox "Usuario No Autorizado.", 48, "Aviso"
		exit sub
	end if
	
	' VisualVar.
	set xVisualVar = VisualVarEditor("Indicador Composición Mensual")
	call AddVarInteger(xVisualVar, "00MES", "Mes", "Indique:", Month(Date))
	call AddVarInteger(xVisualVar, "05ANIO", "Año", "Indique:", Year(Date))
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then exit sub
	
	mes 	= GetValueVisualVar(xVisualVar, "00MES", "Indique:")
	anio 	= GetValueVisualVar(xVisualVar, "05ANIO", "Indique:")
	if mes < 1 or mes > 12 then
		MsgBox "Mes Incorrecto.", 48, "Aviso"
		exit sub
	end if
	if anio < 2000 then
		MsgBox "Año Incorrecto.", 48, "Aviso"
		exit sub
	end if
	fecha = anio & Right("00" & mes, 2)
	
	call ProgressControl(Self.Workspace, "INDICADOR COMPOSICIÓN MENSUAL" , 0, 85)
	
	' EXCEL.
	set HojaExcel = CreateObject("Excel.Application")
	HojaExcel.Workbooks.Add
	
	' Cabecera.
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "ISCOT SERVICES S.A."
	HojaExcel.ActiveSheet.Cells(2, 1).Value = "INDICADOR COMPOSICIÓN MENSUAL"
	HojaExcel.ActiveSheet.Cells(3, 1).Value = "PERIODO: " & Right(fecha, 2) & "/" & Left(fecha, 4)
	for ind = 1 to 3
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(ind, 1), HojaExcel.ActiveSheet.Cells(ind, 8)).Font.Bold 			= true
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(ind, 1), HojaExcel.ActiveSheet.Cells(ind, 8)).Interior.ColorIndex 	= 43
	next
	
	' Columnas.
	HojaExcel.ActiveSheet.Cells(5, 1).Value 		= "Cliente"
	HojaExcel.ActiveSheet.Cells(5, 2).Value 		= "Total Facturado"
	HojaExcel.ActiveSheet.Cells(5, 3).Value 		= "% Factu."
	HojaExcel.ActiveSheet.Cells(5, 4).Value 		= "Total Previsionado"
	HojaExcel.ActiveSheet.Cells(5, 5).Value 		= "% Previ."
	HojaExcel.ActiveSheet.Cells(5, 6).Value 		= "Total"
	HojaExcel.ActiveSheet.Cells(5, 7).Value 		= "Diferencia Fac."
	HojaExcel.ActiveSheet.Cells(5, 8).Value 		= "Total Final"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 8)).Font.Bold 			= true
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(5, 1), HojaExcel.ActiveSheet.Cells(5, 8)).Interior.ColorIndex 	= 15
	' Formatear Celdas.
	HojaExcel.ActiveSheet.Columns(2).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(3).NumberFormat	= "#,##0.00"
	HojaExcel.ActiveSheet.Columns(4).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(5).NumberFormat	= "#,##0.00"
	HojaExcel.ActiveSheet.Columns(6).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(7).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns(8).NumberFormat	= "$ #,##0.00"
	
	' Colores.
	R = 6
	Call CargaExcelManual( Self.Workspace, "exec SP_Ventas_ComposicionMensual '" & fecha & "'", HojaExcel, R, "SinCabe" )
	
	do while HojaExcel.ActiveSheet.Cells(R, 6).Value <> ""
		call ProgressControlAvance(Self.Workspace, "CC: " & HojaExcel.ActiveSheet.Cells(R, 1).Value)
		
		if CDbl(HojaExcel.ActiveSheet.Cells(R, 3).Value) >= 80.0 then		' Total Facturado.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 3), HojaExcel.ActiveSheet.Cells(R, 3)).Interior.ColorIndex 	= 4
		end if
		if CDbl(HojaExcel.ActiveSheet.Cells(R, 5).Value) >= 80.0 then		' Total Previsionado.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 5), HojaExcel.ActiveSheet.Cells(R, 5)).Interior.ColorIndex 	= 3
		end if
		
		R = R + 1
	loop
	
	' Totales.
	HojaExcel.ActiveSheet.Cells(R, 2).Formula 		= "=SUM(B6:B" & R - 1 & ")"
	HojaExcel.ActiveSheet.Cells(R, 4).Formula 		= "=SUM(D6:D" & R - 1 & ")"
	HojaExcel.ActiveSheet.Cells(R, 6).Formula 		= "=SUM(F6:F" & R - 1 & ")"
	HojaExcel.ActiveSheet.Cells(R, 7).Formula 		= "=SUM(G6:G" & R - 1 & ")"
	HojaExcel.ActiveSheet.Cells(R, 8).Formula 		= "=SUM(H6:H" & R - 1 & ")"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R, 1), HojaExcel.ActiveSheet.Cells(R, 8)).Font.Bold 			= true
	
	HojaExcel.ActiveSheet.Columns("A:Z").AutoFit
	
	call ProgressControlFinish(Self.Workspace)
	HojaExcel.Visible = true
	set HojaExcel = nothing
end sub



