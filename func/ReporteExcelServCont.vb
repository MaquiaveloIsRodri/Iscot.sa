Sub main
    Stop
	contador = 0
    Set HojaExcel = CreateObject("Excel.Application")
	HojaExcel.Workbooks.Add
    HojaExcel.Sheets("Hoja1").Select

	HojaExcel.ActiveSheet.Cells(1, 1).Value = "Servicio "& self.nombre
	HojaExcel.ActiveSheet.Cells(1, 1).Font.Bold = True 	' Negrita.
	HojaExcel.ActiveSheet.Cells(1, 1).Interior.ColorIndex = 15 	' Fondo Verde.

	HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "Orden de compra nueva"
	HojaExcel.ActiveSheet.Cells(3, 2).Value 		= "Centro de costos"
	HojaExcel.ActiveSheet.Cells(3, 3).Value 		= "Responsable"
	HojaExcel.ActiveSheet.Cells(3, 4).Value 		= "Importe"

	HojaExcel.ActiveSheet.Cells(3, 5).Value 		= "Precio Actual"
	HojaExcel.ActiveSheet.Cells(3, 6).Value        	= "Q Promedio"
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Font.Bold 			= True 	' Negrita.
	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Interior.ColorIndex 	= 15 	' Fondo Gris.

		'referencias de colores
	  	HojaExcel.ActiveSheet.Cells(4, 8).Value 				= "REFERENCIA"
		HojaExcel.ActiveSheet.Cells(4, 8).Font.Bold 			= True 	' Negrita.
		HojaExcel.ActiveSheet.Cells(4, 8).Interior.ColorIndex 	= 15 	' Fondo Gris.
	  	HojaExcel.ActiveSheet.Cells(5, 8).Interior.ColorIndex 	= 3
	  	HojaExcel.ActiveSheet.Cells(5, 9).Value 				= "Cantidades Superiores al Mes anterior"
	  	HojaExcel.ActiveSheet.Cells(6, 8).Interior.ColorIndex 	= 42
	  	HojaExcel.ActiveSheet.Cells(6, 9).Value 				= "Cantidades Inferior al Mes anterior"
	  	HojaExcel.ActiveSheet.Cells(7, 8).Interior.ColorIndex 	= 43
	  	HojaExcel.ActiveSheet.Cells(7, 9).Value 				= "Cantidades Estables"

	for each xItem in self.boextension.SERVICIOSCONTRATADOS
		contador = contador + 1
        call ProgressControl(Self.Workspace, "PICs POR PERIODO POR CENTRO DE COSTO" , 0, 300)

    next
End sub