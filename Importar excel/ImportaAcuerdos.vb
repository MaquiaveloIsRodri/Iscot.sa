Sub Main
    Stop
	xstring  = stringconexion("Calipso",Self.workspace) : Set xcone = createobject("adodb.connection") : xcone.connectiontimeout = 0 : xcone.connectionstring = xstring : xcone.open	
	xaddress = OpenFileDialog("C:\", "xls files (*.xls)") 
	If xaddress = "" Then 
		Call MsgBox ("No seleccionó ningún archivo. Proceso cancelado.", 64, "Información") : Exit Sub 
	End If

    Set LibroExcel = CreateObject("Excel.Application")
  	LibroExcel.Workbooks.Open xaddress
    LibroExcel.Sheets("ACCIDENTES").Select
	I = 8 
    Contador = 0
    While Trim(LibroExcel.ActiveSheet.Cells(I, 1).Value) <> ""
        Contador = Contador + 1
        I = I + 1
    Wend
    I = 8  :   Creados = 0 ':   Contador = 0
    Call ProgressControl(Self.Workspace, "Cargando Excel - Accidentes", 0, Contador)

    While Trim(LibroExcel.ActiveSheet.Cells(I, 1).Value) <> ""
        call ProgressControlAvance(Self.Workspace, "Cargando Excel - Accidentes")
        xError = False

    '    Set oEmpleadoPotencial  = Nothing
        Set oEmpleado   = Nothing
    	set oCategoria  = Nothing
		Set oCC 	 	= Nothing


        xLegajo                         = "" :	xLegajo			            = Trim(LibroExcel.ActiveSheet.Cells(I,  1).Value)
        xCC		                        = "" :	xCC		                    = Trim(LibroExcel.ActiveSheet.Cells(I,  4).Value)'Consultar
        xPerfil     	                = "" :	xPerfil	                    = Trim(LibroExcel.ActiveSheet.Cells(I,  5).Value)'Consultar
        xCategoria		                = "" :	xCategoria	                = Trim(LibroExcel.ActiveSheet.Cells(I,  7).Value)'Consultar
	    xFechaDeEmision      = "" :	xFechaDeEmision	    		            = Trim(LibroExcel.ActiveSheet.Cells(I,  11).Value)
	    xFechaRecibidoTelegrama = "" :	xFechaRecibidoTelegrama	            = Trim(LibroExcel.ActiveSheet.Cells(I,  12).Value)


        ' ------------------------------------------------------------------------------------------------------------------------------- '
        set oEmpleado = ExisteBo(Self, "EMPLEADO", "CODIGO", xLegajo,Nil, True, False, "=")
        If oEmpleado Is Nothing Then
            xError =True
            LibroExcel.ActiveSheet.Cells(I, 46).Value = chr(13) & "Empleado no encontrado. - " & LibroExcel.ActiveSheet.Cells(I, 46).Value
        End If
        ' ------------------------------------------------------------------------------------------------------------------------------- '

	  Set oCC	 = ExisteBo(Self, "CENTROCOSTOS", "NOMBRE", xCC, nil , True, False, "=")
        If oCC Is Nothing Then
            xError =True
            LibroExcel.ActiveSheet.Cells(I, 14).Value = chr(13) & "tipo de centro de costo no existe. - " & LibroExcel.ActiveSheet.Cells(I, 14).Value 
        End If
        ' ------------------------------------------------------------------------------------------------------------------------------- '
        LibroExcel.ActiveWorkbook.save

        If Not xError Then

            set xItemNuevo = CrearBo("UD_ACCIDENTES", self)
            self.bo_place.bo_owner.bo_owner.boextension.Acuerdos.add(xItemNuevo)

            xItemNuevo.Legajo					= xLegajo
            set xItemNuevo.CentroDeCostos 		= oCC
 		    set xItemNuevo.Perfil         		= oEmpleado.Perfil
		    If IsDate(xFechaDeEmision) Then xItemNuevo.FechaEmision			= cDate(xFechaDeEmision)
		    If IsDate(xFechaDeEmision) Then xItemNuevo.FechaRecibidoTelegrama = cDate(xFechaRecibidoTelegrama)
            Creados = Creados + 1
            LibroExcel.ActiveSheet.Cells(I, 11).Value = "Acuerdos creados. - " & LibroExcel.ActiveSheet.Cells(I, 11).Value
        End If
        I = I + 1
    Wend

    Call WorkspaceCheck(Self.Workspace)
    LibroExcel.Visible = True
    Set LibroExcel = Nothing
    Call ProgressControlFinish(Self.Workspace)
    MsgBox "Proceso Finalizado" & chr(13) & chr(13) & "Se agregaron " & Creados & "registros" & chr(13) & "No procesados: " & Contador-Creados & Chr(13) & "Revisar columna 11 en planilla"

End Sub