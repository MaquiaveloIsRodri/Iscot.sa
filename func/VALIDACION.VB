Sub Main

    Stop
    xaddress = OpenFileDialog("C:\", "xls files (*.xls)")
    If xaddress = "" Then
        Call MsgBox ("No seleccionó ningún archivo. Proceso cancelado.", 64, "Información") : Exit Sub
    End If

    Set LibroExcel = CreateObject("Excel.Application")
    LibroExcel.Workbooks.Open xaddress
    LibroExcel.Sheets("Form1").Select
    I = 2
    Contador = 0
    While Trim(LibroExcel.ActiveSheet.Cells(I, 1).Value) <> ""
        Contador = Contador + 1
        I = I + 1
    Wend
    I = 2  :   Creados = 0 ':   Contador = 0

    Call ProgressControl(Self.Workspace, "Cargando Excel - Accidentes", 0, Contador)


    While Trim(LibroExcel.ActiveSheet.Cells(I, 1).Value) <> ""
        call ProgressControlAvance(Self.Workspace, "Cargando Excel - Celulares")

        xError = False
    '    Set oEmpleadoPotencial  = Nothing
        Set oEmpleado           = Nothing


        xId            			= "" :  xId                   = Trim(LibroExcel.ActiveSheet.Cells(I,   1).Value)
		xlegajo					= "" :  xlegajo               = Trim(LibroExcel.ActiveSheet.Cells(I,   7).Value)
        xlinea                  = "" :  xlinea                = Trim(LibroExcel.ActiveSheet.Cells(I,   11).Value)
        xMarca              	= "" :  xMarca                = Trim(LibroExcel.ActiveSheet.Cells(I,   12).Value)

        set oEmpleado = ExisteBo(Self, "EMPLEADO", "CODIGO", xLegajo,Nil, True, False, "=")
        If oEmpleado Is Nothing Then
            xError = True
            LibroExcel.ActiveSheet.Cells(I, 46).Value = chr(13) & "Empleado no encontrado. - " & LibroExcel.ActiveSheet.Cells(I, 46).Value
        End If

        If Not xError Then
            set xItemNuevo = CrearBo("UD_CELULARES", self)
            self.bo_place.bo_owner.bo_owner.boextension.Celulares.add(xItemNuevo)

            xItemNuevo.CODIGO                = xId
            xItemNuevo.MARCAMODELO           = xMarca
            xItemNuevo.RESPONSABLE           = oEmpleado
            xItemNuevo.NUMEROLINEA           = xlinea


            'If Not oEmpleado            Is Nothing Then set xItemNuevo.RESPONSABLE             = oEmpleado
            Creados = Creados + 1
            LibroExcel.ActiveSheet.Cells(I, 46).Value = "Celular creado creado. - " & LibroExcel.ActiveSheet.Cells(I, 46).Value

        End If
        I = I + 1

    Wend

    Call WorkspaceCheck(Self.Workspace)
    LibroExcel.Visible = True
    Set LibroExcel = Nothing
    Call ProgressControlFinish(Self.Workspace)
    MsgBox "Proceso Finalizado" & chr(13) & chr(13) & "Se agregaron " & Creados & "registros" & chr(13) & "No procesados: " & Contador-Creados & Chr(13) & "Revisar columna 11 en planilla"

End Sub
