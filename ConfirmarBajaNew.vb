Sub Main
    set xFSO 		= CreateObject("Scripting.FileSystemObject")
    set xArchivo    = xFSO.OpenTextFile("C:\util\html\email_bajas_html.html")
    htmlBody        = xArchivo.readAll()  
    set xFSOitems 		= CreateObject("Scripting.FileSystemObject")
    set xArchivoItems    = xFSOitems.OpenTextFile("C:\util\html\email_bajas_items.html")
    linkPowerBI = "https://app.powerbi.com/view?r=eyJrIjoiNmVhMjVlMjMtYjc0YS00Yjk3LWEwZDYtOGIwMzYxNzJmMDM0IiwidCI6IjcyYjZlOGMwLTA3YjEtNDMzNC05MzhlLWZmOTZhODRmYWVlYyIsImMiOjR9" 
    codItems = xArchivoItems.readAll()
	Stop
    Set xObj = Nothing
    contador = 0
	for each xBaja in container
		if xBaja.flag is nothing then
            Call Msgbox("Transacción sin flag, contacte a sistemas",64, "Información")
		   Exit Sub
		else
            if xBaja.flag.id = "{77A2D825-05AF-498A-B91E-814149A20AF0}" then ' cerrada
                Call Msgbox("La Baja " & xBaja.numerodocumento & " ya fué cerrada.",64, "Información")
                Exit sub
            End If
				
		end if
        contador = contador + 1
        Set obj = xBaja
        textoEmpleados = textoEmpleados & vbNewLine & xBaja.destinatario.enteasociado.nombre
    next
    if MsgBox("Confirma la baja de " & Cstr(contador) & " empleados?" & textoEmpleados,36,"Pregunta") <> 6 Then
        Call Msgbox("Proceso cancelado.",64, "Información")
        Exit Sub
    Else
        xSubject = "[Bajas] " & Date() 
        xBody    = ""       
        xadjunto = ""

        Set xDict = NewDic()
        xstring = StringConexion( "CALIPSO", obj.WorkSpace ) 
        set xcone = createobject("adodb.connection")
        xcone.connectionstring  = xstring
        xcone.connectiontimeout = 0
        set xRsMails = RecordSet(xCone, "select top 1 * from producto")
        xRsMails.close
        xRsMails.activeconnection.commandtimeout = 0
        xRsMails.source = "SELECT NOMBRE FROM ITEMTIPOCLASIFICADOR WHERE BO_PLACE_ID = '94477911-84B4-443E-9FE6-775494E7E75B' AND ACTIVESTATUS = 0 "
        xRsMails.open
        xcorreos  = ""
        Do while not xRsMails.eof               
            xcorreos = xcorreos & xRsMails("NOMBRE").Value & ";"        
            xRsMails.MOVENEXT
        Loop
        xRsMails.close  
        Set xVisualVar = VisualVarEditor("Aviso por email")
        call AddVarMemo(xVisualVar, "Correos", "Correos", "Correos", xcorreos )
        xAceptar = ShowVisualVar(xVisualVar)
        If xAceptar Then
            xcorreos = GetValueVisualVar(xVisualVar, "Correos", "Correos")        
        Else
            Call Msgbox("Proceso cancelado.",64, "Información")
            Exit Sub
        End If
    end if



	for each xBaja in container
		'TO-DO : Validar todos los campos antes
        Set xObj = xBaja
		call EjecutarTransicion( xBaja , "Cerrar Baja de Personal" )
		'xBaja.destinatario.fechabaja = xBaja.boextension.FECHABAJA
		'xBaja.destinatario.motivoBaja = xBaja.boextension.MOTIVOEGRESO

		'xBaja.destinatario.boextension.MOTIVOAFIP = xBaja.boextension.MOTIVOAFIP
		'xBaja.destinatario.boextension.TRBAJA = xBaja
		'xBaja.destinatario.boextension.DescuentoTarjeta = xBaja.boextension.DESCUENTATARJETA
		'xBaja.destinatario.boextension.DescuentoIndumentaria = xBaja.boextension.DESCUENTAINDUMENTARIA
		'xBaja.destinatario.boextension.DEPOSITALIQUIDACIONFINAL = xBaja.boextension.DEPOSITALIQUIDACIONFINAL
		
    
       
        htmlItems = codItems
        htmlItems = replace(htmlItems,"xLegajo" , xBaja.destinatario.codigo)
        htmlItems = replace(htmlItems,"xEmpleado" , xBaja.destinatario.enteasociado.nombre)
        'htmlItems = replace(htmlItems,"xCC" , xBaja.destinatario.centrocostos.nombre)
		
		if not xBaja.boextension.motivoegreso is nothing then
            htmlItems = replace(htmlItems,"xCC" , xBaja.boextension.motivoegreso.nombre )
        else
            htmlItems = replace(htmlItems,"xCC" , "" )
        end if 
		
        if not xBaja.boextension.motivoegreso is nothing then
            htmlItems = replace(htmlItems,"xMotivo" , xBaja.boextension.motivoegreso.nombre )
        else
            htmlItems = replace(htmlItems,"xMotivo" , "" )
        end if 
		
        htmlItems = replace(htmlItems,"xFechaIngreso" , xBaja.destinatario.fechaIngreso )
        htmlItems = replace(htmlItems,"xFechaEgreso" , xBaja.boextension.fechabaja)
        xItemsBaja = xItemsBaja & " " & htmlItems
        contador = contador + 1
	Next
    htmlBody = replace(htmlBody,"xFecha" , Cstr(Date))
    htmlBody = replace(htmlBody,"xCantidad" , contador)
    htmlBody = replace(htmlBody,"xItems" , xItemsBaja )
         
    call enviar_aviso(xObj, xcorreos, xSubject, xBody, xadjunto,htmlBody)  
End Sub
