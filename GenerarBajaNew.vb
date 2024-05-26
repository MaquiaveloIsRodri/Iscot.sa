Sub Main
Stop
    Set xFSO 		= CreateObject("Scripting.FileSystemObject")
    Set xArchivo    = xFSO.OpenTextFile("C:\util\html\email_bajas-pruebas.html")
    htmlBody        = xArchivo.readAll()  
    'linkPowerBI     = "https://app.powerbi.com/groups/dc75db2f-ac71-4db2-8658-a0fdea0dc978/reports/6c5e0695-0a6e-4d17-a189-18002322e97f/ReportSection78634f367cd50de51dd1?experience=power-bi&clientSideAuth=0" 

    Set xObj = Nothing
    contador = 0
	
	for each xBaja in container
        'Set xBaja = self
		if xBaja.flag is nothing then
            Call Msgbox("Transacción sin flag, contacte a sistemas",64, "Información")
		   Exit Sub		
        Else
            If xBaja.flag.id = "{77A2D825-05AF-498A-B91E-814149A20AF0}" then ' cerrada
                Call Msgbox("La Baja " & xBaja.numerodocumento & " ya fué cerrada.",64, "Información")
                Exit sub
            End If

		end if
        contador = contador + 1

        'Generamos un item obligatorio de calipso
        Set xBajaItem = CrearItemTransaccion(xBaja)
        xBajaItem.referencia = Nothing

        Set obj = xBaja
        textoEmpleados = textoEmpleados & vbNewLine & xBaja.destinatario.enteasociado.nombre
    next


    if MsgBox("Confirma la baja de " & Cstr(contador) & " empleados?" & textoEmpleados,36,"Pregunta") <> 6 Then
        Call Msgbox("Proceso cancelado.",64, "Información")
        Exit Sub
    Else
        msg = ""
        xSubject = "[AVISO BAJAS] " & Date() 
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
        xRsMails.source = "SELECT ID,CODIGO, NOMBRE FROM ITEMTIPOCLASIFICADOR WHERE BO_PLACE_ID = '94477911-84B4-443E-9FE6-775494E7E75B' AND ACTIVESTATUS = 0 "
        xRsMails.open

        Set xVisualVar = VisualVarEditor("Aviso por email")
        Do while not xRsMails.eof
            Call AddVarBoolean(xVisualVar, xRsMails("ID").Value, xRsMails("CODIGO").Value, "Seleccion", true)
            xRsMails.MOVENEXT
        Loop
        xRsMails.close

        Call AddVarBoolean(xVisualVar, "CC", "Referentes del Servicio", "Seleccion", true)
		
        call AddVarMemo(xVisualVar, "9_Extras", "Extras", "Extras","" )		
        xAceptar = ShowVisualVar(xVisualVar)
		
        If xAceptar Then
            correos = ""
            correosExtras = ""
            correosExtras = GetValueVisualVar(xVisualVar, "9_Extras", "Extras")
            set xRsControl = xRsMails
            xRsControl.open
            
            Do While Not xRsControl.eof
                If GetValueVisualVar(xVisualVar, xRsControl("ID").Value, "Seleccion") Then
                    correos = correos & ";" & xRsControl("NOMBRE").Value
                End If
                xRsControl.MOVENEXT
            Loop
            xRsControl.close
            If correosExtras <> "" Then
                correos = correos & correosExtras
            End If
			enviarOP = GetValueVisualVar(xVisualVar, "CC", "Seleccion")
        Else
            Call Msgbox("Proceso cancelado.",64, "Información")
            Exit Sub
        End If
    end if
    
	
    cantidadBajas = 0
    For Each xBaja In container
		' Busco los mails de los referentes del CC
        msg = msg & VbCrlf & VbCrlf & "Destinatarios del Servicio " & xBaja.centrocostos.name & VbCrlf
	    'For Each xDestinatario in xBaja.centrocostos.BoExtension.AvisoBaja
        '    If xDestinatario.BoExtension.CorreoInstitucional <> "" then
        '        correos = correos & xDestinatario.BoExtension.CorreoInstitucional & ";"
        '        msg =  msg & xDestinatario.name & ", "
        '    Else
        '      	Call Msgbox("El empleado :" & xDestinatario.name & " no tiene configurado el correo institucional.",64, "Información")
        '    End If
        'Next
		correosOp = getMailsDestinatarios(xBaja.Destinatario)
		If enviarOP Then correos = correosOP & ";" & correos
        Set xObj = xBaja
	  call EjecutarTransicion( xBaja , "Cerrar Baja de Personal" )
        'htmlItems = codItems
        items = ""
        items = items & "<tr>"
        items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:60px;  padding-right:10px; padding-left:10px; background:#e2efda'>"& xBaja.destinatario.codigo  &"</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:100px; padding-right:10px; padding-left:10px;'>"& xBaja.destinatario.EnteAsociado.nombre & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:60px;  padding-right:10px; padding-left:10px;'>"& xBaja.BoExtension.Cuit & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:180px; padding-right:10px; padding-left:10px;'>"& xBaja.centrocostos.name & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:25px;  padding-right:10px; padding-left:10px;'>"& xBaja.BoExtension.Perfil.name & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:25px;  padding-right:10px; padding-left:10px;'>"& xBaja.destinatario.fechaIngreso & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:30px;  padding-right:10px; padding-left:10px; background:#1f4e78; color:white'>"& xBaja.boextension.fechabaja & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:30px;  padding-right:10px; padding-left:10px; background:#ddebf7;'>"& xBaja.BoExtension.FECHAENVIO & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:30px;  padding-right:10px; padding-left:10px; background:#ddebf7;'>"& xBaja.BoExtension.FECHARECEPCION & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:90px;  padding-right:10px; padding-left:10px;'>"& xBaja.BoExtension.MOTIVOEGRESO.name & "</td>"
		items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:90px;  padding-right:10px; padding-left:10px;'>"& xBaja.BoExtension.MOTIVOAFIP.name & "</td>"
        If xBaja.BoExtension.DEPOSITALIQUIDACIONFINAL Then 
            items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:10px;  padding-right:10px; padding-left:10px;'>SI</td>" 
        Else
            items = items & "<td style='height:18px; font-family: 'Roboto', sans-serif; font-size: 12px; text-align:center; width:10px;  padding-right:10px; padding-left:10px;'>NO</td>"
        End If
        If xBaja.BoExtension.DESCUENTAINDUMENTARIA Then 
            items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:10px;  padding-right:10px; padding-left:10px; background:#ddebf7;'>SI</td>" 
        Else
            items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:10px;  padding-right:10px; padding-left:10px;background:#ddebf7; '>NO</td>"
        End If
        If xBaja.BoExtension.DESCUENTATARJETA Then 
            items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:10px;  padding-right:10px; padding-left:10px; background:#ddebf7;'>SI</td>" 
        Else
            items = items & "<td style='height:18px; font-family: Roboto, sans-serif; font-size: 12px; text-align:center; width:10px;  padding-right:10px; padding-left:10px; background:#ddebf7;'>NO</td>"
        End If
      
        items = items & "</tr>"



        xItemsBaja  = xItemsBaja & " " & items 
        cantidadBajas = cantidadBajas + 1
	Next

    htmlBody = replace(htmlBody,"xfecha"    , Cstr(Date))
    htmlBody = replace(htmlBody,"xCantidad" , cantidadBajas )
    htmlBody = replace(htmlBody,"xItems"    , xItemsBaja )
	Stop
	if MsgBox("Envíar correo?",36,"Confirma?") <> 6 Then Exit Sub   

    call enviar_aviso(xObj, correos , xSubject, xBody, xadjunto,htmlBody)  
    Call WorkSpaceCheck(xObj.Workspace)
End Sub

Function getMailsDestinatarios(xEmpleado)
    destinatarios = ""


    If Not xEmpleado.centrocostos.boextension.ANALISTARRHH Is Nothing Then destinatarios = destinatarios & ";" & xEmpleado.centrocostos.boextension.ANALISTARRHH.boextension.correoinstitucional
    ' If Not xEmpleado.centrocostos.boextension.GERENTEOPERATIVO Is Nothing Then destinatarios = destinatarios & ";" & xEmpleado.centrocostos.boextension.GERENTEOPERATIVO.boextension.correoinstitucional
    If Not xEmpleado.centrocostos.boextension.COORDINADOROPERATIVO Is Nothing Then 
        For Each xi in xEmpleado.centrocostos.boextension.COORDINADOROPERATIVO
            destinatarios = destinatarios & ";" & xi.boextension.correoinstitucional
        Next
    End If
    ' If Not xEmpleado.centrocostos.boextension.SUPERVISOR Is Nothing Then destinatarios = destinatarios & ";" & xEmpleado.centrocostos.boextension.SUPERVISOR.boextension.correoinstitucional
    If Not xEmpleado.centrocostos.boextension.RESPONSABLESERVICIO Is Nothing Then destinatarios = destinatarios & ";" & xEmpleado.centrocostos.boextension.RESPONSABLESERVICIO.boextension.correoinstitucional
    If Not xEmpleado.centrocostos.boextension.RESPONSABLESERVICIO2 Is Nothing Then destinatarios = destinatarios & ";" & xEmpleado.centrocostos.boextension.RESPONSABLESERVICIO2.boextension.correoinstitucional

    getMailsDestinatarios =  destinatarios
End Function


