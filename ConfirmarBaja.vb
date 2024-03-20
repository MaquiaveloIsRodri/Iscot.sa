Sub Main
Stop
    Set xFSO 		= CreateObject("Scripting.FileSystemObject")
    Set xArchivo    = xFSO.OpenTextFile("C:\util\html\email_bajas.html")
    htmlBody        = xArchivo.readAll()
    linkPowerBI     = "https://app.powerbi.com/view?r=eyJrIjoiNmVhMjVlMjMtYjc0YS00Yjk3LWEwZDYtOGIwMzYxNzJmMDM0IiwidCI6IjcyYjZlOGMwLTA3YjEtNDMzNC05MzhlLWZmOTZhODRmYWVlYyIsImMiOjR9"

    Set xObj = Nothing
    contador = 0

	for each xBaja in container
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
                    correos = correos & xRsControl("NOMBRE").Value
                End If
                xRsControl.MOVENEXT
            Loop
            xRsControl.close
            If correosExtras <> "" Then
                correos = correos & correosExtras
            End If
        Else
            Call Msgbox("Proceso cancelado.",64, "Información")
            Exit Sub
        End If
    end if


    cantidadBajas = 0
    For Each xBaja In container
		' Busco los mails de los referentes del CC
        msg = msg & VbCrlf & VbCrlf & "Destinatarios del Servicio " & xBaja.centrocostos.name & VbCrlf
	    For Each xDestinatario in xBaja.centrocostos.BoExtension.AvisoBaja
            If xDestinatario.BoExtension.CorreoInstitucional <> "" then
                correos = correos & xDestinatario.BoExtension.CorreoInstitucional & ";"
                msg =  msg & xDestinatario.name & ", "
            Else
              	Call Msgbox("El empleado :" & xDestinatario.name & " no tiene configurado el correo institucional.",64, "Información")
            End If
        Next
        Set xObj = xBaja
		call EjecutarTransicion( xBaja , "Cerrar Baja de Personal" )
        'htmlItems = codItems
        items = ""
        items = items & "<tr>"
        items = items & "<td>"& xBaja.destinatario.codigo  &"</td>"
		items = items & "<td>"& xBaja.destinatario.EnteAsociado.nombre & "</td>"
		items = items & "<td>"& xBaja.BoExtension.Cuit & "</td>"
		items = items & "<td>"& xBaja.centrocostos.name & "</td>"
		items = items & "<td>"& xBaja.BoExtension.Perfil.name & "</td>"
		items = items & "<td>"& xBaja.destinatario.fechaIngreso & "</td>"
		items = items & "<td>"& xBaja.boextension.fechabaja & "</td>"
		items = items & "<td>"& xBaja.BoExtension.FECHAENVIO & "</td>"
		items = items & "<td>"& xBaja.BoExtension.FECHARECEPCION & "</td>"
		items = items & "<td>"& xBaja.BoExtension.MOTIVOEGRESO.name & "</td>"
		items = items & "<td>"& xBaja.BoExtension.MOTIVOAFIP.name & "</td>"
        If xBaja.BoExtension.DEPOSITALIQUIDACIONFINAL Then 
            items = items & "<td>SI</td>" 
        Else
            items = items & "<td>NO</td>"
        End If
        If xBaja.BoExtension.DESCUENTAINDUMENTARIA Then 
            items = items & "<td>SI</td>" 
        Else
            items = items & "<td>NO</td>"
        End If
        If xBaja.BoExtension.DESCUENTATARJETA Then 
            items = items & "<td>SI</td>" 
        Else
            items = items & "<td>NO</td>"
        End If
      
        items = items & "</tr>"



        xItemsBaja  = xItemsBaja & " " & items 
        cantidadBajas = cantidadBajas + 1
	Next

    htmlBody = replace(htmlBody,"xFecha"    , Cstr(Date))
    htmlBody = replace(htmlBody,"xCantidad" , cantidadBajas )
    htmlBody = replace(htmlBody,"xItems"    , xItemsBaja )
	Stop

    if MsgBox("Envíar correos a: " & VBCrlf & Replace(Replace(msg,";",VbCrlf),"@iscot.com.ar",""),36,"Confirma?") <> 6 Then Exit Sub   

    call enviar_aviso(xObj, "rodrigofierrro@gmail.com" , xSubject, xBody, xadjunto,htmlBody)  
End Sub