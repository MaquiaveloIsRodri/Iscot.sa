Sub Main
	STOP
	
	set xapp = AWORKSPACE.value
	
	On Error Resume Next
	Set iMsg  	= CreateObject("CDO.Message")
	Set iConf	= CreateObject("CDO.Configuration")
	
	iConf.Load -1    ' CDO Source Defaults
	Set Flds = iConf.Fields
	With Flds
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") 			= 2
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") 			= "ca8.toservers.com"
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") 		= 25
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") 	= 1
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") 		= "mailing@iscot.com.ar"
		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") 		= "Iscot2015"
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") 			= False
		.Update
	End With
	
	strbody 	= cBody.value
	SENDDEBUG "----- HM1 "
	attachs= Split(cAdjunto.value, ";")
	
	With iMsg
		Set .Configuration = iConf
		.To 		= cDestinatario.value
		.CC 		= ""
		.BCC 		= ""
		correoDe 	= "compras@iscot.com.ar"
		set usuario = existebo(xapp, "usuario", "nombre", nombreusuario(), nil, true, false, "=")
		if not usuario is nothing then
			if Trim(usuario.direccionelectronica) <> "" then
				correoDe = usuario.direccionelectronica
			end if
		end if
		.From 		= correoDe
		.Subject 	= "OP y retenciones - Iscot Services"
		.TextBody 	= strbody
		if UBound(attachs)<>-1 then
			For Each attach In attachs
                .AddAttachment attach
            Next
		end if
		SENDDEBUG "----- HM2 " + cDestinatario.value
		.Send
	End With
	
	If Err.Number = 0 Then
		Enviar_Mail_CDO = True
		SENDDEBUG " >> Asunto: OP y retenciones - Iscot Services ----- Correo Enviado!!"
	Else
		SENDDEBUG " >> Email Error: " & Err.Description
	End If
End Sub
