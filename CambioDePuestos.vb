Sub Main
	
	
	set xapp = AWORKSPACE.value
	html = ohtml.value
	On Error Resume Next
	Set iMsg  = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	
	iConf.Load -1    ' CDO Source Defaults
	Set Flds = iConf.Fields
	With Flds
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") 			= 2
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") 			= "ca8.toservers.com"
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") 		= 25
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") 	= 1
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") 		= "calipso@iscot.com.ar"
		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") 		= "calipsoavisos"
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") 			= False
		.Update
	End With
	
	strbody = cBody.value
	SENDDEBUG "----- HM1 "
	'Path_adjunto = "C:\WinInst.txt"
	'If Path_Adjunto <> "" Then
	'   iMsg.AddAttachment (Path_Adjunto)
	'End If
	attach = cAdjunto.value
	
	With iMsg
		Set .Configuration = iConf
		.To 		= cDestinatario.value
		.CC 		= ""
		.BCC 		= ""
		correoDe 	= cCorreo.Value
		set usuario = existebo(xapp, "usuario", "nombre", nombreusuario(), nil,true, false,"=" )
		
		.HTMLBody = html
		
		if not usuario is nothing then
			if Trim(usuario.direccionelectronica) <> "" then
				correoDe = usuario.direccionelectronica
			end if
		end if
		.From 		= correoDe
		.Subject 	= casunto.value
		'.TextBody 	= strbody
		if Trim(attach) <> "" then
			.AddAttachment attach
		end if
		SENDDEBUG "----- HM2 " + cDestinatario.value
		.Send
	End With
	
	If Err.Number = 0 Then
		Enviar_Mail_CDO = True
		SENDDEBUG " >> Asunto: " & casunto.value & " ----- Correo Enviado!!"
		'MsgBox "Asunto: " & casunto.value & vbNewLine & "Correo Enviado!!.", 64, "E-MAIL"
	Else
		SENDDEBUG " >> Email Error: " & Err.Description
		'MsgBox Err.Description, vbCritical, " Error al enviar el E-mail "
	End If
End Sub
