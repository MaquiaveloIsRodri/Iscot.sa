' 25-04-2024 - Rodrigo Fierro.

sub main
	stop

	if Self.Estado <> "C" then
		MsgBox "El Pago debe estar Cerrado.", 48, "Aviso"
		exit sub
	end if

	' Control Email de Usuario.
	set oUsuario 		= ExisteBO(Self, "USUARIO", "NOMBRE", NombreUsuario(), nil, true, false, "=")
	if oUsuario is nothing then
		MsgBox 	"Su Usuario no esta sincronizado en Calipso." & vbNewLine & _
				"Si continua NO recibira la copia del correo enviado." & vbNewLine & _
				"Informe a Sistemas.", 48, "Aviso"
	else
		email_de		= Trim(oUsuario.DIRECCIONELECTRONICA)
		if email_de = "" then
			MsgBox 	"Usuario sin correo electronico." & vbNewLine & _
					"Si continua NO recibira la copia del correo enviado." & vbNewLine & _
					"Informe a Sistemas.", 48, "Aviso"
		end if
	end if


	usuario_nombre		= ""
	if not oUsuario is nothing then
		usuario_nombre	= oUsuario.NOMBRECOMPLETO
	end if

	set xEmpleado = nothing
	set xViewEmpl = NewCompoundView(Self, "EMPLEADO", Self.Workspace, nil, true)
	xViewEmpl.AddJoin(NewJoinSpec(xViewEmpl.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_EMPLEADO", "ID", ""), false))
	xViewEmpl.AddJoin(NewJoinSpec(NewColumnSpec("UD_EMPLEADO", "Usuario", ""), NewColumnSpec("USUARIO", "ID", ""), false))
	xViewEmpl.AddFilter(NewFilterSpec(xViewEmpl.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xViewEmpl.AddFilter(NewFilterSpec(NewColumnSpec("USUARIO", "NOMBRE", ""), " = ", NombreUsuario()))
	xViewEmpl.AddOrderColumn(NewOrderSpec(xViewEmpl.ColumnFromPath("CODIGO"), false))
	if xViewEmpl.ViewItems.Size > 0 then
		set xEmpleado = xViewEmpl.ViewItems.First.Current.BO
	end if
	usuario_puesto		= ""
	if not xEmpleado is nothing then
		if not xEmpleado.Puesto is nothing then
			usuario_puesto	= xEmpleado.Puesto.Descripcion
		end if
	end if

	' Correos del Proveedor.
	MsgBox "Seleccione el E-mail del Proveedor o escrÃ­balo en el campo 'Para'.", 64, "InformaciÃ³n"
	email_para = ""
	set xViewP = NewCompoundView(Self, "DIRECCIONELECTRONICA", Self.Workspace, nil, true)
	xViewP.AddFilter(NewFilterSpec(xViewP.ColumnFromPath("BO_PLACE"), " = ", Self.Destinatario.EnteAsociado.DireccionesElectronicas.ID))
	xViewP.AddFilter(NewFilterSpec(xViewP.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xViewP.AddBOCol("DIRECCIONELECTRONICA")
	if xViewP.ViewItems.Count > 0 then
		set xContainer1 = NewContainer()
		xContainer1.Add(xViewP)
		set xSeleccionados1 = SelectViewItems(Self.Workspace, "EMAIL PROVEEDOR", xContainer1)
		for each xItemSel1 in xSeleccionados1
			email_para = email_para & xItemSel1.BO.DIRECCIONELECTRONICA & "; "
		next
		if email_para <> "" then
			email_para = Left(email_para, Len(email_para) - 2)
		end if
	end if

	email_cuerpo		= "Estimados," & vbNewLine _
						& "A travÃ©s de este medio, se adjuntan los Certificados de las Retenciones aplicadas." & vbNewLine _
						& "Les rogamos que nos comuniquen de inmediato ante cualquier error." & vbNewLine & vbNewLine _
						& "CONFIRMAR RECEPCIÃ“N." & vbNewLine & vbNewLine _
						& "Desde ya muchas gracias!." & vbNewLine _
						& "Saludos cordiales!," & vbNewLine & vbNewLine _
						& "ISCOT SERVICES S.A." & vbNewLine _
						& usuario_nombre & vbNewLine _
						& usuario_puesto

	set xVisualVar = VisualVarEditor("PIP-NORMAL NOTIFICAR CAMBIO DE CENTRO")
	call AddVarString(xVisualVar, "05PARA", "Para:", "Email:", email_para)
	call AddVarMemo(xVisualVar, "10CUERPO", "Cuerpo:", "Email:", email_cuerpo)
	call AddVarBoolean(xVisualVar, "20Pago", "Orden de Pago", "Reportes:", true)
	call AddVarBoolean(xVisualVar, "25Gan", "Certificado Ret. de Ganancias", "Reportes:", true)
	call AddVarBoolean(xVisualVar, "30Arba", "Certificado Ret. IIBB ARBA", "Reportes:", true)
	call AddVarBoolean(xVisualVar, "35SantaFe", "Certificado Ret. IIBB Santa Fe", "Reportes:", true)
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then exit sub

	email_para 		= Trim(GetValueVisualVar(xVisualVar, "05PARA", "Email:"))
	email_cuerpo	= GetValueVisualVar(xVisualVar, "10CUERPO", "Email:")
	rpt_pago		= GetValueVisualVar(xVisualVar, "20Pago", "Reportes:")
	rpt_gan			= GetValueVisualVar(xVisualVar, "25Gan", "Reportes:")
	rpt_arba		= GetValueVisualVar(xVisualVar, "30Arba", "Reportes:")
	rpt_santafe		= GetValueVisualVar(xVisualVar, "35SantaFe", "Reportes:")

	if email_para = "" then
		MsgBox "No indicÃ³ los destinatarios, campo 'Para:'.", 48, "Aviso"
		exit sub
	end if

	set oFS 			= CreateObject("Scripting.FileSystemObject")
	set oImpresora 		= NewDic()
    xAdjunto = ""
	call RegistrarObjeto(oImpresora, "Bullzip PDF Printer", nil)

	if rpt_pago then
		sFileName 		= "C:\util\pdf\ordencompra-" & NombreUsuario() & ".pdf"
		sFileName2 		= "C:\util\pdf\Pago-" & Self.NumeroDocumento & ".pdf"
		if oFS.FileExists(sFileName) then
			oFS.DeleteFile(sFileName)
		end if
		if oFS.FileExists(sFileName2) then
			oFS.DeleteFile(sFileName2)
		end if

		call PrintBO(Self, "Orden de Pago", oImpresora)
		a 				= Esperar(3)
		pdf_listo 		= false
		for b = 1 to 20
			if oFS.FileExists(sFileName) then
				pdf_listo = true
				oFS.MoveFile sFileName, sFileName2
				exit for
			else
				a 		= Esperar(3)
			end if
		next

		if pdf_listo then
			SendDebug " >>> ENVIANDO PAGO: " & Self.NumeroDocumento
			email_asunto	= "PAGO #" & Self.NumeroDocumento & " - ORDEN"
			xAdjunto = sFileName2
			'if oFS.FileExists(sFileName2) then
			'	oFS.DeleteFile(sFileName2)
			'end if
		else
			SendDebug " >>> Fallo el PDF PAGO: " & Self.NumeroDocumento
		end if
	end if

	if rpt_gan then
		sFileName 		= "C:\util\pdf\ordencompra-" & NombreUsuario() & ".pdf"
		sFileName2 		= "C:\util\pdf\Pago-" & Self.NumeroDocumento & "-Gan.pdf"
		if oFS.FileExists(sFileName) then
			oFS.DeleteFile(sFileName)
		end if
		if oFS.FileExists(sFileName2) then
			oFS.DeleteFile(sFileName2)
		end if
		
		call PrintBO(Self, "Certificado Retencion Ganancias", oImpresora)
		a 				= Esperar(3)
		pdf_listo 		= false
		for b = 1 to 20
			if oFS.FileExists(sFileName) then
				pdf_listo = true
				oFS.MoveFile sFileName, sFileName2
				exit for
			else
				a 		= Esperar(3)
			end if
		next
		if pdf_listo then
			SendDebug " >>> ENVIANDO RET GAN PAGO: " & Self.NumeroDocumento
			email_asunto	= "PAGO #" & Self.NumeroDocumento & " - RET. GANANCIAS"
            xAdjunto&";"&sFileName2

			'if oFS.FileExists(sFileName2) then
			'	oFS.DeleteFile(sFileName2)
			'end if
		else
			SendDebug " >>> Fallo el PDF RET GAN PAGO: " & Self.NumeroDocumento
		end if
	end if

	if rpt_arba then
		sFileName 		= "C:\util\pdf\ordencompra-" & NombreUsuario() & ".pdf"
		sFileName2 		= "C:\util\pdf\Pago-" & Self.NumeroDocumento & "-ARBA.pdf"
		if oFS.FileExists(sFileName) then
			oFS.DeleteFile(sFileName)
		end if
		if oFS.FileExists(sFileName2) then
			oFS.DeleteFile(sFileName2)
		end if

		call PrintBO(Self, "Certificado Retencion IIBB ARBA", oImpresora)
		a 				= Esperar(3)
		pdf_listo 		= false
		for b = 1 to 20
			if oFS.FileExists(sFileName) then
				pdf_listo = true
				oFS.MoveFile sFileName, sFileName2
				exit for
			else
				a 		= Esperar(3)
			end if
		next

		if pdf_listo then
			SendDebug " >>> ENVIANDO RET IIBB ARBA PAGO: " & Self.NumeroDocumento
			email_asunto	= "PAGO #" & Self.NumeroDocumento & " - RET. IIBB ARBA"
            xAdjunto = xAdjunto&";"&sFileName2
			'if oFS.FileExists(sFileName2) then
			'	oFS.DeleteFile(sFileName2)
			'end if
		else
			SendDebug " >>> Fallo el PDF RET IIBB ARBA PAGO: " & Self.NumeroDocumento
		end if
	end if

	if rpt_santafe then
		sFileName 		= "C:\util\pdf\ordencompra-" & NombreUsuario() & ".pdf"
		sFileName2 		= "C:\util\pdf\Pago-" & Self.NumeroDocumento & "-StaFe.pdf"
		if oFS.FileExists(sFileName) then
			oFS.DeleteFile(sFileName)
		end if
		if oFS.FileExists(sFileName2) then
			oFS.DeleteFile(sFileName2)
		end if

		call PrintBO(Self, "Certificado Retencion IIBB Santa Fe", oImpresora)

        a 				= Esperar(3)
		pdf_listo 		= false
		for b = 1 to 20
			if oFS.FileExists(sFileName) then
				pdf_listo = true
				oFS.MoveFile sFileName, sFileName2
				exit for
			else
				a 		= Esperar(3)
			end if
		next

		if pdf_listo then
			SendDebug " >>> ENVIANDO RET IIBB STA FE PAGO: " & Self.NumeroDocumento
			email_asunto	= "PAGO #" & Self.NumeroDocumento & " - RET. IIBB STA. FE"
			xAdjunto = xAdjunto&";"&sFileName2

			'if oFS.FileExists(sFileName2) then
			'	oFS.DeleteFile(sFileName2)
			'end if
		else
			SendDebug " >>> Falla el PDF RET IIBB STA FE PAGO: " & Self.NumeroDocumento
		end if
	end if

    If xAdjunto <> "" Then 
    	'Esta funcion deberia enviar todas las retenciones una vez
    	call Enviar_Mails_Gmail(Self, email_para & "; " & email_de, email_asunto, email_cuerpo, xAdjunto)
	Else
		MsgBox "No se pudo enviar ningun Comprobante", 16, "Error"
		Exit Sub
	End If



	MsgBox "Proceso Finalizado!!.", 64, "Informacion"
end sub

function Esperar(Tiempo)
	ComienzoTiempo = Timer 
	FinTiempo = ComienzoTiempo + Tiempo
	do while FinTiempo > Timer
		if ComienzoTiempo > Timer then 
			FinTiempo = FinTiempo - 24 * 60 * 60 
		end if 
	loop 
end function
