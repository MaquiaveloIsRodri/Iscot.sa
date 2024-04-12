' CREADO: 14/02/2014 - Jose Fantasia.
' OTE: Comprobantes de Venta.
' MU: Enviar <Comprobante> como PDF.
sub main
	stop
	
	' Controles.
	if Self.Estado <> "C" then
		MsgBox "Sólo se pueden enviar Comprobantes Cerrados.", 48, "Aviso"
		exit sub
	end if
	if Left(Self.NumeroDocumento, 4) = "9999" or Self.NumeroDocumento = "" then
		MsgBox "No se puede Enviar por Mail un Comprobante sin CAE.", 48, "Aviso"
		exit sub
	end if
	if not Self.BOEXTENSION.MedioEnvio is nothing then
		if MsgBox("El comprobante ya tiene un Medio. ¿Desea sobrescribirlo?", 292, "Pregunta") <> 6 then exit sub
	end if

	' Correo Usuario.
	set xUsuario = ExisteBO(Self, "USUARIO", "NOMBRE", NombreUsuario, nil, true, false, "=")
	if xUsuario is nothing then
		MsgBox 	"Su Usuario no está sincronizado en Calipso." & vbNewLine & _
				"Si continúa NO recibirá la copia del correo enviado." & vbNewLine & _
				"Informe a Sistemas.", 48, "Aviso"
	end if
	correoUsuario = Trim(xUsuario.DIRECCIONELECTRONICA)
	if correoUsuario = "" then
		MsgBox 	"Usuario sin correo electrónico." & vbNewLine & _
				"Si continúa NO recibirá la copia del correo enviado." & vbNewLine & _
				"Informe a Sistemas.", 48, "Aviso"
	end if

	res = MsgBox("¿Confirma que desea enviar por E-mail?" & vbCrlf & vbCrlf _
		& Self.Nombre & vbCrLf _
		& Self.NOMBREDESTINATARIO & vbCrlf _
		& Self.TotalTr, 292, "Confirmar")
	if res <> 6 then exit sub

	tipoTr = ClassName(Self)

	' Bullzip.
	'if UCase(NombreUsuario()) = "CLUNA" or UCase(NombreUsuario()) = "ANTOBOCCO" or UCase(NombreUsuario()) = "FGOZZARINO" then
	'    auxPDF = "facturaventa"
	'else
		auxPDF = "ordencompra-" & NombreUsuario()
	'end if
	
	' Correos del Cliente.
	correos = ""
	set xContainer = NewContainer()
	xContainer.Add(ObtenerView())
	set xSeleccionados = SelectViewItems(Self.Workspace, tipoTr, xContainer)
	for each xCorreo in xSeleccionados
		correos = correos & xCorreo.BO.EMAIL & "; "
	next
	
	' Cuerpo E-mail.
	' 13/01/2020 - Cuerpos de ND y NC.
	cuerpoEmail = ""
	select case tipoTr
	case "TRFACTURAVENTA"
		cuerpoEmail = "Estimados," & vbCrLf _
					& "                      Buenos días. Les hacemos llegar la factura correspondiente al servicio de limpieza." & vbCrLf & vbCrLf _
					& "Por favor, confirmar correcta recepcion." & vbCrLf _
					& "Ante cualquier duda o inconveniente, no duden en comunicarse con nosotros." & vbCrLf & vbCrLf _
					& "Desde ya muchas gracias." & vbCrLf _
					& "Saludos Cordiales."
	case "TRDEBITOVENTA"
		cuerpoEmail = "Estimados," & vbCrLf _
					& "                      Buenos días. Les hacemos llegar la Nota de Debito correspondiente al servicio de limpieza." & vbCrLf & vbCrLf _
					& "Por favor, confirmar correcta recepcion." & vbCrLf _
					& "Ante cualquier duda o inconveniente, no duden en comunicarse con nosotros." & vbCrLf & vbCrLf _
					& "Desde ya muchas gracias." & vbCrLf _
					& "Saludos Cordiales."
	case else
		cuerpoEmail = "Estimados," & vbCrLf _
					& "                      Buenos días. Les hacemos llegar la Nota de Crédito correspondiente al servicio de limpieza." & vbCrLf & vbCrLf _
					& "Por favor, confirmar correcta recepcion." & vbCrLf _
					& "Ante cualquier duda o inconveniente, no duden en comunicarse con nosotros." & vbCrLf & vbCrLf _
					& "Desde ya muchas gracias." & vbCrLf _
					& "Saludos Cordiales."
	end select

	' VisualVar.
	set xVisualVar = VisualVarEditor("ENVIAR COMPROBANTE COMO PDF")
	call AddVarString(xVisualVar, "00CORREOS", "Para:", "Datos", correos)
	call AddVarMemo(xVisualVar, "10CUERPO", "Cuerpo Email:", "Datos", cuerpoEmail)
	aceptar = ShowVisualVar( xVisualVar )
	if not aceptar then	exit sub

	correos		= Trim(GetValueVisualVar(xVisualVar, "00CORREOS", "Datos"))
	cuerpoEmail	= GetValueVisualVar(xVisualVar, "10CUERPO", "Datos")

	if correos = "" then
		MsgBox "No indicó las Direcciones de Correo.", 48, "Aviso"
		exit sub
	end if

	aux_tipo	= ""
	if Self.BOEXTENSION.TipoComprobante.Nombre = "A" then
		select case tipoTr
		case "TRFACTURAVENTA"
			reporte 	= "Factura de Venta A-"
			aux_tipo 	= "FA-A-"
		case "TRDEBITOVENTA"
			reporte 	= "Debito de Venta A"
			aux_tipo 	= "ND-A-"
		case else
			reporte 	= "Credito de Venta A"
			aux_tipo 	= "NC-A-"
		end select
	else
		select case tipoTr
		case "TRFACTURAVENTA"
			reporte 	= "Factura de Venta B-"
			aux_tipo 	= "FA-B-"
		case "TRDEBITOVENTA"
			reporte 	= "Debito de Venta B-"
			aux_tipo 	= "ND-B-"
		case else
			reporte 	= "Credito de Venta B"
			aux_tipo 	= "NC-B-"
		end select
	end if

	set fs 		= CreateObject("Scripting.FileSystemObject")
	sFileName 	= "C:\util\pdf\" & auxPDF & ".pdf"
	sFileName2 	= "C:\util\pdf\" & aux_tipo & Self.NumeroDocumento & ".pdf"
	if fs.FileExists(sFileName) then
		fs.DeleteFile(sFileName)
	end if
	if fs.FileExists(sFileName2) then
		fs.DeleteFile(sFileName2)
	end if
	set fs		= nothing
	
	set Impresoras = NewDic( )
	call RegistrarObjeto( Impresoras, "Bullzip PDF Printer", nil )
	call PrintBO( Self, reporte , Impresoras )
	
	a = Esperar(3)
	EstaPDF = false
	set fs 		= CreateObject("Scripting.FileSystemObject")
	sFileName 	= "C:\util\pdf\" & auxPDF & ".pdf"
	sFileName2 	= "C:\util\pdf\" & aux_tipo & Self.NumeroDocumento & ".pdf"
	
	for b = 1 to 20
		if fs.FileExists(sFileName) then
			EstaPDF = true
			sFileName2 = "C:\util\pdf\" & aux_tipo & Self.NumeroDocumento & ".pdf"
			fs.MoveFile sFileName, sFileName2
			exit for
		else
			a = Esperar(3)
		end if
	next
	
	if EstaPDF then
		xSubject	= aux_tipo & Self.NumeroDocumento
		xBody		= cuerpoEmail
		xCorreos	= correos & "; " & correoUsuario
		xAdjunto 	= sFileName2

		call Enviar_Mails_Gmail(Self, xCorreos, xSubject, xBody, xAdjunto)
		MsgBox "Correo Enviado!!.", 64, "Información"

		if fs.FileExists(sFileName) then
			fs.DeleteFile(sFileName)
		end if
		if fs.FileExists(sFileName2) then
			fs.DeleteFile(sFileName2)
		end if

		set xMedio						= InstanciarBO("{85F05446-DD2C-421B-B20D-2A3F4B680EEE}", "ITEMTIPOCLASIFICADOR", Self.Workspace)
		Self.BOEXTENSION.FechaEnvio 	= Now
		Self.BOEXTENSION.MedioEnvio		= xMedio
		call WorkSpaceCheck(Self.Workspace)
	end if

	set fs = nothing
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

function ObtenerView()
	set xView = NewCompoundView(Self, "UD_EMAIL", Self.Workspace, nil, true)
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("TIPO"), NewColumnSpec("TIPODIRECCIONELECTRONICA", "ID", ""), false))
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("SECTOR"), NewColumnSpec("ITEMTIPOCLASIFICADOR", "ID", ""), true))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", Self.Destinatario.BOEXTENSION.CORREOS.ID))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("TIPO"), " = ", "D11F43B6-76A7-4BBE-9DA3-A5B7738824A1"))	  	' E-Mail para envío de Facturas de Venta.
	xView.AddBOCol("TIPO").Caption						= " "
	xView.AddBOCol("SECTOR").Caption					= " "
	xView.AddColumn(NewColumnSpec("TIPODIRECCIONELECTRONICA", "TIPODIRECCIONELECTRONICA", "")).Caption	= "Tipo"
	xView.AddBOCol("EMAIL").Caption						= "Email"
	xView.AddBOCol("NOMBRECOMPLETO").Caption			= "Nombre y Apellido"
	xView.AddColumn(NewColumnSpec("ITEMTIPOCLASIFICADOR", "NOMBRE", "")).Caption						= "Sector"
	xView.AddBOCol("NOTA").Caption						= "Nota"
	xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NOMBRECOMPLETO"), true))
	set ObtenerView = xView
end function
