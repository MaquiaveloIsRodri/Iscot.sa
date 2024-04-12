' 15/02/2022
' MU: Orden-Servicio Enviar varias en 1 mail.
sub main
	stop
	
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
	
	' Ordenes del Cliente.
	set xContainerOC = NewContainer()
	xContainerOC.Add(ViewOrdenes())
	set xSelOC = SelectViewItems(Self.Workspace, "Ordenes de " & Self.Destinatario.EnteAsociado.Nombre, xContainerOC)
	if xSelOC.Size <= 0 then
		MsgBox "No seleccionó ningúna Orden.", 48, "Aviso"
		exit sub
	end if
	
	' Correos del Cliente.
	correos = ""
	set xContainer = NewContainer()
	xContainer.Add(ObtenerView())
	set xSeleccionados = SelectViewItems(Self.Workspace, "Correos de " & Self.Destinatario.EnteAsociado.Nombre, xContainer)
	for each xCorreo in xSeleccionados
		correos = correos & xCorreo.BO.EMAIL & "; "
	next
	
	' Cuerpo E-mail.
	cuerpoEmail = "Estimados," & vbCrLf _
				& "                      Buenos días. Les hacemos llegar la/s orden/es correspondiente/s al servicio de limpieza." & vbCrLf & vbCrLf _
				& "Por favor, confirmar correcta recepción." & vbCrLf _
				& "Ante cualquier duda o inconveniente, no duden en comunicarse con nosotros." & vbCrLf & vbCrLf _
				& "Desde ya muchas gracias." & vbCrLf _
				& "Saludos Cordiales."
	
	' VisualVar.
	set xVisualVar = VisualVarEditor("ENVIAR MULTIPLES OS POR MAIL")
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
	
	'if UCase(NombreUsuario()) = "CLUNA" or UCase(NombreUsuario()) = "JCHAPARTEGUI" then
	'    auxPDF = "facturaventa"
	'else
		auxPDF = "ordencompra-" & NombreUsuario()
	'end if
	
	set fs 			= CreateObject("Scripting.FileSystemObject")
	set Impresoras 	= NewDic( )
	call RegistrarObjeto( Impresoras, "Bullzip PDF Printer", nil )
	
	' 11/04/2019 - Se puede elegir cuál PDF enviar.
	reporte		= "Orden de Servicio"
	set xVisualVar = VisualVarEditor("IMPRIMIR PDF")
	call AddVarEnum(xVisualVar, "00PDF", "Reporte", "Parametros:", "Orden de Servicio", ListaPDF())
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then exit sub
	reporte = GetValueVisualVar(xVisualVar, "00PDF", "Parametros:")
	
	' Enviar correos.
	call ProgressControl(Self.Workspace, "ENVIANDO ORDENES..." , 0, 10)
	
	pdfsMal = ""
    xSubject	= "Orden de Servicio"
    xBody		= cuerpoEmail
    xCorreos	= correos & "; " & correoUsuario
    xAdjunto = ""
	for each xCpte in xSelOC
		call ProgressControlAvance(Self.Workspace, "Enviando..." & vbNewLine & "OS: " & xCpte.BO.NumeroDocumento)
		
		sFileName 	= "C:\util\pdf\" & auxPDF & ".pdf"
		sFileName2 	= "C:\util\pdf\OrSer-" & xCpte.BO.NumeroDocumento & ".pdf"
		if fs.FileExists(sFileName) then
			fs.DeleteFile(sFileName)
		end if
		if fs.FileExists(sFileName2) then
			fs.DeleteFile(sFileName2)
		end if
		
		call PrintBO( xCpte.BO, reporte , Impresoras )
		a = Esperar(3)
		EstaPDF = false
		for b = 1 to 20
			if fs.FileExists(sFileName) then
				EstaPDF = true
				sFileName2 	= "C:\util\pdf\OrSer-" & xCpte.BO.NumeroDocumento & ".pdf"
				fs.MoveFile sFileName, sFileName2
				exit for
			else
				a = Esperar(3)
			end if
		next

		if EstaPDF then
			xAdjunto 	= xAdjunto&";"&sFileName2
		else
			pdfsMal = pdfsMal & " - " & xCpte.BO. NumeroDocumento & vbNewLine
		end if
	next

    If xAdjunto <> "" Then 
	   xAdjunto = Right(xAdjunto, Len(xAdjunto)-1)
	   Call EnviarEmailSinMensajeAdjuntos(self, xCorreos, xSubject, xBody, xAdjunto)
	Else
		MsgBox "No se pudo enviar ninguna Orden", 16, "Error"
		Exit Sub
	End If

	if pdfsMal <> "" then
		MsgBox "Algunas Ordenes no se enviaron:" & vbNewLine & pdfsMal, 16, "Error"
		exit sub
	end if
	set fs = nothing
	call ProgressControlFinish(Self.Workspace)
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
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("TIPO"), " = ", "E280868E-A980-4D81-8E34-3E1C4D0B2E18"))	  	' E-mail para Orden de Servicio.
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

function ViewOrdenes()
	set xView = NewCompoundView(Self, "TRORDENVENTA", Self.Workspace, nil, true)
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("FLAG"), NewColumnSpec("FLAG", "ID", ""), false))
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_ORDENSERVICIO", "ID", ""), false))
	xView.AddJoin(NewJoinSpec(xView.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), false))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("DESTINATARIO"), " = ", Self.Destinatario.ID))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ESTADO"), " = ", "C"))		' Cerrada.
	xView.AddBOCol("NOMBRE").Caption													= " "
	xView.AddBOCol("FLAG").Caption														= " "
	xView.AddBOCol("NOMBRE").Caption													= "Transacción"
	xView.AddBOCol("NOMBREDESTINATARIO").Caption 										= "Cliente"
	xView.AddColumn(NewColumnSpec("CENTROCOSTOS", "NOMBRE", "")).Caption				= "Centro de Costos"
	xView.AddBOCol("NOTA").Caption 					   	 								= "Cód. Autoriza."
	xView.AddColumn(NewColumnSpec("UD_ORDENSERVICIO", "SOLICITANTE", "")).Caption		= "Solicitante"
	xView.AddColumn(NewColumnSpec("UD_ORDENSERVICIO", "FIRMANTE", "")).Caption			= "Firmante"
	xView.AddBOCol("USUARIO").Caption					 								= "Usuario"
	xView.AddOrderColumn(NewOrderSpec(xView.ColumnFromPath("NUMERODOCUMENTO"), true))
	set ViewOrdenes = xView
end function

function ListaPDF()
	set xDict 	= NewDic()
	set xBucket = NewBucket()
	call RegistrarObjetoBucket(xDict, "Orden de Servicio", "Orden de Servicio")
	call RegistrarObjetoBucket(xDict, "Orden de Servicio - Solo Q", "Orden de Servicio - Solo Q")
	call RegistrarObjetoBucket(xDict, "Orden de Servicio Resumida", "Orden de Servicio Resumida")

	set ListaPDF = xDict
end function
