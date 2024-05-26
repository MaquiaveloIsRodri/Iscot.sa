' 28/04/2022
' MU: Factura Enviar varias en 1 mail.
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
	
	' Ordenes de compra.
	set xContainerOC = NewContainer()
	xContainerOC.Add(ViewTROC())
	set xSelTROC = SelectViewItems(Self.Workspace, "OC de " & Self.Destinatario.EnteAsociado.Nombre, xContainerOC)
	
    if xSelTROC.Size <= 0 then
		MsgBox "No seleccionó ningun comprobante.", 48, "Aviso"
		exit sub
	end if
	
	' Correos a enviar.
	correos = ""
	set xContainer = NewContainer()
	xContainer.Add(ObtenerView())
	set xSeleccionados = SelectViewItems(Self.Workspace, "Correos de " & Self.Destinatario.EnteAsociado.Nombre, xContainer)
	for each xCorreo in xSeleccionados
		correos = correos & xCorreo.BO.EMAIL & "; "
	next
    
	' Cuerpo E-mail.
    cuerpoEmail = ""
    cuerpoEmail = "Estimados," & vbCrLf _
                & "                      Buenos días. Les hacemos llegar la/s " & comprobante & "/s correspondiente/s al servicio de limpieza." & vbCrLf & vbCrLf _
                & "Por favor, confirmar correcta recepcion." & vbCrLf _
                & "Ante cualquier duda o inconveniente, no duden en comunicarse con nosotros." & vbCrLf & vbCrLf _
                & "Desde ya muchas gracias." & vbCrLf _
                & "Saludos Cordiales."

	
    ' VisualVar.
	set xVisualVar = VisualVarEditor("ENVIAR MULTIPLES COMPROBANTES POR MAIL")
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
	
	auxPDF = "ordencompra-" & NombreUsuario()

	
	set fs 			= CreateObject("Scripting.FileSystemObject")
	set Impresoras 	= NewDic( )
	call RegistrarObjeto( Impresoras, "Bullzip PDF Printer", nil )
	
	' Enviar correos.
	call ProgressControl(Self.Workspace, "ENVIANDO Comprobantes..." , 0, 10)
	
	pdfsMal = ""
    xSubject	= comprobante
    xBody		= cuerpoEmail
    xCorreos	= correos & "; " & correoUsuario
    xAdjunto = ""
	for each xCpte in xSelTROC

		call ProgressControlAvance(Self.Workspace, "Enviando..." & comprobante & vbNewLine & "OC" & xCpte.BO.NumeroDocumento)

		sFileName 	= "C:\util\pdf\" & auxPDF & ".pdf"
		sFileName2 	= "C:\util\pdf\OC-" & xCpte.BO.NumeroDocumento & ".pdf"
		
        if fs.FileExists(sFileName) then
			fs.DeleteFile(sFileName)
		end if
		if fs.FileExists(sFileName2) then
			fs.DeleteFile(sFileName2)
		end if

        reporte = "Orden de Compra"
		
		call PrintBO( xCpte.BO, reporte , Impresoras )
		a = Esperar(3)

		EstaPDF = false
		for b = 1 to 20
			if fs.FileExists(sFileName) then
				EstaPDF = true
				sFileName2 	= "C:\util\pdf\OC-" & xCpte.BO.NumeroDocumento & ".pdf"
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

        set xMedio						= InstanciarBO("{85F05446-DD2C-421B-B20D-2A3F4B680EEE}", "ITEMTIPOCLASIFICADOR", Self.Workspace)
        call WorkSpaceCheck(xCpte.BO.Workspace)

	next

    If xAdjunto <> "" Then 
	   xAdjunto = Right(xAdjunto, Len(xAdjunto)-1)
	   Call EnviarEmailSinMensajeAdjuntos(self, xCorreos, xSubject, xBody, xAdjunto)
	Else
		MsgBox "No se pudo enviar ningún Comprobante", 16, "Error"
		Exit Sub
	End If
		
	if pdfsMal <> "" then
		MsgBox "Algunos Comprobantes no se enviaron:" & vbNewLine & pdfsMal, 16, "Error"
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
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ",self.destinatario.enteasociado.direccioneselectronicas.id))
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

function ViewTROC()
	set xViewOC = NewCompoundView(Self, "TRORDENCOMPRA", Self.Workspace, nil, true)
	xViewOC.AddJoin(NewJoinSpec(xViewOC.ColumnFromPath("CENTROCOSTOS"), NewColumnSpec("CENTROCOSTOS", "ID", ""), false))
	xViewOC.AddJoin(NewJoinSpec(xViewOC.ColumnFromPath("FLAG"), NewColumnSpec("FLAG", "ID", ""), false))
	xViewOC.AddJoin(NewJoinSpec(xViewOC.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_ORDENCOMPRA", "ID", ""), false))
	xViewOC.AddFilter(NewFilterSpec(xViewOC.ColumnFromPath("ESTADO"), " = ", "C"))
	xViewOC.AddFilter(NewFilterSpec(xViewOC.ColumnFromPath("FLAG"), " <> ", "{C10833F4-5E1E-47DA-803E-5FBF135BEA51}")) ' Anulado
	xViewOC.AddFilter(NewFilterSpec(xViewOC.ColumnFromPath("FLAG"), " <> ", "{BC12C8D2-C060-4026-9447-A130A688E599}")) ' Facturada
	xViewOC.AddFilter(NewFilterSpec(xViewOC.ColumnFromPath("FLAG"), " <> ", "{3F8FBBA5-200D-492F-9DE7-962ECAAEAA1C}")) ' Pendiente de Entrega
	xViewOC.AddFilter(NewFilterSpec(xViewOC.ColumnFromPath("DESTINATARIO"), " = ", Self.Destinatario.ID))
	xViewOC.AddBOCol("NOMBRE").Caption										= " "
	xViewOC.AddBOCol("FLAG").Caption										= " "
	xViewOC.AddBOCol("NOMBRE").Caption										= "Transacción"
	xViewOC.AddColumn(NewColumnSpec("CENTROCOSTOS", "NOMBRE", "")).Caption	= "Centro Costos"
	xViewOC.AddColumn(NewColumnSpec("FLAG", "DESCRIPCION", "")).Caption		= "Flag"
	xViewOC.AddColumn(NewColumnSpec("UD_ORDENCOMPRA", "PAGOSEMANAL", "")).Caption		= "Pago Semanal"
	xViewOC.AddBOCol("USUARIO").Caption					 					= "Usuario"
	xViewOC.AddOrderColumn(NewOrderSpec(xViewOC.ColumnFromPath("NUMERODOCUMENTO"), true))
	set ViewTROC = xViewOC
end function
