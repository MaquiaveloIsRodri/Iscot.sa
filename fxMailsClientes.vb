sub main

    Stop
    'Container de clientes
    set xVisualVar = VisualVarEditor("Seleccione al cliente")
    Set xView = NewCompoundView(Self, "CLIENTE", Self.Workspace, Nil, False)
    xView.AddJoin(NewJoinSpec(NewColumnSpec("CLIENTE", "ENTEASOCIADO", ""), NewColumnSpec("PERSONA", "ID", ""), false))
	xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
	xView.AddColumn(NewColumnSpec("PERSONA", "NOMBRE", ""))

    Set xContainerCustomer = NewContainer()
    xContainerCustomer.Add (xView)

    Call AddVarObj(xVisualVar,  "00Cliente", "Cliente", "Parametros:", nothing ,xContainerCustomer, Self.WorkSpace )

	accept = ShowVisualVar(xVisualVar)
	if not accept then exit sub


    set customer = GetValueVisualVar( xVisualVar, "00Cliente", "Parametros:" )

    ' Si no se encuentra el cliente sale de la funcion
    if customer is nothing then
        MsgBox "No selecciona un cliente para enviar el mail.", 48, "Aviso"
        exit sub
    end if

	set xCon = CreateObject("adodb.connection")
	xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
	xCon.ConnectionTimeout 	= 150

	set xRs = RecordSet(xCon, "select top 1 * from producto")
	xRs.Close
	xRs.ActiveConnection.CommandTimeout = 0
    dayFilter = Year(now) & Right("0" & Month(now), 2) & Right("0" & Day(now), 2)
    ' Obtenemos la liberacion de las facturas vencidas
    xRs.Source = "exec SP_Cliente_ACobrar_Vencidas_PorCliente '20150131', '" & dayFilter & "', '" & customer.denominacion & "'"
	xRs.Open

    ' Si no tenemos comprobantes, notificamos
    if xRs.EOF = true then 
        MsgBox "El cliente no debe ninguna factura.", 48, "Aviso"
        exit sub
    end if


    items = ""
    Set xFSO 		= CreateObject("Scripting.FileSystemObject")
    Set xArchivo    = xFSO.OpenTextFile("C:\util\html\email-FacturasVencidas.html")
    xAdjunto = ""
    htmlBody        = xArchivo.readAll()

    do while not xRs.EOF
        items = items & "<tr>"
	    items = items & "<td>"& xRs("Comprobante").Value & "</td>"
	    items = items & "<td>"& xRs("Saldo").Value & "</td>"
	    items = items & "</tr>"
        sFileName2 = selectModel(xRs("Comprobante").Value)
        xAdjunto 	= xAdjunto&";"&sFileName2
	    xRs.MoveNext
    loop

    xSubject = "[NOTIFICACION COBRANZAS]" & Date()
    xBody    = ""
    xcorreo = "rodrigofierrro@gmail.com"
    htmlBody = replace(htmlBody,"xItems" , items)
    htmlBody = replace(htmlBody,"xCliente" , customer.denominacion)
    xAdjunto = Right(xAdjunto, Len(xAdjunto)-1)
    call EnviarEmailSinMensajeAdjuntos(self, xcorreo , xSubject , xBody, xAdjunto)
    call enviar_aviso_sinmsg(self, xcorreo , xSubject, xBody, xadjunto,htmlBody)


end sub

' Selecciona segun lo que obtengamos en la consulta
function selectModel(Voucher)
    Select Case true
        Case InStr(Voucher, "FaVen") > 0 'Entra si tiene dicha FaVen "Factura venta"
            'instanciamos las facturas
            set FaVen = ExisteBo(Self, "TRFACTURAVENTA", "NOMBRE", Voucher, nil, true, false, "=")
            auxPDF = "ordencompra-" & NombreUsuario()

            if FaVen.BOEXTENSION.TipoComprobante.Nombre = "A" then
                report 	= "Factura de Venta A-"
                aux_type 	= "FA-A-"

                'mandamos a generar los pdf
                selectModel = printVoucher(FaVen,report , aux_type,auxPDF)
            else
                report 	= "Factura de Venta B-"
                aux_type 	= "FA-B-"

                'mandamos a generar los pdf
                selectModel = printVoucher(FaVen,report,aux_type,auxPDF)
	        end if


        Case InStr(Voucher, "NoCrVe") > 0 'Entra si tiene dicha NoCrVe "Nota credito venta"
            set NoCrVe = ExisteBo(Self, "TRCREDITOVENTA", "NOMBRE", Voucher, nil, true, false, "=")
           auxPDF = "ordencompra-" & NombreUsuario()

            if NoCrVe.BOEXTENSION.TipoComprobante.Nombre = "A" then
                report 	= "Credito de Venta A"
                aux_type 	= "NC-A-"

                'mandamos a generar los pdf
                selectModel = printVoucher(NoCrVe,report,aux_type,auxPDF)
            else
                report 	= "Credito de Venta B"
                aux_type 	= "NC-B-"

                'mandamos a generar los pdf
                selectModel = printVoucher(NoCrVe,report,aux_type,auxPDF)
	        end if
        case Else
            MsgBox "Contactarse con sistema, ya que el comprobante no pudo ser procesado", 64, "InformaciÃƒÂ³n"
		end select
end function



'Esta funcion se encarga de instanciar la impresora
function printVoucher(VoucherFilter,report , aux_type,auxPDF )
    
	set fs 		= CreateObject("Scripting.FileSystemObject")
	sFileName 	= "C:\util\pdf\" & auxPDF & ".pdf"
	sFileName2 	= "C:\util\pdf\" & aux_type & VoucherFilter.NumeroDocumento & ".pdf"
	if fs.FileExists(sFileName) then
		fs.DeleteFile(sFileName)
	end if
	if fs.FileExists(sFileName2) then
		fs.DeleteFile(sFileName2)
	end if

	set fs		= nothing

	set Impresoras = NewDic( )
	call RegistrarObjeto( Impresoras, "Bullzip PDF Printer", nil )
	call PrintBO( VoucherFilter, report , Impresoras )

	a = wait(3)
	EstaPDF = false
	set fs 		= CreateObject("Scripting.FileSystemObject")
	sFileName 	= "C:\util\pdf\" & auxPDF & ".pdf"
	sFileName2 	= "C:\util\pdf\" & aux_type & VoucherFilter.NumeroDocumento & ".pdf"

	for b = 1 to 20
		if fs.FileExists(sFileName) then
			EstaPDF = true
			sFileName2 = "C:\util\pdf\" & aux_type & VoucherFilter.NumeroDocumento & ".pdf"
			fs.MoveFile sFileName, sFileName2
			exit for
		else
			a = wait(3)
		end if
	next
    if EstaPDF then
		printVoucher = sFileName2
	end if
end function

function wait(time)
	startTime = timer
	finishTime = startTime + time
	do while finishTime > timer
		if startTime > timer then 
			finishTime = finishTime - 24 * 60 * 60 
		end if
	loop
end function
