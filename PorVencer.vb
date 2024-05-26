
' 10/04/2024 RFIERRO por vencer
sub main

    Stop
	set xContainerCobro = NewContainer()
	xContainerCobro.Add(ViewCobroCliente)
	set xSelClientes = SelectViewItems(Self.Workspace, Self.Destinatario.name, xContainerCobro)
	if xSelClientes.Size <= 0 then
		MsgBox "No selecciono ninguna Cliente.", 48, "Aviso"
		exit sub
	end if


	for each xfv in xSelClientes

            set xCon = CreateObject("adodb.connection")
            xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
            xCon.ConnectionTimeout 	= 150

            set xRs = RecordSet(xCon, "select top 1 * from producto")
            xRs.Close
            xRs.ActiveConnection.CommandTimeout = 0
            dayFilter = Year(now) & Right("0" & Month(now), 2) & Right("0" & Day(now), 2)


            set customer = existebo(self, "CLIENTE", "id",xfv.bo.DESTINATARIO_ID, nil ,true, false,"=" )

            xRs.Source = "exec SP_Cliente_ACobrar_Por_Cliente '20150131', '" & dayFilter & "', '" & customer.denominacion & "'"
	        xRs.Open

            ' Si no tenemos comprobantes, notificamos
            if xRs.EOF = true then
                If MsgBox("Desea continuar aunque el cliente "& customer.denominacion &" no tenga factura",36,"Pregunta") <> 6 Then
                    Exit Sub
                End If
            end if
            
            if xRs.EOF <> true then
            items = ""
            items2 = ""
            Set xFSO 		= CreateObject("Scripting.FileSystemObject")
            Set xArchivo    = xFSO.OpenTextFile("C:\util\html\email-RecordatorioPagoClientes.html")
            xAdjunto = ""
            htmlBody        = xArchivo.readAll()

                do while not xRs.EOF
                    if xRs("Vencida").Value = "VENCIDA" Then 
                        items2 = items2 & "<tr>"
                        items2 = items2 & "<td>"& xRs("Comprobante").Value & "</td>"
                        items2 = items2 & "<td>"& xRs("Saldo").Value & "</td>"
                        items2 = items2 & "</tr>"
                    else
                        items = items & "<tr>"
                        items = items & "<td>"& xRs("Comprobante").Value & "</td>"
                        items = items & "<td>"& xRs("Saldo").Value & "</td>"
                        items = items & "</tr>"
                    end if
	                xRs.MoveNext
                loop

                xSubject = "[NOTIFICACION COBRANZAS]" & Date() &" " & customer.denominacion
                xBody    = ""
                xcorreo = "rodrigofierrro@gmail.com;carolinaluna@iscot.com.ar"
                'xcorreo = "rodrigofierrro@gmail.com"
                htmlBody = replace(htmlBody,"xItems" , items)
                htmlBody = replace(htmlBody,"xVencidas" , items2)

                xAdjunto = ""
                call Enviar_Mail_Cobro_Cliente(self ,xcorreo , xSubject, xBody, xadjunto,htmlBody)
            end if
	next
    MsgBox "Se envio el correo correctamente", 64, "Información" 

end sub



function ViewCobroCliente()
    Set xView = NewCompoundView(Self, "TRFACTURAVENTA", Self.Workspace, Nil, False)
    xView.AddJoin(NewJoinSpec(NewColumnSpec("TRFACTURAVENTA", "DESTINATARIO", ""), NewColumnSpec("CLIENTE", "ID", ""), false))
    xView.AddJoin(NewJoinSpec(NewColumnSpec("CLIENTE", "ENTEASOCIADO", ""), NewColumnSpec("PERSONA", "ID", ""), false))


    xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ESTADO"), " = ", "C"))

    xView.AddBOCol("DESTINATARIO").Caption = "DESTINATARIO"
    xView.AddBOCol("NOMBREDESTINATARIO").Caption = "NOMBREDESTINATARIO"
    xView.AddColumn(NewColumnSpec("PERSONA", "NOMBRE", ""))
    set ViewCobroCliente = xView
end function
