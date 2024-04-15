'rfierro
Sub Main
Stop
    Set xFSO 		= CreateObject("Scripting.FileSystemObject")
    Set xArchivo    = xFSO.OpenTextFile("C:\util\html\email-ListoParaCompras.html")
    htmlBody        = xArchivo.readAll()

    xSubject = "[AVISO BAJAS] " & Date() 
    xBody    = ""
    xadjunto = ""
    items    = ""

	items = items & "<tr>"
	items = items & "<td>"& Transaccion.name & "</td>"
	items = items & "<td>"& Transaccion.FechaActual & "</td>"
	items = items & "<td>"& Transaccion.BOEXTENSION.SOLICITANTE.DESCRIPCION & "</td>"
	items = items & "<td>"& Transaccion.centrocostos.nombre & "</td>"
	items = items & "<td>"& Transaccion.Detalle & "</td>"
	items = items & "<td>"& Transaccion.BOEXTENSION.Urgente & "</td>"
	items = items & "<td>"& Transaccion.BOEXTENSION.FechaEntrega & "</td>"
	items = items & "</tr>"



    htmlBody = replace(htmlBody,"xItems", items )
    call enviar_aviso(Transaccion, "rodrigofierrro@gmail.com;antonellabocco@iscot.com.ar" , xSubject, xBody, xadjunto,htmlBody)  

end sub
