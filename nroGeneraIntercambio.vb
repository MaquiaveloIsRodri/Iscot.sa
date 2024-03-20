Sub Main
stop

if aObject.NOTA = "EN INTERCAMBIO" then
	If aObject.CONDICION Is Nothing Then
   		addfailedconstraint "No indico la condicion",1
	   	Exit Sub
	End If

	If aObject.motivo.id = "{A4CA0C19-0D12-4ED0-85B1-65AB82C47278}" OR aObject.motivo.id = "{D1401BBF-2F9B-4BD0-AC3C-CA41FC995690}" Then
	If MsgBox("Genera automáticamente el intercambio ? " , 36, "Pregunta") <> 6 Then Exit Sub
        set xIntercambio = crearbo("UD_HISTORICOCELULAR",self)
        Set xIntercambio.empleado = aObject.destinatario
        Set oMotivos = ExisteBo(Self,"TipoClasificador","ID","{D36C3DB2-869F-4B5C-B2E1-99E819748769}",nil,True,False,"=")
 		Set oResicion = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "id", "{2E1F4B42-3133-44AB-B423-F1A965616444}", oMotivos.Valores , True, False, "=")
        Set xIntercambio.motivo = oResicion
	    xIntercambio.FECHADEENVIO = aObject.boextension.FECHAENVIOTLC
	    xIntercambio.FECHAVTO = aObject.boextension.FECHALIMITERESPUESTA
	    xIntercambio.NOTA = "EN INTERCAMBIO"
	    showBo(xIntercambio)
	    If aObject.workspace.intransaction then aObject.workspace.commit
	End If
end if


End Sub
