Sub Main
stop
    If aObject.Workspace.InState(3) = true Then 
        if aObject.NOTA <> "EN INTERCAMBIO" then
            If aObject.boextension.CONDICION Is Nothing Then
                addfailedconstraint "No indico la condicion",1
                Exit Sub
            End If

            If aObject.boextension.CONDICION.id  = "{A4CA0C19-0D12-4ED0-85B1-65AB82C47278}" OR  aObject.boextension.CONDICION.id = "{D1401BBF-2F9B-4BD0-AC3C-CA41FC995690}" Then
                If MsgBox("Genera automáticamente el intercambio ? " , 36, "Pregunta") <> 6 Then Exit Sub
                    stop
                    set xIntercambio = crearbo("UD_INTERCAMBIOEPISTOLAR",aObject)
                    aobject.bo_place.bo_owner.boextension.intercambiopistelar.add(xIntercambio)
                    Set xIntercambio.empleado = aObject.destinatario
                    Set oMotivos              = ExisteBo(aObject,"TipoClasificador","ID","{D36C3DB2-869F-4B5C-B2E1-99E819748769}",nil,True,False,"=")
                    Set oResicion = ExisteBo(aObject, "ITEMTIPOCLASIFICADOR", "id", "{2E1F4B42-3133-44AB-B423-F1A965616444}", oMotivos.Valores , True, False, "=")
                    Set xIntercambio.motivo = oResicion
                    xIntercambio.FECHADEENVIO = aObject.boextension.FECHAENVIOTLC
                    xIntercambio.FECHAVTO = aObject.boextension.FECHALIMITERESPUESTA
                    aObject.NOTA = "EN INTERCAMBIO"
                    showBo(xIntercambio)
                    If aObject.workspace.intransaction then aObject.workspace.commit
            End If
        end if
    End If
End Sub

