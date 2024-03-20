Sub Main

If aObject.motivo Is Nothing Then
	addfailedconstraint "Motivo no ingresado",1
	Exit Sub
End If

If aObject.motivo.id = "{2E1F4B42-3133-44AB-B423-F1A965616444}" OR aObject.motivo.id = "{D07C37A4-3ED1-4E9B-BA61-5CB0EC991258}" Then
If MsgBox("Genera automáticamente la baja ? " , 36, "Pregunta") <> 6 Then Exit Sub
	  	Set xBaja = CrearTransaccion("BAJA", InstanciarBO( "{32DC2522-003F-4694-B193-5AE59009B6AF}", "UOSOLICITUD", aObject.Workspace ))
	   	Set xBaja.destinatario = aObject.empleado
	   	Set oMotivos              = ExisteBo(Self,"TipoClasificador","ID","{D36C3DB2-869F-4B5C-B2E1-99E819748769}",nil,True,False,"=")
 		Set oResicion = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", "{2E1F4B42-3133-44AB-B423-F1A965616444}", oMotivos.Valores , True, False, "=")

	   xBaja.boextension.fechaenvio = aObject.fechadeenvio
	   xBaja.boextension.fecharecepcion = aObject.fecharecepcion
	   xBaja.Nota = "Generado por Intercambio"
	   showBo(xBaja)
	End If
	

End Sub






