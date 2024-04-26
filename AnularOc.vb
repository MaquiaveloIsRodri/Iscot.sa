sub main
	Stop
	res = MsgBox("¿Esta seguro de Anular la OC: [ " & Self.NumeroDocumento & " ] ?", 292, "Anular OC")
	if res <> 6 then exit sub
	Anular = True
	For Each oTr In TransaccionesDestino(Self, Nothing)
		If oTr.Estado = "C" And Left(oTr.NumeroDocumento, 3) <> "ANU" then
		   Anular = False
		End if 
	Next
	
	If Anular Then
	   Set xView = NewCompoundView(Self, "pendiente", Self.Workspace, nill, True)
	   xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("Transaccion"), " = ", Self.ID))
	   For Each xItemV In xView.ViewItems
		   SALDARPENDIENTECOMPLETO (xItemV.bo)
	   Next
	   Self.external_id = "ANULADA"
	   Self.estado 	    = "N"
	   Self.Detalle	    = Self.detalle & " | " & " Anulada por: " & NombreUsuario()
	   Self.flag		= ExisteBO( self, "FLAG", "ID", "C10833F4-5E1E-47DA-803E-5FBF135BEA51", nil, TRUE, FALSE, "=" )
	   For Each oItem In Self.Itemstransaccion
		   oItem.EstadoTr = "N"
	   Next

	   ' AGREGADO: 27/12/2013 - Jose Fantasia.
	   call WorkSpaceCheck(Self.Workspace)
	   tieneDevenga = false
	   set xViewD = NewCompoundView(Self, "TRCONTABLE", Self.Workspace, nil, true)
	   xViewD.AddFilter(NewFilterSpec(xViewD.ColumnFromPath("VINCULOTR"), " = ", Self.ID))
	   xViewD.AddFilter(NewFilterSpec(xViewD.ColumnFromPath("NOTA"), " like ", "%DEVENGAMIENTO%"))
	   for each xItemDev in xViewD.ViewItems
	   	   set xAsientoDev = xItemDev.BO
		   tieneDevenga	   = true
		   exit for
	   next
	   if tieneDevenga then
	   	   call RevertirAsiento( xAsientoDev, false, xAsientoDev.FechaAplicacion )
	   end if
	Else
	   MsgBox "No se puede Anular una Orden de Compra facturada"
	End If
end sub


Public Function CheckDetailList(aAttributeLayout)
	' Esta función se utiliza para discriminar los atributos o sea:
	' Nothing se excluye.
	Dim xIndex
	On Error Resume Next
	
	Set CheckDetailList = Nothing
	Set CheckDetailList = NewLayout(aAttributeLayout.BO, aAttributeLayout.Attribute, aAttributeLayout.Caption, aAttributeLayout.ReadOnly, aAttributeLayout.Visible, "")
	
	Call CheckDetailList.AddListAttribute("NUMEROITEM", "Nro. Ítem")
	Call CheckDetailList.AddListAttribute("Referencia", "Cod. Producto")
	if nombreusuario() = "mleon" then
	   Call CheckDetailList.AddListAttribute("codigoalternativo", "Cod. Proveedor")
	end if
	Call CheckDetailList.AddListAttribute("Descripcion", "Descripción")
	Call CheckDetailList.AddListAttribute("[UD_ITEMORDENCOMPRA]BOExtension.ESPECIFICACIONES", "Especificaciones")
	Call CheckDetailList.AddListAttribute("Cantidad.Cantidad", "Cantidad")
	Call CheckDetailList.AddListAttribute("Valor.Importe", "Precio Compra")
	Call CheckDetailList.AddListAttribute("PorcentajeBonificacion", "% Bonificación")
	Call CheckDetailList.AddListAttribute("ImporteBonificado", "$ Bonificación")
	Call CheckDetailList.AddListAttribute("TotalSinDescuentos", "Total")
	'Call CheckDetailList.AddListAttribute("CentroCostos", "CC")
	Call CheckDetailList.AddListAttribute("[UD_ITEMORDENCOMPRA]BOExtension.MiniCC", "Registro Repuestos")
	'Call CheckDetailList.AddListAttribute("ListaPrecio", "Lista de Precio")
	Call CheckDetailList.AddListAttribute("[UD_ITEMORDENCOMPRA]BOExtension.CRITICO", "Crítico")
	Call CheckDetailList.AddListAttribute("[PRODUCTO]Referencia.CuentasContables.CuentaContable1.Descripcion", "Cuenta")
End Function
