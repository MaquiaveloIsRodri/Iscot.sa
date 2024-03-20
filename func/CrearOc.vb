Sub main
    Stop

    Set xUO             = InstanciarBO( "{CEA52A93-43D9-429A-9708-AD75A1183343}", "UOADMINISTRACIONCOMPRAS", self.Workspace )
	Set xFlag           = InstanciarBO( "{C416D932-A038-4861-9081-AEF8BD66FB0F}", "FLAG", self.Workspace )
	Set xFlagRevisar    = InstanciarBO( "{A32FF5FD-6FFD-4813-BBCE-DDA2F5BF3718}", "FLAG", self.Workspace )
    set oTipoCompra     = ExisteBo(Self,"TipoClasificador","ID","C1AE72F3-E744-4E99-8B7B-2F0A444053D2",nil,True,False,"=")

    If xUO Is Nothing Then
        Call MsgBox ("No se pudo instanciar UO. Contacte a Sistemas.",64,"Información")
        Exit Sub
    End If

    For each xitem in self.boextension.SERVICIOSCONTRATADOS
		If xitem.ordencompranueva is nothing then
            set xoc =  ClonarTransaccion( xitem.ordencompra )
            xOc.boextension.fechregis = self.fechaactual
            xOc.fechaactual = self.boextension.fechaactual
			xOc.nota = Self.BOExtension.Nota
			xOC.FechaEntrega = Self.FechaEntrega
			xOC.BOExtension.FechaEntregaReal = Self.BOExtension.FechaEntregaReal
            xitem.ordencompranueva = xOC
			xOC.BOExtension.SolicitantePIC = xitem.ResponsableOC


            if xitem.TIPOCOMPRA is nothing then
                set xTipoCompra = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "CODIGO", "D0788591-FE6D-4FCF-A949-F68E08FCCF55", oTipoCompra.Valores, True, False, "=")
                xoc.boextension.TIPOPIC = xTipoCompra
            else
                xoc.boextension.TIPOPIC = xitem.TIPOCOMPRA
            end if


            If xitem.revision Then
                xoc.flag = xFlagRevisar
            Else
                xoc.flag = xFlag
            End If

            xOc.boextension.observainterna = xItem.detalle
            If xItem.detalle <> "" Then xOc.nota = xItem.detalle End If
            xOc.centrocostos = xItem.centrocostos
            If xoc.itemstransaccion.count = 1 Then
                for each iOC in xOc.itemstransaccion
                    iOC.cantidad.cantidad = xitem.cantidad
                    iOC.valor.importe = xitem.importe
                next
            End if
		    xitem.ordencompranueva = xOC
		End If
    Next
	If Self.Workspace.InTransaction Then
		Self.Workspace.Commit
	End If
End Sub