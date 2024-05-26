'Fx Cambiar Centro de Costos
Sub Main
	stop

    'view de motivos
    Set xView = NewCompoundView(Self, "ITEMTIPOCLASIFICADOR", Self.Workspace, Nil, True)
    xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
    xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", "{BCF5FCC6-4114-4935-AF85-AC9BA5E15973}"))
    xView.NoFlushBuffers = True
    xView.AddBOCol("NOMBRE")
    xView.ColumnFromPath("NOMBRE")
    Set xContainerMotivos = NewContainer()
    xContainerMotivos.Add (xView)

    ' EL primer visualVar - Se hace para los motivos
    Set xVisualVar = VisualVarEditor("Seleccione el motivo")
    Call AddVarObj(xVisualVar,  "1_motivo",     "Motivo", "Cambio", nothing ,xContainerMotivos, Self.WorkSpace )

    If Not ShowVisualVar(xVisualVar) Then Exit Sub
    Set motivo      =   GetValueVisualVar(xVisualVar, "1_motivo", "Cambio")

    If motivo Is Nothing Then
        MsgBox "No seleccionó Motivo de cambio", 48, "Aviso"
        exit sub
    End If



	Select Case NewValue.value.id
        Case "{85F3F2AD-31D9-419C-AE80-1132995C9A79}" ' Cambio de centro de costos
            Set xViewCC = NewCompoundView(Self, "CENTROCOSTOS", Self.Workspace, Nil, True)
            xViewCC.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
            xViewCC.AddBOCol("NOMBRE")
            xViewCC.ColumnFromPath("NOMBRE")
            Set xContainerCC = NewContainer()
            xContainerCC.Add(xViewCC)

            Set xCCActual = Self.centrocostos

            Set xVisualVarCc = VisualVarEditor("Cambio de centro de costo")
            Call AddVarObj(xVisualVarCc,  "1_ccActual" ,  "Centro de costos Actual", "Cambio", Self.centrocostos ,xContainerCC , Self.WorkSpace )
            Call AddVarDate(xVisualVarCc, "2_fechaHasta", "Fecha Hasta", "Cambio",now())
            Call AddVarObj(xVisualVarCc,  "3_ccNuevo",    "Centro de costos Nuevo", "Cambio", nothing,xContainerCC, Self.WorkSpace )

            If Not ShowVisualVar(xVisualVarCc) Then Exit Sub

            fechaHasta      =   CDate(GetValueVisualVar(xVisualVar, "2_fechaHasta", "Cambio"))
            Set ccActual    =   GetValueVisualVar(xVisualVar, "1_ccActual", "Cambio")
            Set ccNuevo     =   GetValueVisualVar(xVisualVar, "3_ccNuevo", "Cambio")


            set oEmpleado = getEmpleadoDeUsuario( nombreusuario(), self.workspace)
            set xHCc = crearbo("UD_HISTORICOCELULAR",self)
            xHCc.CELULAR    = self 'Asignamos el celular
            xHCc.originante = oEmpleado 'Asignamos el originate del cambio
            xHCc.CENTROCOSTOSANTERIOR = ccActual 'Asigno el cc Anterior
            xHCc.FECHAHASTA = fechaHasta 'Fecha hasta que tuvo activo

            ' Validamos como si es nothing
            If ccNuevo Is Nothing Then

            xMensaje = "¿Estas seguro de quitar el CC? El mismo queda en stock"

            If MsgBox(xMensaje,36,"Pregunta") = 6 Then
                xHCc.CENTROCOSTOSNUEVO    = nothing
                self.stock = true
                exit sub
            End If

            xHCc.CENTROCOSTOSNUEVO    = ccNuevo 'Asignamos un nuevo cc
            self.stock = false
            self.CENTROCOSTOS = ccNuevo 'Le asignamos el nuevo CC al dispositivo

            xMensaje = "¿Esta seguro de cambiar el CC de este dispositivo?"

            If MsgBox(xMensaje,36,"Pregunta") = 6 Then
                if self.workspace.intransaction Then self.workspace.commit
                self.HISTORICOCELULARES.add(xHCc)
                call MsgBox("Se cambio el cc correctamente.",64,"Información")
            Else
                self.workspace.rollback
                Call MsgBox("Proceso cancelado",64,"Información")
                Exit Sub
            End If


        Case "{2FB05F95-3960-452C-82A3-3724CA394590}" ' Cambio de plan
           Set xViewCC = NewCompoundView(Self, "SERVICIOS", Self.Workspace, Nil, True)
            xViewCC.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
            xViewCC.AddBOCol("DENOMINACION")
            xViewCC.ColumnFromPath("DENOMINACION")
            Set xContainerServices = NewContainer()
            xContainerServices.Add(xViewCC)

            Set xServicioActual = Self.Servicios

            Set xVisualVarService = VisualVarEditor("Cambio de servicio")
            Call AddVarObj(xVisualVarService,  "1_ServicioActual" ,  "Servicio Actual", "Cambio", Self.Servicios ,xContainerServices , Self.WorkSpace )
            Call AddVarDate(xVisualVarService, "2_fechaHasta", "Fecha Hasta", "Cambio",now())
            Call AddVarObj(xVisualVarService,  "3_ServicioNuevo",    "Servicio Nuevo", "Cambio", nothing,xContainerCC, Self.WorkSpace )

            If Not ShowVisualVar(xVisualVarCc) Then Exit Sub

            fechaHasta              =   CDate(GetValueVisualVar(xVisualVar, "2_fechaHasta", "Cambio"))
            Set xServicioActual     =   GetValueVisualVar(xVisualVar, "1_ServicioActual", "Cambio")
            Set xServicioNuevo      =   GetValueVisualVar(xVisualVar, "3_ServicioNuevo", "Cambio")


            set oEmpleado = getEmpleadoDeUsuario( nombreusuario(), self.workspace)
            set xHCc = crearbo("UD_HISTORICOCELULAR",self)
            xHCc.CELULAR    = self 'Asignamos el celular
            xHCc.originante = oEmpleado 'Asignamos el originate del cambio
            xHCc.SERVICIOANTERIOR = xServicioActual 'Asigno el SERVICIO Anterior
            xHCc.FECHAHASTA = fechaHasta 'Fecha hasta que tuvo activo

            ' Validamos como si es nothing
            If xServicioNuevo Is Nothing Then

                xMensaje = "¿Estas seguro de quitar el servicio asignado?"

                If MsgBox(xMensaje,36,"Pregunta") = 6 Then
                    xHCc.SERVICIONUEVO    = nothing
                    exit sub
                End If
            end if

            xMensaje = "¿Esta seguro de cambiar el CC de este dispositivo?"

            If MsgBox(xMensaje,36,"Pregunta") = 6 Then
                if self.workspace.intransaction Then self.workspace.commit
                self.HISTORICOCELULAR.add(xHCc)
                call MsgBox("Se cambio el cc correctamente.",64,"Información")
            Else
                self.workspace.rollback
                Call MsgBox("Proceso cancelado",64,"Información")
                Exit Sub
            End If
        Case Else ' Para cualquier otro caso
            ' Código para manejar cualquier otro caso aquí
            call MsgBox("Contactarse con sistemas, para agregar el motivo.",64,"Información")
    End Select


End Sub
