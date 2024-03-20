'Fx Cambiar Centro de Costos
Sub Main
	stop

    'view de motivos
    Set xView = NewCompoundView(Self, "ITEMTIPOCLASIFICADOR", Self.Workspace, Nil, True)
    xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
    xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", "{3A29AA3D-DA9D-419D-B305-5FC3C50171C3}"))
    xView.NoFlushBuffers = True
    xView.AddBOCol("NOMBRE")
    xView.ColumnFromPath("NOMBRE")
    Set xContainerMotivos = NewContainer()
    xContainerMotivos.Add (xView)


    Set xCCActual = Self.centrocostos
    Set xVisualVar = VisualVarEditor("Cambio Centro de Costos")

    Call AddVarObj(xVisualVar,  "1_ccActual" ,  "Centro de costos Actual", "Cambio", Self.centrocostos ,getContainer("CENTROCOSTOS", Self.WorkSpace), Self.WorkSpace )
    Call AddVarDate(xVisualVar, "2_fechaHasta", "Fecha Hasta", "Cambio",now())
    Call AddVarObj(xVisualVar,  "3_ccNuevo",    "Centro de costos Nuevo", "Cambio", nothing,GetContainer("CENTROCOSTOS", Self.WorkSpace), Self.WorkSpace )
    Call AddVarObj(xVisualVar,  "4_motivo",     "Motivo", "Cambio", nothing ,xContainerMotivos, Self.WorkSpace )


    If Not ShowVisualVar(xVisualVar) Then Exit Sub

    fechaHasta      =   CDate(GetValueVisualVar(xVisualVar, "2_fechaHasta", "Cambio"))
    Set ccActual    =   GetValueVisualVar(xVisualVar, "1_ccActual", "Cambio")
    Set ccNuevo     =   GetValueVisualVar(xVisualVar, "3_ccNuevo", "Cambio")
    Set motivo      =   GetValueVisualVar(xVisualVar, "4_motivo", "Cambio")


    If ccNuevo Is Nothing Then
        MsgBox "No seleccionó Centro de Costos.", 48, "Aviso"
        exit sub
    End If
    If motivo Is Nothing Then
        MsgBox "No seleccionó Motivo de cambio", 48, "Aviso"
        exit sub
    End If


	set oEmpleado = getEmpleadoDeUsuario( nombreusuario(), self.workspace)
    set xHCc = crearbo("UD_HISTORICOCELULAR",self)
    xHCc.CELULAR    = self 'Asignamos el celular
    xHCc.originante = oEmpleado 'Asignamos el originate del cambio
    xHCc.CENTROCOSTOSANTERIOR = ccActual 'Asigno el cc Anterior
    xHCc.CENTROCOSTOSNUEVO    = ccNuevo 'Asignamos un nuevo cc
    xHCc.FECHAHASTA = fechaHasta 'Fecha hasta que tuvo activo
    xHCc.MOTIVO = motivo 'Asignamos el motivo

    self.CENTROCOSTOS = ccNuevo 'Le asignamos el nuevo CC al dispositivo
    self.HISTORICOCELULARES.add(XCC)

    xMensaje = "¿Esta seguro de cambiar el CC de este dispositivo?"

    If MsgBox(xMensaje,36,"Pregunta") = 6 Then
        if self.workspace.intransaction Then self.workspace.commit
        call MsgBox("Se cambio el cc correctamente.",64,"Información")
    Else
        self.workspace.rollback
        Call MsgBox("Proceso cancelado",64,"Información")
        Exit Sub
    End If

End Sub



' programado = true
' If xcc.fechahasta =< date() Then
'     self.flag = xFlag
'     self.centrocostos = ccNuevo
' 	self.sector = xcc.sectorhasta
'     self.boextension.horario = xcc.horariohasta
' 	self.Zona = ccNuevo.BoExtension.Zona
'     programado = false
'     xMensaje = "Confirma el cambio de centro de costos?"
' Else
'     xMensaje = "Confirma el cambio de centro de costos para la fecha " & xcc.fechahasta & "?"
' End If