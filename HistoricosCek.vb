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


        Case "{2FB05F95-3960-452C-82A3-3724CA394590}" ' Cambio de plan


    End Select




    Set xViewCC = NewCompoundView(Self, "CENTROCOSTOS", Self.Workspace, Nil, True)
    xViewCC.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
    xViewCC.AddBOCol("NOMBRE")
    xViewCC.ColumnFromPath("NOMBRE")
    Set xContainerCC = NewContainer()
    xContainerCC.Add(xViewCC)

    Set xCCActual = Self.centrocostos

    Call AddVarObj(xVisualVar,  "1_ccActual" ,  "Centro de costos Actual", "Cambio", Self.centrocostos ,xContainerCC , Self.WorkSpace )
    Call AddVarDate(xVisualVar, "2_fechaHasta", "Fecha Hasta", "Cambio",now())
    Call AddVarObj(xVisualVar,  "3_ccNuevo",    "Centro de costos Nuevo", "Cambio", nothing,xContainerCC, Self.WorkSpace )
   


    

    fechaHasta      =   CDate(GetValueVisualVar(xVisualVar, "2_fechaHasta", "Cambio"))
    Set ccActual    =   GetValueVisualVar(xVisualVar, "1_ccActual", "Cambio")
    Set ccNuevo     =   GetValueVisualVar(xVisualVar, "3_ccNuevo", "Cambio")
    

    set oEmpleado = getEmpleadoDeUsuario( nombreusuario(), self.workspace)
    set xHCc = crearbo("UD_HISTORICOCELULAR",self)
    xHCc.CELULAR    = self 'Asignamos el celular
    xHCc.originante = oEmpleado 'Asignamos el originate del cambio
    xHCc.CENTROCOSTOSANTERIOR = ccActual 'Asigno el cc Anterior
    xHCc.FECHAHASTA = fechaHasta 'Fecha hasta que tuvo activo
    
    'Validamos que motivo no sea nothing
   
   

    ' Validamos como si es nothing
    If ccNuevo Is Nothing Then
      xMensaje = "¿Estas seguro de quitar el CC? El mismo queda en stock"
	If MsgBox(xMensaje,36,"Pregunta") = 6 Then 
		xHCc.CENTROCOSTOSNUEVO    = nothing 
		self.stock = true
	else
		exit sub
	end if 
    else 
	xHCc.CENTROCOSTOSNUEVO    = ccNuevo 'Asignamos un nuevo cc
	self.stock = false
    End If


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

End Sub
