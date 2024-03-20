'Fx Cambiar Centro de Costos
Sub Main   
	stop
    If GetParametroEjecucion( "LS_BLOQUEA_CAMBIO_CC", self.WorkSpace ) = "S" Then
        Call MsgBox("Los cambios de centro de costos no estan habilitados, Contactarse con el área de Liquidación de Sueldos",64,"Información")
        Exit Sub
    End If
    If self.flag Is Nothing Then
        ' TO-DO : VALIDAR QUE EL CC HHASTA SEA EL MISMO QUE EL CC DEL ANALAISTA REFERENTE
        Set xView = NewCompoundView(Self, "ITEMTIPOCLASIFICADOR", Self.Workspace, Nil, True)
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", "{0F09F416-F509-4106-B1A7-2505B1B2F651}"))
        xView.NoFlushBuffers = True		
        xView.AddBOCol("NOMBRE")
        xView.ColumnFromPath("NOMBRE")
        Set xContainerMotivos = NewContainer()
        xContainerMotivos.Add (xView)

        Set xView = NewCompoundView(Self, "ITEMTIPOCLASIFICADOR", Self.Workspace, Nil, True)
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", "{CDDB79EC-F609-483A-9951-035DD6A9AE9D}"))
        xView.NoFlushBuffers = True		
        xView.AddBOCol("NOMBRE")
        xView.ColumnFromPath("NOMBRE")
        Set xContainerMotivosRetraso = NewContainer()  
        xContainerMotivosRetraso.Add (xView)

        Set xView = NewCompoundView(Self, "ITEMTIPOCLASIFICADOR", Self.Workspace, Nil, True)
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", "0"))
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", "{F778D42A-7C15-4374-8F4A-F5B9DFF65530}"))
        xView.NoFlushBuffers = True		
        xView.AddBOCol("NOMBRE")
        xView.ColumnFromPath("NOMBRE")
        Set xContainerHorarios = NewContainer()  
        xContainerHorarios.Add (xView)

        Set xView = NewCompoundView(Self, "SECTOR", Self.Workspace, Nil, True)    
        xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("BO_PLACE"), " = ", "{2D172837-4208-4F14-8383-8DE88BA46F66}"))
        xView.NoFlushBuffers = True		
        xView.AddBOCol("NOMBRE")
        xView.ColumnFromPath("NOMBRE")
        Set xContainerSector = NewContainer()  
        xContainerSector.Add (xView)

        Set xCCActual = Self.centrocostos
        Set xVisualVar = VisualVarEditor("Cambio Centro de Costos")

        Call AddVarObj(xVisualVar, "1_ccActual", "Centro de costos Actual", "Cambio", Self.centrocostos ,getContainer("CENTROCOSTOS", Self.WorkSpace), Self.WorkSpace )
        Call AddVarDate(xVisualVar, "2_fechaHasta", "Fecha Hasta", "Cambio",now())
        Call AddVarObj(xVisualVar, "3_ccNuevo", "Centro de costos Nuevo", "Cambio", nothing,GetContainer("CENTROCOSTOS", Self.WorkSpace), Self.WorkSpace )
        Call AddVarObj(xVisualVar, "4_sector", "Sector Nuevo", "Cambio", nothing,xContainerSector, Self.WorkSpace )

        Call AddVarObj(xVisualVar, "5_motivo", "Motivo", "Cambio", nothing ,xContainerMotivos, Self.WorkSpace )
        Call AddVarObj(xVisualVar, "6_horario", "Horario", "Cambio", nothing ,xContainerHorarios, Self.WorkSpace )
        Call AddVarDate(xVisualVar, "7_fechaOperaciones", "Fecha Aviso Operaciones", "Cambio",now())

        call AddVarString(xVisualVar, "9_observacion", "Observacion", "Cambio", "")

        If Not ShowVisualVar(xVisualVar) Then Exit Sub
        
        fechaHasta =   CDate(GetValueVisualVar(xVisualVar, "2_fechaHasta", "Cambio"))
        Set ccNuevo =   GetValueVisualVar(xVisualVar, "3_ccNuevo", "Cambio")
        Set Sector =   GetValueVisualVar(xVisualVar, "4_sector", "Cambio")
        Set motivo =   GetValueVisualVar(xVisualVar, "5_motivo", "Cambio")
        'Set motivoDemora =   GetValueVisualVar(xVisualVar, "99_motivoDemora", "Cambio")

        fechaOperaciones =   GetValueVisualVar(xVisualVar, "7_fechaOperaciones", "Cambio")
        
        Set horario =   GetValueVisualVar(xVisualVar, "6_horario", "Cambio")
        observacion =   GetValueVisualVar(xVisualVar, "9_observacion", "Cambio")
        
        ' TODO - VALIDAR TODO ESTO
        If ccNuevo Is Nothing Then 
            MsgBox "No seleccionó Centro de Costos.", 48, "Aviso"
            exit sub
        End If        
        If Sector Is Nothing Then 
            MsgBox "No seleccionó Sector", 48, "Aviso"
            exit sub
        End If        
        If motivo Is Nothing Then 
            MsgBox "No seleccionó Motivo de cambio", 48, "Aviso"
            exit sub
        End If        
        If horario Is Nothing Then 
            MsgBox "No seleccionó Horario", 48, "Aviso"
            exit sub
        End If
		If ccNuevo.boextension.Zona Is Nothing Then	  
		   Call MsgBox("El centro de costos seleccionado no tiene una zona asignada. Contactese con LLSS",48,"Aviso")
        End If

       ' Correos Supervisores
        xCorreosSupervisores = ""

        

        'If Not self.centrocostos.boextension.analistarrhh.boextension.CORREOINSTITUCIONAL Is Nothing Then
            'xCorreosSupervisores = xCorreosSupervisores & self.centrocostos.boextension.analistarrhh.enteasociado.direcelectronicaprincipal.direccionelectronica & ";"
			xCorreosSupervisores = xCorreosSupervisores & self.centrocostos.boextension.analistarrhh.boextension.CORREOINSTITUCIONAL & ";"
        'End If
        'If Not ccNuevo.boextension.analistarrhh.boextension.CORREOINSTITUCIONAL Is Nothing Then
            'xCorreosSupervisores = xCorreosSupervisores & ccNuevo.boextension.analistarrhh.enteasociado.direcelectronicaprincipal.direccionelectronica & ";"
			xCorreosSupervisores = xCorreosSupervisores & ccNuevo.boextension.analistarrhh.boextension.CORREOINSTITUCIONAL & ";"
        'End If

        'If Not sector.boextension.supervisoroperativo.enteasociado.direcelectronicaprincipal Is Nothing Then
        '    xCorreosSupervisores = xCorreosSupervisores & sector.boextension.supervisoroperativo.enteasociado.direcelectronicaprincipal.direccionelectronica & ";"
        'End If



		set oEmpleado = getEmpleadoDeUsuario( nombreusuario(), self.workspace)
        set xFlag = InstanciarBO( "2D648132-3532-40D8-93A0-68198B398178", "FLAG", self.Workspace) ' Pendiente de Confirmación
        set xcc = crearbo("ud_centrocostoshistoricos",self)
        xcc.empleado = self
        xcc.originante = oEmpleado
        xcc.horariodesde = self.boextension.horario
        xcc.horariohasta = horario
        xcc.referentedesde = self.centrocostos.boextension.analistarrhh
        xcc.referentehasta = ccNuevo.boextension.analistarrhh
        xcc.flag = xFlag
        xcc.centrocostosdesde = Self.centrocostos

        xQueryUltimo = "SELECT TOP 1 HCC.ID HCC_ID, HCC.FECHADESDE " &_
        "FROM   V_UD_CENTROCOSTOSHISTORICOS_ HCC " &_
        "WHERE HCC.BO_PLACE_ID = '" & self.BOEXTENSION.CENTROCOSTOSHISTORICOS.ID &_
        "' ORDER BY HCC.FechaHasta DESC"

        xstring = StringConexion( "CALIPSO", self.WorkSpace ) 
        set xcone = createobject("adodb.connection")
        xcone.connectionstring  = xstring
        xcone.connectiontimeout = 0

        set xRs = RecordSet(xCone, "select top 1 * from producto")
        xRs.close
        xRs.activeconnection.commandtimeout=0
        xRs.source = xQueryUltimo
        xRs.open
		Stop
		Set xCCAnterior = Nothing
        do while not xRs.eof
            set xCCAnterior = InstanciarBO( CStr(xRs("HCC_ID").Value) , "UD_CENTROCOSTOSHISTORICOS", self.Workspace)
            xRs.MOVENEXT
        loop
        
	    xcc.fechahasta = fechaHasta        

        If xCCAnterior is Nothing Then  ' primer cambio de cc
            If fechaHasta > self.fechaingreso Then
		        xcc.fechaDesde = self.fechaingreso
            Else
                'MsgBox "La Fecha de cambio de CC no puede ser anterior a la fecha de ingreso.", 48, "Aviso"
		        'exit sub
            End If
        else
            If fechaHasta > xCCAnterior.fechaHasta Then
                xcc.fechaDesde = xCCAnterior.fechaHasta + 1
            Else
                'MsgBox "La Fecha de cambio de CC no puede ser anterior a la fecha del último cambio.", 48, "Aviso"
		        'exit sub
            End If
		End If
		
        xcc.centrocostos = ccNuevo
        
        xcc.sectordesde = self.sector
        xcc.sectorhasta = sector

        xcc.observacion = observacion
        xcc.fechacarga = now()
        xcc.motivo = motivo	
		
        self.boextension.centrocostoshistoricos.add(XCC)
        ' controlar si la fecha es a futuro
        programado = true
        If xcc.fechahasta =< date() Then
            self.flag = xFlag
            self.centrocostos = ccNuevo
			self.sector = xcc.sectorhasta
		    self.boextension.horario = xcc.horariohasta
			self.Zona = ccNuevo.BoExtension.Zona
            programado = false
            xMensaje = "Confirma el cambio de centro de costos?"
        Else
            xMensaje = "Confirma el cambio de centro de costos para la fecha " & xcc.fechahasta & "?"
        End If
        If MsgBox(xMensaje,36,"Pregunta") = 6 Then           
            Set xDict = NewDic()
            xstring = StringConexion( "CALIPSO", self.WorkSpace ) 
            set xcone = createobject("adodb.connection")
            xcone.connectionstring  = xstring
            xcone.connectiontimeout = 0
            set xRsMails = RecordSet(xCone, "select top 1 * from producto")
            xRsMails.close
            xRsMails.activeconnection.commandtimeout = 0
            xRsMails.source = "SELECT NOMBRE FROM ITEMTIPOCLASIFICADOR WHERE BO_PLACE_ID = 'B21432D4-2631-46F1-93DE-3B5358B3AEE4' AND ACTIVESTATUS = 0 "
            xRsMails.open
            xcorreos  = ""

            Do while not xRsMails.eof           	
                xcorreos = xcorreos & xRsMails("NOMBRE").Value & ";"        
                xRsMails.MOVENEXT
            Loop
            xRsMails.close
            set xFSO 		= CreateObject("Scripting.FileSystemObject")
            set xArchivo = xFSO.OpenTextFile("C:\util\html\email_aviso_cambio_cc.html")    
            htmlBody = xArchivo.readAll()
            htmlBody = Replace(htmlBody,"xEmpleado",xcc.empleado.descripcion)
            htmlBody = Replace(htmlBody,"xMotivo",xcc.motivo.name)
            htmlBody = Replace(htmlBody,"xFechaInicio",xcc.fechahasta)
            htmlBody = Replace(htmlBody,"xCentroCostosDesde",xcc.centrocostosdesde.name)
            htmlBody = Replace(htmlBody,"xReferenteDesde",xcc.referentedesde.name)
            
            If xcc.sectordesde Is Nothing Then
                htmlBody = Replace(htmlBody,"xSectorDesde","-")
            Else
                htmlBody = Replace(htmlBody,"xSectorDesde",xcc.sectordesde.name)
            End If
            If xcc.horariodesde Is Nothing Then
                htmlBody = Replace(htmlBody,"xHorarioDesde","-")
            Else
                htmlBody = Replace(htmlBody,"xHorarioDesde",xcc.horariodesde.name)
            End If

            'htmlBody = Replace(htmlBody,"xFechaInicio",xcc.fechadesde)
            htmlBody = Replace(htmlBody,"xFechaHasta",xcc.fechadesde)
		    htmlBody = Replace(htmlBody,"xFecha",Date())
            htmlBody = Replace(htmlBody,"xCentroCostosHasta",xcc.centrocostos.name)
            htmlBody = Replace(htmlBody,"xReferenteHasta",xcc.referentehasta.name)
            htmlBody = Replace(htmlBody,"xSectorHasta",xcc.sectorhasta.name)
            htmlBody = Replace(htmlBody,"xHorarioHasta",xcc.horariohasta.name)
            htmlBody = Replace(htmlBody,"xObservacion",xcc.observacion)
            htmlBody = Replace(htmlBody,"xAvisoOperaciones",fechaOperaciones)
			htmlBody = Replace(htmlBody,"xLegajo",self.codigo)


            xSubject	= "[Cambio de CC] - " & Self.name
            xBody		= ""
            xcorreos	= xcorreos & ";" & xCorreosSupervisores
			stop
            xadjunto 	= ""
            stop
            call enviar_aviso(Self, xcorreos, xSubject, xBody, xadjunto,htmlBody)
            
            if self.workspace.intransaction Then self.workspace.commit   
            if programado Then 
                xMensaje = "Cambio programado correctamente. Se ejecutará de manera automática el "  & xcc.fechahasta 
            Else
                xUpdate = "update HISTORICOPERSONAL_" & Year(date) & "  set " &_ 
				"CENTROCOSTOS_ID = '"& ccNuevo.ID &"' , " &_
				"CENTROCOSTOS = '"& ccNuevo.NOMBRE &"', " &_
				"SECTOR_ID = '"& xcc.sectorhasta.id &"', " &_
				"LUGAR_TRAB = '"& xcc.sectorhasta.name &"', " &_ 
				"ZONA_ID = '"& ccNuevo.BoExtension.Zona.Id &"', " &_
				"ZONA = '"& ccNuevo.BoExtension.Zona.Nombre &"' " &_ 
				"where LEGAJO = '" & self.codigo & "' AND MOMENTO > '" & xcc.fechahasta & "'"
                'set xCon = CreateObject("adodb.connection")
                'xCon.ConnectionString	= StringConexion("calipso", self.Workspace)
                'xCon.ConnectionTimeout	= 150
                'xCon.Open
                'xCon.Execute()
                'xCon.Close
                xMensaje = "Cambio realizado correctamente"
				set xVector = NewVector()
		   		set xBucket = NewBucket()
		   		xBucket.Value = xUpdate
		   		xVector.Add(xBucket)
		   		ExecutarSQL xVector, "DistrObj", "", self.WorkSpace, -1
		   		Call workspacecheck(self.workspace)
            End If
            call MsgBox(xMensaje,64,"Información")
        Else
            self.workspace.rollback
            Call MsgBox("Proceso cancelado",64,"Información")
            Exit Sub
        End If
    Else
        Set xVisualVar = VisualVarEditor("Cambio Centro de Costos")

        Call AddVarBoolean(xVisualVar, "ingreso", "Tiene ingreso habilitado?", "Controles", false)
        Call AddVarBoolean(xVisualVar, "indumentaria", "Realizó control de Indumentaria?", "Controles", false)
        Call AddVarBoolean(xVisualVar, "firma", "Firmó notificación de cambio?", "Controles", false)
        Call AddVarBoolean(xVisualVar, "tuRecibo", "Realizó control en Tu Recibo?", "Controles", false)
        Call AddVarBoolean(xVisualVar, "reiwin", "Realizó cambio en Reiwin?", "Controles", false)
        If Not ShowVisualVar(xVisualVar) Then Exit Sub
        ' buscar ultimo cambio de cc?
        'Stop
        'xcc.controlingreso = GetValueVisualVar(xVisualVar, "ingreso", "Controles")
        'xcc.controlindumentaria = GetValueVisualVar(xVisualVar, "indumentaria", "Controles")
        'xcc.controlfirma = GetValueVisualVar(xVisualVar, "firma", "Controles")
        'xcc.controlreiwin = GetValueVisualVar(xVisualVar, "reiwin", "Controles")
        'xcc.controlturecibo = GetValueVisualVar(xVisualVar, "tuRecibo", "Controles")

        self.flag = nothing
        if self.workspace.intransaction Then self.workspace.commit
    End If
    
End Sub
