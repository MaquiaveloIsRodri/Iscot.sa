' 17/11/2023 - MU PIP Compra > Agregar Depósito y ubicación Masivamente 
sub main
    stop

    'Valicacion si esta activa
    if Self.ESTADO <> "A" then
        MsgBox "El PIC debe estar Abierto.", 64, "información"
        exit sub
    end if
    if Self.ItemsTransaccion.Size <= 0 then
        MsgBox "No cargó ítems.", 64, "información"
        exit sub
    end if

    ' view Depositos Activos.
    set xViewDep = NewCompoundView(Self, "DEPOSITO", Self.Workspace, nil, true)
    xViewDep.AddFilter(NewFilterSpec(NewColumnSpec("DEPOSITO", "ACTIVESTATUS", ""), " = ", "0"))
    xViewDep.AddBOCol("NOMBRE")
    xViewDep.AddOrderColumn(NewOrderSpec(xViewDep.ColumnFromPath("NOMBRE"), false))
    'Contenedor deposito
    Set xContainerDep = NewContainer()  
    xContainerDep.Add(xViewDep)

    'View ubicacion activa
    set xViewUbi = NewCompoundView(Self, "UBICACION", Self.Workspace, nil, true)
    xViewUbi.AddFilter(NewFilterSpec(NewColumnSpec("UBICACION", "ACTIVESTATUS", ""), " = ", "0"))
    xViewUbi.AddBOCol("NOMBRE")
    xViewUbi.AddOrderColumn(NewOrderSpec(xViewDep.ColumnFromPath("NOMBRE"), false))
    'Contenedor Ubicacion
    Set xContainerUbi = NewContainer()
    xContainerUbi.Add(xViewUbi)


    ' VisualVar Deposito.
    set xVisualVar = VisualVarEditor("PIP Compra - Agregar Depósito y ubicación Masivamente ")
    Call AddVarObj( xVisualVar, "01DEPOSITO" ,"Deposito", "Indique", nothing , xContainerDep, Self.WorkSpace )
    Call AddVarObj( xVisualVar, "01UBICACION","Ubicacion","Indique", Nothing , xContainerUbi, Self.WorkSpace )
    aceptar = ShowVisualVar(xVisualVar)
    if not aceptar then exit sub

    'Getting data Deposito
    set oDeposito   = Nothing
    set oDeposito   = GetValueVisualVar(xVisualVar, "01DEPOSITO", "Indique")

    'Getting data Ubi
    set oUbicacion  = Nothing
    set oUbicacion  = GetValueVisualVar(xVisualVar, "01UBICACION", "Indique")

    if oDeposito Is Nothing then
        MsgBox "No indicó el Depósito.", 48, "Aviso"
        exit sub
    end if
    if oUbicacion Is Nothing then
        MsgBox "No indicó el Ubicacion.", 48, "Aviso"
        exit sub
    end if

    for each xItem in Self.ItemsTransaccion
        'asignar deposito y ubicacion
        xItem.DEPOSITOORI   = oDeposito
        xItem.UBICACIONORI  = oUbicacion
    next

end sub
