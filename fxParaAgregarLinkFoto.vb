sub main
    stop
    set xVisualVar = VisualVarEditor("Escriba el link de la foto")
    call AddVarMemo(xVisualVar, "9_Extras", "Extras", "Extras","" )		
    xAceptar = ShowVisualVar(xVisualVar)

    self.boextension.LINKFOTO = GetValueVisualVar(xVisualVar, "9_Extras", "Extras")
end sub