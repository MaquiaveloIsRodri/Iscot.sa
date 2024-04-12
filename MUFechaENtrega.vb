sub main
    Stop

    If Not NewValue Is Nothing Then
        if NewValue.value.name = "" = "despido indirecto" or NewValue.value.name = "renuncia" or NewValue.value.name = "incapacidad absoluta"    then
            Owner.Attributes("JuridiccionJuicio").ReadOnly = True
            owner.AnalistaART       = ""
        end if


    end



end sub

