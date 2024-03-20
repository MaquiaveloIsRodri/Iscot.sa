sub Main
    stop
    Set xVisualVar = VisualVarEditor("Coloque el costo")

    Call AddVarFloat(xVisualVar,  "1_costo" ,  "Costo Tlc", "Costo", 0)

    If Not ShowVisualVar(xVisualVar) Then Exit Sub

    Costo    =   GetValueVisualVar(xVisualVar, "1_costo", "Costo")
    for each xBaja in container

    next

end sub



