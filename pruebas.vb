sub main

    stop
    self


    If MsgBox("¿Esta seguro de hacer inactivo al cliente?",36,"Pregunta") = 6 Then
        if self.workspace.intransaction Then self.workspace.commit
        call MsgBox("Se cambio el cc correctamente.",64,"Información")
    Else
        self.workspace.rollback
        Call MsgBox("Proceso cancelado",64,"Información")
        Exit Sub
    end if 


end sub