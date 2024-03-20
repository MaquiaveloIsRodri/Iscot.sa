'Neuronal Intercambio Epistolar
Sub Main
    stop
    If aObject.Workspace.InState(3) = true Then
    ' Solo SI presiona en el boton ACEPTAR
        if aObject.contestacion = false then
            If aObject.motivo Is Nothing Then
                addfailedconstraint "Motivo no ingresado",1
                Exit Sub
            End If

            If  aObject.motivo.id = "{2E1F4B42-3133-44AB-B423-F1A965616444}" OR aObject.motivo.id = "{D07C37A4-3ED1-4E9B-BA61-5CB0EC991258}" Then
                ' View bajas del empleado
                Set xViewBajas = NewCompoundView(aObject, "TRSOLICITUD", aObject.Workspace, Nil, True)
                xViewBajas.AddJoin(NewJoinSpec(xViewBajas.ColumnFromPath("BOEXTENSION"), NewColumnSpec("UD_BAJAPERSONAL", "ID", ""), false))
                xViewBajas.AddFilter(NewFilterSpec(NewColumnSpec("TRSOLICITUD", "FLAG", ""), " <> ", "{ABEDCC7F-8CDF-4CF1-AC07-6FAA98521326}")) 'ANULADA
                xViewBajas.AddFilter(NewFilterSpec(NewColumnSpec("TRSOLICITUD", "DESTINATARIO", ""), " = ", aObject.empleado))
                For Each baja In xViewBajas.ViewItems
                    existeBaja = True
                    Set xBaja = baja.BO
                Next
                If Not existeBaja Then
                    If MsgBox("Genera automaticamente la baja ? " , 36, "Pregunta") <> 6 Then
                        Set xBaja = CrearTransaccion("BAJA", InstanciarBO( "{32DC2522-003F-4694-B193-5AE59009B6AF}", "UOSOLICITUD", aObject.Workspace ))
                        Set xBaja.destinatario = aObject.empleado
                        xBaja.boextension.fechaenvio = aObject.fechadeenvio
                        xBaja.boextension.fecharecepcion = aObject.fecharecepcion
                        xBaja.Detalle = "Generado por Intercambio"
                        xBaja.Nota = "Generado por Intercambio"
                        aObject.contestacion = true
                        Call showBo(xBaja)
                    end if
                End If
            End If
        End If
    End If
End Sub