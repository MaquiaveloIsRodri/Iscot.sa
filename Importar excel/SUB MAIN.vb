Sub Main
    ' primero recorrer el XSL
    While hojaExcel
        ' limpiamos las variables
        legajo = ""
        reporta = ""
        gerencia = ""
        ' tomar las variables del xls
        legajo = hojaExcel("codigo")
        reporta = hojaExcel("reporta_legajo")
        gerencia = hojaExcel("gerencia")
        ' revisar que todos los campos esten completos
        If legajo <> "" and reporta <> "" and gerencia <> "" Then
            ' limpiamos las variables
            Set xEmpleado = Nothing
            Set xJefe = Nothing
            Set xGerencia = Nothing
            ' instanciamos todo
            Set xEmpleado = InstanciarBO ("EMPLEADO","CODIGO",legajo)
            Set xJefe = InstanciarBO ("EMPLEADO","CODIGO",reporta)
            Set xGerencia = InstanciarBO ("TIPOCLASIFICADOR","CODIGO",gerencia)
            If Not xEmpleado is nothing and Not xJefe is nothing and Not xGerencia is nothing Then
                ' pudimos instanciar todo bien, actualizamos el empleado
                xEmpleado.boextension.reporta = xJefe
                xEmpleado.boextension.gerencia = gerencia
                ' probablmente agregar un commit
                self.workspace.commit
                 'Escribir en una celda del Excel "Empleado actualizado correctamente"
            Else
                If xEmpleado Is Nothing Then
                     HojaExcel.Worksheets("Base de datos").Cells(i, 7).Value = "No se encontró empleado."
                End If
                If xJefe Is Nothing Then
                    'Escribir en una celda del Excel "El legajo no corresponde a un empleado"
                End If
                If xGerencia Is Nothing Then
                    'Escribir en una celda del Excel "El codigo no corresponde a una gerencia"
                End If
            End If
        Else
            If legajo = "" Then
                'Escribir en una celda del Excel "El legajo no corresponde a un empleado"
            End If
            If  reporta = "" Then
                'Escribir en una celda del Excel "El legajo no corresponde a un empleado"
            End If
            If xGerencia Is Nothing Then
                'Escribir en una celda del Excel "El codigo no corresponde a una gerencia"
            End If
        End If
    End While

    ' Guardar XLS

     HojaExcel.quit

    msgbox "Proceso Finalizado con EXITO, revisar columna 7 las obervaciones"

End Sub
