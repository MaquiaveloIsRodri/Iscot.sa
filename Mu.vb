dim strToday, businessDays3, businessDays5
strToday = date

select case WeekDay( strToday )
    case 1     '-- Domingo
        businessDays3 = 3
    case 2     '-- Lunes
        businessDays3 = 3
    case 3     '-- Martes
        businessDays3 = 3
    case 4     '-- Miercoles
        businessDays3 = 5
    case 5    '-- Jueves
        businessDays3 = 4
    case 6    '-- Viernes
        businessDays3 = 3
    case 7     '-- Sabado
        businessDays3 = 3
end select

strFive = DateAdd( "d", businessDays3, strToday )
strSeven = DateAdd( "d", businessDays5, strToday )


' 23/10/2018 - ME Orden Compra - Proveedor CC.
sub main
    Transaccion.FechaEntrega = CDate("01/01/2000")
	stop
	if (not Transaccion.Destinatario is nothing) and (not Transaccion.CentroCostos is nothing) then
		for each xPlazo in Transaccion.Destinatario.BOExtension.PlazoEntrega
			if xPlazo.CentroCostos.ID = Transaccion.CentroCostos.ID then
				aux_dia 	= Transaccion.FechaActual + xPlazo.Dias
				seguir 		= true
				do while seguir
					if Weekday(aux_dia) = 7 then		' Sábado.
						aux_dia = aux_dia + 2
					end if
					if Weekday(aux_dia) = 1 then		' Domingo.
						aux_dia = aux_dia + 1
					end if
					
					' Feriados.
					es_feriado	= true
					set xViewFe = NewCompoundView(Transaccion, "UD_FERIADOS", Transaccion.Workspace, nil, true)
					xViewFe.AddFilter(NewFilterSpec(xViewFe.ColumnFromPath("FECHA"), " = ", aux_dia))
					if xViewFe.ViewItems.Size = 0 then
						es_feriado = false
					else
						aux_dia = aux_dia + 1
					end if

					if Weekday(aux_dia) <> 7 and Weekday(aux_dia) <> 1 and es_feriado = false then
						seguir = false
					end if
				loop
				' Control de fines de semana | Los sumo por que se cuentan unicamente dias habiles
                For x = 0 to xPlazo.Dias
                    If Weekday(Transaccion.FechaActual + x) = 7 Or Weekday(Transaccion.FechaActual + x) = 1 Then
                        aux_dia = aux_dia + 1
                    End If
                Next
				Transaccion.FechaEntrega = aux_dia
				exit for
			end if
		next
	end if
end sub