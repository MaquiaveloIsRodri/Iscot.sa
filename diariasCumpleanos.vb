relative    = 0.0
ban         = true 
do while not xRs.EOF
    if cc <> xRs("CENTROCOSTOS").Value then
        if cc <> "" then
            ban = false
        end if
        cc          = xRs("CENTROCOSTOS").Value
	end if

	Select Case xRs("PERIODO").Value
		case year(date) & month(Date)
            if ban = false then
                relative = CDbl(xRs("TOTAL").Value)
                xRs.MoveNext
            end if
            MontoActual = MontoActual + CDbl(xRs("TOTAL").Value) + relative
			total       = total + MontoActual
		case year(date) & Mes2Da
            if ban = false then
                relative = CDbl(xRs("TOTAL").Value)
                xRs.MoveNext
            end if
            Monto2da    = Monto2da + CDbl(xRs("TOTAL").Value) + relative
			total       = total + Monto2da

		case year(date) & Mes3r
            if ban = false then
                relative = CDbl(xRs("TOTAL").Value)
                xRs.MoveNext
            end if
			Monto3ra = Monto3ra + CDbl(xRs("TOTAL").Value) + relative
			total    = total + Monto3ra
		Case Else
            if ban = false then
                relative = CDbl(xRs("TOTAL").Value)
                xRs.MoveNext
            end if
			resto = resto + CDbl(xRs("TOTAL").Value) + relative
			total = total + resto
	End Select

    relative = 0.0
    ban      = true
	HojaExcel.ActiveSheet.Cells(I, 3).Value 		= MontoActual
	HojaExcel.ActiveSheet.Cells(I, 4).Value 		= Monto2da
	HojaExcel.ActiveSheet.Cells(I, 5).Value 		= Monto3ra
	HojaExcel.ActiveSheet.Cells(I, 6).Value 		= resto
    MontoActual = 0.0
	Monto2da 	= 0.0
	Monto3ra 	= 0.0
	resto 		= 0.0
	I = I + 1
    xRs.MoveNext
loop
