' CREADO: 23/01/2015 - Daiana Estudia
' 27/09/2019 - Actualizado para que FIAT también se vea Mensual.
' MU ExportarPICxPERIODO.
sub main
	stop

	' 01/07/2016 - Agregado: tilde para ver si el informe se hace desde mes actual o desde el que sigue.

	set xVisualVar = VisualVarEditor("EXPOTAR PICS")
	call AddVarBoolean(xVisualVar, "00DESDE", "Desde Mes Siguiente?", "Parametros:", false)
	aceptar = ShowVisualVar(xVisualVar)
	if not aceptar then exit sub

	desdeSig = GetValueVisualVar(xVisualVar, "00DESDE", "Parametros:")
	if desdeSig then
		fechaActual = DateAdd("m", 2, CDate("01/" & Right("00" & Month(Date), 2) & "/" & Year(Date)))
		fechaActual = DateAdd("d", -1, fechaActual)
	else
		fechaActual = Date
	end if

	ENTRA = FALSE
	if self.id = "{641D686F-BBE6-4B5F-A447-722D2B6B8D13}" and 1 = 2 THEN' FIAT AUTO ARGENTINA '2 IF
		ENTRA = TRUE
		'fechaActual = DATE

		if DAY(fechaActual) < 16 then '1 IF

			fechaActualIncial = YEAR(fechaActual) & Right("00" & Month(fechaActual), 2) & "01"
			fechaActualFinal  = YEAR(fechaActual) & Right("00" & MONTH(fechaActual), 2) & "15"
			nombreMesActual = "1º Quincena de " & nombremes(fechaActual)

			unMesAnterior = dateadd ("m", -1, fechaActual)
			diferencia1 = DATEDIFF("d",unMesAnterior, fechaActual)' ES PARA SACAR SI EL MES TIENE 31 O 30

			nombrePrimeraQuincena  = "2º Quincena de " & NombreMes(unMesAnterior)
			diaInicioMesAnterior   = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "16" ' INICIO SEGUNDA QUINCENA
			diaFinMesAnterior      = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & Right("00" & diferencia1,2) 'FIN SEGUNDA QUINCENA
			nombreSegundaQuincena  = "1º Quincena de " & NombreMes(unMesAnterior)
			diaInicioDosMesAnterior= YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "01" 'INICIO PRIMERA QUINCENA'
			diaFinDosMesAnterior   = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "15" 'FINAL PRIMERA QUINCENA'

			dosMesAnterior = dateadd ("m", -1, unMesAnterior)
			diferencia2 = DATEDIFF("d",dosMesAnterior, unMesAnterior)' ES PARA SACAR SI EL MES TIENE 31 O 30

			nombreTerceraQuincena  = "2º Quincena de " & NombreMes(dosMesAnterior)
			diaInicioTresMesAnterior   = YEAR(dosMesAnterior) & Right("00" & MONTH(dosMesAnterior), 2) & "16" ' INICIO SEGUNDA QUINCENA
			diaFinTresMesAnterior      = YEAR(dosMesAnterior) & Right("00" & MONTH(dosMesAnterior), 2) & Right("00" & diferencia2,2) 'FIN SEGUNDA QUINCENA

		else ' ELSE 1 IF
			fechaActualIncial = YEAR(fechaActual) & Right("00" & Month(fechaActual), 2) & "16"
			fechaActualFinal = YEAR(fechaActual) & Right("00" & MONTH(fechaActual), 2) & Right("00" & DAY(fechaActual), 2) 
			nombreMesActual = "2º Quincena de " & nombremes(fechaActual)

			nombrePrimeraQuincena  = "1º Quincena de " & NombreMes(fechaActual)
			diaInicioMesAnterior   = YEAR(fechaActual) & Right("00" & MONTH(fechaActual), 2) & "01" 'INICIO PRIMERA QUINCENA
			diaFinMesAnterior      = YEAR(fechaActual) & Right("00" & MONTH(fechaActual), 2) & "15" 'FINAL PRIMERA QUINCENA

			unMesAnterior = dateadd ("m", -1, fechaActual)
			diferencia1 = DATEDIFF("d",unMesAnterior, fechaActual)' ES PARA SACAR SI EL MES TIENE 31 O 30

			nombreSegundaQuincena  = "2º Quincena de " & NombreMes(unMesAnterior)
			diaInicioDosMesAnterior= YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "16" ' INICIO SEGUNDA QUINCENA
			diaFinDosMesAnterior   = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & Right("00" & diferencia1,2) 'FIN SEGUNDA QUINCENA

			dosMesAnterior = dateadd ("m", -1, unMesAnterior)

			nombreTerceraQuincena  = "1º Quincena de " & NombreMes(unMesAnterior)
			diaInicioTresMesAnterior   = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "01" ' INICIO PRIMERA QUINCENA
			diaFinTresMesAnterior      = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "15" 'FIN PRIMERA QUINCENA
		end if ' FIN 1 IF

		' 24/06/2016 José Luis Fantasia.
		' Se modificaron: query, query1 y query2.
		' Se agregaron los INNER JOIN de UD_PEDIDOINTERNOCOMPRA.
		' los LEFT(UD.FECHAREGISTRACION, 8)
		' y en el WHERE se agregó: TR.ESTADO <> 'N' para mes actual e = 'C' en los anteriores.
        '27/11/2023 Se agrego una validacion al TIPOPIC = INSUMOS
		' QUERY PARA QUINCENAS FIAT
		query = "select Q.REFERENCIA_ID, Q.PRODUCTO AS PRODUCTO, CAST(SUM(Q.CANTIDAD_MES)AS FLOAT) AS MESACTUAL,"_ 
		  &" CAST(SUM(Q.CANTIDAD_MES1)AS FLOAT) AS UNMESANTERIOR,"_ 
	   	  &" CAST(SUM(Q.CANTIDAD_MES2)AS FLOAT) AS DOSMESANTERIOR,"_ 
		  &" CAST(SUM(Q.CANTIDAD_MES3)AS FLOAT) AS TRESMESANTERIOR,"_ 
	   	  &" ISNULL((SELECT CAST(PR.VALOR2_IMPORTE AS FLOAT) FROM PRECIO PR WITH(NOLOCK)"_ 
 	   	  &" WHERE pr.HABILITADO = 'T'"_ 
 	   	  &" AND PR.REFERENCIA_ID    = Q.REFERENCIA_ID"_
 	   	  &" AND PR.INICIAL2_MOMENTO = (SELECT MAX(PRE.INICIAL2_MOMENTO) FROM PRECIO PRE WHERE PRE.REFERENCIA_ID = PR.REFERENCIA_ID)),0) AS PRECIO,"_
		  &" ROUND(CAST((SUM(Q.CANTIDAD_MES)+SUM(Q.CANTIDAD_MES1)+SUM(Q.CANTIDAD_MES2)+SUM(Q.CANTIDAD_MES3))/4 AS FLOAT),0) as PROMEDIO"_
		  &" from ("_ 
		  &" SELECT IOC.REFERENCIA_ID,IOC.DESCRIPCION as PRODUCTO, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES,0 AS CANTIDAD_MES1, 0 as CANTIDAD_MES2, 0 as CANTIDAD_MES3"_ 
		  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
		  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
		  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
		  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
		  &" WHERE TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
		  &" AND CC.id = '"&SELF.ID&"'"_ 
		  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&fechaActualIncial&"' and '"&fechaActualFinal&"'"_ 
		  &" and TR.ESTADO <> 'N'"_
		  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
		  &" union all"_ 
		  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO,0 AS CANTIDAD_MES, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES1, 0 AS CANTIDAD_MES2, 0 as CANTIDAD_MES3"_ 
		  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
		  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
		  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
		  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
		  &" WHERE  ud.TIPOPIC_N = 'I - Insumos' AND "_
          &" TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
		  &" AND CC.id = '"&SELF.ID&"'"_ 
		  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioMesAnterior&"' and '"&diaFinMesAnterior&"'"_ 
		  &" and TR.ESTADO = 'C'"_
		  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
		  &" union all "_
		  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO, 0 AS CANTIDAD_MES, 0 AS CANTIDAD_MES1, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES2, 0 as CANTIDAD_MES3"_ 
		  &" FROM TRORDENCOMPRA AS TR with(nolock) "_
		  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
		  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
		  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
		  &" WHERE  ud.TIPOPIC_N = 'I - Insumos'"_
          &" TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_
		  &" AND CC.id = '"&SELF.ID&"'"_ 
		  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioDosMesAnterior&"' and '"&diaFinDosMesAnterior&"'"_  
		  &" and TR.ESTADO = 'C'"_
		  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_
		  &" union all"_ 
		  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO, 0 AS CANTIDAD_MES, 0 AS CANTIDAD_MES1, 0 AS CANTIDAD_MES2, SUM(IOC.CANTIDAD2_CANTIDAD) as CANTIDAD_MES3"_
		  &" FROM TRORDENCOMPRA AS TR with(nolock) "_
		  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID "_
		  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID "_
		  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
		  &" WHERE  ud.TIPOPIC_N = 'I - Insumos'"_
          & "TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_
		  &" AND CC.id = '"&SELF.ID&"'"_
		  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioTresMesAnterior&"' and '"&diaFinTresMesAnterior&"'"_ 
		  &" and TR.ESTADO = 'C'"_
		  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION) AS Q"_ 
		  &" group by Q.REFERENCIA_ID,Q.PRODUCTO order by 2"_
		
	else ' ELSE 2 IF
		'FECHAS PARA CUALQUIER CENTRO DE COSTO EXCEPTO FIAT
		' 27/09/2019 - Ahora Fiat Si.
		'fechaActual = DATE

		nombreMesActual = NombreMes(fechaActual)
		fechaActualIncial = YEAR(fechaActual) & Right("00" & Month(fechaActual), 2) & "01"
		fechaActualFinal = YEAR(fechaActual) & Right("00" & MONTH(fechaActual), 2) & Right("00" & DAY(fechaActual), 2)

		unMesAnterior = dateadd ("m", -1, fechaActual)
		nombreUnMesAnterior = NombreMes(unMesAnterior)

		dosMesAnterior = dateadd ("m", -2, fechaActual)
		nombreDosMesAnterior = NombreMes(dosMesAnterior)

		' calculo rango de dias de mes anterior
		diaInicioMesAnterior = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & "01"
		diferencia1 = DATEDIFF("d",unMesAnterior, fechaActual)
		diaFinMesAnterior = YEAR(unMesAnterior) & Right("00" & MONTH(unMesAnterior), 2) & Right("00" & diferencia1,2)

		'calculo rango mes -2
		diaInicioDosMesAnterior = YEAR(dosMesAnterior) & Right("00" & MONTH(dosMesAnterior), 2) & "01"
		diferencia2 = DATEDIFF("d",dosMesAnterior, unMesAnterior)
		diaFinDosMesAnterior = YEAR(dosMesAnterior) & Right("00" & MONTH(dosMesAnterior), 2) & Right("00" & diferencia2,2)

		IF self.id = "{E301E71F-829F-4DA3-ADA3-977F0E6F46BF}" THEN ' FORD ' INICIO IF 3
			'query HOJA 1 - FORD PINTURA
			query1 ="select Q.REFERENCIA_ID, Q.PRODUCTO AS PRODUCTO, CAST(SUM(Q.CANTIDAD_MES)AS FLOAT) AS MESACTUAL,"_ 
			  &" CAST(SUM(Q.CANTIDAD_MES1)AS FLOAT) AS UNMESANTERIOR,"_ 
			  &" CAST(SUM(Q.CANTIDAD_MES2)AS FLOAT) AS DOSMESANTERIOR,"_ 
			  &" ISNULL((SELECT CAST(PR.VALOR2_IMPORTE AS FLOAT) FROM PRECIO PR WITH(NOLOCK)"_ 
			  &" WHERE pr.HABILITADO = 'T'"_ 
			  &" AND PR.REFERENCIA_ID    = Q.REFERENCIA_ID"_
			  &" AND PR.INICIAL2_MOMENTO = (SELECT MAX(PRE.INICIAL2_MOMENTO) FROM PRECIO PRE WHERE PRE.REFERENCIA_ID = PR.REFERENCIA_ID)),0) AS PRECIO,"_
			  &" ROUND(CAST((SUM(Q.CANTIDAD_MES)+SUM(Q.CANTIDAD_MES1)+SUM(Q.CANTIDAD_MES2))/3 AS FLOAT),0) as PROMEDIO"_ 
			  &" from ("_ 
			  &" SELECT IOC.REFERENCIA_ID,IOC.DESCRIPCION as PRODUCTO, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES,0 AS CANTIDAD_MES1, 0 as CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD WITH(NOLOCK) ON UD.ID = TR.BOEXTENSION_ID"_
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"' /*AND UD.CLASIFICADOR_CC_FORD_ID = '7BE98AF1-88CD-42F8-B621-78E71D9BFFBF'*/"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&fechaActualIncial&"' and '"&fechaActualFinal&"'"_ 
			  &" and TR.ESTADO <> 'N'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
			  &" union all"_ 
			  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO,0 AS CANTIDAD_MES, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES1, 0 AS CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD WITH(NOLOCK) ON UD.ID = TR.BOEXTENSION_ID"_
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"' /*AND UD.CLASIFICADOR_CC_FORD_ID = '7BE98AF1-88CD-42F8-B621-78E71D9BFFBF'*/"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioMesAnterior&"' and '"&diaFinMesAnterior&"'"_ 
			  &" and TR.ESTADO = 'C'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
			  &" union all "_
			  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO, 0 AS CANTIDAD_MES, 0 AS CANTIDAD_MES1, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock) "_
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD WITH(NOLOCK) ON UD.ID = TR.BOEXTENSION_ID"_
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"' /*AND UD.CLASIFICADOR_CC_FORD_ID = '7BE98AF1-88CD-42F8-B621-78E71D9BFFBF'*/"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioDosMesAnterior&"' and '"&diaFinDosMesAnterior&"'"_ 
			  &" and TR.ESTADO = 'C'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION) AS Q"_ 
			  &" group by Q.REFERENCIA_ID,Q.PRODUCTO order by 2"_

			' query HOJA 2 - FORD PLANTA
			query2="select Q.REFERENCIA_ID, Q.PRODUCTO AS PRODUCTO, CAST(SUM(Q.CANTIDAD_MES)AS FLOAT) AS MESACTUAL,"_ 
			  &" CAST(SUM(Q.CANTIDAD_MES1)AS FLOAT) AS UNMESANTERIOR,"_ 
			  &" CAST(SUM(Q.CANTIDAD_MES2)AS FLOAT) AS DOSMESANTERIOR,"_ 
			  &" ISNULL((SELECT CAST(PR.VALOR2_IMPORTE AS FLOAT) FROM PRECIO PR WITH(NOLOCK)"_ 
			  &" WHERE pr.HABILITADO = 'T'"_ 
			  &" AND PR.REFERENCIA_ID    = Q.REFERENCIA_ID"_
			  &" AND PR.INICIAL2_MOMENTO = (SELECT MAX(PRE.INICIAL2_MOMENTO) FROM PRECIO PRE WHERE PRE.REFERENCIA_ID = PR.REFERENCIA_ID)),0) AS PRECIO,"_
			  &" ROUND(CAST((SUM(Q.CANTIDAD_MES)+SUM(Q.CANTIDAD_MES1)+SUM(Q.CANTIDAD_MES2))/3 AS FLOAT),0) as PROMEDIO"_ 
			  &" from ("_ 
			  &" SELECT IOC.REFERENCIA_ID,IOC.DESCRIPCION as PRODUCTO, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES,0 AS CANTIDAD_MES1, 0 as CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD WITH(NOLOCK) ON UD.ID = TR.BOEXTENSION_ID"_ 
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"' AND UD.CLASIFICADOR_CC_FORD_ID = '49D577FE-839E-4D46-A59C-C52EC9C7B6EB'"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&fechaActualIncial&"' and '"&fechaActualFinal&"'"_ 
			  &" and TR.ESTADO <> 'N'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
			  &" union all"_ 
			  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO,0 AS CANTIDAD_MES, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES1, 0 AS CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD WITH(NOLOCK) ON UD.ID = TR.BOEXTENSION_ID"_
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"' AND UD.CLASIFICADOR_CC_FORD_ID = '49D577FE-839E-4D46-A59C-C52EC9C7B6EB'"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioMesAnterior&"' and '"&diaFinMesAnterior&"'"_ 
			  &" and TR.ESTADO = 'C'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
			  &" union all "_
			  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO, 0 AS CANTIDAD_MES, 0 AS CANTIDAD_MES1, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock) "_
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD WITH(NOLOCK) ON UD.ID = TR.BOEXTENSION_ID"_
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"' AND UD.CLASIFICADOR_CC_FORD_ID = '49D577FE-839E-4D46-A59C-C52EC9C7B6EB'"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioDosMesAnterior&"' and '"&diaFinDosMesAnterior&"'"_ 
			  &" and TR.ESTADO = 'C'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION) AS Q"_ 
			  &" group by Q.REFERENCIA_ID,Q.PRODUCTO order by 2"_

		else ' ELSE IF 3
			' cualquier centro de costos que no sea Fiat o Ford
			' 27/09/2019 - Ahora Fiat Si.
			query = "select Q.REFERENCIA_ID, Q.PRODUCTO AS PRODUCTO, CAST(SUM(Q.CANTIDAD_MES)AS FLOAT) AS MESACTUAL,"_ 
			  &" CAST(SUM(Q.CANTIDAD_MES1)AS FLOAT) AS UNMESANTERIOR,"_ 
			  &" CAST(SUM(Q.CANTIDAD_MES2)AS FLOAT) AS DOSMESANTERIOR,"_ 
			  &" ISNULL((SELECT CAST(PR.VALOR2_IMPORTE AS FLOAT) FROM PRECIO PR WITH(NOLOCK)"_ 
			  &" WHERE pr.HABILITADO = 'T'"_ 
			  &" AND PR.REFERENCIA_ID    = Q.REFERENCIA_ID"_
			  &" AND PR.INICIAL2_MOMENTO = (SELECT MAX(PRE.INICIAL2_MOMENTO) FROM PRECIO PRE WHERE PRE.REFERENCIA_ID = PR.REFERENCIA_ID)),0) AS PRECIO,"_
			  &" ROUND(CAST((SUM(Q.CANTIDAD_MES)+SUM(Q.CANTIDAD_MES1)+SUM(Q.CANTIDAD_MES2))/3 AS FLOAT),0) as PROMEDIO"_ 
			  &" from ("_ 
			  &" SELECT IOC.REFERENCIA_ID,IOC.DESCRIPCION as PRODUCTO, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES,0 AS CANTIDAD_MES1, 0 as CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"'"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&fechaActualIncial&"' and '"&fechaActualFinal&"'"_ 
			  &" and TR.ESTADO <> 'N'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
			  &" union all"_ 
			  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO,0 AS CANTIDAD_MES, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES1, 0 AS CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock)"_ 
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"'"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioMesAnterior&"' and '"&diaFinMesAnterior&"'"_ 
			  &" and TR.ESTADO = 'C'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION"_ 
			  &" union all "_
			  &" SELECT IOC.REFERENCIA_ID, IOC.DESCRIPCION as PRODUCTO, 0 AS CANTIDAD_MES, 0 AS CANTIDAD_MES1, SUM(IOC.CANTIDAD2_CANTIDAD) AS CANTIDAD_MES2"_ 
			  &" FROM TRORDENCOMPRA AS TR with(nolock) "_
			  &" INNER JOIN ITEMORDENCOMPRA AS IOC with(nolock) ON IOC.PLACEOWNER_ID = TR.ID"_ 
			  &" INNER JOIN V_CENTROCOSTOS AS CC with(nolock) ON CC.ID = TR.CENTROCOSTOS_ID"_ 
			  &" INNER JOIN UD_PEDIDOINTERNOCOMPRA AS UD with(nolock) ON UD.ID = TR.BOEXTENSION_ID"_ 
			  &" WHERE ud.TIPOPIC_N = 'I - Insumos'"_
              &" AND TR.TIPOTRANSACCION_ID = '0F50AC65-F60A-4E27-AB77-3D1B4E37A154'"_ 
			  &" AND CC.id = '"&SELF.ID&"'"_ 
			  &" and LEFT(UD.FECHAREGISTRACION, 8) between '"&diaInicioDosMesAnterior&"' and '"&diaFinDosMesAnterior&"'"_ 
			  &" and TR.ESTADO = 'C'"_
			  &" GROUP BY IOC.REFERENCIA_ID,IOC.DESCRIPCION) AS Q"_ 
			  &" group by Q.REFERENCIA_ID,Q.PRODUCTO order by 2"_
		END IF ' FIN 3 IF
	end if ' FIN 2 IF
	
	' EXCEL.
	call ProgressControl(Self.Workspace, "PICs POR PERIODO POR CENTRO DE COSTO" , 0, 300)
	Set HojaExcel = CreateObject("Excel.Application")
	HojaExcel.Workbooks.Add
	
	IF self.id = "{E301E71F-829F-4DA3-ADA3-977F0E6F46BF}" THEN ' FORD
	    '------HOJA 1---- PINTURA QUERY1
		HojaExcel.Sheets("Hoja1").Select
	    HojaExcel.ActiveSheet.Cells(1, 1).Value = "PICs POR PERIODO POR CENTRO DE COSTO - FORD " 'PINTURA" 
		HojaExcel.ActiveSheet.Cells(1, 1).Font.Bold = True 	' Negrita.
		HojaExcel.ActiveSheet.Cells(1, 1).Interior.ColorIndex = 15 	' Fondo Verde.
		
		HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "Producto"
	    HojaExcel.ActiveSheet.Cells(3, 2).Value 		= nombreMesActual
	   	HojaExcel.ActiveSheet.Cells(3, 3).Value 		= nombreUnMesAnterior
	   	HojaExcel.ActiveSheet.Cells(3, 4).Value 		= nombreDosMesAnterior
	   	HojaExcel.ActiveSheet.Cells(3, 5).Value 		= "Precio Actual"
	   	HojaExcel.ActiveSheet.Cells(3, 6).Value        	= "Q Promedio"
	   	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Font.Bold 			= True 	' Negrita.
	   	HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		
		'referencias de colores
	  	HojaExcel.ActiveSheet.Cells(4, 8).Value 				= "REFERENCIA"
		HojaExcel.ActiveSheet.Cells(4, 8).Font.Bold 			= True 	' Negrita.
		HojaExcel.ActiveSheet.Cells(4, 8).Interior.ColorIndex 	= 15 	' Fondo Gris.
	  	HojaExcel.ActiveSheet.Cells(5, 8).Interior.ColorIndex 	= 3
	  	HojaExcel.ActiveSheet.Cells(5, 9).Value 				= "Cantidades Superiores al Mes anterior"
	  	HojaExcel.ActiveSheet.Cells(6, 8).Interior.ColorIndex 	= 42
	  	HojaExcel.ActiveSheet.Cells(6, 9).Value 				= "Cantidades Inferior al Mes anterior"
	  	HojaExcel.ActiveSheet.Cells(7, 8).Interior.ColorIndex 	= 43
	  	HojaExcel.ActiveSheet.Cells(7, 9).Value 				= "Cantidades Estables"
		 
		'CONSULTA SQL
		R = 4
		
	  	set xCon = CreateObject("adodb.connection")
	  	xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
	  	xCon.ConnectionTimeout 	= 150
		
	  	set xRs = RecordSet(xCon, "select top 1 * from producto")
	  	xRs.Close
	  	xRs.ActiveConnection.CommandTimeout = 0
	  	xRs.Source = query1
	  	xRs.Open
	  	do while not xRs.EOF
			call ProgressControlAvance(Self.Workspace, "Producto: " & CStr(xRs("PRODUCTO").Value))
		 	columna2 = CINT(xRs("MESACTUAL").Value)
			columna3 = CInt(xRs("UNMESANTERIOR").Value)
			if columna2 > columna3 then
				HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 3 	' Fondo Rojo. 
			else
				HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 42    ' Fondo Celeste
			end if
			if columna2 = columna3 then
				HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 43     'Fondo Verde
			end if 
			
			HojaExcel.Sheets("Hoja1").Select
			HojaExcel.ActiveSheet.Cells(R, 1).Value 	= Trim(CStr(xRs("PRODUCTO").Value))
		   	HojaExcel.ActiveSheet.Cells(R, 1).Font.Bold = True 	' Negrita.
			HojaExcel.ActiveSheet.Cells(R, 2).Value 	= CStr(xRs("MESACTUAL").Value)
			HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CStr(xRs("UNMESANTERIOR").Value)
			HojaExcel.ActiveSheet.Cells(R, 4).Value 	= CStr(xRs("DOSMESANTERIOR").Value)
			HojaExcel.ActiveSheet.Cells(R, 5).Value 	= CStr(xRs("precio").Value)
			HojaExcel.ActiveSheet.Cells(R, 6).Value 	= CStr(xRs("PROMEDIO").Value)
			
			calculo1 = calculo1 + (cint(xRs("MESACTUAL").Value) * cdbl(xRs("precio").Value))'importe total mes actual
			calculo2 = calculo2 + (cint(xRs("UNMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  UN MES ANTERIOR
			calculo3 = calculo3 + (cint(xRs("DOSMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  DOS MESES ANTERIORES
			  
			R = R + 1
			xRs.MoveNext
		loop
		
		HojaExcel.Sheets("Hoja1").Select
		HojaExcel.ActiveSheet.Cells(R+1, 1).Value  = "IMPORTE TOTAL"
		HojaExcel.ActiveSheet.Cells(R+1, 2).Value  = cStr(calculo1)
		HojaExcel.ActiveSheet.Cells(R+1, 3).Value  = cStr(calculo2)
		HojaExcel.ActiveSheet.Cells(R+1, 4).Value  = cStr(calculo3)
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 2), HojaExcel.ActiveSheet.Cells(R+1, 4)).NumberFormat			= "$ #,##0.00"
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 1), HojaExcel.ActiveSheet.Cells(R+1, 5)).Font.Bold 		    = True 	' Negrita.
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 2), HojaExcel.ActiveSheet.Cells(R+1, 5)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		
		HojaExcel.ActiveSheet.Columns("A:Z").AutoFit
		
		'------HOJA 2---- PLANTA QUERY2
		'HojaExcel.Sheets("Hoja2").Select
		'HojaExcel.ActiveSheet.Cells(1, 1).Value = "PICs POR PERIODO POR CENTRO DE COSTO - FORD PLANTA" 
		'HojaExcel.ActiveSheet.Cells(1, 1).Font.Bold = True 	' Negrita.
		'HojaExcel.ActiveSheet.Cells(1, 1).Interior.ColorIndex = 15 	' Fondo Verde.
		 
		'HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "PRODUCTO"
		'HojaExcel.ActiveSheet.Cells(3, 2).Value 		= nombreMesActual
		'HojaExcel.ActiveSheet.Cells(3, 3).Value 		= nombreUnMesAnterior
		'HojaExcel.ActiveSheet.Cells(3, 4).Value 		= nombreDosMesAnterior
		'HojaExcel.ActiveSheet.Cells(3, 5).Value 		= "PRECIO ACTUAL"
		'HojaExcel.ActiveSheet.Cells(3, 6).Value        = "CANTIDAD PROMEDIO"
		'HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Font.Bold 			= True 	' Negrita.
		'HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		 
		'referencias de colores
		'HojaExcel.ActiveSheet.Cells(4, 8).Value 		= "REFERENCIA"
		'HojaExcel.ActiveSheet.Cells(5, 8).Interior.ColorIndex = 3
		'HojaExcel.ActiveSheet.Cells(5, 9).Value 		= "Cantidades Superiores al Mes anterior"
		'HojaExcel.ActiveSheet.Cells(6, 8).Interior.ColorIndex = 42
		'HojaExcel.ActiveSheet.Cells(6, 9).Value 		= "Cantidades Inferior al Mes anterior"
		'HojaExcel.ActiveSheet.Cells(7, 8).Interior.ColorIndex = 43
		'HojaExcel.ActiveSheet.Cells(7, 9).Value 		= "Cantidades Estables"
		 
		'CONSULTA SQL
		'R = 4
		
		'set xCon = CreateObject("adodb.connection")
		'xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
		'xCon.ConnectionTimeout 	= 150
		
		'set xRs = RecordSet(xCon, "select top 1 * from producto")
		'xRs.Close
		'xRs.ActiveConnection.CommandTimeout = 0
		'xRs.Source = query2
		'xRs.Open
		'do while not xRs.EOF
		'	call ProgressControlAvance(Self.Workspace, "Producto: " & CStr(xRs("PRODUCTO").Value))
		'	columna2 = CINT(xRs("MESACTUAL").Value)
		'	columna3 = CInt(xRs("UNMESANTERIOR").Value)
		'	if columna2 > columna3 then
		'		HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 3 	' Fondo Rojo. 
		'	else
		'		HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 42    ' Fondo Celeste
		'	end if
		'	if columna2 = columna3 then
		'		HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 43     'Fondo Verde
		'	end if 
			  
		'	HojaExcel.Sheets("Hoja2").Select
		'	HojaExcel.ActiveSheet.Cells(R, 1).Value 	= CStr(xRs("PRODUCTO").Value)
		'	HojaExcel.ActiveSheet.Cells(R, 1).Font.Bold = True 	' Negrita.
		'	HojaExcel.ActiveSheet.Cells(R, 2).Value 	= CStr(xRs("MESACTUAL").Value)
		'	HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CStr(xRs("UNMESANTERIOR").Value)
		'	HojaExcel.ActiveSheet.Cells(R, 4).Value 	= CStr(xRs("DOSMESANTERIOR").Value)
		'	HojaExcel.ActiveSheet.Cells(R, 5).Value 	= CStr(xRs("precio").Value)
		'	HojaExcel.ActiveSheet.Cells(R, 6).Value 	= CStr(xRs("PROMEDIO").Value)
			
		'	calculo1 = calculo1 + (cint(xRs("MESACTUAL").Value) * cdbl(xRs("precio").Value))'importe total mes actual
		'	calculo2 = calculo2 + (cint(xRs("UNMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  UN MES ANTERIOR
		'	calculo3 = calculo3 + (cint(xRs("DOSMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  DOS MESES ANTERIORES
			
		'	R = R + 1
		'	xRs.MoveNext
		'loop
		
		'HojaExcel.Sheets("Hoja2").Select
		'HojaExcel.ActiveSheet.Cells(R+1, 1).Value  = "IMPORTE TOTAL"
		'HojaExcel.ActiveSheet.Cells(R+1, 2).Value  = cStr(calculo1)
		'HojaExcel.ActiveSheet.Cells(R+1, 3).Value  = cStr(calculo2)
		'HojaExcel.ActiveSheet.Cells(R+1, 4).Value  = cStr(calculo3)
		
		'HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 1), HojaExcel.ActiveSheet.Cells(R+1, 5)).Font.Bold 		    = True 	' Negrita.
		'HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 2), HojaExcel.ActiveSheet.Cells(R+1, 5)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		
	else 'EL RESTO DE LOS CENTRO DE COSTOS MENOS FORD 
		
		' --- HOJA1: DETALLE ---.
		HojaExcel.ActiveSheet.Cells(1, 1).Value = "PICs POR PERIODO POR CENTRO DE COSTO - " & Self.Nombre 
		HojaExcel.ActiveSheet.Cells(1, 1).Font.Bold = True 	' Negrita.
		HojaExcel.ActiveSheet.Cells(1, 1).Interior.ColorIndex = 15 	' Fondo Verde.
		
		HojaExcel.ActiveSheet.Cells(3, 1).Value 		= "PRODUCTO"
		IF ENTRA = TRUE THEN
			HojaExcel.ActiveSheet.Cells(3, 2).Value 		= nombreMesActual
			HojaExcel.ActiveSheet.Cells(3, 3).Value 		= nombrePrimeraQuincena
			HojaExcel.ActiveSheet.Cells(3, 4).Value 		= nombreSegundaQuincena
			HojaExcel.ActiveSheet.Cells(3, 5).Value 		= nombreTerceraQuincena
			HojaExcel.ActiveSheet.Cells(3, 6).Value 		= "Precio Actual"
			HojaExcel.ActiveSheet.Cells(3, 7).Value      	= "Q Promedio"
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 7)).Font.Bold 			= True 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 7)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		ELSE
			HojaExcel.ActiveSheet.Cells(3, 2).Value 		= nombreMesActual
			HojaExcel.ActiveSheet.Cells(3, 3).Value 		= nombreUnMesAnterior
			HojaExcel.ActiveSheet.Cells(3, 4).Value 		= nombreDosMesAnterior
			HojaExcel.ActiveSheet.Cells(3, 5).Value 		= "Precio Actual"
			HojaExcel.ActiveSheet.Cells(3, 6).Value      	= "Q Promedio"
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Font.Bold 			= True 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(3, 1), HojaExcel.ActiveSheet.Cells(3, 6)).Interior.ColorIndex 	= 15 	' Fondo Gris.
	    END IF
		
		'referencias de colores
		HojaExcel.ActiveSheet.Cells(4, 8).Value 				= "REFERENCIA"
		HojaExcel.ActiveSheet.Cells(4, 8).Font.Bold 			= True 	' Negrita.
		HojaExcel.ActiveSheet.Cells(4, 8).Interior.ColorIndex 	= 15 	' Fondo Gris.
		HojaExcel.ActiveSheet.Cells(5, 8).Interior.ColorIndex 	= 3
		HojaExcel.ActiveSheet.Cells(5, 9).Value 				= "Cantidades Superiores al Mes anterior"
		HojaExcel.ActiveSheet.Cells(6, 8).Interior.ColorIndex 	= 42
		HojaExcel.ActiveSheet.Cells(6, 9).Value 				= "Cantidades Inferior al Mes anterior"
		HojaExcel.ActiveSheet.Cells(7, 8).Interior.ColorIndex 	= 43
		HojaExcel.ActiveSheet.Cells(7, 9).Value 				= "Cantidades Estables"
		
		R = 4
		
		calculo = 0
		set xCon = CreateObject("adodb.connection")
		xCon.ConnectionString 	= StringConexion("calipso", Self.Workspace)
		xCon.ConnectionTimeout 	= 150
		
		set xRs = RecordSet(xCon, "select top 1 * from producto")
		xRs.Close
		xRs.ActiveConnection.CommandTimeout = 0
		xRs.Source = query
		xRs.Open
		do while not xRs.EOF
			call ProgressControlAvance(Self.Workspace, "Producto: " & CStr(xRs("PRODUCTO").Value))
			columna2 = CINT(xRs("MESACTUAL").Value)
			columna3 = CInt(xRs("UNMESANTERIOR").Value)
			if columna2 > columna3 then
				HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 3 	' Fondo Rojo. 
			else
				HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 42    ' Fondo Celeste
			end if
			if columna2 = columna3 then
				HojaExcel.ActiveSheet.Cells(R, 1).Interior.ColorIndex 	= 43     'Fondo Verde
			end if 
			
			IF ENTRA = TRUE THEN
				HojaExcel.ActiveSheet.Cells(R, 1).Value 	= Trim(CStr(xRs("PRODUCTO").Value))
				HojaExcel.ActiveSheet.Cells(R, 1).Font.Bold = True 	' Negrita.
				HojaExcel.ActiveSheet.Cells(R, 2).Value 	= CStr(xRs("MESACTUAL").Value)
				HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CStr(xRs("UNMESANTERIOR").Value)
				HojaExcel.ActiveSheet.Cells(R, 4).Value 	= CStr(xRs("DOSMESANTERIOR").Value)
				HojaExcel.ActiveSheet.Cells(R, 5).Value 	= CStr(xRs("TRESMESANTERIOR").Value)
				HojaExcel.ActiveSheet.Cells(R, 6).Value 	= CStr(xRs("precio").Value)
				HojaExcel.ActiveSheet.Cells(R, 7).Value 	= CStr(xRs("PROMEDIO").Value)
				calculo4 = calculo4 + (cint(xRs("TRESMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  DOS MESES ANTERIORES
			ELSE
				HojaExcel.ActiveSheet.Cells(R, 1).Value 	= Trim(CStr(xRs("PRODUCTO").Value))
				HojaExcel.ActiveSheet.Cells(R, 1).Font.Bold = True 	' Negrita.
				HojaExcel.ActiveSheet.Cells(R, 2).Value 	= CStr(xRs("MESACTUAL").Value)
				HojaExcel.ActiveSheet.Cells(R, 3).Value 	= CStr(xRs("UNMESANTERIOR").Value)
				HojaExcel.ActiveSheet.Cells(R, 4).Value 	= CStr(xRs("DOSMESANTERIOR").Value)
				HojaExcel.ActiveSheet.Cells(R, 5).Value 	= CStr(xRs("precio").Value)
				HojaExcel.ActiveSheet.Cells(R, 6).Value 	= CStr(xRs("PROMEDIO").Value)
			END IF
			calculo1 = calculo1 + (cint(xRs("MESACTUAL").Value) * cdbl(xRs("precio").Value))'importe total mes actual
			calculo2 = calculo2 + (cint(xRs("UNMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  UN MES ANTERIOR
			calculo3 = calculo3 + (cint(xRs("DOSMESANTERIOR").Value) * cdbl(xRs("precio").Value))'importe total  DOS MESES ANTERIORES
			R = R + 1
			xRs.MoveNext
		loop
		
		HojaExcel.ActiveSheet.Cells(R+1, 1).Value  = "IMPORTE TOTAL"
		HojaExcel.ActiveSheet.Cells(R+1, 2).Value  = cStr(calculo1)
		HojaExcel.ActiveSheet.Cells(R+1, 3).Value  = cStr(calculo2)
		HojaExcel.ActiveSheet.Cells(R+1, 4).Value  = cStr(calculo3)
		HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 2), HojaExcel.ActiveSheet.Cells(R+1, 4)).NumberFormat			= "$ #,##0.00"
		IF ENTRA = TRUE THEN
			HojaExcel.ActiveSheet.Cells(R+1, 5).Value  = cStr(calculo4)
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 1), HojaExcel.ActiveSheet.Cells(R+1, 5)).Font.Bold 		    = True 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 2), HojaExcel.ActiveSheet.Cells(R+1, 5)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		ELSE
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 1), HojaExcel.ActiveSheet.Cells(R+1, 4)).Font.Bold 		    = True 	' Negrita.
			HojaExcel.ActiveSheet.Range(HojaExcel.ActiveSheet.Cells(R+1, 2), HojaExcel.ActiveSheet.Cells(R+1, 4)).Interior.ColorIndex 	= 15 	' Fondo Gris.
		END IF
	END IF
	
	HojaExcel.ActiveSheet.Columns(5).NumberFormat	= "$ #,##0.00"
	HojaExcel.ActiveSheet.Columns("A:I").AutoFit
	call ProgressControlFinish(Self.Workspace)
	HojaExcel.Visible 	= true
	set HojaExcel 		= nothing
end sub
