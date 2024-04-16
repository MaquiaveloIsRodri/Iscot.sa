SELECT Cliente
	,SUM(Total_Facturado) AS Total_Facturado
	,SUM(Total_Facturado) * 100.0 / ((CASE WHEN SUM(Total_Facturado) + SUM(Total_Previsionado) <> 0.0 THEN SUM(Total_Facturado) + SUM(Total_Previsionado) ELSE 1 END)) AS Porce_Facturado
	,SUM(Total_Previsionado) AS Total_Previsionado
	,SUM(Total_Previsionado) * 100.0 / ((CASE WHEN SUM(Total_Facturado) + SUM(Total_Previsionado) <> 0.0 THEN SUM(Total_Facturado) + SUM(Total_Previsionado) ELSE 1 END)) AS Porce_Previsionado
	,(SUM(Total_Facturado) + SUM(Total_Previsionado)) AS Total
	,SUM(Total_Diferencia) AS Total_Diferencia
	,(SUM(Total_Facturado) + SUM(Total_Previsionado)) + SUM(Total_Diferencia) AS Total_Final
	FROM (
		-- FACTURA - OS
		SELECT per.NOMBRE AS Cliente
		,SUM(ifa.TOTAL2_IMPORTE) AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM(ifa.TOTAL2_IMPORTE - (CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENSERVICIO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		INNER JOIN UD_ItemFacturaVenta AS uifa WITH(NOLOCK) ON ios.ID = uifa.ITEMOS_ID
		INNER JOIN ITEMFACTURAVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
		INNER JOIN TRFACTURAVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		WHERE os.ESTADO <> 'N'


        AND (SELECT TOP 1 ep.CODIGO AS CodEP FROM UD_ESTADOFACTURACION AS oef WITH(NOLOCK) 
        LEFT JOIN ITEMTIPOCLASIFICADOR AS ep WITH(NOLOCK) ON oef.ESTADOPAGO_ID = ep.ID
        WHERE oef.BO_PLACE_ID = udfv.ESTADOFACTURA_ID 
        ORDER BY oef.FECHA DESC) <> '05'



		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{0E3520FE-9A43-4CBF-96E0-F7DFA07C8E73}'			-- Pérdida
		AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'			-- Cerrada
		AND fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'		-- Facturado
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL

		-- FACTURA - OC
		SELECT per.NOMBRE AS Cliente
		,SUM(ifa.TOTAL2_IMPORTE) * -1.0 AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM(((CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END) * -1) - ifa.TOTAL2_IMPORTE) AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENCREDITO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		INNER JOIN UD_ItemFacturaVenta AS uifa WITH(NOLOCK) ON ios.ID = uifa.ITEMOS_ID
		INNER JOIN ITEMFACTURAVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
		INNER JOIN TRFACTURAVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'			-- Ganancia
		AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'			-- Cerrada
		AND fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'		-- Facturado
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL
		
		-- NOTA DEBITO - OS
		SELECT per.NOMBRE AS Cliente
		,SUM(ifa.TOTAL2_IMPORTE) AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM(ifa.TOTAL2_IMPORTE - (CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENSERVICIO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		INNER JOIN UD_ITEMDEBITOVENTA AS uifa WITH(NOLOCK) ON ios.ID = uifa.ITEMOS_ID
		INNER JOIN ITEMDEBITOVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
		INNER JOIN TRDEBITOVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{0E3520FE-9A43-4CBF-96E0-F7DFA07C8E73}'			-- Pérdida
		AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'			-- Cerrada
		AND fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'		-- Facturado
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL
		
		-- NOTA DEBITO - OC
		SELECT per.NOMBRE AS Cliente
		,SUM(ifa.TOTAL2_IMPORTE) * -1.0 AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM(((CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END) * -1) - ifa.TOTAL2_IMPORTE) AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENCREDITO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		INNER JOIN UD_ITEMDEBITOVENTA AS uifa WITH(NOLOCK) ON ios.ID = uifa.ITEMOS_ID
		INNER JOIN ITEMDEBITOVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
		INNER JOIN TRDEBITOVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'			-- Ganancia
		AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'			-- Cerrada
		AND fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'		-- Facturado
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL
		
		-- NOTA CREDITO - FV
		SELECT DISTINCT per.NOMBRE AS Cliente
		,fa.VALORTOTAL * -1.0 AS Total_Facturado
		,0.0 AS Total_Previsionado
		,fa.VALORTOTAL - os.VALORTOTAL AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENCREDITO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		INNER JOIN TRCREDITOVENTA AS fa WITH(NOLOCK) ON os.VINCULOTR_ID = fa.ID
		INNER JOIN UD_CREDITOVENTA AS ufa WITH(NOLOCK) ON fa.BOEXTENSION_ID = ufa.ID
		INNER JOIN ITEMCREDITOVENTA AS ifa WITH(NOLOCK) ON fa.ID = ifa.PLACEOWNER_ID
		INNER JOIN UD_ITEMCREDITOVENTA AS uifa WITH(NOLOCK) ON ifa.BOEXTENSION_ID = uifa.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'		-- Anulado
		AND os.FLAG_ID <> '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'		-- Ganancia
		AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'		-- Cerrada
		AND fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		AND uifa.ITEMFV_ID IN (SELECT ifa.ID FROM ITEMFACTURAVENTA AS ifa
				INNER JOIN UD_ITEMFACTURAVENTA AS uifa WITH(NOLOCK) ON ifa.BOEXTENSION_ID = uifa.ID
				INNER JOIN TRFACTURAVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
				WHERE fa.ESTADO = 'C'
				AND fa.EXTERNAL_ID = '')

		UNION ALL
		
		-- NOTA CREDITO - OC
		SELECT per.NOMBRE AS Cliente
		,SUM(ifa.TOTAL2_IMPORTE) * -1.0 AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM(ifa.TOTAL2_IMPORTE - (CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENCREDITO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		INNER JOIN UD_ITEMCREDITOVENTA AS uifa WITH(NOLOCK) ON ios.ID = uifa.ITEMFV_ID
		INNER JOIN ITEMCREDITOVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
		INNER JOIN TRCREDITOVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'			-- Ganancia
		AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'			-- Cerrada
		AND fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'		-- Facturado
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL

		-- OrSer Previsionada
		SELECT per.NOMBRE AS Cliente
		,0.0 AS Total_Facturado
		,SUM((CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) AS Total_Previsionado
		,0.0 AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENSERVICIO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{0E3520FE-9A43-4CBF-96E0-F7DFA07C8E73}'			-- Pérdida
		AND uios.FLAGITEM_ID = '{CBE693B6-0498-4323-8BCD-13384693FBD1}'		-- Pendiente de Facturar
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL

		-- OrCre Previsionada
		SELECT per.NOMBRE AS Cliente
		,0.0 AS Total_Facturado
		,SUM((CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) * -1.0 AS Total_Previsionado
		,0.0 AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENCREDITO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND os.FLAG_ID <> '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'			-- Ganancia
		AND uios.FLAGITEM_ID = '{CBE693B6-0498-4323-8BCD-13384693FBD1}'		-- Pendiente de Facturar
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL

		-- OrSer Pérdida
		SELECT per.NOMBRE AS Cliente
		,0.0 AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM((CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) * -1.0 AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENSERVICIO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND uios.FLAGITEM_ID = '{0E3520FE-9A43-4CBF-96E0-F7DFA07C8E73}'		-- Pérdida
		AND RIGHT(uios.PERIODOPERDIDA, 4) + LEFT(uios.PERIODOPERDIDA, 2) = '20240416'
		GROUP BY cli.ID, per.NOMBRE

		UNION ALL

		-- OrCre Ganancia
		SELECT per.NOMBRE AS Cliente
		,0.0 AS Total_Facturado
		,0.0 AS Total_Previsionado
		,SUM((CASE WHEN ios.VOLUMEN <> 0.0 THEN ios.VOLUMEN ELSE ios.TOTAL2_IMPORTE END)) AS Total_Diferencia
		FROM TRORDENVENTA AS os WITH(NOLOCK)
		INNER JOIN CLIENTE AS cli WITH(NOLOCK) ON os.DESTINATARIO_ID = cli.ID
		INNER JOIN V_PERSONA AS per WITH(NOLOCK) ON cli.ENTEASOCIADO_ID = per.ID
		INNER JOIN UD_ORDENCREDITO AS uos WITH(NOLOCK) ON os.BOEXTENSION_ID = uos.ID
		INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
		INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
		WHERE os.ESTADO <> 'N'
		AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'			-- Anulado
		AND uios.FLAGITEM_ID = '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'		-- Ganancia
		AND LEFT(uos.FECHASERVICIO, 6) = '20240416'
		GROUP BY cli.ID, per.NOMBRE
	) Q
	GROUP BY Cliente
	ORDER BY Cliente



ISNULL((SELECT TOP 1 ep.CODIGO AS CodEP 
FROM UD_ESTADOFACTURACION AS oef WITH(NOLOCK) 
LEFT JOIN ITEMTIPOCLASIFICADOR AS ep WITH(NOLOCK) ON oef.ESTADOPAGO_ID = ep.ID
WHERE oef.BO_PLACE_ID = udfv.ESTADOFACTURA_ID
ORDER BY oef.FECHA DESC), '00') AS EstadoPago