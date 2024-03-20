
	SELECT  cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENSERVICIO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	WHERE os.TIPOTRANSACCION_ID = '{2D7CAB24-3F10-4235-9D0A-2819BDB83633}'		-- Orden de Servicio
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.FLAG_ID <> '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'					-- Facturada
	AND os.FLAG_ID <> '{0E3520FE-9A43-4CBF-96E0-F7DFA07C8E73}'					-- Pérdida
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{CBE693B6-0498-4323-8BCD-13384693FBD1}'				-- Pendiente de Facturar
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL'
    GROUP BY  cc.NOMBRE, LEFT(ud.FECHASERVICIO, 6)
	UNION ALL
	
	-- Segunda consulta
	SELECT  cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENSERVICIO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	INNER JOIN UD_ItemFacturaVenta AS uifa WITH(NOLOCK) ON uifa.ITEMOS_ID = ios.ID
	INNER JOIN ITEMFACTURAVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
	INNER JOIN TRFACTURAVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
	INNER JOIN UD_FACTURAVENTA AS ufa WITH(NOLOCK) ON fa.BOEXTENSION_ID = ufa.ID
	WHERE os.TIPOTRANSACCION_ID = '{2D7CAB24-3F10-4235-9D0A-2819BDB83633}'		-- Orden de Servicio
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'				-- Facturado
	AND CONVERT(date, ud.FECHASERVICIO, 103) <= CONVERT(date, '20231031', 103)
	AND fa.ESTADO = 'C'
	AND fa.EXTERNAL_ID = ''
	AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'					-- Cerrada
	AND CONVERT(date, LEFT(fa.FECHAACTUAL, 8), 103) > CONVERT(date, '20231031', 103)
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)


	UNION ALL

	-- Tercer consulta
	SELECT cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENSERVICIO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENSERVICIO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	INNER JOIN UD_ITEMDEBITOVENTA AS uind WITH(NOLOCK) ON uind.ITEMOS_ID = ios.ID
	INNER JOIN ITEMDEBITOVENTA AS ind WITH(NOLOCK) ON uind.ID = ind.BOEXTENSION_ID
	INNER JOIN TRDEBITOVENTA AS nd WITH(NOLOCK) ON ind.PLACEOWNER_ID = nd.ID
	INNER JOIN UD_DEBITOVENTA AS und WITH(NOLOCK) ON nd.BOEXTENSION_ID = und.ID
	WHERE os.TIPOTRANSACCION_ID = '{2D7CAB24-3F10-4235-9D0A-2819BDB83633}'		-- Orden de Servicio
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'				-- Facturado
	AND CONVERT(date, ud.FECHASERVICIO, 103) <= CONVERT(date, '20231031', 103)
	AND nd.ESTADO = 'C'
	AND nd.EXTERNAL_ID = ''
	AND nd.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'					-- Cerrada
	AND CONVERT(date, LEFT(nd.FECHAACTUAL, 8), 103) > CONVERT(date, '20231031', 103)
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)

	--tercer consulta

	UNION ALL

	SELECT cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) * -1 AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENCREDITO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	WHERE os.TIPOTRANSACCION_ID = '{7A3D273C-3F59-4637-9F3A-7658E7727329}'		-- Orden de Crédito
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.FLAG_ID <> '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'					-- Facturada
	AND os.FLAG_ID <> '{DDC07654-3AA9-42E9-AF8D-E95765E27D16}'					-- Ganancia
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{CBE693B6-0498-4323-8BCD-13384693FBD1}'			-- Pendiente de Facturar
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)

	--Cuarta consulta

	UNION ALL
	-- En Factura.
	SELECT cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) * -1 AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENCREDITO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	INNER JOIN UD_ItemFacturaVenta AS uifa WITH(NOLOCK) ON uifa.ITEMOS_ID = ios.ID
	INNER JOIN ITEMFACTURAVENTA AS ifa WITH(NOLOCK) ON uifa.ID = ifa.BOEXTENSION_ID
	INNER JOIN TRFACTURAVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
	INNER JOIN UD_FACTURAVENTA AS ufa WITH(NOLOCK) ON fa.BOEXTENSION_ID = ufa.ID
	WHERE os.TIPOTRANSACCION_ID = '{7A3D273C-3F59-4637-9F3A-7658E7727329}'		-- Orden de Crédito
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'				-- Facturado
	AND CONVERT(date, ud.FECHASERVICIO, 103) <= CONVERT(date, '20231031', 103)
	AND fa.ESTADO = 'C'
	AND fa.EXTERNAL_ID = ''
	AND fa.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'					-- Cerrada
	AND CONVERT(date, LEFT(fa.FECHAACTUAL, 8), 103) > CONVERT(date, '20231031', 103)
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)

	--quinta consulta
	UNION ALL
	SELECT cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) * -1 AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENCREDITO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	INNER JOIN UD_ITEMDEBITOVENTA AS uind WITH(NOLOCK) ON uind.ITEMOS_ID = ios.ID
	INNER JOIN ITEMDEBITOVENTA AS ind WITH(NOLOCK) ON uind.ID = ind.BOEXTENSION_ID
	INNER JOIN TRDEBITOVENTA AS nd WITH(NOLOCK) ON ind.PLACEOWNER_ID = nd.ID
	INNER JOIN UD_DEBITOVENTA AS und WITH(NOLOCK) ON nd.BOEXTENSION_ID = und.ID
	WHERE os.TIPOTRANSACCION_ID = '{7A3D273C-3F59-4637-9F3A-7658E7727329}'		-- Orden de Crédito
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'				-- Facturado
	AND CONVERT(date, ud.FECHASERVICIO, 103) <= CONVERT(date, '20231031', 103)
	AND nd.ESTADO = 'C'
	AND nd.EXTERNAL_ID = ''
	AND nd.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'					-- Cerrada
	AND CONVERT(date, LEFT(nd.FECHAACTUAL, 8), 103) > CONVERT(date, '20231031', 103)
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)



	UNION ALL
	SELECT cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) * -1 AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENCREDITO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	INNER JOIN UD_ITEMCREDITOVENTA AS uinc WITH(NOLOCK) ON uinc.ITEMFV_ID = ios.ID
	INNER JOIN ITEMCREDITOVENTA AS inc WITH(NOLOCK) ON uinc.ID = inc.BOEXTENSION_ID
	INNER JOIN TRCREDITOVENTA AS nc WITH(NOLOCK) ON inc.PLACEOWNER_ID = nc.ID
	INNER JOIN UD_CREDITOVENTA AS unc WITH(NOLOCK) ON nc.BOEXTENSION_ID = unc.ID
	WHERE os.TIPOTRANSACCION_ID = '{7A3D273C-3F59-4637-9F3A-7658E7727329}'		-- Orden de Crédito
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'				-- Facturado
	AND CONVERT(date, ud.FECHASERVICIO, 103) <= CONVERT(date, '20231031', 103)
	AND nc.ESTADO = 'C'
	AND nc.EXTERNAL_ID = ''
	AND nc.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'					-- Cerrada
	AND CONVERT(date, LEFT(nc.FECHAACTUAL, 8), 103) > CONVERT(date, '20231031', 103)
	and LEFT(ud.FECHASERVICIO, 6) = '202312'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)

	UNION ALL

	SELECT cc.NOMBRE AS CENTROCOSTOS
	,LEFT(ud.FECHASERVICIO, 6) AS PERIODO
	,SUM(ios.TOTAL2_IMPORTE) * -1 AS TOTAL
	FROM TRORDENVENTA AS os WITH(NOLOCK)
	INNER JOIN UD_ORDENCREDITO AS ud WITH(NOLOCK) ON os.BOEXTENSION_ID = ud.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS sec WITH(NOLOCK) ON ud.SECTOR_ID = sec.ID
	INNER JOIN CENTROCOSTOS AS cc WITH(NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
	INNER JOIN UD_CENTROCOSTOS AS ucc WITH(NOLOCK) ON cc.BOEXTENSION_ID = ucc.ID
	LEFT JOIN EMPLEADO AS empl WITH(NOLOCK) ON ucc.RESPONSABLEFACTURACION_ID = empl.ID
	LEFT JOIN V_PERSONA AS pers WITH(NOLOCK) ON empl.ENTEASOCIADO_ID = pers.ID
	INNER JOIN ITEMORDENVENTA AS ios WITH(NOLOCK) ON os.ID = ios.PLACEOWNER_ID
	INNER JOIN UD_ITEMORDENCREDITO AS uios WITH(NOLOCK) ON ios.BOEXTENSION_ID = uios.ID
	INNER JOIN ITEMTIPOCLASIFICADOR AS ts WITH(NOLOCK) ON ud.TIPOSERVICIO_ID = ts.ID
	INNER JOIN TRCREDITOVENTA AS nc WITH(NOLOCK) ON os.VINCULOTR_ID = nc.ID
	INNER JOIN UD_CREDITOVENTA AS unc WITH(NOLOCK) ON nc.BOEXTENSION_ID = unc.ID
	INNER JOIN ITEMCREDITOVENTA AS inc WITH(NOLOCK) ON nc.ID = inc.PLACEOWNER_ID
	INNER JOIN UD_ITEMCREDITOVENTA AS uinc WITH(NOLOCK) ON inc.BOEXTENSION_ID = uinc.ID
	WHERE os.TIPOTRANSACCION_ID = '{7A3D273C-3F59-4637-9F3A-7658E7727329}'		-- Orden de Crédito
	AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}'					-- Anulada
	AND os.ESTADO <> 'N' 
	AND uios.FLAGITEM_ID = '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}'				-- Facturado
	AND CONVERT(date, ud.FECHASERVICIO, 103) <= CONVERT(date, '20231031' , 103)
	AND nc.ESTADO = 'C'
	AND nc.EXTERNAL_ID = ''
	AND nc.FLAG_ID = '{832CECD8-F5FD-408F-9962-E81D3DF36638}'					-- Cerrada
	AND CONVERT(date, LEFT(nc.FECHAACTUAL, 8), 103) > CONVERT(date, '20231031', 103)
	AND uinc.ITEMFV_ID IN (
		SELECT ifa.ID FROM ITEMFACTURAVENTA AS ifa
		INNER JOIN UD_ITEMFACTURAVENTA AS uifa WITH(NOLOCK) ON ifa.BOEXTENSION_ID = uifa.ID
		INNER JOIN TRFACTURAVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		INNER JOIN UD_FACTURAVENTA AS ufa WITH(NOLOCK) ON fa.BOEXTENSION_ID = ufa.ID
		WHERE fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND ufa.CAE <> ''


		UNION ALL

		SELECT ifa.ID FROM ITEMDEBITOVENTA AS ifa
		INNER JOIN UD_ITEMDEBITOVENTA AS uifa WITH(NOLOCK) ON ifa.BOEXTENSION_ID = uifa.ID
		INNER JOIN TRDEBITOVENTA AS fa WITH(NOLOCK) ON ifa.PLACEOWNER_ID = fa.ID
		INNER JOIN UD_DEBITOVENTA AS ufa WITH(NOLOCK) ON fa.BOEXTENSION_ID = ufa.ID
		WHERE fa.ESTADO = 'C'
		AND fa.EXTERNAL_ID = ''
		AND ufa.CAE <> '')
		and cc.NOMBRE = 'CNH FPT INDUSTRIAL'
	and cc.NOMBRE = 'CNH FPT INDUSTRIAL' GROUP BY  cc.NOMBRE,LEFT(ud.FECHASERVICIO, 6)

	ORDER BY [CENTROCOSTOS], [PERIODO]