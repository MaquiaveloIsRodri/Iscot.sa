SELECT
	ALIAS_0.ID ALIAS_0_ID,
	ALIAS_0.TIPOOBJETOESTATICO_ID ALIAS_0_TIPOOBJETOESTATICO_ID,
	ALIAS_0.CODIGO ALIAS_0_CODIGO,
	ALIAS_0.ACTIVESTATUS ALIAS_0_ACTIVESTATUS,
	ALIAS_0.DESCRIPCION ALIAS_0_DESCRIPCION,
	ALIAS_1.DESCRIPCION ALIAS_1_DESCRIPCION,
	ALIAS_2.NOMBRE ALIAS_2_NOMBRE,
	ALIAS_3.NOMBRE ALIAS_3_NOMBRE,
	ALIAS_0.USUARIO ALIAS_0_USUARIO,
	ALIAS_0.MOMENTO ALIAS_0_MOMENTO,
	ALIAS_4.VALOR2_IMPORTE ALIAS_4_VALOR2_IMPORTE,
	ALIAS_5.NOMBRE ALIAS_5_NOMBRE,
	ALIAS_6.DENOMINACION ALIAS_6_DENOMINACION
FROM
	V_SERVICIO_ ALIAS_0
	LEFT OUTER JOIN V_UD_SERVICIO_ ALIAS_8 ON ALIAS_0.BOEXTENSION_ID = ALIAS_8.ID
	LEFT OUTER JOIN V_ITEMTIPOCLASIFICADOR_ ALIAS_2 ON ALIAS_8.SECTOR_ID = ALIAS_2.ID
	LEFT OUTER JOIN V_ITEMTIPOCLASIFICADOR_ ALIAS_3 ON ALIAS_8.TIPOSERVICIO_ID = ALIAS_3.ID
	LEFT OUTER JOIN V_ITEMESQUEMAOPERATIVO_ ALIAS_9 ON ALIAS_0.ID = ALIAS_9.PLACEOWNER_ID
	LEFT OUTER JOIN V_CUENTASCONTABLES_ ALIAS_10 ON ALIAS_0.CUENTASCONTABLES_ID = ALIAS_10.ID
	LEFT OUTER JOIN V_CUENTA_ ALIAS_1 ON ALIAS_10.CUENTACONTABLE1_ID = ALIAS_1.ID
	INNER JOIN V_PRECIO_ ALIAS_4 ON ALIAS_0.ID = ALIAS_4.REFERENCIA_ID,
	V_UD_CENTROCOSTOS_ ALIAS_7
	INNER JOIN V_CENTROCOSTOS_ ALIAS_11 ON ALIAS_7.ID = ALIAS_11.BOEXTENSION_ID
	INNER JOIN V_UD_PLAZOENTREGA_ ALIAS_12 ON ALIAS_7.PLAZOENTREGA_ID = ALIAS_12.ID
	INNER JOIN V_PROVEEDOR_ ALIAS_6 ON ALIAS_12.PROVEEDOR_ID = ALIAS_6.ID
	INNER JOIN V_PERSONA_ ALIAS_5 ON ALIAS_6.ENTEASOCIADO_ID = ALIAS_5.ID
	INNER JOIN V_LISTAPRECIO_ ALIAS_13 ON ALIAS_6.LISTAPRECIO_ID = ALIAS_13.ID
WHERE
	ALIAS_0.BO_PLACE_ID = '{09D21CD8-2D3A-41D5-A18A-2E2DD64EC75D}'
	AND ALIAS_0.ACTIVESTATUS = 0
	AND ALIAS_0.TIPOOBJETOESTATICO_ID IS NULL
	AND ALIAS_9.ESQUEMAOPERATIVO_ID IS NULL
	AND (
		ALIAS_0.UNIDADOPERATIVA_ID = '{CEA52A93-43D9-429A-9708-AD75A1183343}'
		OR ALIAS_0.UNIDADOPERATIVA_ID IS NULL
		OR ALIAS_0.UNIDADOPERATIVA_ID = '{4B4CB6FD-E440-4E78-8130-A936B911D4E1}'
	)
	AND ALIAS_6.ACTIVESTATUS = 0
	AND ALIAS_4.DESDEFECHA <= '20240517072554000'
	AND ALIAS_4.HASTAFECHA >= '20240517072554000'
	AND ALIAS_11.ID = '{75A39710-4738-40DB-A208-DF9ABCAC3603}'