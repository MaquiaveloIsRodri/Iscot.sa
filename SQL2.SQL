DECLARE@FechaParametro AS DATE = '2023-10-26';

WITH
    Periodos AS (
        SELECT
            cc.NOMBRE AS CENTROCOSTOS,
            LEFT (ud.FECHASERVICIO, 6) AS PERIODO,
            SUM(ios.TOTAL2_IMPORTE) AS TOTAL
        FROM
            TRORDENVENTA AS os
        WITH
            (NOLOCK)
            INNER JOIN UD_ORDENSERVICIO AS ud
        WITH
            (NOLOCK) ON os.BOEXTENSION_ID = ud.ID
            INNER JOIN CENTROCOSTOS AS cc
        WITH
            (NOLOCK) ON os.CENTROCOSTOS_ID = cc.ID
            INNER JOIN ITEMORDENVENTA AS ios
        WITH
            (NOLOCK) ON os.ID = ios.PLACEOWNER_ID
            INNER JOIN UD_ITEMORDENSERVICIO AS ud_ios
        WITH
            (NOLOCK) ON ios.BOEXTENSION_ID = ud_ios.ID
        WHERE
            os.TIPOTRANSACCION_ID = '{2D7CAB24-3F10-4235-9D0A-2819BDB83633}' -- Orden de Servicio
            AND os.FLAG_ID <> '{C10833F4-5E1E-47DA-803E-5FBF135BEA51}' -- No Anulada
            AND os.FLAG_ID <> '{D6B3920D-C3FA-4B1F-921E-D653E3B9E170}' -- No Facturada
            AND os.FLAG_ID <> '{0E3520FE-9A43-4CBF-96E0-F7DFA07C8E73}' -- No Pérdida
            AND os.ESTADO <> 'N'
            AND ud_ios.FLAGITEM_ID = '{CBE693B6-0498-4323-8BCD-13384693FBD1}' -- Pendiente de Facturar
            AND LEFT (ud.FECHASERVICIO, 6) >= FORMAT (@FechaParametro, 'yyyyMM')
        GROUP BY
            cc.NOMBRE,
            LEFT (ud.FECHASERVICIO, 6)
    )
SELECT
    CENTROCOSTOS,
    PERIODO,
    SUM(TOTAL) AS TOTAL
FROM
    Periodos
GROUP BY
    CENTROCOSTOS,
    PERIODO
ORDER BY
    CENTROCOSTOS,
    PERIODO;