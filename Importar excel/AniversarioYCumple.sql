select
    ISNULL (
        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
        0
    )
from
    V_ITEMCONTABLE item2
with
    (nolock)
where
    item2.CENTROCOSTOS_ID = "{E7B5CE9A-34BA-4750-A27F-CA928DAB9613}"
    and item2.ESTADOTR = 'C'
    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
    and CAST(SUBSTRING (item2.FECHAVENCIMIENTO, 5, 2) as Int) = 5
    and CAST(LEFT (item2.FECHAVENCIMIENTO, 4) as Int) = 2024
    and item2.REFERENCIA_ID IN (
        select
            ID
        from
            CUENTA
        with
            (nolock)
        where
            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
            and ACTIVESTATUS = 0
    )