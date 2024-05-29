select
    ISNULL(SUM(TotalMaestroEne), 0) as TotalMaestroEne,
    ISNULL(SUM(TotalMaestroFeb), 0) as TotalMaestroFeb,
    ISNULL(SUM(TotalMaestroMar), 0) as TotalMaestroMar,
    ISNULL(SUM(TotalMaestroAbr), 0) as TotalMaestroAbr,
    ISNULL(SUM(TotalMaestroMay), 0) as TotalMaestroMay,
    ISNULL(SUM(TotalMaestroJun), 0) as TotalMaestroJun,
    ISNULL(SUM(TotalMaestroJul), 0) as TotalMaestroJul,
    ISNULL(SUM(TotalMaestroAgo), 0) as TotalMaestroAgo,
    ISNULL(SUM(TotalMaestroSep), 0) as TotalMaestroSep,
    ISNULL(SUM(TotalMaestroOct), 0) as TotalMaestroOct,
    ISNULL(SUM(TotalMaestroNov), 0) as TotalMaestroNov,
    ISNULL(SUM(TotalMaestroDic), 0) as TotalMaestroDic
from
    (
        select
            ISNULL(SUM(item.HABER2_IMPORTE - item.DEBE2_IMPORTE), 0) as TotalMaestroEne,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 2
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroFeb,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 3
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroMar,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 4
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroAbr,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 5
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroMay,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 6
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroJun,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 7
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroJul,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 8
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroAgo,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 9
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroSep,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 10
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroOct,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 11
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroNov,
            (
                select
                    ISNULL(
                        SUM(item2.HABER2_IMPORTE - item2.DEBE2_IMPORTE),
                        0
                    )
                from
                    V_ITEMCONTABLE item2 with(nolock)
                where
                    item2.CENTROCOSTOS_ID = cc.ID
                    and item2.ESTADOTR = 'C'
                    and item2.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
                    and CAST(SUBSTRING(item2.FECHAVENCIMIENTO, 5, 2) as Int) = 12
                    and CAST(LEFT(item2.FECHAVENCIMIENTO, 4) as Int) = 2024
                    and item2.REFERENCIA_ID IN (
                        select
                            ID
                        from
                            CUENTA with(nolock)
                        where
                            ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                            and ACTIVESTATUS = 0
                    )
            ) as TotalMaestroDic
        from
            V_ITEMCONTABLE item with(nolock)
            inner join V_CUENTA cta with(nolock) on cta.ID = item.REFERENCIA_ID
            inner join V_CENTROCOSTOS cc with(nolock) on cc.ID = item.CENTROCOSTOS_ID
        where
            item.ESTADOTR = 'C'
            and item.TIPOTRANSACCION_ID = '{9BB81D09-5EF7-453F-8E29-BC5E33D4FFDA}'
            and CAST(SUBSTRING(item.FECHAVENCIMIENTO, 5, 2) as Int) = 1and CAST(LEFT(item.FECHAVENCIMIENTO, 4) as Int) = 2024
            and cta.ID in (
                select
                    ID
                from
                    V_CUENTA with(nolock)
                where
                    ACUMULA_ID = '{6CC63C25-1886-43DC-A11D-A8E81AE63C10}'
                    and ACTIVESTATUS = 0
            )
            and cc.ID in (
                '{B679CBCD-D7DA-43D1-B84B-4111023351A0}',
                '{1BBD81D6-72F4-4C4E-BC01-2EB7A2EAA75E}',
                '{64438C6B-DD17-456A-9E31-C8D95F983577}',
                '{1970EBB4-4C79-4E76-AD88-33AA1CB005A2}'
            )
        group by
            cc.ID,
            cta.CODIGO,
            cta.DESCRIPCION
    ) Q