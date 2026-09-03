/* QUERY TIEMPOS DE OPERACIÓN SANOFI SIN FINES DE SEMANA CON CLASIFICACIÓN PRODUCTIVO/NO PRODUCTIVO */
/* Ultima modificacion: [03/06/2026] */
/* Incluye clasificación de mercancías - CORREGIDO COLLATION */

WITH ClasificacionMercancias AS (

    SELECT DISTINCT
        [DESCRIPCIÓN_PRODUCTO] COLLATE SQL_Latin1_General_CP1_CI_AS AS [DESCRIPCIÓN_PRODUCTO],
        'PRODUCTIVO' AS Clasificacion
    FROM [BI].[dbo].[Track and Trace productivo]
    WHERE [DESCRIPCIÓN_PRODUCTO] IS NOT NULL 
      AND LTRIM(RTRIM([DESCRIPCIÓN_PRODUCTO])) <> ''
    
    UNION ALL
    
    -- No productivos
    SELECT DISTINCT
        [DESCRIPCIÓN_PRODUCTO] COLLATE SQL_Latin1_General_CP1_CI_AS AS [DESCRIPCIÓN_PRODUCTO],
        'NO PRODUCTIVO' AS Clasificacion
    FROM [BI].[dbo].[Track and Trace No productivo]
    WHERE [DESCRIPCIÓN_PRODUCTO] IS NOT NULL 
      AND LTRIM(RTRIM([DESCRIPCIÓN_PRODUCTO])) <> ''
),
RefEntradaPorPedimento AS (
    SELECT
        REPLACE(TRIM(Referencia), '/', '') AS ReferenciaNormalizada,
        RIGHT(TRIM(CAST(Pedimento AS VARCHAR(20))), 7) AS PedimentoUltimos7,
        MAX(NULLIF(TRIM(Mercancia), '')) AS MercanciaRefPedimento
    FROM [SIR].[Admin].[ADMIN_VT_CGReferenciaEntrada]
    GROUP BY
        REPLACE(TRIM(Referencia), '/', ''),
        RIGHT(TRIM(CAST(Pedimento AS VARCHAR(20))), 7)
),
RefEntradaPorReferencia AS (
    SELECT
        REPLACE(TRIM(Referencia), '/', '') AS ReferenciaNormalizada,
        MAX(NULLIF(TRIM(Mercancia), '')) AS MercanciaRefReferencia
    FROM [SIR].[Admin].[ADMIN_VT_CGReferenciaEntrada]
    GROUP BY REPLACE(TRIM(Referencia), '/', '')
),
ConsultaBase AS (
    SELECT
        Sucursal AS Sucursal,
        CASE
            WHEN Sucursal = 'CIUDAD DE MÉXICO' THEN 'AÉREO AICM'
            WHEN Sucursal = 'AIFA ESTADO DE MEXICO' THEN 'AÉREO AIFA'
            WHEN Sucursal = 'VERACRUZ' THEN 'MARÍTIMO VERACRUZ'
            WHEN CHARINDEX('LT', Referencia) > 0 THEN 'LAREDO'
            WHEN CHARINDEX('MNS', Referencia) > 0 THEN 'MANZANILLO'
            ELSE 'CORRESPONSALIAS'
        END AS [Tipo Sucursal],
        Referencia,

        --INFORME
        Cliente,
        [Patente],
        -- Si la clave NO empieza con R, lo mandamos al Original A1
        CASE 
            WHEN [Clave Pedimento] NOT LIKE 'R%' THEN Pedimento
            ELSE NULL 
        END AS [Pedimento Original A1],
        -- Si la clave empieza con R (como R1, R2, RT, etc.), lo mandamos al R1
        CASE 
            WHEN [Clave Pedimento] LIKE 'R%' THEN Pedimento
            ELSE NULL 
        END AS [Pedimento R1],
        [Tipo Operación Desc],
        [Clave Pedimento],
        [Tipo de Cambio de Pedimento],
        [Aduana/Sección Despacho] AS 'Aduana Despacho',
        [Contenedores],
        CASE 
        WHEN ISNULL([Contenedores], '') = '' THEN 0
        ELSE LEN([Contenedores]) - LEN(REPLACE([Contenedores], ',', '')) + 1 
        END AS [QTY Contenedor],
        --'QTY Contenedor',
        [Peso Bruto],
        --'Tipo Contenedor',
        --'# Pallet'
        --'# Bultos',
        [Total de Bultos],
        [Valor Comercial MXP],
        [Valor Aduana], 
        [Seguros],
        [Fletes],
        [Embalajes],
        [OtrosIncPed] AS 'Otros',
        [Valor Tasa Partida] AS 'TASA IGI/IGE',
        [IGI] AS 'IGI/IGIE',
        [Recargos],
        [Importe DTA FP1] AS 'DTA',
        [Importe IVA 1] AS 'IVA',
        [importeprv] AS 'Prevalidación',
        [IVA PRV] AS 'IVA Prevalidación',
        (
        ISNULL([IGI], 0) + 
        ISNULL([Recargos], 0) + 
        ISNULL([Importe DTA FP1], 0) + 
        ISNULL([Importe IVA 1], 0) + 
        ISNULL([importeprv], 0) + 
        ISNULL([IVA PRV], 0)
        ) AS 'Impuestos',
        [Facturas], 
        [RazonSocial de Proveedores],
        [Clave Incoterm],
        [Clave de País Origen/Destino],
        [Clave de País Vendedor/Comprador],
        [Moneda Factura] AS 'Moneda',
        [Guia Master] AS 'Bls MASTER',
        [Guia House] AS 'Bls HOUSE',
        [Fracciones],
        COALESCE(MERC.MercanciaFinal, main.Mercancía) AS Mercancía,
        [EJE TIPO DE MERCANCÍA] AS 'PT',
        [EMBARQUE REFRIGERADO] AS 'Tempreratura',
        COALESCE(NULLIF([UNIDAD RENTADA], ''), [UNIDAD SUPER EXPRESS]) AS [UNIDAD],
        COALESCE(NULLIF([PLACAS RENTADA], ''), [PLACAS SUPER EXPRESS]) AS [PLACAS],
        [Fecha Entrada/Presentación],


        

        


        
        
        /*---------------Función de Fecha de pago---------------*/
        CONVERT(VARCHAR(10),
        CASE
            WHEN Sucursal = 'CORRESPONSALIAS' THEN [Corresponsalias Fecha de Pago]
            WHEN Sucursal <> 'CORRESPONSALIAS' THEN [ Fecha de Pago]
            ELSE [ Fecha de Pago]
        END, 103) AS [Fecha de Pago funcion],

        /*---------------Entrada de pago-----------------*/
        (
            DATEDIFF(DAY,
                TRY_CAST([Fecha Entrada/Presentación] AS DATE),
                TRY_CAST(
                    CASE
                        WHEN Sucursal = 'CORRESPONSALIAS' THEN [Corresponsalias Fecha de Pago]
                        WHEN Sucursal <> 'CORRESPONSALIAS' THEN [ Fecha de Pago]
                        ELSE [ Fecha de Pago]
                    END AS DATE
                )
            ) + 1
            - (DATEDIFF(WEEK,
                TRY_CAST([Fecha Entrada/Presentación] AS DATE),
                TRY_CAST(
                    CASE
                        WHEN Sucursal = 'CORRESPONSALIAS' THEN [Corresponsalias Fecha de Pago]
                        WHEN Sucursal <> 'CORRESPONSALIAS' THEN [ Fecha de Pago]
                        ELSE [ Fecha de Pago]
                    END AS DATE
                )
            ) * 2)
            - (CASE WHEN DATENAME(WEEKDAY, TRY_CAST([Fecha Entrada/Presentación] AS DATE)) = 'Domingo' THEN 1 ELSE 0 END)
            - (CASE WHEN DATENAME(WEEKDAY, TRY_CAST(
                    CASE
                        WHEN Sucursal = 'CORRESPONSALIAS' THEN [Corresponsalias Fecha de Pago]
                        WHEN Sucursal <> 'CORRESPONSALIAS' THEN [ Fecha de Pago]
                        ELSE [ Fecha de Pago]
                    END AS DATE
                )) = 'Sábado' THEN 1 ELSE 0 END)
        ) AS [Entrada de pago],
        CASE
            WHEN [MOTIVO DE RETRASO] = 'OTRO' THEN [MOTIVO DE RETRASO OTRO]
            ELSE [MOTIVO DE RETRASO]
        END AS "MOTIVO DE RETRASO COMPLETO", 
        [FolioCuentaGastos] AS 'Cuenta de Gastos',
        [PermisosPartida] AS 'Permisos'

    FROM [SIR].[Admin].[SIR_VT_Sabana_Pedimento_ABC] main
    LEFT JOIN RefEntradaPorPedimento AS REF_E
        ON REPLACE(TRIM(main.Referencia), '/', '') = REF_E.ReferenciaNormalizada
        AND RIGHT(TRIM(CAST(main.Pedimento AS VARCHAR(20))), 7) = REF_E.PedimentoUltimos7
    LEFT JOIN RefEntradaPorReferencia AS REF_E2
        ON REPLACE(TRIM(main.Referencia), '/', '') = REF_E2.ReferenciaNormalizada
    OUTER APPLY (
        SELECT
            CASE
                WHEN main.Sucursal = 'CORRESPONSALIAS'
                     OR main.Referencia LIKE '%LT%'
                     OR main.Referencia LIKE '%MNS%'
                    THEN COALESCE(
                        REF_E.MercanciaRefPedimento,
                        NULLIF(TRIM(main.Mercancía), ''),
                        REF_E2.MercanciaRefReferencia
                    )
                ELSE COALESCE(
                    NULLIF(TRIM(main.Mercancía), ''),
                    REF_E.MercanciaRefPedimento,
                    REF_E2.MercanciaRefReferencia
                )
            END AS MercanciaFinal
    ) AS MERC
    /* JOIN para clasificación de mercancías CON COLLATION CORREGIDO */
    LEFT JOIN ClasificacionMercancias cm ON COALESCE(MERC.MercanciaFinal, main.Mercancía) COLLATE SQL_Latin1_General_CP1_CI_AS = cm.[DESCRIPCIÓN_PRODUCTO]

    WHERE 
         ((Cliente LIKE '%AVENTIS%' AND [EJE UNIDAD DE NEGOCIO] LIKE'CHC%'))
        --AND [Tipo Operación Desc] = 'Importación'
        --and [Clave Pedimento] not like 'R%'
)

SELECT 
    *,
    -- NUEVO CAMPO: Clasificación del Financiamiento según monto de Impuestos
    CASE 
        WHEN [Impuestos] BETWEEN 1 AND 24999 THEN 'FINANCIADO'
        WHEN [Impuestos] >= 25000 THEN 'PECE'
        ELSE 'SIN CLASIFICAR' 
    END AS [PECE]
FROM ConsultaBase
WHERE 
     TRY_CONVERT(DATE, [Fecha de Pago funcion], 103) >= '2026-08-01'
     AND TRY_CONVERT(DATE, [Fecha de Pago funcion], 103) <= '2026-08-31'
ORDER BY [Sucursal];
