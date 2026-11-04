import { getFactory } from './factory.helper';

/**
 * SQL of Cat9&12: the page (rows + total count, once for one factory and once per factory for the
 * "ALL" page), the Excel export and the auto-sent CMS payload. The big inner SELECT is the same in both, only the wrapper around it differs.
 * `where` / `where1` are the filters of the two halves of the UNION (see buildCat9AndCat12Filter).
 */

/** Date column of the Excel / CMS queries; they spell it without a space ("ExFty_Date,23"). */
export const EXPORT_DATE_COLUMN = 'sb.ExFty_Date,23';

/**
 * WHERE parts + replacements for the date range and the RY (Product / Component) filter.
 * `dateColumn` is the date expression written inside CONVERT(VARCHAR ,<dateColumn>) (the exports spell it without a space).
 */
export const buildCat9AndCat12Filter = (
  dateFrom: string,
  dateTo: string,
  ry: string,
  dateColumn = 'sb.ExFty_Date ,23',
) => {
  let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  const replacements: any[] = [];

  if (dateFrom && dateTo) {
    where += ` AND CONVERT(VARCHAR ,${dateColumn}) BETWEEN ? AND ?`;
    where1 += ` AND CONVERT(VARCHAR ,${dateColumn}) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo, dateFrom, dateTo);
  }

  if (ry === 'Product') {
    where += ` AND CHARINDEX('-', id.RYNO) > 0`;
    where1 += ` AND CHARINDEX('-', id.RYNO) > 0`;
  } else if (ry === 'Component') {
    where += ` AND id.RYNO LIKE '%[BSU]%' AND CHARINDEX('-' ,id.RYNO)=0`;
    where1 += ` AND id.RYNO LIKE '%[BSU]%' AND CHARINDEX('-' ,id.RYNO)=0`;
  }

  return { where, where1, replacements };
};

/** Inner SELECT (invoices UNION sample invoices) used by the rows query. */
const buildPageInnerSql = (
  where: string,
  where1: string,
) => `SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE 
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                                WHEN do.ShipMode_1 IS NULL THEN ''
                                                ELSE do.ShipMode_1 END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING_D
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE 
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                              ${where1}`;

/** Inner SELECT used by the count query (same columns, but a different Transport_Method CASE and no Booking_No). */
const buildCountInnerSql = (
  where: string,
  where1: string,
) => `SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE 
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN 'AIR'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN 'AIR'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN 'SEA'
                                                WHEN do.Shipmode_1 LIKE '%SC%' THEN 'SEA'
                                                WHEN do.Shipmode_1 LIKE '%AC%' THEN 'AIR'
                                                WHEN do.Shipmode_1 LIKE '%CC%' THEN 'AIR'
                                                WHEN do.ShipMode_1 IS NULL THEN '' 
                                                ELSE 'N/A' END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING_D
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE 
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                              ${where1}`;

/** Rows query of one factory. */
export const buildCat9AndCat12PageQuery = (
  where: string,
  where1: string,
) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              ${buildPageInnerSql(where, where1)}
                          ) AS Cat9AndCat12`;

/** Count query of one factory. */
export const buildCat9AndCat12CountQuery = (
  where: string,
  where1: string,
) => `SELECT COUNT(*) AS total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              ${buildCountInnerSql(where, where1)}
                          ) AS Cat9AndCat12
                        ) AS sub`;

/** Rows query of one factory for the "ALL" page (adds the Factory_Name column). */
export const buildCat9AndCat12AllPageQuery = (
  factoryName: string,
  where: string,
  where1: string,
) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          ,'${factoryName}' AS Factory_Name
                    FROM   (
                              ${buildPageInnerSql(where, where1)}
                          ) AS Cat9AndCat1`;

/** Count query of one factory for the "ALL" page. */
export const buildCat9AndCat12AllCountQuery = (
  factoryName: string,
  where: string,
  where1: string,
) => `SELECT COUNT(*) AS total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          ,'${factoryName}' AS Factory_Name
                    FROM   (
                              ${buildCountInnerSql(where, where1)}
                          ) AS Cat9AndCat12
                        ) AS sub`;

/** Rows query of the Excel export of one factory (like the page query: no Booking_No column). */
export const buildCat9AndCat12ExportQuery = (
  where: string,
  where1: string,
) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE 
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                                WHEN do.ShipMode_1 IS NULL THEN ''
                                                ELSE do.ShipMode_1 END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING_D
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE 
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                              ${where1}
                          ) AS Cat9AndCat12`;

/**
 * Query of the auto-sent CMS payload of one factory (adds Factory / Factory_address / Factory_Name columns).
 * `factoryAddress` is the address text, 'N/A' when the factory has none.
 */
export const buildCat9AndCat12AutoSentQuery = (
  factory: string,
  factoryAddress: string,
  where: string,
  where1: string,
) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          , N'${getFactory(factory)}' AS Factory
                          , N'${factoryAddress}' AS Factory_address
                          ,'${factory}' AS Factory_Name
                  FROM   (
                            SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                                WHEN do.ShipMode_1 IS NULL THEN ''
                                                ELSE do.ShipMode_1 END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING_D
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,IIF(CHARINDEX('-' ,id.RYNO)>0 ,id.RYNO ,'N/A') AS RYProduct
                                    ,IIF(
                                          id.RYNO LIKE '%[BSU]%'
                                          AND CHARINDEX('-' ,id.RYNO)=0
                                        ,id.RYNO
                                        ,'N/A'
                                      )                       AS RYComponent
                                    ,CASE
                                          WHEN id.RYNO LIKE '%B%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'BOTTOM UNIT'
                                          WHEN id.RYNO LIKE '%S%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'SOCKLINER'
                                          WHEN id.RYNO LIKE '%U%'
                                                AND CHARINDEX('-' ,id.RYNO)=0 THEN 'UPPER UNIT'
                                          ELSE 'FINISH SHOE'
                                      END                     AS ComponentName
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                              ${where1}
                        ) AS Cat9AndCat12`;
