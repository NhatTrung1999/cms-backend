export { buildCat6TripQuery } from 'src/helper/cat6.queries';

/** Cat9&12 export for one ERP database; `where` and `where1` are the filters of its two UNION parts. */
export const buildAllFactoryCat9AndCat12Query = (
  factoryName: string,
  where: string,
  where1: string,
) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          ,'${factoryName}' AS Factory_Name
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
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,pc.PortCode             AS Port_Of_Arrival
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
                                              FROM   PACKING
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
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'SEA'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
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
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,pc.PortCode             AS Port_Of_Arrival
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
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'AIR'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where1}
                          ) AS Cat9AndCat12`;

/** Cat5 export for one WMS database. */
export const buildAllFactoryCat5Query = (
  where: string,
) => `SELECT dwo.WASTE_DATE                     AS Waste_disposal_date
                          ,dwc.CONSOLIDATED_WASTE_CODE AS Consolidated_Waste
                          ,dwc.WASTE_CODE AS Waste_Code
                          ,dtv.TREATMENT_VENDOR_NAME          AS Vendor_Name
                          ,dtv.TREATMENT_VENDOR_ID            AS Vendor_ID
                          ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
                          ,dwc.LOCATION_CODE AS Location_Code
                          ,CAST('0' AS INT)                   AS Transportation_Distance_km
                          ,CASE 
                                WHEN dwo.HAZARDOUS<>'N/A' THEN 'hazardous waste'
                                WHEN dwo.NON_HAZARDOUS<>'N/A' THEN 'Non-hazardous waste'
                                ELSE NULL
                          END                                AS The_type_of_waste
                          ,CASE 
                                WHEN dwo.HAZARDOUS<>'N/A' THEN dwo.HAZARDOUS
                                WHEN dwo.NON_HAZARDOUS<>'N/A' THEN dwo.NON_HAZARDOUS
                                ELSE NULL
                          END                                AS Waste_type
                          ,dtm.TREATMENT_METHOD_ENGLISH_NAME  AS Waste_Treatment_method
                          ,dtm.TREATMENT_METHOD_ID            AS Treatment_Method_ID
                          ,dwo.QUANTITY                       AS Weight_of_waste_treated_Unit_kg
                          ,CAST('0' AS INT)                   AS TKT_Ton_km
                    FROM   dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                          LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                                ON  dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                          LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                                ON  dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                          LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                                ON  dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                          LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                                ON  dwc.WASTE_CODE = dwo.WASTE_CODE
                          LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                                ON  td.CODE = dwo.LOCATION_CODE
                    ${where}`;
