import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { getFactory } from './factory.helper';

export const buildQuery = async (
  dateFrom: string,
  dateTo: string,
  factory: string,
  db?: Sequelize,
) => {
  const queryAddress = `SELECT [Address]
                        FROM CMW_Info_Factory
                        WHERE CreatedFactory = '${factory}'`;

  const factoryAddress =
    (await db?.query(queryAddress, {
      type: QueryTypes.SELECT,
    })) || [];

  let where = 'WHERE 1=1';
  const replacements: any[] = [];
  if (dateFrom && dateTo) {
    where += ` AND CONVERT(DATE ,cgzl.CGDate) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo);
  }
  const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY cgzl.CGDate) AS INT) AS [No]
                        ,cgzl.CGDate                     PurDate
                        ,kcrk.ModifyDate                 RKDate
                        ,c.CGNO                          PurNo
                        ,kcrk.RKNO                       ReceivedNo
                        ,c.CLBH                          MatID
                        ,clzl.ywpm                       MatName
                        ,ZLCLSL.CLSL                     Qty_Usage
                        ,KCRKS.Qty                       Qty_Receive
                        ,ISNULL(
                            ISNULL(SN223_UnitWeight.UnitWeight ,imw.Total_Weight)
                            ,SN74A.UnitWeight
                        )                               UnitWeight
                        ,(
                            ISNULL(
                                ISNULL(SN223_UnitWeight.UnitWeight ,imw.Total_Weight)
                                ,SN74A.UnitWeight
                            )*KCRKS.Qty
                        )                               Weight_Unitkg
                        ,ISNULL(z.ZSDH ,CGZL.ZSBH)       SupplierCode
                        ,'${factory}' as FactoryCode
                        ,ISNULL(P.Style ,ZSZL.Style)     Style
                        ,CASE 
                              WHEN ISNULL(ZSZL.Country ,'')='' THEN NULL
                              WHEN (
                                      '${factory}' IN ('LYV' ,'LVL' ,'LHG')
                                      AND ZSZL.Country IN ('Vietnam' ,'Viet nam' ,'VN' ,' VIETNAM' ,'Viet Nam')
                                  ) 
                                  OR ('${factory}'='LYF' AND ZSZL.Country='Indonesia') 
                                  OR (
                                      '${factory}' IN ('LYM' ,'POL')
                                      AND ZSZL.Country IN ('MYANMAR' ,' DA JIA MYANMAR COMPANY LIMITED' ,'MY')
                                  ) THEN 'Land'
                              ELSE 'SEA + Land'
                        END                             Transportationmethod
                        ,isi.SupplierFullAddress      AS Departure
                        ,CASE 
                              WHEN ISNULL(isi.ThirdCountryLandTransport ,'')='' THEN 'N/A'
                              ELSE CAST(isi.ThirdCountryLandTransport AS VARCHAR)
                        END                             ThirdCountryLandTransport
                        ,isi.PortOfDeparture
                        ,isi.PortOfArrival
                        ,isi.Factory_Port             AS FactoryDomesticLandTransport
                        ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Destination
                        ,ISNULL(isi.ThirdCountryLandTransport ,0)+ISNULL(isi.Factory_Port ,0) LandTransportDistance
                        ,isi.SeaTransportDistance
                        ,isi.AirTransportDistance
                        ,CAST('0' AS INT)             AS LandTransportTonKilometers
                        ,CAST('0' AS INT)             AS SeaTransportTonKilometers
                        ,CAST('0' AS INT)             AS AirTransportTonKilometers
                  FROM   CGZLS                        AS c
                        LEFT JOIN cgzl
                              ON  cgzl.CGNO = c.CGNO
                        LEFT JOIN ZSZL
                              ON  CGZL.ZSBH = ZSZL.ZSDH
                        LEFT JOIN ZSZL_Prod P
                              ON  P.ZSDH = ZSZL.zsdh
                                  AND P.GSBH = cgzl.GSBH
                        LEFT JOIN ZSZL z
                              ON  z.zsdh = ISNULL(P.MZSDH ,ZSZL.MZSDH)
                        LEFT JOIN Imp_SuppIDCombine  AS isi
                              ON  isi.ZSDH = ZSZL.zsdh
                        LEFT JOIN clzl
                              ON  cldh = c.CLBH
                        LEFT JOIN kcrk
                              ON  kcrk.ZSNO = c.CGNO
                        LEFT JOIN kcrks
                              ON  kcrks.RKNO = kcrk.RKNO
                                  AND kcrks.CLBH = c.CLBH
                                  AND ISNULL(KCRKS.RKSB ,'')<>'DL'
                                  AND ISNULL(KCRKS.RKSB ,'')<>'NG'
                        LEFT JOIN (
                                  SELECT smi2.CLBH
                                        ,zszl.zsdh
                                        ,MAX(smi2.Supplier_Material_ID) Supplier_Material_ID
                                  FROM   SuppMatID AS smi2
                                        LEFT JOIN zszl
                                              ON  zszl.Zsdh_TW = smi2.CSBH
                                        INNER JOIN Imp_MaterialWeight
                                              ON  Imp_MaterialWeight.Supplier_Material_ID = REPLACE(
                                                      REPLACE(smi2.Supplier_Material_ID ,CHAR(10) ,'')
                                                    ,CHAR(13)
                                                    ,''
                                                  )
                                  WHERE  ISNULL(Total_Weight ,0)<>0
                                  GROUP BY
                                        smi2.CLBH
                                        ,zszl.zsdh
                              ) A
                              ON  A.CLBH = c.CLBH
                                  AND A.zsdh = CGZL.ZSBH
                        LEFT JOIN Imp_MaterialWeight imw
                              ON  imw.Supplier_Material_ID = A.Supplier_Material_ID
                        LEFT JOIN Setup_UnitWeight   AS SN223_UnitWeight
                              ON  SN223_UnitWeight.FormID = 'SN223'
                                  AND SN223_UnitWeight.SupplierID = CGZL.ZSBH
                                  AND SN223_UnitWeight.MatID = c.CLBH
                        LEFT JOIN Setup_UnitWeight   AS SN74A
                              ON  SN74A.FormID = 'SN74A'
                                  AND SN74A.SupplierID = 'ZZZZ'
                                  AND SN74A.MatID = c.CLBH
                        LEFT JOIN (
                                  SELECT SS.CGNO
                                        ,S2.CLBH
                                        ,ISNULL(SUM(S2.CLSL) ,0) AS CLSL
                                  FROM   ZLZLS2 S2
                                        LEFT JOIN CGZLSS SS
                                              ON  SS.ZLBH = S2.ZLBH
                                                  AND SS.CLBH = S2.CLBH
                                  GROUP BY
                                        SS.CGNO
                                        ,S2.CLBH
                              )ZLCLSL
                              ON  ZLCLSL.CGNO = c.CGNO
                                  AND ZLCLSL.CLBH = c.CLBH
                         LEFT JOIN (
                                  SELECT *
                                  FROM   CMW.CMW.dbo.CMW_PortCode_Cat1_4
                                  WHERE  FactoryCode = '${factory}'
                              )                       AS pc
                              ON  pc.SupplierID = ISNULL(z.ZSDH ,CGZL.ZSBH) COLLATE SQL_Latin1_General_CP1_CI_AS
                  ${where}`;
  const countQuery = `SELECT COUNT([No]) as total
                      FROM (
                            SELECT CAST(ROW_NUMBER() OVER(ORDER BY cgzl.CGDate) AS INT) AS [No]
                                  ,cgzl.CGDate                     PurDate
                                  ,kcrk.ModifyDate                 RKDate
                                  ,c.CGNO                          PurNo
                                  ,kcrk.RKNO                       ReceivedNo
                                  ,c.CLBH                          MatID
                                  ,clzl.ywpm                       MatName
                                  ,ZLCLSL.CLSL                     Qty_Usage
                                  ,KCRKS.Qty                       Qty_Receive
                                  ,ISNULL(
                                      ISNULL(SN223_UnitWeight.UnitWeight ,imw.Total_Weight)
                                      ,SN74A.UnitWeight
                                  )                               UnitWeight
                                  ,(
                                      ISNULL(
                                          ISNULL(SN223_UnitWeight.UnitWeight ,imw.Total_Weight)
                                          ,SN74A.UnitWeight
                                      )*KCRKS.Qty
                                  )                               Weight_Unitkg
                                  ,ISNULL(z.ZSDH ,CGZL.ZSBH)       SupplierCode
                                  ,ISNULL(P.Style ,ZSZL.Style)     Style
                                  ,CASE 
                                        WHEN ISNULL(ZSZL.Country ,'')='' THEN NULL
                                        WHEN (
                                                '${factory}' IN ('LYV' ,'LVL' ,'LHG')
                                                AND ZSZL.Country IN ('Vietnam' ,'Viet nam' ,'VN' ,' VIETNAM' ,'Viet Nam')
                                            ) 
                                            OR ('${factory}'='LYF' AND ZSZL.Country='Indonesia') 
                                            OR (
                                                '${factory}' IN ('LYM' ,'POL')
                                                AND ZSZL.Country IN ('MYANMAR' ,' DA JIA MYANMAR COMPANY LIMITED' ,'MY')
                                            ) THEN 'Land'
                                        ELSE 'SEA + Land'
                                  END                             Transportationmethod
                                  ,isi.SupplierFullAddress      AS Departure
                                  ,CASE 
                                        WHEN ISNULL(isi.ThirdCountryLandTransport ,'')='' THEN 'N/A'
                                        ELSE CAST(isi.ThirdCountryLandTransport AS VARCHAR)
                                  END                             ThirdCountryLandTransport
                                  ,isi.PortOfDeparture
                                  ,isi.PortOfArrival
                                  ,isi.Factory_Port             AS FactoryDomesticLandTransport
                                  ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['ADDRESS']}' AS Destination
                                  ,ISNULL(isi.ThirdCountryLandTransport ,0)+ISNULL(isi.Factory_Port ,0) LandTransportDistance
                                  ,isi.SeaTransportDistance
                                  ,isi.AirTransportDistance
                                  ,CAST('0' AS INT)             AS LandTransortTonKilometers
                                  ,CAST('0' AS INT)             AS SeaTransortTonKilometers
                                  ,CAST('0' AS INT)             AS AirTransortTonKilometers
                            FROM   CGZLS                        AS c
                                  LEFT JOIN cgzl
                                        ON  cgzl.CGNO = c.CGNO
                                  LEFT JOIN ZSZL
                                        ON  CGZL.ZSBH = ZSZL.ZSDH
                                  LEFT JOIN ZSZL_Prod P
                                        ON  P.ZSDH = ZSZL.zsdh
                                            AND P.GSBH = cgzl.GSBH
                                  LEFT JOIN ZSZL z
                                        ON  z.zsdh = ISNULL(P.MZSDH ,ZSZL.MZSDH)
                                  LEFT JOIN Imp_SuppIDCombine  AS isi
                                        ON  isi.ZSDH = ZSZL.zsdh
                                  LEFT JOIN clzl
                                        ON  cldh = c.CLBH
                                  LEFT JOIN kcrk
                                        ON  kcrk.ZSNO = c.CGNO
                                  LEFT JOIN kcrks
                                        ON  kcrks.RKNO = kcrk.RKNO
                                            AND kcrks.CLBH = c.CLBH
                                            AND ISNULL(KCRKS.RKSB ,'')<>'DL'
                                            AND ISNULL(KCRKS.RKSB ,'')<>'NG'
                                  LEFT JOIN (
                                            SELECT smi2.CLBH
                                                  ,zszl.zsdh
                                                  ,MAX(smi2.Supplier_Material_ID) Supplier_Material_ID
                                            FROM   SuppMatID AS smi2
                                                  LEFT JOIN zszl
                                                        ON  zszl.Zsdh_TW = smi2.CSBH
                                                  INNER JOIN Imp_MaterialWeight
                                                        ON  Imp_MaterialWeight.Supplier_Material_ID = REPLACE(
                                                                REPLACE(smi2.Supplier_Material_ID ,CHAR(10) ,'')
                                                              ,CHAR(13)
                                                              ,''
                                                            )
                                            WHERE  ISNULL(Total_Weight ,0)<>0
                                            GROUP BY
                                                  smi2.CLBH
                                                  ,zszl.zsdh
                                        ) A
                                        ON  A.CLBH = c.CLBH
                                            AND A.zsdh = CGZL.ZSBH
                                  LEFT JOIN Imp_MaterialWeight imw
                                        ON  imw.Supplier_Material_ID = A.Supplier_Material_ID
                                  LEFT JOIN Setup_UnitWeight   AS SN223_UnitWeight
                                        ON  SN223_UnitWeight.FormID = 'SN223'
                                            AND SN223_UnitWeight.SupplierID = CGZL.ZSBH
                                            AND SN223_UnitWeight.MatID = c.CLBH
                                  LEFT JOIN Setup_UnitWeight   AS SN74A
                                        ON  SN74A.FormID = 'SN74A'
                                            AND SN74A.SupplierID = 'ZZZZ'
                                            AND SN74A.MatID = c.CLBH
                                  LEFT JOIN (
                                            SELECT SS.CGNO
                                                  ,S2.CLBH
                                                  ,ISNULL(SUM(S2.CLSL) ,0) AS CLSL
                                            FROM   ZLZLS2 S2
                                                  LEFT JOIN CGZLSS SS
                                                        ON  SS.ZLBH = S2.ZLBH
                                                            AND SS.CLBH = S2.CLBH
                                            GROUP BY
                                                  SS.CGNO
                                                  ,S2.CLBH
                                        )ZLCLSL
                                        ON  ZLCLSL.CGNO = c.CGNO
                                            AND ZLCLSL.CLBH = c.CLBH
                            ${where}
                      ) as Sub`;
  return { query, countQuery };
};

export const buildQueryTest = async (
  sortField: string = 'No',
  sortOrder: string = 'asc',
  factory: string,
  db?: Sequelize,
  isExport: boolean = false
) => {
  const queryAddress = `SELECT [Address]
                        FROM CMW_Info_Factory
                        WHERE CreatedFactory = '${factory}'`;

  const factoryAddress =
    (await db?.query(queryAddress, {
      type: QueryTypes.SELECT,
    })) || [];

    const pagingSql = isExport
      ? ''
      : `OFFSET :offset ROWS
      FETCH NEXT :limit ROWS ONLY;`;
  const query = `
            IF OBJECT_ID('tempdb..#PurN233_CGZL') IS NOT NULL
            DROP TABLE #PurN233_CGZL

            SELECT cgzl.CGDate                  AS PurDate
                  ,c.CGNO                       AS PurNo
                  ,c.CLBH                       AS MatID
                  ,clzl.ywpm                    AS MatName
                  ,CASE 
                        WHEN LEFT(c.CLBH ,4)='U1AD' AND ISNULL(y.OWeigh ,0)<>0 THEN y.OWeigh
                        ELSE ISNULL(
                              ISNULL(SN223_UnitWeight.UnitWeight ,SN74A.UnitWeight)
                              ,imw.Total_Weight
                        )
                  END                          AS UnitWeight
                  ,ISNULL(z.ZSDH ,CGZL.ZSBH)    AS SupplierCode
                  ,ISNULL(P.Style ,ZSZL.Style)  AS Style
                  ,CASE 
                        WHEN ISNULL(ZSZL.Country ,'')='' THEN NULL
                        WHEN (
                              '${factory}' IN ('LYV' ,'LVL' ,'LHG')
                              AND ZSZL.Country IN ('Vietnam' ,'Viet nam' ,'VN' ,' VIETNAM' ,'Viet Nam')
                        )
                        OR ('${factory}'='LYF' AND ZSZL.Country='Indonesia')
                        OR (
                              '${factory}' IN ('LYM' ,'POL')
                              AND ZSZL.Country IN ('MYANMAR' ,' DA JIA MYANMAR COMPANY LIMITED' ,'MY')
                        ) THEN 'Land'
                        ELSE 'SEA + Land'
                  END                          AS TransportationMethod
                  ,isi.SupplierFullAddress      AS Departure
                  ,CASE 
                        WHEN ISNULL(isi.ThirdCountryLandTransport ,'')='' THEN 'N/A'
                        ELSE CAST(isi.ThirdCountryLandTransport AS VARCHAR)
                  END                          AS ThirdCountryLandTransport
                  ,isi.PortOfDeparture          AS PortOfDeparture
                  ,isi.PortOfArrival            AS PortOfArrival
                  ,isi.Factory_Port             AS FactoryDomesticLandTransport
                  ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['ADDRESS']}' AS Destination
                  ,ISNULL(isi.ThirdCountryLandTransport ,0)+ISNULL(isi.Factory_Port ,0) AS LandTransportDistance
                  ,isi.SeaTransportDistance     AS SeaTransportDistance
                  ,isi.AirTransportDistance     AS AirTransportDistance
            INTO   #PurN233_CGZL
            FROM   CGZLS AS c
                  LEFT JOIN cgzl
                        ON  cgzl.CGNO = c.CGNO
                  LEFT JOIN ZSZL
                        ON  CGZL.ZSBH = ZSZL.ZSDH
                  LEFT JOIN ZSZL_Prod P
                        ON  P.ZSDH = ZSZL.zsdh
                        AND P.GSBH = cgzl.GSBH
                  LEFT JOIN ZSZL z
                        ON  z.zsdh = ISNULL(P.MZSDH ,ZSZL.MZSDH)
                  LEFT JOIN Imp_SuppIDCombine AS isi
                        ON  isi.ZSDH = ZSZL.zsdh
                  LEFT JOIN clzl
                        ON  cldh = c.CLBH
                  LEFT JOIN YWWX2 AS y
                        ON  clzl.cldh = y.CLBH
                  LEFT JOIN (
                        SELECT smi2.CLBH
                              ,zszl.zsdh
                              ,MAX(smi2.Supplier_Material_ID) Supplier_Material_ID
                        FROM   SuppMatID AS smi2
                              LEFT JOIN zszl
                                    ON  zszl.Zsdh_TW = smi2.CSBH
                              INNER JOIN Imp_MaterialWeight
                                    ON  Imp_MaterialWeight.Supplier_Material_ID = REPLACE(
                                                REPLACE(smi2.Supplier_Material_ID ,CHAR(10) ,'')
                                          ,CHAR(13)
                                          ,''
                                          )
                        WHERE  ISNULL(Total_Weight ,0)<>0
                              AND (
                                          (LEFT(smi2.CLBH ,1) NOT IN ('X' ,'Y' ,'Z' ,'V'))
                                          OR (LEFT(smi2.CLBH ,4)='V501')
                                    )
                        GROUP BY
                              smi2.CLBH
                              ,zszl.zsdh
                        ) A
                        ON  A.CLBH = c.CLBH
                        AND A.zsdh = CGZL.ZSBH
                  LEFT JOIN Imp_MaterialWeight imw
                        ON  imw.Supplier_Material_ID = A.Supplier_Material_ID
                  LEFT JOIN Setup_UnitWeight AS SN223_UnitWeight
                        ON  SN223_UnitWeight.FormID = 'SN223'
                        AND SN223_UnitWeight.SupplierID = CGZL.ZSBH
                        AND SN223_UnitWeight.MatID = c.CLBH
                  LEFT JOIN Setup_UnitWeight AS SN74A
                        ON  SN74A.FormID = 'SN74A'
                        AND SN74A.SupplierID = 'ZZZZ'
                        AND SN74A.MatID = c.CLBH
            WHERE  CONVERT(VARCHAR ,CGDate ,23) BETWEEN N'2025-01-01' AND N'2025-12-31'
                  AND ISNULL(cgzl.CGLX ,'') NOT IN ('6' ,'4')
                  AND (
                        (LEFT(c.CLBH ,1) NOT IN ('X' ,'Y' ,'Z' ,'V'))
                        OR (LEFT(c.CLBH ,4)='V501')
                  )
                  AND ISNULL(c.Qty ,0)<>0;

            SELECT ROW_NUMBER() OVER(ORDER BY pnc.PurDate ,pnc.PurNo) AS [No]
                  ,COUNT(*) OVER()            AS TotalRowsCount
                  ,N'${getFactory(factory)}'  AS FactoryCode
                  ,pnc.*
                  ,ZLCLSL.CLSL                AS QtyUsage
                  ,kcrk.ModifyDate            AS RKDate
                  ,KCRK.Qty                   AS QtyReceive
                  ,kcrk.RKNO                  AS ReceivedNo
                  ,(pnc.UnitWeight*KCRK.Qty)  AS WeightUnitkg
                  ,CAST('0' AS INT)           AS LandTransportTonKilometers
                  ,CAST('0' AS INT)           AS SeaTransportTonKilometers
                  ,CAST('0' AS INT)           AS AirTransportTonKilometers
            FROM   #PurN233_CGZL pnc
                  INNER JOIN (
                        SELECT kcrk.RKNO
                              ,kcrk.ZSNO
                              ,kcrks.CLBH
                              ,kcrk.ModifyDate
                              ,SUM(ISNULL(KCRKS.Qty ,0)) Qty
                        FROM   kcrk
                              INNER JOIN kcrks
                                    ON  kcrks.RKNO = kcrk.RKNO
                                          AND ISNULL(KCRKS.RKSB ,'')<>'DL'
                                          AND ISNULL(KCRKS.RKSB ,'')<>'NG'
                        GROUP BY
                              kcrk.RKNO
                              ,kcrk.ZSNO
                              ,kcrks.CLBH
                              ,kcrk.ModifyDate
                        HAVING SUM(ISNULL(KCRKS.Qty ,0))>0
                        )KCRK
                        ON  kcrk.ZSNO = pnc.PurNo
                        AND kcrk.CLBH = pnc.MatID
                  LEFT JOIN (
                        SELECT SS.CGNO
                              ,S2.CLBH
                              ,ISNULL(SUM(S2.CLSL) ,0) AS CLSL
                        FROM   ZLZLS2 S2
                              LEFT JOIN CGZLSS SS
                                    ON  SS.ZLBH = S2.ZLBH
                                          AND SS.CLBH = S2.CLBH
                        GROUP BY
                              SS.CGNO
                              ,S2.CLBH
                        )ZLCLSL
                        ON  ZLCLSL.CGNO = pnc.PurNo
                        AND ZLCLSL.CLBH = pnc.MatID
            ORDER BY
                ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
            ${pagingSql}`;
  return query;
};

export const getADataExcelFactoryCat1AndCat4 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
  dbEIP?: Sequelize,
) => {
  // const { query } = await buildQuery(dateFrom, dateTo, factory, dbEIP);
  // const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

  // const data = await db.query(query, {
  //   replacements,
  //   type: QueryTypes.SELECT,
  // });

  // sheet.columns = [
  //   {
  //     header: 'No.',
  //     key: 'No',
  //   },
  //   {
  //     header: 'Pur Date',
  //     key: 'PurDate',
  //   },
  //   {
  //     header: 'RK Date',
  //     key: 'RKDate',
  //   },
  //   {
  //     header: 'Purchase Order',
  //     key: 'PurchaseOrder',
  //   },
  //   {
  //     header: 'Received No.',
  //     key: 'ReceivedNo',
  //   },
  //   {
  //     header: 'Material No',
  //     key: 'MaterialNo',
  //   },
  //   {
  //     header: 'Qty.(Usage)',
  //     key: 'QtyUsage',
  //   },
  //   {
  //     header: 'Qty.(Receive)',
  //     key: 'QtyReceive',
  //   },
  //   {
  //     header: 'Unit Weight',
  //     key: 'UnitWeight',
  //   },
  //   {
  //     header: 'Weight(Unit: KG)',
  //     key: 'Weight',
  //   },
  //   {
  //     header: 'Supplier Code',
  //     key: 'SupplierCode',
  //   },
  //   {
  //     header: 'Factory Code',
  //     key: 'FactoryCode',
  //   },
  //   {
  //     header: 'Style',
  //     key: 'Style',
  //   },
  //   {
  //     header: 'Transportation Method',
  //     key: 'TransportationMethod',
  //   },
  //   {
  //     header: 'Departure',
  //     key: 'Departure',
  //   },
  //   {
  //     header: 'Third-country Land Transport (A)',
  //     key: 'ThirdcountryLandTransportA',
  //   },
  //   {
  //     header: 'Port of Departure',
  //     key: 'PortofDeparture',
  //   },
  //   {
  //     header: 'Port of Arrival',
  //     key: 'PortofArrival',
  //   },
  //   {
  //     header: 'Factory (Domestic Land Transport) (B)',
  //     key: 'FactoryDomesticLandTransportB',
  //   },
  //   {
  //     header: 'Destination',
  //     key: 'Destination',
  //   },
  //   {
  //     header: 'Land Transport Distance (A+B)',
  //     key: 'LandTransportDistanceAB',
  //   },
  //   {
  //     header: 'Sea Transport Distance',
  //     key: 'SeaTransportDistance',
  //   },
  //   {
  //     header: 'Air Transport Distance',
  //     key: 'AirTransportDistance',
  //   },
  //   {
  //     header: 'Land Transport Ton-Kilometers',
  //     key: 'LandTransortTonKilometers',
  //   },
  //   {
  //     header: 'Sea Transport Ton-Kilometers',
  //     key: 'SeaTransortTonKilometers',
  //   },
  //   {
  //     header: 'Air Transport Ton-Kilometers',
  //     key: 'AirTransortTonKilometers',
  //   },
  // ];

  const query = await buildQueryTest('No', 'asc', factory, dbEIP, true);
  const replacements = {
    startDate: dateFrom,
    endDate: dateTo,
  };

  const data: any[] = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  sheet.columns = [
    { header: 'No.', key: 'No' },
    { header: 'Factory Code', key: 'FactoryCode' },
    { header: 'Pur Date', key: 'PurDate' },
    { header: 'RK Date', key: 'RKDate' },
    { header: 'Purchase Order', key: 'PurNo' },
    { header: 'Received No.', key: 'ReceivedNo' },
    { header: 'Material No.', key: 'MatID' },
    { header: 'Qty.(Usage)', key: 'QtyUsage' },
    { header: 'Qty.(receive)', key: 'QtyReceive' },
    { header: 'Unit weight', key: 'UnitWeight' },
    { header: 'Weight (Unit：KG)', key: 'WeightUnitkg' },
    { header: 'Supplier Code', key: 'SupplierCode' },
    { header: 'Style', key: 'Style' },
    { header: 'Transportation Method', key: 'TransportationMethod' },
    { header: 'Departure', key: 'Departure' },
    {
      header: 'Third-country Land Transport (A)',
      key: 'ThirdCountryLandTransport',
    },
    { header: 'Port of Departure', key: 'PortOfDeparture' },
    { header: 'Port of Arrival', key: 'PortOfArrival' },
    {
      header: 'Factory (Domestic Land Transport)(B)',
      key: 'FactoryDomesticLandTransport',
    },
    { header: 'Destination', key: 'Destination' },
    {
      header: 'Land Transport Distance (A+B)',
      key: 'LandTransportDistance',
    },
    { header: 'Sea Transport Distance', key: 'SeaTransportDistance' },
    { header: 'Air Transport Distance', key: 'AirTransportDistance' },
    {
      header: 'Land Transport Ton-Kilometer',
      key: 'LandTransportTonKilometers',
    },
    {
      header: 'Sea Transport Ton-Kilometers',
      key: 'SeaTransportTonKilometers',
    },
    {
      header: '	Air Transport Ton-Kilomete',
      key: 'AirTransportTonKilometers',
    },
  ];

  data.forEach((item) => sheet.addRow(item));

  sheet.columns.forEach((column) => {
    let maxLength = 0;
    if (typeof column.eachCell === 'function') {
      column.eachCell({ includeEmpty: true }, (cell) => {
        const cellValue = cell.value ? String(cell.value) : '';
        maxLength = Math.max(maxLength, cellValue.length);
      });
    }
    column.width = maxLength * 1.2;
  });

  sheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
  });
};


export const buildQueryAutoSentCMS = async (
  factory: string,
  db?: Sequelize,
) => {
  const queryAddress = `SELECT [Address]
                        FROM CMW_Info_Factory
                        WHERE CreatedFactory = '${factory}'`;

  const factoryAddress =
    (await db?.query(queryAddress, {
      type: QueryTypes.SELECT,
    })) || [];

  const query = `IF OBJECT_ID('tempdb..#PurN233_CGZL') IS NOT NULL
                  DROP TABLE #PurN233_CGZL

                  SELECT cgzl.CGDate                  AS PurDate
                        ,c.CGNO                       AS PurNo
                        ,c.CLBH                       AS MatID
                        ,clzl.ywpm                    AS MatName
                        ,ISNULL(
                        ISNULL(SN223_UnitWeight.UnitWeight ,imw.Total_Weight)
                        ,SN74A.UnitWeight
                        )                            AS UnitWeight
                        ,ISNULL(z.ZSDH ,CGZL.ZSBH)    AS SupplierCode
                        ,ISNULL(P.Style ,ZSZL.Style)  AS Style
                        ,CASE 
                              WHEN ISNULL(ZSZL.Country ,'')='' THEN NULL
                              WHEN (
                                    '${factory}' IN ('LYV' ,'LVL' ,'LHG')
                                    AND ZSZL.Country IN ('Vietnam' ,'Viet nam' ,'VN' ,' VIETNAM' ,'Viet Nam')
                              ) 
                              OR ('${factory}'='LYF' AND ZSZL.Country='Indonesia') 
                              OR (
                                    '${factory}' IN ('LYM' ,'POL')
                                    AND ZSZL.Country IN ('MYANMAR' ,' DA JIA MYANMAR COMPANY LIMITED' ,'MY')
                              ) THEN 'Land'
                              ELSE 'SEA + Land'
                        END                          AS TransportationMethod
                        ,isi.SupplierFullAddress      AS Departure
                        ,CASE 
                              WHEN ISNULL(isi.ThirdCountryLandTransport ,'')='' THEN 'N/A'
                              ELSE CAST(isi.ThirdCountryLandTransport AS VARCHAR)
                        END                          AS ThirdCountryLandTransport
                        ,pc.PortCode                  AS PortOfDeparture
                        --,'${factory === 'LYM' ? 'MMRGN' : 'VNCLP'}' AS PortOfArrival
                        ,'${factory}' AS PortOfArrival
                        ,isi.Factory_Port             AS FactoryDomesticLandTransport
                        ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Destination
                        ,ISNULL(isi.ThirdCountryLandTransport ,0)+ISNULL(isi.Factory_Port ,0) AS LandTransportDistance
                        ,isi.SeaTransportDistance     AS SeaTransportDistance
                        ,isi.AirTransportDistance     AS AirTransportDistance
                  INTO   #PurN233_CGZL
                  FROM   CGZLS AS c
                        LEFT JOIN cgzl
                              ON  cgzl.CGNO = c.CGNO
                        LEFT JOIN ZSZL
                              ON  CGZL.ZSBH = ZSZL.ZSDH
                        LEFT JOIN ZSZL_Prod P
                              ON  P.ZSDH = ZSZL.zsdh
                              AND P.GSBH = cgzl.GSBH
                        LEFT JOIN ZSZL z
                              ON  z.zsdh = ISNULL(P.MZSDH ,ZSZL.MZSDH)
                        LEFT JOIN Imp_SuppIDCombine AS isi
                              ON  isi.ZSDH = ZSZL.zsdh
                        LEFT JOIN clzl
                              ON  cldh = c.CLBH
                        LEFT JOIN (
                              SELECT smi2.CLBH
                                    ,zszl.zsdh
                                    ,MAX(smi2.Supplier_Material_ID) Supplier_Material_ID
                              FROM   SuppMatID AS smi2
                                    LEFT JOIN zszl
                                          ON  zszl.Zsdh_TW = smi2.CSBH
                                    INNER JOIN Imp_MaterialWeight
                                          ON  Imp_MaterialWeight.Supplier_Material_ID = REPLACE(
                                                      REPLACE(smi2.Supplier_Material_ID ,CHAR(10) ,'')
                                                ,CHAR(13)
                                                ,''
                                                )
                              WHERE  ISNULL(Total_Weight ,0)<>0
                              GROUP BY
                                    smi2.CLBH
                                    ,zszl.zsdh
                              ) A
                              ON  A.CLBH = c.CLBH
                              AND A.zsdh = CGZL.ZSBH
                        LEFT JOIN Imp_MaterialWeight imw
                              ON  imw.Supplier_Material_ID = A.Supplier_Material_ID
                        LEFT JOIN Setup_UnitWeight AS SN223_UnitWeight
                              ON  SN223_UnitWeight.FormID = 'SN223'
                              AND SN223_UnitWeight.SupplierID = CGZL.ZSBH
                              AND SN223_UnitWeight.MatID = c.CLBH
                        LEFT JOIN Setup_UnitWeight AS SN74A
                              ON  SN74A.FormID = 'SN74A'
                              AND SN74A.SupplierID = 'ZZZZ'
                              AND SN74A.MatID = c.CLBH
                        LEFT JOIN (
                              SELECT SupplierID
                                    ,PortCode
                                    ,TransportMethod
                              FROM   CMW.CMW.dbo.CMW_PortCode_Cat1_4
                              WHERE  FactoryCode = '${factory}'
                              ) AS pc
                              ON  pc.SupplierID = ISNULL(z.ZSDH ,CGZL.ZSBH) COLLATE Database_Default
                  WHERE  CONVERT(VARCHAR ,CGDate ,23) BETWEEN :startDate AND :endDate
                        AND ISNULL(cgzl.CGLX ,'')<>'6' AND ISNULL(c.Qty ,0)<>0;

                  SELECT ROW_NUMBER() OVER(ORDER BY pnc.PurDate ,pnc.PurNo) AS [No]
                        --,COUNT(*) OVER()            AS TotalRowsCount
                        ,N'${getFactory(factory)}'  AS FactoryCode
                        ,pnc.*
                        ,ZLCLSL.CLSL                AS QtyUsage
                        ,kcrk.ModifyDate            AS RKDate
                        ,KCRK.Qty                   AS QtyReceive
                        ,kcrk.RKNO                  AS ReceivedNo
                        ,(pnc.UnitWeight*KCRK.Qty)  AS WeightUnitkg
                  FROM   #PurN233_CGZL pnc
                        INNER JOIN (
                              SELECT kcrk.RKNO
                                    ,kcrk.ZSNO
                                    ,kcrks.CLBH
                                    ,kcrk.ModifyDate
                                    ,SUM(ISNULL(KCRKS.Qty ,0)) Qty
                              FROM   kcrk
                                    INNER JOIN kcrks
                                          ON  kcrks.RKNO = kcrk.RKNO
                                                AND ISNULL(KCRKS.RKSB ,'')<>'DL'
                                                AND ISNULL(KCRKS.RKSB ,'')<>'NG'
                              GROUP BY
                                    kcrk.RKNO
                                    ,kcrk.ZSNO
                                    ,kcrks.CLBH
                                    ,kcrk.ModifyDate
                              HAVING SUM(ISNULL(KCRKS.Qty ,0)) > 0
                              )KCRK
                              ON  kcrk.ZSNO = pnc.PurNo
                              AND kcrk.CLBH = pnc.MatID
                        LEFT JOIN (
                              SELECT SS.CGNO
                                    ,S2.CLBH
                                    ,ISNULL(SUM(S2.CLSL) ,0) AS CLSL
                              FROM   ZLZLS2 S2
                                    LEFT JOIN CGZLSS SS
                                          ON  SS.ZLBH = S2.ZLBH
                                                AND SS.CLBH = S2.CLBH
                              GROUP BY
                                    SS.CGNO
                                    ,S2.CLBH
                              )ZLCLSL
                              ON  ZLCLSL.CGNO = pnc.PurNo
                              AND ZLCLSL.CLBH = pnc.MatID`;
  return query;
};

