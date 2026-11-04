import { getFactory } from './factory.helper';

/** `AND CGZL.ZSBH ...` clause that leaves out the internal suppliers of a factory ('' for the others). */
export const getSupplierExclusion = (factory: string): string => {
  switch (factory.trim().toLowerCase()) {
    case 'lym':
      return `AND CGZL.ZSBH not in ('MT520', 'VT35')`;
    case 'lhg':
      return `AND CGZL.ZSBH <> 'VL889'`;
    case 'lyv':
      return `AND CGZL.ZSBH <> 'VL888'`;
    case 'lvl':
      return `AND CGZL.ZSBH <> 'VT205'`;
    default:
      return '';
  }
};

/**
 * Shared part of the Cat1&4 SQL: builds the temp tables and ends with
 * `SELECT ... FROM #Cat1AndCat4`. The page query appends its own filter / order / paging,
 * the CMS query uses it as it is.
 */
export const buildCat1AndCat4Base = (
  factory: string,
  factoryAddress: string,
  notIn: string,
) => `
            IF OBJECT_ID('tempdb..#RKList') IS NOT NULL
            DROP TABLE #RKList;
            SELECT ZSNO
                  ,RKNO
                  ,ModifyDate
            INTO   #RKList
            FROM   KCRK
            WHERE  KCRK.ModifyDate >= :startDate AND KCRK.ModifyDate < :endDate
            GROUP BY
                  ZSNO
                  ,RKNO
                  ,ModifyDate

            IF OBJECT_ID('tempdb..#PurN233_CGZL') IS NOT NULL
            DROP TABLE #PurN233_CGZL;            

            SELECT cgzl.CGDate                  AS PurDate
                  ,c.CGNO                       AS PurNo
                  ,c.CLBH                       AS MatID
                  ,clzl.ywpm                    AS MatName
                  ,CASE 
                        WHEN c.CLBH LIKE 'U1AD%' AND ISNULL(y.OWeigh ,0)<>0 THEN y.OWeigh
                        ELSE ISNULL(
                              ISNULL(SN223.UnitWeight ,SN74A.UnitWeight)
                              ,imw.Total_Weight
                        )
                  END                          AS UnitWeight
                  --,ISNULL(z.ZSDH ,CGZL.ZSBH)    AS SupplierCode
                  ,CGZL.ZSBH SupplierCode
                  --,ISNULL(P.Style ,ZSZL.Style)  AS Style
                  ,CASE 
                        WHEN ISNULL(poisi.Transportationmethod ,'')<>'' THEN poisi.Transportationmethod
                        WHEN ISNULL(isi.Transportationmethod ,'')<>'' AND ISNULL(poisi.Transportationmethod ,'')='' THEN 
                        isi.Transportationmethod
                        ELSE (
                              CASE 
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
                                    END
                        )
                  END                          AS TransportationMethod
                  ,CASE 
                        WHEN ISNULL(poisi.SupplierFullAddress ,'')<>'' THEN poisi.SupplierFullAddress
                        ELSE isi.SupplierFullAddress
                  END                             Departure
                  ,CASE 
                        WHEN ISNULL(poisi.ThirdCountryLandTransport ,0)<>0 THEN CAST(poisi.ThirdCountryLandTransport AS VARCHAR)
                        ELSE CASE 
                              WHEN ISNULL(isi.ThirdCountryLandTransport ,'')='' THEN 'N/A'
                              ELSE CAST(isi.ThirdCountryLandTransport AS VARCHAR)
                        END
                  END                          AS ThirdCountryLandTransport
                  --,ISNULL(poisi.PortOfDeparture ,isi.PortOfDeparture)          AS PortOfDeparture
                  --,ISNULL(poisi.PortOfArrival,isi.PortOfArrival)            AS PortOfArrival
                  ,ISNULL(poisi.Factory_Port ,isi.Factory_Port) AS FactoryDomesticLandTransport
                  ,N'${factoryAddress}' AS Destination
                  ,CASE 
                        WHEN ISNULL(poisi.ThirdCountryLandTransport ,0)<>0 OR ISNULL(poisi.Factory_Port ,0)<>0 THEN ISNULL(poisi.ThirdCountryLandTransport ,0)
                        +ISNULL(poisi.Factory_Port ,0)
                        ELSE ISNULL(isi.ThirdCountryLandTransport ,0)+ISNULL(isi.Factory_Port ,0)
                  END                             LandTransportDistance
                  ,ISNULL(poisi.SeaTransportDistance ,isi.SeaTransportDistance) AS SeaTransportDistance
                  ,ISNULL(poisi.AirTransportDistance ,isi.AirTransportDistance) AS AirTransportDistance
                  ,clzl.dwbh
                  ,c.Qty
                  --,KCRK.ModifyDate RKDate
            INTO   #PurN233_CGZL
            FROM   CGZLS AS c
                  LEFT JOIN cgzl
                        ON  cgzl.CGNO = c.CGNO
                  INNER JOIN #RKList AS r
                        ON  r.ZSNO = cgzl.CGNO 
                        -- LEFT JOIN KCRK
                        --      ON  KCRK.ZSNO = cgzl.CGNO
                        
                  LEFT JOIN ZSZL
                        ON  CGZL.ZSBH = ZSZL.ZSDH
                  LEFT JOIN ZSZL_Prod P
                        ON  P.ZSDH = ZSZL.zsdh
                        AND P.GSBH = cgzl.GSBH
                  LEFT JOIN ZSZL z
                        ON  z.zsdh = ISNULL(P.MZSDH ,ZSZL.MZSDH)
                  LEFT JOIN Imp_SuppIDCombine AS isi
                        ON  isi.ZSDH = z.ZSDH
                  LEFT JOIN Imp_SuppIDCombine AS poisi
                        ON  poisi.ZSDH = CGZL.ZSBH
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
                              INNER JOIN Imp_MaterialWeight imw
                                    ON  imw.Supplier_Material_ID = REPLACE(
                                                REPLACE(smi2.Supplier_Material_ID ,CHAR(10) ,'')
                                          ,CHAR(13)
                                          ,''
                                          )
                        WHERE  ISNULL(imw.Total_Weight ,0)<>0
                              AND (smi2.CLBH NOT LIKE '[XYZV]%' OR smi2.CLBH LIKE 'V501%')
                        GROUP BY
                              smi2.CLBH
                              ,zszl.zsdh
                        ) A
                        ON  A.CLBH = c.CLBH
                        AND A.zsdh = CGZL.ZSBH
                  LEFT JOIN Imp_MaterialWeight imw
                        ON  imw.Supplier_Material_ID = A.Supplier_Material_ID
                  LEFT JOIN Setup_UnitWeight AS SN223
                        ON  SN223.FormID = 'SN223'
                        AND SN223.SupplierID = CGZL.ZSBH
                        AND SN223.MatID = c.CLBH
                  LEFT JOIN Setup_UnitWeight AS SN74A
                        ON  SN74A.FormID = 'SN74A'
                        AND SN74A.SupplierID = 'ZZZZ'
                        AND SN74A.MatID = c.CLBH
            WHERE  1 = 1 
                  ${notIn}
                  AND ISNULL(cgzl.CGLX ,'') NOT IN ('6' ,'4')
                  AND (c.CLBH NOT LIKE '[XYZV]%' OR c.CLBH LIKE 'V501%')
                  AND LEFT(c.CLBH,1) NOT IN ('W', 'T')
                  AND (
                        NOT EXISTS (
                              SELECT 1
                              FROM   Setup_Exclusion_Cases_MaterialID AS secmi
                              WHERE  LEFT(c.CLBH ,4) = secmi.TypeMaterial
                                    AND secmi.Fty = 'ALL'
                                    AND secmi.CLBH = 'ALL'
                        )
                        OR EXISTS (
                              SELECT 1
                              FROM   Setup_Exclusion_Cases_MaterialID AS secmi
                              WHERE  LEFT(c.CLBH ,4) = secmi.TypeMaterial
                                    AND secmi.Fty = CGZL.GSBH
                                    AND secmi.CLBH = c.CLBH
                              )
                  )
                  AND ISNULL(c.Qty ,0)<>0;
            CREATE CLUSTERED INDEX IX_PurN233_CGZL_Key ON #PurN233_CGZL(PurNo ,MatID);
            CREATE NONCLUSTERED INDEX IX_RKList_Key ON #RKList(ZSNO ,RKNO ,ModifyDate);

            IF OBJECT_ID('tempdb..#RKQTy') IS NOT NULL
            DROP TABLE #RKQTy;
            SELECT k.ZSNO,ks.CLBH,k.RKNO,k.ModifyDate,SUM(ISNULL(ks.Qty ,0)) AS Qty INTO #RKQTy
            FROM   #RKList k
                  INNER JOIN kcrks ks
                        ON  k.RKNO = ks.RKNO
                  INNER JOIN #PurN233_CGZL t
                        ON  t.PurNo = k.ZSNO
                        AND t.MatID = ks.CLBH
            WHERE  ISNULL(ks.RKSB ,'') NOT IN ('DL' ,'NG')
            GROUP BY
                  k.ZSNO
                  ,ks.CLBH
                  ,k.RKNO
                  ,k.ModifyDate
            HAVING SUM(ISNULL(ks.Qty ,0))>0
                        
            --IF OBJECT_ID('tempdb..#ZLCLSL') IS NOT NULL
            --DROP TABLE #ZLCLSL;
            --SELECT SS.CGNO,S2.CLBH,SUM(ISNULL(S2.CLSL ,0)) AS CLSL INTO #ZLCLSL
            --FROM   #PurN233_CGZL t
            --        INNER JOIN CGZLSS SS
            --            ON  t.PurNo = SS.CGNO
            --        INNER JOIN ZLZLS2 S2
            --            ON  SS.ZLBH = S2.ZLBH
            --                AND SS.CLBH = S2.CLBH
            --WHERE  S2.CLBH = t.MatID
            --GROUP BY SS.CGNO,S2.CLBH

            IF OBJECT_ID('tempdb..#Cat1AndCat4') IS NOT NULL
            DROP TABLE #Cat1AndCat4;
            SELECT N'${getFactory(factory)}'             AS FactoryCode
                  ,pnc.*
                  --,ISNULL(ZLCLSL.CLSL ,0)  AS QtyUsage
                  ,kcrk.ModifyDate         AS RKDate
                  ,KCRK.Qty                AS QtyReceive
                  ,kcrk.RKNO               AS ReceivedNo
                  ,CASE 
                        WHEN dwbh='KGS' THEN pnc.Qty
                        WHEN dwbh='GRAM' THEN pnc.Qty/1000
                        WHEN dwbh='MGR' THEN pnc.Qty/1000000
                        ELSE (pnc.UnitWeight*KCRK.Qty)
                  END                        WeightUnitkg
                  ,CAST('0' AS INT)        AS LandTransportTonKilometers
                  ,CAST('0' AS INT)        AS SeaTransportTonKilometers
                  ,CAST('0' AS INT)        AS AirTransportTonKilometers
                  ,pc.PortCode             AS PortOfDeparture
                  ,CASE 
                        WHEN '${factory}' IN ('LYV' ,'LVL' ,'LHG') AND pnc.TransportationMethod='SEA+LAND' THEN 'VNCLP'
                        WHEN '${factory}' IN ('LYM' ,'POL') AND pnc.TransportationMethod='SEA+LAND' THEN 'MMRGN'
                        WHEN '${factory}' IN ('LYF') AND pnc.TransportationMethod='SEA+LAND' THEN 'IDSRG'
                        ELSE NULL
                  END                     AS PortOfArrival
            INTO   #Cat1AndCat4
            FROM   #PurN233_CGZL pnc
            INNER JOIN #RKQTy KCRK ON KCRK.ZSNO = pnc.PurNo AND KCRK.CLBH = pnc.MatID
               /*   INNER JOIN (
                        SELECT k.ZSNO
                              ,ks.CLBH
                              ,k.RKNO
                              ,k.ModifyDate
                              ,SUM(ISNULL(ks.Qty ,0)) AS Qty
                        FROM   #RKList k
                              INNER JOIN kcrks ks
                                    ON  k.RKNO = ks.RKNO
                              INNER JOIN #PurN233_CGZL t
                                    ON  t.PurNo = k.ZSNO
                                          AND t.MatID = ks.CLBH
                        WHERE  ISNULL(ks.RKSB ,'') NOT IN ('DL' ,'NG')
                        GROUP BY
                              k.ZSNO
                              ,ks.CLBH
                              ,k.RKNO
                              ,k.ModifyDate
                        HAVING SUM(ISNULL(ks.Qty ,0))>0
                        ) KCRK
                        ON  KCRK.ZSNO = pnc.PurNo
                        AND KCRK.CLBH = pnc.MatID
                  LEFT JOIN (
                        SELECT SS.CGNO
                              ,S2.CLBH
                              ,SUM(ISNULL(S2.CLSL ,0)) AS CLSL
                        FROM   #PurN233_CGZL t
                              INNER JOIN CGZLSS SS
                                    ON  t.PurNo = SS.CGNO
                              INNER JOIN ZLZLS2 S2
                                    ON  SS.ZLBH = S2.ZLBH
                                          AND SS.CLBH = S2.CLBH
                        WHERE  S2.CLBH = t.MatID
                        GROUP BY
                              SS.CGNO
                              ,S2.CLBH
                        ) ZLCLSL
                        ON  ZLCLSL.CGNO = pnc.PurNo
                        AND ZLCLSL.CLBH = pnc.MatID */

                  LEFT JOIN (
                        SELECT SupplierID
                              ,PortCode
                        FROM   CMW.CMW.dbo.CMW_PortCode_Cat1_4
                        WHERE  FactoryCode = '${factory}'
                        ) AS pc
                        ON  pc.SupplierID = pnc.SupplierCode COLLATE Database_Default
            GROUP BY
                  pnc.PurDate
                  ,pnc.PurNo
                  ,pnc.MatID
                  ,pnc.MatName
                  ,pnc.UnitWeight
                  ,pnc.SupplierCode
                  --,pnc.Style
                  ,pnc.TransportationMethod
                  ,pnc.Departure
                  ,pnc.ThirdCountryLandTransport
                  ,pnc.FactoryDomesticLandTransport
                  ,pnc.Destination
                  ,pnc.LandTransportDistance
                  ,pnc.SeaTransportDistance
                  ,pnc.AirTransportDistance
                  ,pnc.dwbh
                  ,pnc.Qty
                 -- ,ISNULL(ZLCLSL.CLSL ,0)
                  ,kcrk.ModifyDate
                  ,KCRK.Qty
                  ,kcrk.RKNO
                  ,pc.PortCode;
            CREATE CLUSTERED INDEX IX_Cat1AndCat4_Sort ON #Cat1AndCat4(PurDate ,PurNo);

            SELECT * 
                  ,COUNT(*) OVER() AS TotalRowsCount
                  ,ROW_NUMBER() OVER(ORDER BY PurDate ,PurNo) AS [No]
            FROM   #Cat1AndCat4`;
