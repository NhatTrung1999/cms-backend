import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { getFactory } from './factory.helper';

export const getADataExcelFactoryCat9AndCat12 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
) => {
  let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  const replacements: any[] = [];

  if (dateFrom && dateTo) {
    where += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
    where1 += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo, dateFrom, dateTo);
  }

  const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'SEA'                   AS Transport_Method
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
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
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
  const data = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });
  sheet.columns = [
    {
      header: 'No.',
      key: 'No',
    },
    {
      header: 'Date',
      key: 'Date',
    },
    {
      header: 'Shipment Date',
      key: 'Shipment_Date',
    },
    {
      header: 'Invoice Number',
      key: 'Invoice_Number',
    },
    {
      header: 'Article Name',
      key: 'Article_Name',
    },
    {
      header: 'Article ID',
      key: 'Article_ID',
    },
    {
      header: 'Quantity',
      key: 'Quantity',
    },
    {
      header: 'Gross Weight',
      key: 'Gross_Weight',
    },
    {
      header: 'Customer ID',
      key: 'Customer_ID',
    },
    {
      header: 'Local land transportation',
      key: 'Local_Land_Transportation',
    },
    {
      header: 'Port of Departure',
      key: 'Port_Of_Departure',
    },
    {
      header: 'Port of Arrival',
      key: 'Port_Of_Arrival',
    },
    {
      header: 'Land Transport Distance',
      key: 'Land_Transport_Distance',
    },
    {
      header: 'Sea Transport Distance',
      key: 'Sea_Transport_Distance',
    },
    {
      header: 'Air Transport Distance',
      key: 'Air_Transport_Distance',
    },
    {
      header: 'Transport Method',
      key: 'Transport_Method',
    },
    {
      header: 'Land Transport Ton-Kilometers',
      key: 'Land_Transport_Ton_Kilometers',
    },
    {
      header: 'Sea Transport Ton-Kilometers',
      key: 'Sea_Transport_Ton_Kilometers',
    },
    {
      header: 'Air Transport Ton-Kilometers',
      key: 'Air_Transport_Ton_Kilometers',
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
  dateFrom: string,
  dateTo: string,
  factory: string,
  db?:Sequelize
) => {

  const queryAddress = `SELECT [Address]
                        FROM CMW_Info_Factory
                        WHERE CreatedFactory = '${factory}'`;

  const factoryAddress =
    (await db?.query(queryAddress, {
      type: QueryTypes.SELECT,
    })) || [];

  let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  const replacements: any[] = [];

  if (dateFrom && dateTo) {
    where += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
    where1 += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo, dateFrom, dateTo);
  }

  const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          , N'${getFactory(factory)}' AS Factory
                          , N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
                    FROM   (
                                    SELECT CONVERT(VARCHAR(10), im.INV_DATE, 111)             AS [Date]
                                    ,CONVERT(VARCHAR(10), sb.ExFty_Date, 111)           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,im.INV_NO               AS Invoice_Number
                                    --,id.STYLE_NAME           AS Article_Name
                                    --,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'SEA'                   AS Transport_Method
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
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING
                                              GROUP BY
                                                    INV_NO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
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
                              SELECT CONVERT(VARCHAR(10), im.INV_DATE, 111)             AS [Date]
                                    ,CONVERT(VARCHAR(10), sb.ExFty_Date, 111)           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,is1.Inv_No              AS Invoice_Number
                                    --,'SAMPLE SHOE'           AS Article_Name
                                    --,id.ARTICLE              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
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
  return query;
};
