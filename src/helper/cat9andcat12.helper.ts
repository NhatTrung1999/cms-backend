import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { getFactory } from './factory.helper';

export const getADataExcelFactoryCat9AndCat12 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
  dbEIP?: Sequelize,
) => {
  let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
  const replacements: any[] = [];

  if (dateFrom && dateTo) {
    where += ` AND CONVERT(VARCHAR ,sb.ExFty_Date,23) BETWEEN ? AND ?`;
    where1 += ` AND CONVERT(VARCHAR ,sb.ExFty_Date,23) BETWEEN ? AND ?`;
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
  let data: any = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  data = await Promise.all(
    data.map(async (item) => {
      let portOfDeparture = '';
      if (
        item?.Article_ID.trim().toLowerCase() ===
        'SAMPLE SHOE'.trim().toLowerCase()
      ) {
        item.Transport_Method = 'AIR';
        portOfDeparture = 'SGN';
      } else {
        switch (item?.Transport_Method.trim().toLowerCase()) {
          case '11 SC'.trim().toLowerCase():
          case 'Ocean'.trim().toLowerCase():
          case 'Sea Air Collect'.trim().toLowerCase():
          case 'SC'.trim().toLowerCase():
          case 'Sea/Air'.trim().toLowerCase():
          case 'Partial'.trim().toLowerCase():
            item.Transport_Method = 'SEA';
            break;
          case '20 CC'.trim().toLowerCase():
          case 'AC'.trim().toLowerCase():
          case '10 AC'.trim().toLowerCase():
          case 'CC'.trim().toLowerCase():
          case 'Air PP'.trim().toLowerCase():
          case 'Air Express'.trim().toLowerCase():
          case 'AE'.trim().toLowerCase():
            item.Transport_Method = 'AIR';
            break;
          case 'RAIL'.trim().toLowerCase():
          case 'Rail Collect'.trim().toLowerCase():
          case 'Rail Colle'.trim().toLowerCase():
            item.Transport_Method = 'Land';
            break;
          default:
            item.Transport_Method = 'SEA';
            break;
        }

        switch (factory.trim().toLowerCase()) {
          case 'LYV'.trim().toLowerCase():
          case 'LVL'.trim().toLowerCase():
          case 'LHG'.trim().toLowerCase():
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'SEA'.trim().toLowerCase()
            ) {
              portOfDeparture = 'VNCLP';
            }
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'Land'.trim().toLowerCase()
            ) {
              portOfDeparture = 'SOTH';
            }
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'AIR'.trim().toLowerCase()
            ) {
              portOfDeparture = 'SGN';
            }
            break;
          case 'LYM'.trim().toLowerCase():
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'SEA'.trim().toLowerCase()
            ) {
              portOfDeparture = 'MMRGN';
            } else {
              portOfDeparture = 'RGN';
            }
            break;
          case 'LYF'.trim().toLowerCase():
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'SEA'.trim().toLowerCase()
            ) {
              portOfDeparture = 'IDSRG';
            } else {
              portOfDeparture = 'CGK';
            }
            break;
          default:
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'SEA'.trim().toLowerCase()
            ) {
              portOfDeparture = 'VNCLP';
            }
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'Land'.trim().toLowerCase()
            ) {
              portOfDeparture = 'SOTH';
            }
            if (
              item.Transport_Method.trim().toLowerCase() ===
              'AIR'.trim().toLowerCase()
            ) {
              portOfDeparture = 'SGN';
            }
            break;
            break;
        }
      }
      const queryPoA = `SELECT PortCode
                        FROM CMW_PortCode
                        WHERE TransportMethod LIKE ? AND CustomerNumber LIKE ?`;
      const result: any = await dbEIP?.query(queryPoA, {
        type: QueryTypes.SELECT,
        replacements: [`%${item.Transport_Method}%`, `%${item.Customer_ID}%`],
      });
      return {
        ...item,
        Transport_Method: item.Transport_Method,
        Port_Of_Departure: portOfDeparture,
        Port_Of_Arrival: result[0]?.PortCode.trim().toUpperCase() || 'N/A',
      };
    }),
  );

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
  db?: Sequelize,
) => {
  const queryAddress = `SELECT ${factory?.trim().toLowerCase() === 'lym'.trim().toLowerCase() ? `N'TSANG YIH Co., Ltd /Polo/ချန်းရိ' as [Address]` : '[Address]'}
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
    where += ` AND CONVERT(VARCHAR ,sb.ExFty_Date,23) BETWEEN ? AND ?`;
    where1 += ` AND CONVERT(VARCHAR ,sb.ExFty_Date,23) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo, dateFrom, dateTo);
  }
  const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          , N'${getFactory(factory)}' AS Factory
                          , N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
                          ,'${factory}' AS Factory_Name
                  FROM   (
                            SELECT im.INV_DATE             AS [Date]
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
  return query;
};
