import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

export const getADataExcelFactoryCat1AndCat4 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
) => {
  let where = 'WHERE 1=1';
  const replacements: any[] = [];
  if (dateFrom && dateTo) {
    where += ` AND CONVERT(VARCHAR, c.USERDate, 23) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo);
  }

  const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
                          ,c.USERDate        AS [Date]
                          ,c.CGNO            AS Purchase_Order
                          ,c2.CLBH           AS Material_No
                          ,CAST('0' AS INT)     AS [Weight]
                          ,z.SupplierCode AS Supplier_Code
                          ,z.ThirdCountryLandTransport AS Thirdcountry_Land_Transport
                          ,z.PortofDeparture AS Port_Of_Departure
                          ,z.PortofArrival AS Port_Of_Arrival
                          ,z.Transportationmethod AS Factory_Domestic_Land_Transport
                          ,CAST('0' AS INT)     AS Land_Transport_Distance
                          ,z.SeaTransportDistance AS Sea_Transport_Distance
                          ,CAST('0' AS INT) AS Air_Transport_Distance
                          ,CAST('0' AS INT) AS Land_Transport_Ton_Kilometers
                          ,CAST('0' AS INT) AS Sea_Transport_Ton_Kilometers
                          ,CAST('0' AS INT) AS Air_Transport_Ton_Kilometers
                    FROM   CGZL              AS c
                          INNER JOIN CGZLS  AS c2
                                ON  c2.CGNO = c.CGNO
                          LEFT JOIN zszl    AS z
                                ON  z.zsdh = c.CGNO
                    ${where}`;
  const data = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  sheet.columns = [
    { header: 'No.', key: 'No' },
    { header: 'Date', key: 'Date' },
    { header: 'Purchase Order', key: 'Purchase_Order' },
    { header: 'Material No.', key: 'Material_No' },
    { header: 'Weight (Unit: Ton)', key: 'Weight' },
    { header: 'Supplier Code', key: 'Supplier_Code' },
    {
      header: 'Third-country Land Transport (A)',
      key: 'Thirdcountry_Land_Transport',
    },
    { header: 'Port of Departure', key: 'Port_Of_Departure' },
    { header: 'Port of Arrival', key: 'Port_Of_Arrival' },
    {
      header: 'Factory (Domestic Land Transport) (B)',
      key: 'Factory_Domestic_Land_Transport',
    },
    { header: 'Land Transport Distance (A+B)', key: 'Land_Transport_Distance' },
    { header: 'Sea Transport Distance', key: 'Sea_Transport_Distance' },
    { header: 'Air Transport Distance', key: 'Air_Transport_Distance' },
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
