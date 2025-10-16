import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

export const getDBFactory = async () => {};

export const getADataExcelFactoryCat5 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
) => {
  let where = 'WHERE 1=1';
  const replacements: any[] = [];

  if (dateTo && dateFrom) {
    where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo);
  }

  const query = `SELECT dwo.WASTE_DATE AS Waste_disposal_date
                        ,dtv.TREATMENT_VENDOR_NAME          AS Vender_Name
                        ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
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
  const data = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  sheet.columns = [
    {
      header: 'Waste disposal date',
      key: 'Waste_disposal_date',
    },
    {
      header: 'Vender Name',
      key: 'Vender_Name',
    },
    {
      header: 'Waste collection address',
      key: 'Waste_collection_address',
    },
    {
      header: 'Transportation Distance (km)',
      key: 'Transportation_Distance_km',
    },
    {
      header: '*The type of waste',
      key: 'The_type_of_waste',
    },
    {
      header: '*Waste type',
      key: 'Waste_type',
    },
    {
      header: '*Waste Treatment method',
      key: 'Waste_Treatment_method',
    },
    {
      header: '*Weight of waste treated (Unit：kg)',
      key: 'Weight_of_waste_treated_Unit_kg',
    },
    {
      header: 'TKT (Ton-km)',
      key: 'TKT_Ton_km',
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

export const getAllDataExcelFactoryCat5 = async () => {};
