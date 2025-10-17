import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

export const getADataExcelFactoryCat7 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
) => {
  let where = "WHERE 1=1 AND Work_Or_Not<>'2'";
  const replacements: any[] = [];

  if (dateTo && dateFrom) {
    where += ` AND CHECK_DAY BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo);
  }

  const query = `SELECT DP.Person_ID AS Staff_ID
                          ,DP.Staying_Address            AS Residential_address
                          ,DP.Vehicle                    AS Main_transportation_type
                          ,CAST('0' AS INT)              AS km
                          ,ROUND(SUM(WORKING_TIME) ,2)   AS Number_of_working_days
                          ,CAST('0' AS INT)              AS PKT_p_km
                    FROM   DATA_WORK_TIME WT
                          LEFT JOIN Data_Person_Detail  AS DP
                                ON  WT.Person_Serial_Key = DP.Person_Serial_Key
                    ${where}
                    GROUP BY
                          DP.Person_ID
                          ,DP.Staying_Address
                          ,DP.Vehicle;`;
  const data = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });
  sheet.columns = [
    { header: 'Staff ID', key: 'Staff_ID' },
    { header: 'Residential address', key: 'Residential_address' },
    { header: 'Main transportation type', key: 'Main_transportation_type' },
    { header: 'km', key: 'km' },
    { header: 'Number of working days', key: 'Number_of_working_days' },
    { header: 'PKT (p-km)', key: 'PKT_p_km' },
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
