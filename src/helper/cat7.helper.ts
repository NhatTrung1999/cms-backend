import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

export const buildQuery = (
  factory: string,
  dateFrom?: string,
  dateTo?: string,
) => {
  const isLYM = factory === 'LYM';

  const baseWhere = !isLYM
    ? "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0"
    : 'WHERE 1=1 AND u.Vehicle IS NOT NULL AND dwt.workhours > 0';

  const dateFilter = !isLYM
    ? 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?'
    : 'AND CONVERT(DATE ,dwt.CDate) BETWEEN ? AND ?';

  const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

  const workCol = !isLYM ? 'WORKING_TIME' : 'workhours';

  const table = !isLYM ? 'Data_Work_Time' : 'HR_Attendance';

  const join = !isLYM
    ? 'u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS'
    : 'u.userId = dwt.UserNo';

  const query = `SELECT u.userId                       AS Staff_ID
                        ,CASE 
                              WHEN u.Address_Live IS NULL THEN u.Bus_Route
                              ELSE u.Address_Live
                        END                            AS Residential_address
                        ,u.Vehicle                      AS Main_transportation_type
                        ,'API Calculation'              AS km
                        ,COUNT(${workCol})  AS Number_of_working_days
                        ,'API Calculation'              AS PKT_p_km
                  FROM   ${table}                 AS dwt
                        LEFT JOIN users                AS u
                              ON  ${join}
                  ${where}
                  GROUP BY
                          u.userId
                          ,u.Address_Live
                          ,u.Vehicle
                          ,u.Bus_Route`;

  // console.log(query);
  const countQuery = `SELECT COUNT(*) AS total
                      FROM   (
                                SELECT u.userId                       AS Staff_ID
                                      ,CASE 
                                            WHEN u.Address_Live IS NULL THEN u.Bus_Route
                                            ELSE u.Address_Live
                                      END                            AS Residential_address
                                      ,u.Vehicle                      AS Main_transportation_type
                                      ,'API Calculation'              AS km
                                      ,COUNT(${workCol})  AS Number_of_working_days
                                      ,'API Calculation'              AS PKT_p_km
                                FROM   ${table}                 AS dwt
                                      LEFT JOIN users                AS u
                                            ON  ${join}
                                ${where}
                                GROUP BY
                                        u.userId
                                        ,u.Address_Live
                                        ,u.Vehicle
                                        ,u.Bus_Route	
                      ) AS Sub`;
  return { query, countQuery };
};

export const getADataExcelFactoryCat7 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  const { query } = buildQuery(factory, dateFrom, dateTo);
  const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

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
