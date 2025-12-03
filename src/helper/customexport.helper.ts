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


