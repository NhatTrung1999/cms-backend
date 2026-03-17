import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { getFactory } from './factory.helper';
import dayjs from 'dayjs';
dayjs().format();

// export const buildQuery = (
//   factory: string,
//   dateFrom?: string,
//   dateTo?: string,
// ) => {
//   const isLYM = factory === 'LYM';

//   const baseWhere = !isLYM
//     ? "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.lock = '0'"
//     : `WHERE 1=1 AND u.Vehicle IS NOT NULL AND dwt.workhours > 0 AND u.lock = '0'`;

//   const dateFilter = !isLYM
//     ? 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?'
//     : 'AND CONVERT(DATE ,dwt.CDate) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const workCol = !isLYM ? 'WORKING_TIME' : 'workhours';

//   const table = !isLYM ? 'Data_Work_Time' : 'HR_Attendance';

//   const join = !isLYM
//     ? 'ISNULL(u.Person_Serial_Key ,u.userId) = dwt.Person_Serial_Key COLLATE Database_Default'
//     : 'u.userId = dwt.UserNo';

//   const query = `SELECT u.userId                       AS Staff_ID
//                         ,CASE
//                               WHEN u.Address_Live IS NULL THEN u.PickupDropoffStation
//                               ELSE u.Address_Live
//                         END                            AS Residential_address
//                         ,u.Vehicle                      AS Main_transportation_type
//                         ,'API Calculation'              AS km
//                         ,COUNT(${workCol})  AS Number_of_working_days
//                         ,'API Calculation'              AS PKT_p_km
//                   FROM   ${table}                 AS dwt
//                         LEFT JOIN users                AS u
//                               ON  ${join}
//                   ${where}
//                   GROUP BY
//                           u.userId
//                           ,u.Address_Live
//                           ,u.Vehicle
//                           ,u.PickupDropoffStation`;

//   // console.log(query);
//   const countQuery = `SELECT COUNT(*) AS total
//                       FROM   (
//                                 SELECT u.userId                       AS Staff_ID
//                                       ,CASE
//                                             WHEN u.Address_Live IS NULL THEN u.PickupDropoffStation
//                                             ELSE u.Address_Live
//                                       END                            AS Residential_address
//                                       ,u.Vehicle                      AS Main_transportation_type
//                                       ,'API Calculation'              AS km
//                                       ,COUNT(${workCol})  AS Number_of_working_days
//                                       ,'API Calculation'              AS PKT_p_km
//                                 FROM   ${table}                 AS dwt
//                                       LEFT JOIN users                AS u
//                                             ON  ${join}
//                                 ${where}
//                                 GROUP BY
//                                         u.userId
//                                         ,u.Address_Live
//                                         ,u.Vehicle
//                                         ,u.PickupDropoffStation
//                       ) AS Sub`;
//   return { query, countQuery };
// };

export const buildQuery = (
  factory: string,
  dateFrom?: string,
  dateTo?: string,
) => {
  const isLYM = factory === 'LYM';

  const workCol = isLYM ? 'workhours' : 'WORKING_TIME';
  const table = isLYM ? 'HR_Attendance' : 'Data_Work_Time';
  const dateCol = isLYM ? 'CDate' : 'Check_Day';
  const workFilter = isLYM
    ? 'workhours > 0'
    : "Work_Or_Not <> '2' AND Working_Time > 0";

  const cteGroupKey = isLYM ? 'UserNo' : 'Person_Serial_Key';

  const join = isLYM
    ? 'e.UserNo = u.userId'
    : 'e.Person_Serial_Key COLLATE DATABASE_DEFAULT = ISNULL(u.Person_Serial_Key, u.userId) COLLATE DATABASE_DEFAULT';

  const dateFilter =
    dateFrom && dateTo
      ? `AND CONVERT(DATE, ${dateCol}) >= CONVERT(DATE, ?) AND CONVERT(DATE, ${dateCol}) < CONVERT(DATE, ?)`
      : '';

  const cte = `
    WITH WorkTime AS (
      SELECT ${cteGroupKey}
            ,COUNT(${workCol}) AS Number_of_Working_Days
      FROM   ${table}
      WHERE  ${workFilter}
             ${dateFilter}
      GROUP BY ${cteGroupKey}
    )
  `;

  const selectBody = `
    SELECT u.userId                                                  AS Staff_ID
          ,CASE
                WHEN u.Address_Live IS NULL THEN u.PickupDropoffStation COLLATE DATABASE_DEFAULT
                ELSE u.Address_Live COLLATE DATABASE_DEFAULT
           END                                                       AS Residential_address
          ,u.Vehicle                                                 AS Main_transportation_type
          ,'API Calculation'                                         AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                      AS Number_of_working_days
          ,'API Calculation'                                         AS PKT_p_km
    FROM       users    AS u
    LEFT JOIN  WorkTime AS e ON ${join}
    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  const query = `${cte} ${selectBody} ORDER BY u.userId OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
  const exportQuery = `${cte} ${selectBody} ORDER BY u.userId`;
  const countQuery = `${cte} SELECT COUNT(*) AS total FROM (${selectBody}) AS Sub`;

  return { query, countQuery, exportQuery };
};

export const getADataExcelFactoryCat7 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);
  const { exportQuery } = buildQuery(factory, dateFrom, dateToExclusive);
  const replacements = hasDate ? [dateFrom, dateToExclusive] : [];

  const data = await db.query(exportQuery, {
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

export const buildQueryAutoSentCMS = (
  factory: string,
  dateFrom?: string,
  dateTo?: string,
  factoryAddress?: string,
) => {
  const isLYM = factory === 'LYM';

  const baseWhere = !isLYM
    ? "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.Address_Live IS NOT NULL"
    : 'WHERE 1=1 AND u.Vehicle IS NOT NULL AND dwt.workhours > 0 AND u.Address_Live IS NOT NULL';

  const dateFilter = !isLYM
    ? 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?'
    : 'AND CONVERT(DATE ,dwt.CDate) BETWEEN ? AND ?';

  const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

  const workCol = !isLYM ? 'WORKING_TIME' : 'workhours';

  const table = !isLYM ? 'Data_Work_Time' : 'HR_Attendance';

  const join = !isLYM
    ? 'u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS'
    : 'u.userId = dwt.UserNo';

  const query = `SELECT TOP 30 *, N'${getFactory(factory)}' AS Factory_Name
                  FROM (
                        SELECT u.userId                       AS Staff_ID
                              ,CASE 
                                    WHEN u.Address_Live IS NULL THEN u.PickupDropoffStation
                                    ELSE u.Address_Live
                              END                            AS Residential_address
                              ,u.Vehicle                      AS Main_transportation_type
                              ,'API Calculation'              AS km
                              ,COUNT(${workCol})  AS Number_of_working_days
                              ,'API Calculation'              AS PKT_p_km
                              ,'${factoryAddress || 'N/A'}'       AS Factory_address
                        FROM   ${table}                 AS dwt
                              LEFT JOIN users                AS u
                                    ON  ${join}
                        ${where}
                        GROUP BY
                                u.userId
                                ,u.Address_Live
                                ,u.Vehicle
                                ,u.PickupDropoffStation
                  ) as a`;

  // console.log(query);
  const countQuery = `SELECT COUNT(*) AS total
                      FROM   (
                                SELECT TOP 30 *
                                FROM (
                                      SELECT u.userId                       AS Staff_ID
                                            ,CASE 
                                                  WHEN u.Address_Live IS NULL THEN u.PickupDropoffStation
                                                  ELSE u.Address_Live
                                            END                            AS Residential_address
                                            ,u.Vehicle                      AS Main_transportation_type
                                            ,'API Calculation'              AS km
                                            ,COUNT(${workCol})  AS Number_of_working_days
                                            ,'API Calculation'              AS PKT_p_km
                                            ,'${factoryAddress || 'N/A'}'       AS Factory_address
                                      FROM   ${table}                 AS dwt
                                            LEFT JOIN users                AS u
                                                  ON  ${join}
                                      ${where}
                                      GROUP BY
                                              u.userId
                                              ,u.Address_Live
                                              ,u.Vehicle
                                              ,u.PickupDropoffStation
                                ) as a	
                      ) AS Sub`;
  return { query, countQuery };
};

// export const buildQueryAutoSentCmsLYV = async (
//   dateFrom?: string,
//   dateTo?: string,
//   db?: Sequelize,
// ) => {
//   const queryAddress = `SELECT Coordinates as [Address]
//                         FROM CMW_Info_Factory
//                         WHERE CreatedFactory = 'LYV'`;

//   const factoryAddress =
//     (await db?.query(queryAddress, {
//       type: QueryTypes.SELECT,
//     })) || [];

//   const baseWhere =
//     "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.lock = '0'";
//   // "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.Address_Live IS NOT NULL";

//   const dateFilter = 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT *
//                         ,N'${getFactory('LYV')}'  AS Factory_Name
//                   FROM   (
//                             SELECT u.userId               AS Staff_ID
//                                   ,CASE
//                                       WHEN ISNULL(u.lat ,'')<>''
//                                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat, ', ',u.long)
//                                       ELSE u.Address_Live
//                                   END                 AS Residential_address
//                                   ,u.Vehicle              AS Main_transportation_type
//                                   ,'API Calculation'      AS km
//                                   ,COUNT(WORKING_TIME)    AS Number_of_working_days
//                                   ,'API Calculation'      AS PKT_p_km
//                                   ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
//                                   ,CASE
//                                         WHEN ISNULL(dds.Department_Name ,'')='' THEN dd.Department_Name
//                                         ELSE CONCAT(dd.Department_Name ,' - ' ,dds.Department_Name)
//                                     END                    AS Department_Name
//                             FROM   Data_Work_Time         AS dwt
//                                     LEFT JOIN users        AS u
//                                         ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//                                             Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Person  AS dp
//                                         ON  dp.Person_Serial_Key = dwt.Person_Serial_Key
//                                             AND dp.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = u.Person_Serial_Key COLLATE
//                                                 Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Department AS dd
//                                         ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//                                     LEFT JOIN Data_Department_s AS dds
//                                         ON  dds.Department_Serial_Key = dp.Department_Serial_Keys
//                             ${where}
//                             GROUP BY
//                                     u.userId
//                                     ,u.Address_Live
//                                     ,u.Vehicle
//                                     ,u.PickupDropoffStation
//                                     ,dd.Department_Name
//                                     ,dds.Department_Name
//                                     ,u.lat
//                                     ,u.long
//                     ) AS a`;

//   // console.log(query);

//   //   const query = `SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Bicycle'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Bus'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Car'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Company shuttle bus'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Electric motorcycle'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Motorcycle'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Subway'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a
//   // UNION ALL
//   // SELECT TOP 20*
//   //       ,N'樂億 - LYV'  AS Factory_Name
//   // FROM   (
//   //            SELECT u.userId                          AS Staff_ID
//   //                  ,CASE
//   //                        WHEN ISNULL(u.lat ,'')<>''
//   //                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat ,', ' ,u.long)
//   //                        ELSE u.Address_Live
//   //                   END                    AS Residential_address
//   //                  ,u.Vehicle              AS Main_transportation_type
//   //                  ,'API Calculation'      AS km
//   //                  ,COUNT(WORKING_TIME)    AS Number_of_working_days
//   //                  ,'API Calculation'      AS PKT_p_km
//   //                  ,N'3-5 Ten Lua Street, An Lac Ward, Ho Chi Minh City' AS
//   //                   Factory_address
//   //                  ,dd.Department_Name
//   //            FROM   Data_Work_Time         AS dwt
//   //                   LEFT JOIN users        AS u
//   //                        ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//   //                            Chinese_Taiwan_Stroke_CI_AS
//   //                   LEFT JOIN Data_Person  AS dp
//   //                        ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//   //                            SQL_Latin1_General_CP1_CI_AS
//   //                   LEFT JOIN Data_Department AS dd
//   //                        ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//   //            WHERE  1 = 1
//   //                   AND Work_Or_Not<>'2'
//   //                   AND u.Vehicle IS NOT NULL
//   //                   AND dwt.Working_Time>0
//   //                   AND u.lock = '0'
//   //                   AND CONVERT(DATE ,dwt.Check_Day) BETWEEN N'2025-11-01' AND N'2025-11-30'
//   //                   AND u.Vehicle = 'Walking'
//   //            GROUP BY
//   //                   u.userId
//   //                  ,u.Address_Live
//   //                  ,u.Vehicle
//   //                  ,u.PickupDropoffStation
//   //                  ,dd.Department_Name
//   //                  ,u.lat
//   //                  ,u.long
//   //        )              AS a`;
//   return query;
// };

export const buildQueryAutoSentCmsLYV = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = 'LYV'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  const address =
    factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';

  const { dateToExclusive, hasDate } = buildReplacements(
    dateFrom ?? '',
    dateTo ?? '',
  );

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
      GROUP BY Person_Serial_Key
    )
  `;

  const query = `
    ${cte}
    SELECT u.userId                                                        AS Staff_ID
          ,CASE
                WHEN ISNULL(u.lat, '') <> '' AND ISNULL(u.long, '') <> '' THEN CONCAT(u.lat, ', ', u.long)
                ELSE u.Address_Live COLLATE DATABASE_DEFAULT
           END                                                             AS Residential_address
          ,u.Vehicle                                                       AS Main_transportation_type
          ,'API Calculation'                                               AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                            AS Number_of_working_days
          ,'API Calculation'                                               AS PKT_p_km
          ,N'${address}'                                                   AS Factory_address
          ,CASE
                WHEN ISNULL(dds.Department_Name, '') = '' THEN dd.Department_Name
                ELSE CONCAT(dd.Department_Name, ' - ', dds.Department_Name)
           END                                                             AS Department_Name
          ,N'${getFactory('LYV')}'                                        AS Factory_Name

    FROM       users              AS u

    LEFT JOIN  Data_Person        AS dp
            ON dp.Person_Serial_Key COLLATE DATABASE_DEFAULT = u.Person_Serial_Key COLLATE DATABASE_DEFAULT

    LEFT JOIN  Data_Department    AS dd
            ON dd.Department_Serial_Key  = dp.Department_Serial_Key

    LEFT JOIN  Data_Department_s  AS dds
            ON dds.Department_Serial_Key = dp.Department_Serial_Keys

    LEFT JOIN  WorkTime           AS e
            ON e.Person_Serial_Key COLLATE DATABASE_DEFAULT = u.Person_Serial_Key COLLATE DATABASE_DEFAULT

    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  return query;
};

// export const buildQueryAutoSentCmsLHG = async (
//   dateFrom?: string,
//   dateTo?: string,
//   db?: Sequelize,
// ) => {
//   const queryAddress = `SELECT Coordinates as [Address]
//                         FROM CMW_Info_Factory
//                         WHERE CreatedFactory = 'LHG'`;

//   const factoryAddress =
//     (await db?.query(queryAddress, {
//       type: QueryTypes.SELECT,
//     })) || [];

//   const baseWhere =
//     "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.lock = '0'";

//   const dateFilter = 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT *
//                         ,N'${getFactory('LHG')}'  AS Factory_Name
//                   FROM   (
//                             SELECT u.userId               AS Staff_ID
//                                   ,CASE
//                                       WHEN ISNULL(u.lat ,'')<>''
//                                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat, ', ',u.long)
//                                       ELSE u.Address_Live
//                                   END                 AS Residential_address
//                                   ,u.Vehicle              AS Main_transportation_type
//                                   ,'API Calculation'      AS km
//                                   ,COUNT(WORKING_TIME)    AS Number_of_working_days
//                                   ,'API Calculation'      AS PKT_p_km
//                                   ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
//                                   ,dd.Department_Name
//                             FROM   Data_Work_Time         AS dwt
//                                     LEFT JOIN users        AS u
//                                         ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//                                             Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Person  AS dp
//                                         ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//                                             SQL_Latin1_General_CP1_CI_AS
//                                     LEFT JOIN Data_Department AS dd
//                                         ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//                             ${where}
//                             GROUP BY
//                                     u.userId
//                                   ,u.Address_Live
//                                   ,u.Vehicle
//                                   ,u.PickupDropoffStation
//                                   ,dd.Department_Name
//                                   ,u.lat
//                                   ,u.long
//                         ) AS a`;

//   return query;
// };

export const buildQueryAutoSentCmsLHG = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = 'LHG'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  const address =
    factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
      GROUP BY Person_Serial_Key
    )
  `;

  const query = `
    ${cte}
    SELECT u.userId                                                        AS Staff_ID
          ,CASE
                WHEN ISNULL(u.lat, '') <> '' AND ISNULL(u.long, '') <> '' THEN CONCAT(u.lat, ', ', u.long)
                ELSE u.Address_Live COLLATE DATABASE_DEFAULT
           END                                                             AS Residential_address
          ,u.Vehicle                                                       AS Main_transportation_type
          ,'API Calculation'                                               AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                            AS Number_of_working_days
          ,'API Calculation'                                               AS PKT_p_km
          ,N'${address}'                                                   AS Factory_address
          ,dd.Department_Name
          ,N'${getFactory('LHG')}'                                        AS Factory_Name

    FROM       users              AS u

    LEFT JOIN  Data_Person        AS dp
            ON dp.Person_Serial_Key COLLATE DATABASE_DEFAULT = u.Person_Serial_Key COLLATE DATABASE_DEFAULT

    LEFT JOIN  Data_Department    AS dd
            ON dd.Department_Serial_Key  = dp.Department_Serial_Key

    LEFT JOIN  WorkTime           AS e
            ON e.Person_Serial_Key COLLATE DATABASE_DEFAULT = u.Person_Serial_Key COLLATE DATABASE_DEFAULT

    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  return query;
};

// export const buildQueryAutoSentCmsLVL = async (
//   dateFrom?: string,
//   dateTo?: string,
//   db?: Sequelize,
// ) => {
//   const queryAddress = `SELECT Coordinates as [Address]
//                         FROM CMW_Info_Factory
//                         WHERE CreatedFactory = 'LVL'`;

//   const factoryAddress =
//     (await db?.query(queryAddress, {
//       type: QueryTypes.SELECT,
//     })) || [];

//   const baseWhere =
//     "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.lock = '0'";

//   const dateFilter = 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT *
//                         ,N'${getFactory('LVL')}'  AS Factory_Name
//                   FROM   (
//                             SELECT u.userId                          AS Staff_ID
//                                   ,CASE
//                                       WHEN ISNULL(u.lat ,'')<>''
//                                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat, ', ',u.long)
//                                       ELSE u.Address_Live
//                                   END                 AS Residential_address
//                                   ,u.Vehicle              AS Main_transportation_type
//                                   ,'API Calculation'      AS km
//                                   ,COUNT(WORKING_TIME)    AS Number_of_working_days
//                                   ,'API Calculation'      AS PKT_p_km
//                                   ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
//                                   ,dd.Department_Name
//                             FROM   Data_Work_Time         AS dwt
//                                     LEFT JOIN users        AS u
//                                         ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//                                             Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Person  AS dp
//                                         ON  dp.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = u.Person_Serial_Key COLLATE
//                                             SQL_Latin1_General_CP1_CI_AS
//                                     LEFT JOIN Data_Department AS dd ON dd.Department_Serial_Key = dp.Department_Serial_Key
//                             ${where}
//                             GROUP BY
//                                   u.userId
//                                   ,u.Address_Live
//                                   ,u.Vehicle
//                                   ,u.PickupDropoffStation
//                                   ,dd.Department_Name
//                                   ,u.lat
//                                   ,u.long
//                         )             AS a`;
//   return query;
// };

export const buildQueryAutoSentCmsLVL = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = 'LVL'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  const address =
    factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
      GROUP BY Person_Serial_Key
    )
  `;

  const query = `
    ${cte}
    SELECT u.userId                                                        AS Staff_ID
          ,CASE
                WHEN ISNULL(u.lat, '') <> '' AND ISNULL(u.long, '') <> '' THEN CONCAT(u.lat, ', ', u.long)
                ELSE u.Address_Live COLLATE DATABASE_DEFAULT
           END                                                             AS Residential_address
          ,u.Vehicle                                                       AS Main_transportation_type
          ,'API Calculation'                                               AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                            AS Number_of_working_days
          ,'API Calculation'                                               AS PKT_p_km
          ,N'${address}'                                                   AS Factory_address
          ,dd.Department_Name
          ,N'${getFactory('LVL')}'                                        AS Factory_Name

    FROM       users              AS u

    LEFT JOIN  Data_Person        AS dp
            ON dp.Person_Serial_Key  COLLATE DATABASE_DEFAULT = u.Person_Serial_Key                    COLLATE DATABASE_DEFAULT

    LEFT JOIN  Data_Department    AS dd
            ON dd.Department_Serial_Key = dp.Department_Serial_Key

    LEFT JOIN  WorkTime           AS e
            ON e.Person_Serial_Key   COLLATE DATABASE_DEFAULT = ISNULL(u.Person_Serial_Key, u.userId) COLLATE DATABASE_DEFAULT

    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  return query;
};

// export const buildQueryAutoSentCmsLYM = async (
//   dateFrom?: string,
//   dateTo?: string,
//   db?: Sequelize,
// ) => {
//   const queryAddress = `SELECT Coordinates as [Address]
//                         FROM CMW_Info_Factory
//                         WHERE CreatedFactory = 'LYM'`;

//   const factoryAddress =
//     (await db?.query(queryAddress, {
//       type: QueryTypes.SELECT,
//     })) || [];

//   const baseWhere = `WHERE 1=1 AND dwt.workhours > 0 AND u.lock = '0' AND u.Vehicle IS NOT NULL`;

//   const dateFilter = 'AND CONVERT(DATE ,dwt.CDate) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT *
//                         ,N'${getFactory('LYM')}'  AS Factory_Name
//                   FROM   (
//                             SELECT u.userId                       AS Staff_ID
//                                   ,CASE
//                                       WHEN ISNULL(u.lat ,'')<>''
//                                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat, ', ',u.long)
//                                       ELSE u.Address_Live
//                                   END                 AS Residential_address
//                                   ,u.Vehicle           AS Main_transportation_type
//                                   ,'API Calculation'   AS km
//                                   ,COUNT(workhours)    AS Number_of_working_days
//                                   ,'API Calculation'   AS PKT_p_km
//                                   ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
//                                   ,hu.Part AS Department_Name
//                             FROM   HR_Attendance       AS dwt
//                                     LEFT JOIN users     AS u
//                                         ON  u.userId = dwt.UserNo
//                                     LEFT JOIN HR_Users  AS hu
//                                         ON  hu.UserNo = dwt.UserNo
//                                             AND hu.UserNo = u.userId
//                             ${where}
//                             GROUP BY
//                                   u.userId
//                                   ,u.Address_Live
//                                   ,u.Vehicle
//                                   ,u.PickupDropoffStation
//                                   ,hu.Part
//                                   ,u.lat
//                                   ,u.long
//                         )            AS a`;

//   return query;
// };

export const buildQueryAutoSentCmsLYM = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  // ✅ Lấy factory address
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = 'LYM'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  const address =
    factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT UserNo
            ,COUNT(workhours) AS Number_of_Working_Days
      FROM   HR_Attendance
      WHERE  workhours > 0
             ${hasDate ? `AND CONVERT(DATE, CDate) >= CONVERT(DATE, ?) AND CONVERT(DATE, CDate) < CONVERT(DATE, ?)` : ''}
      GROUP BY UserNo
    )
  `;

  const query = `
    ${cte}
    SELECT u.userId                                                        AS Staff_ID
          ,CASE
                WHEN ISNULL(u.lat, '') <> '' AND ISNULL(u.long, '') <> '' THEN CONCAT(u.lat, ', ', u.long)
                ELSE u.Address_Live
           END                                                             AS Residential_address
          ,u.Vehicle                                                       AS Main_transportation_type
          ,'API Calculation'                                               AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                            AS Number_of_working_days
          ,'API Calculation'                                               AS PKT_p_km
          ,N'${address}'                                                   AS Factory_address
          ,hu.Part                                                         AS Department_Name
          ,N'${getFactory('LYM')}'                                        AS Factory_Name

    FROM       users         AS u

    LEFT JOIN  HR_Users      AS hu
            ON hu.UserNo     = u.userId

    LEFT JOIN  WorkTime      AS e
            ON e.UserNo      = u.userId

    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  return query;
};

// export const buildQueryAutoSentCmsJAZ = async (
//   dateFrom?: string,
//   dateTo?: string,
//   db?: Sequelize,
// ) => {
//   const queryAddress = `SELECT Coordinates as [Address]
//                         FROM CMW_Info_Factory
//                         WHERE CreatedFactory = 'JAZ'`;

//   const factoryAddress =
//     (await db?.query(queryAddress, {
//       type: QueryTypes.SELECT,
//     })) || [];

//   const baseWhere =
//     "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.lock = '0'";
//   // "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.Address_Live IS NOT NULL";

//   const dateFilter = 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT *
//                         ,N'${getFactory('JAZ')}'  AS Factory_Name
//                   FROM   (
//                             SELECT u.userId               AS Staff_ID
//                                   ,CASE
//                                       WHEN ISNULL(u.lat ,'')<>''
//                                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat, ', ',u.long)
//                                       ELSE u.Address_Live
//                                   END                 AS Residential_address
//                                   ,u.Vehicle              AS Main_transportation_type
//                                   ,'API Calculation'      AS km
//                                   ,COUNT(WORKING_TIME)    AS Number_of_working_days
//                                   ,'API Calculation'      AS PKT_p_km
//                                   ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
//                                   ,dd.Department_Name
//                             FROM   Data_Work_Time         AS dwt
//                                     LEFT JOIN users        AS u
//                                         ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//                                             Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Person  AS dp
//                                         ON  dp.Person_Serial_Key = dwt.Person_Serial_Key
//                                             AND dp.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = u.Person_Serial_Key COLLATE
//                                                 Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Department AS dd
//                                         ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//                             ${where}
//                             GROUP BY
//                                     u.userId
//                                     ,u.Address_Live
//                                     ,u.Vehicle
//                                     ,u.PickupDropoffStation
//                                     ,dd.Department_Name
//                                     ,u.lat
//                                     ,u.long
//                     ) AS a`;
//   return query;
// };

export const buildQueryAutoSentCmsJAZ = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = 'JAZ'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  const address =
    factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
      GROUP BY Person_Serial_Key
    )
  `;

  const query = `
    ${cte}
    SELECT u.userId                                                        AS Staff_ID
          ,CASE
                WHEN ISNULL(u.lat, '') <> '' AND ISNULL(u.long, '') <> '' THEN CONCAT(u.lat, ', ', u.long)
                ELSE u.Address_Live COLLATE DATABASE_DEFAULT
           END                                                             AS Residential_address
          ,u.Vehicle                                                       AS Main_transportation_type
          ,'API Calculation'                                               AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                            AS Number_of_working_days
          ,'API Calculation'                                               AS PKT_p_km
          ,N'${address}'                                                   AS Factory_address
          ,dd.Department_Name
          ,N'${getFactory('JAZ')}'                                        AS Factory_Name

    FROM       users              AS u

    LEFT JOIN  Data_Person        AS dp
            ON dp.Person_Serial_Key  COLLATE DATABASE_DEFAULT = u.Person_Serial_Key                    COLLATE DATABASE_DEFAULT

    LEFT JOIN  Data_Department    AS dd
            ON dd.Department_Serial_Key = dp.Department_Serial_Key

    LEFT JOIN  WorkTime           AS e
            ON e.Person_Serial_Key   COLLATE DATABASE_DEFAULT = ISNULL(u.Person_Serial_Key, u.userId) COLLATE DATABASE_DEFAULT

    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  return query;
};

// export const buildQueryAutoSentCmsJZS = async (
//   dateFrom?: string,
//   dateTo?: string,
//   db?: Sequelize,
// ) => {
//   const queryAddress = `SELECT Coordinates as [Address]
//                         FROM CMW_Info_Factory
//                         WHERE CreatedFactory = 'JZS'`;

//   const factoryAddress =
//     (await db?.query(queryAddress, {
//       type: QueryTypes.SELECT,
//     })) || [];

//   const baseWhere =
//     "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0 AND u.lock = '0'";

//   const dateFilter = 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?';

//   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT *
//                         ,N'${getFactory('JZS')}'  AS Factory_Name
//                   FROM   (
//                             SELECT u.userId               AS Staff_ID
//                                   ,CASE
//                                       WHEN ISNULL(u.lat ,'')<>''
//                                             AND ISNULL(u.long ,'')<>'' THEN CONCAT(u.lat, ', ',u.long)
//                                       ELSE u.Address_Live
//                                   END                 AS Residential_address
//                                   ,u.Vehicle              AS Main_transportation_type
//                                   ,'API Calculation'      AS km
//                                   ,COUNT(WORKING_TIME)    AS Number_of_working_days
//                                   ,'API Calculation'      AS PKT_p_km
//                                   ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
//                                   ,dd.Department_Name
//                             FROM   Data_Work_Time         AS dwt
//                                     LEFT JOIN users        AS u
//                                         ON  u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE
//                                             Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Person  AS dp
//                                         ON  dp.Person_Serial_Key = dwt.Person_Serial_Key
//                                             AND dp.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = u.Person_Serial_Key COLLATE
//                                                 Chinese_Taiwan_Stroke_CI_AS
//                                     LEFT JOIN Data_Department AS dd
//                                         ON  dd.Department_Serial_Key = dp.Department_Serial_Key
//                             ${where}
//                             GROUP BY
//                                     u.userId
//                                     ,u.Address_Live
//                                     ,u.Vehicle
//                                     ,u.PickupDropoffStation
//                                     ,dd.Department_Name
//                                     ,u.lat
//                                     ,u.long
//                     ) AS a`;
//   return query;
// };

export const buildQueryAutoSentCmsJZS = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  // ✅ Lấy factory address
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = 'JZS'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  const address =
    factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND Check_Day >= ? AND Check_Day < ?` : ''}
      GROUP BY Person_Serial_Key
    )
  `;

  const query = `
    ${cte}
    SELECT u.userId                                                        AS Staff_ID
          ,CASE
                WHEN ISNULL(u.lat, '') <> '' AND ISNULL(u.long, '') <> '' THEN CONCAT(u.lat, ', ', u.long)
                ELSE u.Address_Live COLLATE DATABASE_DEFAULT
           END                                                             AS Residential_address
          ,u.Vehicle                                                       AS Main_transportation_type
          ,'API Calculation'                                               AS km
          ,ISNULL(e.Number_of_Working_Days, 0)                            AS Number_of_working_days
          ,'API Calculation'                                               AS PKT_p_km
          ,N'${address}'                                                   AS Factory_address
          ,dd.Department_Name
          ,N'${getFactory('JZS')}'                                        AS Factory_Name

    FROM       users              AS u

    LEFT JOIN  Data_Person        AS dp
            ON dp.Person_Serial_Key COLLATE DATABASE_DEFAULT = u.Person_Serial_Key COLLATE DATABASE_DEFAULT

    LEFT JOIN  Data_Department    AS dd
            ON dd.Department_Serial_Key = dp.Department_Serial_Key

    LEFT JOIN  WorkTime           AS e
            ON e.Person_Serial_Key COLLATE DATABASE_DEFAULT = u.Person_Serial_Key COLLATE DATABASE_DEFAULT

    WHERE  u.lock     = '0'
       AND u.Vehicle IS NOT NULL
  `;

  return query;
};

// export const buildQueryCustomExport = (
//   dateFrom: string,
//   dateTo: string,
//   factory: string,
// ) => {
//   const baseWhere: string =
//     factory !== 'LYM'
//       ? `WHERE a.lock = '0' AND e.Work_Or_Not<>'2' AND e.Working_Time>0 AND a.Vehicle IS NOT NULL`
//       : `WHERE  c.workhours>0 AND a.lock = '0' AND a.Vehicle IS NOT NULL`;

//   const dateFilter: string =
//     factory !== 'LYM'
//       ? ` AND CONVERT(DATE ,e.Check_Day) BETWEEN ? AND ?`
//       : ` AND CONVERT(DATE ,c.CDate) BETWEEN ? AND ?`;

//   const table: string =
//     factory !== 'LYM'
//       ? `users                         AS a
//           LEFT JOIN Data_Person         AS b
//                 ON  a.userId COLLATE SQL_Latin1_General_CP1_CI_AS = b.Person_ID COLLATE SQL_Latin1_General_CP1_CI_AS
//           LEFT JOIN Data_Department     AS c
//                 ON  b.Department_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = c.Department_Serial_Key COLLATE
//                     SQL_Latin1_General_CP1_CI_AS
//           LEFT JOIN Data_Person_Detail  AS d
//                 ON  d.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = a.Person_Serial_Key COLLATE
//                     SQL_Latin1_General_CP1_CI_AS
//           LEFT JOIN Data_Work_Time      AS e
//                 ON  e.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = a.Person_Serial_Key COLLATE
//                     SQL_Latin1_General_CP1_CI_AS`
//       : `users                    AS a
//        LEFT JOIN HR_Users       AS b
//             ON  b.UserNo = a.userId
//        LEFT JOIN HR_Attendance  AS c
//             ON  c.UserNo = a.userId
//                 AND c.UserNo = b.UserNo`;
//   const groupBy: string =
//     factory !== 'LYM'
//       ? `c.Department_Name
//       ,a.userId
//       ,a.fullName
//       ,d.Address_Live
//       ,a.Vehicle
//       ,a.Bus_Route
//       ,a.PickupDropoffStation
//       ,a.Address_Live`
//       : `b.Part
//       ,a.userId
//       ,a.fullName
//       ,a.Address_Live
//       ,a.Bus_Route
//       ,a.PickupDropoffStation
//       ,b.Addr_now
// 	  ,a.Vehicle`;

//   const where: string =
//     dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

//   const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY a.userId) AS INT) AS [No]
//                           ,'${factory}'                         AS Factory
//                           ,${factory !== 'LYM' ? 'c.Department_Name' : 'b.Part'}             AS Department
//                           ,a.userId                      AS ID
//                           ,a.fullName                    AS Full_Name
//                           ,CASE
//                                 WHEN a.Vehicle='Company shuttle bus' THEN ${factory !== 'LYM' ? 'd.Address_Live' : 'b.Addr_now'} COLLATE SQL_Latin1_General_CP1_CI_AS
//                                 ELSE a.Address_Live COLLATE SQL_Latin1_General_CP1_CI_AS
//                           END                           AS Current_Address
//                           ,a.Vehicle                     AS Transportation_Mode
//                           ,CASE
//                                 WHEN a.Vehicle='Company shuttle bus' THEN a.Bus_Route COLLATE SQL_Latin1_General_CP1_CI_AS
//                                 ELSE 'N/A' COLLATE SQL_Latin1_General_CP1_CI_AS
//                           END                           AS Bus_Route
//                           ,CASE
//                                 WHEN a.Vehicle='Company shuttle bus' THEN a.PickupDropoffStation COLLATE SQL_Latin1_General_CP1_CI_AS
//                                 ELSE 'N/A' COLLATE SQL_Latin1_General_CP1_CI_AS
//                           END                           AS Pickup_Point
//                           ,COUNT(${factory !== 'LYM' ? 'e.WORKING_TIME' : 'c.workhours'})         AS Number_of_Working_Days
//                     FROM   ${table}
//                     ${where}
//                     GROUP BY ${groupBy}`;

//   const countQuery = `SELECT COUNT(ID) AS total
//                       FROM   (
//                                 SELECT CAST(ROW_NUMBER() OVER(ORDER BY a.userId) AS INT) AS [No]
//                                       ,'${factory}'                         AS Factory
//                                       ,${factory !== 'LYM' ? 'c.Department_Name' : 'b.Part'}             AS Department
//                                       ,a.userId                      AS ID
//                                       ,a.fullName                    AS Full_Name
//                                       ,CASE
//                                             WHEN a.Vehicle='Company shuttle bus' THEN ${factory !== 'LYM' ? 'd.Address_Live' : 'b.Addr_now'} COLLATE SQL_Latin1_General_CP1_CI_AS
//                                             ELSE a.Address_Live COLLATE SQL_Latin1_General_CP1_CI_AS
//                                       END                           AS Current_Address
//                                       ,a.Vehicle                     AS Transportation_Mode
//                                       ,CASE
//                                             WHEN a.Vehicle='Company shuttle bus' THEN a.Bus_Route COLLATE SQL_Latin1_General_CP1_CI_AS
//                                             ELSE 'N/A' COLLATE SQL_Latin1_General_CP1_CI_AS
//                                       END                           AS Bus_Route
//                                       ,CASE
//                                             WHEN a.Vehicle='Company shuttle bus' THEN a.PickupDropoffStation COLLATE SQL_Latin1_General_CP1_CI_AS
//                                             ELSE 'N/A' COLLATE SQL_Latin1_General_CP1_CI_AS
//                                       END                           AS Pickup_Point
//                                       ,COUNT(${factory !== 'LYM' ? 'e.WORKING_TIME' : 'c.workhours'})         AS Number_of_Working_Days
//                                 FROM   ${table}
//                                 ${where}
//                                 GROUP BY ${groupBy}
//                       ) AS Sub`;
//   return { query, countQuery };
// };

export const buildQueryCustomExport = (
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  const isLYM = factory === 'LYM';

  const department = isLYM ? 'b.Part' : 'c.Department_Name';
  const addrLive = isLYM ? 'b.Addr_now' : 'd.Address_Live';
  const cte = isLYM
    ? `
      WITH WorkTime AS (
        SELECT UserNo
              ,COUNT(workhours) AS Number_of_Working_Days
        FROM   HR_Attendance
        WHERE  workhours > 0
               ${dateFrom && dateTo ? `AND CONVERT(DATE, CDate) >= CONVERT(DATE, ?) AND CONVERT(DATE, CDate) < CONVERT(DATE, ?)` : ''}
        GROUP BY UserNo
      )
    `
    : `
      WITH WorkTime AS (
        SELECT Person_Serial_Key
              ,COUNT(WORKING_TIME) AS Number_of_Working_Days
        FROM   Data_Work_Time
        WHERE  Work_Or_Not <> '2'
               AND Working_Time > 0
               ${dateFrom && dateTo ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
        GROUP BY Person_Serial_Key
      )
    `;

  const joins = isLYM
    ? `
      LEFT JOIN HR_Users      AS b ON b.UserNo            = a.userId
      LEFT JOIN WorkTime      AS e ON e.UserNo            = a.userId
    `
    : `
      LEFT JOIN Data_Person        AS b ON b.Person_ID             COLLATE DATABASE_DEFAULT = a.userId                COLLATE DATABASE_DEFAULT
      LEFT JOIN Data_Department    AS c ON c.Department_Serial_Key COLLATE DATABASE_DEFAULT = b.Department_Serial_Key COLLATE DATABASE_DEFAULT
      LEFT JOIN Data_Person_Detail AS d ON d.Person_Serial_Key     COLLATE DATABASE_DEFAULT = a.Person_Serial_Key     COLLATE DATABASE_DEFAULT
      LEFT JOIN WorkTime           AS e ON e.Person_Serial_Key     COLLATE DATABASE_DEFAULT = ISNULL(a.Person_Serial_Key, a.userId) COLLATE DATABASE_DEFAULT
    `;

  const selectBody = `
    SELECT CAST(ROW_NUMBER() OVER (ORDER BY a.userId) AS INT)    AS [No]
          ,'${factory}'                                           AS Factory
          ,${department}                                          AS Department
          ,a.userId                                               AS ID
          ,a.fullName                                             AS Full_Name
          ,CASE
                WHEN a.Vehicle = 'Company shuttle bus' THEN ${addrLive} COLLATE DATABASE_DEFAULT
                ELSE a.Address_Live COLLATE DATABASE_DEFAULT
           END                                                    AS Current_Address
          ,a.Vehicle                                              AS Transportation_Mode
          ,CASE
                WHEN a.Vehicle = 'Company shuttle bus' THEN a.Bus_Route COLLATE DATABASE_DEFAULT
                ELSE 'N/A'
           END                                                    AS Bus_Route
          ,CASE
                WHEN a.Vehicle = 'Company shuttle bus' THEN a.PickupDropoffStation COLLATE DATABASE_DEFAULT
                ELSE 'N/A'
           END                                                    AS Pickup_Point
          ,ISNULL(e.Number_of_Working_Days, 0)                   AS Number_of_Working_Days
    FROM       users AS a
    ${joins}
    WHERE  a.lock     = '0'
       AND a.Vehicle IS NOT NULL
  `;

  const query = `${cte} ${selectBody} ORDER BY a.userId OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
  const countQuery = `${cte} SELECT COUNT(*) AS total FROM (${selectBody}) AS Sub`;

  return { query, countQuery };
};

export const buildReplacements = (
  dateFrom: string,
  dateTo: string,
): {
  dateToExclusive: string;
  hasDate: boolean;
} => {
  return {
    dateToExclusive: dateTo
      ? dayjs(dateTo).add(1, 'day').format('YYYY-MM-DD')
      : '',
    hasDate: !!(dateFrom && dateTo),
  };
};

export const getADataCustomExport = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
  factory: string,
  fields?: string[],
) => {
  const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);
  const { query } = buildQueryCustomExport(dateFrom, dateToExclusive, factory);

  const replacements = hasDate
    ? [dateFrom, dateToExclusive, 0, 999999]
    : [0, 999999];

  const data = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  if (!data?.length) return;

  const selectedFields = fields?.length ? fields : Object.keys(data[0]);

  sheet.columns = selectedFields.map((key) => ({
    header: key.replace(/_/g, ' '),
    key,
    width: 15,
  }));

  // ✅ Style header row
  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E78' },
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    } as ExcelJS.Borders;
  });

  // ✅ Thêm data rows + style luôn
  data.forEach((item, index) => {
    const row = sheet.addRow(
      selectedFields.reduce(
        (acc, key) => {
          acc[key] = item[key] ?? '';
          return acc;
        },
        {} as Record<string, any>,
      ),
    );

    // ✅ Zebra striping cho dễ đọc
    const isEven = index % 2 === 0;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      } as ExcelJS.Borders;
      cell.alignment = { vertical: 'middle' };
      if (isEven) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF2F2F2' },
        };
      }
    });
  });

  // ✅ Auto-fit column width
  sheet.columns.forEach((column) => {
    let maxLength = column.header ? String(column.header).length : 10;
    column.eachCell?.({ includeEmpty: true }, (cell) => {
      maxLength = Math.max(
        maxLength,
        cell.value ? String(cell.value).length : 0,
      );
    });
    column.width = Math.min(maxLength * 1.2, 50);
  });
};
// export const getADataCustomExport = async (
//   sheet: ExcelJS.Worksheet,
//   db: Sequelize,
//   dateFrom: string,
//   dateTo: string,
//   factory: string,
//   fields?: string[],
// ) => {
//   const { query } = buildQueryCustomExport(dateFrom, dateTo, factory);
//   const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

//   const data = await db.query(query, {
//     replacements,
//     type: QueryTypes.SELECT,
//   });
//   if (!data || data.length === 0) return;

//   const allKeys = Object.keys(data[0]);

//   const selectedFields = fields && fields.length > 0 ? fields : allKeys;

//   sheet.columns = selectedFields.map((key) => ({
//     header: key.replace(/_/g, ' '),
//     key,
//   }));

//   data.forEach((item) => {
//     const rowData: Record<string, any> = {};
//     selectedFields.forEach((key) => {
//       rowData[key] = item[key] ?? '';
//     });
//     sheet.addRow(rowData);
//   });

//   // console.log(data);

//   sheet.columns.forEach((column) => {
//     let maxLength = 0;
//     if (typeof column.eachCell === 'function') {
//       column.eachCell({ includeEmpty: true }, (cell) => {
//         const cellValue = cell.value ? String(cell.value) : '';
//         maxLength = Math.max(maxLength, cellValue.length);
//       });
//     }
//     column.width = maxLength * 1.2;
//   });

//   sheet.eachRow({ includeEmpty: true }, (row) => {
//     row.eachCell({ includeEmpty: true }, (cell) => {
//       cell.border = {
//         top: { style: 'thin' },
//         left: { style: 'thin' },
//         bottom: { style: 'thin' },
//         right: { style: 'thin' },
//       };
//     });
//   });
// };
