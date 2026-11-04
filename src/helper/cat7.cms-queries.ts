import { Sequelize } from 'sequelize-typescript';
import { getFactory } from './factory.helper';
import { getFactoryAddress } from './cat7.helper';
import { buildReplacements } from './date.helper';

// SQL sent to each factory HRIS database to read the Cat7 rows of the CMS payload.
// The six queries are really different (tables, joins, collations), so they are not merged.

export const buildQueryAutoSentCmsLYV = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const address = await getFactoryAddress('LYV', db);

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
             AND Check_In <>'1900-01-01 00:00:00.000'
                AND Check_Out <>'1900-01-01 00:00:00.000'
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

    WHERE ISNULL(e.Number_of_Working_Days ,0)<>0
  `;

  return query;
};

export const buildQueryAutoSentCmsLHG = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const address = await getFactoryAddress('LHG', db);

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             AND Check_In <>'1900-01-01 00:00:00.000'
             AND Check_Out <>'1900-01-01 00:00:00.000'
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

    WHERE ISNULL(e.Number_of_Working_Days ,0)<>0
  `;

  return query;
};

export const buildQueryAutoSentCmsLVL = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const address = await getFactoryAddress('LVL', db);

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT dwt.Person_Serial_Key
               ,COUNT(dwt.WORKING_TIME)  AS Number_of_Working_Days
         FROM   Data_Work_Time           AS dwt
                LEFT JOIN Choose_In_Out  AS cio
                     ON  cio.Person_Serial_Key = dwt.Person_Serial_Key
                         AND cio.Check_day = dwt.Check_Day
                         AND cio.Shift_Serial_key = dwt.Shift_Serial_key
         WHERE  dwt.Work_Or_Not<>'2'
                AND dwt.Working_Time>0
                ${hasDate ? `AND CONVERT(DATE, dwt.Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, dwt.Check_Day) < CONVERT(DATE, ?)` : ''}
                AND cio.Check_In <>'1900-01-01 00:00:00.000'
                AND cio.Check_Out <>'1900-01-01 00:00:00.000'
         GROUP BY
                dwt.Person_Serial_Key
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
            ON dp.Person_Serial_Key  COLLATE DATABASE_DEFAULT = ISNULL(u.Person_Serial_Key, u.userId)                    COLLATE DATABASE_DEFAULT

    LEFT JOIN  Data_Department    AS dd
            ON dd.Department_Serial_Key = dp.Department_Serial_Key

    LEFT JOIN  WorkTime           AS e
            ON e.Person_Serial_Key   COLLATE DATABASE_DEFAULT = ISNULL(u.Person_Serial_Key, u.userId) COLLATE DATABASE_DEFAULT

    WHERE ISNULL(e.Number_of_Working_Days ,0)<>0
  `;

  return query;
};

export const buildQueryAutoSentCmsLYM = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  // ✅ Lấy factory address
  const address = await getFactoryAddress('LYM', db);

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT UserNo
            ,COUNT(workhours) AS Number_of_Working_Days
      FROM   HR_Attendance
      WHERE  workhours > 0
             ${hasDate ? `AND CONVERT(DATE, CDate) >= CONVERT(DATE, ?) AND CONVERT(DATE, CDate) < CONVERT(DATE, ?)` : ''}
             AND ISNULL(NULLIF(MatTime1 ,'') ,MatTime1)<>'1900-01-01 00:00:00.000'
              AND ISNULL(NULLIF(MatTime2 ,'') ,MatTime2)<>'1900-01-01 00:00:00.000'
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

    WHERE ISNULL(e.Number_of_Working_Days ,0)<>0
  `;

  return query;
};

export const buildQueryAutoSentCmsJAZ = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  const address = await getFactoryAddress('JAZ', db);

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND CONVERT(DATE, Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, Check_Day) < CONVERT(DATE, ?)` : ''}
             AND Check_In <> '1900-01-01 00:00:00.000'
                 AND Check_Out <>'1900-01-01 00:00:00.000'
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

    WHERE ISNULL(e.Number_of_Working_Days ,0)<>0
  `;

  return query;
};

export const buildQueryAutoSentCmsJZS = async (
  dateFrom?: string,
  dateTo?: string,
  db?: Sequelize,
) => {
  // ✅ Lấy factory address
  const address = await getFactoryAddress('JZS', db);

  const { hasDate } = buildReplacements(dateFrom ?? '', dateTo ?? '');

  const cte = `
    WITH WorkTime AS (
      SELECT Person_Serial_Key
            ,COUNT(WORKING_TIME) AS Number_of_Working_Days
      FROM   Data_Work_Time
      WHERE  Work_Or_Not <> '2'
             AND Working_Time > 0
             ${hasDate ? `AND Check_Day >= ? AND Check_Day < ?` : ''}
             AND Check_In <>'1900-01-01 00:00:00.000'
                AND Check_Out <>'1900-01-01 00:00:00.000'
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

    WHERE ISNULL(e.Number_of_Working_Days ,0)<>0
  `;

  return query;
};
