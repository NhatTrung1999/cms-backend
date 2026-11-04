import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import dayjs from 'dayjs';
import { buildReplacements } from './date.helper';
import { fillSheet } from './excel.helper';
import { CAT7_EXPORT_COLUMNS } from './cat7.columns';

/** Address (coordinates) of the factory from CMW_Info_Factory, or 'N/A' when there is no db or no row. */
export const getFactoryAddress = async (factory: string, db?: Sequelize) => {
  const factoryAddress = db
    ? await db.query(
        `SELECT Coordinates AS [Address] FROM CMW_Info_Factory WHERE CreatedFactory = '${factory}'`,
        { type: QueryTypes.SELECT },
      )
    : [];

  return factoryAddress.length > 0 ? factoryAddress[0]['Address'] : 'N/A';
};
dayjs().format();

export const buildQuery = (
  factory: string,
  dateFrom?: string,
  dateTo?: string,
) => {
  const isLYM = factory === 'LYM';
  const isLYV = factory === 'LYV';
  const isLVL = factory === 'LVL';

  const workCol = isLYM ? 'workhours' : 'WORKING_TIME';
  const table = isLYM ? 'HR_Attendance' : 'Data_Work_Time';
  const dateCol = isLYM ? 'CDate' : 'Check_Day';
  const cteGroupKey = isLYM ? 'UserNo' : 'Person_Serial_Key';
  const workFilter = isLYM
    ? 'workhours > 0'
    : "Work_Or_Not <> '2' AND Working_Time > 0";

  const join = isLYM
    ? 'e.UserNo = u.userId'
    : 'e.Person_Serial_Key COLLATE DATABASE_DEFAULT = ISNULL(u.Person_Serial_Key, u.userId) COLLATE DATABASE_DEFAULT';

  const dateFilter =
    dateFrom && dateTo
      ? `AND CONVERT(DATE ,${dateCol}) >= CONVERT(DATE ,?) AND CONVERT(DATE ,${dateCol}) < CONVERT(DATE ,?)`
      : '';

  const checkInOutFilter = isLYM
    ? `AND ISNULL(NULLIF(MatTime1, ''), MatTime1) <> '1900-01-01 00:00:00.000'
       AND ISNULL(NULLIF(MatTime2, ''), MatTime2) <> '1900-01-01 00:00:00.000'`
    : isLYV
      ? `AND Check_In  <> '1900-01-01 00:00:00.000'
       AND Check_Out <> '1900-01-01 00:00:00.000'`
      : isLVL
        ? `AND cio.Check_In  <> '1900-01-01 00:00:00.000'
       AND cio.Check_Out <> '1900-01-01 00:00:00.000'`
        : `AND Check_In  <> '1900-01-01 00:00:00.000'
       AND Check_Out <> '1900-01-01 00:00:00.000'`;

  const cteFrom = isLVL
    ? `Data_Work_Time AS dwt
       LEFT JOIN Choose_In_Out AS cio
              ON cio.Person_Serial_Key = dwt.Person_Serial_Key
             AND cio.Check_day         = dwt.Check_Day
             AND cio.Shift_Serial_key  = dwt.Shift_Serial_key`
    : table;

  const cte = isLVL
    ? `
      WITH WorkTime AS (
        SELECT dwt.Person_Serial_Key
              ,COUNT(dwt.${workCol}) AS Number_of_Working_Days
        FROM   ${cteFrom}
        WHERE  dwt.Work_Or_Not <> '2'
               AND dwt.Working_Time > 0
               ${dateFrom && dateTo ? `AND dwt.${dateCol} >= ? AND dwt.${dateCol} < ?` : ''}
               ${checkInOutFilter}
        GROUP BY dwt.Person_Serial_Key
      )
    `
    : `
      WITH WorkTime AS (
        SELECT ${cteGroupKey}
              ,COUNT(${workCol}) AS Number_of_Working_Days
        FROM   ${table}
        WHERE  ${workFilter}
               ${dateFilter}
               ${checkInOutFilter}
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
    WHERE  ISNULL(e.Number_of_Working_Days, 0) <> 0
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
  fillSheet(sheet, CAT7_EXPORT_COLUMNS, data);
};
