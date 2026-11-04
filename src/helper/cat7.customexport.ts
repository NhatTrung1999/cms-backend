import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildReplacements } from './date.helper';

export const buildQueryCustomExport = (
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  const isLYM = factory === 'LYM';
  const isLYV = factory === 'LYV';
  const isLVL = factory === 'LVL';

  const department = isLYM ? 'b.Part' : 'c.Department_Name';
  const addrLive = isLYM ? 'b.Addr_now' : 'd.Address_Live';

  const dateFilter =
    dateFrom && dateTo
      ? `AND CONVERT(DATE, ${isLYM ? 'CDate' : 'Check_Day'}) >= CONVERT(DATE, ?) AND CONVERT(DATE, ${isLYM ? 'CDate' : 'Check_Day'}) < CONVERT(DATE, ?)`
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

  const cte = isLYM
    ? `
      WITH WorkTime AS (
        SELECT UserNo
              ,COUNT(workhours) AS Number_of_Working_Days
        FROM   HR_Attendance
        WHERE  workhours > 0
               ${dateFilter}
               ${checkInOutFilter}
        GROUP BY UserNo
      )
    `
    : isLVL
      ? `
      WITH WorkTime AS (
        SELECT dwt.Person_Serial_Key
              ,COUNT(dwt.WORKING_TIME) AS Number_of_Working_Days
        FROM   Data_Work_Time          AS dwt
        LEFT JOIN Choose_In_Out        AS cio
               ON cio.Person_Serial_Key = dwt.Person_Serial_Key
              AND cio.Check_day         = dwt.Check_Day
              AND cio.Shift_Serial_key  = dwt.Shift_Serial_key
        WHERE  dwt.Work_Or_Not <> '2'
               AND dwt.Working_Time > 0
               ${dateFrom && dateTo ? `AND CONVERT(DATE, dwt.Check_Day) >= CONVERT(DATE, ?) AND CONVERT(DATE, dwt.Check_Day) < CONVERT(DATE, ?)` : ''}
               ${checkInOutFilter}
        GROUP BY dwt.Person_Serial_Key
      )
    `
      : `
      WITH WorkTime AS (
        SELECT Person_Serial_Key
              ,COUNT(WORKING_TIME) AS Number_of_Working_Days
        FROM   Data_Work_Time
        WHERE  Work_Or_Not <> '2'
               AND Working_Time > 0
               ${dateFilter}
               ${checkInOutFilter}
        GROUP BY Person_Serial_Key
      )
    `;

  const joins = isLYM
    ? `
      LEFT JOIN HR_Users      AS b ON b.UserNo        = a.userId
      LEFT JOIN WorkTime      AS e ON e.UserNo        = a.userId
    `
    : `
      LEFT JOIN Data_Person        AS b ON b.Person_ID             COLLATE DATABASE_DEFAULT = a.userId                COLLATE DATABASE_DEFAULT
      LEFT JOIN Data_Department    AS c ON c.Department_Serial_Key COLLATE DATABASE_DEFAULT = b.Department_Serial_Key COLLATE DATABASE_DEFAULT
      LEFT JOIN Data_Person_Detail AS d ON d.Person_Serial_Key     COLLATE DATABASE_DEFAULT = a.Person_Serial_Key     COLLATE DATABASE_DEFAULT
      LEFT JOIN WorkTime           AS e ON e.Person_Serial_Key     COLLATE DATABASE_DEFAULT = ISNULL(a.Person_Serial_Key, a.userId) COLLATE DATABASE_DEFAULT
    `;

  const selectBody = `
    SELECT CAST(ROW_NUMBER() OVER (ORDER BY a.userId) AS INT) AS [No]
          ,'${factory}'                                            AS Factory
          ,${department}                                           AS Department
          ,a.userId                                                AS ID
          ,a.fullName                                              AS Full_Name
          ,CASE
                WHEN a.Vehicle = 'Company shuttle bus' THEN ${addrLive} COLLATE DATABASE_DEFAULT
                ELSE a.Address_Live COLLATE DATABASE_DEFAULT
           END                                                     AS Current_Address
          ,a.Vehicle                                               AS Transportation_Mode
          ,CASE
                WHEN a.Vehicle = 'Company shuttle bus' THEN a.Bus_Route COLLATE DATABASE_DEFAULT
                ELSE 'N/A'
           END                                                     AS Bus_Route
          ,CASE
                WHEN a.Vehicle = 'Company shuttle bus' THEN a.PickupDropoffStation COLLATE DATABASE_DEFAULT
                ELSE 'N/A'
           END                                                     AS Pickup_Point
          ,ISNULL(e.Number_of_Working_Days, 0)                    AS Number_of_Working_Days
    FROM       users AS a
    ${joins}
    WHERE ISNULL(e.Number_of_Working_Days, 0) <> 0
  `;

  const query = `${cte} ${selectBody} ORDER BY a.userId OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
  const countQuery = `${cte} SELECT COUNT(*) AS total FROM (${selectBody}) AS Sub`;

  return { query, countQuery };
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
