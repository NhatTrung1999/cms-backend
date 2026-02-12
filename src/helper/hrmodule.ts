export const buildQueryHRModule = (
  dateFrom: string,
  dateTo: string,
  fullName: string,
  id: string,
  department: string,
  joinDate: string,
  factory: string,
) => {
  const isLYM = factory === 'LYM';
  const baseWhere: string = isLYM
    ? `WHERE  c.workhours>0 AND a.lock = '0'`
    : `WHERE a.lock = '0' AND e.Work_Or_Not<>'2' AND e.Working_Time>0`;

  const conditions: string[] = [];
  if (dateFrom && dateTo) {
    const dateField = isLYM ? 'c.CDate' : 'e.Check_Day';
    conditions.push(`CONVERT(DATE, ${dateField}) BETWEEN ? AND ?`);
  }

  if (fullName) {
    conditions.push(`a.fullName LIKE ?`);
  }

  if (id) {
    conditions.push(`a.userId LIKE ?`);
  }

  if (department) {
    conditions.push(
      `${isLYM ? `d.__sno = ?` : `c.Department_Serial_Key LIKE ?`}`,
    );
  }

  if (joinDate) {
    // const dateField = isLYM ? 'b.[Start_Date]' : 'b.Date_Come_In';
    conditions.push(
      `${isLYM ? `CAST(b.[Start_Date] AS DATE) = ?` : `CONVERT(VARCHAR ,b.Date_Come_In ,23) = ?`}`,
    );
  }

  const select: string = isLYM
    ? `a.userId                 AS ID
      ,b.Part                   AS Department
      ,a.fullName               AS FullName
      ,b.[Start_Date]           AS JoinDate
      ,b.Addr_now               AS PermanentAddress
      ,a.Address_Live           AS CurrentAddress
      ,a.Vehicle                AS TransportationMethod
      ,COUNT(c.workhours)       AS Number_of_Working_Days`
    : `a.userId                      AS ID
      ,c.Department_Name             AS Department
      ,a.fullName                    AS FullName
      ,b.Date_Come_In                AS JoinDate
      ,d.Address_Live                AS PermanentAddress
      ,a.Address_Live                AS CurrentAddress
      ,a.Vehicle                     AS TransportationMethod
      ,COUNT(e.WORKING_TIME)         AS Number_of_Working_Days`;

  const table: string = isLYM
    ? `users                    AS a
                    LEFT JOIN HR_Users       AS b
                            ON  b.UserNo = a.userId
                    LEFT JOIN HR_Attendance  AS c
                            ON  c.UserNo = a.userId
                                AND c.UserNo = b.UserNo
                    LEFT JOIN HR_DEPT AS d ON d.DeptName = b.Part`
    : `users                         AS a
       LEFT JOIN Data_Person         AS b
            ON  a.userId = b.Person_ID COLLATE Database_Default
       LEFT JOIN Data_Department     AS c
            ON  b.Department_Serial_Key = c.Department_Serial_Key COLLATE Database_Default
       LEFT JOIN Data_Person_Detail  AS d
            ON  d.Person_Serial_Key = a.Person_Serial_Key COLLATE Database_Default
       LEFT JOIN Data_Work_Time      AS e
            ON  e.Person_Serial_Key = ISNULL(a.Person_Serial_Key, a.userId) COLLATE Database_Default`;

  const groupBy: string = isLYM
    ? `a.userId
                    ,b.Part
                    ,a.fullName
                    ,b.[Start_Date]
                    ,b.Addr_now
                    ,a.Address_Live
                    ,a.Vehicle`
    : `a.userId
                    ,c.Department_Name
                    ,a.fullName
                    ,b.Date_Come_In
                    ,d.Address_Live
                    ,a.Address_Live
                    ,a.Vehicle`;
  const where: string =
    conditions.length > 0
      ? `${baseWhere} AND ${conditions.join(' AND ')}`
      : baseWhere;

  const query = `SELECT ${select}
                    FROM   ${table}
                    ${where}
                    GROUP BY ${groupBy}`;
  const countQuery = `SELECT COUNT(ID) AS total
                        FROM   (
                                SELECT ${select}
                                FROM   ${table}
                                ${where}
                                GROUP BY ${groupBy}
                        ) AS Sub`;
  return { query, countQuery };
};
