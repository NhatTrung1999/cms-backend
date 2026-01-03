export const buildQueryHRModule = (
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  const baseWhere: string =
    factory !== 'LYM'
      ? `WHERE a.lock = '0' AND e.Work_Or_Not<>'2' AND e.Working_Time>0`
      : `WHERE  c.workhours>0 AND a.lock = '0'`;

  const dateFilter: string =
    factory !== 'LYM'
      ? ` AND CONVERT(DATE ,e.Check_Day) BETWEEN ? AND ?`
      : ` AND CONVERT(DATE ,c.CDate) BETWEEN ? AND ?`;

  const select: string =
    factory !== 'LYM'
      ? `a.userId                      AS ID
      ,c.Department_Name             AS Department
      ,a.fullName                    AS FullName
      ,d.Address_Live                AS PermanentAddress
      ,a.Address_Live                AS CurrentAddress
      ,a.Vehicle                     AS TransportationMode
      ,COUNT(e.WORKING_TIME)         AS Number_of_Working_Days`
      : `a.userId                 AS ID
      ,b.Part                   AS Department
      ,a.fullName               AS Full_Name
      ,b.Addr_now               AS PermanentAddress
      ,a.Address_Live           AS CurrentAddress
      ,a.Vehicle                AS TransportationMode
      ,COUNT(c.workhours)       AS Number_of_Working_Days`;

  const table: string =
    factory !== 'LYM'
      ? `users                         AS a
                    LEFT JOIN Data_Person         AS b
                            ON  a.userId COLLATE SQL_Latin1_General_CP1_CI_AS = b.Person_ID COLLATE SQL_Latin1_General_CP1_CI_AS
                    LEFT JOIN Data_Department     AS c
                            ON  b.Department_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = c.Department_Serial_Key COLLATE
                                SQL_Latin1_General_CP1_CI_AS
                    LEFT JOIN Data_Person_Detail  AS d
                            ON  d.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = a.Person_Serial_Key COLLATE 
                                SQL_Latin1_General_CP1_CI_AS
                    LEFT JOIN Data_Work_Time      AS e
                            ON  e.Person_Serial_Key COLLATE SQL_Latin1_General_CP1_CI_AS = a.Person_Serial_Key COLLATE
                                SQL_Latin1_General_CP1_CI_AS`
      : `users                    AS a
                    LEFT JOIN HR_Users       AS b
                            ON  b.UserNo = a.userId
                    LEFT JOIN HR_Attendance  AS c
                            ON  c.UserNo = a.userId
                                AND c.UserNo = b.UserNo`;
  const groupBy: string =
    factory !== 'LYM'
      ? `a.userId
                    ,c.Department_Name
                    ,a.fullName
                    ,d.Address_Live
                    ,a.Address_Live
                    ,a.Vehicle`
      : `a.userId
                    ,b.Part
                    ,a.fullName
                    ,b.Addr_now
                    ,a.Address_Live
                    ,a.Vehicle`;
  const where: string =
    dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

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
