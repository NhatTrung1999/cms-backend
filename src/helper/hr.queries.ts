/**
 * Department list SQL of each factory's HRIS database (label / value pairs of the department filter).
 * JAZ and JZS use the same text.
 */
const JAZ_AND_JZS_DEPARTMENT_QUERY = `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                    ORDER BY
                          Department_Name`;

export const DEPARTMENT_QUERY_BY_FACTORY: Record<string, string> = {
  LYV: `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                          AND ISNULL(Hide ,'0')<>'1'
                    ORDER BY Department_Name`,
  LHG: `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                          AND Building_Overtime<>''
                    ORDER BY
                          Department_Name`,
  LVL: `SELECT DISTINCT Department_Name  AS label
                          ,Department_Serial_Key     AS value
                    FROM   Data_Department
                    WHERE  NoUse_Date = '2099-12-31 00:00:00.000'
                          AND ISNULL(Hide ,'0')<>'1'
                    ORDER BY
                          Department_Name`,
  LYM: `SELECT DISTINCT DeptName  AS label
                          ,__sno              AS value
                    FROM   HR_DEPT
                    ORDER BY
                          DeptName`,
  JAZ: JAZ_AND_JZS_DEPARTMENT_QUERY,
  JZS: JAZ_AND_JZS_DEPARTMENT_QUERY,
};
