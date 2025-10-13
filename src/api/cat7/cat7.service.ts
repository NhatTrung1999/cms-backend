import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class Cat7Service {
  constructor(@Inject('HRIS') private readonly HRIS: Sequelize) {}

  async getDataCat7(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Staff_ID',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;
    let where = "WHERE 1=1 AND Work_Or_Not<>'2'";
    const replacements: any[] = [];

    if (dateTo && dateFrom) {
      where += ` AND CHECK_DAY BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    const query = `SELECT DP.Person_ID                  AS Staff_ID
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
    // ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    // OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;

    // console.log(query);
    const countQuery = `SELECT COUNT(*) AS total
                        FROM   (
                                  SELECT DP.Person_ID        AS Staff_ID
                                        ,DP.Staying_Address  AS Residential_address
                                        ,DP.Vehicle          AS Main_transportation_type
                                        ,CAST('0' AS INT)    AS km
                                        ,ROUND(SUM(WORKING_TIME) ,2) AS Number_of_working_days
                                        ,CAST('0' AS INT)    AS PKT_p_km
                                  FROM   DATA_WORK_TIME WT
                                          LEFT JOIN Data_Person_Detail AS DP
                                              ON  WT.Person_Serial_Key = DP.Person_Serial_Key
                                  ${where}
                                  GROUP BY
                                          DP.Person_ID
                                        ,DP.Staying_Address
                                        ,DP.Vehicle	
                        ) AS Sub`;

    const [dataResults, countResults] = await Promise.all([
      this.HRIS.query(query, { replacements, type: QueryTypes.SELECT }),
      this.HRIS.query(countQuery, {
        type: QueryTypes.SELECT,
        replacements,
      }),
    ]);

    console.log(dataResults, countResults);

    let data = dataResults;
    data.sort((a, b) => {
      const aValue = a[sortField];
      const bValue = b[sortField];
      if (sortOrder === 'asc') {
        return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
      } else {
        return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
      }
    });

    data = data.slice(offset, offset + limit);

    const total = (countResults[0] as { total: number })?.total || 0;

    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
  }
}
