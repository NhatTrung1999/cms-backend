import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class Cat5Service {
  constructor(
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
  ) {}

  async getDataWMS(
    date: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Waste_disposal_date',
    sortOrder: string = 'asc',
  ) {
    try {
      const offset = (page - 1) * limit;

      let where = 'WHERE 1=1';
      const replacements: any[] = [];

      if (date) {
        where += ` AND CONVERT(VARCHAR, dwo.WASTE_DATE, 23) = ?`;
        replacements.push(date);
      }

      const query = `SELECT dwo.WASTE_DATE                     AS Waste_disposal_date
                              ,dtv.TREATMENT_VENDOR_NAME          AS Vender_Name
                              ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
                              ,CAST('0' AS INT)                   AS Transportation_Distance_km
                              ,CASE 
                                    WHEN dwo.HAZARDOUS<>'N/A' THEN 'hazardous waste'
                                    WHEN dwo.NON_HAZARDOUS<>'N/A' THEN 'Non-hazardous waste'
                                    ELSE NULL
                              END                                AS The_type_of_waste
                              ,CASE 
                                    WHEN dwo.HAZARDOUS<>'N/A' THEN dwo.HAZARDOUS
                                    WHEN dwo.NON_HAZARDOUS<>'N/A' THEN dwo.NON_HAZARDOUS
                                    ELSE NULL
                              END                                AS Waste_type
                              ,dtm.TREATMENT_METHOD_ENGLISH_NAME  AS Waste_Treatment_method
                              ,dwo.QUANTITY                       AS Weight_of_waste_treated_Unit_kg
                              ,CAST('0' AS INT)                   AS TKT_Ton_km
                        FROM   dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                              LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                                    ON  dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                              LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                                    ON  dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                              LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                                    ON  dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                              LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                                    ON  dwc.WASTE_CODE = dwo.WASTE_CODE
                              LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                                    ON  td.CODE = dwo.LOCATION_CODE
                        ${where}
                        ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
                        OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;

      const countQuery = `SELECT COUNT(*) AS total
                        FROM dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                        LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                              ON dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                        LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                              ON dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                        LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                              ON dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                        LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                              ON dwc.WASTE_CODE = dwo.WASTE_CODE
                        LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                              ON td.CODE = dwo.LOCATION_CODE
                        ${where}`;
      const connects = [this.LYV_WMS, this.LVL_WMS];

      const [dataResults, countResults] = await Promise.all([
        Promise.all(
          connects.map((conn) => {
            return conn.query(query, { type: QueryTypes.SELECT, replacements });
          }),
        ),
        Promise.all(
          connects.map((conn) => {
            return conn.query(countQuery, {
              type: QueryTypes.SELECT,
              replacements,
            });
          }),
        ),
      ]);

      const data = dataResults.flat();

      const total = countResults.reduce((sum, result) => {
        return sum + ((result[0] as { total: number })?.total || 0);
      }, 0);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }
}
