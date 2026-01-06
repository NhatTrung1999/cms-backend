import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQueryAutoSentCMS } from 'src/helper/cat5.helper';
import dayjs from 'dayjs';
dayjs().format();

@Injectable()
export class Cat5Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
  ) {}

  async getDataWMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Waste_disposal_date',
    sortOrder: string = 'asc',
  ) {
    let db: Sequelize;
    switch (factory) {
      case 'LYV':
        db = this.LYV_WMS;
        break;
      case 'LHG':
        db = this.LHG_WMS;
        break;
      case 'LYM':
        db = this.LYM_WMS;
        break;
      case 'LVL':
        db = this.LVL_WMS;
        break;
      case 'JAZ':
        db = this.JAZ_WMS;
        break;
      case 'JZS':
        db = this.JZS_WMS;
        break;
      default:
        return await this.getAllDataFactory(
          dateFrom,
          dateTo,
          page,
          limit,
          sortField,
          sortOrder,
        );
    }

    return await this.getDataFactory(
      db,
      dateFrom,
      dateTo,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  private async getDataFactory(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
    const offset = (page - 1) * limit;
    let where = 'WHERE 1=1 AND dwc.DISABLED = 0 AND dwo.QUANTITY<>0';
    const replacements: any[] = [];

    if (dateTo && dateFrom) {
      where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    const query = `SELECT dwo.WASTE_DATE                     AS Waste_disposal_date
                          ,dwc.CONSOLIDATED_WASTE_CODE AS Consolidated_Waste
                          ,dwc.WASTE_CODE AS Waste_Code
                          ,dtv.TREATMENT_VENDOR_NAME          AS Vendor_Name
                          ,dtv.TREATMENT_VENDOR_ID            AS Vendor_ID
                          ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
                          ,dwc.LOCATION_CODE AS Location_Code
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
                          ,dtm.TREATMENT_METHOD_ID            AS Treatment_Method_ID
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
                    ${where}`;
    // ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    // OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;

    // console.log(query);
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

    const [dataResults, countResults] = await Promise.all([
      db.query(query, { replacements, type: QueryTypes.SELECT }),
      db.query(countQuery, {
        type: QueryTypes.SELECT,
        replacements,
      }),
    ]);

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

    // const total = totalResult[0]?.total || 0;
    // const hasMore = offset + data.length < total;
    // return { data, page, limit, total, hasMore };
  }

  private async getAllDataFactory(
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;

      let where = 'WHERE 1=1 AND dwc.DISABLED = 0 AND dwo.QUANTITY<>0';
      const replacements: any[] = [];

      if (dateTo && dateFrom) {
        where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
        replacements.push(dateFrom, dateTo);
      }

      const query = `SELECT dwo.WASTE_DATE                     AS Waste_disposal_date
                            ,dwc.CONSOLIDATED_WASTE_CODE AS Consolidated_Waste
                            ,dwc.WASTE_CODE AS Waste_Code
                            ,dtv.TREATMENT_VENDOR_NAME          AS Vendor_Name
                            ,dtv.TREATMENT_VENDOR_ID            AS Vendor_ID
                            ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
                            ,dwc.LOCATION_CODE AS Location_Code
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
                            ,dtm.TREATMENT_METHOD_ID            AS Treatment_Method_ID
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
                         ${where}`;
      // ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
      // OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;

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
      const connects = [
        this.LYV_WMS,
        this.LHG_WMS,
        this.LYM_WMS,
        this.LVL_WMS,
        this.JAZ_WMS,
        this.JZS_WMS,
      ];

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

      let data = dataResults.flat();

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

      const total = countResults.reduce((sum, result) => {
        return sum + ((result[0] as { total: number })?.total || 0);
      }, 0);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    try {
      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

      const connects = [
        this.LYV_WMS,
        // this.LHG_WMS,
        // this.LYM_WMS,
        // this.LVL_WMS,
        // this.JAZ_WMS,
        // this.JZS_WMS,
      ];

      const dataResult = await Promise.all(
        connects.map(async (conn) => {
          return conn.query(
            await buildQueryAutoSentCMS(dateFrom, dateTo, this.EIP),
            {
              type: QueryTypes.SELECT,
              replacements,
            },
          );
        }),
      );

      // console.log(dataResult);
      const data = dataResult.flat();

      // console.log(data);

      const formatData = data.map((item: any) => {
        const custVenName = item.Vendor_ID;
        const departure = item.Factory_address;
        const destination = item.Waste_collection_address;
        const product = `${item.The_type_of_waste} - ${item.Waste_type}`;
        const activityData = item.Weight_of_waste_treated_Unit_kg;
        const memo = item.Waste_Treatment_method;
        const factoryName = item.Factory_Name;
        const wasteDisposalDate = item.Waste_disposal_date;

        return {
          System: 'CMW', //default
          Corporation: 'LAI YIH', //default
          Factory: factoryName,
          Department: '',
          DocKey: 'S3.C5', //default
          SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
          EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
          ActivityType: '3.6', //default
          DataType: '1', //default
          DocType: 'CMS Web', //default
          UndDoc: '',
          DocFlow: '',
          DocDate: dayjs(wasteDisposalDate).format('YYYY/MM/DD'),
          DocDate2: dayjs(wasteDisposalDate).format('YYYY/MM/DD'),
          DocNo: '',
          UndDocNo: '',
          CustVenName: custVenName,
          InvoiceNo: '',
          TransType: '陸運', //default
          Departure: departure,
          Destination: destination,
          PortType: '',
          StPort: '',
          ThPort: '',
          EndPort: '',
          Product: product,
          Quity: '',
          Amount: '',
          ActivityData: activityData,
          ActivityUnit: 'KG', //default
          Unit: '',
          UnitWeight: '',
          Memo: memo,
          CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
          Creator: '',
        };
      });

      return formatData;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  //logging
  async getLoggingCat5(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const [dataResults, countResults] = await Promise.all([
      this.EIP.query(
        `SELECT *
        FROM CMW_Category_5_Log`,
        { type: QueryTypes.SELECT },
      ),
      this.EIP.query(
        `SELECT COUNT(*) AS total
        FROM CMW_Category_5_Log`,
        {
          type: QueryTypes.SELECT,
        },
      ),
    ]);

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
    // const records: any[] = await this.EIP.query(
    //   `SELECT *
    //     FROM CMW_Category_7_Log
    //     ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    //         `,
    //   { type: QueryTypes.SELECT },
    // );
    // return records;
  }
}
