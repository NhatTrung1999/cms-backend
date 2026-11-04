import * as ExcelJS from 'exceljs';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { getFactory } from './factory.helper';
import { CAT5_COLUMNS } from './cat5.columns';
import { fillSheet } from './excel.helper';

export const getADataExcelFactoryCat5 = async (
  sheet: ExcelJS.Worksheet,
  db: Sequelize,
  dateFrom: string,
  dateTo: string,
) => {
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
  const data = await db.query(query, {
    replacements,
    type: QueryTypes.SELECT,
  });

  fillSheet(sheet, CAT5_COLUMNS, data);
};

export const buildQueryAutoSentCMS = async (
  dateFrom?: string,
  dateTo?: string,
  factory?: string,
  db?: Sequelize,
) => {
  const queryAddress = `SELECT ${factory?.trim().toLowerCase() === 'lym'.trim().toLowerCase() ? `N'TSANG YIH Co., Ltd /Polo/ချန်းရိ' as [Address]` : '[Address]'}
                        FROM CMW_Info_Factory
                        WHERE CreatedFactory = '${factory}'`;

  const factoryAddress =
    (await db?.query(queryAddress, {
      type: QueryTypes.SELECT,
    })) || [];

  const baseWhere = 'WHERE 1=1 AND dwc.DISABLED = 0 AND dwo.QUANTITY<>0';

  const dateFilter = 'AND dwo.WASTE_DATE BETWEEN ? AND ?';

  const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

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
                          ,N'${factoryAddress.length === 0 ? 'N/A' : factoryAddress[0]['Address']}' AS Factory_address
                          ,N'${getFactory(factory || 'unknow_factory')}'  AS Factory_Name
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

  return query;
};
