/** WHERE part + replacements of the Cat5 (waste) page queries. */
export const buildCat5Filter = (dateFrom: string, dateTo: string) => {
  let where = 'WHERE 1=1 AND dwc.DISABLED = 0 AND dwo.QUANTITY<>0';
  const replacements: any[] = [];

  if (dateTo && dateFrom) {
    where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo);
  }

  return { where, replacements };
};

/** Rows query (the same SQL for one factory and for each factory of the "ALL" page). */
export const buildCat5PageQuery = (
  where: string,
) => `SELECT dwo.WASTE_DATE                     AS Waste_disposal_date
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

/** Count query of the same page. */
export const buildCat5CountQuery = (where: string) => `SELECT COUNT(*) AS total
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
