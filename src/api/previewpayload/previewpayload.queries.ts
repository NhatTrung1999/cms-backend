/** SQL of the Cat6 business trips (UOF database): one row per traveler, with the department of the traveler. */
export const buildCat6PreviewQuery = (where: string) => `
      WITH Travelers AS (
                          SELECT chb.*
                                ,chb.UserCreate          AS TravelerID
                                ,0                       AS TravelerOrder
                          FROM   CDS_HRBUSS_BusTripData     chb
                          WHERE  chb.BPMStatus = 'F'
                                  AND (
                                          chb.AssisstedIDs IS NULL
                                          OR LTRIM(RTRIM(chb.AssisstedIDs))=''
                                      )
                                  ${where}
                          UNION
                          ALL
                          SELECT chb.*
                                ,LTRIM(RTRIM(t.v.value('.' ,'nvarchar(50)'))) AS TravelerID
                                ,ROW_NUMBER() OVER(
                                      PARTITION BY chb.TripID
                                      ORDER BY(
                                          SELECT NULL
                                      )
                                  )  AS TravelerOrder
                          FROM   CDS_HRBUSS_BusTripData chb
                                  CROSS APPLY (
                              SELECT CAST(
                                          '<x>'
                                        +REPLACE(REPLACE(chb.AssisstedIDs ,',' ,'$') ,'$' ,'</x><x>')
                                        +'</x>' AS XML
                                      ) AS DATA
                          )         AS s
                          CROSS APPLY s.data.nodes('/x') AS t(v)
                          WHERE  chb.BPMStatus = 'F'
                                  AND chb.AssisstedIDs IS NOT NULL
                                  AND LTRIM(RTRIM(chb.AssisstedIDs))<>''
                                  AND LTRIM(RTRIM(t.v.value('.' ,'nvarchar(50)')))<>''
                                  ${where}
                      )

                  SELECT tr.*
                        ,COALESCE(
                            dp.Department_Name COLLATE SQL_Latin1_General_CP1_CI_AS
                            ,vwd.GROUP_NAME
                            ,(
                                SELECT TOP 1 vwd2.GROUP_NAME
                                FROM   TB_EB_USER teu2
                                        OUTER APPLY (
                                    SELECT teed2.GROUP_ID
                                    FROM   TB_EB_EMPL_DEP AS teed2
                                    WHERE  teed2.USER_GUID = teu2.USER_GUID
                                            AND teed2.ORDERS = 0
                                ) teed2
                                LEFT JOIN vwDepartment_Factory vwd2
                                            ON  vwd2.GROUP_ID = teed2.GROUP_ID
                                WHERE  teu2.ACCOUNT = ISNULL(tr.Factory_User ,'')+tr.TravelerID
                                        AND vwd2.GROUP_NAME IS NOT NULL
                            )
                            ,(
                                SELECT TOP 1 vwd3.GROUP_NAME
                                FROM   TB_EB_USER teu3
                                        OUTER APPLY (
                                    SELECT teed3.GROUP_ID
                                    FROM   TB_EB_EMPL_DEP AS teed3
                                    WHERE  teed3.USER_GUID = teu3.USER_GUID
                                            AND teed3.ORDERS = 0
                                ) teed3
                                LEFT JOIN vwDepartment_Factory vwd3
                                            ON  vwd3.GROUP_ID = teed3.GROUP_ID
                                WHERE  teu3.ACCOUNT = ISNULL(tr.Departure ,'')+tr.TravelerID
                                        AND vwd3.GROUP_NAME IS NOT NULL
                            )
                            ,(
                                SELECT TOP 1 vwd4.GROUP_NAME
                                FROM   CDS_FMEval_Employee cfe
                                        JOIN TB_EB_USER teu4
                                            ON  teu4.ACCOUNT = cfe.BPMAccount
                                        OUTER APPLY (
                                    SELECT teed4.GROUP_ID
                                    FROM   TB_EB_EMPL_DEP AS teed4
                                    WHERE  teed4.USER_GUID = teu4.USER_GUID
                                            AND teed4.ORDERS = 0
                                ) teed4
                                LEFT JOIN vwDepartment_Factory vwd4
                                            ON  vwd4.GROUP_ID = teed4.GROUP_ID
                                WHERE  cfe.EmpID = tr.TravelerID
                                        AND vwd4.GROUP_NAME IS NOT NULL
                            )
                            ,(
                                SELECT TOP 1 vwd5.GROUP_NAME
                                FROM   TB_EB_USER teu5
                                        OUTER APPLY (
                                    SELECT teed5.GROUP_ID
                                    FROM   TB_EB_EMPL_DEP AS teed5
                                    WHERE  teed5.USER_GUID = teu5.USER_GUID
                                            AND teed5.ORDERS = 0
                                ) teed5
                                LEFT JOIN vwDepartment_Factory vwd5
                                            ON  vwd5.GROUP_ID = teed5.GROUP_ID
                                WHERE  teu5.ACCOUNT = tr.TravelerID
                                        AND vwd5.GROUP_NAME IS NOT NULL
                            )
                            ,(
                                SELECT TOP 1 vwd6.GROUP_NAME
                                FROM   TB_EB_USER teu6
                                        JOIN TB_EB_EMPL_DEP teed6
                                            ON  teed6.USER_GUID = teu6.USER_GUID
                                        LEFT JOIN vwDepartment_Factory vwd6
                                            ON  vwd6.GROUP_ID = teed6.GROUP_ID
                                WHERE  teu6.ACCOUNT = tr.TravelerID
                                        AND vwd6.GROUP_NAME IS NOT NULL
                                ORDER BY
                                        teed6.ORDERS
                            )
                            ,(
                                SELECT TOP 1 vwd7.GROUP_NAME
                                FROM   TB_EB_USER teu7
                                        JOIN TB_EB_EMPL_DEP teed7
                                            ON  teed7.USER_GUID = teu7.USER_GUID
                                        LEFT JOIN vwDepartment_Factory vwd7
                                            ON  vwd7.GROUP_ID = teed7.GROUP_ID
                                WHERE  teu7.ACCOUNT LIKE '%'+tr.TravelerID
                                        AND vwd7.GROUP_NAME IS NOT NULL
                                ORDER BY
                                        teed7.ORDERS
                            )
                        )                        AS Dept
                        ,COUNT(*) OVER()          AS TotalRow
                  FROM   Travelers tr
                        OUTER APPLY (
                      SELECT TOP 1 teu.USER_GUID
                      FROM   TB_EB_USER teu
                      WHERE  (
                                tr.Factory_User IS NOT NULL
                                AND teu.ACCOUNT=tr.Factory_User+tr.TravelerID
                            )
                            OR (tr.Factory_User IS NULL AND teu.ACCOUNT=tr.TravelerID)
                  )                               AS teu
                  OUTER APPLY (
                      SELECT teed.GROUP_ID
                      FROM   TB_EB_EMPL_DEP AS teed
                      WHERE  teed.USER_GUID = teu.USER_GUID
                            AND teed.ORDERS = 0
                  )                               AS teed
                  LEFT JOIN vwDepartment_Factory  AS vwd
                              ON  vwd.GROUP_ID = teed.GROUP_ID
                        LEFT JOIN [JZS_HRIS].[HRIS].[dbo].[View_Data_Person] dp
                              ON  dp.Person_ID COLLATE SQL_Latin1_General_CP1_CI_AS = tr.TravelerID COLLATE SQL_Latin1_General_CP1_CI_AS
                  WHERE  tr.DOC_NBR LIKE 'LYV-HR-BT%'
                        OR tr.DOC_NBR LIKE 'LHG-SUGG%'
                        OR tr.DOC_NBR LIKE 'LVL-HR-BTF%'
                        OR tr.DOC_NBR LIKE 'LVL-ODBT%'
                        OR tr.DOC_NBR LIKE 'LYM-HR-BT%'
                        OR tr.DOC_NBR LIKE 'JZS-SUGG%'
                        OR tr.DOC_NBR LIKE 'JAZ_BizTrip%'
                  ORDER BY
                        tr.CreatedAt ASC
                        ,tr.DOC_NBR
                        ,tr.TravelerOrder;
    `;
