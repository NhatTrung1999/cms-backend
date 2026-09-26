import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes, Sequelize } from 'sequelize';
import dayjs from 'dayjs';
import { getCat6AirportName } from 'src/helper/cat6-airport.helper';
import {
  appendCat6DocumentSuffix,
  getCat6AssistedIds,
} from 'src/helper/cat6-assisted.helper';
import { getFactory } from 'src/helper/factory.helper';

dayjs().format();

type MergedLegItem = {
  LegNo?: number;
  From?: string;
  To?: string;
  Transport?: string;
  StayType?: string | null;
  Nights?: number | null;
};

type AccommodationItem = {
  nights?: number | string;
  type?: string;
};

type Cat6ActiveRow = {
  Application_Day: string;
  Document_Number: string;
  Document_Number_Base: string;
  Staff_ID: string;
  Dept: string;
  Round_trip_One_way: string;
  Start_Time: string;
  End_Time: string;
  Business_Trip_Type: string;
  Departure: string;
  Destination: string;
  Transport: string;
  Leg_Index: number;
  Number_of_nights_stayed: number;
  TotalRow?: number;
  [key: string]: string | number | undefined;
};

type CMSPayload = {
  System: string;
  Corporation: string;
  Factory: string;
  Department: string;
  DocKey: string;
  ActivitySource: string;
  SPeriodData: string;
  EPeriodData: string;
  ActivityType: string;
  DataType: string;
  DocType: string;
  DocDate: string;
  DocDate2: string;
  DocNo: string;
  UndDocNo: string;
  TransType: string;
  Departure: string;
  Destination: string;
  Memo: string;
  CreateDateTime: string;
  Creator: string;
};

type CMSAccommodationPayload = {
  System: string;
  Corporation: string;
  Factory: string;
  Department: string;
  DocKey: string;
  ActivitySource: string;
  SPeriodData: string;
  EPeriodData: string;
  ActivityType: string;
  DataType: string;
  DocType: string;
  DocDate: string;
  DocDate2: string;
  DocNo: string;
  UndDocNo: string;
  TransType: string;
  ActivityData: number;
  Memo: string;
  CreateDateTime: string;
  Creator: string;
};

@Injectable()
export class Cat6Service {
  constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

  private parseJsonArray<T = unknown>(value: unknown): T[] {
    if (Array.isArray(value)) return value as T[];
    if (typeof value !== 'string' || !value.trim()) return [];
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? (parsed as T[]) : [];
    } catch {
      return [];
    }
  }

  private formatTripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const t = type.toLowerCase().trim();
    if (t === 'round' || t === 'roundtrip' || t === 'round trip')
      return 'Round trip';
    return 'One-way';
  }

  private formatBusinessTripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const map: Record<string, string> = {
      factory_domestic: 'Domestic business trip within the group',
      outside_domestic: 'Domestic business trip to third party entities',
      factory_oversea: 'Overseas business trip within the group',
      outside_oversea: 'Overseas business trip to third party entities',
    };
    return map[type.toLowerCase()] ?? type;
  }

  private extractPlacesAndTransports(routesValue: unknown) {
    const legs = this.parseJsonArray<MergedLegItem>(routesValue)
      .slice()
      .sort((a, b) => (a.LegNo ?? 0) - (b.LegNo ?? 0));

    const places: string[] = [];
    const transports: string[] = [];

    const pushPlace = (value?: string) => {
      const place = getCat6AirportName(value);
      if (!place) return;
      const lastPlace = places[places.length - 1];
      if (lastPlace === place) return;
      places.push(place);
    };

    legs.forEach((leg, index) => {
      if (index === 0) pushPlace(leg.From);
      pushPlace(leg.To);
      transports.push(leg.Transport?.trim() ?? '');
    });

    const transportLimit = Math.max(places.length - 1, 0);
    const normalizedTransports = transports.slice(0, transportLimit);
    while (normalizedTransports.length < transportLimit) {
      normalizedTransports.push('');
    }

    return { places, transports: normalizedTransports };
  }

  private buildLegsFromPlaces(
    places: string[],
    transports: string[],
  ): { dep: string; dest: string; transType: string }[] {
    const legs: { dep: string; dest: string; transType: string }[] = [];
    const maxLegs = Math.max(places.length - 1, 0);

    for (let index = 0; index < maxLegs; index += 1) {
      const dep = places[index]?.trim() ?? '';
      const dest = places[index + 1]?.trim() ?? '';
      const transType = transports[index]?.trim() ?? '';

      if (!dep || !dest || !transType) {
        continue;
      }

      legs.push({ dep, dest, transType });
    }

    return legs;
  }

  private getAccommodationNights(accommodationValue: unknown) {
    const accommodations =
      this.parseJsonArray<AccommodationItem>(accommodationValue);
    return accommodations.reduce((sum, item) => {
      if (item?.type?.trim().toLowerCase() !== 'hotel') return sum;
      const nights = Number(item?.nights ?? 0);
      return sum + (Number.isFinite(nights) ? nights : 0);
    }, 0);
  }

  private hasDormAccommodation(accommodationValue: unknown) {
    const accommodations =
      this.parseJsonArray<AccommodationItem>(accommodationValue);
    return accommodations.some(
      (item) => item.type?.trim().toLowerCase() === 'dorm',
    );
  }

  private hasCompanyShuttleCar(routesValue: unknown) {
    const routes = this.parseJsonArray<MergedLegItem>(routesValue);
    return routes.some(
      (route) =>
        route.Transport?.trim().toLowerCase() === 'company shuttle car',
    );
  }

  private expandRowsByAssistedIds(
    row: Record<string, any>,
  ): Record<string, any>[] {
    const assistedIds = getCat6AssistedIds(row.AssisstedIDs);

    if (assistedIds.length === 0) {
      return [
        {
          ...row,
          DOC_NBR: row.DOC_NBR ?? '',
          EffectiveStaffID: row.UserCreate ?? '',
        },
      ];
    }

    if (assistedIds.length === 1) {
      return [
        {
          ...row,
          DOC_NBR: row.DOC_NBR ?? '',
          EffectiveStaffID: assistedIds[0],
        },
      ];
    }

    return assistedIds.map((assistedId, index) => ({
      ...row,
      DOC_NBR: appendCat6DocumentSuffix(String(row.DOC_NBR ?? ''), index),
      EffectiveStaffID: assistedId,
    }));
  }

  private compareValues(left: unknown, right: unknown, sortOrder: string) {
    const direction = sortOrder?.toLowerCase() === 'desc' ? -1 : 1;
    const leftValue = left ?? '';
    const rightValue = right ?? '';

    if (typeof leftValue === 'number' && typeof rightValue === 'number') {
      return (leftValue - rightValue) * direction;
    }

    const leftText = String(leftValue);
    const rightText = String(rightValue);
    return (
      leftText.localeCompare(rightText, undefined, {
        numeric: true,
        sensitivity: 'base',
      }) * direction
    );
  }

  private getCMSDocKey(businessTripType: string): string {
    switch (businessTripType.trim().toLowerCase()) {
      case 'domestic business trip within the group':
        return '3.5.1';
      case 'domestic business trip to third party entities':
        return '3.5.2';
      case 'overseas business trip within the group':
        return '3.5.3';
      case 'overseas business trip to third party entities':
        return '3.5.4';
      default:
        return '';
    }
  }

  private mapCMSTransType(transport: string): string {
    const normalized = transport.trim().toLowerCase();

    switch (normalized) {
      case 'car':
        return '計程車/出租車';
      case 'company shuttle car':
        return '公司車';
      case 'flight':
        return '飛機';
      case 'personal motocycle':
        return '個人機車';
      default:
        return transport.trim();
    }
  }

  private mapRowToCMSPayload(
    row: Cat6ActiveRow,
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): CMSPayload | null {
    if (!row.Departure || !row.Destination || !row.Transport) {
      return null;
    }

    const docKey = this.getCMSDocKey(row.Business_Trip_Type ?? '');

    return {
      System: 'BPM',
      Corporation: 'LAI YIH',
      Factory: getFactory(factory ?? ''),
      Department: row.Dept ?? '',
      DocKey: docKey,
      ActivitySource: '',
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: '3.5',
      DataType: '2',
      DocType: '洽公單',
      DocDate: row.Application_Day
        ? dayjs(row.Application_Day).format('YYYY/MM/DD')
        : '',
      DocDate2: row.Start_Time
        ? dayjs(row.Start_Time).format('YYYY/MM/DD')
        : '',
      DocNo: row.Document_Number_Base ?? '',
      UndDocNo: `${row.Document_Number_Base ?? ''}-${row.Leg_Index ?? 1}`,
      TransType: this.mapCMSTransType(row.Transport),
      Departure: row.Departure,
      Destination: row.Destination,
      Memo: row.Business_Trip_Type ?? '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    };
  }

  async transformCMSData(
    rows: Cat6ActiveRow[],
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): Promise<CMSPayload[]> {
    return rows
      .map((row) => this.mapRowToCMSPayload(row, factory, dateFrom, dateTo))
      .filter((payload): payload is CMSPayload => payload !== null);
  }

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    const result = await this.getDataCat6(
      dateFrom,
      dateTo,
      factory,
      1,
      999999,
      'CreatedAt',
      'asc',
      false,
    );

    const rows = result.data as Cat6ActiveRow[];
    return this.transformCMSData(rows, factory, dateFrom, dateTo);
  }

  private mapRowToAccommodationCMSPayload(
    row: Cat6ActiveRow,
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): CMSAccommodationPayload | null {
    const nights = Number(row.Number_of_nights_stayed ?? 0);
    if (!Number.isFinite(nights) || nights <= 0) {
      return null;
    }

    return {
      System: 'BPM',
      Corporation: 'LAI YIH',
      Factory: getFactory(factory ?? ''),
      Department: row.Dept ?? '',
      DocKey: '3.5.5',
      ActivitySource: '住宿',
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: '3.5',
      DataType: '3',
      DocType: '出差住宿單',
      DocDate: row.Application_Day
        ? dayjs(row.Application_Day).format('YYYY/MM/DD')
        : '',
      DocDate2: row.Start_Time
        ? dayjs(row.Start_Time).format('YYYY/MM/DD')
        : '',
      DocNo: row.Document_Number_Base ?? '',
      UndDocNo: row.Document_Number_Base ?? '',
      TransType: 'double',
      ActivityData: nights,
      Memo: row.Business_Trip_Type ?? '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    };
  }

  async transformAccommodationCMSData(
    rows: Cat6ActiveRow[],
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): Promise<CMSAccommodationPayload[]> {
    return rows
      .map((row) =>
        this.mapRowToAccommodationCMSPayload(row, factory, dateFrom, dateTo),
      )
      .filter((payload): payload is CMSAccommodationPayload => payload !== null);
  }

  async autoSentCMSAccommodation(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    const result = await this.getDataCat6(
      dateFrom,
      dateTo,
      factory,
      1,
      999999,
      'CreatedAt',
      'asc',
      false,
    );

    const rows = result.data as Cat6ActiveRow[];
    const uniqueTripRows = Array.from(
      new Map(
        rows.map((row) => [row.Document_Number_Base, row]),
      ).values(),
    );
    return this.transformAccommodationCMSData(
      uniqueTripRows,
      factory,
      dateFrom,
      dateTo,
    );
  }

  async getDataCat6(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'CreatedAt',
    sortOrder: string = 'asc',
    checkedDormShuttle: boolean = false,
  ) {
    const replacements: any[] = [];

    let where = '';

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,chb.CreatedAt ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory.trim().toLowerCase() !== 'ALL'.trim().toLowerCase()) {
      where += ` AND chb.DOC_NBR LIKE ?`;
      replacements.push(`%${factory}%`);
    }

    const finalReplacements = [...replacements, ...replacements];

    const allowedSortFields: Record<string, string> = {
      Application_Day: 'CreatedAt',
      Document_Number: 'DOC_NBR',
      Staff_ID: 'UserCreate',
      Dept: 'Dept',
      Round_trip_One_way: 'TypeTravel',
      Start_Time: 'DateStart',
      End_Time: 'DateEnd',
      Business_Trip_Type: 'Factory',
      Number_of_nights_stayed: 'StayNight',
      CreatedAt: 'CreatedAt',
    };

    const safeSortField = allowedSortFields[sortField] ?? 'CreatedAt';
    const safeSortOrder = sortOrder?.toLowerCase() === 'desc' ? 'DESC' : 'ASC';
    const docNbrPatterns = `(chb.DOC_NBR LIKE 'LYV-HR-BT%'
                 OR chb.DOC_NBR LIKE 'LHG-SUGG%'
                 OR chb.DOC_NBR LIKE 'LVL-HR-BTF%'
                 OR chb.DOC_NBR LIKE 'LVL-ODBT%'
                 OR chb.DOC_NBR LIKE 'LYM-HR-BT%'
                 OR chb.DOC_NBR LIKE 'JZS-SUGG%'
                 OR chb.DOC_NBR LIKE 'JAZ_BizTrip%')`;
    const query = `
IF OBJECT_ID('tempdb..#Travelers')    IS NOT NULL DROP TABLE #Travelers;
IF OBJECT_ID('tempdb..#RouteNodes')   IS NOT NULL DROP TABLE #RouteNodes;
IF OBJECT_ID('tempdb..#AccomSeq')     IS NOT NULL DROP TABLE #AccomSeq;
IF OBJECT_ID('tempdb..#StopSeq')      IS NOT NULL DROP TABLE #StopSeq;
IF OBJECT_ID('tempdb..#Legs')         IS NOT NULL DROP TABLE #Legs;
IF OBJECT_ID('tempdb..#Itinerary')    IS NOT NULL DROP TABLE #Itinerary;
IF OBJECT_ID('tempdb..#DeptDistinct') IS NOT NULL DROP TABLE #DeptDistinct;
IF OBJECT_ID('tempdb..#DeptLookup')   IS NOT NULL DROP TABLE #DeptLookup;

SELECT  *
INTO    #Travelers
FROM
(
    SELECT  chb.*
           ,chb.UserCreate          AS TravelerID
           ,0                       AS TravelerOrder
    FROM    CDS_HRBUSS_BusTripData chb
    WHERE   chb.BPMStatus = 'F'
            AND (
                    chb.AssisstedIDs IS NULL
                    OR LTRIM(RTRIM(chb.AssisstedIDs))=''
                )
            AND ${docNbrPatterns}
            ${where}

    UNION ALL

    SELECT  chb.*
           ,LTRIM(RTRIM(t.v.value('.' ,'nvarchar(50)'))) AS TravelerID
           ,ROW_NUMBER() OVER(
                PARTITION BY chb.TripID
                ORDER BY(
                    SELECT NULL
                )
            )  AS TravelerOrder
    FROM    CDS_HRBUSS_BusTripData chb
            CROSS APPLY (
        SELECT CAST(
                    '<x>'
                   +REPLACE(REPLACE(chb.AssisstedIDs ,',' ,'$') ,'$' ,'</x><x>')
                   +'</x>' AS XML
                ) AS DATA
    )         AS s
    CROSS APPLY s.data.nodes('/x') AS t(v)
    WHERE   chb.BPMStatus = 'F'
            AND chb.AssisstedIDs IS NOT NULL
            AND LTRIM(RTRIM(chb.AssisstedIDs))<>''
            AND LTRIM(RTRIM(t.v.value('.' ,'nvarchar(50)')))<>''
            AND ${docNbrPatterns}
            ${where}
) AS trUnion;

CREATE CLUSTERED INDEX IX_Travelers_Key
    ON #Travelers (DOC_NBR, TravelerID, TravelerOrder);

SELECT  tr.DOC_NBR
       ,tr.TravelerID
       ,tr.TravelerOrder
       ,CAST(r.[key] AS INT)                       AS NodeIdx
       ,JSON_VALUE(r.value, '$.AddressName')       AS AddressName
       ,JSON_VALUE(r.value, '$.AddressDetail')     AS AddressDetail
       ,JSON_VALUE(r.value, '$.Transport')         AS Transport
       ,COALESCE(JSON_VALUE(r.value, '$.From'), '') AS AirFrom
       ,COALESCE(JSON_VALUE(r.value, '$.To'), '')   AS AirTo
       ,CASE WHEN LOWER(ISNULL(COALESCE(JSON_VALUE(r.value, '$.isAirport')
                                        ,JSON_VALUE(r.value, '$.IsAirport')), 'false'))
                  IN ('true', '1')
             THEN 1 ELSE 0 END                       AS IsAirport
INTO    #RouteNodes
FROM    #Travelers tr
        CROSS APPLY OPENJSON(NULLIF(LTRIM(RTRIM(tr.Routes)), '')) AS r
WHERE   ISJSON(tr.Routes) = 1;

CREATE CLUSTERED INDEX IX_RouteNodes_Key
    ON #RouteNodes (DOC_NBR, TravelerID, TravelerOrder, NodeIdx);

SELECT  tr.DOC_NBR
       ,tr.TravelerID
       ,tr.TravelerOrder
       ,ROW_NUMBER() OVER (PARTITION BY tr.DOC_NBR, tr.TravelerID, tr.TravelerOrder
                           ORDER BY CAST(a.[key] AS INT))   AS AccomIdx
       ,JSON_VALUE(a.value, '$.type')                       AS StayType
       ,JSON_VALUE(a.value, '$.address')                    AS StayAddress
       ,TRY_CONVERT(INT, JSON_VALUE(a.value, '$.nights'))    AS StayNights
       ,CASE WHEN LOWER(ISNULL(JSON_VALUE(a.value, '$.isSameAsAbove'), 'false')) IN ('true', '1')
             THEN 1 ELSE 0 END                               AS SameAsAbove
INTO    #AccomSeq
FROM    #Travelers tr
        CROSS APPLY OPENJSON(NULLIF(LTRIM(RTRIM(tr.Accommodation)), '')) AS a
WHERE   ISJSON(tr.Accommodation) = 1;

CREATE CLUSTERED INDEX IX_AccomSeq_Key
    ON #AccomSeq (DOC_NBR, TravelerID, TravelerOrder, AccomIdx);

SELECT  s.*
       ,ROW_NUMBER() OVER (PARTITION BY s.DOC_NBR, s.TravelerID, s.TravelerOrder
                           ORDER BY s.NodeIdx, s.SubIdx)    AS StopIdx
INTO    #StopSeq
FROM
(
    SELECT  rn.DOC_NBR, rn.TravelerID, rn.TravelerOrder
           ,rn.NodeIdx
           ,0                                               AS SubIdx
           ,COALESCE(NULLIF(LTRIM(RTRIM(rn.AddressDetail)), ''), rn.AddressName, '') AS StopName
           ,ISNULL(rn.Transport, '')                         AS LeaveBy
    FROM    #RouteNodes rn
    WHERE   rn.IsAirport = 0

    UNION ALL

    SELECT  rn.DOC_NBR, rn.TravelerID, rn.TravelerOrder
           ,rn.NodeIdx
           ,0                                               AS SubIdx
           ,rn.AirFrom                                       AS StopName
           ,N'Flight'                                        AS LeaveBy
    FROM    #RouteNodes rn
    WHERE   rn.IsAirport = 1
            AND NULLIF(LTRIM(RTRIM(rn.AirFrom)), '') IS NOT NULL

    UNION ALL

    SELECT  rn.DOC_NBR, rn.TravelerID, rn.TravelerOrder
           ,rn.NodeIdx
           ,1                                               AS SubIdx
           ,rn.AirTo                                         AS StopName
           ,NULL                                             AS LeaveBy
    FROM    #RouteNodes rn
    WHERE   rn.IsAirport = 1
            AND NULLIF(LTRIM(RTRIM(rn.AirTo)), '') IS NOT NULL
) s;

CREATE CLUSTERED INDEX IX_StopSeq_Key
    ON #StopSeq (DOC_NBR, TravelerID, TravelerOrder, StopIdx);

SELECT  lr.DOC_NBR, lr.TravelerID, lr.TravelerOrder
       ,ROW_NUMBER() OVER (PARTITION BY lr.DOC_NBR, lr.TravelerID, lr.TravelerOrder
                           ORDER BY lr.StopIdx)              AS LegNo
       ,lr.LegFrom
       ,lr.LegTo
       ,lr.LegTransport
INTO    #Legs
FROM
(
    SELECT  cur.DOC_NBR
           ,cur.TravelerID
           ,cur.TravelerOrder
           ,cur.StopIdx
           ,cur.StopName                                    AS LegFrom
           ,nxt.StopName                                    AS LegTo
           ,COALESCE(NULLIF(cur.LeaveBy, ''), NULLIF(nxt.LeaveBy, ''), '') AS LegTransport
    FROM    #StopSeq cur
            JOIN #StopSeq nxt
                 ON  nxt.DOC_NBR       = cur.DOC_NBR
                 AND nxt.TravelerID    = cur.TravelerID
                 AND nxt.TravelerOrder = cur.TravelerOrder
                 AND nxt.StopIdx       = cur.StopIdx + 1
    WHERE   cur.StopName <> nxt.StopName
) lr;

CREATE CLUSTERED INDEX IX_Legs_Key
    ON #Legs (DOC_NBR, TravelerID, TravelerOrder, LegNo);

SELECT  l.DOC_NBR, l.TravelerID, l.TravelerOrder
       ,(
            SELECT  l2.LegNo         AS LegNo
                   ,l2.LegFrom       AS [From]
                   ,l2.LegTo         AS [To]
                   ,l2.LegTransport  AS Transport
                   ,CASE WHEN acc.SameAsAbove = 1 THEN N'sameasabove' ELSE acc.StayType END AS StayType
                   ,acc.StayNights   AS Nights
            FROM    #Legs l2
                    LEFT JOIN #AccomSeq acc
                         ON  acc.DOC_NBR       = l2.DOC_NBR
                         AND acc.TravelerID    = l2.TravelerID
                         AND acc.TravelerOrder = l2.TravelerOrder
                         AND acc.AccomIdx      = l2.LegNo
            WHERE   l2.DOC_NBR       = l.DOC_NBR
                    AND l2.TravelerID    = l.TravelerID
                    AND l2.TravelerOrder = l.TravelerOrder
            ORDER BY l2.LegNo
            FOR JSON PATH, INCLUDE_NULL_VALUES
        )                                                     AS RoutesMerged
INTO    #Itinerary
FROM    #Legs l
GROUP BY l.DOC_NBR, l.TravelerID, l.TravelerOrder;

CREATE CLUSTERED INDEX IX_Itinerary_Key
    ON #Itinerary (DOC_NBR, TravelerID, TravelerOrder);

SELECT DISTINCT
        tr.Factory_User
       ,tr.Departure
       ,tr.TravelerID
INTO    #DeptDistinct
FROM    #Travelers tr;

SELECT  d.Factory_User
       ,d.Departure
       ,d.TravelerID
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
                WHERE  teu2.ACCOUNT = ISNULL(d.Factory_User ,'')+d.TravelerID
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
                WHERE  teu3.ACCOUNT = ISNULL(d.Departure ,'')+d.TravelerID
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
                WHERE  cfe.EmpID = d.TravelerID
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
                WHERE  teu5.ACCOUNT = d.TravelerID
                        AND vwd5.GROUP_NAME IS NOT NULL
            )
           ,(
                SELECT TOP 1 vwd6.GROUP_NAME
                FROM   TB_EB_USER teu6
                        JOIN TB_EB_EMPL_DEP teed6
                            ON  teed6.USER_GUID = teu6.USER_GUID
                        LEFT JOIN vwDepartment_Factory vwd6
                            ON  vwd6.GROUP_ID = teed6.GROUP_ID
                WHERE  teu6.ACCOUNT = d.TravelerID
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
                WHERE  teu7.ACCOUNT LIKE '%'+d.TravelerID
                        AND vwd7.GROUP_NAME IS NOT NULL
                ORDER BY
                        teed7.ORDERS
            )
        )                        AS Dept
INTO    #DeptLookup
FROM    #DeptDistinct d
        OUTER APPLY (
    SELECT TOP 1 teu.USER_GUID
    FROM   TB_EB_USER teu
    WHERE  (
               d.Factory_User IS NOT NULL
               AND teu.ACCOUNT=d.Factory_User+d.TravelerID
           )
           OR (d.Factory_User IS NULL AND teu.ACCOUNT=d.TravelerID)
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
            ON  dp.Person_ID COLLATE SQL_Latin1_General_CP1_CI_AS = d.TravelerID COLLATE SQL_Latin1_General_CP1_CI_AS;

CREATE CLUSTERED INDEX IX_DeptLookup_Key
    ON #DeptLookup (Factory_User, Departure, TravelerID);

SELECT  tr.TripID
       ,tr.DOC_NBR
       ,tr.Factory
       ,tr.Departure
       ,tr.Destination
       ,tr.TypeTravel
       ,ISNULL(it.RoutesMerged, tr.Routes)  AS Routes
       ,tr.DuringDay
       ,tr.Distance
       ,tr.TotalMoney
       ,tr.Accommodation
       ,tr.CreatedAt
       ,tr.UserCreate
       ,tr.BPMStatus
       ,tr.YN
       ,tr.Factory_User
       ,tr.DateStart
       ,tr.DateEnd
       ,tr.StayNight
       ,tr.AssisstedIDs
       ,tr.TravelerID
       ,tr.TravelerOrder
       ,dl.Dept
       ,COUNT(*) OVER()          AS TotalRow
FROM    #Travelers tr
        LEFT JOIN #Itinerary it
               ON  it.DOC_NBR       = tr.DOC_NBR
               AND it.TravelerID    = tr.TravelerID
               AND it.TravelerOrder = tr.TravelerOrder
        LEFT JOIN #DeptLookup dl
               ON  ISNULL(dl.Factory_User ,'') = ISNULL(tr.Factory_User ,'')
               AND ISNULL(dl.Departure ,'')    = ISNULL(tr.Departure ,'')
               AND dl.TravelerID                = tr.TravelerID
ORDER BY ${safeSortField} ${safeSortOrder} ,tr.DOC_NBR ,tr.TravelerOrder;

DROP TABLE #Travelers, #RouteNodes, #AccomSeq, #StopSeq, #Legs, #Itinerary, #DeptDistinct, #DeptLookup;`;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements: finalReplacements,
    })) as Record<string, any>[];

    const filteredRows = checkedDormShuttle
      ? rawRows.filter(
          (row) =>
            this.hasCompanyShuttleCar(row.Routes) &&
            this.hasDormAccommodation(row.Accommodation),
        )
      : rawRows;
    const total = filteredRows.length;

    const transformed = filteredRows.flatMap((row) =>
      this.expandRowsByAssistedIds(row).flatMap((expandedRow) => {
        const documentNumberBase = expandedRow.DOC_NBR ?? '';
        const tripFields = {
          Application_Day: expandedRow.CreatedAt
            ? dayjs(expandedRow.CreatedAt).format('YYYY-MM-DD')
            : '',
          Document_Number_Base: documentNumberBase,
          Staff_ID: expandedRow.EffectiveStaffID ?? '',
          Dept: expandedRow.Dept ?? expandedRow.Department ?? '',
          Round_trip_One_way:
            this.formatTripType(expandedRow.TypeTravel) ?? '',
          Start_Time: expandedRow.DateStart
            ? dayjs(expandedRow.DateStart).format('YYYY-MM-DD')
            : '',
          End_Time: expandedRow.DateEnd
            ? dayjs(expandedRow.DateEnd).format('YYYY-MM-DD')
            : '',
          Business_Trip_Type:
            this.formatBusinessTripType(expandedRow.Factory) ?? '',
          Number_of_nights_stayed: this.getAccommodationNights(
            expandedRow.Accommodation,
          ),
        };

        const { places, transports } = this.extractPlacesAndTransports(
          expandedRow.Routes,
        );
        const legs = this.buildLegsFromPlaces(places, transports);

        if (legs.length === 0) {
          return [
            {
              ...tripFields,
              Document_Number: `${documentNumberBase}-1`,
              Departure: '',
              Destination: '',
              Transport: '',
              Leg_Index: 1,
              TotalRow: 0,
            },
          ];
        }

        return legs.map((leg, index) => ({
          ...tripFields,
          Document_Number: `${documentNumberBase}-${index + 1}`,
          Departure: leg.dep,
          Destination: leg.dest,
          Transport: leg.transType,
          Leg_Index: index + 1,
          TotalRow: 0,
        }));
      }),
    );

    const withNightsStayed = checkedDormShuttle
      ? transformed
      : transformed.filter(
          (row) => Number(row.Number_of_nights_stayed ?? 0) > 0,
        );

    const transformedTotal = withNightsStayed.length;
    withNightsStayed.forEach((row) => {
      row.TotalRow = transformedTotal;
    });

    withNightsStayed.sort((left, right) =>
      this.compareValues(left[sortField], right[sortField], sortOrder),
    );

    const offset = (page - 1) * limit;
    const data = withNightsStayed.slice(offset, offset + limit);

    return {
      data,
      page,
      limit,
      total: Number(transformedTotal),
      hasMore: offset + data.length < Number(transformedTotal),
    };
  }
}
