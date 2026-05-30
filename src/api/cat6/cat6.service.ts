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

type RouteItem = {
  AddressDetail?: string;
  AddressName?: string;
  Transport?: string;
  From?: string;
  To?: string;
  isAirport?: boolean;
  IsAirport?: boolean;
};

type AccommodationItem = {
  nights?: number | string;
  type?: string;
};

type Cat6ActiveRow = {
  Document_Date: string;
  Document_Number: string;
  Staff_ID: string;
  Dept: string;
  Round_trip_One_way: string;
  Start_Time: string;
  End_Time: string;
  Business_Trip_Type: string;
  Number_of_nights_stayed: number;
  TotalRow?: number;
  [key: string]: string | number | undefined;
};

type CMSLeg = {
  transType: string;
  dep: string;
  dest: string;
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

  private formatPlacesAndTransports(routesValue: unknown) {
    const routes = this.parseJsonArray<RouteItem>(routesValue);
    const places: string[] = [];
    const transports: string[] = [];

    const pushPlace = (value?: string) => {
      const place = getCat6AirportName(value);
      if (!place) return;
      const lastPlace = places[places.length - 1];
      if (lastPlace === place) return;
      places.push(place);
    };

    const isFlightRoute = (route: RouteItem) => {
      const transport = route.Transport?.trim().toLowerCase() ?? '';
      return (
        route.isAirport === true ||
        route.IsAirport === true ||
        transport === 'flight'
      );
    };

    let index = 0;
    while (index < routes.length) {
      const route = routes[index];
      const transport = route.Transport?.trim() ?? '';
      const nextRoute = routes[index + 1];

      if (isFlightRoute(route)) {
        const flightStart = route.From?.trim() ?? '';
        let flightEnd = route.To?.trim() ?? '';
        let cursor = index;

        while (
          cursor + 1 < routes.length &&
          isFlightRoute(routes[cursor + 1])
        ) {
          cursor += 1;
          flightEnd = routes[cursor].To?.trim() ?? flightEnd;
        }

        pushPlace(flightStart);
        pushPlace(flightEnd);
        transports.push('Flight');
        index = cursor + 1;
        continue;
      }

      if (
        !transport &&
        places.length === 0 &&
        nextRoute &&
        isFlightRoute(nextRoute)
      ) {
        pushPlace(nextRoute.From);
        index += 1;
        continue;
      }

      pushPlace(route.AddressDetail ?? route.AddressName);
      if (transport) {
        transports.push(transport);
      }
      index += 1;
    }

    const transportLimit = Math.max(places.length - 1, 0);
    const normalizedTransports = transports.slice(0, transportLimit);
    while (normalizedTransports.length < transportLimit) {
      normalizedTransports.push('');
    }

    return {
      ...places.reduce<Record<string, string>>((acc, place, index) => {
        acc[`Place${index + 1}`] = place;
        return acc;
      }, {}),
      ...normalizedTransports.reduce<Record<string, string>>(
        (acc, transport, index) => {
          acc[`Transport_${index + 1}`] = transport;
          return acc;
        },
        {},
      ),
    };
  }

  private getAccommodationNights(accommodationValue: unknown) {
    const accommodations =
      this.parseJsonArray<AccommodationItem>(accommodationValue);
    return accommodations.reduce((sum, item) => {
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
    const routes = this.parseJsonArray<RouteItem>(routesValue);
    return routes.some(
      (route) =>
        route.Transport?.trim().toLowerCase() === 'company shuttle car',
    );
  }

  private expandRowsByAssistedIds(row: Record<string, any>): Record<string, any>[] {
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

  private extractPlaces(row: Cat6ActiveRow): string[] {
    return Object.keys(row)
      .filter((key) => /^Place\d+$/.test(key))
      .sort((left, right) => {
        const leftIndex = Number(left.replace('Place', ''));
        const rightIndex = Number(right.replace('Place', ''));
        return leftIndex - rightIndex;
      })
      .map((key) => String(row[key] ?? '').trim())
      .filter(Boolean);
  }

  private extractTransports(row: Cat6ActiveRow): string[] {
    return Object.keys(row)
      .filter((key) => /^Transport_\d+$/.test(key))
      .sort((left, right) => {
        const leftIndex = Number(left.replace('Transport_', ''));
        const rightIndex = Number(right.replace('Transport_', ''));
        return leftIndex - rightIndex;
      })
      .map((key) => String(row[key] ?? '').trim());
  }

  private buildCMSLegs(row: Cat6ActiveRow): CMSLeg[] {
    const places = this.extractPlaces(row);
    const transports = this.extractTransports(row);
    const legs: CMSLeg[] = [];
    const maxLegs = Math.max(places.length - 1, 0);

    for (let index = 0; index < maxLegs; index += 1) {
      const dep = places[index]?.trim() ?? '';
      const dest = places[index + 1]?.trim() ?? '';
      const transType = transports[index]?.trim() ?? '';

      if (!dep || !dest || !transType) {
        continue;
      }

      legs.push({
        transType,
        dep,
        dest,
      });
    }

    return legs;
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
  ): CMSPayload[] {
    const legs = this.buildCMSLegs(row);
    const docKey = this.getCMSDocKey(row.Business_Trip_Type ?? '');

    return legs.map((leg, index) => ({
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
      DocDate: row.Document_Date
        ? dayjs(row.Document_Date).format('YYYY/MM/DD')
        : '',
      DocDate2: row.Start_Time
        ? dayjs(row.Start_Time).format('YYYY/MM/DD')
        : '',
      DocNo: row.Document_Number ?? '',
      UndDocNo: `${row.Document_Number ?? ''}-${index + 1}`,
      TransType: this.mapCMSTransType(leg.transType),
      Departure: leg.dep,
      Destination: leg.dest,
      Memo: row.Business_Trip_Type ?? '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    }));
  }

  async transformCMSData(
    rows: Cat6ActiveRow[],
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): Promise<CMSPayload[]> {
    return rows.flatMap((row) =>
      this.mapRowToCMSPayload(row, factory, dateFrom, dateTo),
    );
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

    let where = `WHERE  chb.BPMStatus = 'F'
                        AND ISNULL(chb.Factory_User ,'')<>''`;

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,chb.CreatedAt ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory.trim().toLowerCase() !== 'ALL'.trim().toLowerCase()) {
      where += ` AND chb.Factory_User LIKE ?`;
      replacements.push(`%${factory}%`);
    }

    const allowedSortFields: Record<string, string> = {
      Document_Date: 'CreatedAt',
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
    const query = `
                    SELECT chb.*
                          --,twt.Current_doc
                          ,COALESCE(
                              vwd.GROUP_NAME
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
                                  WHERE  teu2.ACCOUNT = chb.Factory_User+chb.UserCreate
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
                                  WHERE  teu3.ACCOUNT = chb.Departure+chb.UserCreate
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
                                  WHERE  cfe.EmpID = chb.UserCreate
                                          AND vwd4.GROUP_NAME IS NOT NULL
                              )
                          )                         AS Dept
                          ,COUNT(chb.TripID) OVER()  AS TotalRow
                    FROM   CDS_HRBUSS_BusTripData    AS chb
                          LEFT JOIN TB_EB_USER      AS teu
                                ON  (
                                        teu.ACCOUNT LIKE chb.Factory_User+'[0-9][0-9][0-9][0-9][0-9]'
                                        AND teu.ACCOUNT LIKE '%'+chb.UserCreate+'%'
                                    )
                                    OR (
                                        teu.ACCOUNT NOT LIKE chb.Factory_User+'[0-9][0-9][0-9][0-9][0-9]'
                                        AND teu.ACCOUNT=chb.UserCreate
                                    )
                          OUTER APPLY (
                        SELECT teed2.GROUP_ID
                        FROM   TB_EB_EMPL_DEP AS teed2
                        WHERE  teed2.USER_GUID = teu.USER_GUID
                              AND teed2.ORDERS = 0
                    )                                AS teed
                    LEFT JOIN vwDepartment_Factory   AS vwd
                                ON  vwd.GROUP_ID = teed.GROUP_ID
                          LEFT JOIN tb_wkf_task     AS twt
                                ON  twt.DOC_NBR = chb.DOC_NBR
                                    AND (
                                            twt.DOC_NBR LIKE 'LYV-HR-BT%'
                                            OR twt.DOC_NBR LIKE 'LHG-SUGG%'
                                            OR twt.DOC_NBR LIKE 'LVL-HR-BTF%'
                                            OR twt.DOC_NBR LIKE 'LVL-ODBT%'
                                            OR twt.DOC_NBR LIKE 'LYM-HR-BT%'
                                        )
      ${where}
      ORDER BY ${safeSortField} ${safeSortOrder}
    `;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
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
      this.expandRowsByAssistedIds(row).map((expandedRow) => ({
        Document_Date: expandedRow.CreatedAt
          ? dayjs(expandedRow.CreatedAt).format('YYYY-MM-DD')
          : '',
        Document_Number: expandedRow.DOC_NBR ?? '',
        Staff_ID: expandedRow.EffectiveStaffID ?? '',
        Dept: expandedRow.Dept ?? expandedRow.Department ?? '',
        Round_trip_One_way: this.formatTripType(expandedRow.TypeTravel) ?? '',
        Start_Time: expandedRow.DateStart
          ? dayjs(expandedRow.DateStart).format('YYYY-MM-DD')
          : '',
        End_Time: expandedRow.DateEnd
          ? dayjs(expandedRow.DateEnd).format('YYYY-MM-DD')
          : '',
        Business_Trip_Type: this.formatBusinessTripType(expandedRow.Factory) ?? '',
        ...this.formatPlacesAndTransports(expandedRow.Routes),
        Number_of_nights_stayed: this.getAccommodationNights(
          expandedRow.Accommodation,
        ),
        TotalRow: 0,
      })),
    );

    const transformedTotal = transformed.length;
    transformed.forEach((row) => {
      row.TotalRow = transformedTotal;
    });

    transformed.sort((left, right) =>
      this.compareValues(left[sortField], right[sortField], sortOrder),
    );

    const offset = (page - 1) * limit;
    const data = transformed.slice(offset, offset + limit);

    return {
      data,
      page,
      limit,
      total: Number(transformedTotal),
      hasMore: offset + data.length < Number(transformedTotal),
    };
  }
}
