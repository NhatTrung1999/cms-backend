import { Inject, Injectable } from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import dayjs from 'dayjs';
dayjs().format();

export interface AutoCMSCat6Data {
  DocumentDate: string;
  DocumentNumber: string;
  StaffID: number | string;
  Dept: string;
  TripType: string;
  StartTime: string;
  EndTime: string;
  BusinessTripType: string;
  PlaceOfDeparture: string;
  DepartureAirport: string;
  LandTransportDistanceA: string;
  LandTrasportationTypeA: string;
  DestinationAirport: string;
  ThirdCountryTransferDestination: string;
  LandTransportDistanceB: string;
  LandTrasportationTypeB: string;
  Destination2: string;
  Destination3: string;
  Destination4: string;
  Destination5: string;
  Destination6: string;
  LandTransportDistance: string;
  LandTrasportationType: string;
  AirTransportDistance: string;
  NumberOfNightsStayed: number;
}

function getFactory(factory: string): string {
  return factory ?? '';
}

type Cat6RouteItem = {
  AddressName: string;
  Transport: string;
  AddressDetail: string;
  isAirport: boolean;
  From: string;
  To: string;
};

type AccommodationItem = {
  id: string;
  isSameAsAbove: boolean;
  type: string;
  nights: number;
  addressID: string;
  address?: string;
};
type Cat6ExcelRow = {
  Document_Date: string | null;
  Document_Number: string | null;
  Staff_ID: string | null;
  Round_trip_One_way: string | null;
  Start_Time: string | null;
  End_Time: string | null;
  Business_Trip_Type: string | null;
  Place_of_Departure: string | null;
  Departure_Airport: string | null;
  Land_Transport_Distance_km_A: string;
  Land_Trasportation_Type_A: string | null;
  Destination_Airport: string | null;
  Third_country_transfer_Destination: string | null;
  Land_Transport_Distance_km_B: string;
  Land_Transportation_Type_B: string | null;
  Destination_2: string | null;
  Destination_3: string | null;
  Destination_4: string | null;
  Destination_5: string | null;
  Destination_6: string | null;
  Land_Transport_Distance_km: string;
  Land_Transportation_Type: string | null;
  Air_Transport_Distance_km: string;
  Number_of_nights_stayed: number;
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

  private normalizeRoute(item: unknown): Cat6RouteItem {
    const r = (item ?? {}) as Record<string, any>;
    const isAirport = Boolean(r.isAirport ?? r.IsAirport);
    const transport = typeof r.Transport === 'string' ? r.Transport : '';
    return {
      AddressName: r.AddressName ?? '',
      Transport: transport,
      AddressDetail: r.AddressDetail ?? '',
      isAirport: isAirport || transport.trim().toLowerCase() === 'flight',
      From: r.From ?? '',
      To: r.To ?? '',
    };
  }
  private splitRoutesByFlight(routes: Cat6RouteItem[]): Cat6RouteItem[][] {
    const flightIndices = routes
      .map((r, i) =>
        r.isAirport || r.Transport.trim().toLowerCase() === 'flight' ? i : -1,
      )
      .filter((i) => i >= 0);

    if (flightIndices.length <= 1) return [routes];

    const splitPoints = flightIndices.slice(1).map((idx) => idx - 1);
    const boundaries = [0, ...splitPoints, routes.length - 1];

    const groups: Cat6RouteItem[][] = [];
    for (let i = 0; i < boundaries.length - 1; i++) {
      groups.push(routes.slice(boundaries[i], boundaries[i + 1] + 1));
    }

    return groups;
  }

  private transformRow(row: Record<string, any>): Cat6ExcelRow {
    const routes = this.parseJsonArray(row.Routes)
      .flat()
      .map((item) => this.normalizeRoute(item));

    const accommodation = this.parseJsonArray<AccommodationItem>(
      row.Accommodation,
    );
    const flights = routes.filter(
      (r) =>
        r.isAirport === true || r.Transport.trim().toLowerCase() === 'flight',
    );

    const firstFlight = flights[0] ?? null;
    const lastFlight = flights[flights.length - 1] ?? null;
    const lastFlightIdx = routes
      .map((r) => r.isAirport || r.Transport.trim().toLowerCase() === 'flight')
      .lastIndexOf(true);
    const totalNights =
      Number(row.StayNight) ||
      accommodation.reduce((sum, a) => sum + (a.nights || 0), 0) ||
      0;

    return {
      Document_Date: row.CreatedAt
        ? dayjs(row.CreatedAt).format('YYYY-MM-DD')
        : null,
      Document_Number: row.DOC_NBR ?? null,
      Staff_ID: row.UserCreate ?? null,
      Round_trip_One_way: this.formatTripType(row.TypeTravel),
      Start_Time: row.DateStart
        ? dayjs(row.DateStart).format('YYYY-MM-DD')
        : null,
      End_Time: row.DateEnd ? dayjs(row.DateEnd).format('YYYY-MM-DD') : null,
      Business_Trip_Type: this.formatBusinessTripType(row.Factory),
      Place_of_Departure: routes[0]?.AddressDetail ?? null,
      Departure_Airport: firstFlight?.From ?? null,
      Land_Transport_Distance_km_A: 'API Calculation',
      Land_Trasportation_Type_A: routes[0]?.Transport ?? null,
      Destination_Airport: lastFlight?.To ?? null,
      Third_country_transfer_Destination:
        lastFlightIdx >= 0
          ? (routes[lastFlightIdx + 1]?.AddressDetail ?? null)
          : (routes[1]?.AddressDetail ?? null),
      Land_Transport_Distance_km_B: 'API Calculation',
      Land_Transportation_Type_B:
        lastFlightIdx >= 0
          ? (routes[lastFlightIdx + 1]?.Transport ?? null)
          : (routes[1]?.Transport ?? null),
      ...(() => {
        const hotelAddresses: (string | null)[] = accommodation
          .filter((a) => a.type?.toLowerCase() === 'hotel' && a.address)
          .map((a) => a.address ?? null);
        while (hotelAddresses.length < 5) hotelAddresses.push(null);
        return {
          Destination_2: hotelAddresses[0],
          Destination_3: hotelAddresses[1],
          Destination_4: hotelAddresses[2],
          Destination_5: hotelAddresses[3],
          Destination_6: hotelAddresses[4],
        };
      })(),
      Land_Transport_Distance_km: 'API Calculation',
      Land_Transportation_Type:
        lastFlightIdx >= 0
          ? (routes[lastFlightIdx + 1]?.Transport ?? null)
          : (routes[1]?.Transport ?? null),
      Air_Transport_Distance_km: 'API Calculation',
      Number_of_nights_stayed: totalNights,
    };
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
      outside_domestic: 'Domestic business trip to third-party entities',
      factory_oversea: 'Overseas business trip within the group',
      outside_oversea: 'Overseas business trip to third-party entities',
    };
    return map[type.toLowerCase()] ?? type;
  }
  private buildAccommodationMap(
    fullRoutes: Cat6RouteItem[],
    accommodation: AccommodationItem[],
  ): Map<string, AccommodationItem> {
    const stopsWithAccommodation = fullRoutes
      .slice(1)
      .filter((r) => !r.isAirport && r.Transport.toLowerCase() !== 'flight');

    const map = new Map<string, AccommodationItem>();
    stopsWithAccommodation.forEach((stop, i) => {
      if (accommodation[i] && stop.AddressDetail) {
        map.set(stop.AddressDetail.trim(), accommodation[i]);
      }
    });
    return map;
  }
  private getAccommodationForGroup(
    groupRoutes: Cat6RouteItem[],
    accMap: Map<string, AccommodationItem>,
  ): AccommodationItem[] {
    return groupRoutes
      .slice(1)
      .filter((r) => !r.isAirport && r.Transport.toLowerCase() !== 'flight')
      .map((r) => accMap.get(r.AddressDetail?.trim() ?? ''))
      .filter((a): a is AccommodationItem => !!a);
  }

  async getDataCat6(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'CreatedAt',
    sortOrder: string = 'asc',
  ) {
    const replacements: any[] = [];

    let where = `WHERE 1=1
                  AND BPMStatus = 'F'
                  AND ISNULL(Accommodation, '') <> ''
                  AND ISNULL(Factory_User, '') <> ''`;

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ` AND Factory_User LIKE ?`;
      replacements.push(`%${factory}%`);
    }
    const allowedSortFields: Record<string, string> = {
      Document_Date: 'CreatedAt',
      Document_Number: 'DOC_NBR',
      Staff_ID: 'UserCreate',
      Start_Time: 'DateStart',
      End_Time: 'DateEnd',
      Business_Trip_Type: 'Factory',
      Place_of_Departure: 'Departure',
      Number_of_nights_stayed: 'StayNight',
      CreatedAt: 'CreatedAt',
    };

    const safeSortField = allowedSortFields[sortField] ?? 'CreatedAt';
    const safeSortOrder = sortOrder?.toLowerCase() === 'desc' ? 'DESC' : 'ASC';
    const query = `
      SELECT *,
        DATEDIFF(day, DateStart, DateEnd) AS DuringDay,
        COUNT(TripID) OVER() AS TotalRow
      FROM CDS_HRBUSS_BusTripData
      ${where}
      ORDER BY ${safeSortField} ${safeSortOrder}
    `;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as Record<string, any>[];
    const transformed = rawRows.flatMap((row) => {
      const routes = this.parseJsonArray(row.Routes)
        .flat()
        .map((item) => this.normalizeRoute(item));

      const accommodation = this.parseJsonArray<AccommodationItem>(
        row.Accommodation,
      );
      const accMap = this.buildAccommodationMap(routes, accommodation);

      const routeGroups = this.splitRoutesByFlight(routes);

      return routeGroups.map((groupRoutes) => {
        const groupAccommodation = this.getAccommodationForGroup(
          groupRoutes,
          accMap,
        );
        return this.transformRow({
          ...row,
          Routes: groupRoutes,
          Accommodation: groupAccommodation,
        });
      });
    });
    const total = rawRows[0]?.TotalRow ?? transformed.length;
    const offset = (page - 1) * limit;
    const data = transformed.slice(offset, offset + limit);

    return {
      data,
      page,
      limit,
      total: Number(total),
      hasMore: offset + data.length < Number(total),
    };
  }

  async transformData(
    data: AutoCMSCat6Data[],
    factory: string,
    dateFrom: string,
    dateTo: string,
  ) {
    return data.flatMap((item) => {
      const routes = this.buildRoutes(item);

      return routes.map((route, index) => {
        let transType = '';
        switch (route.transType.trim().toLowerCase()) {
          case 'car':
            transType = '計程車/出租車';
            break;
          case 'company shuttle car':
            transType = '公司車';
            break;
          case 'flight':
            transType = '飛機';
            break;
          default:
            transType = route.transType.trim();
            break;
        }

        let docKey = '';
        switch (item.BusinessTripType.trim().toLowerCase()) {
          case 'domestic business trip within the group':
            docKey = '3.5.1';
            break;
          case 'domestic business trip to third-party entities':
            docKey = '3.5.2';
            break;
          case 'overseas business trip within the group':
            docKey = '3.5.3';
            break;
          case 'overseas business trip to third-party entities':
            docKey = '3.5.4';
            break;
        }

        return {
          System: 'BPM',
          Corporation: 'LAI YIH',
          Factory: getFactory(factory),
          Department: item.Dept,
          DocKey: docKey,
          ActivitySource: '',
          SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
          EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
          ActivityType: '3.5',
          DataType: '2',
          DocType: '洽公單',
          DocDate: dayjs(item.DocumentDate).format('YYYY/MM/DD'),
          DocDate2: dayjs(item.StartTime).format('YYYY/MM/DD'),
          DocNo: item.DocumentNumber,
          UndDocNo: `${item.DocumentNumber}-${index + 1}`,
          TransType: transType,
          Departure: route.dep,
          Destination: route.dest,
          Memo: item.BusinessTripType,
          CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
          Creator: '',
        };
      });
    });
  }

  private mapToAutoCMSFormat(row: Cat6ExcelRow): AutoCMSCat6Data {
    return {
      DocumentDate: row.Document_Date ?? '',
      DocumentNumber: row.Document_Number ?? '',
      StaffID: row.Staff_ID ?? '',
      Dept: '',
      TripType: row.Round_trip_One_way ?? '',
      StartTime: row.Start_Time ?? '',
      EndTime: row.End_Time ?? '',
      BusinessTripType: row.Business_Trip_Type ?? '',
      PlaceOfDeparture: row.Place_of_Departure ?? '',
      DepartureAirport: row.Departure_Airport ?? '',
      LandTransportDistanceA: row.Land_Transport_Distance_km_A ?? '',
      LandTrasportationTypeA: row.Land_Trasportation_Type_A ?? '',
      DestinationAirport: row.Destination_Airport ?? '',
      ThirdCountryTransferDestination:
        row.Third_country_transfer_Destination ?? '',
      LandTransportDistanceB: row.Land_Transport_Distance_km_B ?? '',
      LandTrasportationTypeB: row.Land_Transportation_Type_B ?? '',
      Destination2: row.Destination_2 ?? '',
      Destination3: row.Destination_3 ?? '',
      Destination4: row.Destination_4 ?? '',
      Destination5: row.Destination_5 ?? '',
      Destination6: row.Destination_6 ?? '',
      LandTransportDistance: row.Land_Transport_Distance_km ?? '',
      LandTrasportationType: row.Land_Transportation_Type ?? '',
      AirTransportDistance: row.Air_Transport_Distance_km ?? '',
      NumberOfNightsStayed: row.Number_of_nights_stayed ?? 0,
    };
  }

  private buildRoutes(item: AutoCMSCat6Data) {
    const routes: { transType: string; dep: string; dest: string }[] = [];

    if (item.PlaceOfDeparture) {
      if (item.DepartureAirport) {
        routes.push({
          transType: item.LandTrasportationTypeA,
          dep: item.PlaceOfDeparture,
          dest: item.DepartureAirport,
        });
      }

      if (item.DepartureAirport && item.DestinationAirport) {
        routes.push({
          transType: 'Flight',
          dep: item.DepartureAirport,
          dest: item.DestinationAirport,
        });
      }

      if (item.DestinationAirport && item.ThirdCountryTransferDestination) {
        routes.push({
          transType: item.LandTrasportationTypeB,
          dep: item.DestinationAirport,
          dest: item.ThirdCountryTransferDestination,
        });
      }

      if (!item.DepartureAirport && item.ThirdCountryTransferDestination) {
        routes.push({
          transType: item.LandTrasportationTypeB,
          dep: item.PlaceOfDeparture,
          dest: item.ThirdCountryTransferDestination,
        });
      }
    }

    return routes;
  }

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    const result = await this.getDataCat6(dateFrom, dateTo, factory, 1, 999999);
    const cmData = result.data.map((row) => this.mapToAutoCMSFormat(row));
    return this.transformData(cmData, factory, dateFrom, dateTo);
  }
}
