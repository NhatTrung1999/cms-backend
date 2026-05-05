import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes, Sequelize } from 'sequelize';
import dayjs from 'dayjs';

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
      outside_domestic: 'Domestic business trip to third-party entities',
      factory_oversea: 'Overseas business trip within the group',
      outside_oversea: 'Overseas business trip to third-party entities',
    };
    return map[type.toLowerCase()] ?? type;
  }

  private formatPlacesAndTransports(routesValue: unknown) {
    const routes = this.parseJsonArray<RouteItem>(routesValue);
    const places: string[] = [];
    const transports: string[] = [];

    const pushPlace = (value?: string) => {
      const place = value?.trim() ?? '';
      if (!place) return;
      const lastPlace = places[places.length - 1];
      if (lastPlace === place) return;
      places.push(place);
    };

    const isFlightRoute = (route: RouteItem) => {
      const transport = route.Transport?.trim().toLowerCase() ?? '';
      return route.isAirport === true || route.IsAirport === true || transport === 'flight';
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

        while (cursor + 1 < routes.length && isFlightRoute(routes[cursor + 1])) {
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
      SELECT *,
        COUNT(TripID) OVER() AS TotalRow
      FROM CDS_HRBUSS_BusTripData
      ${where}
      ORDER BY ${safeSortField} ${safeSortOrder}
    `;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as Record<string, any>[];

    console.log(rawRows);
    const transformed = rawRows.map((row) => ({
      Document_Date: row.CreatedAt
        ? dayjs(row.CreatedAt).format('YYYY-MM-DD')
        : '',
      Document_Number: row.DOC_NBR ?? '',
      Staff_ID: row.UserCreate ?? '',
      Dept: row.Dept ?? row.Department ?? '',
      Round_trip_One_way: this.formatTripType(row.TypeTravel) ?? '',
      Start_Time: row.DateStart
        ? dayjs(row.DateStart).format('YYYY-MM-DD')
        : '',
      End_Time: row.DateEnd ? dayjs(row.DateEnd).format('YYYY-MM-DD') : '',
      Business_Trip_Type: this.formatBusinessTripType(row.Factory) ?? '',
      ...this.formatPlacesAndTransports(row.Routes),
      Number_of_nights_stayed: this.getAccommodationNights(row.Accommodation),
      TotalRow: Number(row.TotalRow ?? 0),
    }));

    transformed.sort((left, right) =>
      this.compareValues(left[sortField], right[sortField], sortOrder),
    );

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
}
