import { Inject, Injectable } from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import dayjs from 'dayjs';
dayjs().format();

// ─── Types ────────────────────────────────────────────────────────────────────

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
};

// Format map sang đúng cột Excel
type Cat6ExcelRow = {
  // A–H: Thông tin chung
  documentDate: string | null;
  documentNumber: string | null;
  staffID: string | null;
  roundTripOrOneWay: string | null;
  startTime: string | null;
  endTime: string | null;
  businessTripType: string | null;
  placeOfDeparture: string | null;
  // I–K: Điểm khởi hành
  departureAirport: string | null;
  landTransportDistanceA: number | null;
  landTransportTypeA: string | null;
  // L–O: Điểm đến
  destinationAirport: string | null;
  thirdCountryTransferDestination: string | null;
  landTransportDistanceB: number | null;
  landTransportTypeB: string | null;
  // P–T: Transit destinations
  transitDestination2: string | null;
  transitDestination3: string | null;
  transitDestination4: string | null;
  transitDestination5: string | null;
  transitDestination6: string | null;
  // U–X: Tổng hợp
  landTransportDistanceTotal: number | null;
  landTransportTypeTotal: string | null;
  airTransportDistance: number | null;
  numberOfNightsStayed: number;
};

// ─── Service ──────────────────────────────────────────────────────────────────

@Injectable()
export class Cat6Service {
  constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

  // ── Parse JSON field từ DB ────────────────────────────────────────────────

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
    // ✅ Handle cả 2 case: isAirport (lowercase) và IsAirport (uppercase)
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

  // ── Tách routes thành child rows theo số lượng flight ────────────────────
  // Mỗi flight → 1 child row
  // Land stop ngay trước flight tiếp theo được lặp lại ở 2 child liền kề
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

  // ── Transform 1 DB row → Excel row ───────────────────────────────────────

  private transformRow(row: Record<string, any>): Cat6ExcelRow {
    const routes = this.parseJsonArray(row.Routes)
      .flat()
      .map((item) => this.normalizeRoute(item));

    const accommodation = this.parseJsonArray<AccommodationItem>(
      row.Accommodation,
    );

    // Tách flights và land transports
    const flights = routes.filter(
      (r) =>
        r.isAirport === true || r.Transport.trim().toLowerCase() === 'flight',
    );
    const landLegs = routes.filter(
      (r) =>
        r.isAirport === false && r.Transport.trim().toLowerCase() !== 'flight',
    );

    // Index của tất cả các flight trong routes
    const flightIndices = routes
      .map((r, i) =>
        r.isAirport || r.Transport.trim().toLowerCase() === 'flight' ? i : -1,
      )
      .filter((i) => i >= 0);

    const firstFlight = flights[0] ?? null;
    const lastFlight = flights[flights.length - 1] ?? null;

    // Transit = các flight ở giữa (bỏ first và last)
    // Với 2 flights: RGN→SGN, SGN→HKG thì transit là SGN (From của flight thứ 2)
    const transitAirports =
      flights.length > 1
        ? flights
            .slice(1)
            .map((f) => f.From)
            .filter(Boolean)
        : [];

    // Land transport trước flight đầu (A)
    const firstFlightIdx = routes.findIndex(
      (r) => r.isAirport || r.Transport.trim().toLowerCase() === 'flight',
    );
    const landBeforeFlight = routes
      .slice(0, firstFlightIdx < 0 ? 0 : firstFlightIdx)
      .filter((r) => !r.isAirport && r.Transport.toLowerCase() !== 'flight');

    // Land transport sau flight cuối (B)
    const lastFlightIdx = routes
      .map((r) => r.isAirport || r.Transport.trim().toLowerCase() === 'flight')
      .lastIndexOf(true);
    const landAfterFlight = routes
      .slice(lastFlightIdx + 1)
      .filter((r) => !r.isAirport && r.Transport.toLowerCase() !== 'flight');

    // DuringDay: tính lại từ DateStart/DateEnd vì DB trả về 0
    const duringDay =
      row.DateStart && row.DateEnd
        ? Math.abs(
            Math.round(
              (new Date(row.DateEnd).getTime() -
                new Date(row.DateStart).getTime()) /
                (1000 * 60 * 60 * 24),
            ),
          )
        : Number(row.DuringDay) || 0;

    // Số đêm lưu trú
    const totalNights =
      Number(row.StayNight) ||
      accommodation.reduce((sum, a) => sum + (a.nights || 0), 0) ||
      0;

    // Pad transit destinations đủ 5 cột (Destination 2–6)
    const padded = [...transitAirports];
    while (padded.length < 5) padded.push(null as any);

    return {
      // A–H
      documentDate: row.CreatedAt
        ? dayjs(row.CreatedAt).format('YYYY-MM-DD')
        : null,
      documentNumber: row.DOC_NBR ?? null,
      staffID: row.UserCreate ?? null,
      roundTripOrOneWay: this.formatTripType(row.TypeTravel),
      startTime: row.DateStart
        ? dayjs(row.DateStart).format('YYYY-MM-DD')
        : null,
      endTime: row.DateEnd ? dayjs(row.DateEnd).format('YYYY-MM-DD') : null,
      businessTripType: row.Factory ?? null,
      // ✅ Lấy AddressDetail của điểm xuất phát đầu tiên
      placeOfDeparture: routes[0]?.AddressDetail ?? null,

      // I–K
      departureAirport: firstFlight?.From ?? null,
      landTransportDistanceA: row.Distance ?? null,
      // ✅ Lấy Transport của điểm xuất phát đầu tiên
      landTransportTypeA: routes[0]?.Transport ?? null,

      // L–O
      destinationAirport: lastFlight?.To ?? null,
      // Có flight: routes[lastFlightIdx + 1].AddressDetail
      // Không có flight: routes[1].AddressDetail
      thirdCountryTransferDestination:
        lastFlightIdx >= 0
          ? (routes[lastFlightIdx + 1]?.AddressDetail ?? null)
          : (routes[1]?.AddressDetail ?? null),
      landTransportDistanceB: row.Distance ?? null,
      // Có flight: routes[lastFlightIdx + 1].Transport
      // Không có flight: routes[1].Transport
      landTransportTypeB:
        lastFlightIdx >= 0
          ? (routes[lastFlightIdx + 1]?.Transport ?? null)
          : (routes[1]?.Transport ?? null),

      // P–T: AddressDetail các điểm kế tiếp sau thirdCountryTransferDestination
      // Có flight: bắt đầu từ routes[lastFlightIdx + 2]
      // Không có flight: bắt đầu từ routes[2]
      ...(() => {
        const startIdx = lastFlightIdx >= 0 ? lastFlightIdx + 2 : 2;
        const destinations: (string | null)[] = routes
          .slice(startIdx)
          .map((r) => r.AddressDetail ?? null);
        while (destinations.length < 5) destinations.push(null);
        return {
          transitDestination2: destinations[0],
          transitDestination3: destinations[1],
          transitDestination4: destinations[2],
          transitDestination5: destinations[3],
          transitDestination6: destinations[4],
        };
      })(),

      // U–X
      landTransportDistanceTotal: row.Distance ?? null,
      // ✅ Transport của điểm gần cuối (routes[length - 2])
      landTransportTypeTotal: routes[routes.length - 2]?.Transport ?? null,
      airTransportDistance: row.Distance ?? null,
      numberOfNightsStayed: totalNights,
    };
  }

  private formatTripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const map: Record<string, string> = {
      oneway: 'One-way',
      roundtrip: 'Round trip',
      'round trip': 'Round trip',
    };
    return map[type.toLowerCase()] ?? type;
  }

  // ─── Main method ──────────────────────────────────────────────────────────

  // ── Map accommodation theo thứ tự các điểm dừng ──────────────────────────
  // Bỏ routes[0] (điểm xuất phát) và các airport
  // accommodation[0] → stop[1], accommodation[1] → stop[2], ...
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

  // Lấy accommodations cho 1 child row (bỏ stop đầu tiên của child = điểm xuất phát)
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
                  --AND ISNULL(Accommodation, '') <> ''
                  --AND ISNULL(Factory_User, '') <> ''
                  AND DOC_NBR = 'LYV_HR_BT260100010'`;

    // ✅ Đã bỏ hardcode DOC_NBR

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ` AND Factory_User LIKE ?`;
      replacements.push(`%${factory}%`);
    }

    // ✅ Tính DuringDay ngay trong query
    const query = `
      SELECT *,
        DATEDIFF(day, DateStart, DateEnd) AS DuringDay,
        COUNT(TripID) OVER() AS TotalRow
      FROM CDS_HRBUSS_BusTripData
      ${where}
    `;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as Record<string, any>[];

    // Expand mỗi DB row thành nhiều child rows theo số lượng flight
    // rồi transform từng child sang Excel format
    const transformed = rawRows.flatMap((row) => {
      const routes = this.parseJsonArray(row.Routes)
        .flat()
        .map((item) => this.normalizeRoute(item));

      const accommodation = this.parseJsonArray<AccommodationItem>(
        row.Accommodation,
      );

      // Build map: AddressDetail → AccommodationItem theo thứ tự stops
      const accMap = this.buildAccommodationMap(routes, accommodation);

      const routeGroups = this.splitRoutesByFlight(routes);

      return routeGroups.map((groupRoutes) => {
        // Lấy đúng accommodation cho child này
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

    // Phân trang
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
