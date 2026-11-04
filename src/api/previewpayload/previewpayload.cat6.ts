import dayjs from 'dayjs';
import { getCat6AirportName } from 'src/helper/cat6-airport.helper';
import {
  expandCat6RowsByAssistedIds,
  formatCat6BusinessTripType,
  formatCat6TripType,
  getCat6AccommodationNights,
  getCat6CMSDocKey,
  mapCat6CMSTransType,
} from 'src/helper/cat6.helper';

// These two moved to helper/cat6.helper (also used by the Cat6 service); kept exported from here as before.
export { getCat6CMSDocKey, mapCat6CMSTransType };
import { getFactory } from 'src/helper/factory.helper';
import { parseJsonArray } from 'src/helper/json.helper';

type Cat6RouteItem = {
  AddressName: string;
  Transport: string;
  AddressDetail: string;
  isAirport: boolean;
  From: string;
  To: string;
};

type Cat6ActivePreviewRow = {
  Application_Day: string;
  Document_Number: string;
  Staff_ID: string;
  Dept: string;
  Round_trip_One_way: string;
  Start_Time: string;
  End_Time: string;
  Business_Trip_Type: string;
  Number_of_nights_stayed: number;
  Factory_Code: string;
  [key: string]: string | number | undefined;
};

type Cat6PreviewPayload = {
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

type Cat6AccommodationPreviewPayload = {
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

export const formatCat6PlacesAndTransports = (routesValue: unknown) => {
  const routes = parseJsonArray<Cat6RouteItem>(routesValue);
  const places: string[] = [];
  const transports: string[] = [];

  const pushPlace = (value?: string) => {
    const place = getCat6AirportName(value);
    if (!place) return;
    if (places[places.length - 1] === place) return;
    places.push(place);
  };

  const isFlightRoute = (route: Cat6RouteItem) => {
    const transport = route.Transport?.trim().toLowerCase() ?? '';
    return route.isAirport === true || transport === 'flight';
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
};

export const transformCat6ActiveRow = (
  row: Record<string, any>,
): Cat6ActivePreviewRow => {
  return {
    Application_Day: row.CreatedAt
      ? dayjs(row.CreatedAt).format('YYYY-MM-DD')
      : '',
    Document_Number: row.DOC_NBR ?? '',
    Staff_ID: row.EffectiveStaffID ?? '',
    Dept: row.Dept ?? row.Department ?? '',
    Round_trip_One_way: formatCat6TripType(row.TypeTravel) ?? '',
    Start_Time: row.DateStart ? dayjs(row.DateStart).format('YYYY-MM-DD') : '',
    End_Time: row.DateEnd ? dayjs(row.DateEnd).format('YYYY-MM-DD') : '',
    Business_Trip_Type: formatCat6BusinessTripType(row.Factory) ?? '',
    ...formatCat6PlacesAndTransports(row.Routes),
    Number_of_nights_stayed: getCat6AccommodationNights(row.Accommodation),
    Factory_Code: String(row.Factory_User ?? '')
      .trim()
      .toUpperCase(),
  };
};

export const extractCat6Places = (row: Cat6ActivePreviewRow): string[] => {
  return Object.keys(row)
    .filter((key) => /^Place\d+$/.test(key))
    .sort((left, right) => {
      const leftIndex = Number(left.replace('Place', ''));
      const rightIndex = Number(right.replace('Place', ''));
      return leftIndex - rightIndex;
    })
    .map((key) => String(row[key] ?? '').trim())
    .filter(Boolean);
};

export const extractCat6Transports = (row: Cat6ActivePreviewRow): string[] => {
  return Object.keys(row)
    .filter((key) => /^Transport_\d+$/.test(key))
    .sort((left, right) => {
      const leftIndex = Number(left.replace('Transport_', ''));
      const rightIndex = Number(right.replace('Transport_', ''));
      return leftIndex - rightIndex;
    })
    .map((key) => String(row[key] ?? '').trim());
};

export const buildCat6CMSLegs = (row: Cat6ActivePreviewRow) => {
  const places = extractCat6Places(row);
  const transports = extractCat6Transports(row);
  const legs: { transType: string; dep: string; dest: string }[] = [];
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
};

export const mapCat6RowToPreviewPayload = (
  row: Cat6ActivePreviewRow,
  factory: string,
  dateFrom: string,
  dateTo: string,
): Cat6PreviewPayload[] => {
  const legs = buildCat6CMSLegs(row);
  const docKey = getCat6CMSDocKey(row.Business_Trip_Type ?? '');
  const resolvedFactoryCode =
    factory.trim().toUpperCase() === 'ALL'
      ? row.Factory_Code
      : factory.trim().toUpperCase();

  return legs.map((leg, index) => ({
    System: 'BPM',
    Corporation: 'LAI YIH',
    Factory: getFactory(resolvedFactoryCode),
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
    DocDate2: row.Start_Time ? dayjs(row.Start_Time).format('YYYY/MM/DD') : '',
    DocNo: row.Document_Number ?? '',
    UndDocNo: `${row.Document_Number ?? ''}-${index + 1}`,
    TransType: mapCat6CMSTransType(leg.transType),
    Departure: leg.dep,
    Destination: leg.dest,
    Memo: row.Business_Trip_Type ?? '',
    CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
    Creator: '',
  }));
};

export const mapCat6RowToAccommodationPreviewPayload = (
  row: Cat6ActivePreviewRow,
  factory: string,
  dateFrom: string,
  dateTo: string,
): Cat6AccommodationPreviewPayload => {
  const resolvedFactoryCode =
    factory.trim().toUpperCase() === 'ALL'
      ? row.Factory_Code
      : factory.trim().toUpperCase();

  return {
    System: 'BPM',
    Corporation: 'LAI YIH',
    Factory: getFactory(resolvedFactoryCode),
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
    DocDate2: row.Start_Time ? dayjs(row.Start_Time).format('YYYY/MM/DD') : '',
    DocNo: row.Document_Number ?? '',
    UndDocNo: row.Document_Number ?? '',
    TransType: 'double',
    ActivityData: Number(row.Number_of_nights_stayed ?? 0),
    Memo: row.Business_Trip_Type ?? '',
    CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
    Creator: '',
  };
};

/** Raw rows of the UOF business-trip query -> one preview row per traveler (assisted IDs expanded). */
export const toCat6ActiveRows = (rawRows: Record<string, any>[]) =>
  rawRows.flatMap((row) =>
    expandCat6RowsByAssistedIds(row).map((expandedRow) =>
      transformCat6ActiveRow(expandedRow),
    ),
  );
