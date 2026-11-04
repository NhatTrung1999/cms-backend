import dayjs from 'dayjs';
import { getCat6AirportName } from './cat6-airport.helper';
import {
  appendCat6DocumentSuffix,
  getCat6AssistedIds,
} from './cat6-assisted.helper';
import { parseJsonArray } from './json.helper';
import { getOwn } from './object.helper';

type Cat6MergedLegItem = {
  LegNo?: number;
  From?: string;
  To?: string;
  Transport?: string;
  StayType?: string | null;
  Nights?: number | null;
};

type Cat6AccommodationItem = {
  id?: string;
  isSameAsAbove?: boolean;
  type?: string;
  nights?: number | string;
  addressID?: string;
};

export const formatCat6TripType = (type: unknown): string | null => {
  if (!type || typeof type !== 'string') return null;
  const t = type.toLowerCase().trim();
  if (t === 'round' || t === 'roundtrip' || t === 'round trip')
    return 'Round trip';
  return 'One-way';
};

export const formatCat6BusinessTripType = (type: unknown): string | null => {
  if (!type || typeof type !== 'string') return null;
  const map: Record<string, string> = {
    factory_domestic: 'Domestic business trip within the group',
    outside_domestic: 'Domestic business trip to third party entities',
    factory_oversea: 'Overseas business trip within the group',
    outside_oversea: 'Overseas business trip to third party entities',
  };
  return getOwn(map, type.toLowerCase()) ?? type;
};

export const extractCat6PlacesAndTransports = (routesValue: unknown) => {
  const legs = parseJsonArray<Cat6MergedLegItem>(routesValue)
    .slice()
    .sort((a, b) => (a.LegNo ?? 0) - (b.LegNo ?? 0));

  const places: string[] = [];
  const transports: string[] = [];

  const pushPlace = (value?: string) => {
    const place = getCat6AirportName(value);
    if (!place) return;
    if (places[places.length - 1] === place) return;
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
};

export const buildCat6LegsFromPlaces = (
  places: string[],
  transports: string[],
): { dep: string; dest: string; transType: string }[] => {
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
};

export const getCat6AccommodationNights = (accommodationValue: unknown) => {
  const accommodations =
    parseJsonArray<Cat6AccommodationItem>(accommodationValue);
  return accommodations.reduce((sum, item) => {
    if (item?.type?.trim().toLowerCase() !== 'hotel') return sum;
    const nights = Number(item?.nights ?? 0);
    return sum + (Number.isFinite(nights) ? nights : 0);
  }, 0);
};

export const expandCat6RowsByAssistedIds = (
  row: Record<string, any>,
): Record<string, any>[] => {
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
};

/** Turns the raw business-trip rows of the UOF query into the rows of the Cat6 Excel files. */
export const toCat6ExportRows = (rawRows: Record<string, any>[]) =>
  rawRows.flatMap((row) =>
    expandCat6RowsByAssistedIds(row).map((expandedRow) => ({
      Application_Day: expandedRow.CreatedAt
        ? dayjs(expandedRow.CreatedAt).format('YYYY-MM-DD')
        : '',
      Document_Number: expandedRow.DOC_NBR ?? '',
      Staff_ID: expandedRow.EffectiveStaffID ?? '',
      Dept: expandedRow.Dept ?? expandedRow.Department ?? '',
      Round_trip_One_way: formatCat6TripType(expandedRow.TypeTravel) ?? '',
      Start_Time: expandedRow.DateStart
        ? dayjs(expandedRow.DateStart).format('YYYY-MM-DD')
        : '',
      End_Time: expandedRow.DateEnd
        ? dayjs(expandedRow.DateEnd).format('YYYY-MM-DD')
        : '',
      Business_Trip_Type: formatCat6BusinessTripType(expandedRow.Factory) ?? '',
      Routes: expandedRow.Routes,
      Number_of_nights_stayed: getCat6AccommodationNights(
        expandedRow.Accommodation,
      ),
    })),
  );

export const getCat6CMSDocKey = (businessTripType: string): string => {
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
};

export const mapCat6CMSTransType = (transport: string): string => {
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
};

export const hasCat6DormAccommodation = (accommodationValue: unknown) => {
  const accommodations =
    parseJsonArray<Cat6AccommodationItem>(accommodationValue);
  return accommodations.some(
    (item) => item.type?.trim().toLowerCase() === 'dorm',
  );
};

export const hasCat6CompanyShuttleCar = (routesValue: unknown) => {
  const routes = parseJsonArray<Cat6MergedLegItem>(routesValue);
  return routes.some(
    (route) => route.Transport?.trim().toLowerCase() === 'company shuttle car',
  );
};

export const compareCat6Values = (
  left: unknown,
  right: unknown,
  sortOrder: string,
) => {
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
};
