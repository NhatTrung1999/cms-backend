import { ICat6Record, Route } from 'src/types/cat6';

export const splitRecordByAirport = (record: ICat6Record): ICat6Record[] => {
  const routes = record.Routes;
  let remainingAccommodations = [...(record.Accommodation || [])];
  const result: ICat6Record[] = [];

  let startIndex = 0;
  let airportCount = 0;

  for (let i = 0; i < routes.length; i++) {
    if (routes[i].isAirport) {
      airportCount++;

      if (airportCount > 1) {
        const cutIndex = i - 1;

        const segmentRoutes = routes.slice(startIndex, cutIndex + 1);

        const destinationsCount = segmentRoutes
          .slice(1)
          .filter((r) => !r.isAirport).length;

        const segmentAccommodations = remainingAccommodations.slice(
          0,
          destinationsCount,
        );

        remainingAccommodations =
          remainingAccommodations.slice(destinationsCount);

        result.push({
          ...record,
          Routes: segmentRoutes,
          Accommodation: segmentAccommodations,
        });
        startIndex = cutIndex;
      }
    }
  }

  const lastRoutes = routes.slice(startIndex);

  result.push({
    ...record,
    Routes: lastRoutes,
    Accommodation: remainingAccommodations,
  });

  return result;
};

export const normalizeRoute = (route: any): Route => {
  return {
    AddressName: route.AddressName ?? null,
    AddressDetail: route.AddressDetail ?? null,
    Transport: route.Transport,
    isAirport: route.isAirport,
    From: route.From ?? null,
    To: route.To ?? null,
  };
};

export const buildQueryCat6 = async (
  dateFrom: string,
  dateTo: string,
  factory: string,
) => {
  let where = 'WHERE 1=1';
  const replacements: any[] = [];

  if (dateTo && dateFrom) {
    where += ` AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?`;
    replacements.push(dateFrom, dateTo);
  }

  if (factory) {
    where += ` AND Factory_User LIKE ?`;
    replacements.push(`%${factory}%`);
  }

  const query = `SELECT *, COUNT(TripID) OVER() AS TotalRow
                    FROM CDS_HRBUSS_BusTripData
                    ${where}`;
  return query;
};
