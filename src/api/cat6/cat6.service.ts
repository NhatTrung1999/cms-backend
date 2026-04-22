import { Inject, Injectable } from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import dayjs from 'dayjs';
import { getFactory } from 'src/helper/factory.helper';
import { AutoCMSCat6Data } from 'src/fakedata';
dayjs().format();

type Cat6UiRow = {
  Document_Date: string;
  Document_Number: string;
  Staff_ID: string;
  Round_trip_One_way: string;
  Start_Time: string;
  End_Time: string;
  Business_Trip_Type: string;
  Place_of_Departure: string;
  Land_Trasportation_Type_A: string;
  Land_Transport_Distance_km_A: string;
  Departure_Airport: string;
  Destination_Airport: string;
  Air_Transport_Distance_km: string;
  Third_country_transfer_Destination: string;
  Land_Transportation_Type_B: string;
  Land_Transport_Distance_km_B: string;
  Destination_2: string;
  Destination_3: string;
  Destination_4: string;
  Destination_5: string;
  Destination_6: string;
  Land_Transportation_Type: string;
  Land_Transport_Distance_km: string;
  Number_of_nights_stayed: number;
  TotalRow: number;
  Routes: ReturnType<Cat6Service['normalizeRoute']>[];
  Accommodation: ReturnType<Cat6Service['normalizeAccommodation']>[];
};

@Injectable()
export class Cat6Service {
  constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

  async getDataCat6(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Document_Date',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;
    let where = `WHERE 1=1 AND BPMStatus = 'F'
                  AND ISNULL(Accommodation ,'')<>''
                  AND ISNULL(Factory_User ,'')<>''`;
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

    const dataResults = await this.fetchRawCat6Rows(replacements, query);

    let data = dataResults.flatMap((item) => {
      const routes = this.parseJsonArray(item.Routes).map((route) =>
        this.normalizeRoute(route),
      );
      const accommodations = this.parseJsonArray(item.Accommodation).map(
        (accommodation) => this.normalizeAccommodation(accommodation),
      );

      return this.splitByFlight({
        ...item,
        Routes: routes,
        Accommodation: accommodations,
      }).map((segment) => this.formatUiRow(segment));
    });

    data.sort((a, b) => {
      const aValue = a[sortField];
      const bValue = b[sortField];
      if (sortOrder === 'asc') {
        return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
      } else {
        return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
      }
    });

    const total = data.length;
    data = data.slice(offset, offset + limit);

    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
  }

  private parseJsonArray<T>(value: string | null | undefined): T[] {
    if (!value) {
      return [];
    }

    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  private normalizeRoute(route: any) {
    return {
      AddressName: route?.AddressName ?? null,
      AddressDetail: route?.AddressDetail ?? null,
      Transport: route?.Transport ?? '',
      isAirport: Boolean(route?.isAirport),
      From: route?.From ?? null,
      To: route?.To ?? null,
    };
  }

  private normalizeAccommodation(item: any) {
    return {
      id: item?.id ?? '',
      isSameAsAbove: Boolean(item?.isSameAsAbove),
      type: item?.type ?? '',
      address: item?.address ?? '',
      nights: Number(item?.nights ?? 0),
    };
  }

  private formatUiRow(record: {
    Routes: ReturnType<Cat6Service['normalizeRoute']>[];
    Accommodation: ReturnType<Cat6Service['normalizeAccommodation']>[];
    [key: string]: any;
  }): Cat6UiRow {
    const routes = record.Routes || [];
    const flightSegment = routes.find((route) => route.isAirport);
    const startSegment = routes[0];
    const otherDestinations = routes.filter(
      (route) => route !== startSegment && !route.isAirport,
    );

    return {
      Document_Date: this.formatDate(record.CreatedAt),
      Document_Number: record.DOC_NBR || '',
      Staff_ID: record.UserCreate || '',
      Round_trip_One_way: this.formatRoundTrip(record.TypeTravel || ''),
      Start_Time: this.formatDate(record.DateStart),
      End_Time: this.formatDate(record.DateEnd),
      Business_Trip_Type: this.formatBusinessTripType(record.Factory || ''),
      Place_of_Departure: this.getAddress(startSegment),
      Land_Trasportation_Type_A: startSegment?.Transport?.trim() || '',
      Land_Transport_Distance_km_A: 'API Calculation',
      Departure_Airport: flightSegment?.From || '',
      Destination_Airport: flightSegment?.To || '',
      Air_Transport_Distance_km: flightSegment ? 'API Calculation' : '',
      Third_country_transfer_Destination: this.getAddress(otherDestinations[0]),
      Land_Transportation_Type_B: otherDestinations[0]?.Transport?.trim() || '',
      Land_Transport_Distance_km_B: otherDestinations[0]
        ? 'API Calculation'
        : '',
      Destination_2: this.getAddress(otherDestinations[1]),
      Destination_3: this.getAddress(otherDestinations[2]),
      Destination_4: this.getAddress(otherDestinations[3]),
      Destination_5: this.getAddress(otherDestinations[4]),
      Destination_6: this.getAddress(otherDestinations[5]),
      Land_Transportation_Type: otherDestinations[0]?.Transport?.trim() || '',
      Land_Transport_Distance_km: 'API Calculation',
      Number_of_nights_stayed: this.sumHotelNights(record.Accommodation),
      TotalRow: record.TotalRow || 0,
      Routes: routes,
      Accommodation: record.Accommodation,
    };
  }

  private formatDate(value: string | Date | null | undefined) {
    if (!value) {
      return '';
    }

    const formatted = dayjs(value);
    return formatted.isValid() ? formatted.format('YYYY/MM/DD') : '';
  }

  private formatBusinessTripType(factory: string) {
    switch (factory?.trim().toLowerCase()) {
      case 'factory_domestic':
        return 'Domestic business trip within the group';
      case 'outside_domestic':
        return 'Domestic business trip to third-party entities';
      case 'factory_oversea':
        return 'Overseas business trip within the group';
      case 'outside_oversea':
        return 'Overseas business trip to third-party entities';
      default:
        return '';
    }
  }

  private formatBusinessTripTypeNoDash(factory: string) {
    return this.formatBusinessTripType(factory).replace(
      /third-party/g,
      'third party',
    );
  }

  private normalizeBusinessTripTypeForDocKey(value: string) {
    return (value || '')
      .trim()
      .toLowerCase()
      .replace(/-/g, ' ')
      .replace(/\s+/g, ' ');
  }

  private getDocKeyFromBusinessTripType(value: string) {
    switch (this.normalizeBusinessTripTypeForDocKey(value)) {
      case 'factory domestic':
      case 'domestic business trip within the group':
        return '3.5.1';
      case 'outside domestic':
      case 'domestic business trip to third party entities':
        return '3.5.2';
      case 'factory oversea':
      case 'overseas business trip within the group':
        return '3.5.3';
      case 'outside oversea':
      case 'overseas business trip to third party entities':
        return '3.5.4';
      default:
        return '';
    }
  }

  private formatRoundTrip(typeTravel: string) {
    return typeTravel?.trim().toLowerCase() === 'round'
      ? 'Round trip'
      : 'One-way';
  }

  private getAddress(route?: ReturnType<Cat6Service['normalizeRoute']>) {
    return route?.AddressDetail || route?.AddressName || '';
  }

  private sumHotelNights(
    accommodations: ReturnType<Cat6Service['normalizeAccommodation']>[] = [],
  ) {
    return (
      accommodations
        .filter(
          (item) =>
            item.isSameAsAbove === false &&
            item.type?.trim().toLowerCase() === 'hotel',
        )
        .reduce((sum, acc) => sum + (acc.nights || 0), 0) || 0
    );
  }

  private splitByFlight(record: {
    Routes: ReturnType<Cat6Service['normalizeRoute']>[];
    Accommodation: ReturnType<Cat6Service['normalizeAccommodation']>[];
    [key: string]: any;
  }) {
    const routes = record.Routes || [];
    let remainingAccommodations = [...(record.Accommodation || [])];
    const result: (typeof record)[] = [];

    if (routes.length === 0) {
      return [
        {
          ...record,
          Routes: [],
        },
      ];
    }

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
            .filter((route) => !route.isAirport).length;
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

    result.push({
      ...record,
      Routes: routes.slice(startIndex),
      Accommodation: remainingAccommodations,
    });

    return result;
  }

  async transformData(
    data: AutoCMSCat6Data[],
    factory: string,
    dateFrom: string,
    dateTo: string,
  ) {
    return data.flatMap((item) => {
      const routes: { transType: string; dep: string; dest: string }[] = [];

      if (item.PlaceOfDeparture && item.DepartureAirport) {
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

      return routes.map((route, index) => {
        let transType = '';
        switch (route.transType.trim().toLowerCase()) {
          case 'Car'.trim().toLowerCase():
            transType = '計程車/出租車';
            break;
          case 'Company Shuttle Car'.trim().toLowerCase():
            transType = '公司車';
            break;
          case 'Flight'.trim().toLowerCase():
            transType = '飛機';
            break;
          default:
            transType = route.transType.trim();
            break;
        }

        const businessTripType = this.formatBusinessTripTypeNoDash(
          item.BusinessTripType,
        );
        return {
          System: 'BPM',
          Corporation: 'LAI YIH',
          Factory: getFactory(factory),
          Department: item.Dept,
          DocKey: this.getDocKeyFromBusinessTripType(businessTripType),
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
          Memo: businessTripType,
          CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
          Creator: '',
        };
      });
    });
  }

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    const data = await this.buildAutoCmsDataFromDb(dateFrom, dateTo, factory);
    return await this.transformData(data, factory, dateFrom, dateTo);
  }

  async autoSentCMSV2(dateFrom: string, dateTo: string, factory: string) {
    const data = await this.buildAutoCmsV2DataFromDb(dateFrom, dateTo, factory);

    const mapped = data.map((item) => {
      return {
        System: 'BPM',
        Corporation: 'LAI YIH',
        Factory: getFactory(factory),
        Department: item.Dept,
        DocKey: this.getDocKeyFromBusinessTripType(item.BusinessTripType),
        ActivitySource: '住宿',
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: '3.5',
        DataType: '3',
        DocType: '出差住宿單',
        DocDate: dayjs(item.DocumentDate).format('YYYY/MM/DD'),
        DocDate2: dayjs(item.StartTime).format('YYYY/MM/DD'),
        DocNo: item.DocumentNumber,
        UndDocNo: item.DocumentNumber,
        TransType: 'double',
        ActivityData: item.NumberOfNightsStayed,
        Memo: '',
        CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
        Creator: '',
      };
    });

    const grouped = new Map<string, (typeof mapped)[number]>();

    for (const item of mapped) {
      const key = (item.DocNo || '').trim();
      const existing = grouped.get(key);

      if (!existing) {
        grouped.set(key, item);
        continue;
      }

      existing.ActivityData =
        Number(existing.ActivityData || 0) + Number(item.ActivityData || 0);
    }

    return Array.from(grouped.values());
  }

  private async buildAutoCmsDataFromDb(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ): Promise<AutoCMSCat6Data[]> {
    const { query, replacements } = this.buildCat6Query(
      dateFrom,
      dateTo,
      factory,
    );
    const dataResults = await this.fetchRawCat6Rows(replacements, query);

    return dataResults.flatMap((item) => {
      const routes = this.parseJsonArray(item.Routes).map((route) =>
        this.normalizeRoute(route),
      );
      const accommodations = this.parseJsonArray(item.Accommodation).map(
        (accommodation) => this.normalizeAccommodation(accommodation),
      );

      return this.splitByFlight({
        ...item,
        Routes: routes,
        Accommodation: accommodations,
      }).map((segment) => this.mapRawRowToAutoCmsData(segment));
    });
  }

  private async buildAutoCmsV2DataFromDb(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ): Promise<AutoCMSCat6Data[]> {
    const { query, replacements } = this.buildCat6Query(
      dateFrom,
      dateTo,
      factory,
    );
    const dataResults = await this.fetchRawCat6Rows(replacements, query);

    const mergedRows = new Map<string, AutoCMSCat6Data>();
    let fallbackIndex = 0;

    for (const item of dataResults) {
      const documentNumber = (item.DOC_NBR || '').trim();
      const key = documentNumber || `__missing_doc_no__${fallbackIndex++}`;
      const nights = this.sumHotelNights(
        this.parseJsonArray(item.Accommodation).map((accommodation) =>
          this.normalizeAccommodation(accommodation),
        ),
      );

      const existing = mergedRows.get(key);
      if (existing) {
        existing.NumberOfNightsStayed += nights;
        continue;
      }

      mergedRows.set(
        key,
        this.mapRawRowToAutoCmsData({
          ...item,
          Routes: [],
          Accommodation: [],
        }),
      );
      const inserted = mergedRows.get(key);
      if (inserted) {
        inserted.NumberOfNightsStayed = nights;
      }
    }

    return Array.from(mergedRows.values());
  }

  private async fetchRawCat6Rows(
    replacements: any[],
    query: string,
  ): Promise<any[]> {
    return (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as any[];
  }

  private buildCat6Query(dateFrom: string, dateTo: string, factory: string) {
    let where = `WHERE 1=1 AND BPMStatus = 'F'
                  AND ISNULL(Accommodation ,'')<>'' 
                  AND ISNULL(Factory_User ,'')<>''`;
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

    return { query, replacements };
  }

  private mapRawRowToAutoCmsData(record: {
    Routes: ReturnType<Cat6Service['normalizeRoute']>[];
    Accommodation: ReturnType<Cat6Service['normalizeAccommodation']>[];
    [key: string]: any;
  }): AutoCMSCat6Data {
    const routes = record.Routes || [];
    const flightSegment = routes.find((route) => route.isAirport);
    const startSegment = routes[0];
    const otherDestinations = routes.filter(
      (route) => route !== startSegment && !route.isAirport,
    );

    return {
      DocumentDate: this.formatDate(record.CreatedAt),
      DocumentNumber: record.DOC_NBR || '',
      StaffID: record.UserCreate || '',
      Dept: record.Dept ?? '',
      TripType: this.formatRoundTrip(record.TypeTravel || ''),
      StartTime: this.formatDate(record.DateStart),
      EndTime: this.formatDate(record.DateEnd),
      BusinessTripType: this.formatBusinessTripType(record.Factory || ''),
      PlaceOfDeparture: this.getAddress(startSegment),
      DepartureAirport: flightSegment?.From || '',
      LandTransportDistanceA: 'API Auto calculate',
      LandTrasportationTypeA: startSegment?.Transport?.trim() || '',
      DestinationAirport: flightSegment?.To || '',
      ThirdCountryTransferDestination: this.getAddress(otherDestinations[0]),
      LandTransportDistanceB: otherDestinations[0] ? 'API Auto calculate' : '',
      LandTrasportationTypeB: otherDestinations[0]?.Transport?.trim() || '',
      Destination2: this.getAddress(otherDestinations[1]),
      Destination3: this.getAddress(otherDestinations[2]),
      Destination4: this.getAddress(otherDestinations[3]),
      Destination5: this.getAddress(otherDestinations[4]),
      Destination6: this.getAddress(otherDestinations[5]),
      LandTransportDistance: 'API Auto calculate',
      LandTrasportationType: otherDestinations[0]?.Transport?.trim() || '',
      AirTransportDistance: flightSegment ? 'API Auto calculate' : '',
      NumberOfNightsStayed: this.sumHotelNights(record.Accommodation),
    };
  }
}
