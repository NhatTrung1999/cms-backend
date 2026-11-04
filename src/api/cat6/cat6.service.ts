import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes, Sequelize } from 'sequelize';
import dayjs from 'dayjs';
import { getFactory } from 'src/helper/factory.helper';
import {
  buildCat6LegsFromPlaces,
  compareCat6Values,
  expandCat6RowsByAssistedIds,
  extractCat6PlacesAndTransports,
  formatCat6BusinessTripType,
  formatCat6TripType,
  getCat6AccommodationNights,
  getCat6CMSDocKey,
  hasCat6CompanyShuttleCar,
  hasCat6DormAccommodation,
  mapCat6CMSTransType,
} from 'src/helper/cat6.helper';
import { buildCat6Filter, buildCat6TripQuery } from 'src/helper/cat6.queries';
import { getOwn } from 'src/helper/object.helper';

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
  /** Only set when loadCat6Rows is asked for it (the CMS payloads of the "ALL" factory need it). */
  Factory_Code?: string;
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

/** Factory code of a payload row: the factory chosen, or the factory of the row itself when 'ALL' is chosen. */
const factoryCodeOf = (row: Cat6ActiveRow, factory: string) => {
  const code = (factory ?? '').trim().toUpperCase();
  return code === 'ALL' ? (row.Factory_Code ?? '') : code;
};

@Injectable()
export class Cat6Service {
  constructor(@Inject('UOF') private readonly UOF: Sequelize) {}

  private mapRowToCMSPayload(
    row: Cat6ActiveRow,
    factory: string,
    dateFrom: string,
    dateTo: string,
  ): CMSPayload | null {
    if (!row.Departure || !row.Destination || !row.Transport) {
      return null;
    }

    const docKey = getCat6CMSDocKey(row.Business_Trip_Type ?? '');

    return {
      System: 'BPM',
      Corporation: 'LAI YIH',
      Factory: getFactory(factoryCodeOf(row, factory)),
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
      TransType: mapCat6CMSTransType(row.Transport),
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
    const result = await this.loadCat6Rows(
      dateFrom,
      dateTo,
      factory,
      1,
      999999,
      'CreatedAt',
      'asc',
      false,
      true,
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
      Factory: getFactory(factoryCodeOf(row, factory)),
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
      .filter(
        (payload): payload is CMSAccommodationPayload => payload !== null,
      );
  }

  async autoSentCMSAccommodation(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    const result = await this.loadCat6Rows(
      dateFrom,
      dateTo,
      factory,
      1,
      999999,
      'CreatedAt',
      'asc',
      false,
      true,
    );

    const rows = result.data as Cat6ActiveRow[];
    const uniqueTripRows = Array.from(
      new Map(rows.map((row) => [row.Document_Number_Base, row])).values(),
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
    return this.loadCat6Rows(
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
      checkedDormShuttle,
      false,
    );
  }

  /** The rows of getDataCat6; `withFactoryCode` adds the Factory_Code column (the factory of each row). */
  private async loadCat6Rows(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
    checkedDormShuttle: boolean,
    withFactoryCode: boolean,
  ) {
    const { where, replacements: finalReplacements } = buildCat6Filter(
      dateFrom,
      dateTo,
      factory.trim().toLowerCase() !== 'ALL'.trim().toLowerCase()
        ? factory
        : null,
    );

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

    const safeSortField = getOwn(allowedSortFields, sortField) ?? 'CreatedAt';
    const safeSortOrder = sortOrder?.toLowerCase() === 'desc' ? 'DESC' : 'ASC';
    const query = buildCat6TripQuery(
      where,
      `${safeSortField} ${safeSortOrder}`,
    );

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements: finalReplacements,
    })) as Record<string, any>[];

    const filteredRows = checkedDormShuttle
      ? rawRows.filter(
          (row) =>
            hasCat6CompanyShuttleCar(row.Routes) &&
            hasCat6DormAccommodation(row.Accommodation),
        )
      : rawRows;

    const transformed = filteredRows.flatMap((row) =>
      expandCat6RowsByAssistedIds(row).flatMap((expandedRow) => {
        const documentNumberBase = expandedRow.DOC_NBR ?? '';
        const tripFields = {
          Application_Day: expandedRow.CreatedAt
            ? dayjs(expandedRow.CreatedAt).format('YYYY-MM-DD')
            : '',
          Document_Number_Base: documentNumberBase,
          Staff_ID: expandedRow.EffectiveStaffID ?? '',
          Dept: expandedRow.Dept ?? expandedRow.Department ?? '',
          Round_trip_One_way: formatCat6TripType(expandedRow.TypeTravel) ?? '',
          Start_Time: expandedRow.DateStart
            ? dayjs(expandedRow.DateStart).format('YYYY-MM-DD')
            : '',
          End_Time: expandedRow.DateEnd
            ? dayjs(expandedRow.DateEnd).format('YYYY-MM-DD')
            : '',
          Business_Trip_Type:
            formatCat6BusinessTripType(expandedRow.Factory) ?? '',
          Number_of_nights_stayed: getCat6AccommodationNights(
            expandedRow.Accommodation,
          ),
          ...(withFactoryCode && {
            Factory_Code: String(expandedRow.Factory_User ?? '')
              .trim()
              .toUpperCase(),
          }),
        };

        const { places, transports } = extractCat6PlacesAndTransports(
          expandedRow.Routes,
        );
        const legs = buildCat6LegsFromPlaces(places, transports);

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
      compareCat6Values(left[sortField], right[sortField], sortOrder),
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
