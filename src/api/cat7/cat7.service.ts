import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQuery } from 'src/helper/cat7.helper';
import { buildReplacements } from 'src/helper/date.helper';
import { sortByField } from 'src/helper/sort.helper';
import { getOwn } from 'src/helper/object.helper';
import {
  buildQueryAutoSentCmsJAZ,
  buildQueryAutoSentCmsJZS,
  buildQueryAutoSentCmsLHG,
  buildQueryAutoSentCmsLVL,
  buildQueryAutoSentCmsLYM,
  buildQueryAutoSentCmsLYV,
} from 'src/helper/cat7.cms-queries';
import { buildQueryCustomExport } from 'src/helper/cat7.customexport';
import dayjs from 'dayjs';
import { ACTIVITY_TYPES, ActivityType, BuildQueryFn } from 'src/types/cat7';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';

type BuildSql = (
  factory: string,
  dateToExclusive: string,
) => { query: string; countQuery: string };

/** The columns of a CMS query row that mapToCMSFormat reads. */
type Cat7CMSRow = {
  Staff_ID?: string | null;
  Main_transportation_type?: string | null;
  Number_of_working_days?: string | number | null;
  Factory_address?: string | null;
  Residential_address?: string | null;
  Factory_Name?: string | null;
  Department_Name?: string | null;
};

type SortSql = {
  fieldMap: Record<string, string>;
  defaultField: string;
  marker: string;
  suffix: string;
};

@Injectable()
export class Cat7Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
  ) {}

  private get factoryDbMap(): Record<string, Sequelize> {
    return {
      LYV: this.LYV_HRIS,
      LHG: this.LHG_HRIS,
      LYM: this.LYM_HRIS,
      LVL: this.LVL_HRIS,
      JAZ: this.JAZ_HRIS,
      JZS: this.JZS_HRIS,
    };
  }

  private readonly buildQueryMap: Record<FactoryCode, BuildQueryFn> = {
    LYV: buildQueryAutoSentCmsLYV,
    LHG: buildQueryAutoSentCmsLHG,
    LVL: buildQueryAutoSentCmsLVL,
    LYM: buildQueryAutoSentCmsLYM,
    JAZ: buildQueryAutoSentCmsJAZ,
    JZS: buildQueryAutoSentCmsJZS,
  };

  private readonly sortFieldMap: Record<string, string> = {
    Staff_ID: 'u.userId',
    Residential_Address: 'u.Address_Live',
    Main_Transportation_Type: 'u.Vehicle',
    Number_of_Working_Days: 'e.Number_of_Working_Days',
  };

  /** How the sort chosen by the user is put into the SQL of a page query. */
  private get pageSort(): SortSql {
    return {
      fieldMap: this.sortFieldMap,
      defaultField: 'u.userId',
      marker: 'ORDER BY u.userId',
      suffix: '',
    };
  }

  private get customExportSort(): SortSql {
    return {
      fieldMap: this.sortFieldMapCustomExport,
      defaultField: 'a.userId',
      marker: 'ORDER BY a.userId OFFSET',
      suffix: ' OFFSET',
    };
  }

  // The "all factories" custom export replaces the FIRST 'ORDER BY a.userId' of the query (the one inside
  // ROW_NUMBER() OVER of its CTE), the single-factory one replaces the last one: kept as it was.
  private get customExportAllSort(): SortSql {
    return {
      fieldMap: this.sortFieldMapCustomExport,
      defaultField: 'a.userId',
      marker: 'ORDER BY a.userId',
      suffix: '',
    };
  }

  /** One page of one factory (sorting and paging happen in the SQL). */
  private async getPageOfFactory(
    db: Sequelize,
    factory: string,
    build: BuildSql,
    sort: SortSql,
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;
    const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);

    const { query, countQuery } = build(factory, dateToExclusive);

    const safeSortField = getOwn(sort.fieldMap, sortField) ?? sort.defaultField;
    const safeSortOrder = sortOrder === 'asc' ? 'ASC' : 'DESC';
    const sortedQuery = query.replace(
      sort.marker,
      `ORDER BY ${safeSortField} ${safeSortOrder}${sort.suffix}`,
    );

    const dataReplacements = hasDate
      ? [dateFrom, dateToExclusive, offset, limit]
      : [offset, limit];
    const countReplacements = hasDate ? [dateFrom, dateToExclusive] : [];

    const [dataResults, countResults] = await Promise.all([
      db.query(sortedQuery, {
        replacements: dataReplacements,
        type: QueryTypes.SELECT,
      }),
      db.query(countQuery, {
        replacements: countReplacements,
        type: QueryTypes.SELECT,
      }),
    ]);

    const total = (countResults[0] as { total: number })?.total || 0;
    const hasMore = offset + dataResults.length < total;

    return { data: dataResults, page, limit, total, hasMore };
  }

  /**
   * One page over every factory: all rows of each factory are read, merged, sorted and cut here.
   * `numberRows` renumbers the merged rows (column No) before sorting.
   */
  private async getAllFactoriesPage(
    build: BuildSql,
    sort: SortSql,
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
    numberRows: boolean,
  ) {
    try {
      const offset = (page - 1) * limit;
      const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);

      const safeSortField =
        getOwn(sort.fieldMap, sortField) ?? sort.defaultField;
      const safeSortOrder = sortOrder === 'asc' ? 'ASC' : 'DESC';
      const countReplacements = hasDate ? [dateFrom, dateToExclusive] : [];
      const dataReplacements = hasDate
        ? [dateFrom, dateToExclusive, 0, 999999]
        : [0, 999999];

      const results = await Promise.all(
        Object.entries(this.factoryDbMap).map(([factoryName, conn]) => {
          const { query, countQuery } = build(factoryName, dateToExclusive);

          const sortedQuery = query.replace(
            sort.marker,
            `ORDER BY ${safeSortField} ${safeSortOrder}${sort.suffix}`,
          );

          return Promise.all([
            conn.query(sortedQuery, {
              replacements: dataReplacements,
              type: QueryTypes.SELECT,
            }),
            conn.query(countQuery, {
              replacements: countReplacements,
              type: QueryTypes.SELECT,
            }),
          ]);
        }),
      );

      const mergedRows = results.flatMap(([data]) => data);
      const allData = numberRows
        ? mergedRows.map((item, index) => ({ ...item, No: index + 1 }))
        : mergedRows;

      sortByField(allData, sortField, sortOrder);

      const total = results.reduce(
        (sum, [, count]) => sum + ((count[0] as { total: number })?.total || 0),
        0,
      );
      const data = allData.slice(offset, offset + limit);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async getDataCat7(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Staff_ID',
    sortOrder: string = 'asc',
  ) {
    const db = getOwn(this.factoryDbMap, factory);
    const build = (factoryName: string, dateToExclusive: string) =>
      buildQuery(factoryName, dateFrom, dateToExclusive);

    if (!db) {
      return this.getAllFactoriesPage(
        build,
        this.pageSort,
        dateFrom,
        dateTo,
        page,
        limit,
        sortField,
        sortOrder,
        false,
      );
    }

    return this.getPageOfFactory(
      db,
      factory,
      build,
      this.pageSort,
      dateFrom,
      dateTo,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  private readonly sortFieldMapCustomExport: Record<string, string> = {
    No: 'a.userId',
    ID: 'a.userId',
    Department: 'c.Department_Name',
    Full_Name: 'a.fullName',
    Current_Address: 'a.Address_Live',
    Transportation_Mode: 'a.Vehicle',
    Bus_Route: 'a.Bus_Route',
    Pickup_Point: 'a.PickupDropoffStation',
    Number_of_Working_Days: 'e.Number_of_Working_Days',
  };
  async getCustomExport(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'ID',
    sortOrder: string = 'asc',
  ) {
    const db = getOwn(this.factoryDbMap, factory);
    const build = (factoryName: string, dateToExclusive: string) =>
      buildQueryCustomExport(dateFrom, dateToExclusive, factoryName);

    if (!db) {
      return this.getAllFactoriesPage(
        build,
        this.customExportAllSort,
        dateFrom,
        dateTo,
        page,
        limit,
        sortField,
        sortOrder,
        true,
      );
    }

    return this.getPageOfFactory(
      db,
      factory,
      build,
      this.customExportSort,
      dateFrom,
      dateTo,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }
  // auto send cms

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    try {
      if (factory.trim().toUpperCase() === 'ALL') {
        return this.autoSentCMSAllFactories(dateFrom, dateTo);
      }

      return this.getCMSByFactory(factory, dateFrom, dateTo);
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  private async autoSentCMSAllFactories(dateFrom: string, dateTo: string) {
    const results = await Promise.all(
      FACTORY_LIST.map((factory) =>
        this.getCMSByFactory(factory, dateFrom, dateTo),
      ),
    );

    return results.flat();
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
  ) {
    const db = getOwn(this.factoryDbMap, factory);
    if (!db) return [];

    const buildQuery = getOwn(this.buildQueryMap, factory);
    if (!buildQuery) return [];

    const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);

    const query = await buildQuery(dateFrom, dateTo, this.EIP);

    const replacements = hasDate ? [dateFrom, dateToExclusive] : [];

    const data = await db.query<Cat7CMSRow>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    return data.flatMap((item) => this.mapToCMSFormat(item, dateFrom, dateTo));
  }

  private mapToCMSFormat(item: Cat7CMSRow, dateFrom: string, dateTo: string) {
    const staffId = item.Staff_ID ?? '';
    const Main_transportation_type = item.Main_transportation_type ?? '';
    const Number_of_working_days = item.Number_of_working_days ?? '';
    const Factory_address = item.Factory_address ?? '';
    const Residential_address = item.Residential_address ?? Factory_address;
    const FactoryName = item.Factory_Name ?? '';
    const DepartmentName = item.Department_Name ?? '';

    let dockey = '';
    let transType = '';

    switch (Main_transportation_type.trim().toLowerCase()) {
      case 'Walking'.trim().toLowerCase(): //步行
        dockey = '3.3.1';
        transType = '步行';
        break;
      case 'Bicycle'.trim().toLowerCase(): //自行車
        dockey = '3.3.2';
        transType = '自行車';
        break;
      case 'Electric motorcycle'.trim().toLowerCase(): //電動機車
        dockey = '3.3.3';
        transType = '電動機車';
        break;
      case 'Motorcycle'.trim().toLowerCase(): //機車
        dockey = '3.3.4';
        transType = '機車';
        break;
      case 'Electric car'.trim().toLowerCase(): //電動汽車
        dockey = '3.3.5';
        transType = '電動汽車';
        break;
      case 'Car'.trim().toLowerCase(): //汽車
        dockey = '3.3.6';
        transType = '汽車';
        break;
      case 'Bus'.trim().toLowerCase(): //公車/客運
        dockey = '3.3.7';
        transType = '公車/客運';
        break;
      case 'Company shuttle bus'.trim().toLowerCase(): //員工接駁車
        dockey = '3.3.8';
        transType = '員工接駁車';
        break;
      case 'Subway'.trim().toLowerCase(): //地鐵
        dockey = '3.3.9';
        transType = '地鐵';
        break;
      default:
        break;
    }

    return ACTIVITY_TYPES.map((activityType: ActivityType) => ({
      System: 'CMS Web',
      Corporation: 'LAI YIH',
      Factory: FactoryName,
      Department: DepartmentName.slice(0, 20),
      DocKey: dockey,
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: activityType,
      DataType: '2',
      DocType: 'CMS Web',
      DocDate: dayjs('2025/12/31').format('YYYY/MM/DD'),
      DocDate2: dayjs('2025/12/31').format('YYYY/MM/DD'),
      UndDocNo: staffId,
      TransType: transType,
      Departure: Residential_address,
      Destination: Factory_address,
      Attendance: Number_of_working_days.toString(),
      Memo: '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    }));
  }
}
