import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import {
  buildCat5CountQuery,
  buildCat5Filter,
  buildCat5PageQuery,
} from 'src/helper/cat5.queries';
import { getCat5CMSRows } from 'src/helper/cat5.cms';
import { sortByField } from 'src/helper/sort.helper';
import { getOwn } from 'src/helper/object.helper';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';

// The factories that have a WMS database (the order is the order of the "ALL" page).
const CAT5_FACTORIES: FactoryCode[] = [
  'LYV',
  'LHG',
  'LYM',
  'LVL',
  'JAZ',
  'JZS',
];

@Injectable()
export class Cat5Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
  ) {}

  async getDataWMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Waste_disposal_date',
    sortOrder: string = 'asc',
  ) {
    const db = CAT5_FACTORIES.includes(factory)
      ? this.getDbByFactory(factory)
      : null;
    if (!db) {
      return await this.getAllDataFactory(
        dateFrom,
        dateTo,
        page,
        limit,
        sortField,
        sortOrder,
      );
    }

    return await this.getDataFactory(
      db,
      dateFrom,
      dateTo,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  private async getDataFactory(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;
    const { where, replacements } = buildCat5Filter(dateFrom, dateTo);

    const query = buildCat5PageQuery(where);

    const countQuery = buildCat5CountQuery(where);

    const [dataResults, countResults] = await Promise.all([
      db.query(query, { replacements, type: QueryTypes.SELECT }),
      db.query(countQuery, {
        type: QueryTypes.SELECT,
        replacements,
      }),
    ]);

    let data = dataResults;
    sortByField(data, sortField, sortOrder);

    data = data.slice(offset, offset + limit);

    const total = (countResults[0] as { total: number })?.total || 0;

    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
  }

  private async getAllDataFactory(
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;

      const { where, replacements } = buildCat5Filter(dateFrom, dateTo);

      const query = buildCat5PageQuery(where);

      const countQuery = buildCat5CountQuery(where);
      const connects = CAT5_FACTORIES.map(
        (factoryName) => this.getDbByFactory(factoryName)!,
      );

      const [dataResults, countResults] = await Promise.all([
        Promise.all(
          connects.map((conn) => {
            return conn.query(query, { type: QueryTypes.SELECT, replacements });
          }),
        ),
        Promise.all(
          connects.map((conn) => {
            return conn.query(countQuery, {
              type: QueryTypes.SELECT,
              replacements,
            });
          }),
        ),
      ]);

      let data = dataResults.flat();

      sortByField(data, sortField, sortOrder);

      data = data.slice(offset, offset + limit);

      const total = countResults.reduce((sum, result) => {
        return sum + ((result[0] as { total: number })?.total || 0);
      }, 0);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  // auto sent cms
  async autoSentCMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    try {
      if (factory.trim().toUpperCase() === 'ALL') {
        return this.autoSentCMSAllFactories(dateFrom, dateTo, dockeyCMS);
      }

      return this.getCMSByFactory(factory, dateFrom, dateTo, dockeyCMS);
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  private async autoSentCMSAllFactories(
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const results = await Promise.all(
      FACTORY_LIST.map((factory) =>
        this.getCMSByFactory(factory, dateFrom, dateTo, dockeyCMS),
      ),
    );

    return results.flat();
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    return getCat5CMSRows(db, this.EIP, factory, dateFrom, dateTo, dockeyCMS);
  }

  private getDbByFactory(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<FactoryCode, Sequelize> = {
      LYV: this.LYV_WMS,
      LHG: this.LHG_WMS,
      LVL: this.LVL_WMS,
      LYM: this.LYM_WMS,
      JAZ: this.JAZ_WMS,
      JZS: this.JZS_WMS,
    };

    return getOwn(dbMap, factory) ?? null;
  }
}
