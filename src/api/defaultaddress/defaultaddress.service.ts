import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class DefaultaddressService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}
  async getDefaultAddress(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Factory LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
    SELECT *,
           COUNT(ID) OVER() AS total
    FROM CMW_Default_Address
    ${where}
    ORDER BY CreatedAt DESC
    `,
      {
        type: QueryTypes.SELECT,
        replacements,
      },
    );

    let data = dataResults as any[];

    data.sort((a, b) => {
      const aValue = a[sortField];
      const bValue = b[sortField];
      if (sortOrder === 'asc') {
        return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
      } else {
        return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
      }
    });

    data = data.slice(offset, offset + limit);

    const cleanData = data.map(({ total, ...rest }) => rest);

    const total = dataResults.length
      ? Number((dataResults[0] as any).total)
      : 0;

    const hasMore = offset + data.length < total;

    return { data: cleanData, page, limit, total, hasMore };
  }
}
