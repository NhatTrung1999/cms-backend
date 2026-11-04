import { Inject, Injectable } from '@nestjs/common';
import { orderByClause } from 'src/helper/sql.helper';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class InfofactoryService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}
  async getInfoFactory(
    companyName: string,
    city: string,
    sortField: string,
    sortOrder: string,
  ) {
    let where = 'WHERE 1=1';
    const replacements: any[] = [];

    if (companyName) {
      where += ` AND CompanyName LIKE ?`;
      replacements.push(`%${companyName}%`);
    }

    if (city) {
      where += ` AND City LIKE ?`;
      replacements.push(`%${city}%`);
    }

    const query = `SELECT * FROM CMW_Info_Factory ${where} ORDER BY ${orderByClause(sortField, sortOrder)}`;

    const data = await this.EIP.query(query, {
      replacements,
      type: QueryTypes.SELECT,
    });
    return data;
  }
}
