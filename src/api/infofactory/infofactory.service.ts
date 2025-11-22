import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class InfofactoryService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}
  async getInfoFactory() {
    const query = 'SELECT * FROM CMW_Info_Factory';
    const data = await this.EIP.query(query, { type: QueryTypes.SELECT})
    console.log(data);
  }
}
