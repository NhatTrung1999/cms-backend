import {
  Inject,
  Injectable,
  UnauthorizedException,
  InternalServerErrorException,
} from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';

@Injectable()
export class Cat9andcat12Service {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}

  async getData(date, userID) {
    if (!userID) throw new UnauthorizedException();

    try {
      let where = ' 1 = 1 ';

      if (date !== '') {
      }

      const payload: any = await this.EIP.query<any>(
        `
          SELECT *
          FROM CMS_File_Management
          WHERE CreatedAt = '${userID}' ${where}
        `,
        {
          replacements: [],
          type: QueryTypes.SELECT,
        },
      );
      return payload;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }
}
