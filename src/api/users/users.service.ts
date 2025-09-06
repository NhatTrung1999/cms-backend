import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class UsersService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}

  async validateUser(userid: string, password: string, factory: string) {
    const payload: any = await this.EIP.query<any>(
      `
      SELECT *
      FROM Carbon_Management_System_Account
      WHERE UserID = ? AND [Password] = ?
    `,
      {
        replacements: [userid, password],
        type: QueryTypes.SELECT,
      },
    );
    if (payload.length === 0) return false;
    return payload[0];
  }
}
