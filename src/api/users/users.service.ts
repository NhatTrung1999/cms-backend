import { Inject, Injectable } from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class UsersService {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('ERP') private readonly ERP: Sequelize,
  ) {}

  async validateUser(userid: string, password?: string, factory?: string) {
    // console.log(password);
    let query = `
    SELECT *
    FROM CMS_Account
    WHERE UserID = ?
  `;
    const replacements: any[] = [userid];
    if (password) {
      query += ` AND [Password] = ?`;
      replacements.push(password);
    }
    const payload: any = await this.EIP.query<any>(query, {
      replacements,
      type: QueryTypes.SELECT,
    });
    // const payload: any = await this.EIP.query<any>(
    // `
    //   SELECT *
    //   FROM CMS_Account
    //   WHERE UserID = ? AND [Password] = ?
    // `,
    //   {
    //     replacements: [userid, password],
    //     type: QueryTypes.SELECT,
    //   },
    // );
    if (payload.length === 0) return false;
    return payload[0];
  }

  async validateErpUser(userid: string, password: string) {
    const payload: any = await this.ERP.query<any>(
      `
        SELECT USERID, USERNAME, EMAIL, PWD, LOCK
        FROM Busers
        WHERE USERID = ? AND PWD = ?  
      `,
      {
        replacements: [userid, password],
        type: QueryTypes.SELECT,
      },
    );

    if (payload.length === 0) return false;
    return payload[0];
  }

  async checkLock(userid: string) {
    const payload: any = await this.ERP.query<any>(
      `
      SELECT USERID, USERNAME, EMAIL, PWD, LOCK
      FROM Busers
      WHERE USERID = ? AND LOCK = 'Y'  
    `,
      {
        replacements: [userid],
        type: QueryTypes.SELECT,
      },
    );

    if (payload.length === 1) return false;

    return true;
  }

  async checkExistUser(userid: string) {
    const payload: any = await this.EIP.query<any>(
    `
      SELECT *
      FROM CMS_Account
      WHERE UserID = ?
    `,
      {
        replacements: [userid],
        type: QueryTypes.SELECT,
      },
    );
    if (payload.length === 0) return false;
    return payload[0];
  }
}
